import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { stdin as input, stdout as output } from "node:process";
import * as readline from "node:readline/promises";
import { type BrowserContext, chromium } from "playwright";
import {
  type AccountData,
  getAccount,
  getActiveAccountId,
  getProfileDirForAccount,
  saveAccount,
} from "./account-manager.js";

export type BrowserAuthApp = "outlook" | "teams";

export interface DecodedJwtPayload {
  aud?: string | string[];
  iss?: string;
  exp?: number;
  nbf?: number;
  upn?: string;
  unique_name?: string;
  email?: string;
  preferred_username?: string;
  name?: string;
  scp?: string;
  roles?: string[];
  appid?: string;
  tid?: string;
  oid?: string;
  sub?: string;
  [key: string]: unknown;
}

export interface BrowserAuthOptions {
  app?: BrowserAuthApp | undefined;
  accountId?: string | undefined;
  timeoutMs?: number | undefined;
  readOnly?: boolean | undefined;
}

export interface BrowserAuthResult {
  token: string;
  authInfo: AccountData;
}

export const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Parse JWT payload safely in Node.js */
export function parseJwt(token: string): DecodedJwtPayload | null {
  try {
    const parts = token.split(".");
    if (parts.length !== 3) return null;
    const base64Url = parts[1];
    const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
    const jsonPayload = Buffer.from(base64, "base64").toString("utf8");
    return JSON.parse(jsonPayload) as DecodedJwtPayload;
  } catch {
    return null;
  }
}

/** Check if token is a valid, unexpired Microsoft Graph JWT token */
export function isLikelyGraphToken(token: string): boolean {
  const payload = parseJwt(token);
  if (!payload) return false;

  // Expiration check (must have at least 10 seconds of validity left)
  if (payload.exp && payload.exp * 1000 <= Date.now() + 10_000) {
    return false;
  }

  // Audience check
  const aud = payload.aud;
  if (!aud) return false;
  const audiences = Array.isArray(aud) ? aud : [aud];
  return audiences.some(
    (a) =>
      typeof a === "string" &&
      (a === "https://graph.microsoft.com" ||
        a === "https://graph.microsoft.com/" ||
        a === "00000003-0000-0000-c000-000000000000" ||
        a.includes("graph.microsoft.com"))
  );
}

/** Prompt user interactively in terminal for app selection (Outlook vs Teams) */
export async function promptAppSelection(): Promise<BrowserAuthApp> {
  if (!process.stdin.isTTY) {
    console.log("Non-interactive terminal detected. Defaulting to Outlook Web.");
    return "outlook";
  }

  const rl = readline.createInterface({ input, output });
  try {
    console.log("\nSelect application to open for authentication:");
    console.log("  1) Outlook Web (https://outlook.office.com)");
    console.log("  2) Microsoft Teams Web (https://teams.microsoft.com)");
    const answer = (await rl.question("\nEnter choice [1/2] (default: 1): ")).trim();

    if (answer === "2" || answer.toLowerCase() === "teams" || answer.toLowerCase() === "msteams") {
      return "teams";
    }
    return "outlook";
  } finally {
    rl.close();
  }
}

/** Launch Playwright persistent browser context with fallbacks for Windows / local environments */
export async function launchPersistentBrowserContext(
  profileDir: string,
  headless = false
): Promise<BrowserContext> {
  const args = ["--start-maximized"];
  const viewport = headless ? { width: 1280, height: 720 } : null;

  // 1. Try bundled Chromium
  try {
    return await chromium.launchPersistentContext(profileDir, {
      headless,
      args,
      viewport,
    });
  } catch (chromiumError) {
    // 2. Try Microsoft Edge
    try {
      return await chromium.launchPersistentContext(profileDir, {
        headless,
        channel: "msedge",
        args,
        viewport,
      });
    } catch {
      // 3. Try Google Chrome
      try {
        return await chromium.launchPersistentContext(profileDir, {
          headless,
          channel: "chrome",
          args,
          viewport,
        });
      } catch {
        const origMsg =
          chromiumError instanceof Error ? chromiumError.message : String(chromiumError);
        throw new Error(
          "Could not launch browser (Playwright Chromium, Edge, or Chrome).\n" +
            `Original error: ${origMsg}\n` +
            "Please run: npx playwright install chromium"
        );
      }
    }
  }
}

/**
 * Authenticate via Playwright browser by navigating to Outlook or Teams
 * and capturing the Authorization header of requests sent to graph.microsoft.com.
 * Uses a persistent browser profile so session cookies are stored permanently.
 */
export async function authenticateViaBrowser(
  options: BrowserAuthOptions = {}
): Promise<BrowserAuthResult> {
  const targetApp = options.app ?? (await promptAppSelection());
  const targetUrl =
    targetApp === "teams" ? "https://teams.microsoft.com/v2/" : "https://outlook.office.com/mail/";
  const appDisplayName = targetApp === "teams" ? "Microsoft Teams Web" : "Outlook Web";
  const timeoutMs = options.timeoutMs ?? 300_000; // 5 minutes

  // Determine initial profile directory: prioritize existing active account's profile if available
  const existingAccount = await getAccount(options.accountId);
  const activeAccountId = options.accountId || existingAccount?.id || (await getActiveAccountId());
  const initialProfileId = activeAccountId || `temp_${Date.now()}`;
  const profileDir = getProfileDirForAccount(initialProfileId);

  console.log("\n🌐 Microsoft Graph Browser Authentication");
  console.log("=".repeat(50));
  console.log(`Target Application: ${appDisplayName} (${targetUrl})`);
  console.log(`Persistent Profile: ${profileDir}`);
  if (options.accountId) {
    console.log(`Account ID / Alias: ${options.accountId}`);
  }
  console.log("\n⏳ Launching browser...");

  let context: BrowserContext | undefined;

  try {
    context = await launchPersistentBrowserContext(profileDir, false /* non-headless */);
    const pages = context.pages();
    const page = pages.length > 0 ? pages[0] : await context.newPage();

    let capturedToken: string | undefined;
    let isResolved = false;

    console.log("🌐 Browser window opened.");
    console.log("👉 Please sign in to your Microsoft account in the browser window.");
    console.log("⏳ Listening for requests to graph.microsoft.com...\n");

    const tokenPromise = new Promise<string>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        if (!isResolved) {
          isResolved = true;
          reject(
            new Error(`Authentication timed out after ${Math.round(timeoutMs / 1000)} seconds.`)
          );
        }
      }, timeoutMs);

      // Handle request interception
      const onRequest = (request: any) => {
        if (isResolved) return;
        try {
          const url = request.url();
          if (url.includes("graph.microsoft.com")) {
            const headers = request.headers();
            const authHeader = headers.authorization || headers.Authorization;
            if (authHeader && typeof authHeader === "string") {
              const match = authHeader.match(/^Bearer\s+(.+)$/i);
              const token = match ? match[1].trim() : authHeader.trim();
              if (token && isLikelyGraphToken(token) && !capturedToken) {
                capturedToken = token;
                isResolved = true;
                clearTimeout(timeoutId);
                resolve(token);
              }
            }
          }
        } catch {
          // ignore error in request handler
        }
      };

      context?.on("request", onRequest);

      // Close event
      context?.on("close", () => {
        if (!isResolved) {
          isResolved = true;
          clearTimeout(timeoutId);
          reject(new Error("Browser was closed before authentication completed."));
        }
      });
    });

    // Navigate to target URL
    await page.goto(targetUrl).catch(() => {
      // Navigation redirects during login are expected
    });

    const token = await tokenPromise;
    const payload = parseJwt(token);

    let accountName =
      payload?.upn ||
      payload?.unique_name ||
      payload?.email ||
      payload?.preferred_username ||
      payload?.name ||
      "Unknown User";
    let displayName = payload?.name || accountName;

    // Verify token with Graph API /me endpoint (best effort)
    try {
      const meResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (meResponse.ok) {
        const me = (await meResponse.json()) as any;
        if (me.displayName) displayName = me.displayName;
        if (me.userPrincipalName) accountName = me.userPrincipalName;
        else if (me.mail) accountName = me.mail;
      }
    } catch {
      // Best-effort
    }

    const expiresAt = payload?.exp ? new Date(payload.exp * 1000).toISOString() : undefined;
    const grantedScopes = payload?.scp ? payload.scp.split(" ") : undefined;

    const accountId = options.accountId || accountName;
    const finalProfileDir = getProfileDirForAccount(accountId);

    // If a temporary profile was used and differs from finalProfileDir, copy over
    if (profileDir !== finalProfileDir) {
      try {
        await fs.cp(profileDir, finalProfileDir, { recursive: true, force: true });
        await fs.rm(profileDir, { recursive: true, force: true }).catch(() => {});
      } catch {
        // Best-effort profile sync
      }
    }

    const accountData: AccountData = {
      id: accountId,
      account: accountName,
      displayName,
      targetApp,
      authMethod: "browser",
      profileDir: finalProfileDir,
      token,
      expiresAt,
      lastRefreshed: new Date().toISOString(),
      grantedScopes,
      authenticated: true,
      clientId: (payload?.appid as string) || (payload?.aud as string) || "browser-session",
    };

    // Save to multi-account store
    await saveAccount(accountData, true /* makeActive */);

    console.log("✅ Authorization token captured successfully!");
    console.log(`👤 Account: ${displayName} (${accountName})`);
    console.log(`🆔 Account ID: ${accountId}`);
    console.log(`🌐 Target: ${appDisplayName}`);
    if (expiresAt) {
      console.log(`⏰ Access token expires: ${new Date(expiresAt).toLocaleString()}`);
    }
    console.log(`📁 Persistent Profile saved to: ${finalProfileDir}`);
    console.log("🔄 Headless silent auto-refresh: ENABLED");
    console.log("\n🚀 You can now use the MCP server in Cursor / Claude / Antigravity!");
    console.log("   The server will automatically refresh credentials silently in the background.");

    return { token, authInfo: accountData };
  } finally {
    if (context) {
      try {
        await context.close();
      } catch {
        // Ignore close errors
      }
    }
  }
}

/**
 * Silently refresh the access token in headless mode using the account's persistent profile.
 * No UI or login prompt is displayed.
 */
export async function refreshAccountTokenSilent(
  accountIdOrAlias?: string | undefined,
  timeoutMs = 45_000
): Promise<string | undefined> {
  const account = await getAccount(accountIdOrAlias);
  if (!account?.profileDir) {
    return undefined;
  }

  const targetUrl =
    account.targetApp === "teams"
      ? "https://teams.microsoft.com/v2/"
      : "https://outlook.office.com/mail/";

  let context: BrowserContext | undefined;

  try {
    context = await launchPersistentBrowserContext(account.profileDir, true /* headless */);
    const pages = context.pages();
    const page = pages.length > 0 ? pages[0] : await context.newPage();

    let capturedToken: string | undefined;
    let isResolved = false;

    const tokenPromise = new Promise<string>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        if (!isResolved) {
          isResolved = true;
          reject(new Error(`Silent token refresh timed out after ${timeoutMs}ms`));
        }
      }, timeoutMs);

      const onRequest = (request: any) => {
        if (isResolved) return;
        try {
          const url = request.url();
          if (url.includes("graph.microsoft.com")) {
            const headers = request.headers();
            const authHeader = headers.authorization || headers.Authorization;
            if (authHeader && typeof authHeader === "string") {
              const match = authHeader.match(/^Bearer\s+(.+)$/i);
              const token = match ? match[1].trim() : authHeader.trim();
              if (token && isLikelyGraphToken(token) && !capturedToken) {
                capturedToken = token;
                isResolved = true;
                clearTimeout(timeoutId);
                resolve(token);
              }
            }
          }
        } catch {
          // ignore
        }
      };

      context?.on("request", onRequest);

      context?.on("close", () => {
        if (!isResolved) {
          isResolved = true;
          clearTimeout(timeoutId);
          reject(new Error("Browser context was closed."));
        }
      });
    });

    await page.goto(targetUrl, { waitUntil: "domcontentloaded", timeout: timeoutMs }).catch(() => {
      // Ignore navigation redirect errors during login
    });

    const token = await tokenPromise;
    const payload = parseJwt(token);

    account.token = token;
    account.expiresAt = payload?.exp ? new Date(payload.exp * 1000).toISOString() : undefined;
    account.lastRefreshed = new Date().toISOString();
    account.authenticated = true;

    await saveAccount(account, false);
    return token;
  } catch (err) {
    console.error(
      `Silent token refresh for '${account.account}' failed:`,
      err instanceof Error ? err.message : err
    );
    return undefined;
  } finally {
    if (context) {
      try {
        await context.close();
      } catch {
        // ignore
      }
    }
  }
}
