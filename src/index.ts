#!/usr/bin/env node

import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import {
  type AuthenticationResult,
  type Configuration,
  PublicClientApplication,
} from "@azure/msal-node";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { cachePlugin } from "./msal-cache.js";
import {
  getAccount,
  getAccountsStore,
  listAccounts,
  removeAccount,
  removeAllAccounts,
  saveAccount,
  setActiveAccount,
} from "./services/account-manager.js";
import { authenticateViaBrowser, type BrowserAuthApp } from "./services/browser-auth.js";
import { FULL_SCOPES, GraphService, READ_ONLY_SCOPES } from "./services/graph.js";
import { registerAuthTools } from "./tools/auth.js";
import { registerChatTools } from "./tools/chats.js";
import { registerSearchTools } from "./tools/search.js";
import { registerTeamsTools } from "./tools/teams.js";
import { registerUsersTools } from "./tools/users.js";
import { setupDnsLookupFallback } from "./utils/dns-patch.js";

// Initialize DNS fallback for resilient network connections
setupDnsLookupFallback();

// Microsoft Graph CLI app ID (default public client)
const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = "https://login.microsoftonline.com/common";

const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Check whether CLI args contain --read-only. */
function hasReadOnlyFlag(args: string[]): boolean {
  return args.includes("--read-only");
}

/** Read the persisted auth info file (best-effort). */
async function readAuthInfo(): Promise<Record<string, unknown> | undefined> {
  try {
    const data = await fs.readFile(AUTH_INFO_PATH, "utf8");
    return JSON.parse(data) as Record<string, unknown>;
  } catch {
    return undefined;
  }
}

// Authentication functions
async function authenticate(
  readOnly: boolean,
  isBrowser = false,
  selectedApp?: BrowserAuthApp,
  accountId?: string
) {
  if (isBrowser) {
    try {
      await authenticateViaBrowser({ app: selectedApp, accountId, readOnly });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      console.error("\n❌ Browser authentication failed:", errorMessage);
      process.exit(1);
    }
    return;
  }

  const scopes = readOnly ? READ_ONLY_SCOPES : FULL_SCOPES;
  const modeLabel = readOnly ? "read-only" : "full access";

  console.log("🔐 Microsoft Graph Authentication for MCP Server");
  console.log("=".repeat(50));
  console.log(`Using Microsoft Graph CLI app (${modeLabel})`);

  try {
    console.log("\n📱 Using device code flow...");

    const msalConfig: Configuration = {
      auth: {
        clientId: CLIENT_ID,
        authority: AUTHORITY,
      },
      cache: {
        cachePlugin, // Use our custom file-based cache for refresh tokens
      },
    };

    const client = new PublicClientApplication(msalConfig);

    const result: AuthenticationResult | null = await client.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        console.log("\n📱 Please complete authentication:");
        console.log(`🌐 Visit: ${response.verificationUri}`);
        console.log(`🔑 Enter code: ${response.userCode}`);
        console.log("\n⏳ Waiting for you to complete authentication...");
      },
    });

    if (result) {
      const accountName = result.account?.username || "default";
      const targetAccountId = accountId || accountName;

      await saveAccount(
        {
          id: targetAccountId,
          account: accountName,
          displayName: result.account?.name || accountName,
          targetApp: "teams",
          authMethod: "device_code",
          token: result.accessToken,
          expiresAt: result.expiresOn?.toISOString(),
          lastRefreshed: new Date().toISOString(),
          grantedScopes: result.scopes,
          authenticated: true,
          clientId: CLIENT_ID,
        },
        true
      );

      console.log("\n✅ Authentication successful!");
      console.log(`👤 Signed in as: ${result.account?.username || "Unknown"}`);
      console.log(`🆔 Account ID: ${targetAccountId}`);
      console.log(`🔒 Mode: ${modeLabel}`);
      console.log("🔄 Refresh token cached for automatic renewal");
      console.log("\n🚀 You can now use the MCP server in Cursor!");
      console.log("   The server will automatically use these credentials.");
    }
  } catch (error) {
    const errorMessage = error instanceof Error ? error.message : String(error);

    // Provide helpful error messages for common issues
    if (errorMessage.includes("AADSTS50020")) {
      console.error("\n❌ Authentication failed: User account not in tenant");
    } else if (errorMessage.includes("AADSTS65001")) {
      console.error("\n❌ Authentication failed: Admin consent required");
      console.error("   Grant admin consent for the required permissions in Azure Portal");
    } else {
      console.error("\n❌ Authentication failed:", errorMessage);
    }
    console.log("\n💡 Tip: If device code login is restricted in your tenant, try browser login:");
    console.log("   npx @kida7/teams-mcp@latest authenticate --browser");
    process.exit(1);
  }
}

async function handleListAccounts() {
  const store = await listAccounts();
  if (store.accounts.length === 0) {
    console.log("❌ No accounts configured.");
    console.log("👉 Add an account: npx @kida7/teams-mcp@latest authenticate --browser");
    return;
  }
  console.log("\n👥 Authenticated Microsoft 365 Accounts:");
  console.log("=".repeat(50));
  for (const acc of store.accounts) {
    const isActive = acc.id === store.activeAccount;
    const autoRefresh = acc.profileDir ? "Auto-refresh: ENABLED" : "Auto-refresh: DISABLED";
    const expires = acc.expiresAt
      ? new Date(acc.expiresAt) > new Date()
        ? `Expires: ${new Date(acc.expiresAt).toLocaleTimeString()}`
        : "Token expired (Auto-refresh on request)"
      : "Token: active";
    console.log(
      `${isActive ? "👉 [ACTIVE] " : "   • "} ${acc.displayName || acc.account} (${acc.account})`
    );
    console.log(
      `     ID: ${acc.id} | Target App: ${acc.targetApp} | ${autoRefresh} | ${expires}\n`
    );
  }
  console.log("💡 Switch active account: npx @kida7/teams-mcp@latest use <accountId>");
  console.log(
    "💡 Run MCP for specific account: TEAMS_MCP_ACCOUNT=<accountId> npx @kida7/teams-mcp@latest"
  );
}

async function handleSwitchAccount(accountIdOrAlias?: string) {
  if (!accountIdOrAlias) {
    console.error("❌ Please specify an account ID or email to switch to.");
    console.error("   Usage: npx @kida7/teams-mcp@latest use <accountId>");
    process.exit(1);
  }
  try {
    const acc = await setActiveAccount(accountIdOrAlias);
    console.log(
      `✅ Switched active account to: ${acc.displayName || acc.account} (${acc.account}) [ID: ${acc.id}]`
    );
  } catch (err) {
    console.error(`❌ Error: ${err instanceof Error ? err.message : err}`);
    process.exit(1);
  }
}

async function checkAuth(targetAccountQuery?: string) {
  try {
    const targetAccount = await getAccount(targetAccountQuery);

    if (targetAccount?.authenticated) {
      console.log("✅ Authentication found");
      console.log(
        `👤 Account: ${targetAccount.displayName || targetAccount.account} (${targetAccount.account})`
      );
      console.log(`🆔 Account ID: ${targetAccount.id}`);
      console.log(
        `🌐 Target App: ${targetAccount.targetApp === "teams" ? "Microsoft Teams Web" : "Outlook Web"}`
      );
      console.log(
        `🔄 Silent Auto-Refresh: ${targetAccount.profileDir ? "Enabled (Persistent Profile)" : "Disabled"}`
      );

      if (targetAccount.expiresAt) {
        const expiresAt = new Date(targetAccount.expiresAt);
        const now = new Date();

        if (expiresAt > now) {
          console.log(`⏰ Access token expires: ${expiresAt.toLocaleString()}`);
          console.log("🎯 Ready to use with MCP server!");
        } else {
          console.log(
            `⏰ Access token expired on ${expiresAt.toLocaleString()} (Will auto-refresh silently on request)`
          );
          console.log("🎯 Ready to use with MCP server!");
        }
      } else {
        console.log("🎯 Ready to use with MCP server!");
      }
      return true;
    }

    const legacyAuth = await readAuthInfo();
    if (legacyAuth?.authenticated) {
      console.log("✅ Authentication found (legacy)");
      console.log(`👤 Account: ${legacyAuth.account || "Unknown"}`);
      return true;
    }
  } catch (_error) {
    console.log("❌ No authentication found");
    return false;
  }

  console.log("❌ No authentication found");
  console.log("👉 Run 'npx @kida7/teams-mcp@latest authenticate --browser' to log in.");
  return false;
}

async function logout(targetAccount?: string, removeAll = false) {
  if (removeAll) {
    await removeAllAccounts();
    console.log("✅ Successfully logged out of ALL accounts and cleared browser profiles.");
    return;
  }

  if (targetAccount) {
    const removed = await removeAccount(targetAccount);
    if (removed) {
      console.log(`✅ Successfully logged out and removed account '${targetAccount}'.`);
    } else {
      console.log(`⚠️ Account '${targetAccount}' not found.`);
    }
    return;
  }

  const store = await getAccountsStore();
  if (store.activeAccount) {
    const active = store.activeAccount;
    await removeAccount(active);
    console.log(`✅ Successfully logged out active account '${active}'.`);
  } else {
    await removeAllAccounts();
    console.log("✅ Successfully logged out.");
  }
}

// MCP Server setup
async function startMcpServer(readOnly: boolean, targetAccount?: string) {
  // Create MCP server
  const server = new McpServer({
    name: "teams-mcp",
    version: "1.0.0",
  });

  // Initialize Graph service (singleton)
  const graphService = GraphService.getInstance();
  graphService.readOnlyMode = readOnly;
  if (targetAccount) {
    graphService.requestedAccount = targetAccount;
  }

  // Detect scope mismatch: warn when switching from read-only → full mode
  if (!readOnly && !process.env.AUTH_TOKEN) {
    const authInfo = await readAuthInfo();
    if (authInfo) {
      const grantedScopes = authInfo.grantedScopes as string[] | undefined;
      const hasWriteScopes = grantedScopes?.some(
        (s: string) =>
          s === "ChannelMessage.Send" ||
          s === "ChannelMessage.ReadWrite" ||
          s === "Chat.ReadWrite" ||
          s === "Files.ReadWrite.All"
      );
      if (grantedScopes && !hasWriteScopes) {
        console.error(
          "⚠️  Warning: You authenticated with read-only scopes but the server is running in full mode."
        );
        console.error("   Write operations may fail. Re-authenticate without --read-only:");
        console.error("   npx @kida7/teams-mcp@latest authenticate");
      }
    }
  }

  // Register all tools (write tools are skipped when readOnly is true)
  registerAuthTools(server, graphService, readOnly);
  registerUsersTools(server, graphService, readOnly);
  registerTeamsTools(server, graphService, readOnly);
  registerChatTools(server, graphService, readOnly);
  registerSearchTools(server, graphService, readOnly);

  // Start server
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error(
    `Microsoft Graph MCP Server started${readOnly ? " (read-only mode)" : ""}${targetAccount ? ` [Account: ${targetAccount}]` : ""}`
  );
}

// Main function to handle both CLI and MCP server modes
async function main() {
  const args = process.argv.slice(2);
  const readOnly = hasReadOnlyFlag(args) || process.env.TEAMS_MCP_READ_ONLY === "true";
  const isBrowser =
    args.includes("--browser") ||
    args.includes("-b") ||
    args.includes("auth:browser") ||
    args.includes("browser-auth");
  const removeAll = args.includes("--all");

  let selectedApp: BrowserAuthApp | undefined;
  let targetAccountId: string | undefined;

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    if (arg === "--app" || arg === "-a") {
      const nextArg = args[i + 1]?.toLowerCase();
      if (nextArg === "teams" || nextArg === "msteams") selectedApp = "teams";
      if (nextArg === "outlook") selectedApp = "outlook";
    } else if (arg.startsWith("--app=")) {
      const val = arg.split("=")[1]?.toLowerCase();
      if (val === "teams" || val === "msteams") selectedApp = "teams";
      if (val === "outlook") selectedApp = "outlook";
    } else if (arg === "--outlook") {
      selectedApp = "outlook";
    } else if (arg === "--teams" || arg === "--msteams") {
      selectedApp = "teams";
    } else if (arg === "--account" || arg === "-u" || arg === "--profile") {
      targetAccountId = args[i + 1];
    } else if (arg.startsWith("--account=")) {
      targetAccountId = arg.split("=")[1];
    } else if (arg.startsWith("--profile=")) {
      targetAccountId = arg.split("=")[1];
    }
  }

  // Find the primary command or help flag
  const command = args.find(
    (arg) =>
      !arg.startsWith("--app") &&
      arg !== "-a" &&
      arg !== "--outlook" &&
      arg !== "--teams" &&
      arg !== "--msteams" &&
      arg !== "--browser" &&
      arg !== "-b" &&
      arg !== "--read-only" &&
      arg !== "--all" &&
      !arg.startsWith("--account") &&
      !arg.startsWith("--profile") &&
      arg !== "-u" &&
      arg !== targetAccountId
  );

  // CLI commands
  switch (command) {
    case "authenticate":
    case "auth":
      await authenticate(readOnly, isBrowser, selectedApp, targetAccountId);
      return;
    case "auth:browser":
    case "browser-auth":
      await authenticate(readOnly, true, selectedApp, targetAccountId);
      return;
    case "accounts":
    case "list-accounts":
    case "auth:list":
      await handleListAccounts();
      return;
    case "use":
    case "switch":
    case "switch-account": {
      const targetParam =
        targetAccountId || args.find((arg) => !arg.startsWith("-") && arg !== command);
      await handleSwitchAccount(targetParam);
      return;
    }
    case "check":
      await checkAuth(targetAccountId);
      return;
    case "logout": {
      const targetParam =
        targetAccountId || args.find((arg) => !arg.startsWith("-") && arg !== command);
      await logout(targetParam, removeAll);
      return;
    }
    case "help":
    case "--help":
    case "-h":
      console.log("Microsoft Graph MCP Server");
      console.log("");
      console.log("Usage:");
      console.log(
        "  npx @kida7/teams-mcp@latest authenticate                          # Authenticate via Device Code (default)"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest authenticate --browser                # Authenticate via Playwright browser"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest authenticate --browser --account work # Authenticate and save under alias 'work'"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest authenticate --browser --app outlook  # Open Outlook in browser to capture token"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest authenticate --browser --app teams    # Open MS Teams in browser to capture token"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest accounts                             # List all configured accounts"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest use <accountId>                      # Switch active account"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest check                                # Check authentication status"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest logout [accountId] [--all]           # Clear authentication"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest                                      # Start MCP server (active account)"
      );
      console.log(
        "  npx @kida7/teams-mcp@latest --account <accountId>                # Start MCP server for specific account"
      );
      console.log("");
      console.log("Options:");
      console.log(
        "  --browser, -b              Use Playwright browser login to capture Graph token"
      );
      console.log(
        "  --account <id>, -u         Account ID or alias for multi-account login or server target"
      );
      console.log(
        "  --app <outlook|teams>, -a  Target web application for browser login (default: prompt)"
      );
      console.log("  --outlook                  Shortcut for --app outlook");
      console.log("  --teams                    Shortcut for --app teams");
      console.log("  --read-only                Authenticate or run with read-only scopes");
      console.log("  --all                      Apply to all accounts (e.g. logout --all)");
      console.log("");
      console.log("Environment variables:");
      console.log("  TEAMS_MCP_ACCOUNT=<id>     # Select specific account for MCP server");
      console.log("  TEAMS_MCP_READ_ONLY=true   # Start MCP server in read-only mode");
      console.log("  AUTH_TOKEN=<jwt>           # Use a pre-existing access token");
      return;
    case undefined:
      if (isBrowser) {
        await authenticate(readOnly, true, selectedApp, targetAccountId);
        return;
      }
      // No command = start MCP server
      await startMcpServer(readOnly, targetAccountId);
      return;
    default:
      console.error(`Unknown command: ${command}`);
      console.error("Use --help to see available commands");
      process.exit(1);
  }
}

// Handle uncaught errors
process.on("uncaughtException", (error) => {
  console.error("Uncaught exception:", error);
  process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
  console.error("Unhandled rejection at:", promise, "reason:", reason);
  process.exit(1);
});

main().catch((error) => {
  console.error("Failed to start:", error);
  process.exit(1);
});
