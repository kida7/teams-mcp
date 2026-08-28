import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

export interface AccountData {
  id: string;
  account: string;
  displayName?: string | undefined;
  targetApp: "outlook" | "teams";
  authMethod: "browser" | "device_code" | "token";
  profileDir?: string | undefined;
  token?: string | undefined;
  expiresAt?: string | undefined;
  lastRefreshed?: string | undefined;
  grantedScopes?: string[] | undefined;
  authenticated: boolean;
  clientId?: string | undefined;
}

export interface AccountsStore {
  activeAccount?: string | undefined;
  accounts: Record<string, AccountData>;
}

export const TEAMS_MCP_CONFIG_DIR = join(homedir(), ".teams-mcp");
export const ACCOUNTS_FILE_PATH = join(TEAMS_MCP_CONFIG_DIR, "accounts.json");
export const PROFILES_BASE_DIR = join(TEAMS_MCP_CONFIG_DIR, "profiles");
export const LEGACY_AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Convert an account identifier / email to a safe directory name */
export function sanitizeAccountName(accountName: string): string {
  return accountName.replace(/[^a-zA-Z0-9_-]/g, "_").toLowerCase();
}

/** Get the persistent browser profile directory for a given account */
export function getProfileDirForAccount(accountId: string): string {
  const safeName = sanitizeAccountName(accountId);
  return join(PROFILES_BASE_DIR, safeName);
}

/** Ensure the configuration directory and profiles base directory exist */
export async function ensureConfigDirs(): Promise<void> {
  try {
    await fs.mkdir(TEAMS_MCP_CONFIG_DIR, { recursive: true });
    await fs.mkdir(PROFILES_BASE_DIR, { recursive: true });
  } catch {
    // Ignore if already exists
  }
}

/** Read accounts store, automatically migrating legacy auth file if needed */
export async function getAccountsStore(): Promise<AccountsStore> {
  await ensureConfigDirs();

  try {
    const data = await fs.readFile(ACCOUNTS_FILE_PATH, "utf8");
    if (data && typeof data === "string" && data.trim()) {
      const parsed = JSON.parse(data);
      if (parsed && typeof parsed.accounts === "object") {
        return parsed as AccountsStore;
      }
      // If parsed data contains token directly (legacy auth info)
      if (parsed?.token && typeof parsed.token === "string") {
        const accountId = (parsed.account as string) || "default";
        return {
          activeAccount: accountId,
          accounts: {
            [accountId]: {
              id: accountId,
              account: accountId,
              displayName: parsed.displayName,
              targetApp: parsed.targetApp || "outlook",
              authMethod: parsed.authMethod || "browser",
              profileDir: parsed.profileDir,
              token: parsed.token,
              expiresAt: parsed.expiresAt,
              authenticated: true,
              clientId: parsed.clientId,
            },
          },
        };
      }
    }
  } catch {
    // accounts.json does not exist yet
  }

  // Attempt migration from legacy ~/.msgraph-mcp-auth.json
  try {
    const legacyData = await fs.readFile(LEGACY_AUTH_INFO_PATH, "utf8");
    if (legacyData && typeof legacyData === "string" && legacyData.trim()) {
      const legacyInfo = JSON.parse(legacyData);
      if (legacyInfo?.authenticated) {
        const accountId = (legacyInfo.account as string) || "default";
        const legacyAccount: AccountData = {
          id: accountId,
          account: accountId,
          displayName: legacyInfo.displayName as string | undefined,
          targetApp: (legacyInfo.targetApp as "outlook" | "teams") || "outlook",
          authMethod: (legacyInfo.authMethod as "browser" | "device_code" | "token") || "browser",
          profileDir:
            legacyInfo.authMethod === "browser" ? getProfileDirForAccount(accountId) : undefined,
          token: legacyInfo.token as string | undefined,
          expiresAt: legacyInfo.expiresAt as string | undefined,
          grantedScopes: legacyInfo.grantedScopes as string[] | undefined,
          authenticated: true,
          clientId: legacyInfo.clientId as string | undefined,
        };

        const initialStore: AccountsStore = {
          activeAccount: accountId,
          accounts: {
            [accountId]: legacyAccount,
          },
        };

        await fs.writeFile(ACCOUNTS_FILE_PATH, JSON.stringify(initialStore, null, 2), "utf8");
        return initialStore;
      }
    }
  } catch {
    // No legacy auth file
  }

  return {
    activeAccount: undefined,
    accounts: {},
  };
}

/** Save accounts store to disk and sync active account to legacy auth file */
export async function saveAccountsStore(store: AccountsStore): Promise<void> {
  await ensureConfigDirs();
  await fs.writeFile(ACCOUNTS_FILE_PATH, JSON.stringify(store, null, 2), "utf8");

  // Sync active account to legacy auth file for backwards compatibility
  if (store.activeAccount && store.accounts[store.activeAccount]) {
    const active = store.accounts[store.activeAccount];
    const legacyPayload = {
      clientId: active.clientId || "teams-mcp",
      authenticated: active.authenticated,
      timestamp: active.lastRefreshed || new Date().toISOString(),
      expiresAt: active.expiresAt,
      account: active.account,
      displayName: active.displayName,
      grantedScopes: active.grantedScopes,
      authMethod: active.authMethod,
      targetApp: active.targetApp,
      token: active.token,
      profileDir: active.profileDir,
    };
    try {
      await fs.writeFile(LEGACY_AUTH_INFO_PATH, JSON.stringify(legacyPayload, null, 2), "utf8");
    } catch {
      // Ignore error writing legacy file
    }
  }
}

/** Get a specific account by ID or alias, or get the active account */
export async function getAccount(
  accountIdOrAlias?: string | undefined
): Promise<AccountData | undefined> {
  const store = await getAccountsStore();

  const query = accountIdOrAlias || process.env.TEAMS_MCP_ACCOUNT;

  if (query) {
    // 1. Exact match by ID
    if (store.accounts[query]) {
      return store.accounts[query];
    }

    // 2. Case-insensitive search by ID or email
    const lowerQuery = query.toLowerCase();
    const foundKey = Object.keys(store.accounts).find(
      (key) =>
        key.toLowerCase() === lowerQuery ||
        store.accounts[key]?.account.toLowerCase() === lowerQuery
    );
    if (foundKey && store.accounts[foundKey]) {
      return store.accounts[foundKey];
    }

    return undefined;
  }

  // Return active account if set
  if (store.activeAccount && store.accounts[store.activeAccount]) {
    return store.accounts[store.activeAccount];
  }

  // Fallback to first available account
  const allKeys = Object.keys(store.accounts);
  if (allKeys.length > 0 && store.accounts[allKeys[0]]) {
    return store.accounts[allKeys[0]];
  }

  return undefined;
}

/** Save or update an account in the store */
export async function saveAccount(account: AccountData, makeActive = true): Promise<void> {
  const store = await getAccountsStore();
  store.accounts[account.id] = account;

  if (makeActive || !store.activeAccount) {
    store.activeAccount = account.id;
  }

  await saveAccountsStore(store);
}

/** Set the active account */
export async function setActiveAccount(accountIdOrAlias: string): Promise<AccountData> {
  const store = await getAccountsStore();
  const lowerQuery = accountIdOrAlias.toLowerCase();

  const foundKey = Object.keys(store.accounts).find(
    (key) =>
      key.toLowerCase() === lowerQuery || store.accounts[key]?.account.toLowerCase() === lowerQuery
  );

  if (!foundKey || !store.accounts[foundKey]) {
    throw new Error(`Account '${accountIdOrAlias}' not found in saved accounts.`);
  }

  store.activeAccount = foundKey;
  await saveAccountsStore(store);
  return store.accounts[foundKey];
}

/** Remove an account and its profile directory */
export async function removeAccount(accountIdOrAlias: string): Promise<boolean> {
  const store = await getAccountsStore();
  const lowerQuery = accountIdOrAlias.toLowerCase();

  const foundKey = Object.keys(store.accounts).find(
    (key) =>
      key.toLowerCase() === lowerQuery || store.accounts[key]?.account.toLowerCase() === lowerQuery
  );

  if (!foundKey || !store.accounts[foundKey]) {
    return false;
  }

  const account = store.accounts[foundKey];

  // Remove profile directory if it exists
  if (account.profileDir) {
    try {
      await fs.rm(account.profileDir, { recursive: true, force: true });
    } catch {
      // Ignore deletion error
    }
  }

  delete store.accounts[foundKey];

  if (store.activeAccount === foundKey) {
    const remaining = Object.keys(store.accounts);
    store.activeAccount = remaining.length > 0 ? remaining[0] : undefined;
  }

  await saveAccountsStore(store);
  return true;
}

/** Remove all accounts and profile directories */
export async function removeAllAccounts(): Promise<void> {
  try {
    await fs.rm(TEAMS_MCP_CONFIG_DIR, { recursive: true, force: true });
  } catch {
    // Ignore
  }

  try {
    await fs.unlink(LEGACY_AUTH_INFO_PATH);
  } catch {
    // Ignore
  }

  try {
    await fs.unlink(join(homedir(), ".teams-mcp-token-cache.json"));
  } catch {
    // Ignore
  }
}

/** List all accounts with status metadata */
export async function listAccounts(): Promise<{
  activeAccount?: string | undefined;
  accounts: AccountData[];
}> {
  const store = await getAccountsStore();
  return {
    activeAccount: store.activeAccount,
    accounts: Object.values(store.accounts),
  };
}
