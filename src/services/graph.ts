import { promises as fs } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { type AccountInfo, PublicClientApplication } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";
import { cachePlugin } from "../msal-cache.js";
import { type AccountData, getAccount, listAccounts, setActiveAccount } from "./account-manager.js";
import { parseJwt, refreshAccountTokenSilent } from "./browser-auth.js";
import { setupDnsLookupFallback } from "../utils/dns-patch.js";

setupDnsLookupFallback();

const CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
const AUTHORITY = "https://login.microsoftonline.com/common";
const AUTH_INFO_PATH = join(homedir(), ".msgraph-mcp-auth.json");

/** Scopes sufficient for read-only operations (no message sending, no file uploads). */
export const READ_ONLY_SCOPES = [
  "User.Read",
  "User.ReadBasic.All",
  "Team.ReadBasic.All",
  "Channel.ReadBasic.All",
  "ChannelMessage.Read.All",
  "TeamMember.Read.All",
  "Chat.Read",
];

/** Full scopes including write operations. */
export const FULL_SCOPES = [
  ...READ_ONLY_SCOPES,
  "ChannelMessage.Send",
  "ChannelMessage.ReadWrite",
  "Chat.ReadWrite",
  "Files.ReadWrite.All",
];

export interface AuthStatus {
  isAuthenticated: boolean;
  userPrincipalName?: string | undefined;
  displayName?: string | undefined;
  expiresAt?: string | undefined;
  accountId?: string | undefined;
  targetApp?: string | undefined;
  authMethod?: string | undefined;
  autoRefresh?: boolean | undefined;
  availableAccounts?:
    | Array<{ id: string; account: string; displayName?: string | undefined }>
    | undefined;
}

export class GraphService {
  private static instance: GraphService;
  private client: Client | undefined;
  private isInitialized = false;
  private tokenExpiresAt: Date | undefined;
  private msalApp: PublicClientApplication | undefined;
  private msalAccount: AccountInfo | undefined;
  private _readOnlyMode = false;
  private _requestedAccountId: string | undefined;

  static getInstance(): GraphService {
    if (!GraphService.instance) {
      GraphService.instance = new GraphService();
    }
    return GraphService.instance;
  }

  /** Whether the service operates in read-only mode (reduced permission scopes). */
  get readOnlyMode(): boolean {
    return this._readOnlyMode;
  }

  set readOnlyMode(value: boolean) {
    this._readOnlyMode = value;
  }

  /** Specific account to use (if multi-account) */
  get requestedAccount(): string | undefined {
    return this._requestedAccountId;
  }

  set requestedAccount(value: string | undefined) {
    if (this._requestedAccountId !== value) {
      this._requestedAccountId = value;
      this.isInitialized = false;
      this.client = undefined;
    }
  }

  /** Switch the active account */
  async switchAccount(accountId: string): Promise<AccountData> {
    const account = await setActiveAccount(accountId);
    this.requestedAccount = account.id;
    this.isInitialized = false;
    this.client = undefined;
    await this.initializeClient();
    return account;
  }

  /** Returns the scopes to request based on the current mode. */
  get scopes(): string[] {
    return this._readOnlyMode ? READ_ONLY_SCOPES : FULL_SCOPES;
  }

  private async initializeClient(): Promise<void> {
    if (this.isInitialized) return;

    try {
      // Priority 1: AUTH_TOKEN environment variable (direct token injection)
      const envToken = process.env.AUTH_TOKEN;
      if (envToken) {
        const validatedToken = this.validateToken(envToken);
        if (validatedToken) {
          this.client = Client.initWithMiddleware({
            authProvider: {
              getAccessToken: async () => validatedToken,
            },
          });
          this.isInitialized = true;
          return;
        }
        console.warn("AUTH_TOKEN in environment is expired or invalid. Falling back to account store.");
      }

      // Priority 2: Multi-Account Store (~/.teams-mcp/accounts.json)
      const targetQuery = this._requestedAccountId || process.env.TEAMS_MCP_ACCOUNT;
      const account = await getAccount(targetQuery);

      if (account) {
        let activeToken = account.token;

        // Check if token needs refresh
        const isExpiredOrMissing =
          !activeToken ||
          !this.validateToken(activeToken) ||
          (account.expiresAt && new Date(account.expiresAt).getTime() - Date.now() < 5 * 60 * 1000);

        if (isExpiredOrMissing && account.authMethod === "browser" && account.profileDir) {
          // Attempt silent headless refresh with persistent browser profile
          const refreshed = await refreshAccountTokenSilent(account.id);
          if (refreshed) {
            activeToken = refreshed;
          }
        }

        if (activeToken) {
          const validated = this.validateToken(activeToken);
          if (validated) {
            if (account.expiresAt) {
              this.tokenExpiresAt = new Date(account.expiresAt);
            }
            this.client = Client.initWithMiddleware({
              authProvider: {
                getAccessToken: async () => {
                  // If token is about to expire mid-session, attempt silent refresh
                  if (
                    this.tokenExpiresAt &&
                    this.tokenExpiresAt.getTime() - Date.now() < 5 * 60 * 1000 &&
                    account.authMethod === "browser" &&
                    account.profileDir
                  ) {
                    const refreshed = await refreshAccountTokenSilent(account.id);
                    if (refreshed) {
                      const payload = parseJwt(refreshed);
                      if (payload?.exp) {
                        this.tokenExpiresAt = new Date(payload.exp * 1000);
                      }
                      return refreshed;
                    }
                  }
                  return account.token || activeToken || "";
                },
              },
            });
            this.isInitialized = true;
            return;
          }
        }
      }

      // Priority 3: MSAL with cached refresh token for automatic token renewal
      this.msalApp = new PublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: AUTHORITY,
        },
        cache: {
          cachePlugin,
        },
      });

      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      if (accounts.length > 0) {
        this.msalAccount = accounts[0];

        // Verify we can acquire a token
        const result = await this.msalApp.acquireTokenSilent({
          scopes: this.scopes,
          account: this.msalAccount,
        });

        if (result) {
          this.tokenExpiresAt = result.expiresOn ?? undefined;

          // Create Graph client with MSAL-backed auth provider for automatic token refresh
          this.client = Client.initWithMiddleware({
            authProvider: {
              getAccessToken: () => this.acquireToken(),
            },
          });

          this.isInitialized = true;
          return;
        }
      }

      // Priority 4: Legacy Stored token from browser authentication (AUTH_INFO_PATH)
      try {
        const authData = await fs.readFile(AUTH_INFO_PATH, "utf8");
        const authInfo = JSON.parse(authData);
        if (authInfo?.token && typeof authInfo.token === "string") {
          const validatedToken = this.validateToken(authInfo.token);
          if (validatedToken) {
            if (authInfo.expiresAt) {
              this.tokenExpiresAt = new Date(authInfo.expiresAt);
            }
            this.client = Client.initWithMiddleware({
              authProvider: {
                getAccessToken: async () => validatedToken,
              },
            });
            this.isInitialized = true;
            return;
          }
        }
      } catch {
        // No stored auth file or invalid JSON
      }
    } catch (error) {
      console.error("Failed to initialize Graph client:", error);
    }
  }

  private async acquireToken(): Promise<string> {
    if (!this.msalApp || !this.msalAccount) {
      throw new Error("MSAL not initialized");
    }

    const result = await this.msalApp.acquireTokenSilent({
      scopes: this.scopes,
      account: this.msalAccount,
    });

    if (!result) {
      throw new Error(
        "Failed to acquire access token. Please re-authenticate: npx @kida7/teams-mcp@latest authenticate"
      );
    }

    this.tokenExpiresAt = result.expiresOn ?? undefined;
    return result.accessToken;
  }

  async getAuthStatus(): Promise<AuthStatus> {
    await this.initializeClient();

    const accountsList = await listAccounts().catch(() => ({
      activeAccount: undefined,
      accounts: [],
    }));
    const targetQuery = this._requestedAccountId || process.env.TEAMS_MCP_ACCOUNT;
    const currentAccount = await getAccount(targetQuery).catch(() => undefined);

    const availableAccounts = accountsList.accounts.map((acc) => ({
      id: acc.id,
      account: acc.account,
      displayName: acc.displayName,
    }));

    if (!this.client) {
      return {
        isAuthenticated: false,
        availableAccounts: availableAccounts.length > 0 ? availableAccounts : undefined,
      };
    }

    try {
      const me = await this.client.api("/me").get();
      const status: AuthStatus = {
        isAuthenticated: true,
        userPrincipalName: me?.userPrincipalName ?? currentAccount?.account ?? undefined,
        displayName: me?.displayName ?? currentAccount?.displayName ?? undefined,
        expiresAt: this.tokenExpiresAt?.toISOString() ?? currentAccount?.expiresAt,
      };

      if (currentAccount?.id) status.accountId = currentAccount.id;
      if (currentAccount?.targetApp) status.targetApp = currentAccount.targetApp;
      if (currentAccount?.authMethod) status.authMethod = currentAccount.authMethod;
      if (currentAccount?.profileDir) status.autoRefresh = true;
      if (availableAccounts.length > 0) status.availableAccounts = availableAccounts;

      return status;
    } catch (error) {
      console.error("Error getting user info:", error);
      return {
        isAuthenticated: false,
        availableAccounts: availableAccounts.length > 0 ? availableAccounts : undefined,
      };
    }
  }

  async getClient(): Promise<Client> {
    await this.initializeClient();

    if (!this.client) {
      throw new Error(
        "Not authenticated. Please run the authentication CLI tool first: npx @kida7/teams-mcp@latest authenticate"
      );
    }
    return this.client;
  }

  isAuthenticated(): boolean {
    return !!this.client && this.isInitialized;
  }

  validateToken(token: string): string | undefined {
    const tokenSplits = token.split(".");
    if (tokenSplits.length !== 3) {
      console.error("Invalid JWT token: missing claims");
      return undefined;
    }

    try {
      const base64Url = tokenSplits[1];
      const base64 = base64Url.replace(/-/g, "+").replace(/_/g, "/");
      const payload = JSON.parse(Buffer.from(base64, "base64").toString("utf8"));

      // Expiration check
      if (payload.exp && payload.exp * 1000 <= Date.now()) {
        console.error("JWT token is expired");
        return undefined;
      }

      const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
      const isGraphAudience = audiences.some(
        (a: string) =>
          typeof a === "string" &&
          (a === "https://graph.microsoft.com" ||
            a === "https://graph.microsoft.com/" ||
            a === "00000003-0000-0000-c000-000000000000" ||
            a.includes("graph.microsoft.com"))
      );

      if (!isGraphAudience) {
        console.error("Invalid JWT token: Not a valid Microsoft Graph token");
        return undefined;
      }
    } catch (error) {
      console.error("Invalid JWT token: Failed to parse payload", error);
      return undefined;
    }

    return token;
  }
}
