import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { listAccounts } from "../services/account-manager.js";
import type { GraphService } from "../services/graph.js";

export function registerAuthTools(
  server: McpServer,
  graphService: GraphService,
  _readOnly: boolean
) {
  // Authentication status tool
  server.registerTool(
    "auth_status",
    {
      title: "Auth Status",
      description:
        "Check the authentication status of the Microsoft Graph connection. Returns current active account details, expiration, auto-refresh status, and available accounts.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async () => {
      const status = await graphService.getAuthStatus();

      if (!status.isAuthenticated) {
        let msg =
          "❌ Not authenticated. Please run: npx @kida7/teams-mcp@latest authenticate --browser";
        if (status.availableAccounts && status.availableAccounts.length > 0) {
          msg += `\n\nAvailable saved accounts: ${status.availableAccounts.map((a) => a.account).join(", ")}. You can select one via switch_account.`;
        }
        return {
          content: [{ type: "text", text: msg }],
        };
      }

      let infoText = `✅ Authenticated as ${status.displayName || "Unknown User"} (${status.userPrincipalName || "No email available"})`;
      if (status.accountId) {
        infoText += `\n🆔 Account ID: ${status.accountId}`;
      }
      if (status.targetApp) {
        infoText += `\n🌐 Target App: ${status.targetApp === "teams" ? "Microsoft Teams Web" : "Outlook Web"}`;
      }
      if (status.expiresAt) {
        infoText += `\n⏰ Access token expires: ${new Date(status.expiresAt).toLocaleString()}`;
      }
      infoText += `\n🔄 Persistent Silent Auto-Refresh: ${status.autoRefresh ? "Enabled (Headless)" : "Disabled"}`;

      if (status.availableAccounts && status.availableAccounts.length > 1) {
        infoText += `\n\n👥 Other available accounts (${status.availableAccounts.length}):`;
        for (const acc of status.availableAccounts) {
          const isCurrent = acc.id === status.accountId;
          infoText += `\n  ${isCurrent ? "👉 [ACTIVE] " : "   "} ${acc.account} (${acc.displayName || "No Name"}) [ID: ${acc.id}]`;
        }
      }

      return {
        content: [{ type: "text", text: infoText }],
      };
    }
  );

  // List accounts tool
  server.registerTool(
    "list_accounts",
    {
      title: "List Accounts",
      description: "List all authenticated Microsoft 365 accounts and their status.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async () => {
      const store = await listAccounts();
      if (store.accounts.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No accounts found. Add an account using: npx @kida7/teams-mcp@latest authenticate --browser",
            },
          ],
        };
      }

      let text = "👥 Authenticated Microsoft 365 Accounts:\n";
      for (const acc of store.accounts) {
        const isActive = acc.id === store.activeAccount;
        const autoRefresh = acc.profileDir ? "Auto-refresh: ON" : "Auto-refresh: OFF";
        const expires = acc.expiresAt
          ? `Expires: ${new Date(acc.expiresAt).toLocaleTimeString()}`
          : "";
        text += `\n${isActive ? "👉 [ACTIVE] " : "   • "} ${acc.displayName || acc.account} (${acc.account})`;
        text += `\n     ID: ${acc.id} | App: ${acc.targetApp} | ${autoRefresh} ${expires ? `| ${expires}` : ""}\n`;
      }

      return {
        content: [{ type: "text", text: text.trim() }],
      };
    }
  );

  // Switch account tool
  server.registerTool(
    "switch_account",
    {
      title: "Switch Account",
      description: "Switch the active Microsoft 365 account used by the MCP server.",
      inputSchema: {
        accountId: z
          .string()
          .describe("The account ID, alias, or email of the account to switch to"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    async ({ accountId }) => {
      try {
        const switched = await graphService.switchAccount(accountId);
        return {
          content: [
            {
              type: "text",
              text: `✅ Switched active account to: ${switched.displayName || switched.account} (${switched.account}) [ID: ${switched.id}]`,
            },
          ],
        };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text", text: `❌ Failed to switch account: ${message}` }],
        };
      }
    }
  );
}
