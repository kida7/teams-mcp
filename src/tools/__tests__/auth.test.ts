import { beforeEach, describe, expect, it, vi } from "vitest";
import { listAccounts } from "../../services/account-manager.js";
import {
  createMockGraphService,
  createMockMcpServer,
  createMockUnauthenticatedGraphService,
} from "../../test-utils/setup.js";
import { registerAuthTools } from "../auth.js";

// Mock account-manager
vi.mock("../../services/account-manager.js", () => ({
  listAccounts: vi.fn(),
  getAccount: vi.fn(),
  setActiveAccount: vi.fn(),
  saveAccount: vi.fn(),
}));

describe("Authentication Tools", () => {
  let mockServer: any;
  let mockGraphService: any;

  beforeEach(() => {
    mockServer = createMockMcpServer();
    vi.clearAllMocks();

    vi.mocked(listAccounts).mockResolvedValue({
      activeAccount: "user1@company.com",
      accounts: [
        {
          id: "user1@company.com",
          account: "user1@company.com",
          displayName: "User One",
          targetApp: "outlook",
          profileDir: "/mock/profile/dir",
          authMethod: "browser",
          authenticated: true,
          expiresAt: new Date(Date.now() + 3600000).toISOString(),
        },
      ],
    });
  });

  describe("auth_status tool", () => {
    it("should register auth_status tool correctly", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      expect(mockServer.registerTool).toHaveBeenCalledWith(
        "auth_status",
        expect.objectContaining({
          title: "Auth Status",
          description: expect.stringContaining("Check the authentication status"),
          inputSchema: {},
          annotations: expect.objectContaining({
            readOnlyHint: true,
            destructiveHint: false,
          }),
        }),
        expect.any(Function)
      );
    });

    it("should return authenticated status when user is authenticated", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result.content[0].text).toContain(
        "✅ Authenticated as Test User (test.user@example.com)"
      );
      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });

    it("should return unauthenticated status when user is not authenticated", async () => {
      mockGraphService = createMockUnauthenticatedGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      const authTool = mockServer.getTool("auth_status");
      const result = await authTool.handler();

      expect(result.content[0].text).toContain("❌ Not authenticated");
      expect(mockGraphService.getAuthStatus).toHaveBeenCalledTimes(1);
    });
  });

  describe("list_accounts tool", () => {
    it("should list all configured accounts", async () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      const listTool = mockServer.getTool("list_accounts");
      const result = await listTool.handler();

      expect(result.content[0].text).toContain("user1@company.com");
      expect(result.content[0].text).toContain("[ACTIVE]");
    });
  });

  describe("switch_account tool", () => {
    it("should switch account and return confirmation", async () => {
      mockGraphService = {
        ...createMockGraphService(),
        switchAccount: vi.fn().mockResolvedValue({
          id: "user2@company.com",
          account: "user2@company.com",
          displayName: "User Two",
        }),
      };
      registerAuthTools(mockServer, mockGraphService, false);

      const switchTool = mockServer.getTool("switch_account");
      const result = await switchTool.handler({ accountId: "user2@company.com" });

      expect(result.content[0].text).toContain("✅ Switched active account to: User Two");
      expect(mockGraphService.switchAccount).toHaveBeenCalledWith("user2@company.com");
    });
  });

  describe("tool registration", () => {
    it("should register all expected authentication tools", () => {
      mockGraphService = createMockGraphService();
      registerAuthTools(mockServer, mockGraphService, false);

      const registeredTools = mockServer.getAllTools();
      expect(registeredTools).toContain("auth_status");
      expect(registeredTools).toContain("list_accounts");
      expect(registeredTools).toContain("switch_account");
      expect(registeredTools).toHaveLength(3);
    });
  });
});
