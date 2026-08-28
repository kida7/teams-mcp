import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  authenticateViaBrowser,
  isLikelyGraphToken,
  launchPersistentBrowserContext,
  parseJwt,
  promptAppSelection,
  refreshAccountTokenSilent,
} from "../browser-auth.js";

// Mock playwright
vi.mock("playwright", () => ({
  chromium: {
    launchPersistentContext: vi.fn(),
  },
}));

import { chromium } from "playwright";

describe("browser-auth", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("parseJwt", () => {
    it("should parse standard base64 and base64url JWT payloads", () => {
      const payload = {
        aud: "https://graph.microsoft.com",
        upn: "test@example.com",
        name: "Test User",
        exp: Math.floor(Date.now() / 1000) + 3600,
      };

      const base64Url = Buffer.from(JSON.stringify(payload)).toString("base64url");
      const token = `header.${base64Url}.signature`;

      const parsed = parseJwt(token);
      expect(parsed).toEqual(expect.objectContaining(payload));
    });

    it("should return null for invalid token formats", () => {
      expect(parseJwt("not-a-jwt")).toBeNull();
      expect(parseJwt("header.payload")).toBeNull();
      expect(parseJwt("header.invalid-base64-%%%.signature")).toBeNull();
      expect(parseJwt("")).toBeNull();
    });
  });

  describe("isLikelyGraphToken", () => {
    it("should return true for valid unexpired token with graph audience", () => {
      const payload = {
        aud: "https://graph.microsoft.com",
        exp: Math.floor(Date.now() / 1000) + 3600,
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;
      expect(isLikelyGraphToken(token)).toBe(true);
    });

    it("should return true for valid token with Microsoft Graph app ID GUID", () => {
      const payload = {
        aud: "00000003-0000-0000-c000-000000000000",
        exp: Math.floor(Date.now() / 1000) + 3600,
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;
      expect(isLikelyGraphToken(token)).toBe(true);
    });

    it("should return true when aud is an array containing graph.microsoft.com", () => {
      const payload = {
        aud: ["https://graph.microsoft.com", "other-aud"],
        exp: Math.floor(Date.now() / 1000) + 3600,
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;
      expect(isLikelyGraphToken(token)).toBe(true);
    });

    it("should return false for expired tokens", () => {
      const payload = {
        aud: "https://graph.microsoft.com",
        exp: Math.floor(Date.now() / 1000) - 60, // expired 1m ago
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;
      expect(isLikelyGraphToken(token)).toBe(false);
    });

    it("should return false for tokens with unrelated audience", () => {
      const payload = {
        aud: "https://api.github.com",
        exp: Math.floor(Date.now() / 1000) + 3600,
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;
      expect(isLikelyGraphToken(token)).toBe(false);
    });

    it("should return false for non-JWT strings", () => {
      expect(isLikelyGraphToken("invalid-token")).toBe(false);
    });
  });

  describe("launchPersistentBrowserContext", () => {
    it("should launch default chromium persistent context if available", async () => {
      const mockContext = { close: vi.fn() };
      vi.mocked(chromium.launchPersistentContext).mockResolvedValueOnce(mockContext as any);

      const context = await launchPersistentBrowserContext("/mock/dir", false);
      expect(context).toBe(mockContext);
      expect(chromium.launchPersistentContext).toHaveBeenCalledWith("/mock/dir", {
        headless: false,
        args: ["--start-maximized"],
        viewport: null,
      });
    });

    it("should fallback to msedge if chromium fails", async () => {
      const mockContext = { close: vi.fn() };
      vi.mocked(chromium.launchPersistentContext)
        .mockRejectedValueOnce(new Error("Chromium not found"))
        .mockResolvedValueOnce(mockContext as any);

      const context = await launchPersistentBrowserContext("/mock/dir", false);
      expect(context).toBe(mockContext);
      expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(2, "/mock/dir", {
        headless: false,
        channel: "msedge",
        args: ["--start-maximized"],
        viewport: null,
      });
    });

    it("should fallback to chrome if msedge fails", async () => {
      const mockContext = { close: vi.fn() };
      vi.mocked(chromium.launchPersistentContext)
        .mockRejectedValueOnce(new Error("Chromium not found"))
        .mockRejectedValueOnce(new Error("Edge not found"))
        .mockResolvedValueOnce(mockContext as any);

      const context = await launchPersistentBrowserContext("/mock/dir", false);
      expect(context).toBe(mockContext);
      expect(chromium.launchPersistentContext).toHaveBeenNthCalledWith(3, "/mock/dir", {
        headless: false,
        channel: "chrome",
        args: ["--start-maximized"],
        viewport: null,
      });
    });

    it("should throw a helpful error if all browser channels fail", async () => {
      vi.mocked(chromium.launchPersistentContext)
        .mockRejectedValueOnce(new Error("Chromium not found"))
        .mockRejectedValueOnce(new Error("Edge not found"))
        .mockRejectedValueOnce(new Error("Chrome not found"));

      await expect(launchPersistentBrowserContext("/mock/dir", false)).rejects.toThrow(
        /Could not launch browser/
      );
    });
  });

  describe("promptAppSelection", () => {
    it("should default to outlook when stdin is not a TTY", async () => {
      const originalIsTTY = process.stdin.isTTY;
      process.stdin.isTTY = false;
      try {
        const app = await promptAppSelection();
        expect(app).toBe("outlook");
      } finally {
        process.stdin.isTTY = originalIsTTY;
      }
    });
  });

  describe("authenticateViaBrowser", () => {
    it("should capture token from requests to graph.microsoft.com", async () => {
      const payload = {
        aud: "https://graph.microsoft.com",
        upn: "user@tenant.com",
        name: "Test User",
        scp: "User.Read Chat.Read",
        exp: Math.floor(Date.now() / 1000) + 3600,
      };
      const token = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;

      let requestListener: ((req: any) => void) | undefined;
      const mockPage = {
        goto: vi.fn().mockImplementation(async () => {
          if (requestListener) {
            requestListener({
              url: () => "https://graph.microsoft.com/v1.0/me",
              headers: () => ({
                authorization: `Bearer ${token}`,
              }),
            });
          }
        }),
      };

      const mockContext = {
        pages: vi.fn().mockReturnValue([mockPage]),
        newPage: vi.fn().mockResolvedValue(mockPage),
        on: vi.fn().mockImplementation((event, callback) => {
          if (event === "request") {
            requestListener = callback;
          }
        }),
        close: vi.fn().mockResolvedValue(undefined),
      };

      vi.mocked(chromium.launchPersistentContext).mockResolvedValueOnce(mockContext as any);

      const fs = await import("node:fs");
      vi.mocked(fs.promises.writeFile).mockResolvedValueOnce();

      const result = await authenticateViaBrowser({ app: "outlook", timeoutMs: 5000 });

      expect(result.token).toBe(token);
      expect(result.authInfo.account).toBe("test.user@example.com");
      expect(result.authInfo.authMethod).toBe("browser");
      expect(result.authInfo.targetApp).toBe("outlook");
      expect(mockContext.close).toHaveBeenCalled();
    });

    it("should handle timeout error", async () => {
      const mockPage = {
        goto: vi.fn().mockResolvedValue(undefined),
      };

      const mockContext = {
        pages: vi.fn().mockReturnValue([mockPage]),
        newPage: vi.fn().mockResolvedValue(mockPage),
        on: vi.fn(),
        close: vi.fn().mockResolvedValue(undefined),
      };

      vi.mocked(chromium.launchPersistentContext).mockResolvedValueOnce(mockContext as any);

      await expect(authenticateViaBrowser({ app: "teams", timeoutMs: 50 })).rejects.toThrow(
        /Authentication timed out/
      );
    });
  });

  describe("refreshAccountTokenSilent", () => {
    it("should launch headless persistent context and capture fresh token", async () => {
      const payload = {
        aud: "https://graph.microsoft.com",
        upn: "test@example.com",
        exp: Math.floor(Date.now() / 1000) + 7200,
      };
      const freshToken = `header.${Buffer.from(JSON.stringify(payload)).toString("base64url")}.sig`;

      let requestListener: ((req: any) => void) | undefined;
      const mockPage = {
        goto: vi.fn().mockImplementation(async () => {
          if (requestListener) {
            requestListener({
              url: () => "https://graph.microsoft.com/v1.0/me",
              headers: () => ({
                authorization: `Bearer ${freshToken}`,
              }),
            });
          }
        }),
      };

      const mockContext = {
        pages: vi.fn().mockReturnValue([mockPage]),
        newPage: vi.fn().mockResolvedValue(mockPage),
        on: vi.fn().mockImplementation((event, callback) => {
          if (event === "request") {
            requestListener = callback;
          }
        }),
        close: vi.fn().mockResolvedValue(undefined),
      };

      vi.mocked(chromium.launchPersistentContext).mockResolvedValueOnce(mockContext as any);

      const fs = await import("node:fs");
      vi.mocked(fs.promises.readFile).mockResolvedValueOnce(
        JSON.stringify({
          activeAccount: "test@example.com",
          accounts: {
            "test@example.com": {
              id: "test@example.com",
              account: "test@example.com",
              targetApp: "outlook",
              profileDir: "/mock/profile",
              authenticated: true,
            },
          },
        })
      );
      vi.mocked(fs.promises.writeFile).mockResolvedValueOnce();

      const refreshed = await refreshAccountTokenSilent("test@example.com", 5000);
      expect(refreshed).toBe(freshToken);
      expect(mockContext.close).toHaveBeenCalled();
    });
  });
});
