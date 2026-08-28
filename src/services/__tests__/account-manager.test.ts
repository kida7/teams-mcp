import { promises as fs } from "node:fs";
import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  type AccountData,
  getAccount,
  getAccountsStore,
  getProfileDirForAccount,
  listAccounts,
  removeAccount,
  removeAllAccounts,
  sanitizeAccountName,
  saveAccount,
  setActiveAccount,
} from "../account-manager.js";

vi.mock("node:fs", async () => {
  const actual = (await vi.importActual("node:fs")) as any;
  return {
    ...actual,
    promises: {
      ...(actual.promises || {}),
      readFile: vi.fn(),
      writeFile: vi.fn(),
      mkdir: vi.fn(),
      rm: vi.fn(),
      unlink: vi.fn(),
    },
  };
});

describe("AccountManager", () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  describe("sanitizeAccountName", () => {
    it("should convert email or special characters into a safe directory name", () => {
      expect(sanitizeAccountName("user@company.com")).toBe("user_company_com");
      expect(sanitizeAccountName("User.Name+Test@Example.org")).toBe("user_name_test_example_org");
      expect(sanitizeAccountName("work-profile_1")).toBe("work-profile_1");
    });
  });

  describe("getProfileDirForAccount", () => {
    it("should return path containing sanitized account name", () => {
      const dir = getProfileDirForAccount("test.user@domain.com");
      expect(dir).toContain("test_user_domain_com");
    });
  });

  describe("getAccountsStore", () => {
    it("should return empty store if file not found and no legacy file", async () => {
      vi.mocked(fs.readFile).mockRejectedValue(new Error("ENOENT"));

      const store = await getAccountsStore();
      expect(store.accounts).toEqual({});
      expect(store.activeAccount).toBeUndefined();
    });

    it("should return parsed accounts when file exists", async () => {
      const existingStore = {
        activeAccount: "user1@company.com",
        accounts: {
          "user1@company.com": {
            id: "user1@company.com",
            account: "user1@company.com",
            displayName: "User One",
            targetApp: "outlook" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
        },
      };

      vi.mocked(fs.readFile).mockResolvedValueOnce(JSON.stringify(existingStore));

      const store = await getAccountsStore();
      expect(store.activeAccount).toBe("user1@company.com");
      expect(store.accounts["user1@company.com"].displayName).toBe("User One");
    });

    it("should migrate legacy ~/.msgraph-mcp-auth.json if accounts.json does not exist", async () => {
      // First read for accounts.json fails
      vi.mocked(fs.readFile)
        .mockRejectedValueOnce(new Error("ENOENT"))
        // Second read for legacy auth succeeds
        .mockResolvedValueOnce(
          JSON.stringify({
            account: "legacy@company.com",
            displayName: "Legacy User",
            targetApp: "teams",
            authMethod: "browser",
            token: "legacy-token",
            authenticated: true,
          })
        );

      const store = await getAccountsStore();
      expect(store.activeAccount).toBe("legacy@company.com");
      expect(store.accounts["legacy@company.com"]).toBeDefined();
      expect(store.accounts["legacy@company.com"].token).toBe("legacy-token");
      expect(fs.writeFile).toHaveBeenCalled();
    });
  });

  describe("saveAccount and getAccount", () => {
    it("should save a new account and make it active", async () => {
      vi.mocked(fs.readFile).mockRejectedValue(new Error("ENOENT"));

      const newAccount: AccountData = {
        id: "work",
        account: "work@company.com",
        displayName: "Work Account",
        targetApp: "outlook",
        authMethod: "browser",
        token: "tok-123",
        authenticated: true,
      };

      await saveAccount(newAccount, true);

      expect(fs.writeFile).toHaveBeenCalled();
    });

    it("should find account by ID, email, or active default", async () => {
      const existingStore = {
        activeAccount: "work",
        accounts: {
          work: {
            id: "work",
            account: "work@company.com",
            displayName: "Work Account",
            targetApp: "outlook" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
          personal: {
            id: "personal",
            account: "personal@outlook.com",
            displayName: "Personal Account",
            targetApp: "teams" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
        },
      };

      vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(existingStore));

      // Query by ID
      const byId = await getAccount("personal");
      expect(byId?.account).toBe("personal@outlook.com");

      // Query by Email
      const byEmail = await getAccount("work@company.com");
      expect(byEmail?.id).toBe("work");

      // Query default active
      const def = await getAccount();
      expect(def?.id).toBe("work");
    });
  });

  describe("setActiveAccount", () => {
    it("should switch active account in store", async () => {
      const existingStore = {
        activeAccount: "work",
        accounts: {
          work: {
            id: "work",
            account: "work@company.com",
            targetApp: "outlook" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
          personal: {
            id: "personal",
            account: "personal@outlook.com",
            targetApp: "teams" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
        },
      };

      vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(existingStore));

      const updated = await setActiveAccount("personal");
      expect(updated.id).toBe("personal");
      expect(fs.writeFile).toHaveBeenCalled();
    });

    it("should throw if account does not exist", async () => {
      vi.mocked(fs.readFile).mockResolvedValue(
        JSON.stringify({ activeAccount: undefined, accounts: {} })
      );

      await expect(setActiveAccount("nonexistent")).rejects.toThrow(/not found/);
    });
  });

  describe("removeAccount and removeAllAccounts", () => {
    it("should remove specified account and its profile dir", async () => {
      const existingStore = {
        activeAccount: "work",
        accounts: {
          work: {
            id: "work",
            account: "work@company.com",
            targetApp: "outlook" as const,
            authMethod: "browser" as const,
            profileDir: "/mock/profiles/work",
            authenticated: true,
          },
          personal: {
            id: "personal",
            account: "personal@outlook.com",
            targetApp: "teams" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
        },
      };

      vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(existingStore));

      const result = await removeAccount("work");
      expect(result).toBe(true);
      expect(fs.rm).toHaveBeenCalledWith("/mock/profiles/work", { recursive: true, force: true });
    });

    it("should remove all accounts and configs on removeAllAccounts", async () => {
      await removeAllAccounts();
      expect(fs.rm).toHaveBeenCalled();
    });
  });

  describe("listAccounts", () => {
    it("should return array of accounts with active marker", async () => {
      const existingStore = {
        activeAccount: "work",
        accounts: {
          work: {
            id: "work",
            account: "work@company.com",
            targetApp: "outlook" as const,
            authMethod: "browser" as const,
            authenticated: true,
          },
        },
      };

      vi.mocked(fs.readFile).mockResolvedValue(JSON.stringify(existingStore));

      const list = await listAccounts();
      expect(list.activeAccount).toBe("work");
      expect(list.accounts).toHaveLength(1);
    });
  });
});
