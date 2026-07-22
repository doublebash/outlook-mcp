import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the toolkit's proxied-OAuth store so each failure mode can be driven
// deterministically. These are the exact strings the toolkit throws
// (node_modules/@bashco/mcp-toolkit/src/upstreamAuth/proxiedOAuth.ts) — if a
// toolkit upgrade changes them, these tests fail loudly rather than letting
// getValidAccessToken silently fall back to the generic message.
const storeMock = {
	saveTokens: vi.fn(),
	getTokens: vi.fn(),
	getValidAccessToken: vi.fn(),
	clearTokens: vi.fn(),
};

vi.mock("@bashco/mcp-toolkit", () => ({
	createProxiedOAuthStore: () => storeMock,
}));

import { getValidAccessToken, saveTokens, clearTokens } from "../src/auth.js";

const MARKER_KEY = "oauth_tokens:connected";

function makeEnv() {
	const kv = new Map<string, string>();
	return {
		env: {
			OAUTH_KV: {
				get: vi.fn(async (k: string) => kv.get(k) ?? null),
				put: vi.fn(async (k: string, v: string) => {
					kv.set(k, v);
				}),
				delete: vi.fn(async (k: string) => {
					kv.delete(k);
				}),
			},
			MCP_APPROVAL_CODE: "test-code",
			MICROSOFT_TENANT_ID: "common",
			MICROSOFT_CLIENT_ID: "client-id",
			MICROSOFT_CLIENT_SECRET: "client-secret",
		} as never,
		kv,
	};
}

beforeEach(() => {
	vi.clearAllMocks();
});

describe("getValidAccessToken — failure attribution", () => {
	it("returns the access token on the happy path", async () => {
		const { env } = makeEnv();
		storeMock.getValidAccessToken.mockResolvedValue("access-abc");

		await expect(getValidAccessToken(env)).resolves.toBe("access-abc");
	});

	it("costs no extra KV reads on the happy path", async () => {
		const { env } = makeEnv();
		storeMock.getValidAccessToken.mockResolvedValue("access-abc");

		await getValidAccessToken(env);

		// getValidAccessToken runs on every Graph call (graphFetch), so the
		// diagnostic marker must only be read when something has already failed.
		expect((env as never as ReturnType<typeof makeEnv>["env"]).OAUTH_KV.get).not.toHaveBeenCalled();
	});

	it("reports 'never connected' when no marker exists", async () => {
		const { env } = makeEnv();
		storeMock.getValidAccessToken.mockRejectedValue(new Error("upstream not authenticated"));

		await expect(getValidAccessToken(env)).rejects.toThrow(/^Not authenticated\./);
	});

	it("distinguishes undecryptable tokens from 'never connected' when a marker exists", async () => {
		const { env, kv } = makeEnv();
		kv.set(
			MARKER_KEY,
			JSON.stringify({ connected_at: "2026-07-01T00:00:00.000Z", namespace: "outlook-mcp-v2" }),
		);
		// The toolkit swallows a decrypt error, deletes the ciphertext, and
		// surfaces it as plain "not authenticated" — identical to never having
		// connected. The marker is what tells them apart.
		storeMock.getValidAccessToken.mockRejectedValue(new Error("upstream not authenticated"));

		await expect(getValidAccessToken(env)).rejects.toThrow(
			/could no longer be decrypted/,
		);
		await expect(getValidAccessToken(env)).rejects.toThrow(/MCP_APPROVAL_CODE changed/);
		await expect(getValidAccessToken(env)).rejects.toThrow(/2026-07-01/);
	});

	it("surfaces the HTTP status when Microsoft rejects the refresh", async () => {
		const { env, kv } = makeEnv();
		kv.set(MARKER_KEY, JSON.stringify({ connected_at: "x", namespace: "outlook-mcp-v2" }));
		storeMock.getValidAccessToken.mockRejectedValue(new Error("token refresh failed: 400"));

		await expect(getValidAccessToken(env)).rejects.toThrow(
			/Microsoft rejected the token refresh \(HTTP 400\)/,
		);
	});

	it("does not misreport a refresh rejection as a decryption failure", async () => {
		const { env, kv } = makeEnv();
		kv.set(MARKER_KEY, JSON.stringify({ connected_at: "x", namespace: "outlook-mcp-v2" }));
		storeMock.getValidAccessToken.mockRejectedValue(new Error("token refresh failed: 401"));

		await expect(getValidAccessToken(env)).rejects.not.toThrow(/decrypted/);
	});

	it("reports lock contention as transient rather than an auth failure", async () => {
		const { env } = makeEnv();
		storeMock.getValidAccessToken.mockRejectedValue(
			new Error("token refresh: could not acquire lock"),
		);

		await expect(getValidAccessToken(env)).rejects.toThrow(/Retry shortly/);
		await expect(getValidAccessToken(env)).rejects.not.toThrow(/oauth\/start/);
	});
});

describe("connection marker lifecycle", () => {
	it("writes a marker on save, holding no secret material", async () => {
		const { env, kv } = makeEnv();
		storeMock.saveTokens.mockResolvedValue(undefined);

		await saveTokens(env, {
			access_token: "a",
			refresh_token: "r",
			expires_at: 1,
		} as never);

		const raw = kv.get(MARKER_KEY);
		expect(raw).toBeDefined();
		expect(JSON.parse(raw!)).toMatchObject({ namespace: "outlook-mcp-v2" });
		// The marker sits in KV unencrypted, so it must never carry the approval
		// code, the tokens, or anything derived from them.
		expect(raw).not.toContain("test-code");
		expect(raw).not.toContain("refresh");
	});

	it("does not fail the connect when the marker write fails", async () => {
		const { env } = makeEnv();
		storeMock.saveTokens.mockResolvedValue(undefined);
		vi.mocked(env.OAUTH_KV.put).mockRejectedValue(new Error("KV down"));

		await expect(
			saveTokens(env, { access_token: "a", refresh_token: "r", expires_at: 1 } as never),
		).resolves.toBeUndefined();
	});

	it("clears the marker on disconnect so a reconnect starts clean", async () => {
		const { env, kv } = makeEnv();
		storeMock.saveTokens.mockResolvedValue(undefined);
		storeMock.clearTokens.mockResolvedValue(undefined);

		await saveTokens(env, {
			access_token: "a",
			refresh_token: "r",
			expires_at: 1,
		} as never);
		expect(kv.get(MARKER_KEY)).toBeDefined();

		await clearTokens(env);
		expect(kv.get(MARKER_KEY)).toBeUndefined();

		// And with the marker gone, a later failure reads as "never connected".
		storeMock.getValidAccessToken.mockRejectedValue(new Error("upstream not authenticated"));
		await expect(getValidAccessToken(env)).rejects.toThrow(/^Not authenticated\./);
	});
});
