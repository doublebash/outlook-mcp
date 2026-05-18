import type { CloudflareRateLimiter } from "@bashco/mcp-toolkit";

export interface Env {
	OAUTH_KV: KVNamespace;
	MICROSOFT_CLIENT_ID: string;
	MICROSOFT_CLIENT_SECRET: string;
	MICROSOFT_TENANT_ID: string;
	WORKER_URL: string;
	/**
	 * Operator-only approval code. Gates `/approve` and `/oauth/start`.
	 * Also the HKDF-derived encryption secret for Microsoft tokens stored in KV.
	 * NEVER returned to clients. NEVER used as a runtime bearer.
	 */
	MCP_APPROVAL_CODE: string;

	RATE_LIMIT_APPROVE: CloudflareRateLimiter;
	RATE_LIMIT_TOKEN: CloudflareRateLimiter;
	RATE_LIMIT_REGISTER: CloudflareRateLimiter;
	RATE_LIMIT_MCP: CloudflareRateLimiter;
	RATE_LIMIT_OAUTH_START: CloudflareRateLimiter;
}

/** Microsoft token shape stored encrypted at rest. */
export interface OutlookTokenData {
	access_token: string;
	refresh_token: string;
	expires_at: number; // Unix ms
}
