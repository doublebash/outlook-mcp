export interface Env {
  OAUTH_KV: KVNamespace;
  OUTLOOK_OAUTH_KV: KVNamespace;
  MICROSOFT_CLIENT_ID: string;
  MICROSOFT_CLIENT_SECRET: string;
  MICROSOFT_TENANT_ID: string;
  WORKER_URL: string;
  MCP_SECRET: string;
  OUTLOOK_OAUTH_CLIENT_SECRET: string;
}

export interface TokenData {
  access_token: string;
  refresh_token: string;
  expires_at: number; // Unix ms
}
