import { TOOLS, callTool } from './tools';
import type { ToolArgs } from './tools';
import type { Env } from './types';

// ── JSON-RPC types ─────────────────────────────────────────────────────────────

interface JsonRpcRequest {
  jsonrpc: '2.0';
  id?: string | number | null;
  method: string;
  params?: Record<string, unknown>;
}

function ok(id: string | number | null | undefined, result: unknown): Response {
  return Response.json({ jsonrpc: '2.0', id: id ?? null, result });
}

function err(id: string | number | null | undefined, code: number, message: string): Response {
  return Response.json({ jsonrpc: '2.0', id: id ?? null, error: { code, message } });
}

// ── Timing-safe comparison ────────────────────────────────────────────────────
// Prevents token enumeration via timing attacks on the hot MCP path.
// Uses HMAC signatures so comparison is always constant-time regardless
// of where the strings differ.

async function timingSafeEqual(a: string, b: string): Promise<boolean> {
  const encoder = new TextEncoder();
  // A zero key is fine here — we care about constant-time comparison, not HMAC security.
  const key = await crypto.subtle.importKey(
    'raw',
    new Uint8Array(32),
    { name: 'HMAC', hash: 'SHA-256' },
    false,
    ['sign']
  );
  const [sigA, sigB] = await Promise.all([
    crypto.subtle.sign('HMAC', key, encoder.encode(a)),
    crypto.subtle.sign('HMAC', key, encoder.encode(b)),
  ]);
  const bytesA = new Uint8Array(sigA);
  const bytesB = new Uint8Array(sigB);
  if (bytesA.length !== bytesB.length) return false;
  let diff = 0;
  for (let i = 0; i < bytesA.length; i++) diff |= bytesA[i]! ^ bytesB[i]!;
  return diff === 0;
}

// ── MCP request handler ────────────────────────────────────────────────────────

export async function handleMCP(request: Request, env: Env): Promise<Response> {
  // Accept either:
  //   X-MCP-Secret: <secret>          — Claude Desktop via proxy.mjs
  //   Authorization: Bearer <secret>  — Claude.ai web via OAuth token
  const mcpSecret = request.headers.get('X-MCP-Secret');
  const authHeader = request.headers.get('Authorization');
  const bearerToken = authHeader?.startsWith('Bearer ') ? authHeader.slice(7) : null;

  const bearerOk = bearerToken !== null && await timingSafeEqual(bearerToken, env.MCP_SECRET);
  const secretOk = mcpSecret !== null && await timingSafeEqual(mcpSecret, env.MCP_SECRET);
  const authenticated = bearerOk || secretOk;

  if (!authenticated) {
    return new Response('Unauthorised', {
      status: 401,
      headers: {
        'WWW-Authenticate':
          `Bearer realm="Outlook MCP Server", resource_metadata="${env.WORKER_URL}/.well-known/oauth-protected-resource"`,
      },
    });
  }

  let body: JsonRpcRequest;

  try {
    body = await request.json() as JsonRpcRequest;
  } catch {
    return err(null, -32700, 'Parse error');
  }

  const { id, method, params } = body;

  // Notifications (no id) don't need a response — return 202 Accepted
  if (id === undefined && method.startsWith('notifications/')) {
    return new Response(null, { status: 202 });
  }

  try {
    switch (method) {
      case 'ping':
        return ok(id, {});

      case 'initialize': {
        const clientVersion = (params?.protocolVersion as string | undefined) ?? '2025-03-26';
        const negotiated = clientVersion >= '2025-03-26' ? '2025-03-26' : '2024-11-05';
        const initResponse = ok(id, {
          protocolVersion: negotiated,
          capabilities: { tools: {} },
          serverInfo: { name: 'outlook-mcp', version: '1.0.0' },
        });
        const sessionId = crypto.randomUUID();
        const r = new Response(initResponse.body, initResponse);
        r.headers.set('Mcp-Session-Id', sessionId);
        return r;
      }

      case 'tools/list':
        return ok(id, { tools: TOOLS });

      case 'tools/call': {
        const toolName = params?.name as string | undefined;
        const toolArgs = (params?.arguments ?? {}) as ToolArgs;

        if (!toolName) {
          return err(id, -32602, 'Missing tool name');
        }

        const result = await callTool(toolName, toolArgs, env);

        return ok(id, {
          content: [{ type: 'text', text: result }],
          isError: false,
        });
      }

      default:
        return err(id, -32601, `Method not found: ${method}`);
    }
  } catch (e) {
    const message = e instanceof Error ? e.message : String(e);
    return ok(id, {
      content: [{ type: 'text', text: `Error: ${message}` }],
      isError: true,
    });
  }
}
