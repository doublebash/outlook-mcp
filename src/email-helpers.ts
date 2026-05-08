// Shared helpers for email composition: HTML sanitisation, attachment
// validation, and message building. Centralising these means every code path
// that sends or drafts mail goes through the same security checks.

import { graphRequestRaw } from './graph';

// ── Limits ─────────────────────────────────────────────────────────────────────

export const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;       // 3 MB per file
export const MAX_TOTAL_ATTACHMENT_BYTES = 10 * 1024 * 1024; // 10 MB per email

// MIME allowlist. Anything not on this list is rejected. Keep tight.
const ALLOWED_MIME_TYPES = new Set([
  // Documents
  'application/pdf',
  'application/msword',
  'application/vnd.ms-excel',
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  // Images
  // Note: image/svg+xml deliberately excluded. SVG is XML and can carry
  // <script> tags that execute when rendered; it is a known XSS vector.
  // Use PNG/JPG/WebP for inline images instead.
  'image/png',
  'image/jpeg',
  'image/jpg',
  'image/gif',
  'image/webp',
  // Text
  'text/plain',
  'text/csv',
  'text/html',
  'application/json',
  // Archives
  'application/zip',
]);

// ── Outbound HTML safety pass ──────────────────────────────────────────────────
// Strips script-execution vectors before mail leaves the account.
// This is defence-in-depth: it does not replace the recipient's mail-client
// sanitiser, but it ensures we never use this account to deliver weaponised
// HTML even if Claude is prompt-injected.

export function sanitizeOutboundHtml(html: string): string {
  return html
    // Remove entire <script>, <iframe>, <object>, <embed>, <form>, <link>, <meta>, <base> blocks
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/<object[\s\S]*?<\/object>/gi, '')
    .replace(/<embed[^>]*>/gi, '')
    .replace(/<form[\s\S]*?<\/form>/gi, '')
    .replace(/<link[^>]*>/gi, '')
    .replace(/<meta[^>]*>/gi, '')
    .replace(/<base[^>]*>/gi, '')
    // Strip on*= event handler attributes (onclick, onload, onerror, etc.)
    .replace(/\son[a-z]+\s*=\s*"[^"]*"/gi, '')
    .replace(/\son[a-z]+\s*=\s*'[^']*'/gi, '')
    .replace(/\son[a-z]+\s*=\s*[^\s>]+/gi, '')
    // Neutralise javascript: URLs
    .replace(/javascript\s*:/gi, 'blocked:')
    // Strip data: URLs except for inline images we explicitly allow
    .replace(/\sdata:(?!image\/(png|jpeg|jpg|gif|webp);base64,)/gi, ' blocked:');
}

// ── Attachment validation ──────────────────────────────────────────────────────

// An attachment can specify ONE of three sources:
//   - content_base64: raw base64 bytes inline (Claude has the bytes already)
//   - onedrive_path:  fetched server-side from the user's OneDrive
//   - url:            fetched server-side from a public HTTPS URL
// content_type is auto-detected for the latter two if omitted.
export interface AttachmentInput {
  name: string;
  content_base64?: string;
  onedrive_path?: string;
  url?: string;
  content_type?: string;
  is_inline?: boolean;
  content_id?: string;
}

interface GraphAttachment {
  '@odata.type': string;
  name: string;
  contentBytes: string;
  contentType: string;
  isInline?: boolean;
  contentId?: string;
}

export function validateAndBuildAttachments(
  attachments: unknown
): GraphAttachment[] {
  if (!attachments) return [];
  if (!Array.isArray(attachments)) {
    throw new Error('attachments must be an array');
  }

  const built: GraphAttachment[] = [];
  let totalBytes = 0;

  for (const raw of attachments) {
    const a = raw as AttachmentInput;
    if (!a.name || !a.content_base64 || !a.content_type) {
      throw new Error('Each attachment requires name, content_base64, and content_type');
    }

    // MIME check
    const mime = a.content_type.toLowerCase();
    if (!ALLOWED_MIME_TYPES.has(mime)) {
      throw new Error(`MIME type not allowed: ${mime}. Allowed types: ${Array.from(ALLOWED_MIME_TYPES).join(', ')}`);
    }

    // Filename extension defence-in-depth — block executable extensions
    // even if the MIME type lies.
    const dangerousExt = /\.(exe|bat|cmd|scr|msi|dll|ps1|vbs|com|cpl|jar|app)$/i;
    if (dangerousExt.test(a.name)) {
      throw new Error(`Filename extension not allowed: ${a.name}`);
    }

    // Size check (base64 expands ~33% — derive raw size)
    const rawBytes = Math.floor((a.content_base64.length * 3) / 4);
    if (rawBytes > MAX_ATTACHMENT_BYTES) {
      throw new Error(`Attachment "${a.name}" is ${(rawBytes / 1024 / 1024).toFixed(1)} MB. Max is ${MAX_ATTACHMENT_BYTES / 1024 / 1024} MB per file.`);
    }
    totalBytes += rawBytes;
    if (totalBytes > MAX_TOTAL_ATTACHMENT_BYTES) {
      throw new Error(`Total attachments exceed ${MAX_TOTAL_ATTACHMENT_BYTES / 1024 / 1024} MB.`);
    }

    // Inline image must have a content_id so the HTML body's cid: reference can resolve
    if (a.is_inline && !a.content_id) {
      throw new Error(`Inline attachment "${a.name}" requires a content_id (referenced in HTML as cid:THE_ID)`);
    }

    built.push({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: a.name,
      contentBytes: a.content_base64,
      contentType: a.content_type,
      ...(a.is_inline && { isInline: true, contentId: a.content_id }),
    });
  }

  return built;
}

// ── Message body builder ───────────────────────────────────────────────────────

export function buildMessageBody(
  body: string,
  bodyType: string | undefined
): { contentType: 'Text' | 'HTML'; content: string } {
  const isHtml = (bodyType ?? 'text').toLowerCase() === 'html';
  return {
    contentType: isHtml ? 'HTML' : 'Text',
    content: isHtml ? sanitizeOutboundHtml(body) : body,
  };
}

// ── Recipient list builder ─────────────────────────────────────────────────────

export function buildRecipients(csv: string | undefined): Array<{ emailAddress: { address: string } }> {
  if (!csv) return [];
  return csv.split(',').map(e => ({ emailAddress: { address: e.trim() } })).filter(r => r.emailAddress.address);
}

// ── Server-side attachment fetching (SSRF-safe) ────────────────────────────────

const URL_FETCH_TIMEOUT_MS = 10_000;

// Block private/internal/loopback addresses to prevent SSRF.
// Note: this only catches URL-literal IPs, not DNS rebinding. Cloudflare
// Workers' fetch already refuses to connect to private ranges as additional
// defence, but we still validate up-front for clearer errors and to reject
// hostnames like "localhost".
function rejectIfPrivateHost(hostname: string): void {
  const lower = hostname.toLowerCase();
  if (lower === 'localhost' || lower.endsWith('.localhost') || lower.endsWith('.internal') || lower.endsWith('.local')) {
    throw new Error(`Refusing to fetch from internal hostname: ${hostname}`);
  }
  // IPv4 literal checks
  const ipv4 = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(lower);
  if (ipv4) {
    const [a, b] = [parseInt(ipv4[1]), parseInt(ipv4[2])];
    const isPrivate =
      a === 10 ||                          // 10.0.0.0/8
      a === 127 ||                         // loopback
      a === 0 ||                           // 0.0.0.0/8
      (a === 169 && b === 254) ||          // link-local / metadata
      (a === 172 && b >= 16 && b <= 31) || // 172.16.0.0/12
      (a === 192 && b === 168) ||          // 192.168.0.0/16
      a >= 224;                            // multicast / reserved
    if (isPrivate) throw new Error(`Refusing to fetch from private IP: ${hostname}`);
  }
  // IPv6 literal checks (Cloudflare URL might wrap in [])
  const stripped = lower.replace(/^\[|\]$/g, '');
  if (stripped === '::1' || stripped.startsWith('fc') || stripped.startsWith('fd') || stripped.startsWith('fe80:') || stripped.startsWith('::ffff:')) {
    throw new Error(`Refusing to fetch from private IPv6: ${hostname}`);
  }
}

function validateExternalUrl(rawUrl: string): URL {
  let parsed: URL;
  try {
    parsed = new URL(rawUrl);
  } catch {
    throw new Error(`Invalid URL: ${rawUrl}`);
  }
  if (parsed.protocol !== 'https:') {
    throw new Error(`Only https:// URLs are allowed (got ${parsed.protocol})`);
  }
  rejectIfPrivateHost(parsed.hostname);
  return parsed;
}

// Convert ArrayBuffer to base64. Chunked to avoid stack overflow on
// String.fromCharCode.apply with large arrays.
function arrayBufferToBase64(buf: ArrayBuffer): string {
  const bytes = new Uint8Array(buf);
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(chunk));
  }
  return btoa(binary);
}

// Read a Response body, enforcing a hard size cap mid-stream so a malicious
// or runaway server can't drain Worker memory.
async function readBodyWithCap(response: Response, maxBytes: number): Promise<ArrayBuffer> {
  // Sanity-check Content-Length up front when present
  const cl = response.headers.get('content-length');
  if (cl && parseInt(cl, 10) > maxBytes) {
    throw new Error(`Remote file is ${(parseInt(cl, 10) / 1024 / 1024).toFixed(1)} MB. Max is ${maxBytes / 1024 / 1024} MB.`);
  }

  const reader = response.body?.getReader();
  if (!reader) {
    // Fallback: response without a stream — load directly with implicit cap from Workers
    const buf = await response.arrayBuffer();
    if (buf.byteLength > maxBytes) {
      throw new Error(`Remote file exceeds ${maxBytes / 1024 / 1024} MB cap`);
    }
    return buf;
  }

  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    total += value.byteLength;
    if (total > maxBytes) {
      try { await reader.cancel(); } catch { /* ignore */ }
      throw new Error(`Remote file exceeds ${maxBytes / 1024 / 1024} MB cap`);
    }
    chunks.push(value);
  }

  const out = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) { out.set(c, offset); offset += c.byteLength; }
  return out.buffer;
}

// Fetch a public HTTPS URL with SSRF protection, timeout, and size cap.
// Redirects are handled manually so each hop's URL is re-validated against
// the SSRF blocklist — preventing a public URL that 302s to a private IP.
const MAX_REDIRECTS = 3;

async function fetchAsBase64FromUrl(rawUrl: string): Promise<{ content_base64: string; content_type: string }> {
  let currentUrl = rawUrl;
  validateExternalUrl(currentUrl);

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), URL_FETCH_TIMEOUT_MS);

  try {
    let response: Response | null = null;
    for (let hop = 0; hop <= MAX_REDIRECTS; hop++) {
      response = await fetch(currentUrl, {
        method: 'GET',
        redirect: 'manual',
        signal: controller.signal,
      });

      // Manual redirect handling: 301/302/303/307/308 → re-validate target
      if ([301, 302, 303, 307, 308].includes(response.status)) {
        if (hop === MAX_REDIRECTS) {
          throw new Error(`Too many redirects (max ${MAX_REDIRECTS})`);
        }
        const location = response.headers.get('location');
        if (!location) throw new Error('Redirect response had no Location header');
        // Resolve relative redirects against the current URL
        const next = new URL(location, currentUrl).toString();
        validateExternalUrl(next);
        currentUrl = next;
        continue;
      }

      // Not a redirect — break out and process the response
      break;
    }

    if (!response) throw new Error('No response received');
    if (!response.ok) {
      throw new Error(`URL fetch failed: ${response.status} ${response.statusText}`);
    }
    const contentType = (response.headers.get('content-type') ?? 'application/octet-stream').split(';')[0].trim();
    const buf = await readBodyWithCap(response, MAX_ATTACHMENT_BYTES);
    return { content_base64: arrayBufferToBase64(buf), content_type: contentType };
  } finally {
    clearTimeout(timer);
  }
}

// Fetch a file from the user's OneDrive by path, returning base64 + MIME.
async function fetchAsBase64FromOneDrive(
  token: string,
  path: string
): Promise<{ content_base64: string; content_type: string }> {
  // Strip any leading slash for the Graph URL template
  const clean = path.replace(/^\/+/, '');
  const response = await graphRequestRaw(token, `/me/drive/root:/${encodeURI(clean)}:/content`);
  const contentType = (response.headers.get('content-type') ?? 'application/octet-stream').split(';')[0].trim();
  const buf = await readBodyWithCap(response, MAX_ATTACHMENT_BYTES);
  return { content_base64: arrayBufferToBase64(buf), content_type: contentType };
}

// ── Attachment resolver ────────────────────────────────────────────────────────
// Async because OneDrive/URL sources require fetches. Returns Graph-ready
// attachment objects (after security validation).

export async function resolveAndValidateAttachments(
  token: string,
  raw: unknown
): Promise<GraphAttachment[]> {
  if (!raw) return [];
  if (!Array.isArray(raw)) throw new Error('attachments must be an array');

  // First pass: resolve every source-bearing input to a concrete content_base64.
  const resolved: AttachmentInput[] = [];
  for (const item of raw) {
    const a = item as AttachmentInput;
    if (!a.name) throw new Error('Each attachment requires a "name" field');

    const sources = [a.content_base64, a.onedrive_path, a.url].filter(Boolean).length;
    if (sources === 0) {
      throw new Error(`Attachment "${a.name}" needs one of: content_base64, onedrive_path, url`);
    }
    if (sources > 1) {
      throw new Error(`Attachment "${a.name}" specifies multiple sources — pick one of content_base64, onedrive_path, url`);
    }

    if (a.content_base64) {
      if (!a.content_type) throw new Error(`Attachment "${a.name}" with content_base64 requires content_type`);
      resolved.push(a);
    } else if (a.onedrive_path) {
      const fetched = await fetchAsBase64FromOneDrive(token, a.onedrive_path);
      resolved.push({
        ...a,
        content_base64: fetched.content_base64,
        content_type: a.content_type ?? fetched.content_type,
      });
    } else if (a.url) {
      const fetched = await fetchAsBase64FromUrl(a.url);
      resolved.push({
        ...a,
        content_base64: fetched.content_base64,
        content_type: a.content_type ?? fetched.content_type,
      });
    }
  }

  // Second pass: existing strict validation (size, MIME allowlist, extension blocklist).
  return validateAndBuildAttachments(resolved);
}
