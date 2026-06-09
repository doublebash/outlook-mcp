import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPost, graphPutBinary } from "../graph.js";
import { sanitizeFileList } from "../sanitize.js";
import type { Env } from "../types.js";
import { encodeOneDrivePath } from "./_shared.js";

// OneDrive paths: alphanumerics, spaces, dots, dashes, underscores, slashes,
// parentheses, apostrophes, and a few common safe punctuation marks. Disallows
// query/fragment chars (`?`, `#`, `&`, `=`) and control chars to prevent path
// injection. Path is encoded per-segment before interpolation regardless.
const pathSchema = z
	.string()
	.min(1)
	.max(1024)
	.regex(
		/^[\w \-.,()'!&+@$\/]+$/,
		"path may contain letters, numbers, spaces, and . - _ , ( ) ' ! & + @ $ /",
	);

// ── Upload limits + MIME allowlist ────────────────────────────────────────────

// Microsoft Graph's "simple upload" (PUT to :/content) supports up to 4 MB.
// We cap slightly below to leave headroom for transport overhead.
export const MAX_UPLOAD_BYTES = 4 * 1024 * 1024;

// Uploads go to the user's own OneDrive (no third-party delivery), so the
// allowlist is broader than the email-attachment one — it includes plain-text
// formats the caller is likely to need (vCard, iCalendar, Markdown, XML). It
// still rejects executables and dangerous MIME types.
export const UPLOAD_ALLOWED_MIME_TYPES = new Set([
	// Documents
	"application/pdf",
	"application/msword",
	"application/vnd.ms-excel",
	"application/vnd.ms-powerpoint",
	"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
	"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
	"application/vnd.openxmlformats-officedocument.presentationml.presentation",
	// Images (excludes SVG — XML, can carry <script>)
	"image/png",
	"image/jpeg",
	"image/jpg",
	"image/gif",
	"image/webp",
	// Text formats
	"text/plain",
	"text/csv",
	"text/html",
	"text/markdown",
	"text/vcard",
	"text/x-vcard",
	"text/calendar",
	"application/json",
	"application/xml",
	"text/xml",
	// Archives
	"application/zip",
]);

// Executable extensions blocked regardless of declared MIME type.
const DANGEROUS_FILENAME_EXT = /\.(exe|bat|cmd|scr|msi|dll|ps1|vbs|com|cpl|jar|app)$/i;

// ── Pure helpers (exported for unit tests) ────────────────────────────────────

// Decode standard base64 to bytes. Throws ToolError on invalid input rather
// than letting atob's DOMException leak out as an unhandled error.
export function base64ToBytes(b64: string): Uint8Array {
	let binary: string;
	try {
		binary = atob(b64);
	} catch {
		throw ToolError.validation("content_base64 is not valid base64");
	}
	const bytes = new Uint8Array(binary.length);
	for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
	return bytes;
}

// ── Tool implementations ──────────────────────────────────────────────────────

async function listFilesImpl(
	env: Env,
	args: { folder_path?: string; count?: number },
): Promise<unknown> {
	const count = args.count ?? 25;
	const basePath = args.folder_path
		? `/me/drive/root:/${encodeOneDrivePath(args.folder_path)}:/children`
		: `/me/drive/root/children`;

	const data = (await graphGet(env, basePath, {
		$select: "id,name,size,file,folder,lastModifiedDateTime,webUrl",
		$top: count,
	})) as { value: unknown[] };
	return sanitizeFileList(data.value);
}

async function shareFileImpl(
	env: Env,
	args: {
		item_path: string;
		link_type?: "view" | "edit";
		scope?: "anonymous" | "organization";
	},
): Promise<unknown> {
	const data = (await graphPost(
		env,
		`/me/drive/root:/${encodeOneDrivePath(args.item_path)}:/createLink`,
		{ type: args.link_type ?? "view", scope: args.scope ?? "anonymous" },
	)) as { link: { webUrl: string; type: string; scope: string } };

	return {
		success: true,
		url: data.link.webUrl,
		type: data.link.type,
		scope: data.link.scope,
	};
}

async function uploadOneDriveFileImpl(
	env: Env,
	args: {
		item_path: string;
		content_base64: string;
		content_type: string;
		conflict_behavior?: "rename" | "replace" | "fail";
	},
): Promise<unknown> {
	const mime = args.content_type.toLowerCase();
	if (!UPLOAD_ALLOWED_MIME_TYPES.has(mime)) {
		throw ToolError.validation(
			`MIME type not allowed: ${mime}. Allowed types: ${Array.from(
				UPLOAD_ALLOWED_MIME_TYPES,
			).join(", ")}`,
		);
	}

	const bytes = base64ToBytes(args.content_base64);
	if (bytes.byteLength === 0) {
		throw ToolError.validation("upload body is empty");
	}
	if (bytes.byteLength > MAX_UPLOAD_BYTES) {
		throw ToolError.validation(
			`File is ${(bytes.byteLength / 1024 / 1024).toFixed(1)} MB. Max is ${
				MAX_UPLOAD_BYTES / 1024 / 1024
			} MB for simple upload.`,
		);
	}
	if (DANGEROUS_FILENAME_EXT.test(args.item_path)) {
		throw ToolError.validation(`Filename extension not allowed: ${args.item_path}`);
	}

	const encodedPath = encodeOneDrivePath(args.item_path);
	// `@microsoft.graph.conflictBehavior` is a Graph-defined query param.
	// URLSearchParams handles the URL-encoding of the `@` and `.` correctly.
	const query = new URLSearchParams({
		"@microsoft.graph.conflictBehavior": args.conflict_behavior ?? "replace",
	});
	const path = `/me/drive/root:/${encodedPath}:/content?${query.toString()}`;

	const data = (await graphPutBinary(env, path, bytes, mime)) as {
		id?: string;
		name?: string;
		size?: number;
		webUrl?: string;
		file?: { mimeType?: string };
		folder?: unknown;
		lastModifiedDateTime?: string;
	};

	return { success: true, file: sanitizeFileList([data])[0] };
}

// ── Tool definitions ──────────────────────────────────────────────────────────

export const filesTools = defineTools<Env>({
	list_files: {
		description: "List files and folders in OneDrive.",
		schema: z.object({
			folder_path: pathSchema.optional(),
			count: z.number().int().min(1).max(200).optional(),
		}),
		handler: (env, args) => listFilesImpl(env, args),
	},

	share_file: {
		description:
			'Create a shareable link for a file or folder in OneDrive. Default scope is "anonymous" — anyone with the link can access. Use "organization" to restrict to org members only.',
		schema: z.object({
			item_path: pathSchema,
			link_type: z.enum(["view", "edit"]).optional(),
			scope: z.enum(["anonymous", "organization"]).optional(),
		}),
		handler: (env, args) => shareFileImpl(env, args),
	},

	upload_onedrive_file: {
		description:
			'Upload a new file to OneDrive via Microsoft Graph simple upload. Pass file bytes as base64 in content_base64 and the MIME type in content_type. Max 4 MB. conflict_behavior defaults to "replace"; use "rename" to auto-suffix `(1)` on collision or "fail" to error. Returns the uploaded driveItem metadata (id, name, size, webUrl) so the caller can immediately share it via share_file — useful for any host-then-share workflow (e.g. generating a .vcf in your own folder and returning a download link).',
		schema: z.object({
			item_path: pathSchema,
			content_base64: z.string().min(1).max(8 * 1024 * 1024),
			content_type: z.string().min(1).max(128),
			conflict_behavior: z.enum(["rename", "replace", "fail"]).optional(),
		}),
		handler: (env, args) => uploadOneDriveFileImpl(env, args),
	},
});
