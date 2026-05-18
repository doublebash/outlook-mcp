import { defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPost } from "../graph.js";
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
});
