import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { PR_DEFERRED_SEND_TIME } from "../calendar-helpers.js";
import {
	buildMessageBody,
	buildRecipients,
	resolveAndValidateAttachments,
} from "../email-helpers.js";
import { graphDelete, graphGet, graphPatch, graphPost } from "../graph.js";
import { sanitizeDraftMessage, sanitizeEmailFull, sanitizeEmailList } from "../sanitize.js";
import type { Env } from "../types.js";
import {
	ATTACHMENTS_DESC,
	BODY_TYPE_DESC,
	attachmentSchema,
	auditLog,
	escapeOdataString,
	summariseAttachmentSources,
} from "./_shared.js";

// ── Schemas ────────────────────────────────────────────────────────────────────

const bodyTypeSchema = z.enum(["text", "html"]).optional();
// Folder is interpolated into a URL path. Restrict to well-known names or
// Graph folder IDs (alphanumeric + `_-=` for base64url padding).
const folderSchema = z
	.string()
	.regex(/^[A-Za-z0-9_=\-]+$/, "folder must be a well-known name or folder ID")
	.max(256)
	.optional();
const emailIdSchema = z.string().min(1).max(512);

// ── Helpers ────────────────────────────────────────────────────────────────────

// Build the Graph "message" object from common args. Attachments are NOT
// included here — they are uploaded separately via uploadAttachmentsToDraft
// because they may need server-side fetching (OneDrive / URL).
function buildMessageObject(args: {
	to?: string;
	cc?: string;
	bcc?: string;
	subject?: string;
	body?: string;
	body_type?: string;
}): Record<string, unknown> {
	const message: Record<string, unknown> = {};
	if (args.subject !== undefined) message.subject = args.subject;
	if (args.body !== undefined) {
		message.body = buildMessageBody(args.body, args.body_type);
	}
	if (args.to !== undefined) message.toRecipients = buildRecipients(args.to);
	const cc = buildRecipients(args.cc);
	const bcc = buildRecipients(args.bcc);
	if (cc.length) message.ccRecipients = cc;
	if (bcc.length) message.bccRecipients = bcc;
	return message;
}

async function uploadAttachmentsToDraft(
	env: Env,
	draftId: string,
	attachmentsRaw: unknown,
): Promise<void> {
	const attachments = await resolveAndValidateAttachments(env, attachmentsRaw);
	for (const att of attachments) {
		await graphPost(env, `/me/messages/${draftId}/attachments`, att);
	}
}

// ── Impls ──────────────────────────────────────────────────────────────────────

async function listEmailsImpl(
	env: Env,
	args: { count?: number; folder?: string; search?: string },
): Promise<unknown> {
	const count = Math.min(args.count ?? 20, 50);
	const folder = args.folder ?? "inbox";
	const selectFields =
		"id,conversationId,subject,from,receivedDateTime,bodyPreview,isRead,isDraft";

	if (args.search) {
		// Note: /me/messages?$search is keyword-based. We DON'T combine with
		// $orderby here because Graph requires the advanced-query opt-in for that
		// combination and behaviour varies by tenant.
		const data = (await graphGet(env, "/me/messages", {
			$search: `"${args.search.replace(/"/g, "")}"`,
			$select: selectFields,
			$top: count,
		})) as { value: unknown[] };
		return sanitizeEmailList(data.value);
	}
	const data = (await graphGet(env, `/me/mailFolders/${folder}/messages`, {
		$select: selectFields,
		$top: count,
		$orderby: "receivedDateTime desc",
	})) as { value: unknown[] };
	return sanitizeEmailList(data.value);
}

async function readEmailImpl(env: Env, args: { id: string }): Promise<unknown> {
	const data = await graphGet(env, `/me/messages/${args.id}`, {
		$select:
			"id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,isDraft",
	});
	return sanitizeEmailFull(data);
}

async function sendEmailImpl(
	env: Env,
	args: {
		to: string;
		subject: string;
		body: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
	},
): Promise<unknown> {
	// If the message has attachments we must use the create-draft-then-send
	// pattern, because /me/sendMail enforces a 4 MB total request cap. Drafts
	// can have attachments uploaded individually and are not subject to that
	// single-request cap.
	const hasAttachments = Array.isArray(args.attachments) && args.attachments.length > 0;

	if (!hasAttachments) {
		await graphPost(env, "/me/sendMail", { message: buildMessageObject(args) });
		auditLog("email_sent", { to: args.to, subject: args.subject, attachments: 0 });
		return { success: true, message: "Email sent." };
	}

	const draft = (await graphPost(env, "/me/messages", buildMessageObject(args))) as {
		id: string;
	};
	await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	await graphPost(env, `/me/messages/${draft.id}/send`);
	auditLog("email_sent", {
		to: args.to,
		subject: args.subject,
		attachments: args.attachments!.length,
		attachment_sources: summariseAttachmentSources(args.attachments),
	});
	return { success: true, message: "Email sent." };
}

type GraphRecipient = { emailAddress: { address: string } };

interface GraphReplyOriginal {
	from?: { emailAddress?: { address?: string } } | null;
	toRecipients?: Array<{ emailAddress?: { address?: string } }> | null;
	ccRecipients?: Array<{ emailAddress?: { address?: string } }> | null;
}

// Decide whether a reply needs its recipients overridden.
//
// Microsoft Graph's createReply/createReplyAll sets the draft's To: from the
// ORIGINAL message's sender. That's correct when replying to someone else, but
// when you reply to your OWN sent message the sender is you — so the reply
// bounces back to yourself instead of reaching the person you originally wrote
// to. This recomputes the recipients for that self-sent case.
//
// Returns null when no override is needed (the message was sent by someone
// else — Graph's default is right). Returns recipient lists when the owner sent
// the original, so the caller can redirect To: (and, for reply-all, CC:) to the
// original recipients with the owner removed.
export function computeSelfReplyRecipients(
	original: GraphReplyOriginal,
	ownerAddress: string,
	replyAll: boolean,
): { toRecipients: GraphRecipient[]; ccRecipients: GraphRecipient[] } | null {
	const owner = ownerAddress.trim().toLowerCase();
	const fromAddr = original.from?.emailAddress?.address?.trim().toLowerCase() ?? "";

	// Only intervene when the message being replied to was sent BY the owner.
	if (!owner || fromAddr !== owner) return null;

	const dedupeExcludingOwner = (
		list: Array<{ emailAddress?: { address?: string } }> | null | undefined,
	): GraphRecipient[] => {
		const seen = new Set<string>();
		const out: GraphRecipient[] = [];
		for (const r of list ?? []) {
			const addr = r.emailAddress?.address?.trim();
			if (!addr) continue;
			const key = addr.toLowerCase();
			if (key === owner || seen.has(key)) continue;
			seen.add(key);
			out.push({ emailAddress: { address: addr } });
		}
		return out;
	};

	return {
		toRecipients: dedupeExcludingOwner(original.toRecipients),
		ccRecipients: replyAll ? dedupeExcludingOwner(original.ccRecipients) : [],
	};
}

// The mailbox owner's primary SMTP address, used to detect replies to self.
async function getMailboxOwnerAddress(env: Env): Promise<string | null> {
	const me = (await graphGet(env, "/me", {
		$select: "mail,userPrincipalName",
	})) as { mail?: string | null; userPrincipalName?: string | null };
	return me.mail ?? me.userPrincipalName ?? null;
}

async function replyToEmailImpl(
	env: Env,
	args: {
		id: string;
		body: string;
		body_type?: string;
		reply_all?: boolean;
		attachments?: unknown[];
	},
): Promise<unknown> {
	const action = args.reply_all ? "createReplyAll" : "createReply";

	// Fetch the original sender/recipients and the mailbox owner alongside
	// creating the reply draft, so we can fix the reply-to-self bounce below.
	const [original, owner, draft] = await Promise.all([
		graphGet(env, `/me/messages/${args.id}`, {
			$select: "from,toRecipients,ccRecipients",
		}) as Promise<GraphReplyOriginal>,
		getMailboxOwnerAddress(env),
		graphPost(env, `/me/messages/${args.id}/${action}`) as Promise<{ id: string }>,
	]);

	const updates: Record<string, unknown> = {
		body: buildMessageBody(args.body, args.body_type),
	};

	const override = owner
		? computeSelfReplyRecipients(original, owner, !!args.reply_all)
		: null;
	if (override) {
		if (override.toRecipients.length === 0) {
			// Replying to your own message that had no other recipient. Refuse
			// rather than silently emailing yourself; clean up the orphan draft.
			try {
				await graphDelete(env, `/me/messages/${draft.id}`);
			} catch {
				/* leave the orphan draft if cleanup fails */
			}
			throw ToolError.validation(
				"This message was sent by you and has no other recipient to reply to. Use send_email to start a new message instead.",
			);
		}
		updates.toRecipients = override.toRecipients;
		updates.ccRecipients = override.ccRecipients;
	}

	await graphPatch(env, `/me/messages/${draft.id}`, updates);

	if (args.attachments && args.attachments.length > 0) {
		await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	}

	await graphPost(env, `/me/messages/${draft.id}/send`);
	auditLog("reply_sent", {
		in_reply_to: args.id,
		reply_all: !!args.reply_all,
		self_reply_redirected: !!override,
		attachments: args.attachments?.length ?? 0,
		attachment_sources: summariseAttachmentSources(args.attachments),
	});
	return { success: true, message: "Reply sent." };
}

async function forwardEmailImpl(
	env: Env,
	args: {
		id: string;
		to: string;
		body?: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
	},
): Promise<unknown> {
	const draft = (await graphPost(env, `/me/messages/${args.id}/createForward`)) as {
		id: string;
	};

	const updates: Record<string, unknown> = {
		toRecipients: buildRecipients(args.to),
	};
	const cc = buildRecipients(args.cc);
	const bcc = buildRecipients(args.bcc);
	if (cc.length) updates.ccRecipients = cc;
	if (bcc.length) updates.bccRecipients = bcc;
	if (args.body) {
		updates.body = buildMessageBody(args.body, args.body_type);
	}
	await graphPatch(env, `/me/messages/${draft.id}`, updates);

	if (args.attachments && args.attachments.length > 0) {
		await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	}

	await graphPost(env, `/me/messages/${draft.id}/send`);
	auditLog("email_forwarded", {
		forwarded_from: args.id,
		to: args.to,
		additional_attachments: args.attachments?.length ?? 0,
		attachment_sources: summariseAttachmentSources(args.attachments),
	});
	return { success: true, message: "Email forwarded." };
}

// ── Threaded draft replies/forwards ─────────────────────────────────────────────
// Unlike reply_to_email / forward_email (which create a draft via Graph's
// createReply/createReplyAll/createForward and immediately send it), these tools
// leave the generated draft in Drafts so the user can review and send it later
// from Outlook — while still preserving the original conversation thread.
//
// We POST to the create* endpoint with NO body, then PATCH body/recipients
// separately. Graph rejects createReply when both a comment AND a body are
// supplied in the same call (HTTP 400), so the two-step approach is required.

const DRAFT_META_SELECT =
	"id,subject,conversationId,webLink,toRecipients,ccRecipients,createdDateTime,lastModifiedDateTime";

// Map a failed create*-draft Graph call to a clear "not found" message when the
// original message can't be reached; otherwise pass the (already friendly)
// Graph error through unchanged.
function mapCreateDraftError(e: unknown, kind: "reply" | "forward"): unknown {
	const isMissing =
		e instanceof ToolError &&
		(e.status === 404 ||
			(typeof e.internalMessage === "string" &&
				/ErrorItemNotFound|not found/i.test(e.internalMessage)));
	if (isMissing) {
		return ToolError.validation(
			kind === "forward"
				? "Could not create forward draft: message not found or not accessible."
				: "Could not create reply draft: message not found or not accessible.",
		);
	}
	return e;
}

// Apply optional updates to the freshly-created draft and return its metadata.
// If the PATCH fails we DON'T throw — the draft already exists, so we return its
// id plus a warning so the user never loses the generated draft. If the final
// metadata read also fails we still return the id.
async function finaliseThreadDraft(
	env: Env,
	draftId: string,
	updates: Record<string, unknown>,
	audit: { event: string; details: Record<string, unknown> },
): Promise<Record<string, unknown>> {
	let patchError: string | null = null;
	if (Object.keys(updates).length > 0) {
		try {
			await graphPatch(env, `/me/messages/${draftId}`, updates);
		} catch (e) {
			patchError = e instanceof Error ? e.message : String(e);
		}
	}

	auditLog(audit.event, {
		...audit.details,
		draft_id: draftId,
		patch_failed: !!patchError,
	});

	let meta: Record<string, unknown>;
	try {
		const raw = await graphGet(env, `/me/messages/${draftId}`, {
			$select: DRAFT_META_SELECT,
		});
		meta = sanitizeDraftMessage(raw) as unknown as Record<string, unknown>;
	} catch {
		meta = { id: draftId };
	}

	if (patchError) {
		return {
			success: false,
			draft_id: draftId,
			...meta,
			warning:
				"The draft reply was created in your Drafts, but updating its content failed. You can still edit and send it in Outlook, or retry with update_draft.",
			error: patchError,
		};
	}

	return { success: true, message: "Draft created.", draft_id: draftId, ...meta };
}

export async function createReplyDraftImpl(
	env: Env,
	args: { id: string; body?: string; body_type?: string; reply_all?: boolean },
): Promise<unknown> {
	const action = args.reply_all ? "createReplyAll" : "createReply";

	let draft: { id: string };
	try {
		draft = (await graphPost(env, `/me/messages/${args.id}/${action}`)) as {
			id: string;
		};
	} catch (e) {
		throw mapCreateDraftError(e, "reply");
	}

	// Self-reply redirect: Graph sets the draft's To: from the ORIGINAL sender,
	// so replying to your OWN sent message bounces back to yourself. Recompute
	// recipients exactly as reply_to_email does.
	const [original, owner] = await Promise.all([
		graphGet(env, `/me/messages/${args.id}`, {
			$select: "from,toRecipients,ccRecipients",
		}) as Promise<GraphReplyOriginal>,
		getMailboxOwnerAddress(env),
	]);

	const updates: Record<string, unknown> = {};
	if (args.body !== undefined) {
		updates.body = buildMessageBody(args.body, args.body_type);
	}

	const override = owner
		? computeSelfReplyRecipients(original, owner, !!args.reply_all)
		: null;
	if (override) {
		if (override.toRecipients.length === 0) {
			// Replying to your own message that had no other recipient. Refuse
			// rather than drafting a mail to yourself; clean up the orphan draft.
			try {
				await graphDelete(env, `/me/messages/${draft.id}`);
			} catch {
				/* leave the orphan draft if cleanup fails */
			}
			throw ToolError.validation(
				"This message was sent by you and has no other recipient to reply to. Use create_draft to start a new message instead.",
			);
		}
		updates.toRecipients = override.toRecipients;
		updates.ccRecipients = override.ccRecipients;
	}

	return finaliseThreadDraft(env, draft.id, updates, {
		event: "reply_draft_created",
		details: {
			in_reply_to: args.id,
			reply_all: !!args.reply_all,
			self_reply_redirected: !!override,
		},
	});
}

export async function createForwardDraftImpl(
	env: Env,
	args: {
		id: string;
		to?: string;
		cc?: string;
		bcc?: string;
		body?: string;
		body_type?: string;
	},
): Promise<unknown> {
	let draft: { id: string };
	try {
		draft = (await graphPost(env, `/me/messages/${args.id}/createForward`)) as {
			id: string;
		};
	} catch (e) {
		throw mapCreateDraftError(e, "forward");
	}

	const updates: Record<string, unknown> = {};
	if (args.to !== undefined) updates.toRecipients = buildRecipients(args.to);
	const cc = buildRecipients(args.cc);
	const bcc = buildRecipients(args.bcc);
	if (cc.length) updates.ccRecipients = cc;
	if (bcc.length) updates.bccRecipients = bcc;
	if (args.body !== undefined) {
		updates.body = buildMessageBody(args.body, args.body_type);
	}

	return finaliseThreadDraft(env, draft.id, updates, {
		event: "forward_draft_created",
		details: { forwarded_from: args.id },
	});
}

async function deleteEmailImpl(env: Env, args: { id: string }): Promise<unknown> {
	await graphDelete(env, `/me/messages/${args.id}`);
	return { success: true, message: "Email deleted." };
}

async function moveEmailImpl(
	env: Env,
	args: { id: string; destination_folder: string },
): Promise<unknown> {
	const data = (await graphPost(env, `/me/messages/${args.id}/move`, {
		destinationId: args.destination_folder,
	})) as { id: string };
	return { success: true, message: "Email moved.", new_id: data.id };
}

async function createDraftImpl(
	env: Env,
	args: {
		to?: string;
		subject: string;
		body: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
	},
): Promise<unknown> {
	const draft = (await graphPost(env, "/me/messages", buildMessageObject(args))) as {
		id: string;
		webLink?: string;
	};

	if (args.attachments && args.attachments.length > 0) {
		await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	}

	auditLog("draft_created", {
		draft_id: draft.id,
		to: args.to,
		subject: args.subject,
		attachments: args.attachments?.length ?? 0,
		attachment_sources: summariseAttachmentSources(args.attachments),
	});
	return {
		success: true,
		message: "Draft created.",
		draft_id: draft.id,
		webLink: draft.webLink,
	};
}

async function updateDraftImpl(
	env: Env,
	args: {
		id: string;
		to?: string;
		subject?: string;
		body?: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
	},
): Promise<unknown> {
	const updates = buildMessageObject(args);
	let didChange = false;

	if (Object.keys(updates).length > 0) {
		await graphPatch(env, `/me/messages/${args.id}`, updates);
		didChange = true;
	}

	if (args.attachments !== undefined) {
		const existing = (await graphGet(
			env,
			`/me/messages/${args.id}/attachments`,
			{ $select: "id" },
		)) as { value: Array<{ id: string }> };
		for (const a of existing.value) {
			await graphDelete(env, `/me/messages/${args.id}/attachments/${a.id}`);
		}
		if (args.attachments.length > 0) {
			await uploadAttachmentsToDraft(env, args.id, args.attachments);
		}
		didChange = true;
	}

	if (!didChange) {
		throw ToolError.validation(
			"No fields provided to update. Pass at least one of: subject, body, to, cc, bcc, attachments.",
		);
	}

	auditLog("draft_updated", {
		draft_id: args.id,
		fields_updated: Object.keys(updates),
		attachments_replaced: args.attachments !== undefined,
	});
	return { success: true, message: "Draft updated." };
}

async function sendDraftImpl(env: Env, args: { id: string }): Promise<unknown> {
	await graphPost(env, `/me/messages/${args.id}/send`);
	auditLog("draft_sent", { draft_id: args.id });
	return { success: true, message: "Draft sent." };
}

async function scheduleSendImpl(
	env: Env,
	args: {
		to: string;
		subject: string;
		body: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
		send_at: string;
	},
): Promise<unknown> {
	const sendAtDate = new Date(args.send_at);
	if (Number.isNaN(sendAtDate.getTime())) {
		throw ToolError.validation("send_at is not a valid ISO 8601 datetime");
	}
	if (sendAtDate.getTime() <= Date.now()) {
		throw ToolError.validation("send_at must be in the future");
	}

	const message = buildMessageObject(args);
	message.singleValueExtendedProperties = [
		{ id: PR_DEFERRED_SEND_TIME, value: sendAtDate.toISOString() },
	];

	const draft = (await graphPost(env, "/me/messages", message)) as { id: string };

	if (args.attachments && args.attachments.length > 0) {
		await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	}

	await graphPost(env, `/me/messages/${draft.id}/send`);
	auditLog("scheduled_send", {
		draft_id: draft.id,
		send_at: sendAtDate.toISOString(),
		to: args.to,
		subject: args.subject,
		attachments: args.attachments?.length ?? 0,
		attachment_sources: summariseAttachmentSources(args.attachments),
	});
	return {
		success: true,
		message: `Email scheduled for ${sendAtDate.toISOString()}. The mail server will hold it until then.`,
		scheduled_id: draft.id,
	};
}

async function getConversationImpl(
	env: Env,
	args: { conversation_id?: string; message_id?: string; count?: number },
): Promise<unknown> {
	let conversationId = args.conversation_id;
	if (!conversationId) {
		if (!args.message_id) {
			throw ToolError.validation("Either conversation_id or message_id is required");
		}
		const msg = (await graphGet(env, `/me/messages/${args.message_id}`, {
			$select: "conversationId",
		})) as { conversationId?: string };
		if (!msg.conversationId)
			throw ToolError.validation("Could not resolve conversationId from message");
		conversationId = msg.conversationId;
	}

	const count = args.count ?? 50;
	const safeId = escapeOdataString(conversationId);
	// `ConsistencyLevel: eventual` enables server-side $orderby combined with
	// $filter — the advanced-query opt-in Microsoft Graph requires.
	const data = (await graphGet(
		env,
		"/me/messages",
		{
			$filter: `conversationId eq '${safeId}'`,
			$select:
				"id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,isDraft",
			$top: count,
			$orderby: "receivedDateTime asc",
		},
		{ ConsistencyLevel: "eventual" },
	)) as { value: unknown[] };

	return {
		conversationId,
		count: data.value.length,
		messages: data.value.map((m) => sanitizeEmailFull(m)),
	};
}

// ── defineTools map ────────────────────────────────────────────────────────────

export const emailTools = defineTools<Env>({
	list_emails: {
		description:
			"Retrieve recent emails. Supports folder selection and keyword search.",
		schema: z.object({
			count: z.number().int().min(1).max(50).optional(),
			folder: folderSchema,
			search: z.string().min(1).max(256).optional(),
		}),
		handler: (env, args) => listEmailsImpl(env, args),
	},
	read_email: {
		description: "Read the full content of an email by its ID.",
		schema: z.object({ id: emailIdSchema }),
		handler: (env, args) => readEmailImpl(env, args),
	},
	send_email: {
		description:
			"Send a new email. Supports plain text or HTML body, attachments, and inline images.",
		schema: z.object({
			to: z.string().min(1).max(2048),
			subject: z.string().min(1).max(998),
			body: z.string().max(2 * 1024 * 1024),
			body_type: bodyTypeSchema,
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
		}),
		handler: (env, args) => sendEmailImpl(env, args),
	},
	reply_to_email: {
		description:
			"Reply to an email. Supports plain text or HTML, attachments, and inline images.",
		schema: z.object({
			id: emailIdSchema,
			body: z.string().max(2 * 1024 * 1024),
			body_type: bodyTypeSchema,
			reply_all: z.boolean().optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
		}),
		handler: (env, args) => replyToEmailImpl(env, args),
	},
	forward_email: {
		description:
			"Forward an email to new recipients. The original message and its attachments are included automatically; you can add an optional comment above and additional attachments.",
		schema: z.object({
			id: emailIdSchema,
			to: z.string().min(1).max(2048),
			body: z.string().max(2 * 1024 * 1024).optional(),
			body_type: bodyTypeSchema,
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
		}),
		handler: (env, args) => forwardEmailImpl(env, args),
	},
	create_reply_draft: {
		description:
			"Create a saved Outlook draft reply INSIDE an existing email thread (saved to Drafts, NOT sent). Use this when the user wants to review or manually send a reply later while preserving the original conversation. Requires the original message ID. To send a reply immediately instead, use reply_to_email. The returned id can be passed to send_draft. Attachments are not supported by this tool yet — add them with update_draft on the returned draft id.",
		schema: z.object({
			id: emailIdSchema,
			body: z.string().max(2 * 1024 * 1024).optional(),
			body_type: bodyTypeSchema,
		}),
		handler: (env, args) => createReplyDraftImpl(env, args),
	},
	create_reply_all_draft: {
		description:
			"Create a saved Outlook draft reply-all INSIDE an existing email thread (saved to Drafts, NOT sent). Use this when the original thread has multiple recipients and the user wants to review or manually send the reply later. Graph preserves the original To/CC recipients. Requires the original message ID. To send immediately instead, use reply_to_email with reply_all. The returned id can be passed to send_draft. Attachments are not supported by this tool yet — add them with update_draft on the returned draft id.",
		schema: z.object({
			id: emailIdSchema,
			body: z.string().max(2 * 1024 * 1024).optional(),
			body_type: bodyTypeSchema,
		}),
		handler: (env, args) => createReplyDraftImpl(env, { ...args, reply_all: true }),
	},
	create_forward_draft: {
		description:
			"Create a saved Outlook draft forward from an existing email (saved to Drafts, NOT sent). The original message is included automatically; you can add an optional comment above it (body) and set recipients (to/cc/bcc). Use this when the user wants to review or manually send a forwarded message later. To forward and send immediately instead, use forward_email. The returned id can be passed to send_draft. Attachments are not supported by this tool yet — add them with update_draft on the returned draft id.",
		schema: z.object({
			id: emailIdSchema,
			to: z.string().max(2048).optional(),
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			body: z.string().max(2 * 1024 * 1024).optional(),
			body_type: bodyTypeSchema,
		}),
		handler: (env, args) => createForwardDraftImpl(env, args),
	},
	delete_email: {
		description: "Delete an email (moves it to the Deleted Items folder).",
		schema: z.object({ id: emailIdSchema }),
		handler: (env, args) => deleteEmailImpl(env, args),
	},
	move_email: {
		description: "Move an email to a different folder.",
		schema: z.object({
			id: emailIdSchema,
			destination_folder: z
				.string()
				.regex(/^[A-Za-z0-9_=\-]+$/, "folder must be a well-known name or folder ID")
				.max(256),
		}),
		handler: (env, args) => moveEmailImpl(env, args),
	},
	create_draft: {
		description:
			"Create a new email draft (saved but not sent). Returns the draft ID for later editing or sending.",
		schema: z.object({
			to: z.string().min(1).max(2048).optional(),
			subject: z.string().min(1).max(998),
			body: z.string().max(2 * 1024 * 1024),
			body_type: bodyTypeSchema,
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
		}),
		handler: (env, args) => createDraftImpl(env, args),
	},
	update_draft: {
		description:
			"Update an existing draft email. Only provide the fields you want to change. To replace attachments, provide the full new attachment list (existing attachments will be removed first). Empty calls (no fields provided) are rejected.",
		schema: z.object({
			id: emailIdSchema,
			to: z.string().max(2048).optional(),
			subject: z.string().max(998).optional(),
			body: z.string().max(2 * 1024 * 1024).optional(),
			body_type: bodyTypeSchema,
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
		}),
		handler: (env, args) => updateDraftImpl(env, args),
	},
	send_draft: {
		description: "Send an existing draft email by its ID.",
		schema: z.object({ id: emailIdSchema }),
		handler: (env, args) => sendDraftImpl(env, args),
	},
	schedule_send: {
		description:
			"Send an email at a future date/time. The Microsoft 365 mail server holds the message until the scheduled time. Same options as send_email plus a send_at datetime.",
		schema: z.object({
			to: z.string().min(1).max(2048),
			subject: z.string().min(1).max(998),
			body: z.string().max(2 * 1024 * 1024),
			body_type: bodyTypeSchema,
			cc: z.string().max(2048).optional(),
			bcc: z.string().max(2048).optional(),
			attachments: z.array(attachmentSchema).max(20).optional(),
			send_at: z.string().min(1).max(64),
		}),
		handler: (env, args) => scheduleSendImpl(env, args),
	},
	get_conversation: {
		description:
			"Get all messages in an email thread (conversation), oldest first. Returns messages from across all folders (inbox, sent, archive, etc.). Pass either a conversation_id or a message_id from any email in the thread.",
		schema: z
			.object({
				conversation_id: z.string().min(1).max(512).optional(),
				message_id: z.string().min(1).max(512).optional(),
				count: z.number().int().min(1).max(200).optional(),
			})
			.refine((d) => !!(d.conversation_id || d.message_id), {
				message: "Provide either conversation_id or message_id",
			}),
		handler: (env, args) => getConversationImpl(env, args),
	},
});

// Re-export tool descriptions so the tool file is self-contained for tests.
export { ATTACHMENTS_DESC, BODY_TYPE_DESC };
