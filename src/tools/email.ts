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
import {
	buildBodyWithSignature,
	findQuoteBoundary,
	insertBeforeQuote,
	maybeSign,
	renderBodyHtml,
	resolveSignature,
} from "../signature.js";
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

// Opt-in server-side signature. Default false so existing callers are
// unaffected. The signature itself lives in deployment config (SIGNATURE_HTML),
// never in this repo — see src/signature.ts.
const includeSignatureSchema = z.boolean().optional().default(false);
const INCLUDE_SIGNATURE_DESC =
	"Append the sender's configured email signature server-side. Do NOT write the signature into `body` yourself — set this flag and the Worker adds it verbatim. Forces the body to HTML (a signature in a plain-text body renders as raw markup). No-op if the deployment has no signature configured. Default false.";

// ── Helpers ────────────────────────────────────────────────────────────────────

// Build the Graph "message" object from common args. Attachments are NOT
// included here — they are uploaded separately via uploadAttachmentsToDraft
// because they may need server-side fetching (OneDrive / URL).
function buildMessageObject(
	args: {
		to?: string;
		cc?: string;
		bcc?: string;
		subject?: string;
		body?: string;
		body_type?: string;
		include_signature?: boolean;
	},
	env?: Env,
	notes?: string[],
): Record<string, unknown> {
	const message: Record<string, unknown> = {};
	if (args.subject !== undefined) message.subject = args.subject;

	// Signature path replaces the normal body build: the signature must be
	// concatenated after sanitisation (see src/signature.ts), so it cannot go
	// through buildMessageBody.
	const signed = env ? maybeSign(env, args.include_signature, args.body, args.body_type) : null;
	if (signed) {
		message.body = { contentType: signed.contentType, content: signed.content };
		notes?.push(...signed.notes);
	} else if (args.body !== undefined) {
		message.body = buildMessageBody(args.body, args.body_type);
	}
	if (args.to !== undefined) message.toRecipients = buildRecipients(args.to);
	const cc = buildRecipients(args.cc);
	const bcc = buildRecipients(args.bcc);
	if (cc.length) message.ccRecipients = cc;
	if (bcc.length) message.bccRecipients = bcc;
	return message;
}

// Attach signature-related notes (e.g. a body_type override) to a tool result
// so the caller sees what the Worker changed instead of it happening silently.
function withNotes(
	result: Record<string, unknown>,
	notes: string[],
): Record<string, unknown> {
	const unique = [...new Set(notes)];
	return unique.length > 0 ? { ...result, notes: unique } : result;
}

/**
 * Compose the outgoing body for a reply/forward draft that Graph has already
 * generated.
 *
 * Graph's createReply/createReplyAll/createForward draft already contains the
 * quoted original. PATCHing `body` with just the author's new text REPLACES
 * that whole document, which silently discards the conversation history — the
 * recipient gets a bare reply with no "On <date>, X wrote:" beneath it.
 *
 * So instead of replacing, we read the generated body back and insert the new
 * content ABOVE the quote boundary, leaving the quote intact below. This is
 * why replies are always HTML: the generated quote block is HTML, and a Text
 * body cannot carry it.
 *
 * Returns undefined when there is nothing to write (no body, no signature), so
 * the caller leaves Graph's draft untouched.
 */
async function composeThreadBody(
	env: Env,
	draftId: string,
	args: { body?: string; body_type?: string; include_signature?: boolean },
	notes: string[],
): Promise<{ contentType: "HTML"; content: string } | undefined> {
	const signature = args.include_signature ? resolveSignature(env) : null;

	// The author's own content, rendered and (when asked) signed.
	let top: string;
	if (signature) {
		const built = buildBodyWithSignature(args.body, args.body_type, signature);
		notes.push(...built.notes);
		top = built.content;
	} else {
		if (args.body === undefined) return undefined;
		top = renderBodyHtml(args.body, args.body_type);
	}

	if (args.body !== undefined && (args.body_type ?? "text").toLowerCase() !== "html") {
		notes.push(
			"Sent as HTML so the quoted original could be preserved below your message.",
		);
	}

	let generated = "";
	try {
		const raw = (await graphGet(env, `/me/messages/${draftId}`, { $select: "body" })) as {
			body?: { content?: string } | null;
		};
		generated = raw.body?.content ?? "";
	} catch {
		// Reading the draft failed. Fall back to the author's content alone
		// rather than failing the send, but say so — the quote will be missing.
		notes.push(
			"Could not read the generated draft, so the quoted original is not included below your message.",
		);
		return { contentType: "HTML", content: top };
	}

	if (generated.trim() === "") return { contentType: "HTML", content: top };

	if (findQuoteBoundary(generated) !== null) {
		return { contentType: "HTML", content: insertBeforeQuote(generated, top) };
	}

	// No recognisable boundary: treat the whole generated document as the quote
	// and place the new content at the top of it, inside <body> when present.
	const bodyOpen = /<body\b[^>]*>/i.exec(generated);
	if (bodyOpen) {
		const at = bodyOpen.index + bodyOpen[0].length;
		return {
			contentType: "HTML",
			content: `${generated.slice(0, at)}${top}<br><br>${generated.slice(at)}`,
		};
	}
	return { contentType: "HTML", content: `${top}<br><br>${generated}` };
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	// If the message has attachments we must use the create-draft-then-send
	// pattern, because /me/sendMail enforces a 4 MB total request cap. Drafts
	// can have attachments uploaded individually and are not subject to that
	// single-request cap.
	const hasAttachments = Array.isArray(args.attachments) && args.attachments.length > 0;
	const notes: string[] = [];

	if (!hasAttachments) {
		await graphPost(env, "/me/sendMail", { message: buildMessageObject(args, env, notes) });
		auditLog("email_sent", {
			to: args.to,
			subject: args.subject,
			attachments: 0,
			signed: !!args.include_signature,
		});
		return withNotes({ success: true, message: "Email sent." }, notes);
	}

	const draft = (await graphPost(env, "/me/messages", buildMessageObject(args, env, notes))) as {
		id: string;
	};
	await uploadAttachmentsToDraft(env, draft.id, args.attachments);
	await graphPost(env, `/me/messages/${draft.id}/send`);
	auditLog("email_sent", {
		to: args.to,
		subject: args.subject,
		attachments: args.attachments!.length,
		attachment_sources: summariseAttachmentSources(args.attachments),
		signed: !!args.include_signature,
	});
	return withNotes({ success: true, message: "Email sent." }, notes);
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	const action = args.reply_all ? "createReplyAll" : "createReply";
	const notes: string[] = [];

	// Fetch the original sender/recipients and the mailbox owner alongside
	// creating the reply draft, so we can fix the reply-to-self bounce below.
	const [original, owner, draft] = await Promise.all([
		graphGet(env, `/me/messages/${args.id}`, {
			$select: "from,toRecipients,ccRecipients",
		}) as Promise<GraphReplyOriginal>,
		getMailboxOwnerAddress(env),
		graphPost(env, `/me/messages/${args.id}/${action}`) as Promise<{ id: string }>,
	]);

	const updates: Record<string, unknown> = {};
	const composed = await composeThreadBody(env, draft.id, args, notes);
	if (composed) updates.body = composed;

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
		signed: !!args.include_signature,
	});
	return withNotes({ success: true, message: "Reply sent." }, notes);
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	const notes: string[] = [];
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

	const composed = await composeThreadBody(env, draft.id, args, notes);
	if (composed) updates.body = composed;
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
		signed: !!args.include_signature,
	});
	return withNotes({ success: true, message: "Email forwarded." }, notes);
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
	args: {
		id: string;
		body?: string;
		body_type?: string;
		reply_all?: boolean;
		include_signature?: boolean;
	},
): Promise<unknown> {
	const action = args.reply_all ? "createReplyAll" : "createReply";
	const notes: string[] = [];

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
	const composed = await composeThreadBody(env, draft.id, args, notes);
	if (composed) updates.body = composed;

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

	const result = await finaliseThreadDraft(env, draft.id, updates, {
		event: "reply_draft_created",
		details: {
			in_reply_to: args.id,
			reply_all: !!args.reply_all,
			self_reply_redirected: !!override,
			signed: !!args.include_signature,
		},
	});
	return withNotes(result, notes);
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	const notes: string[] = [];
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

	const composed = await composeThreadBody(env, draft.id, args, notes);
	if (composed) updates.body = composed;

	const result = await finaliseThreadDraft(env, draft.id, updates, {
		event: "forward_draft_created",
		details: { forwarded_from: args.id, signed: !!args.include_signature },
	});
	return withNotes(result, notes);
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

export async function createDraftImpl(
	env: Env,
	args: {
		to?: string;
		subject: string;
		body: string;
		body_type?: string;
		cc?: string;
		bcc?: string;
		attachments?: unknown[];
		include_signature?: boolean;
	},
): Promise<unknown> {
	// Inject the signature at draft-creation time (not at send). A standalone
	// draft is reviewed and then sent via send_draft, which never touches the
	// body — so create time is the only point the signature can be added, and it
	// means the user sees the signed body while reviewing.
	const notes: string[] = [];
	const draft = (await graphPost(env, "/me/messages", buildMessageObject(args, env, notes))) as {
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
		signed: !!args.include_signature,
	});
	return withNotes(
		{
			success: true,
			message: "Draft created.",
			draft_id: draft.id,
			webLink: draft.webLink,
		},
		notes,
	);
}

export async function updateDraftImpl(
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	const notes: string[] = [];

	// Only sign when a new body is actually supplied. update_draft is a partial
	// update ("only provide the fields you want to change"), so signing with no
	// body would build a signature-only body and wipe whatever the draft held.
	const signBody = args.include_signature === true && args.body !== undefined;
	if (args.include_signature === true && args.body === undefined) {
		notes.push(
			"include_signature was ignored: it only applies when you also provide a new body. " +
				"The existing draft body was left unchanged.",
		);
	}
	const updates = buildMessageObject(args, signBody ? env : undefined, notes);
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
		signed: signBody,
	});
	return withNotes({ success: true, message: "Draft updated." }, notes);
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
		include_signature?: boolean;
	},
): Promise<unknown> {
	const sendAtDate = new Date(args.send_at);
	if (Number.isNaN(sendAtDate.getTime())) {
		throw ToolError.validation("send_at is not a valid ISO 8601 datetime");
	}
	if (sendAtDate.getTime() <= Date.now()) {
		throw ToolError.validation("send_at must be in the future");
	}

	const notes: string[] = [];
	const message = buildMessageObject(args, env, notes);
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
		signed: !!args.include_signature,
	});
	return withNotes(
		{
			success: true,
			message: `Email scheduled for ${sendAtDate.toISOString()}. The mail server will hold it until then.`,
			scheduled_id: draft.id,
		},
		notes,
	);
}

// ── Conversation lookup ───────────────────────────────────────────────────────

const CONVERSATION_DEFAULT_COUNT = 50;
const CONVERSATION_MAX_COUNT = 200;
const CONVERSATION_SELECT =
	"id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,isDraft";

interface ConversationMessage {
	receivedDateTime?: string;
}

// Oldest first, in the Worker rather than in Graph — see getConversationImpl for
// why the server-side sort had to go. `Array#sort` is stable (ES2019+), so
// messages sharing a timestamp keep the order Graph returned them in, and
// messages with no parseable `receivedDateTime` (Graph can omit it on drafts)
// fall to the end instead of scattering through the thread.
function sortByReceivedAsc<T extends ConversationMessage>(messages: T[]): T[] {
	return [...messages].sort((a, b) => {
		const ta = Date.parse(a.receivedDateTime ?? "");
		const tb = Date.parse(b.receivedDateTime ?? "");
		if (Number.isNaN(ta)) return Number.isNaN(tb) ? 0 : 1;
		if (Number.isNaN(tb)) return -1;
		return ta - tb;
	});
}

// Both inputs are supported: a message_id is resolved to its thread first.
async function resolveConversationIdFromMessage(env: Env, messageId: string): Promise<string> {
	let msg: { conversationId?: string } | null;
	try {
		msg = (await graphGet(env, `/me/messages/${messageId}`, {
			$select: "conversationId",
		})) as { conversationId?: string } | null;
	} catch (e) {
		// Otherwise this surfaces as "Outlook error 404: ErrorItemNotFound: ..."
		// — technically true, but it doesn't say which lookup failed.
		if (e instanceof ToolError && e.status === 404) {
			throw ToolError.notFound("Message", messageId);
		}
		throw e;
	}
	if (!msg?.conversationId) throw ToolError.notFound("Message", messageId);
	return msg.conversationId;
}

export async function getConversationImpl(
	env: Env,
	args: { conversation_id?: string; message_id?: string; count?: number },
): Promise<unknown> {
	let conversationId = args.conversation_id;
	if (!conversationId) {
		if (!args.message_id) {
			throw ToolError.validation("Either conversation_id or message_id is required");
		}
		conversationId = await resolveConversationIdFromMessage(env, args.message_id);
	}

	// Clamp defensively: the schema already bounds `count`, but this function is
	// called directly from tests and keeping the Graph page bounded is the whole
	// reason the request stays cheap enough to sort locally.
	const count = Math.min(
		Math.max(args.count ?? CONVERSATION_DEFAULT_COUNT, 1),
		CONVERSATION_MAX_COUNT,
	);
	const safeId = escapeOdataString(conversationId);

	// Filter only — no `$orderby`, no `ConsistencyLevel: eventual`. Pairing a
	// conversationId restriction with a server-side date sort makes Graph reject
	// the whole request:
	//
	//   400 InefficientFilter: The restriction or sort order is too complex for
	//   this operation.
	//
	// A mailbox has no index that serves both at once, and the advanced-query
	// opt-in that unlocks filter+orderby on directory endpoints does not apply to
	// /me/messages. A bare conversationId filter is supported and cross-folder
	// (Graph searches the whole mailbox — inbox, sent, archive), so the shape of
	// the result is unchanged; only the ordering moves into the Worker below.
	const data = (await graphGet(env, "/me/messages", {
		$filter: `conversationId eq '${safeId}'`,
		$select: CONVERSATION_SELECT,
		$top: count,
	})) as { value?: ConversationMessage[] } | null;

	const messages = Array.isArray(data?.value) ? data.value : [];
	if (messages.length === 0) throw ToolError.notFound("Conversation", conversationId);

	return {
		conversationId,
		count: messages.length,
		messages: sortByReceivedAsc(messages).map((m) => sanitizeEmailFull(m)),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
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
