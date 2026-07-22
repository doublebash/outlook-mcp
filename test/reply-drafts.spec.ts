import { describe, it, expect, vi, beforeEach } from "vitest";
import { ToolError } from "@bashco/mcp-toolkit";

// Mock the Graph client layer so we can assert which endpoints the threaded
// draft tools call without standing up a live Microsoft Graph session.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	// email-helpers.ts imports these; include them so the module resolves even
	// though the reply-draft path never calls them.
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphPost, graphPatch, graphDelete } from "../src/graph.js";
import { createReplyDraftImpl, createForwardDraftImpl } from "../src/tools/email.js";

const env = {} as never;

const OWNER = "me@bashco.co.nz";
const ORIG_OTHER = "orig-from-someone-else";
const ORIG_SELF = "orig-from-me";
const DRAFT_ID = "draft-123";

const DRAFT_META = {
	id: DRAFT_ID,
	subject: "Re: Quarterly numbers",
	conversationId: "conv-abc",
	webLink: "https://outlook.office.com/owa/?ItemID=DRAFT",
	toRecipients: [{ emailAddress: { name: "Adam", address: "adam@example.com" } }],
	ccRecipients: [],
	createdDateTime: "2026-06-15T00:00:00Z",
	lastModifiedDateTime: "2026-06-15T00:01:00Z",
};

beforeEach(() => {
	vi.clearAllMocks();

	// Default: createReply/createReplyAll/createForward return a draft id.
	vi.mocked(graphPost).mockResolvedValue({ id: DRAFT_ID });
	vi.mocked(graphPatch).mockResolvedValue({});
	vi.mocked(graphDelete).mockResolvedValue({});

	vi.mocked(graphGet).mockImplementation(
		async (_e: unknown, path: string): Promise<unknown> => {
			if (path === "/me") return { mail: OWNER };
			if (path === `/me/messages/${ORIG_OTHER}`) {
				// Sent by someone else → Graph's default recipients are correct.
				return {
					from: { emailAddress: { address: "adam@example.com" } },
					toRecipients: [{ emailAddress: { address: OWNER } }],
					ccRecipients: [],
				};
			}
			if (path === `/me/messages/${ORIG_SELF}`) {
				// Sent by the owner → self-reply redirect should kick in.
				return {
					from: { emailAddress: { address: OWNER } },
					toRecipients: [{ emailAddress: { address: "client@example.com" } }],
					ccRecipients: [],
				};
			}
			if (path === `/me/messages/${DRAFT_ID}`) return DRAFT_META;
			throw new Error(`unexpected graphGet path: ${path}`);
		},
	);
});

// Graph's createReply/createForward draft already contains the quoted original.
// Historically the tools PATCHed `body` with just the author's new text, which
// replaced that whole document and silently discarded the conversation history.
// These lock in the fix: the new content goes ABOVE the quote, quote intact.
describe("quoted original is preserved", () => {
	const QUOTE =
		'<div id="divRplyFwdMsg">From: adam@example.com<br>Sent: Monday<br>Subject: Quarterly numbers</div><div>Original message text.</div>';
	const GENERATED = `<html><body><div></div><div id="appendonsend"></div><hr>${QUOTE}</body></html>`;

	function withGeneratedBody() {
		const base = vi.mocked(graphGet).getMockImplementation()!;
		vi.mocked(graphGet).mockImplementation(async (e: unknown, path: string, q?: unknown) => {
			const query = q as { $select?: string } | undefined;
			if (path === `/me/messages/${DRAFT_ID}` && query?.$select === "body") {
				return { body: { contentType: "HTML", content: GENERATED } };
			}
			return base(e, path, q);
		});
	}

	function patchedBody(): string {
		const call = vi.mocked(graphPatch).mock.calls.at(-1);
		const updates = call?.[2] as { body?: { content?: string } };
		return updates.body?.content ?? "";
	}

	it("keeps the quoted original when replying with a body", async () => {
		withGeneratedBody();
		await createReplyDraftImpl(env, { id: ORIG_OTHER, body: "Thanks, noted." });

		const content = patchedBody();
		expect(content).toContain("Original message text.");
		expect(content).toContain("From: adam@example.com");
	});

	it("places the reply above the quote, not below it", async () => {
		withGeneratedBody();
		await createReplyDraftImpl(env, { id: ORIG_OTHER, body: "Thanks, noted." });

		const content = patchedBody();
		expect(content.indexOf("Thanks, noted.")).toBeLessThan(content.indexOf("divRplyFwdMsg"));
	});

	it("keeps the quoted original when forwarding with a comment", async () => {
		withGeneratedBody();
		await createForwardDraftImpl(env, {
			id: ORIG_OTHER,
			to: "bob@example.com",
			body: "FYI",
		});

		const content = patchedBody();
		expect(content).toContain("Original message text.");
		expect(content.indexOf("FYI")).toBeLessThan(content.indexOf("divRplyFwdMsg"));
	});

	it("converts a plain-text reply to HTML so the quote can be carried", async () => {
		withGeneratedBody();
		await createReplyDraftImpl(env, {
			id: ORIG_OTHER,
			body: "line one\nline two",
			body_type: "text",
		});

		const call = vi.mocked(graphPatch).mock.calls.at(-1);
		const updates = call?.[2] as { body?: { contentType?: string; content?: string } };
		expect(updates.body?.contentType).toBe("HTML");
		expect(updates.body?.content).toContain("line one<br>line two");
		expect(updates.body?.content).toContain("Original message text.");
	});

	it("still sanitises the author's HTML while keeping the quote", async () => {
		withGeneratedBody();
		await createReplyDraftImpl(env, {
			id: ORIG_OTHER,
			body: '<p onclick="x()">Hi</p><script>alert(1)</script>',
			body_type: "html",
		});

		const content = patchedBody();
		expect(content).not.toContain("<script");
		expect(content).not.toContain("onclick");
		expect(content).toContain("Original message text.");
	});

	it("falls back to the author's text when the draft body cannot be read", async () => {
		const base = vi.mocked(graphGet).getMockImplementation()!;
		vi.mocked(graphGet).mockImplementation(async (e: unknown, path: string, q?: unknown) => {
			const query = q as { $select?: string } | undefined;
			if (path === `/me/messages/${DRAFT_ID}` && query?.$select === "body") {
				throw new Error("Graph 500");
			}
			return base(e, path, q);
		});

		// The reply must still go out — losing the quote beats losing the draft.
		const result = (await createReplyDraftImpl(env, {
			id: ORIG_OTHER,
			body: "Thanks, noted.",
		})) as Record<string, unknown>;

		expect(patchedBody()).toContain("Thanks, noted.");
		expect(JSON.stringify(result.notes)).toMatch(/quoted original is not included/);
	});
});

describe("create_reply_draft", () => {
	it("calls POST /me/messages/{id}/createReply", async () => {
		await createReplyDraftImpl(env, { id: ORIG_OTHER });
		expect(graphPost).toHaveBeenCalledWith(env, `/me/messages/${ORIG_OTHER}/createReply`);
	});

	it("patches the returned draft when a body is supplied (HTML)", async () => {
		await createReplyDraftImpl(env, {
			id: ORIG_OTHER,
			body: "<p>Thanks!</p>",
			body_type: "html",
		});
		expect(graphPatch).toHaveBeenCalledWith(
			env,
			`/me/messages/${DRAFT_ID}`,
			expect.objectContaining({
				body: { contentType: "HTML", content: expect.stringContaining("Thanks!") },
			}),
		);
	});

	it("does not patch when body is omitted (and recipients need no override)", async () => {
		await createReplyDraftImpl(env, { id: ORIG_OTHER });
		expect(graphPatch).not.toHaveBeenCalled();
	});

	it("returns draft id, subject, conversation id, and web link", async () => {
		const result = (await createReplyDraftImpl(env, { id: ORIG_OTHER })) as Record<
			string,
			unknown
		>;
		expect(result).toMatchObject({
			success: true,
			id: DRAFT_ID,
			draft_id: DRAFT_ID,
			subject: "Re: Quarterly numbers",
			conversationId: "conv-abc",
			webLink: expect.stringContaining("outlook.office.com"),
		});
	});

	it("returns a clear error when the original message is not found", async () => {
		vi.mocked(graphPost).mockRejectedValueOnce(
			new ToolError({
				userMessage: "Outlook error 404: ErrorItemNotFound",
				internalMessage: "Graph 404: ErrorItemNotFound: not found",
				status: 404,
				upstreamName: "Graph",
			}),
		);
		await expect(createReplyDraftImpl(env, { id: "missing" })).rejects.toThrow(
			/message not found or not accessible/,
		);
	});

	it("passes a non-404 Graph create failure through unchanged", async () => {
		vi.mocked(graphPost).mockRejectedValueOnce(new Error("network boom"));
		await expect(createReplyDraftImpl(env, { id: ORIG_OTHER })).rejects.toThrow(
			/network boom/,
		);
	});

	it("keeps the created draft and warns when the body patch fails", async () => {
		vi.mocked(graphPatch).mockRejectedValueOnce(new Error("patch boom"));
		const result = (await createReplyDraftImpl(env, {
			id: ORIG_OTHER,
			body: "hi",
		})) as Record<string, unknown>;
		expect(result).toMatchObject({
			success: false,
			draft_id: DRAFT_ID,
			error: "patch boom",
		});
		expect(result.warning).toEqual(expect.stringContaining("draft"));
		// The draft must NOT be deleted — the user shouldn't lose it.
		expect(graphDelete).not.toHaveBeenCalled();
	});

	it("redirects a reply to your own sent message back to the original recipient", async () => {
		await createReplyDraftImpl(env, { id: ORIG_SELF, body: "follow-up" });
		const patchArg = vi.mocked(graphPatch).mock.calls[0]?.[2] as {
			toRecipients?: unknown;
		};
		expect(patchArg.toRecipients).toEqual([
			{ emailAddress: { address: "client@example.com" } },
		]);
	});
});

describe("create_reply_all_draft (reply_all flag)", () => {
	it("calls POST /me/messages/{id}/createReplyAll", async () => {
		await createReplyDraftImpl(env, { id: ORIG_OTHER, reply_all: true });
		expect(graphPost).toHaveBeenCalledWith(
			env,
			`/me/messages/${ORIG_OTHER}/createReplyAll`,
		);
	});

	it("preserves Graph's recipients (no manual rebuild) for a normal thread", async () => {
		await createReplyDraftImpl(env, { id: ORIG_OTHER, reply_all: true, body: "ok" });
		const patchArg = vi.mocked(graphPatch).mock.calls[0]?.[2] as Record<string, unknown>;
		expect(patchArg).not.toHaveProperty("toRecipients");
		expect(patchArg).not.toHaveProperty("ccRecipients");
		expect(patchArg).toHaveProperty("body");
	});
});

describe("create_forward_draft", () => {
	it("calls POST /me/messages/{id}/createForward", async () => {
		await createForwardDraftImpl(env, { id: ORIG_OTHER, to: "bob@example.com" });
		expect(graphPost).toHaveBeenCalledWith(
			env,
			`/me/messages/${ORIG_OTHER}/createForward`,
		);
	});

	it("patches recipients and body when supplied", async () => {
		await createForwardDraftImpl(env, {
			id: ORIG_OTHER,
			to: "bob@example.com",
			cc: "carol@example.com",
			body: "FYI",
		});
		expect(graphPatch).toHaveBeenCalledWith(
			env,
			`/me/messages/${DRAFT_ID}`,
			expect.objectContaining({
				toRecipients: [{ emailAddress: { address: "bob@example.com" } }],
				ccRecipients: [{ emailAddress: { address: "carol@example.com" } }],
				// HTML, not Text: the outgoing body has to carry Graph's quoted
			// original below the comment, and a Text body cannot.
			body: { contentType: "HTML", content: "FYI" },
			}),
		);
	});

	it("does not patch when no recipients or body are supplied", async () => {
		await createForwardDraftImpl(env, { id: ORIG_OTHER });
		expect(graphPatch).not.toHaveBeenCalled();
	});

	it("returns the draft metadata", async () => {
		const result = (await createForwardDraftImpl(env, {
			id: ORIG_OTHER,
			to: "bob@example.com",
		})) as Record<string, unknown>;
		expect(result).toMatchObject({
			success: true,
			id: DRAFT_ID,
			conversationId: "conv-abc",
		});
	});

	it("returns a clear error when the original message is not found", async () => {
		vi.mocked(graphPost).mockRejectedValueOnce(
			new ToolError({
				userMessage: "Outlook error 404: ErrorItemNotFound",
				internalMessage: "Graph 404: ErrorItemNotFound: not found",
				status: 404,
				upstreamName: "Graph",
			}),
		);
		await expect(
			createForwardDraftImpl(env, { id: "missing", to: "bob@example.com" }),
		).rejects.toThrow(/forward draft: message not found or not accessible/);
	});
});
