import { ToolError } from "@bashco/mcp-toolkit";
import { beforeEach, describe, expect, it, vi } from "vitest";

// Mock the Graph client layer so we can assert exactly which query options
// get_conversation sends, without standing up a live Microsoft Graph session.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	// email-helpers.ts imports these; include them so the module resolves even
	// though the conversation path never calls them.
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet } from "../src/graph.js";
import { getConversationImpl } from "../src/tools/email.js";

const env = {} as never;

const CONV_ID = "AAQkAGI2conversation";
const MSG_IN_THREAD = "msg-middle-of-thread";

// Deliberately out of order: Graph makes no ordering promise once $orderby is
// gone, so the fixture must not accidentally arrive sorted.
const THREAD = [
	{
		id: "m2",
		conversationId: CONV_ID,
		subject: "Re: Quarterly numbers",
		from: { emailAddress: { name: "Adam", address: "adam@example.com" } },
		toRecipients: [{ emailAddress: { address: "me@bashco.co.nz" } }],
		ccRecipients: [],
		receivedDateTime: "2026-08-05T11:00:00Z",
		body: { contentType: "text", content: "Second." },
		hasAttachments: false,
		isDraft: false,
	},
	{
		id: "m3",
		conversationId: CONV_ID,
		subject: "Re: Quarterly numbers",
		from: { emailAddress: { name: "Me", address: "me@bashco.co.nz" } },
		toRecipients: [{ emailAddress: { address: "adam@example.com" } }],
		ccRecipients: [],
		receivedDateTime: "2026-08-06T09:30:00Z",
		body: { contentType: "text", content: "Third." },
		hasAttachments: false,
		isDraft: false,
	},
	{
		id: "m1",
		conversationId: CONV_ID,
		subject: "Quarterly numbers",
		from: { emailAddress: { name: "Adam", address: "adam@example.com" } },
		toRecipients: [{ emailAddress: { address: "me@bashco.co.nz" } }],
		ccRecipients: [],
		receivedDateTime: "2026-08-04T08:00:00Z",
		body: { contentType: "text", content: "First." },
		hasAttachments: false,
		isDraft: false,
	},
];

function messagesQuery(): Record<string, unknown> {
	const call = vi
		.mocked(graphGet)
		.mock.calls.find(([, path]) => path === "/me/messages");
	if (!call) throw new Error("expected a GET /me/messages call");
	return (call[2] ?? {}) as Record<string, unknown>;
}

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphGet).mockImplementation(
		async (_e: unknown, path: string): Promise<unknown> => {
			if (path === "/me/messages") return { value: THREAD };
			if (path === `/me/messages/${MSG_IN_THREAD}`) return { conversationId: CONV_ID };
			throw new Error(`unexpected graphGet path: ${path}`);
		},
	);
});

describe("lookup by conversation_id", () => {
	it("returns the thread", async () => {
		const result = (await getConversationImpl(env, { conversation_id: CONV_ID })) as {
			conversationId: string;
			count: number;
			messages: Array<{ id: string }>;
		};

		expect(result.conversationId).toBe(CONV_ID);
		expect(result.count).toBe(3);
		expect(result.messages).toHaveLength(3);
	});

	it("preserves the sanitised message shape", async () => {
		const result = (await getConversationImpl(env, { conversation_id: CONV_ID })) as {
			messages: Array<Record<string, unknown>>;
		};

		expect(result.messages[0]).toEqual({
			id: "m1",
			conversationId: CONV_ID,
			subject: "Quarterly numbers",
			from: "Adam <adam@example.com>",
			to: ["me@bashco.co.nz"],
			cc: [],
			received: "2026-08-04T08:00:00Z",
			body: "First.",
			hasAttachments: false,
			isDraft: false,
		});
	});
});

describe("lookup by message_id", () => {
	it("resolves the conversation from the message first", async () => {
		const result = (await getConversationImpl(env, { message_id: MSG_IN_THREAD })) as {
			conversationId: string;
			count: number;
		};

		expect(vi.mocked(graphGet).mock.calls[0]?.[1]).toBe(`/me/messages/${MSG_IN_THREAD}`);
		expect(result.conversationId).toBe(CONV_ID);
		expect(result.count).toBe(3);
		// The thread query filters on the resolved id, not the message id.
		expect(messagesQuery().$filter).toBe(`conversationId eq '${CONV_ID}'`);
	});
});

describe("ordering", () => {
	// Graph used to do this server-side. It now happens in the Worker, so the
	// guarantee needs its own test rather than riding on the upstream contract.
	it("returns messages oldest first regardless of Graph's order", async () => {
		const result = (await getConversationImpl(env, { conversation_id: CONV_ID })) as {
			messages: Array<{ id: string; received: string }>;
		};

		expect(result.messages.map((m) => m.id)).toEqual(["m1", "m2", "m3"]);
	});

	it("sorts messages with no receivedDateTime to the end", async () => {
		vi.mocked(graphGet).mockResolvedValue({
			value: [{ id: "no-date" }, ...THREAD],
		});

		const result = (await getConversationImpl(env, { conversation_id: CONV_ID })) as {
			messages: Array<{ id: string }>;
		};

		expect(result.messages.map((m) => m.id)).toEqual(["m1", "m2", "m3", "no-date"]);
	});
});

// The bug this fixes: `$filter=conversationId eq '...'` plus
// `$orderby=receivedDateTime asc` makes Graph return
// `400 InefficientFilter: The restriction or sort order is too complex for this
// operation.` Neither the sort nor the advanced-query header may come back.
describe("Graph request shape", () => {
	it("sends the conversation filter without a server-side sort", async () => {
		await getConversationImpl(env, { conversation_id: CONV_ID });

		const query = messagesQuery();
		expect(query.$filter).toBe(`conversationId eq '${CONV_ID}'`);
		expect(query.$orderby).toBeUndefined();
	});

	it("does not send the ConsistencyLevel advanced-query header", async () => {
		await getConversationImpl(env, { conversation_id: CONV_ID });

		const call = vi
			.mocked(graphGet)
			.mock.calls.find(([, path]) => path === "/me/messages");
		expect(call?.[3]).toBeUndefined();
	});

	it("bounds the page with $top", async () => {
		await getConversationImpl(env, { conversation_id: CONV_ID });
		expect(messagesQuery().$top).toBe(50);

		vi.clearAllMocks();
		vi.mocked(graphGet).mockResolvedValue({ value: THREAD });
		await getConversationImpl(env, { conversation_id: CONV_ID, count: 5 });
		expect(messagesQuery().$top).toBe(5);

		vi.clearAllMocks();
		vi.mocked(graphGet).mockResolvedValue({ value: THREAD });
		await getConversationImpl(env, { conversation_id: CONV_ID, count: 5000 });
		expect(messagesQuery().$top).toBe(200);
	});

	it("escapes single quotes in the conversation id", async () => {
		await getConversationImpl(env, { conversation_id: "a'b" });
		expect(messagesQuery().$filter).toBe("conversationId eq 'a''b'");
	});
});

describe("errors", () => {
	it("reports a missing conversation compactly instead of returning nothing", async () => {
		vi.mocked(graphGet).mockResolvedValue({ value: [] });

		await expect(
			getConversationImpl(env, { conversation_id: "no-such-thread" }),
		).rejects.toMatchObject({
			userMessage: "Conversation not found",
			status: 404,
		});
	});

	it("reports a missing message rather than passing Graph's error through", async () => {
		vi.mocked(graphGet).mockRejectedValue(
			new ToolError({
				userMessage: "Outlook error 404: ErrorItemNotFound: The specified object was not found",
				internalMessage: "Graph 404: {}",
				status: 404,
				upstreamName: "Graph",
			}),
		);

		await expect(getConversationImpl(env, { message_id: "gone" })).rejects.toMatchObject({
			userMessage: "Message not found",
			status: 404,
		});
	});

	it("reports a message that carries no conversationId", async () => {
		vi.mocked(graphGet).mockResolvedValue({});

		await expect(
			getConversationImpl(env, { message_id: "orphan" }),
		).rejects.toMatchObject({
			userMessage: "Message not found",
			status: 404,
		});
	});

	it("rejects a call with neither id", async () => {
		await expect(getConversationImpl(env, {})).rejects.toMatchObject({
			userMessage: "Either conversation_id or message_id is required",
		});
	});

	it("lets non-404 Graph failures through untouched", async () => {
		vi.mocked(graphGet).mockRejectedValue(
			new ToolError({
				userMessage: "Outlook error 503: ServiceUnavailable",
				status: 503,
				upstreamName: "Graph",
			}),
		);

		await expect(getConversationImpl(env, { message_id: "any" })).rejects.toMatchObject({
			userMessage: "Outlook error 503: ServiceUnavailable",
			status: 503,
		});
	});
});
