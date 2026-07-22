import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph layer so we can assert exactly what body gets written to the
// draft, without a live Microsoft Graph session.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphPost, graphPatch } from "../src/graph.js";
import { createDraftImpl, updateDraftImpl } from "../src/tools/email.js";

const DRAFT_ID = "draft-xyz";

// A stand-in signature with the same shape as the real one: inline styles plus a
// distinctive tagline the idempotency guard can key on.
const SIG = `<table style="border-collapse:collapse;"><tr><td><div style="font-weight:700;">Test Sender</div><div style="letter-spacing:0.4px;">WIN THE LISTING. SELL THE PROPERTY.</div></td></tr></table>`;

const env = { SIGNATURE_HTML: SIG, SIGNATURE_LOGO_URL: "https://cdn.example.com/l.png" } as never;
const envNoSig = {} as never;

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphPost).mockResolvedValue({ id: DRAFT_ID, webLink: "https://outlook/DRAFT" });
	vi.mocked(graphPatch).mockResolvedValue({});
});

function postedBody(): { contentType?: string; content?: string } {
	// createDraft posts the message object to /me/messages.
	const call = vi.mocked(graphPost).mock.calls.find((c) => c[1] === "/me/messages");
	const msg = call?.[2] as { body?: { contentType?: string; content?: string } };
	return msg.body ?? {};
}

describe("create_draft — signature injection", () => {
	it("appends the signature when include_signature is true", async () => {
		await createDraftImpl(env, {
			subject: "Proposal",
			body: "Hi, here is our proposal.",
			include_signature: true,
		});
		const body = postedBody();
		expect(body.contentType).toBe("HTML");
		expect(body.content).toContain("Test Sender");
		expect(body.content).toContain("Hi, here is our proposal.<br><br>");
	});

	it("bakes the signature in at create time so send_draft sends it as-is", async () => {
		// send_draft never touches the body, so the signature has to be present in
		// the draft the moment it is created.
		await createDraftImpl(env, {
			subject: "Proposal",
			body: "Body text",
			include_signature: true,
		});
		expect(postedBody().content).toContain("WIN THE LISTING. SELL THE PROPERTY.");
	});

	it("leaves the body untouched when include_signature is false", async () => {
		await createDraftImpl(env, {
			subject: "Proposal",
			body: "Body text",
			include_signature: false,
		});
		const body = postedBody();
		expect(body.content).not.toContain("Test Sender");
		expect(body.contentType).toBe("Text");
	});

	it("is a no-op when the deployment has no signature configured", async () => {
		await createDraftImpl(envNoSig, {
			subject: "Proposal",
			body: "Body text",
			include_signature: true,
		});
		expect(postedBody().content).not.toContain("Test Sender");
	});

	it("reports the body_type override in notes", async () => {
		const res = (await createDraftImpl(env, {
			subject: "Proposal",
			body: "Body text",
			body_type: "text",
			include_signature: true,
		})) as Record<string, unknown>;
		expect(JSON.stringify(res.notes)).toMatch(/overridden from 'text' to 'html'/);
	});

	it("does not double up when the agent already pasted the signature", async () => {
		await createDraftImpl(env, {
			subject: "Proposal",
			body: `Hi there<br><br>${SIG}`,
			body_type: "html",
			include_signature: true,
		});
		const content = postedBody().content ?? "";
		expect(content.split("WIN THE LISTING").length - 1).toBe(1);
	});
});

describe("update_draft — signature injection", () => {
	function patchedBody(): { contentType?: string; content?: string } {
		const call = vi.mocked(graphPatch).mock.calls.at(-1);
		const updates = call?.[2] as { body?: { contentType?: string; content?: string } };
		return updates?.body ?? {};
	}

	it("signs the new body when include_signature is true and a body is given", async () => {
		await updateDraftImpl(env, { id: DRAFT_ID, body: "Revised text", include_signature: true });
		const body = patchedBody();
		expect(body.contentType).toBe("HTML");
		expect(body.content).toContain("Revised text<br><br>");
		expect(body.content).toContain("Test Sender");
	});

	it("does NOT wipe the draft to a signature-only body when no body is provided", async () => {
		// The guard: include_signature with no body must not build a signature-only
		// body that replaces whatever the draft held.
		const res = (await updateDraftImpl(env, {
			id: DRAFT_ID,
			subject: "New subject only",
			include_signature: true,
		})) as Record<string, unknown>;

		const call = vi.mocked(graphPatch).mock.calls.at(-1);
		const updates = call?.[2] as Record<string, unknown>;
		expect(updates.body).toBeUndefined();
		expect(updates.subject).toBe("New subject only");
		expect(JSON.stringify(res.notes)).toMatch(/only applies when you also provide a new body/);
	});

	it("does not sign when include_signature is false", async () => {
		await updateDraftImpl(env, { id: DRAFT_ID, body: "Plain revision", include_signature: false });
		expect(patchedBody().content).not.toContain("Test Sender");
	});
});
