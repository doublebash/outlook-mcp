import { describe, it, expect, vi, beforeEach } from "vitest";

// Mock the Graph layer so we can assert the exact event body posted/patched.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
}));

import { graphPost, graphPatch } from "../src/graph.js";
import {
	createCalendarEventImpl,
	updateCalendarEventImpl,
} from "../src/tools/calendar.js";

const EVENT_ID = "evt-1";
const SIG = `<table style="border-collapse:collapse;"><tr><td><div style="font-weight:700;">Test Sender</div><div style="letter-spacing:0.4px;">WIN THE LISTING. SELL THE PROPERTY.</div></td></tr></table>`;
const env = { SIGNATURE_HTML: SIG, SIGNATURE_LOGO_URL: "https://cdn.example.com/l.png" } as never;
const envNoSig = {} as never;

const base = {
	subject: "Shoot briefing",
	start_datetime: "2026-08-01T10:00:00",
	end_datetime: "2026-08-01T10:30:00",
};

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphPost).mockResolvedValue({ id: EVENT_ID });
	vi.mocked(graphPatch).mockResolvedValue({});
});

function postedEventBody(): { contentType?: string; content?: string } {
	const call = vi.mocked(graphPost).mock.calls.find((c) => c[1] === "/me/events");
	const event = call?.[2] as { body?: { contentType?: string; content?: string } };
	return event.body ?? {};
}

function patchedBody(): { contentType?: string; content?: string } {
	const call = vi.mocked(graphPatch).mock.calls.at(-1);
	const updates = call?.[2] as { body?: { contentType?: string; content?: string } };
	return updates?.body ?? {};
}

describe("create_calendar_event — signature injection", () => {
	it("appends the signature to the description when include_signature is true", async () => {
		await createCalendarEventImpl(env, {
			...base,
			description: "Agenda: locations, timing, gear.",
			include_signature: true,
		});
		const body = postedEventBody();
		expect(body.contentType).toBe("HTML");
		expect(body.content).toContain("Agenda: locations, timing, gear.<br><br>");
		expect(body.content).toContain("Test Sender");
	});

	it("preserves the signature's inline styles (sanitiser boundary)", async () => {
		await createCalendarEventImpl(env, {
			...base,
			description: "Hi",
			include_signature: true,
		});
		expect(postedEventBody().content).toContain("border-collapse:collapse");
	});

	it("sends the signature alone for an event with no description", async () => {
		await createCalendarEventImpl(env, { ...base, include_signature: true });
		const body = postedEventBody();
		expect(body.content).toContain("Test Sender");
		expect(body.content?.startsWith("<br>")).toBe(false);
	});

	it("does not sign when include_signature is false", async () => {
		await createCalendarEventImpl(env, {
			...base,
			description: "Plain agenda",
			include_signature: false,
		});
		const body = postedEventBody();
		expect(body.content).not.toContain("Test Sender");
		expect(body.contentType).toBe("Text");
	});

	it("is a no-op when the deployment has no signature configured", async () => {
		await createCalendarEventImpl(envNoSig, {
			...base,
			description: "Agenda",
			include_signature: true,
		});
		expect(postedEventBody().content).not.toContain("Test Sender");
	});

	it("does not double up when the description already carries the signature", async () => {
		await createCalendarEventImpl(env, {
			...base,
			description: `Agenda<br><br>${SIG}`,
			body_type: "html",
			include_signature: true,
		});
		expect((postedEventBody().content ?? "").split("WIN THE LISTING").length - 1).toBe(1);
	});
});

describe("update_calendar_event — signature injection", () => {
	it("signs the new description when a description is supplied", async () => {
		await updateCalendarEventImpl(env, {
			id: EVENT_ID,
			description: "Revised agenda",
			include_signature: true,
		});
		const body = patchedBody();
		expect(body.contentType).toBe("HTML");
		expect(body.content).toContain("Revised agenda<br><br>");
		expect(body.content).toContain("Test Sender");
	});

	it("does NOT wipe the event to a signature-only body when no description is given", async () => {
		const res = (await updateCalendarEventImpl(env, {
			id: EVENT_ID,
			subject: "New subject only",
			include_signature: true,
		})) as Record<string, unknown>;

		const call = vi.mocked(graphPatch).mock.calls.at(-1);
		const updates = call?.[2] as Record<string, unknown>;
		expect(updates.body).toBeUndefined();
		expect(updates.subject).toBe("New subject only");
		expect(JSON.stringify(res.notes)).toMatch(/only applies when you also provide a new description/);
	});
});
