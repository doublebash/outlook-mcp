import { describe, it, expect, vi, beforeEach } from "vitest";
import { ToolError } from "@bashco/mcp-toolkit";

// Mock the Graph client layer so we can assert exactly which properties the
// contact tools send upstream without standing up a live Microsoft Graph
// session — the `phones` regression was invisible until a real request ran.
vi.mock("../src/graph.js", () => ({
	graphGet: vi.fn(),
	graphPost: vi.fn(),
	graphPatch: vi.fn(),
	graphDelete: vi.fn(),
	graphRequestRaw: vi.fn(),
	graphPutBinary: vi.fn(),
}));

import { graphGet, graphPost, graphPatch } from "../src/graph.js";
import {
	listContactsImpl,
	createContactImpl,
	updateContactImpl,
} from "../src/tools/contacts.js";
import { sanitizeContactList } from "../src/sanitize.js";

const env = {} as never;
const ID = "AAMkADNkContactId=";

beforeEach(() => {
	vi.clearAllMocks();
	vi.mocked(graphGet).mockResolvedValue({ value: [] });
	vi.mocked(graphPost).mockResolvedValue({ id: ID });
	vi.mocked(graphPatch).mockResolvedValue({ id: ID });
});

function lastPatchBody(): Record<string, unknown> {
	const call = vi.mocked(graphPatch).mock.calls.at(-1);
	return call?.[2] as Record<string, unknown>;
}

describe("list_contacts projection", () => {
	// Regression guard: `phones` is not a property of Graph's contact resource.
	// Selecting it made every list_contacts call fail with
	// `RequestBroker--ParseUri: Could not find a property named 'phones'`.
	it("never requests a generic `phones` property", async () => {
		await listContactsImpl(env, { count: 5 });

		const query = vi.mocked(graphGet).mock.calls[0]?.[2] as Record<string, string>;
		expect(query.$select).not.toContain("phones");
	});

	it("selects the real Graph phone fields", async () => {
		await listContactsImpl(env, {});

		const query = vi.mocked(graphGet).mock.calls[0]?.[2] as Record<string, string>;
		const selected = query.$select.split(",");
		expect(selected).toContain("mobilePhone");
		expect(selected).toContain("businessPhones");
		expect(selected).toContain("homePhones");
		expect(selected).toContain("emailAddresses");
	});

	it("escapes quotes in the search filter", async () => {
		await listContactsImpl(env, { search: "O'Brien" });

		const query = vi.mocked(graphGet).mock.calls[0]?.[2] as Record<string, string>;
		expect(query.$filter).toBe("contains(displayName,'O''Brien')");
	});
});

describe("create_contact", () => {
	it("writes a phone to mobilePhone, not a `phones` array", async () => {
		await createContactImpl(env, {
			first_name: "Rhys",
			last_name: "Theodorou",
			phone: "0275588575",
		});

		const body = vi.mocked(graphPost).mock.calls[0]?.[2] as Record<string, unknown>;
		expect(body.mobilePhone).toBe("0275588575");
		expect(body).not.toHaveProperty("phones");
	});
});

describe("update_contact", () => {
	it("rejects a call with no fields to change", async () => {
		await expect(updateContactImpl(env, { id: ID })).rejects.toThrow(ToolError);
		expect(graphPatch).not.toHaveBeenCalled();
	});

	it("sends only the fields that were supplied", async () => {
		await updateContactImpl(env, { id: ID, job_title: "Director" });

		expect(lastPatchBody()).toEqual({ jobTitle: "Director" });
	});

	it("patches the contact by id", async () => {
		await updateContactImpl(env, { id: ID, company: "Fluid Plumbing & Gas" });

		expect(vi.mocked(graphPatch).mock.calls[0]?.[1]).toBe(`/me/contacts/${ID}`);
	});

	// An empty string is the caller's way of saying "remove this value"; Graph
	// clears a scalar property when it receives null, not "".
	it("clears a text field with null when given an empty string", async () => {
		await updateContactImpl(env, { id: ID, mobile: "", job_title: "" });

		expect(lastPatchBody()).toEqual({ mobilePhone: null, jobTitle: null });
	});

	it("distinguishes clearing a field from leaving it alone", async () => {
		await updateContactImpl(env, { id: ID, company: "" });
		expect(lastPatchBody()).toEqual({ companyName: null });

		await updateContactImpl(env, { id: ID, mobile: "021 555 0000" });
		expect(lastPatchBody()).not.toHaveProperty("companyName");
	});

	it("clears phone lists with an empty array", async () => {
		await updateContactImpl(env, { id: ID, business_phones: [], home_phones: [] });

		expect(lastPatchBody()).toEqual({ businessPhones: [], homePhones: [] });
	});

	it("replaces the email collection, and empties it when cleared", async () => {
		await updateContactImpl(env, { id: ID, email: "new@example.com" });
		expect(lastPatchBody().emailAddresses).toEqual([{ address: "new@example.com" }]);

		await updateContactImpl(env, { id: ID, email: "" });
		expect(lastPatchBody().emailAddresses).toEqual([]);
	});

	// Graph does not recompute displayName from givenName/surname, so a rename
	// would otherwise leave the old name showing in every contact listing.
	it("rebuilds displayName when a name part changes", async () => {
		vi.mocked(graphGet).mockResolvedValue({ givenName: "Rhys", surname: "Theodorou" });

		const result = (await updateContactImpl(env, {
			id: ID,
			last_name: "Theodorou-Smith",
		})) as { notes?: string[] };

		expect(lastPatchBody().displayName).toBe("Rhys Theodorou-Smith");
		expect(result.notes?.[0]).toContain("Rhys Theodorou-Smith");
	});

	it("honours an explicit display_name without re-reading the contact", async () => {
		await updateContactImpl(env, {
			id: ID,
			first_name: "Rhys",
			display_name: "Rhys T. (Fluid)",
		});

		expect(lastPatchBody().displayName).toBe("Rhys T. (Fluid)");
		expect(graphGet).not.toHaveBeenCalled();
	});

	it("does not re-read the contact when no name part changes", async () => {
		await updateContactImpl(env, { id: ID, mobile: "021 555 0000" });

		expect(graphGet).not.toHaveBeenCalled();
		expect(lastPatchBody()).not.toHaveProperty("displayName");
	});

	it("leaves displayName alone when clearing both name parts would empty it", async () => {
		vi.mocked(graphGet).mockResolvedValue({ givenName: "Rhys", surname: "Theodorou" });

		await updateContactImpl(env, { id: ID, first_name: "", last_name: "" });

		const body = lastPatchBody();
		expect(body).toEqual({ givenName: null, surname: null });
		expect(body).not.toHaveProperty("displayName");
	});
});

describe("sanitizeContactList", () => {
	it("maps the three Graph phone fields", () => {
		const [contact] = sanitizeContactList([
			{
				id: ID,
				displayName: "Rhys Theodorou",
				givenName: "Rhys",
				surname: "Theodorou",
				emailAddresses: [{ address: "info@fluidsystems.co.nz" }],
				mobilePhone: "0275588575",
				businessPhones: ["09 555 1234"],
				homePhones: [],
				companyName: "Fluid Plumbing & Gas",
				jobTitle: "Owner",
			},
		]) as Array<Record<string, unknown>>;

		expect(contact).toMatchObject({
			name: "Rhys Theodorou",
			firstName: "Rhys",
			lastName: "Theodorou",
			email: "info@fluidsystems.co.nz",
			mobile: "0275588575",
			businessPhones: ["09 555 1234"],
			homePhones: [],
			company: "Fluid Plumbing & Gas",
			jobTitle: "Owner",
		});
	});

	it("falls back mobile -> business -> home for the single `phone` field", () => {
		const rows = sanitizeContactList([
			{ mobilePhone: "mob", businessPhones: ["biz"], homePhones: ["home"] },
			{ businessPhones: ["biz"], homePhones: ["home"] },
			{ homePhones: ["home"] },
			{},
		]) as Array<Record<string, unknown>>;

		expect(rows.map(r => r.phone)).toEqual(["mob", "biz", "home", null]);
	});

	// A contact with nothing populated must come back as empty values rather
	// than throwing — Outlook address books are full of partial records.
	it("returns empty values for an unpopulated contact", () => {
		const [contact] = sanitizeContactList([{ id: ID }]) as Array<Record<string, unknown>>;

		expect(contact).toEqual({
			id: ID,
			name: undefined,
			firstName: null,
			lastName: null,
			email: null,
			phone: null,
			mobile: null,
			businessPhones: [],
			homePhones: [],
			company: null,
			jobTitle: null,
		});
	});
});
