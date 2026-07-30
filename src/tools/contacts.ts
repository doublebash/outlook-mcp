import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPatch, graphPost } from "../graph.js";
import { sanitizeContactList } from "../sanitize.js";
import type { Env } from "../types.js";
import { escapeOdataString } from "./_shared.js";

export async function listContactsImpl(
	env: Env,
	args: { search?: string; count?: number },
): Promise<unknown> {
	const count = args.count ?? 20;
	const query: Record<string, string | number> = {
		// Graph contacts have no generic `phones` property — phone numbers live in
		// mobilePhone (string) plus businessPhones / homePhones (string arrays).
		// Selecting `phones` 400s with a ParseUri error.
		$select:
			"id,displayName,givenName,surname,emailAddresses,mobilePhone,businessPhones,homePhones,companyName,jobTitle",
		$top: count,
		$orderby: "displayName",
	};
	if (args.search) {
		// Escape single quotes so a name containing `'` (e.g. `O'Brien`) doesn't
		// break out of the OData literal.
		query.$filter = `contains(displayName,'${escapeOdataString(args.search)}')`;
	}
	const data = (await graphGet(env, "/me/contacts", query)) as { value: unknown[] };
	return sanitizeContactList(data.value);
}

export async function createContactImpl(
	env: Env,
	args: {
		first_name: string;
		last_name: string;
		email?: string;
		phone?: string;
		company?: string;
		job_title?: string;
	},
): Promise<unknown> {
	const contact: Record<string, unknown> = {
		givenName: args.first_name,
		surname: args.last_name,
	};

	if (args.email) {
		contact.emailAddresses = [
			{
				address: args.email,
				name: `${args.first_name} ${args.last_name}`,
			},
		];
	}
	if (args.phone) contact.mobilePhone = args.phone;
	if (args.company) contact.companyName = args.company;
	if (args.job_title) contact.jobTitle = args.job_title;

	const data = await graphPost(env, "/me/contacts", contact);
	return { success: true, message: "Contact created.", contact: data };
}

// Graph stores `displayName` as its own property rather than deriving it from
// givenName/surname, so renaming a contact without touching displayName leaves
// the old name showing everywhere the contact is listed. When a name part
// changes and no explicit display_name is given, fetch whichever half wasn't
// supplied and rebuild displayName so the two stay consistent.
async function resolveDisplayName(
	env: Env,
	id: string,
	firstName: string | undefined,
	lastName: string | undefined,
): Promise<string | undefined> {
	const current = (await graphGet(env, `/me/contacts/${id}`, {
		$select: "givenName,surname",
	})) as { givenName?: string | null; surname?: string | null };

	const first = (firstName ?? current.givenName ?? "").trim();
	const last = (lastName ?? current.surname ?? "").trim();
	const rebuilt = [first, last].filter(Boolean).join(" ");
	return rebuilt || undefined;
}

export async function updateContactImpl(
	env: Env,
	args: {
		id: string;
		first_name?: string;
		last_name?: string;
		display_name?: string;
		email?: string;
		mobile?: string;
		business_phones?: string[];
		home_phones?: string[];
		company?: string;
		job_title?: string;
	},
): Promise<unknown> {
	const updates: Record<string, unknown> = {};
	const notes: string[] = [];

	// Every text field below is checked with `!== undefined` rather than for
	// truthiness: an explicit empty string means "clear this", which Graph
	// expects as null. Omitting the field leaves it untouched.
	if (args.first_name !== undefined) updates.givenName = args.first_name || null;
	if (args.last_name !== undefined) updates.surname = args.last_name || null;
	if (args.company !== undefined) updates.companyName = args.company || null;
	if (args.job_title !== undefined) updates.jobTitle = args.job_title || null;
	if (args.mobile !== undefined) updates.mobilePhone = args.mobile || null;
	if (args.business_phones !== undefined) updates.businessPhones = args.business_phones;
	if (args.home_phones !== undefined) updates.homePhones = args.home_phones;

	if (args.email !== undefined) {
		// Graph replaces the whole emailAddresses collection on PATCH, so this
		// drops any additional addresses the contact had. Called out in the tool
		// description so the caller isn't surprised.
		updates.emailAddresses = args.email ? [{ address: args.email }] : [];
	}

	if (args.display_name !== undefined) {
		updates.displayName = args.display_name || null;
	} else if (args.first_name !== undefined || args.last_name !== undefined) {
		const rebuilt = await resolveDisplayName(env, args.id, args.first_name, args.last_name);
		if (rebuilt) {
			updates.displayName = rebuilt;
			notes.push(`displayName was updated to "${rebuilt}" to match the new name.`);
		}
	}

	if (Object.keys(updates).length === 0) {
		throw ToolError.validation(
			"No fields provided to update. Pass at least one of: first_name, last_name, " +
				"display_name, email, mobile, business_phones, home_phones, company, job_title. " +
				"Pass an empty string to clear a text field, or an empty array to clear phone lists.",
		);
	}

	const data = await graphPatch(env, `/me/contacts/${args.id}`, updates);
	const [contact] = sanitizeContactList([data]);
	const result: Record<string, unknown> = {
		success: true,
		message: "Contact updated.",
		contact,
	};
	if (notes.length > 0) result.notes = notes;
	return result;
}

export const contactsTools = defineTools<Env>({
	list_contacts: {
		description: "List contacts from your Outlook address book, with optional search.",
		schema: z.object({
			search: z.string().min(1).max(256).optional(),
			count: z.number().int().min(1).max(200).optional(),
		}),
		handler: (env, args) => listContactsImpl(env, args),
	},

	create_contact: {
		description: "Create a new contact in Outlook.",
		schema: z.object({
			first_name: z.string().min(1).max(64),
			last_name: z.string().min(1).max(64),
			email: z.string().email().max(256).optional(),
			phone: z.string().min(1).max(64).optional(),
			company: z.string().min(1).max(256).optional(),
			job_title: z.string().min(1).max(256).optional(),
		}),
		handler: (env, args) => createContactImpl(env, args),
	},

	update_contact: {
		description:
			"Update an existing Outlook contact. Use list_contacts to find the contact's id first. " +
			"Only provide the fields you want to change; omitted fields are left untouched. " +
			"To clear a field, pass an empty string (text fields) or an empty array (business_phones, " +
			"home_phones). Note that email replaces the contact's entire email list, and " +
			"business_phones/home_phones replace the whole list rather than appending. " +
			"Empty calls (no fields provided) are rejected.",
		schema: z.object({
			id: z.string().min(1).max(512),
			first_name: z.string().max(64).optional(),
			last_name: z.string().max(64).optional(),
			display_name: z.string().max(256).optional(),
			email: z.union([z.string().email().max(256), z.literal("")]).optional(),
			mobile: z.string().max(64).optional(),
			business_phones: z.array(z.string().min(1).max(64)).max(10).optional(),
			home_phones: z.array(z.string().min(1).max(64)).max(10).optional(),
			company: z.string().max(256).optional(),
			job_title: z.string().max(256).optional(),
		}),
		handler: (env, args) => updateContactImpl(env, args),
	},
});
