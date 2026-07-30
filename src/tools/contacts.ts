import { defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPost } from "../graph.js";
import { sanitizeContactList } from "../sanitize.js";
import type { Env } from "../types.js";
import { escapeOdataString } from "./_shared.js";

async function listContactsImpl(
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

async function createContactImpl(
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
});
