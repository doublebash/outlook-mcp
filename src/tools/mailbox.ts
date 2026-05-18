import { defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphPatch } from "../graph.js";
import { sanitizeMailboxSettings } from "../sanitize.js";
import type { Env } from "../types.js";

async function getMailboxSettingsImpl(env: Env): Promise<unknown> {
	const data = await graphGet(env, "/me/mailboxSettings");
	return sanitizeMailboxSettings(data);
}

async function setOutOfOfficeImpl(
	env: Env,
	args: {
		enabled: boolean;
		internal_message?: string;
		external_message?: string;
		start_datetime?: string;
		end_datetime?: string;
	},
): Promise<unknown> {
	const internalMsg = args.internal_message;
	const externalMsg = args.external_message ?? internalMsg;

	const replySettings: Record<string, unknown> = {
		status: args.enabled ? "alwaysEnabled" : "disabled",
	};

	if (args.enabled && internalMsg) {
		replySettings.internalReplyMessage = internalMsg;
		replySettings.externalReplyMessage = externalMsg;
	}

	if (args.enabled && args.start_datetime && args.end_datetime) {
		replySettings.status = "scheduled";
		replySettings.scheduledStartDateTime = {
			dateTime: args.start_datetime,
			timeZone: "UTC",
		};
		replySettings.scheduledEndDateTime = {
			dateTime: args.end_datetime,
			timeZone: "UTC",
		};
	}

	await graphPatch(env, "/me/mailboxSettings", {
		automaticRepliesSetting: replySettings,
	});

	return {
		success: true,
		message: `Out-of-office ${args.enabled ? "enabled" : "disabled"}.`,
	};
}

export const mailboxTools = defineTools<Env>({
	get_mailbox_settings: {
		description:
			"Get mailbox settings including out-of-office status, timezone, and working hours.",
		schema: z.object({}),
		handler: (env) => getMailboxSettingsImpl(env),
	},

	set_out_of_office: {
		description: "Enable or disable the automatic out-of-office reply.",
		schema: z.object({
			enabled: z.boolean(),
			internal_message: z.string().max(8192).optional(),
			external_message: z.string().max(8192).optional(),
			start_datetime: z.string().min(1).max(64).optional(),
			end_datetime: z.string().min(1).max(64).optional(),
		}),
		handler: (env, args) => setOutOfOfficeImpl(env, args),
	},
});
