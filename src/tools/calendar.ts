import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { buildRecurrence, type RecurrenceInput } from "../calendar-helpers.js";
import { buildMessageBody } from "../email-helpers.js";
import { graphDelete, graphGet, graphPatch, graphPost } from "../graph.js";
import { sanitizeEventList } from "../sanitize.js";
import { maybeSign } from "../signature.js";
import type { Env } from "../types.js";
import {
	BODY_TYPE_DESC,
	auditLog,
	escapeOdataString,
	recurrenceSchema,
} from "./_shared.js";

// Opt-in signature on the event description, default false. Reuses the same
// SIGNATURE_HTML config as email. The note about marketing CTAs is deliberate:
// the signature carries sales buttons that suit a client-facing invite but not
// an internal one, so the flag is left to the caller per-event.
const includeSignatureSchema = z.boolean().optional().default(false);
const INCLUDE_SIGNATURE_DESC =
	"Append the sender's configured signature to the event description, server-side. Set this flag rather than writing the signature into `description` yourself. Forces the description to HTML. No-op if no signature is configured. Note the signature includes marketing call-to-action buttons, which may not suit internal meetings. Default false.";

// Common $select for events — kept here so list and instances stay in sync.
// Note: `recurrence` is included so sanitizeEventList can flag `isRecurring`.
const EVENT_SELECT =
	"id,seriesMasterId,type,subject,start,end,location,attendees,organizer,bodyPreview,isOnlineMeeting,onlineMeetingUrl,categories,recurrence,isCancelled,responseStatus";

// Microsoft Graph's OData $filter interprets a date-only string like
// "2026-05-19" as midnight UTC at the START of that day — so an end_date of
// "2026-05-19" silently EXCLUDES every event on May 19. v2.2.3 fixes this by
// expanding date-only inputs to inclusive day boundaries before they reach
// the filter. Strings that already carry a time component pass through.
const DATE_ONLY = /^\d{4}-\d{2}-\d{2}$/;

export function normaliseStartDate(input: string): string {
	if (DATE_ONLY.test(input)) return `${input}T00:00:00.000Z`;
	return input;
}

export function normaliseEndDate(input: string): string {
	if (DATE_ONLY.test(input)) return `${input}T23:59:59.999Z`;
	return input;
}

const eventIdSchema = z.string().min(1).max(512);

// ── Impls ──────────────────────────────────────────────────────────────────────

async function listCalendarEventsImpl(
	env: Env,
	args: {
		start_date?: string;
		end_date?: string;
		count?: number;
		subject_contains?: string;
	},
): Promise<unknown> {
	const now = new Date();
	// Default window: 7 days past → 7 days future. Symmetric around "now" so
	// both "what just happened?" and "what's coming up?" queries hit. The old
	// default was now → now+7d (future only), which silently dropped past
	// events — a Claude asking for "the most recent meeting" would never see
	// anything that had already ended. v2.2.2 fixed this.
	const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
	const weekLater = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
	const start = args.start_date
		? normaliseStartDate(args.start_date)
		: weekAgo.toISOString();
	const end = args.end_date
		? normaliseEndDate(args.end_date)
		: weekLater.toISOString();
	// Default raised from 20 → 100 in v2.2.0. The old default silently truncated
	// wide-range queries; Claude has no way to know it missed events on later
	// pages. 100 covers ~95% of practical "search across the last month" cases.
	const count = args.count ?? 100;

	const dateFilter = `start/dateTime ge '${start}' and start/dateTime le '${end}'`;
	const subjectFilter = args.subject_contains
		? ` and contains(subject,'${escapeOdataString(args.subject_contains)}')`
		: "";

	const data = (await graphGet(env, "/me/events", {
		$select: EVENT_SELECT,
		$top: count,
		$filter: `${dateFilter}${subjectFilter}`,
		$orderby: "start/dateTime",
	})) as { value: unknown[] };
	return sanitizeEventList(data.value);
}

export async function createCalendarEventImpl(
	env: Env,
	args: {
		subject: string;
		start_datetime: string;
		end_datetime: string;
		timezone?: string;
		location?: string;
		description?: string;
		body_type?: string;
		attendees?: string;
		is_online_meeting?: boolean;
		categories?: string[];
		recurrence?: RecurrenceInput;
		include_signature?: boolean;
	},
): Promise<unknown> {
	const timezone = args.timezone ?? "UTC";
	const attendeeList = args.attendees
		? args.attendees
				.split(",")
				.map((e) => e.trim())
				.filter(Boolean)
		: [];

	const event: Record<string, unknown> = {
		subject: args.subject,
		start: { dateTime: args.start_datetime, timeZone: timezone },
		end: { dateTime: args.end_datetime, timeZone: timezone },
	};

	if (args.location) event.location = { displayName: args.location };

	// Signature path replaces the plain body build: like email, the signature is
	// concatenated after sanitisation, so it cannot go through buildMessageBody.
	const notes: string[] = [];
	const signed = maybeSign(env, args.include_signature, args.description, args.body_type);
	if (signed) {
		event.body = { contentType: signed.contentType, content: signed.content };
		notes.push(...signed.notes);
	} else if (args.description) {
		event.body = buildMessageBody(args.description, args.body_type);
	}

	if (attendeeList.length > 0) {
		event.attendees = attendeeList.map((email) => ({
			emailAddress: { address: email },
			type: "required",
		}));
	}
	if (args.is_online_meeting) {
		event.isOnlineMeeting = true;
		event.onlineMeetingProvider = "teamsForBusiness";
	}
	if (args.categories && args.categories.length > 0) {
		event.categories = args.categories;
	}
	if (args.recurrence) {
		event.recurrence = buildRecurrence(args.recurrence, args.start_datetime, timezone);
	}

	const data = await graphPost(env, "/me/events", event);
	const result: Record<string, unknown> = {
		success: true,
		message: "Event created.",
		event: data,
	};
	if (notes.length > 0) result.notes = notes;
	return result;
}

export async function updateCalendarEventImpl(
	env: Env,
	args: {
		id: string;
		subject?: string;
		start_datetime?: string;
		end_datetime?: string;
		timezone?: string;
		location?: string;
		description?: string;
		body_type?: string;
		attendees?: string;
		categories?: string[];
		include_signature?: boolean;
	},
): Promise<unknown> {
	const timezone = args.timezone ?? "UTC";
	const updates: Record<string, unknown> = {};
	const notes: string[] = [];

	if (args.subject) updates.subject = args.subject;
	if (args.start_datetime)
		updates.start = { dateTime: args.start_datetime, timeZone: timezone };
	if (args.end_datetime)
		updates.end = { dateTime: args.end_datetime, timeZone: timezone };
	if (args.location) updates.location = { displayName: args.location };
	// Sign only when a new description is supplied. This is a partial update, so
	// signing with no description would build a signature-only body and replace
	// the event's existing description.
	const signBody = args.include_signature === true && !!args.description;
	if (args.include_signature === true && !args.description) {
		notes.push(
			"include_signature was ignored: it only applies when you also provide a new description. " +
				"The existing event description was left unchanged.",
		);
	}
	if (signBody) {
		const signed = maybeSign(env, true, args.description, args.body_type);
		updates.body = signed
			? { contentType: signed.contentType, content: signed.content }
			: buildMessageBody(args.description!, args.body_type);
		if (signed) notes.push(...signed.notes);
	} else if (args.description) {
		updates.body = buildMessageBody(args.description, args.body_type);
	}
	if (args.attendees) {
		updates.attendees = args.attendees
			.split(",")
			.map((e) => e.trim())
			.filter(Boolean)
			.map((address) => ({ emailAddress: { address }, type: "required" }));
	}
	if (args.categories !== undefined) {
		updates.categories = args.categories;
	}

	if (Object.keys(updates).length === 0) {
		throw ToolError.validation(
			"No fields provided to update. Pass at least one of: subject, start_datetime, end_datetime, location, description, attendees, categories.",
		);
	}

	await graphPatch(env, `/me/events/${args.id}`, updates);
	const result: Record<string, unknown> = { success: true, message: "Event updated." };
	if (notes.length > 0) result.notes = notes;
	return result;
}

async function deleteCalendarEventImpl(
	env: Env,
	args: { id: string },
): Promise<unknown> {
	await graphDelete(env, `/me/events/${args.id}`);
	return { success: true, message: "Event deleted." };
}

async function cancelCalendarEventImpl(
	env: Env,
	args: { id: string; comment?: string },
): Promise<unknown> {
	const body = args.comment ? { comment: args.comment } : {};
	await graphPost(env, `/me/events/${args.id}/cancel`, body);
	auditLog("event_cancelled", { event_id: args.id, has_comment: !!args.comment });
	return { success: true, message: "Event cancelled and attendees notified." };
}

async function respondToEventImpl(
	env: Env,
	args: {
		id: string;
		response: "accept" | "decline" | "tentatively_accept";
		comment?: string;
		send_response?: boolean;
	},
): Promise<unknown> {
	const endpoint =
		args.response === "accept"
			? "accept"
			: args.response === "decline"
				? "decline"
				: "tentativelyAccept";

	const body: Record<string, unknown> = {
		sendResponse: args.send_response ?? true,
	};
	if (args.comment) body.comment = args.comment;

	await graphPost(env, `/me/events/${args.id}/${endpoint}`, body);

	auditLog("event_response", {
		event_id: args.id,
		response: args.response,
		send_response: body.sendResponse,
	});
	return { success: true, message: `Responded with: ${args.response}` };
}

async function listEventOccurrencesImpl(
	env: Env,
	args: { series_id: string; start_date: string; end_date: string; count?: number },
): Promise<unknown> {
	const count = args.count ?? 50;
	const data = (await graphGet(env, `/me/events/${args.series_id}/instances`, {
		startDateTime: normaliseStartDate(args.start_date),
		endDateTime: normaliseEndDate(args.end_date),
		$select: EVENT_SELECT,
		$top: count,
		$orderby: "start/dateTime",
	})) as { value: unknown[] };

	return {
		series_id: args.series_id,
		count: data.value.length,
		occurrences: sanitizeEventList(data.value),
	};
}

// ── defineTools map ────────────────────────────────────────────────────────────

export const calendarTools = defineTools<Env>({
	list_calendar_events: {
		description:
			"Retrieve calendar events within a date range. Default window when no dates are supplied is **7 days past → 7 days future** (symmetric around now) — so 'most recent meeting' and 'next meeting' queries both work without prompting. Default count is 100 (max 200). Use `subject_contains` to filter server-side by substring match on the event subject (case-insensitive) — much more efficient than fetching everything and grepping client-side.",
		schema: z.object({
			start_date: z.string().min(1).max(64).optional(),
			end_date: z.string().min(1).max(64).optional(),
			count: z.number().int().min(1).max(200).optional(),
			subject_contains: z.string().min(1).max(256).optional(),
		}),
		handler: (env, args) => listCalendarEventsImpl(env, args),
	},

	create_calendar_event: {
		description:
			"Create a new calendar event. Supports rich-text descriptions (HTML), recurring schedules, categories, and Teams meetings. " +
			BODY_TYPE_DESC,
		schema: z.object({
			subject: z.string().min(1).max(256),
			start_datetime: z.string().min(1).max(64),
			end_datetime: z.string().min(1).max(64),
			timezone: z.string().min(1).max(64).optional(),
			location: z.string().max(256).optional(),
			description: z.string().max(2 * 1024 * 1024).optional(),
			body_type: z.enum(["text", "html"]).optional(),
			attendees: z.string().max(4096).optional(),
			is_online_meeting: z.boolean().optional(),
			categories: z.array(z.string().min(1).max(64)).max(32).optional(),
			recurrence: recurrenceSchema.optional(),
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
		}),
		handler: (env, args) => createCalendarEventImpl(env, args),
	},

	update_calendar_event: {
		description:
			"Update an existing calendar event. Only provide the fields you want to change. To remove an attendee, provide the full updated attendee list without them. To edit a single occurrence of a recurring series, pass the occurrence ID (use list_event_occurrences to find it). To edit the whole series, pass the seriesMasterId. Empty calls (no fields provided) are rejected.",
		schema: z.object({
			id: eventIdSchema,
			subject: z.string().min(1).max(256).optional(),
			start_datetime: z.string().min(1).max(64).optional(),
			end_datetime: z.string().min(1).max(64).optional(),
			timezone: z.string().min(1).max(64).optional(),
			location: z.string().max(256).optional(),
			description: z.string().max(2 * 1024 * 1024).optional(),
			body_type: z.enum(["text", "html"]).optional(),
			attendees: z.string().max(4096).optional(),
			categories: z.array(z.string().min(1).max(64)).max(32).optional(),
			include_signature: includeSignatureSchema.describe(INCLUDE_SIGNATURE_DESC),
		}),
		handler: (env, args) => updateCalendarEventImpl(env, args),
	},

	delete_calendar_event: {
		description:
			"Delete a calendar event by its ID. Removes the event silently without sending notifications. To formally cancel a meeting and notify attendees, use cancel_calendar_event instead.",
		schema: z.object({ id: eventIdSchema }),
		handler: (env, args) => deleteCalendarEventImpl(env, args),
	},

	cancel_calendar_event: {
		description:
			"Formally cancel a meeting you organised. Sends cancellation notices to all attendees with an optional comment. Only works for events you are the organizer of — to remove an event you do not organise, use delete_calendar_event.",
		schema: z.object({
			id: eventIdSchema,
			comment: z.string().max(2048).optional(),
		}),
		handler: (env, args) => cancelCalendarEventImpl(env, args),
	},

	respond_to_event: {
		description:
			"Accept, decline, or tentatively accept a meeting invite. The event must be one you have been invited to (you are an attendee, not the organiser).",
		schema: z.object({
			id: eventIdSchema,
			response: z.enum(["accept", "decline", "tentatively_accept"]),
			comment: z.string().max(2048).optional(),
			send_response: z.boolean().optional(),
		}),
		handler: (env, args) => respondToEventImpl(env, args),
	},

	list_event_occurrences: {
		description:
			"List the individual occurrences of a recurring event series within a date range. Use this to find the occurrence ID for editing or cancelling a single instance of a recurring event.",
		schema: z.object({
			series_id: eventIdSchema,
			start_date: z.string().min(1).max(64),
			end_date: z.string().min(1).max(64),
			count: z.number().int().min(1).max(500).optional(),
		}),
		handler: (env, args) => listEventOccurrencesImpl(env, args),
	},
});
