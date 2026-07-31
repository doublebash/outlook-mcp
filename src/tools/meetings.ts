import { ToolError, defineTools } from "@bashco/mcp-toolkit";
import { z } from "zod";
import { graphGet, graphRequestRaw } from "../graph.js";
import { parseVtt } from "../transcript-helpers.js";
import type { Env } from "../types.js";
import { auditLog, escapeOdataString } from "./_shared.js";

// Three sources of truth for which meeting to query. Exactly one must be provided.
//
// - `meeting_id` — the onlineMeeting id (use when known)
// - `calendar_event_id` — a calendar event id (Graph resolves to onlineMeeting via joinUrl)
// - `join_url` — the Teams meeting join URL (Graph resolves via JoinWebUrl filter).
//   This is the canonical handle for ad-hoc / Meet-now / "Start recording now" calls
//   that don't have a calendar event attached.
const meetingRefSchema = z
	.object({
		meeting_id: z.string().min(1).max(512).optional(),
		calendar_event_id: z.string().min(1).max(512).optional(),
		join_url: z.string().url().max(2048).optional(),
	})
	.refine(
		(v) => [v.meeting_id, v.calendar_event_id, v.join_url].filter(Boolean).length === 1,
		"Provide exactly one of: meeting_id, calendar_event_id, join_url.",
	);

type MeetingRef = {
	meeting_id?: string;
	calendar_event_id?: string;
	join_url?: string;
};

// ── onlineMeeting id resolution ───────────────────────────────────────────────
// Graph's recording/transcript endpoints want an onlineMeeting id, which is
// distinct from a calendarEvent id and from the Teams join URL. The supported
// translation is to filter onlineMeetings by JoinWebUrl. Calendar events expose
// their join URL via `onlineMeetingUrl` (or `onlineMeeting.joinUrl`), so the
// event path is event → joinUrl → JoinWebUrl filter.
//
// Ad-hoc / Meet-now calls don't create a calendar event; the join URL is then
// the canonical handle (copied from the Teams chat thread for the call).

interface GraphEventOnlineMeeting {
	isOnlineMeeting?: boolean;
	onlineMeetingUrl?: string;
	onlineMeeting?: { joinUrl?: string };
	bodyPreview?: string;
}

// Pulls a Teams join URL out of an event's bodyPreview text. Used as a
// fallback when Microsoft Graph's structured fields are empty — a real
// quirk observed in production: events created via the legacy Outlook Teams
// add-in (or some mobile flows) can have `isOnlineMeeting: true` while both
// `onlineMeetingUrl` and `onlineMeeting.joinUrl` are null; the URL only
// lives as plain text in the body.
//
// Matches both:
//   - new short format: https://teams.microsoft.com/meet/<code>?p=<passcode>
//   - legacy long form: https://teams.microsoft.com/l/meetup-join/<thread>@thread.v2/0?context=...
//   - personal account variant: https://teams.live.com/meet/...
//
// Whitespace, `<`, `>`, `"`, `'` and end-of-string close the URL. Returns
// the first match (Teams meeting bodies typically contain one canonical
// join URL plus optional non-meeting links like "Meeting options" that
// don't match these path prefixes).
export function extractJoinUrlFromBody(body: string | undefined): string | null {
	if (!body) return null;
	const re =
		/https?:\/\/(?:teams\.microsoft\.com|teams\.live\.com)\/(?:l\/meetup-join|meet)\/[^\s<>"']+/;
	const match = body.match(re);
	return match ? match[0] : null;
}

// Pull the join URL off a Graph event, trying the structured fields first
// (cheapest, exact) and falling back to body-preview parsing. Returns null
// if neither path yields a URL.
function pickJoinUrlFromEvent(
	event: GraphEventOnlineMeeting,
): string | null {
	if (!event.isOnlineMeeting) return null;
	const structured = event.onlineMeetingUrl ?? event.onlineMeeting?.joinUrl;
	if (structured) return structured;
	return extractJoinUrlFromBody(event.bodyPreview);
}

async function resolveJoinUrl(env: Env, ref: MeetingRef): Promise<string> {
	if (ref.join_url) return ref.join_url.trim();
	if (!ref.calendar_event_id) {
		throw ToolError.validation(
			"Provide exactly one of: meeting_id, calendar_event_id, join_url.",
		);
	}

	const event = (await graphGet(env, `/me/events/${ref.calendar_event_id}`, {
		$select: "isOnlineMeeting,onlineMeetingUrl,onlineMeeting,bodyPreview",
	})) as GraphEventOnlineMeeting;

	const joinUrl = pickJoinUrlFromEvent(event);
	if (!joinUrl) {
		throw ToolError.validation(
			event.isOnlineMeeting
				? "Calendar event is marked as a Teams online meeting but no joinUrl could be found in either the structured fields (onlineMeetingUrl, onlineMeeting.joinUrl) or the body preview. The Teams binding may be incomplete — try passing join_url directly if you can copy the meeting link from elsewhere."
				: "Calendar event is not a Teams online meeting. Recordings and transcripts are only available for Teams meetings.",
		);
	}
	return joinUrl;
}

// Shared lookup: given a join URL, find the corresponding onlineMeeting id.
// Returns null if no match (rather than throwing) so bulk callers can skip
// individual events without aborting the whole query.
async function lookupOnlineMeetingIdByJoinUrl(
	env: Env,
	joinUrl: string,
): Promise<string | null> {
	const filtered = (await graphGet(env, "/me/onlineMeetings", {
		$filter: `JoinWebUrl eq '${escapeOdataString(joinUrl)}'`,
	})) as { value: Array<{ id?: string }> };
	return filtered.value?.[0]?.id ?? null;
}

async function resolveMeetingId(env: Env, ref: MeetingRef): Promise<string> {
	if (ref.meeting_id) return ref.meeting_id;

	const joinUrl = await resolveJoinUrl(env, ref);
	const onlineMeetingId = await lookupOnlineMeetingIdByJoinUrl(env, joinUrl);
	if (!onlineMeetingId) {
		throw ToolError.validation(
			ref.join_url
				? "Could not resolve onlineMeeting from join URL. The URL may be wrong, the meeting may have been deleted, or you may not have permission. The URL must match Microsoft's stored JoinWebUrl exactly — including the `?context=…` query string. Copy the link directly from the Teams chat for the call."
				: "Could not resolve onlineMeeting from calendar event. The meeting may not have started yet, or the join URL has rotated. Try again with meeting_id or join_url.",
		);
	}
	return onlineMeetingId;
}

// ── Sanitizers ────────────────────────────────────────────────────────────────
// Whitelist fields Graph returns; everything else stays internal.

interface GraphRecording {
	id?: string;
	meetingId?: string;
	callId?: string;
	contentCorrelationId?: string;
	recordingContentUrl?: string;
	createdDateTime?: string;
	endDateTime?: string;
	meetingOrganizer?: { user?: { id?: string; displayName?: string } };
}

function pickRecording(r: GraphRecording): Record<string, unknown> {
	return {
		id: r.id,
		meeting_id: r.meetingId,
		call_id: r.callId,
		created_at: r.createdDateTime,
		ended_at: r.endDateTime,
		content_url: r.recordingContentUrl,
		organizer: r.meetingOrganizer?.user?.displayName,
	};
}

interface GraphTranscript {
	id?: string;
	meetingId?: string;
	transcriptContentUrl?: string;
	createdDateTime?: string;
	meetingOrganizer?: { user?: { id?: string; displayName?: string } };
}

function pickTranscript(t: GraphTranscript): Record<string, unknown> {
	return {
		id: t.id,
		meeting_id: t.meetingId,
		created_at: t.createdDateTime,
		content_url: t.transcriptContentUrl,
		organizer: t.meetingOrganizer?.user?.displayName,
	};
}

interface GraphOnlineMeeting {
	id?: string;
	subject?: string;
	startDateTime?: string;
	endDateTime?: string;
	joinWebUrl?: string;
	creationDateTime?: string;
}

function pickOnlineMeeting(m: GraphOnlineMeeting): Record<string, unknown> {
	return {
		id: m.id,
		subject: m.subject,
		start: m.startDateTime,
		end: m.endDateTime,
		join_url: m.joinWebUrl,
		created_at: m.creationDateTime,
	};
}

// ── Impls ──────────────────────────────────────────────────────────────────────

async function findOnlineMeetingImpl(env: Env, args: MeetingRef): Promise<unknown> {
	const meetingId = await resolveMeetingId(env, args);
	const meeting = (await graphGet(env, `/me/onlineMeetings/${meetingId}`, {
		$select: "id,subject,startDateTime,endDateTime,joinWebUrl,creationDateTime",
	})) as GraphOnlineMeeting;

	auditLog("online_meeting_found", { meeting_id: meetingId });
	return pickOnlineMeeting(meeting);
}

async function listMeetingRecordingsImpl(env: Env, args: MeetingRef): Promise<unknown> {
	const meetingId = await resolveMeetingId(env, args);
	const data = (await graphGet(
		env,
		`/me/onlineMeetings/${meetingId}/recordings`,
	)) as { value: GraphRecording[] };

	auditLog("meeting_recordings_listed", {
		meeting_id: meetingId,
		count: data.value?.length ?? 0,
	});

	return {
		meeting_id: meetingId,
		count: data.value?.length ?? 0,
		recordings: (data.value ?? []).map(pickRecording),
	};
}

async function listMeetingTranscriptsImpl(env: Env, args: MeetingRef): Promise<unknown> {
	const meetingId = await resolveMeetingId(env, args);
	const data = (await graphGet(
		env,
		`/me/onlineMeetings/${meetingId}/transcripts`,
	)) as { value: GraphTranscript[] };

	auditLog("meeting_transcripts_listed", {
		meeting_id: meetingId,
		count: data.value?.length ?? 0,
	});

	return {
		meeting_id: meetingId,
		count: data.value?.length ?? 0,
		transcripts: (data.value ?? []).map(pickTranscript),
	};
}

async function getTranscriptContentImpl(
	env: Env,
	args: MeetingRef & { transcript_id: string; format?: "vtt" | "json" },
): Promise<unknown> {
	const meetingId = await resolveMeetingId(env, args);
	const format = args.format ?? "vtt";

	// Microsoft Graph's transcript content endpoint only returns text/vtt
	// (docx was deprecated 2023-05; JSON has never been supported — passing
	// `?$format=application/json` returns a 400). We always fetch VTT, and if
	// the caller asked for JSON we parse it locally into a structured
	// {cues:[{start,end,speaker,text}]} shape.
	const response = await graphRequestRaw(
		env,
		`/me/onlineMeetings/${meetingId}/transcripts/${args.transcript_id}/content?$format=text/vtt`,
	);
	const vtt = await response.text();

	auditLog("meeting_transcript_fetched", {
		meeting_id: meetingId,
		transcript_id: args.transcript_id,
		format,
		bytes: vtt.length,
	});

	if (format === "json") {
		const parsed = parseVtt(vtt);
		return {
			meeting_id: meetingId,
			transcript_id: args.transcript_id,
			format,
			...parsed,
		};
	}

	return {
		meeting_id: meetingId,
		transcript_id: args.transcript_id,
		format,
		content: vtt,
	};
}

// ── Recent recordings discovery ───────────────────────────────────────────────
// Lists recent calendar events that are Teams online meetings, fans out to
// `/recordings` and `/transcripts` per event in parallel, and returns the
// subset that actually have content. Avoids forcing the caller to chain
// list_calendar_events → list_meeting_recordings(event_id) → ... themselves.

interface GraphEventForDiscovery {
	id?: string;
	subject?: string;
	start?: { dateTime?: string; timeZone?: string };
	end?: { dateTime?: string; timeZone?: string };
	isOnlineMeeting?: boolean;
	onlineMeetingUrl?: string;
	onlineMeeting?: { joinUrl?: string };
	bodyPreview?: string;
	organizer?: { emailAddress?: { name?: string; address?: string } };
}

interface DiscoveryRow {
	subject: string | undefined;
	recorded_at: string | undefined;
	start: string | undefined;
	end: string | undefined;
	duration_minutes: number | undefined;
	calendar_event_id: string | undefined;
	meeting_id: string;
	recording_count: number;
	transcript_count: number;
	organizer: string | undefined;
}

function computeDurationMinutes(
	startIso: string | undefined,
	endIso: string | undefined,
): number | undefined {
	if (!startIso || !endIso) return undefined;
	const startMs = Date.parse(startIso);
	const endMs = Date.parse(endIso);
	if (Number.isNaN(startMs) || Number.isNaN(endMs)) return undefined;
	return Math.round((endMs - startMs) / 60000);
}

const DISCOVERY_PAGE_SIZE = 200;
const DISCOVERY_MAX_PAGES = 10;
const DISCOVERY_MAX_CANDIDATES = 200;

// Collect Teams events in the window, newest first.
//
// `isOnlineMeeting` is not a filterable property — putting it in $filter makes
// Graph reject the whole request with
// `ErrorInvalidProperty: The property 'isOnlineMeeting' does not support
// filtering`. So the date range is filtered server-side and the online-meeting
// test happens here instead.
//
// That shifts what $top bounds: it now counts every event in the window rather
// than only Teams ones, so a single page could be entirely all-day blocks and
// 1:1s while real meetings sit just past the cut. Page through on $skiptoken
// until there are enough candidates, capping pages so a long window on a busy
// calendar can't run unbounded.
async function fetchOnlineMeetingCandidates(
	env: Env,
	startIso: string,
	endIso: string,
): Promise<{ candidates: GraphEventForDiscovery[]; scanned: number }> {
	// bodyPreview is selected so we can fall back to URL extraction when the
	// structured Teams fields are empty (Outlook add-in quirk).
	const baseQuery: Record<string, string | number> = {
		$select:
			"id,subject,start,end,isOnlineMeeting,onlineMeetingUrl,onlineMeeting,bodyPreview,organizer",
		$top: DISCOVERY_PAGE_SIZE,
		$filter: `start/dateTime ge '${startIso}' and start/dateTime le '${endIso}'`,
		$orderby: "start/dateTime desc",
	};

	const candidates: GraphEventForDiscovery[] = [];
	let query: Record<string, string | number> = baseQuery;
	let scanned = 0;

	for (let page = 0; page < DISCOVERY_MAX_PAGES; page++) {
		const resp = (await graphGet(env, "/me/events", query)) as {
			value?: GraphEventForDiscovery[];
			"@odata.nextLink"?: string;
		};

		const events = resp.value ?? [];
		scanned += events.length;
		for (const event of events) {
			if (event.isOnlineMeeting) candidates.push(event);
		}

		if (candidates.length >= DISCOVERY_MAX_CANDIDATES) break;

		// Graph hands back an absolute nextLink; reuse just its $skiptoken so the
		// request keeps going through the same query-building path.
		const nextLink = resp["@odata.nextLink"];
		if (!nextLink) break;
		let skipToken: string | null = null;
		try {
			skipToken = new URL(nextLink).searchParams.get("$skiptoken");
		} catch {
			skipToken = null;
		}
		if (!skipToken) break;
		query = { ...baseQuery, $skiptoken: skipToken };
	}

	return {
		candidates: candidates.slice(0, DISCOVERY_MAX_CANDIDATES),
		scanned,
	};
}

export async function listRecentMeetingRecordingsImpl(
	env: Env,
	args: { within_days?: number; limit?: number },
): Promise<unknown> {
	const withinDays = args.within_days ?? 14;
	const limit = args.limit ?? 10;

	const now = Date.now();
	// Look backwards `withinDays` and a tiny 1-day forward buffer (events that
	// started today and run into tonight should still surface).
	const startIso = new Date(now - withinDays * 86_400_000).toISOString();
	const endIso = new Date(now + 86_400_000).toISOString();

	const { candidates: events, scanned } = await fetchOnlineMeetingCandidates(
		env,
		startIso,
		endIso,
	);

	// Per-event work: resolve onlineMeeting id, fetch recordings + transcripts.
	// One failed event must not abort the whole call — return null and filter.
	const rows = await Promise.all(
		events.map(async (event): Promise<DiscoveryRow | null> => {
			const joinUrl = pickJoinUrlFromEvent(event);
			if (!joinUrl) return null;

			try {
				const meetingId = await lookupOnlineMeetingIdByJoinUrl(env, joinUrl);
				if (!meetingId) return null;

				const [recordingsResp, transcriptsResp] = await Promise.all([
					graphGet(
						env,
						`/me/onlineMeetings/${meetingId}/recordings`,
					) as Promise<{ value: GraphRecording[] }>,
					graphGet(
						env,
						`/me/onlineMeetings/${meetingId}/transcripts`,
					) as Promise<{ value: GraphTranscript[] }>,
				]);

				const recordings = recordingsResp.value ?? [];
				const transcripts = transcriptsResp.value ?? [];
				if (recordings.length === 0 && transcripts.length === 0) {
					return null;
				}

				// Sort key — prefer the latest recording's createdDateTime;
				// fall back to latest transcript; final fallback is event end.
				const latestRecordingAt = recordings
					.map((r) => r.createdDateTime)
					.filter((t): t is string => Boolean(t))
					.sort()
					.pop();
				const latestTranscriptAt = transcripts
					.map((t) => t.createdDateTime)
					.filter((t): t is string => Boolean(t))
					.sort()
					.pop();
				const recordedAt =
					latestRecordingAt ?? latestTranscriptAt ?? event.end?.dateTime;

				return {
					subject: event.subject,
					recorded_at: recordedAt,
					start: event.start?.dateTime,
					end: event.end?.dateTime,
					duration_minutes: computeDurationMinutes(
						event.start?.dateTime,
						event.end?.dateTime,
					),
					calendar_event_id: event.id,
					meeting_id: meetingId,
					recording_count: recordings.length,
					transcript_count: transcripts.length,
					organizer: event.organizer?.emailAddress?.name,
				};
			} catch {
				// Per-event failure (event without onlineMeeting binding, permission
				// glitch, etc.) — skip and keep going.
				return null;
			}
		}),
	);

	const withContent = rows.filter((r): r is DiscoveryRow => r !== null);
	withContent.sort((a, b) =>
		(b.recorded_at ?? "").localeCompare(a.recorded_at ?? ""),
	);
	const top = withContent.slice(0, limit);

	auditLog("recent_meeting_recordings_listed", {
		within_days: withinDays,
		events_scanned: scanned,
		online_meetings: events.length,
		with_content: withContent.length,
		returned: top.length,
	});

	return {
		within_days: withinDays,
		events_scanned: scanned,
		online_meetings: events.length,
		count: top.length,
		recordings: top,
	};
}

// ── defineTools map ────────────────────────────────────────────────────────────

export const meetingsTools = defineTools<Env>({
	list_recent_meeting_recordings: {
		description:
			"Find recent Teams meetings that have recordings or transcripts. **Use this as the first step** when the user asks about 'recent recordings', 'my latest meeting transcript', 'last week's standups', or anything similar — instead of trying to find a specific meeting via list_calendar_events. Scans calendar events in the past N days (default 14, max 90), checks each for recordings/transcripts, returns those with at least one, sorted newest first. The `calendar_event_id` in each result can be passed straight to list_meeting_transcripts or get_transcript_content. Only surfaces meetings the current user organised or attended where they have permission to read recordings/transcripts.",
		schema: z.object({
			within_days: z.number().int().min(1).max(90).optional(),
			limit: z.number().int().min(1).max(50).optional(),
		}),
		handler: (env, args) => listRecentMeetingRecordingsImpl(env, args),
	},

	find_online_meeting: {
		description:
			"Resolve a Teams onlineMeeting from any of meeting_id, calendar_event_id, or join_url. Returns the meeting's id, subject, start/end time, join URL, and creation time. Useful as a 'does this meeting exist and am I allowed to read it?' check before calling list_meeting_recordings or list_meeting_transcripts. For ad-hoc / Meet-now calls (no calendar event), pass join_url copied from the Teams chat thread for the call.",
		schema: meetingRefSchema,
		handler: (env, args) => findOnlineMeetingImpl(env, args),
	},

	list_meeting_recordings: {
		description:
			"List Teams meeting recordings for an online meeting. Pass exactly one of: meeting_id (the onlineMeeting id), calendar_event_id (auto-resolved via the event's join URL), or join_url (for ad-hoc / Meet-now calls that have no calendar event). Recording content URLs returned here are short-lived — call this tool again to refresh them. Only works for Teams meetings, not external meeting providers.",
		schema: meetingRefSchema,
		handler: (env, args) => listMeetingRecordingsImpl(env, args),
	},

	list_meeting_transcripts: {
		description:
			"List transcripts for a Teams meeting. Pass exactly one of: meeting_id, calendar_event_id, or join_url (the Teams meeting link — useful for Meet-now calls with no calendar event). Returns transcript metadata only — use get_transcript_content to fetch the actual text. Transcripts are only generated if live captions / transcript was on during the meeting; they typically appear 5–15 minutes after the call ends.",
		schema: meetingRefSchema,
		handler: (env, args) => listMeetingTranscriptsImpl(env, args),
	},

	get_transcript_content: {
		description:
			'Fetch the full text of a Teams meeting transcript. Pass one of meeting_id, calendar_event_id, or join_url, plus the transcript_id from list_meeting_transcripts. Format defaults to "vtt" (WebVTT — raw caption text with timestamps and speaker tags). Use "json" for a structured shape: {count, cues: [{start, end, speaker, text}]} with timestamps as seconds (float, possibly negative if the transcription started mid-call) and inline tags stripped from text.',
		schema: meetingRefSchema.and(
			z.object({
				transcript_id: z.string().min(1).max(512),
				format: z.enum(["vtt", "json"]).optional(),
			}),
		),
		handler: (env, args) => getTranscriptContentImpl(env, args),
	},
});
