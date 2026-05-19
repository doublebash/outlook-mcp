// WebVTT parser for Microsoft Teams meeting transcripts.
//
// Microsoft Graph's `/onlineMeetings/{id}/transcripts/{id}/content` endpoint
// returns text/vtt only (the docx format was deprecated 2023-05; JSON has never
// been supported). Callers who want structured speaker/turn data have to parse
// VTT locally — this helper provides that.
//
// Output shape:
//   {
//     count: number,
//     cues: Array<{
//       start: number       // seconds (may be negative; see below)
//       end: number         // seconds
//       speaker: string|null   // from VTT voice tag <v Speaker Name>
//       text: string           // inline tags stripped
//     }>
//   }
//
// Negative timestamps: per Graph docs, "negative offsets indicate that the
// transcription began while the conversation was ongoing." We surface them
// as-is — callers can decide whether to clamp to 0.

export interface TranscriptCue {
	start: number;
	end: number;
	speaker: string | null;
	text: string;
}

export interface TranscriptParse {
	count: number;
	cues: TranscriptCue[];
}

// Matches a WebVTT timing line. Hours portion is optional. Allows an optional
// trailing settings clause (e.g. `align:start position:0%`).
//   00:00:05.000 --> 00:00:10.000
//   01:23:45.123 --> 01:23:46.456 align:start
//   -00:00:30.000 --> 00:00:00.000   (negative offset)
const TIMING_LINE =
	/^(-?(?:\d{1,2}:)?\d{1,2}:\d{1,2}\.\d{1,3})\s*-->\s*(-?(?:\d{1,2}:)?\d{1,2}:\d{1,2}\.\d{1,3})(?:\s+\S.*)?$/;

// Voice tag: `<v Speaker Name>` or `<v.class1.class2 Speaker Name>`.
// The class portion is allowed but ignored.
const VOICE_TAG = /<v(?:\.[^ >]*)? +([^>]+)>/;

// Strip every inline tag (including the voice tag once we've extracted speaker,
// closing `</v>`, timestamp markers `<00:00:05.000>`, style markers `<c.color>`,
// etc.). Caller-visible cue text is plain prose.
const ALL_INLINE_TAGS = /<[^>]+>/g;

// Parse `HH:MM:SS.mmm` or `MM:SS.mmm`, with optional leading `-`.
// Returns seconds as a float, or NaN if the input is malformed.
function parseTime(s: string): number {
	const m = s.match(/^(-?)(?:(\d+):)?(\d+):(\d+)\.(\d+)$/);
	if (!m) return Number.NaN;
	const sign = m[1] === "-" ? -1 : 1;
	const h = m[2] ? Number.parseInt(m[2], 10) : 0;
	const minutes = Number.parseInt(m[3] ?? "0", 10);
	const seconds = Number.parseInt(m[4] ?? "0", 10);
	const msRaw = m[5] ?? "0";
	const fractional = Number.parseInt(msRaw, 10) / 10 ** msRaw.length;
	return sign * (h * 3600 + minutes * 60 + seconds + fractional);
}

export function parseVtt(vtt: string): TranscriptParse {
	// Normalise line endings.
	const normalised = vtt.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
	// Split into blocks on one or more blank lines.
	const blocks = normalised.split(/\n[ \t]*\n/);

	const cues: TranscriptCue[] = [];

	for (const blockRaw of blocks) {
		const block = blockRaw.trim();
		if (!block) continue;
		// Skip the WEBVTT header (always the first block).
		if (/^WEBVTT(\s|$)/.test(block)) continue;
		// Skip metadata-only blocks. These can also be cue text if a line in
		// the cue starts with these words, but in transcripts they're never
		// the leading line of a real cue.
		if (/^(NOTE|STYLE|REGION)(\s|$)/.test(block)) continue;

		const lines = block.split("\n");

		// Find the timing line. A cue block may have an optional identifier
		// line BEFORE the timing line, so we scan rather than assume index 0.
		let timingIdx = -1;
		let timingMatch: RegExpMatchArray | null = null;
		for (let i = 0; i < lines.length; i++) {
			const line = lines[i];
			if (line === undefined) continue;
			const m = line.match(TIMING_LINE);
			if (m) {
				timingIdx = i;
				timingMatch = m;
				break;
			}
		}
		if (timingIdx === -1 || !timingMatch) continue;

		const startStr = timingMatch[1] ?? "";
		const endStr = timingMatch[2] ?? "";
		const start = parseTime(startStr);
		const end = parseTime(endStr);
		if (Number.isNaN(start) || Number.isNaN(end)) continue;

		const textLines = lines.slice(timingIdx + 1);
		if (textLines.length === 0) continue;

		const joined = textLines.join("\n");
		let speaker: string | null = null;
		const voiceMatch = joined.match(VOICE_TAG);
		if (voiceMatch?.[1]) {
			// Trim whitespace and any trailing slashes from self-closing forms.
			speaker = voiceMatch[1].trim().replace(/\/+$/, "") || null;
		}

		const cleaned = joined.replace(ALL_INLINE_TAGS, "").trim();
		if (!cleaned) continue;

		cues.push({ start, end, speaker, text: cleaned });
	}

	return { count: cues.length, cues };
}
