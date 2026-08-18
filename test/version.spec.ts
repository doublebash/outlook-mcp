import { describe, expect, it } from "vitest";
import { SERVER_VERSION } from "../src/index.js";
import pkg from "../package.json";

describe("SERVER_VERSION", () => {
	it("matches package.json", () => {
		// This drifted silently for four releases: package.json reached 2.4.0 while
		// the handshake kept telling clients 2.0.0. Nothing branches on the value,
		// so nothing broke and nothing complained — it just made every
		// version-specific question misleading.
		//
		// If this fails, you bumped one and not the other. Bump both.
		expect(SERVER_VERSION).toBe(pkg.version);
	});

	it("is a plain semver triple", () => {
		// Guards against a "2.4.0-dev" or "v2.4.0" sneaking in — MCP clients display
		// this string verbatim.
		expect(SERVER_VERSION).toMatch(/^\d+\.\d+\.\d+$/);
	});
});
