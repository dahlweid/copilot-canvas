// The Word application icon, taken from the Word installed on this machine.
//
// Not a file in this repository, deliberately. The repo is public, so committing
// Microsoft's mark would be redistributing it, and the open icon sets have
// removed the whole Microsoft family for trademark reasons -- measured by
// `spikes/word-icon/probes/probe-icon-sources.mjs`: simple-icons answers 404 for
// `microsoft`, `microsoftword`, `microsoft-word` and `microsoftoffice`, while
// `libreoffice` answers 200, so the instrument works and the family is genuinely
// gone. Borrowing another suite's mark for a Word document would name the wrong
// product. What is left is the user's own installation, which redistributes
// nothing.
//
// Extracted once and held. The same probe measures the extraction at 32x32,
// 2349 bytes, ~1.5s, with the machine's `WINWORD.EXE` count unchanged across it
// -- it reads the executable's resources rather than automating Word. Per
// request that would still be absurd, and the answer cannot change while this
// process lives.

import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

const SCRIPT_PATH = path.join(path.dirname(fileURLToPath(import.meta.url)), "word-icon.ps1");

/**
 * How long to wait for the extraction.
 *
 * Short by this repo's standards, and for a reason the generous deadlines
 * elsewhere do not share: nothing waits on this. It decorates a bar that is
 * already drawn, so a slow answer and no answer are the same to the user, and
 * the cost of guessing wrong is a missing 32px image rather than a hung open.
 */
const TIMEOUT_MS = 10_000;

/** Run the extractor, resolving to base64 or null when there is no icon to be had. */
function runExtractor({ signal } = {}) {
    return new Promise((resolve) => {
        // Discrete argv elements, never an interpolated `-Command` string, and
        // the script takes no parameters at all -- so there is no path for a
        // shell to reparse. See word-icon.ps1's header.
        const child = spawn(
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", SCRIPT_PATH],
            { stdio: ["ignore", "pipe", "pipe"], windowsHide: true, signal },
        );

        let out = "";
        child.stdout.setEncoding("utf8");
        child.stdout.on("data", (chunk) => {
            out += chunk;
        });
        child.stderr.resume();

        // Every failure resolves rather than rejecting: a missing Word, a
        // PowerShell that is not there at all, a non-Windows host. None of them
        // is an error the user should be told about, because none of them stops
        // the viewer doing its job.
        child.on("error", () => resolve(null));
        child.on("close", (code) => resolve(code === 0 && out.trim() ? out.trim() : null));
    });
}

/**
 * Decodes what the extractor produced, refusing anything that is not a PNG.
 *
 * The signature check is not defensive padding: `-NoProfile` suppresses a
 * profile's output but a machine policy banner or a transcription notice can
 * still reach stdout ahead of the payload, and base64 decoding is happy to turn
 * that into bytes. Serving those bytes as `image/png` would produce a broken
 * image rather than no image, which is the worse of the two.
 */
export function decodeIcon(base64) {
    if (typeof base64 !== "string" || !base64.trim()) return null;
    let buffer;
    try {
        buffer = Buffer.from(base64.trim(), "base64");
    } catch {
        return null;
    }
    const PNG_MAGIC = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    if (buffer.length <= PNG_MAGIC.length || !buffer.subarray(0, PNG_MAGIC.length).equals(PNG_MAGIC)) return null;
    return { buffer, etag: `"${createHash("sha256").update(buffer).digest("hex").slice(0, 32)}"` };
}

/**
 * A source of the Word icon that extracts at most once.
 *
 * `run` is injectable so the extraction has a reachable test: CI has no Word
 * and no PowerShell, and a test that skipped there would leave the caching and
 * the failure handling -- the parts with branches -- unmeasured on every machine
 * that matters.
 */
export function createWordIconSource({ run = runExtractor, timeoutMs = TIMEOUT_MS } = {}) {
    let pending = null;

    return {
        /** The icon, or null when this machine cannot produce one. Extracted once. */
        async get() {
            // The *promise* is memoized, not the value: two requests arriving
            // together must share one extraction rather than spawning two.
            // Memoized on failure too -- a machine without Word does not grow
            // one, and re-spawning PowerShell per request to rediscover that is
            // exactly the per-request work this must not do.
            pending ??= (async () => {
                const controller = new AbortController();
                const timer = setTimeout(() => controller.abort(), timeoutMs);
                try {
                    return decodeIcon(await run({ signal: controller.signal }));
                } catch {
                    return null;
                } finally {
                    clearTimeout(timer);
                }
            })();
            return await pending;
        },
    };
}

export const wordIcon = createWordIconSource();
