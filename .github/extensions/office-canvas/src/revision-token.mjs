// The revision token: a hash of the file, returned with a structure map and
// required by any edit.
//
// Measured behaviour, which is exactly what optimistic concurrency needs. The
// isolation spike that recorded it has been retired, so the measurement is
// reproduced here rather than cited: a SHA-256 prefix over the file, computed
// in 3 ms.
//
//   | our own edit and save          | token changes     |
//   | save with nothing dirty        | token unchanged   |
//   | external regeneration          | token changes     |
//
// The "nothing dirty" row is the important one: Word skips writing a document
// it considers clean, so merely inspecting a document does not churn the token
// and force a re-read. The first row is what lets an edit response hand back a
// fresh token, so the agent is forced to re-read after somebody else's edit and
// never after its own.
//
// That is the whole concurrency story for transient locking: we hold nothing
// between operations, so we cannot assume anything stayed put, and the token is
// what converts that from a hazard into a detectable, refusable condition.
//
// A SHA-256 prefix rather than mtime+size, because a script that regenerates a
// document can easily reproduce both while changing every word.

import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";

/** Named so the scheme can change later without an edit silently comparing apples to oranges. */
export const TOKEN_PREFIX = "sha256:";
const TOKEN_HEX_LENGTH = 16;

export function revisionTokenOf(bytes) {
    return TOKEN_PREFIX + createHash("sha256").update(bytes).digest("hex").slice(0, TOKEN_HEX_LENGTH);
}

/**
 * Token for a file on disk. Streamed rather than read whole: a token is taken
 * on every read and every edit, and a large document should not cost its own
 * size in resident memory each time.
 */
export function fileRevisionToken(filePath) {
    return new Promise((resolve, reject) => {
        const hash = createHash("sha256");
        const stream = createReadStream(filePath);
        stream.on("error", reject);
        stream.on("data", (chunk) => hash.update(chunk));
        stream.on("end", () => resolve(TOKEN_PREFIX + hash.digest("hex").slice(0, TOKEN_HEX_LENGTH)));
    });
}

/** Exact comparison, but null-safe, so a missing token never passes as a match. */
export function tokensMatch(a, b) {
    return typeof a === "string" && typeof b === "string" && a.length > 0 && a === b;
}
