// Serves the vendored pdf.js assets, reassembling the worker from its parts.
//
// `pdf.worker.min.mjs` is 1,262,398 bytes and the install envelope caps a file
// at 1,000,000, so it ships as three committed parts (see tools/vendor-pdfjs.mjs
// for why the split is at vendoring time rather than packaging time). The cap
// constrains *files in the installed folder*; it says nothing about what a
// loopback server may return. So the workaround stays in this one module and
// pdf.js sees an ordinary `workerSrc` URL, which is also what lets a version
// bump leave the viewer untouched.

import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";

const VENDOR_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "ui", "vendor");
const MANIFEST_NAME = "pdfjs.manifest.json";

export class VendorAssetError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "VendorAssetError";
        this.code = code;
    }
}

/**
 * Reads and validates the vendoring manifest.
 *
 * The manifest is the single source of truth for what was vendored, including
 * the part order. Nothing derives that order from the filenames: a lexicographic
 * sort places `part10` before `part2`, and the three parts pdf.js needs today
 * would hide that until a version bump produced ten.
 */
export async function readManifest(dir = VENDOR_DIR) {
    const file = path.join(dir, MANIFEST_NAME);
    let raw;
    try {
        raw = await readFile(file, "utf8");
    } catch (err) {
        throw new VendorAssetError(
            "vendor_missing",
            `The pdf.js manifest is missing at ${file} (${err.code ?? err.message}). ` +
                `Run 'node tools/vendor-pdfjs.mjs' to vendor pdf.js.`,
        );
    }
    let manifest;
    try {
        manifest = JSON.parse(raw);
    } catch (err) {
        throw new VendorAssetError("vendor_invalid", `The pdf.js manifest is not valid JSON: ${err.message}`);
    }
    const parts = manifest?.worker?.parts;
    if (!Array.isArray(parts) || parts.length === 0) {
        throw new VendorAssetError("vendor_invalid", "The pdf.js manifest lists no worker parts.");
    }
    return manifest;
}

/**
 * Concatenates the worker parts in manifest order and returns the whole thing.
 *
 * Returns a complete buffer rather than a stream deliberately. Streaming the
 * parts as they are read would mean a part that fails on the third read arrives
 * after a `200` and two thirds of a file are already on the wire — the client
 * gets a truncated, syntactically broken worker and no error. Reading fully
 * first turns that into an ordinary failed request.
 */
export async function joinWorker(dir = VENDOR_DIR) {
    const manifest = await readManifest(dir);
    const buffers = [];
    for (const part of manifest.worker.parts) {
        // Names come from the manifest but still get basenamed: the manifest is
        // a committed file, not user input, yet a path in it would otherwise be
        // able to read anything the process can.
        const file = path.join(dir, path.basename(String(part.name)));
        try {
            buffers.push(await readFile(file));
        } catch (err) {
            throw new VendorAssetError(
                "vendor_incomplete",
                `The pdf.js worker part '${part.name}' could not be read (${err.code ?? err.message}). ` +
                    `The vendored copy is incomplete; run 'node tools/vendor-pdfjs.mjs'.`,
            );
        }
    }
    const joined = Buffer.concat(buffers);
    if (joined.byteLength !== manifest.worker.bytes) {
        throw new VendorAssetError(
            "vendor_incomplete",
            `The reassembled pdf.js worker is ${joined.byteLength} bytes, but the manifest records ` +
                `${manifest.worker.bytes}. The vendored parts do not match the manifest.`,
        );
    }
    // The byte count alone cannot see a part served out of order -- the parts are
    // two equal 600,000-byte blocks and a remainder, so any permutation totals
    // the same and yields a worker that is the right size and syntactically
    // ruined. The digest is also what the ETag claims about the body, and an
    // ETag asserting a hash nobody computed is a cause the code never
    // determined.
    const digest = createHash("sha256").update(joined).digest("hex");
    if (digest !== manifest.worker.sha256) {
        throw new VendorAssetError(
            "vendor_corrupt",
            `The reassembled pdf.js worker hashes to ${digest}, but the manifest records ` +
                `${manifest.worker.sha256}. The vendored parts are corrupt or out of order.`,
        );
    }
    return { buffer: joined, etag: `"${digest}"` };
}

// One copy for the whole process. The bytes are identical for every canvas, and
// 1.26 MB per open panel would be waste. The in-flight promise is kept, not just
// the result, so two canvases opening at once read the parts once between them.
//
// Keyed by directory: a single slot would answer a second directory with the
// first one's bytes, which is exactly the kind of "right size, wrong content"
// failure the digest check exists to catch and would not see.
const workerPromises = new Map();

export function loadWorker(dir = VENDOR_DIR) {
    let pending = workerPromises.get(dir);
    if (!pending) {
        pending = joinWorker(dir).catch((err) => {
            // A failed load must not be cached: the fix is to re-run the
            // vendoring script, and the next request should see the result.
            workerPromises.delete(dir);
            throw err;
        });
        workerPromises.set(dir, pending);
    }
    return pending;
}

/** Test seam: drops the cached worker so a test can vary what is on disk. */
export function resetWorkerCache() {
    workerPromises.clear();
}

export { VENDOR_DIR, MANIFEST_NAME };
