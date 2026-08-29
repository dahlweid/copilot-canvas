#!/usr/bin/env node
// Vendors pdf.js into the extension, splitting the worker so every committed
// file fits the install envelope.
//
// Run: node tools/vendor-pdfjs.mjs [--version 6.2.108]
//
// ## Why this exists at all
//
// `pdf.worker.min.mjs` is 1,262,398 bytes against a 1,000,000-byte per-file
// cap that `install_extension` enforces on *both* install paths — gist and
// repo folder alike. It is UTF-8 JavaScript, so it is not refused for being
// binary; it is simply too large for any single file.
//
// ## Why the split happens here and not in the packager
//
// C3 requires an extension folder to run exactly as committed, and repo-folder
// install never runs `tools/package-extension.mjs`. Splitting at packaging time
// would produce an installable artefact from a source tree that is not itself
// installable — the folder on `main` would be uninstallable while the release
// zip worked. So the parts are committed, and this script is what mints them.
//
// ## Why the output lands under src/ui/
//
// `tools/validate-extensions.mjs` rejects `console.*` outside `src/ui/`,
// because stdout carries JSON-RPC in the extension process. Minified pdf.js
// contains `console.*` calls. `src/ui/` is exempt because it runs in the canvas
// iframe — which is exactly where pdf.js runs, so this placement is the honest
// one rather than a way around the check.
//
// ## Why the parts are not named *.mjs
//
// The validator runs `node --check` on every `.mjs`. A part on its own is a
// fragment of a program and does not parse. The `.partN` suffix keeps them out
// of that check, and out of the `console.*` and import-graph scans, none of
// which can say anything meaningful about half an expression.

import { mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const VENDOR_DIR = path.join(ROOT, ".github", "extensions", "office-canvas", "src", "ui", "vendor");
const MANIFEST_NAME = "pdfjs.manifest.json";

const DEFAULT_VERSION = "6.2.108";

/**
 * Target size for one part.
 *
 * Deliberately well under the 1,000,000-byte cap rather than just below it.
 * The cap is not restated here: the last step of this script runs
 * `tools/validate-extensions.mjs`, so the real gate — the one CI runs — is what
 * confirms the output fits. A second copy of the number would be a second thing
 * to keep in step, and this repo has had constants drift three times.
 */
const PART_TARGET_BYTES = 600_000;

/** How far back we will hunt for a newline before falling back. */
const NEWLINE_SEARCH_WINDOW = 65_536;

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");

/**
 * Splits a buffer into parts of at most `target` bytes, each independently
 * decodable as UTF-8.
 *
 * The rule is **not** "the boundary byte must be < 0x80", which was the first
 * guess and is stricter than the encoding requires. UTF-8 continuation bytes
 * are `0x80`-`0xBF`; a boundary is safe as long as it does not land on one,
 * because every other byte starts a new code point. Measured on the real
 * worker: the newline branch never fires (minified JavaScript is one enormous
 * line), so the continuation walk-back does all the work — which is why both
 * branches have to exist and both have to be tested.
 */
export function splitUtf8Safe(buffer, target) {
    if (!(target > 0)) throw new Error(`part target must be positive, got ${target}`);
    const parts = [];
    let start = 0;
    while (start < buffer.length) {
        let end = Math.min(start + target, buffer.length);
        if (end < buffer.length) {
            const newline = buffer.lastIndexOf(0x0a, end - 1);
            if (newline > start && end - newline < NEWLINE_SEARCH_WINDOW) {
                end = newline + 1; // just after a newline, so the parts stay diffable
            } else {
                while (end > start && (buffer[end] & 0xc0) === 0x80) end--;
                if (end === start) throw new Error("no safe UTF-8 boundary within one part");
            }
        }
        parts.push(buffer.subarray(start, end));
        start = end;
    }
    return parts;
}

/**
 * Custom properties the *bundle* sets on text-layer nodes at render time.
 *
 * These are the whole contract between `TextLayer` and the stylesheet: it sets
 * them inline per span and the CSS is what turns them into position and size.
 * Extracting the stylesheet without checking them would ship a text layer that
 * silently renders every span at font-size 0 -- invisible, unselectable, and
 * still "working" as far as any test of the PDF itself is concerned.
 *
 * Checked in both directions below: each must appear in the bundle (so a rename
 * upstream is caught) and in the extracted CSS (so a bad extraction is caught).
 */
const TEXT_LAYER_CONTRACT = ["--font-height", "--scale-x", "--rotate", "--min-font-size", "--total-scale-factor"];

/**
 * Cuts one brace-balanced rule out of a stylesheet.
 *
 * pdf.js ships nested CSS, so `.textLayer{...}` contains every rule that applies
 * to the layer. Taking the whole balanced block gets all of them and nothing
 * else -- no annotation-editor, XFA or signature styling for layers we never
 * mount, which is 160 KB of the 163 KB file.
 */
function extractRule(css, selector) {
    const start = css.indexOf(`${selector}{`);
    if (start < 0) throw new Error(`${selector} not found in the stylesheet`);
    let depth = 0;
    for (let i = css.indexOf("{", start); i < css.length; i++) {
        if (css[i] === "{") depth++;
        else if (css[i] === "}" && --depth === 0) return css.slice(start, i + 1);
    }
    throw new Error(`${selector} is not brace-balanced`);
}

function assertUtf8(buffer, what) {
    try {
        new TextDecoder("utf-8", { fatal: true }).decode(buffer);
    } catch {
        throw new Error(`${what} is not independently valid UTF-8`);
    }
}

async function fetchBuffer(url) {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`${url} -> ${res.status} ${res.statusText}`);
    return Buffer.from(await res.arrayBuffer());
}

function parseArgs(argv) {
    let version = DEFAULT_VERSION;
    for (let i = 0; i < argv.length; i++) {
        if (argv[i] === "--version") version = argv[++i] ?? version;
    }
    if (!/^\d+\.\d+\.\d+$/.test(version)) throw new Error(`--version must be semver, got '${version}'`);
    return { version };
}

async function main() {
    const { version } = parseArgs(process.argv.slice(2));
    const base = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${version}/build`;
    const web = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${version}/web`;

    process.stdout.write(`vendoring pdfjs-dist@${version}\n`);

    const [main_, worker, viewerCss] = await Promise.all([
        fetchBuffer(`${base}/pdf.min.mjs`),
        fetchBuffer(`${base}/pdf.worker.min.mjs`),
        fetchBuffer(`${web}/pdf_viewer.css`),
    ]);

    assertUtf8(main_, "pdf.min.mjs");
    assertUtf8(worker, "pdf.worker.min.mjs");

    // The text layer is positioned entirely by these properties, so the CSS is
    // taken from the same release rather than written out here from what the
    // minified source appears to do. Asserting the contract in both directions
    // is what makes that safe across a version bump.
    const textLayerCss = extractRule(viewerCss.toString("utf8"), ".textLayer");
    const bundleText = main_.toString("utf8");
    for (const property of TEXT_LAYER_CONTRACT) {
        if (!bundleText.includes(property)) {
            throw new Error(`pdf.min.mjs no longer sets ${property} — the text-layer contract has changed`);
        }
        if (!textLayerCss.includes(property)) {
            throw new Error(`the extracted .textLayer CSS does not use ${property}`);
        }
    }

    const parts = splitUtf8Safe(worker, PART_TARGET_BYTES);
    parts.forEach((part, i) => assertUtf8(part, `part ${i}`));

    const rejoined = Buffer.concat(parts);
    if (sha256(rejoined) !== sha256(worker)) {
        throw new Error("reassembly is not byte-exact — refusing to write parts that do not round-trip");
    }

    // Clear the directory first so a version bump that produces *fewer* parts
    // cannot leave an orphan behind. An orphaned part is invisible: the manifest
    // would not list it, the route would not serve it, and it would sit in the
    // installed folder counting against the total budget forever.
    await rm(VENDOR_DIR, { recursive: true, force: true });
    await mkdir(VENDOR_DIR, { recursive: true });

    await writeFile(path.join(VENDOR_DIR, "pdf.min.mjs"), main_);

    const textLayerBuffer = Buffer.from(
        "/* Extracted from pdfjs-dist@" +
            version +
            " web/pdf_viewer.css by tools/vendor-pdfjs.mjs.\n" +
            "   Only the .textLayer rule: the other 160 KB style layers this viewer\n" +
            "   never mounts. Do not edit — re-run the vendoring script. */\n" +
            textLayerCss +
            "\n",
        "utf8",
    );
    await writeFile(path.join(VENDOR_DIR, "pdf-text-layer.css"), textLayerBuffer);

    const partEntries = [];
    for (const [i, part] of parts.entries()) {
        const name = `pdf.worker.min.mjs.part${i}`;
        await writeFile(path.join(VENDOR_DIR, name), part);
        partEntries.push({ name, bytes: part.byteLength });
    }

    // `parts` is an ordered array, and that order *is* the concatenation order.
    // Nothing anywhere derives the order from the filenames, which is deliberate:
    // a lexicographic sort puts `part10` before `part2`, and three parts today
    // would hide that until a version bump produced ten.
    const manifest = {
        version,
        source: base,
        generatedBy: "tools/vendor-pdfjs.mjs",
        files: [
            { name: "pdf.min.mjs", bytes: main_.byteLength, sha256: sha256(main_) },
            {
                name: "pdf-text-layer.css",
                bytes: textLayerBuffer.byteLength,
                sha256: sha256(textLayerBuffer),
                extractedFrom: `${web}/pdf_viewer.css`,
                rule: ".textLayer",
            },
        ],
        worker: {
            name: "pdf.worker.min.mjs",
            bytes: worker.byteLength,
            sha256: sha256(worker),
            parts: partEntries,
        },
    };
    await writeFile(path.join(VENDOR_DIR, MANIFEST_NAME), `${JSON.stringify(manifest, null, 4)}\n`);

    process.stdout.write(
        `  pdf.min.mjs           ${main_.byteLength} bytes\n` +
            `  pdf.worker.min.mjs    ${worker.byteLength} bytes -> ${parts.length} parts ` +
            `(${partEntries.map((p) => p.bytes).join(" + ")})\n`,
    );

    // The envelope is not re-checked against a local copy of the cap. The
    // validator is the thing CI runs, so running it here means this script
    // cannot disagree with the gate about what fits.
    process.stdout.write("\nverifying against the real install envelope:\n");
    const { stdout } = await execFileAsync(process.execPath, [path.join(ROOT, "tools", "validate-extensions.mjs")]);
    process.stdout.write(stdout);
}

// Importable for tests without running the fetch.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    try {
        await main();
    } catch (err) {
        process.stderr.write(`vendor-pdfjs: ${err.message}\n`);
        process.exit(1);
    }
}

export { PART_TARGET_BYTES, VENDOR_DIR, MANIFEST_NAME, TEXT_LAYER_CONTRACT, extractRule };
