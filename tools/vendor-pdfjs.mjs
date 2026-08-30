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
//
// ## Why the Apache-2.0 notices are emitted here rather than added to the files
//
// We redistribute pdf.js verbatim, so Apache-2.0 §4 applies to us. §4(a) wants
// recipients to get a copy of the licence, and §4(c) wants the copyright and
// attribution notices retained from the Source form. The two minified bundles
// carry their own `@licstart` block and need nothing; the extracted stylesheet
// is a cut out of `web/pdf_viewer.css`, and the first version of this script
// dropped that file's notice on the floor.
//
// Patching the stylesheet by hand would not have fixed it: this script rewrites
// the whole vendor directory on the next version bump and would have dropped it
// again, silently. So the notice is *read out of the upstream file* and the
// licence is *fetched from the same release*, which also means neither can go
// stale relative to what was actually vendored.

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

/** The vendored copy of Apache-2.0, discharging §4(a) for anyone given this folder. */
const LICENSE_NAME = "LICENSE";

/**
 * Phrases that identify the fetched licence as Apache-2.0.
 *
 * Taken from the file actually served for 6.2.108, not from the canonical
 * Apache text: pdf.js ships the licence without its "APPENDIX: How to apply"
 * section, so asserting on the appendix would fail against a correct file. The
 * point of the check is narrow — if a future release relicenses pdf.js, this
 * script must stop rather than vendor new terms under a header claiming the old
 * ones.
 */
const APACHE_LICENSE_MARKERS = ["Apache License", "Version 2.0, January 2004", "END OF TERMS AND CONDITIONS"];

/** The one rule we keep out of the viewer stylesheet. */
const TEXT_LAYER_RULE = ".textLayer";

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

/**
 * The leading comment block of a stylesheet, when it carries a copyright notice.
 *
 * Apache-2.0 §4(c) asks that we retain "all copyright ... and attribution
 * notices from the Source form of the Work". Reading the notice out of the file
 * the rule is cut from is what makes that literally true: a notice typed out
 * here would be our text *about* Mozilla's file, and it would go stale the
 * first time upstream changed the year.
 *
 * An absent notice throws rather than quietly emitting nothing. Emitting
 * nothing is precisely the regression this exists to stop, and a second silent
 * drop would be indistinguishable from a clean run.
 */
function extractLeadingNotice(css) {
    const start = css.indexOf("/*");
    if (start < 0 || css.slice(0, start).trim() !== "") {
        throw new Error("the upstream stylesheet does not open with a comment, so it carries no notice to retain");
    }
    const end = css.indexOf("*/", start + 2);
    if (end < 0) throw new Error("the upstream stylesheet's leading comment is unterminated");
    const notice = css.slice(start, end + 2);
    if (!/copyright/i.test(notice)) {
        throw new Error("the upstream stylesheet's leading comment carries no copyright notice");
    }
    return notice;
}

/**
 * The stylesheet we commit: upstream's notice, our provenance line, the rule.
 *
 * Split out of `main` so a test can drive the exact bytes this writes without
 * going near the network. The generator is the load-bearing half of the
 * attribution fix — an artifact patched by hand is undone by the next version
 * bump — so what it emits has to be testable on its own.
 *
 * The provenance paragraph is also what Apache-2.0 §4(b) asks for: a prominent
 * notice that we changed the file. It says which file, which release and what
 * was taken, so it states the modification rather than merely admitting one.
 */
function buildTextLayerCss(viewerCss, version) {
    const notice = extractLeadingNotice(viewerCss);
    const rule = extractRule(viewerCss, TEXT_LAYER_RULE);
    return (
        `${notice}\n\n` +
        `/* Extracted from pdfjs-dist@${version} web/pdf_viewer.css by tools/vendor-pdfjs.mjs.\n` +
        `   Only the ${TEXT_LAYER_RULE} rule: the other 160 KB style layers this viewer\n` +
        `   never mounts. Do not edit — re-run the vendoring script.\n` +
        `   The Apache-2.0 licence text is in ${LICENSE_NAME}, beside this file. */\n` +
        `${rule}\n`
    );
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
    const pkg = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${version}`;
    const base = `${pkg}/build`;
    const web = `${pkg}/web`;

    process.stdout.write(`vendoring pdfjs-dist@${version}\n`);

    const [main_, worker, viewerCss, license] = await Promise.all([
        fetchBuffer(`${base}/pdf.min.mjs`),
        fetchBuffer(`${base}/pdf.worker.min.mjs`),
        fetchBuffer(`${web}/pdf_viewer.css`),
        fetchBuffer(`${pkg}/LICENSE`),
    ]);

    assertUtf8(main_, "pdf.min.mjs");
    assertUtf8(worker, "pdf.worker.min.mjs");
    assertUtf8(license, LICENSE_NAME);

    // The licence is taken from the same release as the code, so the copy we
    // ship cannot describe a different version than the one we vendored. The
    // marker check is the stop: if pdf.js ever relicenses, this refuses rather
    // than committing new terms under a stylesheet header naming the old ones.
    const licenseText = license.toString("utf8");
    for (const marker of APACHE_LICENSE_MARKERS) {
        if (!licenseText.includes(marker)) {
            throw new Error(
                `${pkg}/LICENSE does not read as Apache-2.0 (no '${marker}') — ` +
                    `check what pdfjs-dist@${version} is licensed under before vendoring it`,
            );
        }
    }

    // The text layer is positioned entirely by these properties, so the CSS is
    // taken from the same release rather than written out here from what the
    // minified source appears to do. Asserting the contract in both directions
    // is what makes that safe across a version bump.
    const textLayerCss = buildTextLayerCss(viewerCss.toString("utf8"), version);
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
    await writeFile(path.join(VENDOR_DIR, LICENSE_NAME), license);

    const textLayerBuffer = Buffer.from(textLayerCss, "utf8");
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
        // Deliberately its own key rather than an entry in `files`. `files` means
        // "vendored asset the viewer fetches over /vendor/", and the integration
        // suite asserts exactly that by walking `files` and expecting HTTP 200
        // for each. The licence is shipped to be *read in the folder*, not
        // served, so listing it there would assert a route that does not exist.
        // It is still covered by a digest, in the unit suite, like everything else.
        license: {
            name: LICENSE_NAME,
            spdx: "Apache-2.0",
            holder: "Mozilla Foundation",
            source: `${pkg}/LICENSE`,
            bytes: license.byteLength,
            sha256: sha256(license),
        },
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
            `  pdf-text-layer.css    ${textLayerBuffer.byteLength} bytes (${TEXT_LAYER_RULE}, notice retained)\n` +
            `  ${LICENSE_NAME}               ${license.byteLength} bytes (Apache-2.0, from ${pkg}/LICENSE)\n` +
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

export {
    PART_TARGET_BYTES,
    VENDOR_DIR,
    MANIFEST_NAME,
    LICENSE_NAME,
    APACHE_LICENSE_MARKERS,
    TEXT_LAYER_CONTRACT,
    TEXT_LAYER_RULE,
    extractRule,
    extractLeadingNotice,
    buildTextLayerCss,
};
