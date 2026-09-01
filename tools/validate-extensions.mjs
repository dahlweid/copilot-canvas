#!/usr/bin/env node
// Checks the invariants that make an extension folder installable.
//
// These are not style rules. Each is a constraint the extension runtime or
// `install_extension` imposes, stated inline at the check that enforces it
// (C1-C4 below), and each has a failure mode that otherwise only
// shows up at install time or as a silently dead extension process.
//
// Run: node tools/validate-extensions.mjs

import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const EXTENSIONS_DIR = path.join(ROOT, ".github", "extensions");

// These are the limits `install_extension` actually enforces, measured against
// the running app (v1.0.80) rather than taken from the documentation: a
// 1,200,000-byte file is refused with "too large (1200000 bytes > 1000000 byte
// limit)" and 7.68 MB of under-cap files with "exceed the 5000000 byte total
// limit". They are decimal, not binary, and they apply to *both* install
// sources — gist and repo folder alike. Using 1 MiB / 5 MiB here, as this file
// previously did, left a 48,576-byte window in which an extension passed
// validation and then failed at install, which is the one thing this check
// exists to prevent.
const MAX_FILE_BYTES = 1_000_000; // C4: install per-file limit
const MAX_TOTAL_BYTES = 5_000_000; // C4: install total limit

// Bytes are printed only once they are about to decide something. Below this
// point, no single file the per-file limit still permits can carry the total
// past the envelope, so the number answers no question anyone is asking; above
// it, the very next legal file can. Deriving the band from the two limits
// rather than picking a percentage leaves no third number to quote, go stale,
// or be mistaken for a property of the tree -- which is the misuse the routine
// figure was removed for (#82).
const WARN_TOTAL_BYTES = MAX_TOTAL_BYTES - MAX_FILE_BYTES;

// The `version` key in copilot-extension.json is the manifest format version and
// is parsed as a u32 — see checkManifest for the measurement.
const MANIFEST_FORMAT_VERSION = 1;

const problems = [];
const fail = (extension, message) => problems.push(`${extension}: ${message}`);

async function listFiles(dir) {
    const out = [];
    for (const entry of await readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...(await listFiles(full)));
        else if (entry.isFile()) out.push(full);
    }
    return out;
}

/** C1/C4: the entry point and manifest must exist, and agree on the name. */
async function checkManifest(name, dir) {
    try {
        await stat(path.join(dir, "extension.mjs"));
    } catch {
        fail(name, "missing extension.mjs (C1: the entry file name is fixed)");
    }

    let manifest;
    try {
        manifest = JSON.parse(await readFile(path.join(dir, "copilot-extension.json"), "utf8"));
    } catch (err) {
        return fail(name, `copilot-extension.json missing or invalid: ${err.message}`);
    }
    if (manifest.name !== name) {
        fail(name, `manifest name '${manifest.name}' does not match its folder`);
    }

    // `version` is the manifest *format* version, not the product version.
    // Measured: install_extension parses it as a Rust u32, so a semver string is
    // rejected outright — `invalid type: string "1.0.0", expected u32`. The app's
    // own authoring guide states the shape as `{ "name": ..., "version": 1 }`, and
    // the one extension shipped on this machine (mobile-canvas) uses `1`.
    // The in-place loader is lenient about this; the installer is not, so an
    // extension with a semver `version` runs in the dev checkout and cannot be
    // installed by anyone. That is exactly the divergence this file exists to catch.
    if (manifest.version !== MANIFEST_FORMAT_VERSION) {
        fail(
            name,
            `manifest version must be the number ${MANIFEST_FORMAT_VERSION} ` +
                `(install_extension parses it as u32), not ${JSON.stringify(manifest.version)}`,
        );
    }

    // The product version therefore needs its own field. Unknown keys are tolerated
    // by the installer (measured: a manifest carrying `productVersion` installs), so
    // this is a safe place to keep the single in-repo source of truth for release tags.
    if (!/^\d+\.\d+\.\d+$/.test(manifest.productVersion ?? "")) {
        fail(name, `manifest productVersion '${manifest.productVersion}' is not semver`);
    }
}

/** C2: the SDK is auto-resolved; a package.json implies a build we cannot run. */
function checkNoPackageJson(name, files, dir) {
    for (const file of files) {
        if (path.basename(file) === "package.json") {
            fail(name, `${path.relative(dir, file)} exists (C2: no package.json in an extension)`);
        }
    }
}

/**
 * stdout carries JSON-RPC, so a stray console.log kills the extension. src/ui/
 * runs in the canvas iframe, where console is legitimate — a repo-wide rule
 * would be wrong. Tests run as their own process and are exempt too.
 */
async function checkNoConsoleLog(name, files, dir) {
    const uiDir = path.join(dir, "src", "ui") + path.sep;
    const testDir = path.join(dir, "test") + path.sep;
    for (const file of files) {
        if (!file.endsWith(".mjs") && !file.endsWith(".js")) continue;
        if (file.startsWith(uiDir) || file.startsWith(testDir)) continue;

        const lines = (await readFile(file, "utf8")).split(/\r?\n/);
        lines.forEach((line, i) => {
            if (line.trim().startsWith("//")) return;
            if (/(^|[^.\w])console\.(log|info|warn|debug|error)\s*\(/.test(line)) {
                fail(name, `${path.relative(dir, file)}:${i + 1} writes to console (corrupts JSON-RPC)`);
            }
        });
    }
}

/** C3: install copies one folder, so an import escaping it breaks once installed. */
async function checkSelfContained(name, files, dir) {
    // Four shapes reach a module: `import x from`, `export x from`, a bare
    // side-effect `import "..."`, and dynamic `import("...")`. Missing any one
    // of them makes this check quietly useless.
    const pattern = new RegExp(
        [
            /(?:import|export)\s[^'"()\n]*?from\s*['"](\.[^'"]*)['"]/.source,
            /import\s*['"](\.[^'"]*)['"]/.source,
            /import\s*\(\s*['"](\.[^'"]*)['"]\s*\)/.source,
        ].join("|"),
        "g",
    );
    for (const file of files) {
        if (!file.endsWith(".mjs") && !file.endsWith(".js")) continue;
        const text = await readFile(file, "utf8");
        for (const match of text.matchAll(pattern)) {
            const specifier = match[1] ?? match[2] ?? match[3];
            const resolved = path.resolve(path.dirname(file), specifier);
            if (!resolved.startsWith(dir + path.sep)) {
                fail(name, `${path.relative(dir, file)} imports '${specifier}' from outside the extension (C3)`);
            }
        }
    }
}

/** C4: install is UTF-8 text only, 1,000,000 bytes per file and 5,000,000 in total. */
async function checkPackagingEnvelope(name, files, dir) {
    let total = 0;
    const decoder = new TextDecoder("utf-8", { fatal: true });
    for (const file of files) {
        const buffer = await readFile(file);
        total += buffer.byteLength;
        if (buffer.byteLength > MAX_FILE_BYTES) {
            fail(
                name,
                `${path.relative(dir, file)} is ${buffer.byteLength} bytes (C4 limit ${MAX_FILE_BYTES})`,
            );
        }
        try {
            decoder.decode(buffer);
        } catch {
            fail(name, `${path.relative(dir, file)} is not valid UTF-8 (C4: gists refuse binary files)`);
        }
    }
    if (total > MAX_TOTAL_BYTES) {
        fail(name, `total ${total} bytes exceeds the C4 limit of ${MAX_TOTAL_BYTES}`);
    } else if (total > WARN_TOTAL_BYTES) {
        // Working-tree bytes, and a CRLF checkout holds more of them than an LF
        // one at the same commit -- so this figure describes this checkout, not
        // the tree, which is why it is not printed routinely. It is still the
        // right number to print here: install copies the working tree, and the
        // only question being answered is how much room is left in it.
        console.log(
            `${name}: ${total} of ${MAX_TOTAL_BYTES} bytes on disk, ${MAX_TOTAL_BYTES - total} left — ` +
                `less than the ${MAX_FILE_BYTES}-byte per-file limit, so one more file could exceed the total`,
        );
    }
}

/** Every .mjs must at least parse. */
async function checkSyntax(name, files, dir) {
    for (const file of files) {
        if (!file.endsWith(".mjs")) continue;
        try {
            await execFileAsync(process.execPath, ["--check", file]);
        } catch (err) {
            fail(name, `${path.relative(dir, file)} does not parse: ${err.stderr?.trim() ?? err.message}`);
        }
    }
}

let extensions;
try {
    extensions = (await readdir(EXTENSIONS_DIR, { withFileTypes: true }))
        .filter((entry) => entry.isDirectory())
        .map((entry) => entry.name);
} catch {
    console.error(`No extensions directory at ${EXTENSIONS_DIR}`);
    process.exit(1);
}

// A validator that passes because it found nothing is worse than no validator.
if (extensions.length === 0) {
    console.error("No extensions found — refusing to pass vacuously.");
    process.exit(1);
}

for (const name of extensions) {
    const dir = path.join(EXTENSIONS_DIR, name);
    const files = await listFiles(dir);
    await checkManifest(name, dir);
    checkNoPackageJson(name, files, dir);
    await checkNoConsoleLog(name, files, dir);
    await checkSelfContained(name, files, dir);
    await checkSyntax(name, files, dir);
    await checkPackagingEnvelope(name, files, dir);
    // The file count, and no byte total. The total was accumulated from disk, so
    // a CRLF checkout and an LF checkout disagreed about the same commit
    // (measured: ~23 KB apart across two clones) while the count was identical
    // in both -- and the figure was being quoted between sessions as if it
    // identified a tree. Bytes are reported by checkPackagingEnvelope instead,
    // only where they are being compared against the envelope. See #82.
    console.log(`${name}: ${files.length} files`);
}

if (problems.length > 0) {
    console.error(`\n${problems.length} problem(s):`);
    for (const problem of problems) console.error(`  - ${problem}`);
    process.exit(1);
}

console.log(`\nOK — ${extensions.length} extension(s) valid.`);
