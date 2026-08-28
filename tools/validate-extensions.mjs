#!/usr/bin/env node
// Checks the invariants that make an extension folder installable.
//
// These are not style rules. Each corresponds to a documented constraint in
// docs/repo-restructure.md §1, and each has a failure mode that otherwise only
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

const MAX_FILE_BYTES = 1024 * 1024; // C4: gist per-file limit
const MAX_TOTAL_BYTES = 5 * 1024 * 1024; // C4: gist total limit

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
    if (!/^\d+\.\d+\.\d+$/.test(manifest.version ?? "")) {
        fail(name, `manifest version '${manifest.version}' is not semver`);
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

/** C4: gist sharing is UTF-8 text only, ~1 MB per file and ~5 MB in total. */
async function checkPackagingEnvelope(name, files, dir) {
    let total = 0;
    const decoder = new TextDecoder("utf-8", { fatal: true });
    for (const file of files) {
        const buffer = await readFile(file);
        total += buffer.byteLength;
        if (buffer.byteLength > MAX_FILE_BYTES) {
            fail(name, `${path.relative(dir, file)} is ${(buffer.byteLength / 1024).toFixed(0)} KB (C4 limit 1 MB)`);
        }
        try {
            decoder.decode(buffer);
        } catch {
            fail(name, `${path.relative(dir, file)} is not valid UTF-8 (C4: gists refuse binary files)`);
        }
    }
    if (total > MAX_TOTAL_BYTES) {
        fail(name, `total ${(total / 1024 / 1024).toFixed(2)} MB exceeds the C4 limit of 5 MB`);
    }
    return total;
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
    const total = await checkPackagingEnvelope(name, files, dir);
    console.log(`${name}: ${files.length} files, ${(total / 1024).toFixed(0)} KB`);
}

if (problems.length > 0) {
    console.error(`\n${problems.length} problem(s):`);
    for (const problem of problems) console.error(`  - ${problem}`);
    process.exit(1);
}

console.log(`\nOK — ${extensions.length} extension(s) valid.`);
