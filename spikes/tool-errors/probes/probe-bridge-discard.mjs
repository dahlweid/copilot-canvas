#!/usr/bin/env node
// Do the two discard sites still read the way #45 measured them?
//
// The end-to-end measurement in ../FINDINGS.md is the primary evidence and does
// not depend on this file. This is the *citation* check: it asserts that the two
// blobs quoted in that document still contain the code quoted from them, so a
// CLI update that fixes -- or moves -- the discard turns the citation red
// instead of leaving a document asserting something about a version nobody runs
// any more.
//
// It reads the installed CLI, so it is machine-local by nature and is a spike
// probe rather than a test. It is not in the unit suite and cannot be: CI has no
// Copilot CLI installation to read.
//
// ## Why two different roots
//
// Measured while writing this, and worth stating because it is an easy citation
// to get wrong. The extension log for a live run reports:
//
//     COPILOT_SDK_PATH=%LOCALAPPDATA%\Programs\GitHub Copilot\copilot-sdk
//     __dir=%LOCALAPPDATA%\copilot\pkg\win32-x64\<version>
//
// So the SDK the extension *imports* is the app copy, and the runtime *hosting*
// it is the pkg blob. Hop 1 lives in the first, hop 2 in the second. Citing
// either root for both would name a file that does not contain the line.
//
// Do not look for `Tool execution failed` in `github.exe`: searched as UTF-8
// across all 302 MB of it, 0 hits. The string lives in the pkg's bundled JS.
//
// Usage:  node spikes/tool-errors/probes/probe-bridge-discard.mjs

import { readFileSync, existsSync, readdirSync } from "node:fs";
import { join } from "node:path";

const localAppData = process.env.LOCALAPPDATA;
if (!localAppData) {
    console.error("LOCALAPPDATA is not set; this probe reads the installed Copilot CLI and only runs on Windows.");
    process.exit(2);
}

/** Hop 1: the SDK reduces a thrown error to its message and drops the rest. */
const HOP_1 = {
    what: "SDK: a throw is reduced to error.message",
    path: join(localAppData, "Programs", "GitHub Copilot", "copilot-sdk", "extension.js"),
    // Split so this file's own text cannot satisfy a naive grep of the tree.
    needles: [
        "const message = error instanceof Error ? error.message : String(error);",
        "handlePendingToolCall({ requestId, error: message })",
    ],
};

/**
 * The handler call itself, in the same blob and a few lines above hop 1.
 *
 * Not a discard site -- it is here because `reportingFailures` quotes it as the
 * reason it forwards variadically. A wrapper that names one parameter drops the
 * SDK's context object, including the W3C trace headers, and no handler and no
 * test downstream can see that happen. So the quote is load-bearing and belongs
 * under the same citation check as the discards it sits between.
 *
 * A pattern, not a literal, for the same reason as hop 2: this file is the app
 * copy today and is not guaranteed to stay unminified.
 */
const CALL_SITE = {
    what: "SDK: a handler is called with (args, context)",
    path: join(localAppData, "Programs", "GitHub Copilot", "copilot-sdk", "extension.js"),
    needles: [/await handler\(\s*args\s*,\s*\{/],
};

/**
 * Hop 2's needle, as a pattern rather than a string.
 *
 * Measured while writing this, and the reason it is not a literal: the runtime
 * blob is minified, and the *same* code is spelled differently between
 * versions. 1.0.80 emits `rejectExternalTool(r,new Error(...))`, 1.0.73 emits
 * `rejectExternalTool(e.requestId,new Error(...))`. A literal needle taken from
 * one reports the discard as *absent* from the other, which is a false all-clear
 * on the exact question this file exists to ask.
 */
const HOP_2 = /rejectExternalTool\([A-Za-z0-9_$.]+,\s*new Error\("Tool execution failed"\)\)/;

/**
 * Every installed runtime, newest last.
 *
 * Only the newest is asserted on: that is the one hosting the extension, and it
 * is the version FINDINGS.md cites. The others are printed because a version
 * that stops matching is the signal that the bridge changed -- but an old blob
 * left on disk is not evidence about the CLI anyone is running, so it must not
 * turn this red.
 */
function installedRuntimes() {
    const pkgRoot = join(localAppData, "copilot", "pkg");
    if (!existsSync(pkgRoot)) return [];
    const found = [];
    for (const platform of readdirSync(pkgRoot)) {
        const platformDir = join(pkgRoot, platform);
        let versions;
        try {
            versions = readdirSync(platformDir);
        } catch {
            continue;
        }
        for (const version of versions) {
            const candidate = join(platformDir, version, "sdk", "index.js");
            if (existsSync(candidate)) found.push({ version, label: `${platform}/${version}`, path: candidate });
        }
    }
    // Numeric-aware, so 1.0.9 does not sort after 1.0.80.
    const parts = (v) => v.split(/[.-]/).map((n) => Number.parseInt(n, 10) || 0);
    found.sort((a, b) => {
        const [x, y] = [parts(a.version), parts(b.version)];
        for (let i = 0; i < Math.max(x.length, y.length); i++) {
            if ((x[i] ?? 0) !== (y[i] ?? 0)) return (x[i] ?? 0) - (y[i] ?? 0);
        }
        return 0;
    });
    return found;
}

let failures = 0;

function check(label, filePath, needles, fatal = true) {
    if (!existsSync(filePath)) {
        console.log(`MISSING  ${label}\n         ${filePath}`);
        if (fatal) failures++;
        return;
    }
    const source = readFileSync(filePath, "utf8");
    for (const needle of needles) {
        const ok = typeof needle === "string" ? source.includes(needle) : needle.test(source);
        if (!ok && fatal) failures++;
        console.log(`${ok ? "PRESENT " : "ABSENT  "} ${label}\n         ${needle}`);
    }
    console.log(`         ${filePath}\n`);
}

console.log("#45 bridge citations\n");

check("hop 1 -- " + HOP_1.what, HOP_1.path, HOP_1.needles);
check("call site -- " + CALL_SITE.what, CALL_SITE.path, CALL_SITE.needles);

const runtimes = installedRuntimes();
if (runtimes.length === 0) {
    console.log("MISSING  hop 2 -- no pkg runtime found under %LOCALAPPDATA%\\copilot\\pkg");
    failures++;
} else {
    const newest = runtimes.at(-1);
    for (const runtime of runtimes) {
        const isNewest = runtime === newest;
        check(
            `hop 2 (${runtime.label})${isNewest ? " -- newest, asserted" : " -- older, informational"}`,
            runtime.path,
            [HOP_2],
            isNewest,
        );
    }
}

console.log(failures === 0 ? "OK: every citation still resolves." : `FAILED: ${failures} citation(s) no longer resolve.`);
process.exit(failures === 0 ? 0 : 1);
