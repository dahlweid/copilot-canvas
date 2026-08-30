// probe-close-during-retry.mjs -- what an un-cancelled #autoRefresh actually
// reaches after ViewerInstance.close() returns.
//
// Run: node spikes/viewer-disposal/probes/probe-close-during-retry.mjs
// Needs a real, licensed Word. Three arms.
//
// ## Why this probe exists
//
// #81 measured the continuation with a **fake** RenderCache and concluded it was
// inert. That conclusion is only as strong as the fake: a fake cannot open a
// file, so the most damaging consequence of a continuation is invisible to it by
// construction. The follow-up hypothesis on #81 is that a real continuation goes
// through `RenderCache.refresh` -> `WordHost` -> opens the document, and that
// this is what made a freshly re-opened panel fail with
//
//     Another process is holding demo.docx open more strictly than Word does.
//
// eight seconds before the same open succeeded unaided.
//
// That is two claims, and they are separable. This probe separates them,
// because the second can be false while the first is true:
//
//   arm 1  does the loop outlive close() and enter WordHost at all?
//   arm 2  does a WordHost open of a document make that document refuse a
//          concurrent `Copy-Item` -- i.e. can it *be* the reported failure?
//   arm 3  positive control: can arm 2's counter count a failure at all?
//
// Arm 3 is not ceremony. Arm 2 answers with a count, and a count of zero from a
// dead instrument reads exactly like a count of zero from a live one.
//
// ## What is real here and what is injected
//
// The RenderCache, the WordHost, the FileWatcher, the document and the Word
// process are all real. One thing is injected: `cache.refresh` is made to throw
// `file_locked` for the *first two* calls, to arm the retry loop deterministically
// without depending on a race. The call that matters -- the one after `close()`
// returns -- runs the genuine `RenderCache.refresh`, so the decisive observation
// is made on unmodified code.
//
// `WordHost.openDocument` is wrapped to count and timestamp entries. The wrapper
// delegates to the real method; it changes nothing about what happens.

import assert from "node:assert/strict";
import { execFile, spawn } from "node:child_process";
import { mkdtemp, rm, stat, utimes } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { RenderCache } from "../../../.github/extensions/office-canvas/src/render-cache.mjs";
import { ViewerInstance } from "../../../.github/extensions/office-canvas/src/server.mjs";
import {
    assertNoLeakedWord,
    ownedWordLedger,
    wordPids,
} from "../../../.github/extensions/office-canvas/test/integration/word-pids.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT = path.resolve(HERE, "../../../.github/extensions/office-canvas");

const say = (line) => process.stderr.write(`${line}\n`);
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const results = [];
async function arm(name, fn) {
    try {
        const detail = await fn();
        results.push({ name, ok: true });
        say(`  ok    ${name}`);
        if (detail) for (const line of [].concat(detail)) say(`        ${line}`);
    } catch (err) {
        results.push({ name, ok: false, err });
        say(`  FAIL  ${name}`);
        say(`        ${err.message}`);
    }
}

async function waitFor(predicate, label, budgetMs = 30_000) {
    const deadline = Date.now() + budgetMs;
    while (Date.now() < deadline) {
        if (predicate()) return;
        await sleep(20);
    }
    throw new Error(`timed out after ${budgetMs}ms waiting for ${label}`);
}

const workRoot = await mkdtemp(path.join(tmpdir(), "viewer-disposal-"));
const fixture = path.join(workRoot, "demo.docx");
const home = path.join(workRoot, "home");
// `addRecent` writes under $COPILOT_HOME on every open; keep it out of the real one.
process.env.COPILOT_HOME = home;

const pidsBefore = await wordPids();
say(`WINWORD.EXE before: ${pidsBefore.length} (${pidsBefore.join(", ") || "none"})`);

say("generating fixture...");
await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(EXT, "test/integration/make-fixture.ps1"),
    "-Out",
    fixture,
]);

const cache = new RenderCache({
    cacheRoot: path.join(workRoot, "artifacts"),
    log: (m) => say(`        [cache] ${m}`),
});
const ledger = ownedWordLedger().watch(cache.host);

// --- instrumentation ---------------------------------------------------------

/** Every entry into the real WordHost.openDocument, with a timestamp. */
const hostOpens = [];
const realHostOpen = cache.host.openDocument.bind(cache.host);
cache.host.openDocument = async (args) => {
    hostOpens.push(Date.now());
    return realHostOpen(args);
};

/** Every call into RenderCache.refresh, with a timestamp. */
const refreshCalls = [];
const realRefresh = cache.refresh.bind(cache);
let scriptedFailures = 0;
cache.refresh = async (docPath) => {
    refreshCalls.push(Date.now());
    if (scriptedFailures > 0) {
        scriptedFailures -= 1;
        const err = new Error("file_locked injected by the probe to arm the retry loop");
        err.code = "file_locked";
        throw err;
    }
    return realRefresh(docPath);
};

let closedAt = 0;
const after = (stamps) => stamps.filter((t) => t >= closedAt).length;

try {
    // --- arm 1 --------------------------------------------------------------
    // Asserted in the direction we want the code to hold, not in the direction
    // it currently behaves: this arm **fails on today's main**, which is what
    // makes it evidence rather than a description, and passes once close()
    // cancels. A probe whose assertion matches the defect would go red the day
    // the defect is fixed.
    await arm("nothing reaches cache.refresh or WordHost after close() returns", async () => {
        const viewer = new ViewerInstance({
            cache,
            instanceId: "probe",
            workspacePath: null,
            // The shipped delays, deliberately: the question is about the real
            // ~6s window, not about a collapsed one.
            autoRefreshDelaysMs: [500, 1500, 4000],
            log: (m) => say(`        [viewer] ${m}`),
        });
        await viewer.openDocument(fixture);
        assert.equal(viewer.status, "ready", "precondition: the document opened");
        const opensAfterOpen = hostOpens.length;
        assert.ok(opensAfterOpen >= 1, "precondition: opening the document entered WordHost");

        // Two injected failures: one consumed before close(), one after -- so the
        // *third* attempt, ~2s past close(), runs the genuine RenderCache.refresh.
        scriptedFailures = 2;

        // A real mtime change: the watcher fingerprints `mtimeMs|size`, and the
        // render key is `hash(path|mtimeMs|size)`, so this both fires the watcher
        // and guarantees `RenderCache.open` cannot short-circuit on an unchanged
        // key. Touching rather than writing keeps the bytes a valid .docx.
        const before = await stat(fixture);
        const moved = new Date(before.mtimeMs + 5000);
        await utimes(fixture, moved, moved);

        await waitFor(() => refreshCalls.length >= 1, "the watcher to deliver and the first attempt to fail");

        // Mid-retry by construction: attempt 0 has just thrown and the loop is
        // sitting in its 500ms backoff.
        closedAt = Date.now();
        await viewer.close();
        const closeReturnedAt = Date.now();
        const refreshesAtClose = refreshCalls.length;
        const opensAtClose = hostOpens.length;

        // Longer than the ~6s the backoff sums to, so the whole window is observed.
        await sleep(12_000);

        const refreshesAfter = after(refreshCalls);
        const opensAfter = after(hostOpens);
        const lines = [
            `close() returned in ${closeReturnedAt - closedAt}ms`,
            `cache.refresh: ${refreshesAtClose} before close, ${refreshesAfter} after it`,
            `WordHost.openDocument: ${opensAtClose} before close, ${opensAfter} after it`,
            `viewer.status after the window: ${viewer.status}` +
                (viewer.error ? ` (${viewer.error.code})` : ""),
            `scripted failures left unconsumed: ${scriptedFailures}`,
        ];
        // Both counts are reported before either is asserted. Asserting them in
        // sequence would hide the second number behind the first failure, and
        // the second is the one #81's fake cache could not see.
        assert.deepEqual(
            { refreshesAfterClose: refreshesAfter, wordHostOpensAfterClose: opensAfter },
            { refreshesAfterClose: 0, wordHostOpensAfterClose: 0 },
            "work landed on a disposed viewer after close() returned",
        );
        return lines;
    });

    // --- arm 2 --------------------------------------------------------------
    await arm("a WordHost open of a document does NOT make it refuse a concurrent Copy-Item", async () => {
        const dst = path.join(workRoot, "hammer-copy.docx");
        const hammer = execFileAsync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path.join(HERE, "probe-copy-hammer.ps1"),
            "-Src",
            fixture,
            "-Dst",
            dst,
            "-Seconds",
            "10",
        ]);

        // Three real opens of the *original* through the real host, inside the
        // hammer's window. Open-Doc copies the original every time, so each of
        // these is exactly the file access a continuation would make.
        await sleep(500);
        const opensBefore = hostOpens.length;
        for (let i = 0; i < 3; i++) {
            await cache.host.openDocument({
                docId: `probe-arm2-${i}`,
                path: fixture,
                workDir: path.join(workRoot, `arm2-${i}`),
            });
        }
        const opensDuring = hostOpens.length - opensBefore;

        const { stdout } = await hammer;
        const report = JSON.parse(stdout.trim());
        const types = [...new Set(report.failures.map((f) => f.type))];
        const lines = [
            `${opensDuring} WordHost opens of the original during the window`,
            `Copy-Item: ${report.attempts} attempts, ${report.failures.length} refused` +
                (types.length ? ` (${types.join(", ")})` : ""),
        ];
        assert.equal(
            report.failures.length,
            0,
            `a WordHost open refused ${report.failures.length} of ${report.attempts} copies (${types.join(", ")})`,
        );
        return lines;
    });

    // --- arm 3 --------------------------------------------------------------
    await arm("positive control: the same counter does count a genuinely strict holder", async () => {
        const dst = path.join(workRoot, "hammer-control.docx");
        const hammer = execFileAsync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            path.join(HERE, "probe-copy-hammer.ps1"),
            "-Src",
            fixture,
            "-Dst",
            dst,
            "-Seconds",
            "6",
        ]);

        await sleep(500);
        const locker = spawn(
            "powershell.exe",
            [
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                path.join(HERE, "probe-strict-lock.ps1"),
                "-Path",
                fixture,
                "-Seconds",
                "3",
            ],
            { stdio: ["ignore", "pipe", "inherit"] },
        );
        await new Promise((resolve, reject) => {
            locker.stdout.on("data", (chunk) => {
                if (String(chunk).includes("locked")) resolve();
            });
            locker.on("error", reject);
            locker.on("exit", () => resolve());
        });
        await new Promise((resolve) => locker.on("exit", resolve));

        const { stdout } = await hammer;
        const report = JSON.parse(stdout.trim());
        const types = [...new Set(report.failures.map((f) => f.type))];
        const lines = [
            `Copy-Item: ${report.attempts} attempts, ${report.failures.length} refused` +
                (types.length ? ` (${types.join(", ")})` : ""),
        ];
        assert.ok(
            report.failures.length > 0,
            "the counter never fired against FileShare::None, so arm 2's zero means nothing",
        );
        assert.ok(
            types.includes("System.IO.IOException"),
            `expected System.IO.IOException -- the type word-host.ps1 maps to file_locked -- got ${types.join(", ")}`,
        );
        return lines;
    });
} finally {
    cache.refresh = realRefresh;
    await cache.dispose().catch((err) => say(`dispose failed: ${err.message}`));
}

const pidsAfter = await wordPids();
say(`WINWORD.EXE after: ${pidsAfter.length} (${pidsAfter.join(", ") || "none"})`);

await arm("this probe left no Word process behind", async () => {
    await assertNoLeakedWord(pidsBefore, { ledger });
    return [`host reported owning: ${ledger.pids().join(", ") || "none"}`];
});

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
say(`\n${results.length} arm(s), ${failed.length} failed`);
process.exit(failed.length ? 1 : 0);
