// Smoke test for the Word bridge. Run with:
//   node .github/extensions/office-canvas/test/integration/host-smoke.mjs
//
// Requires Word. Generates its own fixture, exercises every host command, and
// asserts that no WINWORD.EXE is left behind.

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { mkdtemp, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import assert from "node:assert/strict";

import { WordHost } from "../../src/word/word-host.mjs";
import { assertNoLeakedWord, killOwnedWord, ownedWordLedger, wordPids } from "./word-pids.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));

const results = [];
function check(name, fn) {
    return (async () => {
        try {
            await fn();
            results.push({ name, ok: true });
            process.stderr.write(`  ok   ${name}\n`);
        } catch (err) {
            results.push({ name, ok: false, err });
            process.stderr.write(`  FAIL ${name}\n       ${err.message}\n`);
        }
    })();
}

const workRoot = await mkdtemp(path.join(tmpdir(), "word-canvas-test-"));
const fixture = path.join(workRoot, "fixture.docx");
const pidsBefore = await wordPids();

process.stderr.write("generating fixture...\n");
await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(HERE, "make-fixture.ps1"),
    "-Out",
    fixture,
]);

const ledger = ownedWordLedger();
const host = new WordHost({
    log: (m) => process.stderr.write(`[host] ${m}\n`),
    // The only sanctioned source of "this WINWORD is ours". Everything this
    // suite kills comes from here; see word-pids.mjs for why differencing is
    // not allowed to authorize a kill.
    onOwnedPid: (pid) => ledger.record(pid),
    // Deliberately stable across runs: a host that crashed in a previous run
    // records its Word pid here, and the next run reaps it.
    pidDir: path.join(tmpdir(), "word-canvas-test-pids"),
});
const docId = "smoke";
let opened;

try {
    await check("ping starts Word and reports a version", async () => {
        const t0 = Date.now();
        const res = await host.ping();
        process.stderr.write(`       word ${res.wordVersion}, owned=${res.owned}, ${Date.now() - t0}ms\n`);
        assert.equal(res.ready, true);
        assert.ok(res.wordVersion, "expected a Word version");
    });

    await check("open returns metadata and a multi-page count", async () => {
        opened = await host.openDocument({ docId, path: fixture, workDir: path.join(workRoot, "work") });
        process.stderr.write(`       ${opened.name}: ${opened.pageCount} pages, ${opened.wordCount} words\n`);
        assert.ok(opened.pageCount > 1, `expected a multi-page fixture, got ${opened.pageCount}`);
        assert.equal(opened.title, "Word Canvas Fixture");
        assert.notEqual(path.resolve(opened.workingCopy), path.resolve(fixture), "must not open the original file");
    });

    await check("the original file is never locked or modified", async () => {
        const before = await stat(fixture);
        // If Word held the original open, this rename would fail on Windows.
        const probe = `${fixture}.probe`;
        const { renameSync } = await import("node:fs");
        renameSync(fixture, probe);
        renameSync(probe, fixture);
        const after = await stat(fixture);
        assert.equal(before.mtimeMs, after.mtimeMs, "original mtime changed");
        assert.equal(before.size, after.size, "original size changed");
    });

    await check("export produces a real PDF", async () => {
        const out = path.join(workRoot, "out.pdf");
        const t0 = Date.now();
        const res = await host.exportPdf({ docId, out });
        process.stderr.write(`       ${res.pageCount} pages, ${res.sizeBytes} bytes, ${Date.now() - t0}ms\n`);
        assert.ok(res.sizeBytes > 1000, "PDF looks too small");
        const { readFileSync } = await import("node:fs");
        assert.equal(readFileSync(out).subarray(0, 5).toString("latin1"), "%PDF-", "not a PDF");
    });

    await check("export honours a page range", async () => {
        const out = path.join(workRoot, "range.pdf");
        const full = (await stat(path.join(workRoot, "out.pdf"))).size;
        const res = await host.exportPdf({ docId, out, from: 1, to: 1 });
        assert.ok(res.sizeBytes < full, "single-page export should be smaller than the full export");
    });

    await check("outline finds headings with levels and page numbers", async () => {
        const res = await host.outline({ docId });
        process.stderr.write(`       ${res.count} headings, first: ${JSON.stringify(res.headings[0])}\n`);
        assert.ok(res.count >= 10, `expected >= 10 headings, got ${res.count}`);
        assert.ok(Array.isArray(res.headings), "headings must be an array");
        const h1 = res.headings.filter((h) => h.level === 1);
        const h2 = res.headings.filter((h) => h.level === 2);
        assert.ok(h1.length >= 4, `expected >= 4 level-1 headings, got ${h1.length}`);
        assert.ok(h2.length >= 6, `expected >= 6 level-2 headings, got ${h2.length}`);
        assert.ok(
            res.headings.every((h) => h.page >= 1),
            "every heading needs a page number",
        );
        // The whole point of the outline is jumping to the right page, so page
        // numbers must actually vary across a multi-page document.
        const pages = new Set(res.headings.map((h) => h.page));
        assert.ok(pages.size > 1, `headings all resolved to one page: ${[...pages].join(",")}`);
        // Document order, not style order.
        const starts = res.headings.map((h) => h.start);
        assert.deepEqual(starts, [...starts].sort((a, b) => a - b), "headings must be in document order");
        assert.equal(res.headings[0].text, "Chapter 1");
    });

    await check("search returns hits with page numbers and snippets", async () => {
        const res = await host.search({ docId, query: "ZORBLAX" });
        assert.equal(res.count, 1, `expected exactly one hit, got ${res.count}`);
        assert.ok(res.hits[0].page >= 1, "hit needs a page number");
        assert.match(res.hits[0].snippet, /ZORBLAX/);
        assert.doesNotMatch(res.hits[0].snippet, /[\r\n]/, "snippet must be single-line");
    });

    await check("search with no hits returns an empty array", async () => {
        const res = await host.search({ docId, query: "definitely-not-in-the-document-xyzzy" });
        assert.equal(res.count, 0);
        assert.ok(Array.isArray(res.hits), `hits must stay an array, got ${JSON.stringify(res.hits)}`);
    });

    await check("text extracts the whole document", async () => {
        const res = await host.text({ docId });
        assert.match(res.text, /Chapter 1/);
        assert.match(res.text, /ZORBLAX/);
        assert.equal(res.fromPage, 1);
    });

    await check("text extracts a single page", async () => {
        const whole = await host.text({ docId });
        const page1 = await host.text({ docId, fromPage: 1, toPage: 1 });
        assert.ok(page1.text.length > 0, "page 1 should not be empty");
        assert.ok(page1.text.length < whole.text.length, "one page must be shorter than the whole document");
    });

    await check("text rejects a page past the end", async () => {
        await assert.rejects(() => host.text({ docId, fromPage: 9999 }), /beyond the end/i);
    });

    await check("non-ASCII survives the stdio round trip", async () => {
        const res = await host.search({ docId, query: "Muenchen" });
        assert.ok(res.count > 0, "expected to find the umlaut-check sentence");
    });

    await check("info reports document properties", async () => {
        const res = await host.info({ docId });
        assert.equal(res.title, "Word Canvas Fixture");
        assert.ok(res.pageCount > 1);
        assert.ok(res.paragraphs > 10);
    });

    await check("a missing file produces a clean error, not a hang", async () => {
        await assert.rejects(
            () => host.openDocument({ docId: "missing", path: path.join(workRoot, "nope.docx"), workDir: workRoot }),
            /File not found/i,
        );
    });

    await check("an unknown command is reported, not swallowed", async () => {
        await assert.rejects(() => host.request("frobnicate", {}), /Unknown command/i);
    });

    await check("commands recover after the host process dies", async () => {
        // Simulates a Word crash / stuck-dialog teardown: the bridge must
        // transparently restart and replay the open.
        await host.request("__force_kill_for_test", {}).catch(() => {});
        killOwnedWord(ledger);
        const res = await host.info({ docId });
        assert.ok(res.pageCount > 1, "expected the bridge to recover and reopen the document");
    });

    await check("close releases the document", async () => {
        const res = await host.closeDocument({ docId });
        assert.equal(res.closed, true);
    });

    // The reap in Clear-OrphanedWord is the most exposed kill in the host: a pid
    // file outlives a crashed host by an unbounded interval, so "there is a
    // WINWORD at this pid" says even less there than it does inside Stop-Word's
    // bounded wait. Pid reuse cannot be forced, but the state it produces can:
    // a live WINWORD at the recorded pid whose identity is not the recorded one.
    // Recording the wrong identity for a Word we know is alive reaches the same
    // branch, and the Word at risk is ours, so a failure damages nothing but us.
    await check("the orphan sweep does not kill a Word whose identity does not match", async () => {
        const { ownedPid } = await host.ping();
        assert.ok(ownedPid, "expected the host to report the pid of the Word it owns");

        // A measured dead pid rather than an assumed one: this child has exited
        // by the time its own $PID comes back to us.
        const { stdout } = await execFileAsync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", "$PID"]);
        const deadHostPid = Number.parseInt(stdout.trim(), 10);
        assert.ok(Number.isInteger(deadHostPid) && deadHostPid > 0, "expected a pid from the probe child");

        const reapDir = await mkdtemp(path.join(tmpdir(), "word-canvas-reap-"));
        // Two entries, both naming our live Word, reaching the two ways identity
        // can fail to be proved: a recorded start time that is not this Word's,
        // and a single-field file from a host predating the identity field --
        // which is what an upgrade leaves lying in the pid directory. Both must
        // decline, and one sweep covers both.
        const stale = [
            { path: path.join(reapDir, `${deadHostPid}.pid`), body: `${ownedPid} 1`, why: "mismatched start time" },
            { path: path.join(reapDir, `${deadHostPid + 1}.pid`), body: `${ownedPid}`, why: "legacy entry with no identity" },
        ];
        for (const entry of stale) await writeFile(entry.path, entry.body, "utf8");

        const reaper = new WordHost({
            log: (m) => process.stderr.write(`[reaper] ${m}\n`),
            onOwnedPid: (pid) => ledger.record(pid),
            pidDir: reapDir,
        });
        try {
            await reaper.ping(); // startup runs the sweep
        } finally {
            await reaper.dispose();
        }

        // Assert the branch was reached before asserting on what it did. The
        // sweep deletes every entry it processes, so a surviving file means it
        // skipped that one -- its recording pid was reused and looked alive --
        // and the check below would then be green without having tested
        // anything. Skipped-but-green is the failure a guard test can least
        // afford, so an unprocessed entry is a hard failure and not a skip.
        for (const entry of stale) {
            const processed = await stat(entry.path).then(() => false, () => true);
            assert.equal(processed, true, `the sweep never processed the ${entry.why} entry, so it proved nothing`);
        }

        assert.ok(
            (await wordPids()).includes(ownedPid),
            `the sweep killed pid ${ownedPid} without having proved it owned it`,
        );
        await rm(reapDir, { recursive: true, force: true }).catch(() => {});
    });
} finally {
    await host.dispose();
}

await check("no new WINWORD.EXE is left behind", async () => {
    // Print the attribution before asserting on it. When this fails under
    // concurrency the first question is whether the host named the right pid,
    // and that is not reconstructable after the fact.
    process.stderr.write(`       host reported owning: ${ledger.pids().join(", ") || "(none)"}\n`);
    await assertNoLeakedWord(pidsBefore, { ledger });
});

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = results.filter((r) => !r.ok);
process.stderr.write(`\n${results.length - failed.length}/${results.length} passed\n`);
process.exit(failed.length === 0 ? 0 : 1);
