// Smoke test for the Word bridge. Run with:
//   node .github/extensions/office-canvas/test/integration/host-smoke.mjs
//
// Requires Word. Generates its own fixture, exercises every host command, and
// asserts that no WINWORD.EXE is left behind.

import { execFile, execFileSync } from "node:child_process";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { mkdtemp, readFile, readdir, rm, stat, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import assert from "node:assert/strict";

import { WordHost } from "../../src/word/word-host.mjs";
import { codepoints } from "./docx-zip.mjs";
import { assertNoLeakedWord, killOwnedWord, ownedWordLedger, wordPids } from "./word-pids.mjs";
import { acquireWordSuiteLock } from "./word-suite-lock.mjs";

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

/**
 * Waits for one WINWORD pid to disappear, and reports how long that took.
 *
 * "This Word is gone" is a **state at a deadline**, which is the kind of claim a
 * poll can settle -- unlike "no Word appeared during this interval", for which
 * no later moment makes it true. #37 measured the difference and CONTEXT.md
 * records it beside `probe-suite-contention.mjs`; sampling once here would be
 * the avoidable half of that.
 *
 * The deadline matches `assertNoLeakedWord`'s for the same reason: process exit
 * on this machine is load-dependent and not bounded by idle measurements, and
 * polling makes a generous deadline free on the runs that pass. Returns rather
 * than asserts, so the caller can say what the timeout means in its own terms.
 */
async function waitForWordGone(pid, { timeoutMs = 90000, intervalMs = 250 } = {}) {
    const started = Date.now();
    const deadline = started + timeoutMs;
    let present = (await wordPids()).includes(pid);
    while (present && Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, intervalMs));
        present = (await wordPids()).includes(pid);
    }
    return { gone: !present, waitedMs: Date.now() - started };
}

/**
 * Ends a PowerShell host process this suite started, and waits for it to go.
 *
 * `pid` must come from the host's own record -- word-host.ps1 names its pid file
 * `$PID.pid`, so the file name in a pid directory only this suite created is the
 * host reporting its own identity, not something differenced out of a process
 * list. The name check is the same one `killOwnedWord` makes and for the same
 * reason: pids are reused.
 *
 * It waits, and re-asks with `Get-Process`, because the caller's next step
 * depends on `Clear-OrphanedWord` seeing this pid as dead -- and that is the
 * predicate the sweep itself uses. `Kill()` only requests termination.
 *
 * Never used on a WINWORD. Word is killed through `killOwnedWord` alone.
 */
function killHostProcess(hostPid) {
    assert.ok(Number.isInteger(hostPid) && hostPid > 4, `refusing to kill '${hostPid}', which is not a pid`);
    return String(
        execFileSync("powershell.exe", [
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            `$p = Get-Process -Id ${hostPid} -ErrorAction SilentlyContinue; ` +
                "if (-not $p) { 'gone' } " +
                "elseif ($p.ProcessName -ne 'powershell') { 'notpowershell' } " +
                "else { $p.Kill(); $null = $p.WaitForExit(30000); " +
                `if (Get-Process -Id ${hostPid} -ErrorAction SilentlyContinue) { 'still-running' } else { 'killed' } }`,
        ]) ?? "",
    ).trim();
}

const workRoot = await mkdtemp(path.join(tmpdir(), "word-canvas-test-"));
const fixture = path.join(workRoot, "fixture.docx");
// Before the census, never after: every pid assertion below diffs against it,
// and a neighbour's Word starting in between lands inside the window. See
// word-suite-lock.mjs. Released when this process exits.
await acquireWordSuiteLock("host-smoke");
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
        process.stderr.write(
            `       word ${res.wordVersion}, owned=${res.owned}, via ${res.attribution}, ${Date.now() - t0}ms\n`,
        );
        assert.equal(res.ready, true);
        assert.ok(res.wordVersion, "expected a Word version");
    });

    // Ownership is asserted on *how* it was decided, not only on what came back.
    // A pid alone cannot carry that: the differencing this replaced returned a
    // plausible integer on every run, including the runs where it named a
    // stranger's Word, so `assert.ok(ownedPid)` was green throughout the defect
    // it was supposed to catch. `attribution` is the observable that separates a
    // sound answer from a lucky one.
    await check("ownership is attributed by window handle, not by inference", async () => {
        const res = await host.ping();
        assert.equal(
            res.attribution,
            "hwnd",
            `expected ownership proven from our own window handle, got '${res.attribution}'. ` +
                "'attached' means we found an existing Word (start this suite with no Word running); " +
                "'unattributed' means attribution failed, so teardown loses its kill fallback.",
        );
        assert.ok(Number.isInteger(res.ownedPid) && res.ownedPid > 4, "expected a real pid");

        // The census negative, checked from the outside: a pid that was already
        // alive before this suite started cannot be one we created. This is the
        // assertion that fails if attribution ever regresses to picking a
        // stranger -- which is precisely what differencing did, measured, when a
        // second Word was created concurrently.
        assert.ok(
            !pidsBefore.includes(res.ownedPid),
            `the host claims to own pid ${res.ownedPid}, but that Word was already running before this suite started`,
        );
        assert.ok(
            (await wordPids()).includes(res.ownedPid),
            `the host reported owning pid ${res.ownedPid}, but no WINWORD is running there`,
        );
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
        // It said `Muenchen` -- no umlaut in it -- and so proved nothing about
        // the property in its own name. The whole of issue #40 lived through two
        // merged layers behind tests like this one.
        //
        // Two probes, because the boundary has two directions and one request
        // cannot separate them. Both terms are built from codepoints so neither
        // assertion depends on how this file's own bytes are decoded, which is
        // the class of question being measured.
        const strasse = `Stra${String.fromCharCode(0x00df)}e`;
        const umlautWord = `Gr${String.fromCharCode(0x00fc, 0x00df)}e`;

        // Outbound, and deliberately reached with an **ASCII** query. The
        // obvious construction -- one non-ASCII search, asserting the echoed
        // query and the snippet together -- cannot do this job, and measuring it
        // is what showed why: with stdin broken the search matches nothing, so
        // there are no hits and the snippet assertion never runs. It is
        // unreachable in precisely the case it exists to rule out.
        //
        // An ASCII query that matches a paragraph whose *text* is non-ASCII has
        // no inbound dependency at all, so this stays green through an inbound
        // regression and goes red only for an outbound one.
        const outbound = await host.search({ docId, query: "UMLAUTMARKER" });
        assert.ok(outbound.count > 0, "the fixture has no UMLAUTMARKER paragraph to read back");
        assert.ok(
            outbound.hits.some((h) => h.snippet.includes(umlautWord)),
            `text the host read out of the document lost its umlauts on the way back: ${outbound.hits
                .map((h) => codepoints(h.snippet))
                .join(" | ")}`,
        );

        // Inbound. `query` is echoed from the host's own decoding of the bytes
        // we wrote to its stdin, so a mismatch here is the defect in #40.
        //
        // It is not a *pure* inbound probe -- the echo crosses stdout on the way
        // back too -- and the two failures are still told apart, by signature
        // rather than by site. Measured: with InputEncoding unset the sz arrives
        // as `U+251C U+0192`, two characters where one was sent, because UTF-8
        // bytes were decoded as OEM and expanded. With OutputEncoding unset
        // instead it arrives as a single `U+FFFD`, because one OEM byte was
        // decoded as UTF-8 and was not valid. Expansion means inbound;
        // replacement means outbound. The control above then confirms which.
        const inbound = await host.search({ docId, query: strasse });
        assert.equal(inbound.query, strasse, `the host decoded the query as ${codepoints(inbound.query)}, sent ${codepoints(strasse)}`);
        assert.ok(inbound.count > 0, `searching for ${codepoints(strasse)} found nothing, though the fixture contains it`);
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
            {
                path: path.join(reapDir, `${deadHostPid}.pid`),
                body: `${ownedPid} 1`,
                why: "mismatched start time",
                expect: "does not match the recorded",
            },
            {
                path: path.join(reapDir, `${deadHostPid + 1}.pid`),
                body: `${ownedPid}`,
                why: "legacy entry with no identity",
                expect: "no start time was ever recorded",
            },
        ];
        for (const entry of stale) await writeFile(entry.path, entry.body, "utf8");

        const reaperLog = [];
        const reaper = new WordHost({
            log: (m) => {
                reaperLog.push(m);
                process.stderr.write(`[reaper] ${m}\n`);
            },
            onOwnedPid: (pid) => ledger.record(pid),
            pidDir: reapDir,
        });
        try {
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

            // Surviving is a SHARED outcome: both entries name the same Word, so
            // either guard alone keeps it alive and the assertion above stays
            // green while the other branch is gone. A mutation check found that
            // -- removing the null-identity guard left the legacy entry declined
            // anyway by the comparison below it, and nothing reddened. So each
            // entry must be attributed to the branch that declined it, and the
            // decline has to be found on the log the extension actually reads,
            // which is the channel a refusal is worthless without.
            for (const entry of stale) {
                const reported = reaperLog.some(
                    (line) => line.includes(`refusing to reap pid ${ownedPid}`) && line.includes(entry.expect),
                );
                assert.equal(
                    reported,
                    true,
                    `the ${entry.why} entry was not reported as declined for its own reason; the log said: ${
                        reaperLog.join(" | ") || "(nothing)"
                    }`,
                );
            }
        } finally {
            // A failing assertion above must not also leak a temp directory:
            // this test runs on every suite run, so the accumulation would be
            // silent and unbounded.
            await rm(reapDir, { recursive: true, force: true }).catch(() => {});
        }
    });

    // The positive half of that pair, and the branch that actually carries the
    // weight. The case above proves the sweep *declines* a Word it cannot prove
    // it owns; nothing proved it reaps one it can. When this branch fails the
    // consequence is not a process that lingers -- it is a hidden WINWORD that
    // nothing will ever end, which is on this repo's short list of real bugs --
    // and by the standard `Get-WordStartTime`'s own comment sets, an instrument
    // that has never produced a positive cannot tell you the branch works.
    //
    // Every destructive step here is authorised from what the host itself
    // recorded: the pid it proved from its own window handle, plus the start
    // time that turns that coordinate into an identity
    // (word-host.ps1 `Get-WordStartTime` / `Stop-VerifiedWord`). Nothing is
    // differenced. The header of word-pids.mjs sets out why that is not a
    // stylistic preference -- measured in
    // `spikes/isolation/probes/probe-word-ownership.ps1`, differencing reported
    // 2 new pids for the 1 instance created, so a set difference cannot name a
    // kill target however plausible the integer it hands back looks.
    await check("the orphan sweep reaps a Word its ledger entry proves it owns", async () => {
        // The bystander control, captured before anything is destroyed. This
        // Word is a WINWORD, it is alive, and it is *not* named by the entry the
        // reaper will read -- so it is the observation that separates "reaped
        // the Word the ledger named" from "swept the machine".
        const bystander = (await host.ping()).ownedPid;
        assert.ok(bystander, "expected this suite's own host to report the Word it owns");
        assert.ok(
            (await wordPids()).includes(bystander),
            `the bystander Word ${bystander} was already gone, so it can control nothing`,
        );

        const victimDir = await mkdtemp(path.join(tmpdir(), "word-canvas-orphan-"));
        const victimLedger = ownedWordLedger();
        const victim = new WordHost({
            log: (m) => process.stderr.write(`[victim] ${m}\n`),
            // Recorded twice on purpose. The suite ledger is what the final leak
            // assertion reads; the private one scopes this case's own cleanup to
            // the Word it created, so a failure here cannot reach into the rest
            // of the suite and kill a Word another check still needs.
            onOwnedPid: (pid) => {
                ledger.record(pid);
                victimLedger.record(pid);
            },
            pidDir: victimDir,
        });
        const reaperLog = [];
        let reaper = null;
        try {
            const started = await victim.ping();
            assert.equal(
                started.attribution,
                "hwnd",
                `the victim host proved nothing about which Word it started (got '${started.attribution}'), ` +
                    "so there is no identity to reap by",
            );
            const orphanPid = started.ownedPid;
            assert.ok(Number.isInteger(orphanPid) && orphanPid > 4, "expected a real pid for the Word to orphan");
            assert.notEqual(orphanPid, bystander, "the victim attached to the bystander instead of creating a Word");
            assert.ok(
                !pidsBefore.includes(orphanPid),
                `the victim claims to own pid ${orphanPid}, which was already running before this suite started`,
            );

            // The ledger entry exactly as the host wrote it: the file *name* is
            // the recording host's own $PID, the contents are "<wordPid>
            // <startTicks>". Read rather than constructed -- a synthesized entry
            // would test this suite's idea of the format, and the format is the
            // thing the reap trusts.
            const names = (await readdir(victimDir)).filter((n) => n.endsWith(".pid"));
            assert.equal(names.length, 1, `expected one ledger entry, got: ${names.join(", ") || "(none)"}`);
            const entryPath = path.join(victimDir, names[0]);
            const victimHostPid = Number.parseInt(path.basename(names[0], ".pid"), 10);
            const [recordedPid, recordedTicks] = (await readFile(entryPath, "utf8")).trim().split(/\s+/);
            assert.equal(
                Number(recordedPid),
                orphanPid,
                "the ledger entry names a different Word than the host attributed by window handle",
            );
            // Without this the sweep declines with "no start time was ever
            // recorded", the Word survives, and every assertion below would be
            // testing the legacy path the case above already covers.
            assert.match(
                recordedTicks ?? "",
                /^[1-9][0-9]*$/,
                `the entry recorded no start time (${JSON.stringify(recordedTicks ?? null)}), so a pid is all it has ` +
                    "and the sweep will refuse it -- this case would then prove nothing",
            );

            // The bridge carries its own reaper: `#onExit` kills `ownedPid` when
            // the host process dies, and `dispose()` does the same. That net sits
            // in *front* of the one under test, so leaving it armed would make
            // this case green whichever of the two fired, with no way to tell
            // them apart. Forgetting the pid here is what the PowerShell-side
            // ledger exists for in the first place: the Node process went away
            // too, so nothing in this address space remembers the Word.
            victim.ownedPid = null;

            const ended = killHostProcess(victimHostPid);
            assert.equal(
                ended,
                "killed",
                `could not end the victim host process ${victimHostPid} (${ended}); the sweep skips an entry whose ` +
                    "recording host is still alive, so it would never have reached the Word",
            );

            // A genuine orphan, on every count the ledger comment names: its host
            // is gone, no client holds it, no document is open, and nothing in
            // this process remembers it.
            assert.ok(
                (await wordPids()).includes(orphanPid),
                `Word ${orphanPid} was already gone before the sweep ran, so there was no orphan to reap`,
            );

            reaper = new WordHost({
                log: (m) => {
                    reaperLog.push(m);
                    process.stderr.write(`[reaper] ${m}\n`);
                },
                onOwnedPid: (pid) => ledger.record(pid),
                pidDir: victimDir,
            });
            await reaper.ping(); // startup runs the sweep

            // Reached, before what it did -- the same guard the case above needs
            // and for the same reason. The sweep deletes every entry it
            // processes, so a surviving file means it skipped this one, and the
            // rest of this case would then be settled by whatever else happened
            // to be true.
            const processed = await stat(entryPath).then(() => false, () => true);
            assert.equal(processed, true, "the sweep never processed the orphan's ledger entry, so it proved nothing");

            // A decline is the failure this case is the inverse of, and it is
            // reported on the log the extension actually reads. Naming the pid
            // keeps this from being satisfied by a refusal about some other one.
            const declined = reaperLog.filter((line) => line.includes(`refusing to reap pid ${orphanPid}`));
            assert.deepEqual(
                declined,
                [],
                `the sweep refused the one entry whose identity it could prove: ${declined.join(" | ")}`,
            );

            // The positive instrument, and the assertion this case turns on.
            // Everything above is satisfied without a kill: the entry file is
            // removed on every outcome, and a refusal is silent on 'gone' as
            // well as on 'killed' -- `Stop-VerifiedWord` returns 'gone' both for
            // no process at that pid and for a non-Word one. So without this
            // line the case would rest on a Word being unable to end itself
            // inside 90 s, which is exactly what word-host.ps1's own note
            // declines to claim from one 45 s observation. Presence-shaped on
            // purpose: an absence check here would also pass if the log were
            // never fed at all.
            const reaped = reaperLog.filter((line) => line.includes(`reaped orphaned WINWORD ${orphanPid}`));
            assert.equal(
                reaped.length,
                1,
                `the sweep never reported reaping Word ${orphanPid}, so nothing separates "the sweep ended it" from ` +
                    `"it was already gone". The reaper said: ${reaperLog.join(" | ") || "(nothing)"}`,
            );

            // Confirmation, not the evidence: the line above says the terminate
            // was accepted, and `Kill()` only requests termination.
            const { gone, waitedMs } = await waitForWordGone(orphanPid);
            process.stderr.write(`       orphan ${orphanPid} gone ${(waitedMs / 1000).toFixed(1)}s after the sweep\n`);
            assert.equal(
                gone,
                true,
                `the sweep reported reaping Word ${orphanPid}, but that Word was still running ` +
                    `${(waitedMs / 1000).toFixed(1)}s later`,
            );

            // And it reaped *that* Word, not Word in general. The sweep read one
            // entry; the bystander is a live WINWORD no entry named, so it must
            // still be here. This is the assertion that would go red if the reap
            // ever widened from an identity to a process name.
            assert.ok(
                (await wordPids()).includes(bystander),
                `the sweep also ended pid ${bystander}, which none of the entries it read named`,
            );
        } finally {
            if (reaper) await reaper.dispose().catch(() => {});
            await victim.dispose().catch(() => {});
            // This case deliberately creates a Word that nothing else in the
            // suite will end, so a failure above must not strand one on the
            // machine. Through the sanctioned path -- a ledger fed only by the
            // host's own attribution -- and scoped to this case's Word alone. A
            // no-op when the sweep did its job: the pid reports 'gone'.
            killOwnedWord(victimLedger);
            await rm(victimDir, { recursive: true, force: true }).catch(() => {});
        }
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
