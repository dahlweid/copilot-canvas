// Measures the Word suite lock: does it actually exclude, does it survive a
// holder being killed, does reclamation stay exclusive when several waiters
// race for the same abandoned lock, and -- the question that matters most --
// does holding it blind the pid assertions it is meant to protect?
//
// Issue #37's lock is only worth having if all four answers are right. The
// last is the one that could quietly make things worse: a lock that made
// `newWordPids` stop reporting would turn a flaky assertion into an
// unfalsifiable one, which this repo treats as the more serious defect.
//
// Part 4 was added after review. Its absence is why a real reclamation race
// shipped in the first commit of the lock: parts 1 and 2 between them look
// like they cover concurrency and staleness, and neither covers the two at
// once, which is the only condition under which the defect appears.
//
// Run:
//   node spikes/isolation/probes/probe-suite-lock.mjs          # all four parts
//   node spikes/isolation/probes/probe-suite-lock.mjs --no-word  # skip part 3
//
// Part 3 needs Word. Parts 1, 2 and 4 are Office-free.
//
// Nothing here kills a process it did not spawn. The holder killed in part 2 is
// a node child this probe started; the Word in part 3 is created and quit by
// `make-fixture.ps1` and is never signalled.

import { execFile, spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import { unlinkSync, writeFileSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { hostname, tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

import { acquireWordSuiteLock } from "../../../.github/extensions/office-canvas/test/integration/word-suite-lock.mjs";
import { newWordPids, wordPids } from "../../../.github/extensions/office-canvas/test/integration/word-pids.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const SELF = fileURLToPath(import.meta.url);
const EXT = path.join(HERE, "..", "..", "..", ".github", "extensions", "office-canvas");

// ---------------------------------------------------------------- child mode
// Re-executed by part 1 and part 2 so the contenders are genuinely separate
// processes: an in-process test of a lock whose whole job is cross-process
// exclusion would prove nothing.
if (process.argv[2] === "--hold") {
    const holdMs = Number(process.argv[3]);
    const handle = await acquireWordSuiteLock(`child-${process.pid}`, {
        timeoutMs: 60_000,
        pollMs: 25,
        log: () => {},
    });
    process.stdout.write(`${handle.held ? "held" : "unlocked"} ${Date.now()} `);
    if (holdMs < 0) {
        // Hold forever: part 2 kills this to simulate a crashed run.
        process.stdout.write(`${process.pid}\n`);
        await new Promise(() => {});
    }
    await new Promise((r) => setTimeout(r, holdMs));
    // Stamp the end BEFORE releasing, not after. Stamping afterwards was an
    // instrument defect that manufactured overlaps: `release()` unlinks and then
    // the clock is read, and on a loaded machine this process can be descheduled
    // in between -- long enough for the next waiter to acquire and stamp its
    // start before this one stamps its end. That reported an overlap where the
    // holds were strictly sequential, and it did so in part 1, whose plain
    // `wx`-create path is exclusive by construction. Stamping first makes the
    // reported span a strict *subset* of the real hold, so the instrument can
    // under-report an overlap but can never invent one.
    process.stdout.write(`${Date.now()}\n`);
    handle.release();
    process.exit(0);
}

const lockPath = path.join(await mkdtemp(path.join(tmpdir(), "probe-suite-lock-")), "suite.lock");
process.env.OFFICE_CANVAS_WORD_LOCK_PATH = lockPath;
console.log(`lock path: ${lockPath}\n`);

// ------------------------------------------------- part 1: does it exclude?
const CONTENDERS = 8;
const HOLD_MS = 150;

function runChild(holdMs) {
    return new Promise((resolve, reject) => {
        const child = spawn(process.execPath, [SELF, "--hold", String(holdMs)], {
            env: { ...process.env, OFFICE_CANVAS_WORD_LOCK_PATH: lockPath },
        });
        let out = "";
        let err = "";
        child.stdout.on("data", (d) => (out += d));
        child.stderr.on("data", (d) => (err += d));
        child.on("error", reject);
        child.on("exit", (code) => {
            // A child that died without printing a span is not a child that ran
            // unlocked, and silently folding the two together would let a crash
            // in the lock be reported as a scheduling artefact.
            if (code !== 0 || !out.trim()) {
                const line =
                    err.split("\n").find((l) => /Error|Errno|throw/.test(l))?.trim() ??
                    err.trim().split("\n")[0] ??
                    "";
                resolve(`crashed:${code} 0 0 ${line}`);
                return;
            }
            resolve(out.trim());
        });
    });
}

console.log(`part 1: ${CONTENDERS} processes race for the lock, each holding ${HOLD_MS} ms`);
const outputs = await Promise.all(Array.from({ length: CONTENDERS }, () => runChild(HOLD_MS)));
const spans = outputs.map((line) => {
    const [state, from, to] = line.split(/\s+/);
    return { state, from: Number(from), to: Number(to) };
});
spans.sort((a, b) => a.from - b.from);

let overlaps = 0;
for (let i = 1; i < spans.length; i++) {
    if (spans[i].from < spans[i - 1].to) overlaps++;
}
const unlocked = spans.filter((s) => s.state !== "held").length;
for (const s of spans) console.log(`  ${s.state} ${s.to - s.from} ms  [${s.from} .. ${s.to}]`);
console.log(`  overlapping holds: ${overlaps}   ran unlocked: ${unlocked}`);
console.log(`  verdict: ${overlaps === 0 && unlocked === 0 ? "EXCLUSIVE" : "NOT EXCLUSIVE"}\n`);

// -------------------------------------- part 2: is a killed holder reclaimed?
console.log("part 2: a holder is killed without releasing (the crashed-run case)");
const victim = spawn(process.execPath, [SELF, "--hold", "-1"], {
    env: { ...process.env, OFFICE_CANVAS_WORD_LOCK_PATH: lockPath },
});
const victimPid = await new Promise((resolve) => {
    let buf = "";
    victim.stdout.on("data", (d) => {
        buf += d;
        const m = buf.match(/held \d+ (\d+)/);
        if (m) resolve(Number(m[1]));
    });
});
console.log(`  holder ${victimPid} has the lock; killing it (a child this probe spawned)`);
victim.kill("SIGKILL");
await new Promise((r) => victim.on("exit", r));

const reclaimStart = Date.now();
const messages = [];
const reclaimed = await acquireWordSuiteLock("probe-reclaim", {
    timeoutMs: 30_000,
    pollMs: 100,
    log: (m) => messages.push(m.trim()),
});
console.log(`  reclaimed: ${reclaimed.held} after ${Date.now() - reclaimStart} ms`);
for (const m of messages) console.log(`    ${m}`);
console.log(`  verdict: ${reclaimed.held ? "RECLAIMED" : "WEDGED"}\n`);

// ------------- part 4: several waiters race to reclaim ONE abandoned lock
// The arm that was missing, and its absence is why a real defect shipped in
// the first commit of this file.
//
// Part 1 races 8 contenders but every one of them releases normally, so the
// lock is never stale, the reclaim path is never entered, and "0 overlapping
// holds" says nothing whatever about reclamation. Part 2 does reclaim, but
// with exactly one waiter, so nothing races. The defect lives precisely where
// those two meet -- several waiters waking together on an abandoned lock --
// and that is not an exotic combination: waiters poll on a common interval, so
// a reclamation event is *when they wake together* rather than a rare
// coincidence of it.
//
// ## Why this runs rounds instead of one race
//
// The first version of this arm raced once, and it was a bad instrument: it
// caught the known defect on its first run and then passed that same defect 5
// times out of 5. A roughly 1-in-6 detector reporting a pass is
// indistinguishable from a fixed bug -- which is the very disease #37 is about
// -- so a single race would have let the next version of this defect through
// while looking like evidence. Repeating the race turns one unreliable
// observation into a reliable one.
const ROUNDS = Number(process.env.PROBE_RECLAIM_ROUNDS ?? 12);
const RECLAIMERS = 6;
console.log(`part 4: ${RECLAIMERS} processes race to reclaim ONE abandoned lock, ${ROUNDS} rounds`);

// A dead pid obtained honestly -- spawn a child, let it exit, reap it. Inventing
// a pid risks naming a live stranger, which is a different scenario from an
// abandoned run and would exercise a different branch.
const corpse = spawn(process.execPath, ["-e", ""], { stdio: "ignore" });
const deadPid = corpse.pid;
await new Promise((r) => corpse.on("exit", r));

let totalOverlaps = 0;
let totalNeverHeld = 0;
let totalCrashed = 0;
let roundsWithOverlap = 0;
let worstMs = 0;
for (let round = 1; round <= ROUNDS; round++) {
    try {
        unlinkSync(lockPath);
    } catch {}
    writeFileSync(
        lockPath,
        JSON.stringify({
            token: randomUUID(),
            pid: deadPid,
            label: "abandoned-run",
            host: hostname(),
            startedAt: Date.now(),
        }),
    );

    const spans = (await Promise.all(Array.from({ length: RECLAIMERS }, () => runChild(HOLD_MS))))
        .map((line) => {
            const [state, from, to] = line.split(/\s+/);
            const detail = line.split(/\s+/).slice(3).join(" ");
            return { state, from: Number(from), to: Number(to), detail };
        })
        .sort((a, b) => a.from - b.from);

    let overlaps = 0;
    let worstOverlapMs = 0;
    // Only spans that actually HELD can overlap meaningfully. A child that timed
    // out ran unlocked for its whole hold, so counting its span would
    // manufacture an overlap against the legitimate holder and report a false
    // red -- an instrument that cries wolf is no more use here than one that
    // sleeps through the fire. It is counted separately, below.
    const heldSpans = spans.filter((s) => s.state === "held");
    for (let i = 1; i < heldSpans.length; i++) {
        const by = heldSpans[i - 1].to - heldSpans[i].from;
        if (by > 0) {
            overlaps++;
            worstOverlapMs = Math.max(worstOverlapMs, by);
        }
    }
    const neverHeld = spans.filter((s) => s.state !== "held").length;
    const crashed = spans.filter((s) => s.state?.startsWith("crashed")).length;
    totalOverlaps += overlaps;
    totalNeverHeld += neverHeld;
    totalCrashed += crashed;
    if (crashed > 0) {
        // A child that died is not a child that waited. Reporting the two as one
        // number would let a throw inside the lock hide behind a scheduling
        // story, which is the failure mode this probe exists to refuse.
        for (const s of spans.filter((x) => x.state?.startsWith("crashed"))) {
            console.log(`  round ${String(round).padStart(2)}: child ${s.state} ${s.detail}`);
        }
    }
    worstMs = Math.max(worstMs, worstOverlapMs);
    if (overlaps > 0) {
        roundsWithOverlap++;
        console.log(`  round ${String(round).padStart(2)}: ${overlaps} OVERLAPPING HOLD(S), worst ${worstOverlapMs} ms`);
        for (const s of spans) console.log(`      ${s.state} ${s.to - s.from} ms  [${s.from} .. ${s.to}]`);
    }
}

console.log(
    `  rounds with an overlap: ${roundsWithOverlap}/${ROUNDS}   ` +
        `total overlaps: ${totalOverlaps}   worst ${worstMs} ms   never held: ${totalNeverHeld} (of which crashed: ${totalCrashed})`,
);
if (totalNeverHeld === ROUNDS * RECLAIMERS) {
    // Distinguish "the lock held correctly" from "nothing was measured". If the
    // planted owner was not judged stale, every child simply timed out and the
    // overlap count is zero for a reason that has nothing to do with the code
    // under test -- an instrument silently reporting a pass is the exact
    // disease #37 is about.
    console.log("  verdict: INCONCLUSIVE -- the planted lock was never judged stale, nothing raced\n");
} else if (totalCrashed > 0) {
    // A child that threw out of the acquire path is a defect in the lock, not a
    // quiet waiter. Folding it into "never held" would let a throw be reported
    // as patience.
    console.log(`  verdict: ${totalCrashed} CHILD CRASH(ES) in the acquire path -- see above\n`);
} else {
    console.log(
        `  verdict: ${totalOverlaps === 0 ? "RECLAIM IS EXCLUSIVE" : "TWO HOLDERS -- reclamation races"}\n`,
    );
}

// ------------------- part 3: can the pid assertion still see a Word appearing?
if (process.argv.includes("--no-word")) {
    console.log("part 3: skipped (--no-word)");
    reclaimed.release();
    process.exit(0);
}

console.log("part 3: with the lock HELD, does a Word that starts still get reported?");
console.log("        (if this says NO, the lock has made the assertion unfalsifiable)");
const workRoot = await mkdtemp(path.join(tmpdir(), "probe-suite-lock-fixture-"));
const out = path.join(workRoot, "fixture.docx");
const before = await wordPids();

// A Word really does start here: make-fixture.ps1 does New-Object -ComObject
// Word.Application. It quits its own instance; this probe only watches.
let sawDuring = [];
const watcher = (async () => {
    const deadline = Date.now() + 120_000;
    while (Date.now() < deadline && sawDuring.length === 0) {
        sawDuring = await newWordPids(before);
        // Sleep between samples. `wordPids` spawns a PowerShell per call, so an
        // unslept loop spawns them back to back for up to two minutes on a
        // shared machine -- and that load lands on the very Word start this is
        // timing. `assertNoLeakedWord` polls at 250 ms; match it.
        if (sawDuring.length === 0) await new Promise((r) => setTimeout(r, 250));
    }
})();
await execFileAsync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(EXT, "test", "integration", "make-fixture.ps1"),
    "-Out",
    out,
    "-Chapters",
    "1",
]);
await watcher;

console.log(`  lock held during this: ${reclaimed.held}`);
console.log(`  newWordPids reported: [${sawDuring.join(",")}]`);
console.log(`  verdict: ${sawDuring.length > 0 ? "STILL DETECTS -- the assertion can go red" : "BLIND -- do not ship"}`);

await rm(workRoot, { recursive: true, force: true }).catch(() => {});
reclaimed.release();
