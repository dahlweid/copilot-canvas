// Measures the Word suite lock: does it actually exclude, does it survive a
// holder being killed, and -- the question that matters most -- does holding it
// blind the pid assertions it is meant to protect?
//
// Issue #37's lock is only worth having if all three answers are right. The
// third is the one that could quietly make things worse: a lock that made
// `newWordPids` stop reporting would turn a flaky assertion into an
// unfalsifiable one, which this repo treats as the more serious defect.
//
// Run:
//   node spikes/isolation/probes/probe-suite-lock.mjs          # all three parts
//   node spikes/isolation/probes/probe-suite-lock.mjs --no-word  # skip part 3
//
// Part 3 needs Word. Parts 1 and 2 are Office-free.
//
// Nothing here kills a process it did not spawn. The holder killed in part 2 is
// a node child this probe started; the Word in part 3 is created and quit by
// `make-fixture.ps1` and is never signalled.

import { execFile, spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
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
    handle.release();
    process.stdout.write(`${Date.now()}\n`);
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
        child.stdout.on("data", (d) => (out += d));
        child.on("error", reject);
        child.on("exit", () => resolve(out.trim()));
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
