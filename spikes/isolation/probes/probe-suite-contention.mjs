// Reproduces issue #37's false red on demand, and turns it into a rate.
//
// The issue's evidence is `read-smoke` scoring 17/18, 18/18, 17/18 on an
// unchanged tree "with two sibling sessions driving Word probes throughout".
// That is a real measurement and it is not reproducible on demand: it depends
// on what someone else's session happened to be doing. So it cannot be used to
// tell whether a fix worked -- a green run on a quiet machine is green either
// way, and it is exactly the 17-in-18 case.
//
// This supplies the missing half: a **neighbour** that drives Word on purpose,
// so the contention is a controlled variable rather than the weather. The
// neighbour is `make-fixture.ps1` in a loop, which is what a sibling session's
// suite does -- `New-Object -ComObject Word.Application`, work, `Quit()`.
//
// Two arms, and the difference between them is the whole claim:
//
//   --neighbour unlocked   the neighbour ignores the suite lock, as any session
//                          on a branch predating word-suite-lock.mjs does.
//   --neighbour locked     the neighbour holds the lock around each pass, as a
//                          lock-aware sibling suite does.
//
// Run the first against a tree without the lock, the second against a tree with
// it, and the two pass rates are the before and after.
//
// It also reports *which* check failed, because #37 records two different
// assertions being conflated: `assertNoLeakedWord` at the end of the run, and
// the "a missing file is reported without starting Word" check at the start,
// which is an assertion of absence over an interval and cannot be rescued by
// any deadline.
//
// Run:
//   node spikes/isolation/probes/probe-suite-contention.mjs 18 --neighbour unlocked
//   node spikes/isolation/probes/probe-suite-contention.mjs 18 --neighbour locked
//   node spikes/isolation/probes/probe-suite-contention.mjs 6  --neighbour none
//
// The neighbour is a child this probe spawns and is the only process it stops.
// Every Word involved is created and quit by make-fixture.ps1 itself.

import { spawn } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { acquireWordSuiteLock } from "../../../.github/extensions/office-canvas/test/integration/word-suite-lock.mjs";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SELF = fileURLToPath(import.meta.url);
const EXT = path.join(HERE, "..", "..", "..", ".github", "extensions", "office-canvas");
const FIXTURE = path.join(EXT, "test", "integration", "make-fixture.ps1");
const SUITE = path.join(EXT, "test", "integration", "read-smoke.mjs");

const argv = process.argv.slice(2);
const flag = (name, fallback) => {
    const i = argv.indexOf(`--${name}`);
    return i === -1 ? fallback : argv[i + 1];
};

// ------------------------------------------------------------ neighbour mode
// A separate process on purpose: the contention being modelled is between
// operating-system processes, and an in-process imitation would not produce a
// WINWORD another run's census can see.
if (argv[0] === "--be-neighbour") {
    const locked = argv[1] === "locked";
    const work = await mkdtemp(path.join(tmpdir(), "neighbour-"));
    let pass = 0;
    // Stop on a byte from stdin rather than on a signal. On Windows a signal
    // sent by the parent is TerminateProcess: no handler runs, and a
    // make-fixture pass caught mid-flight is orphaned along with the Word it
    // drove. That orphan is exactly the uncontrolled variable this probe
    // exists to remove, so the neighbour is asked to stop and finishes the
    // pass it is in.
    let stopping = false;
    process.stdin.on("data", () => (stopping = true));
    process.stdin.on("end", () => (stopping = true));
    process.stdin.resume();
    for (;;) {
        if (stopping) break;
        const handle = locked
            ? await acquireWordSuiteLock("neighbour", { pollMs: 250, log: () => {} })
            : null;
        try {
            await new Promise((resolve) => {
                const child = spawn(
                    "powershell.exe",
                    [
                        "-NoProfile",
                        "-NonInteractive",
                        "-ExecutionPolicy",
                        "Bypass",
                        "-File",
                        FIXTURE,
                        "-Out",
                        path.join(work, `n${pass % 3}.docx`),
                        "-Chapters",
                        "1",
                    ],
                    { stdio: "ignore" },
                );
                child.on("exit", resolve);
                child.on("error", resolve);
            });
        } finally {
            handle?.release();
        }
        process.stdout.write(`neighbour pass ${++pass}\n`);
        await new Promise((r) => setTimeout(r, 500));
    }
    // Explicit: the loop is no longer infinite, and falling through from here
    // would run the measuring half of this file inside the neighbour.
    process.exit(0);
}

// ------------------------------------------------------------------ measuring
const runs = Number(argv[0] ?? 18);
const neighbour = flag("neighbour", "unlocked");

function runSuite() {
    return new Promise((resolve) => {
        const started = Date.now();
        const child = spawn(process.execPath, [SUITE], { stdio: ["ignore", "ignore", "pipe"] });
        let err = "";
        child.stderr.on("data", (d) => (err += d));
        child.on("exit", (code) => {
            // Key on the check's name, not on a glyph or a line position: the
            // suite prints "  FAIL <name>" and the two candidate failures are
            // told apart by which name follows.
            const failed = [...err.matchAll(/^\s*FAIL\s+(.+)$/gm)].map((m) => m[1].trim());
            resolve({
                code,
                failed,
                seconds: (Date.now() - started) / 1000,
                startedAt: new Date(started).toISOString(),
            });
        });
    });
}

let neighbourProc = null;
if (neighbour !== "none") {
    neighbourProc = spawn(process.execPath, [SELF, "--be-neighbour", neighbour], {
        stdio: ["pipe", "ignore", "ignore"],
    });
    // Let it get a Word going before the first run, so run 1 is not the only
    // one measured on a quiet machine.
    await new Promise((r) => setTimeout(r, 15_000));
}

console.log(`read-smoke x ${runs}, neighbour: ${neighbour}\n`);
const results = [];
for (let i = 1; i <= runs; i++) {
    const r = await runSuite();
    results.push(r);
    const why = r.failed.length ? `  <- ${r.failed.join(" | ")}` : "";
    // Stamp every run. A contention figure whose runs cannot be dated cannot
    // afterwards be checked against what else was on the machine, and "I cannot
    // tell whether a sibling was running" is not a recoverable answer later.
    console.log(
        `run ${String(i).padStart(2)}: ${r.startedAt}  exit ${r.code}  ${r.seconds.toFixed(0)}s${why}`,
    );
}

if (neighbourProc) {
    // Ask, then wait. Killing it outright would strand a make-fixture pass and
    // its Word; the wait is what makes the shutdown clean rather than merely
    // quiet. The kill is a backstop for a neighbour that is somehow stuck, and
    // it says so out loud, because a probe that leaks silently is worse than
    // one that leaks.
    const stopped = new Promise((resolve) => neighbourProc.on("exit", resolve));
    neighbourProc.stdin.write("stop\n");
    neighbourProc.stdin.end();
    const clean = await Promise.race([
        stopped.then(() => true),
        new Promise((r) => setTimeout(() => r(false), 90_000)),
    ]);
    if (!clean) {
        console.log("WARNING: neighbour did not stop in 90s; killing it. Check for a stray WINWORD.");
        neighbourProc.kill();
    }
}

const green = results.filter((r) => r.code === 0).length;
const byCheck = new Map();
for (const r of results) for (const name of r.failed) byCheck.set(name, (byCheck.get(name) ?? 0) + 1);

console.log(`\npass rate: ${green}/${runs}   neighbour: ${neighbour}`);
for (const [name, n] of byCheck) console.log(`  failed ${n}x: ${name}`);
await rm(path.join(tmpdir(), "never"), { force: true }).catch(() => {});
process.exit(0);
