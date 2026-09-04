// Does the *fixture generator's own* Word still exist when the "a missing file
// is reported without starting Word" check in read-smoke.mjs runs?
//
// Issue #37 attributes that check's 17-in-18 pass rate to another session
// driving Word. That is an inference, and there is a nearer candidate nobody
// measured: read-smoke takes its `pidsBefore` census first, and only *then*
// calls `make-fixture.ps1` through its `makeFixture` helper, which does `New-Object -ComObject
// Word.Application` and `$w.Quit()`. So the fixture's own WINWORD is created
// **inside the differencing window**, and `Quit()` returns 3-28 ms before the
// process actually goes (probe-quit-exit-gap.ps1; measured tail 2.7-6.1 s on an
// idle machine). If that tail outlives the PowerShell host, the fixture's dying
// Word is a "new pid" to the very next census -- and the next census is a
// doomed read away.
//
// This reproduces exactly that prefix and nothing else: census, generate the
// fixture, then census again as fast as read-smoke would. It reports what
// read-smoke's missing-file check would have concluded, then polls until the pid is gone so the
// tail past `execFile` resolving is a number rather than a guess.
//
// It kills nothing. Every pid it names is one make-fixture.ps1 created and
// asked to quit; the point is how long that takes, not to hurry it.
//
// Run:
//   node spikes/isolation/probes/probe-fixture-word-tail.mjs [reps]

import { execFile } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const FIXTURE_SCRIPT = path.join(
    HERE,
    "..",
    "..",
    "..",
    ".github",
    "extensions",
    "office-canvas",
    "test",
    "integration",
    "make-fixture.ps1",
);

async function wordPids() {
    const { stdout } = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        "@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -join ','",
    ]);
    return stdout.trim().split(",").filter(Boolean).map(Number);
}

const reps = Number(process.argv[2] ?? 5);
const workRoot = await mkdtemp(path.join(tmpdir(), "probe-fixture-tail-"));
const rows = [];

console.log(`make-fixture.ps1 -> WINWORD tail, ${reps} rep(s)`);
console.log("(the census is taken before the fixture runs, as read-smoke.mjs does)\n");

for (let rep = 1; rep <= reps; rep++) {
    const out = path.join(workRoot, `fixture-${rep}.docx`);
    const before = await wordPids();

    const startedAt = Date.now();
    await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        FIXTURE_SCRIPT,
        "-Out",
        out,
        "-Chapters",
        "2",
        "-Duplicates",
    ]);
    const returnedAt = Date.now();

    // The observation read-smoke:157 makes, at the same distance from the
    // fixture returning: one census, no settle, no retry.
    const appeared = (await wordPids()).filter((pid) => !before.includes(pid));
    const observedAt = Date.now();

    let goneAt = null;
    if (appeared.length > 0) {
        const deadline = Date.now() + 60_000;
        while (Date.now() < deadline) {
            const still = (await wordPids()).filter((pid) => appeared.includes(pid));
            if (still.length === 0) {
                goneAt = Date.now();
                break;
            }
            await new Promise((r) => setTimeout(r, 100));
        }
    }

    const verdict = appeared.length === 0 ? "would PASS" : "would FAIL";
    rows.push({ rep, appeared, verdict });
    console.log(
        `rep ${rep}: fixture took ${returnedAt - startedAt} ms, census ${observedAt - returnedAt} ms later ` +
            `-> ${appeared.length} new pid(s) [${appeared.join(",")}]  ${verdict}`,
    );
    if (appeared.length > 0) {
        const tail = goneAt === null ? "still alive after 60 s" : `${goneAt - returnedAt} ms past execFile resolving`;
        console.log(`        survivor gone: ${tail}`);
    }
    await rm(out, { force: true }).catch(() => {});
}

await rm(workRoot, { recursive: true, force: true }).catch(() => {});

const failed = rows.filter((r) => r.verdict === "would FAIL").length;
console.log(`\nread-smoke:157 would have failed ${failed}/${reps} time(s) on the fixture's own Word alone.`);
