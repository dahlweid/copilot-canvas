// When two Words are alive at once, whose autocorrect settings win?
//
// Run: node spikes/isolation/probes/probe-autocorrect-concurrency.mjs
// Requires Word. Perturbs the user's Word settings and puts them back; the found
// state is captured first and restored in a `finally`.
//
// THE QUESTION (issue #51)
//
// The five settings word-host.ps1 used to suppress around an authoring call --
// AutoCorrect.{ReplaceText,CorrectSentenceCaps,CorrectInitialCaps} and
// Options.AutoFormatAsYouType{ReplaceQuotes,ReplaceSymbols} -- persist for the
// user. That was established by probe-autocorrect-untouched.mjs and by the
// retraction of arm C in probe-autocorrect.ps1: a reader started *after* the
// writer exits sees the new value, a reader running *alongside* it sees the old
// one.
//
// Two observations, no mechanism. What is left open is the one that decides
// whether "capture, suppress, restore" is a guarantee or only a best effort:
// **when two instances are alive at once, does the one that exits last overwrite
// what the other wrote?** If it does, then a Word the user already had open
// carries a copy of these settings from before we ran, and its exit lands that
// copy over our restore -- or over our suppression, depending on the order.
//
// Both orderings are ordinary. A user with Word open, asking Copilot to author a
// document, produces one of them by default.
//
// The answer, measured below, is yes: whichever Word exits cleanly last decides
// the stored value (arms 4 and 5). That is what made capture-and-restore a best
// effort rather than a guarantee, and it is half of why the suppression was
// deleted instead of hardened -- the other half being that it was measured to
// buy nothing (probe-autocorrect-necessity.ps1).
//
// This probe is KEPT after that deletion, and deliberately so. It is the answer
// to the issue's open question, it prices any future proposal to write a global
// Word setting from this host, and arm 6 is now an end-to-end regression guard:
// it runs the real tool against a user's Word holding all five ON and asserts
// they are still ON afterwards.
//
// WHAT THIS MEASURES, AND THE DISCRIMINATOR IT IS BUILT AROUND
//
// Every persisted value is read from a **third, fresh instance started after
// every other Word this probe created has exited**. A read taken while any of
// them is alive measures nothing about persistence -- that is precisely the
// defect that made arm C wrong, and it is avoided here by construction rather
// than by care: `readPersisted` waits for every worker pid to leave the process
// list before it starts its reader.
//
// The reader is itself a Word, so it is also a potential writer at its own exit.
// It only ever reads, so whatever it flushes is what it found, and every reading
// below is taken by exactly one such reader with nothing else of ours alive.
//
// The "user's Word" role is a **second real powershell.exe**, not a Start-Job:
// a job runspace cannot drive Word (SaveAs2 wedges there -- probe-saveas-apartment.ps1),
// and the whole point is a genuinely separate process holding a genuinely
// separate instance.
//
// THE ARMS
//
//   1  writer quits             control: does a lone writer's value persist
//   2  writer's host killed,    the crash path word-host.mjs actually takes --
//      Word left to self-exit   it kills the host process after 20 s, and a
//                               hidden Word with no documents exits when the
//                               last COM reference goes
//   3  writer's Word killed     the orphan reaper's path (Stop-VerifiedWord)
//   4  user's Word exits AFTER  the ordering that could undo our restore
//      the writer
//   5  user's Word exits BEFORE the ordering that could land our suppression on
//      the writer                the user
//   6  the real tool, with a    the field scenario, end to end
//      user's Word alive across
//      it and exiting after
//
// Arms 2 and 3 are what deliverable "the restore must survive a host crash and
// the kill fallback" turns on: the restore in Cmd-Create runs in a `finally`
// while Word is still held, so a crash *during* the authoring call skips it.
// Whether that leaks anything to the user depends entirely on whether a Word
// that is killed, or that exits on a refcount drop, persists what it was holding.
// That is a measurement, not a deduction.
//
// SAFETY
//
// Every Word this probe kills was created by a worker *this probe started*, and
// the pid comes out of that worker's own RCW via ActiveWindow.Hwnd ->
// GetWindowThreadProcessId -- ours by construction, not by differencing
// (probe-word-ownership.ps1 measured differencing reporting 2 new pids for 1
// instance). No pid that existed before this probe started is ever touched.
//
// Nothing is interpolated into a command string: the worker is a script file and
// every path reaches it as a discrete argv element.

import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const here = path.dirname(fileURLToPath(import.meta.url));
const extension = path.resolve(here, "../../../.github/extensions/office-canvas");
// Imported, never copied: the census helper is shared with the integration
// suites and a second copy would drift from the one that gets fixed.
const { wordPids } = await import(pathToFileURL(path.join(extension, "test/integration/word-pids.mjs")).href);
// pathToFileURL, not the bare path: on Windows a dynamic import of `C:\...`
// fails with ERR_UNSUPPORTED_ESM_URL_SCHEME because the drive letter parses as
// a URL scheme.
const { RenderCache } = await import(pathToFileURL(path.join(extension, "src/render-cache.mjs")).href);

// The five, listed once, for the same reason $script:AC_SETTINGS exists in
// word-host.ps1: a setting added to one pass and not another reads as handled
// while being unchecked.
const NAMES = [
    "ReplaceText",
    "CorrectSentenceCaps",
    "CorrectInitialCaps",
    "AutoFormatAsYouTypeReplaceQuotes",
    "AutoFormatAsYouTypeReplaceSymbols",
];

// Generous, and deliberately so. Word's start and exit latencies are
// load-dependent and this probe runs three instances at once in places; a
// deadline fitted to an idle machine turns a slow run into a false finding.
const DEADLINE_MS = 180000;
const POLL_MS = 200;
// How long an orphaned Word is given to exit by itself before the arm concludes
// it will not. Measured once at 180 s with no exit, so this only has to be long
// enough to be fair, not long enough to be conclusive twice.
const SELF_EXIT_MS = 45000;

const WORKER = `param(
    [Parameter(Mandatory = $true)][string] $ResultPath,
    [Parameter(Mandatory = $true)][string] $ReadyPath,
    [Parameter(Mandatory = $true)][string] $ReleasePath,
    # Five booleans, comma separated, in the order this script reads and writes
    # them. Empty means read only. A per-setting list rather than one flag,
    # because the found state on a user's machine need not be uniform and
    # putting back "whatever the first one was" would be a silent corruption
    # dressed as a restore.
    [string] $Values = '',
    # Written after the release signal and before Quit, in the same five-value
    # shape. This is the restore half of what Cmd-Create does: suppress, hold,
    # put the priors back, quit. Empty means no second write.
    [string] $RestoreValues = ''
)

$ErrorActionPreference = 'Stop'

# The pid comes from a window handle read off the RCW this script created, so it
# names this script's own instance by construction. Differencing the process
# list would not: with a concurrent Word alive -- which is the entire point of
# this probe -- it reports pids that are not ours.
Add-Type -Namespace Win32 -Name Wnd -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true)]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
'@

function Get-Five($a) {
    [ordered]@{
        ReplaceText                       = $a.AutoCorrect.ReplaceText
        CorrectSentenceCaps               = $a.AutoCorrect.CorrectSentenceCaps
        CorrectInitialCaps                = $a.AutoCorrect.CorrectInitialCaps
        AutoFormatAsYouTypeReplaceQuotes  = $a.Options.AutoFormatAsYouTypeReplaceQuotes
        AutoFormatAsYouTypeReplaceSymbols = $a.Options.AutoFormatAsYouTypeReplaceSymbols
    }
}

function Set-Five($a, [bool[]] $v) {
    $a.AutoCorrect.ReplaceText = $v[0]
    $a.AutoCorrect.CorrectSentenceCaps = $v[1]
    $a.AutoCorrect.CorrectInitialCaps = $v[2]
    $a.Options.AutoFormatAsYouTypeReplaceQuotes = $v[3]
    $a.Options.AutoFormatAsYouTypeReplaceSymbols = $v[4]
}

$w = $null
try {
    $w = New-Object -ComObject Word.Application
    $w.Visible = $false
    $w.DisplayAlerts = 0

    # Application.Hwnd does not exist on Word; ActiveWindow.Hwnd does, once a
    # document is open. Close(0) takes a by-value VARIANT and was measured not to
    # throw under 5.1 -- unlike Quit, which is why that one below has no argument.
    $scratch = $w.Documents.Add()
    [uint32] $procId = 0
    [void][Win32.Wnd]::GetWindowThreadProcessId([IntPtr][int64] $w.ActiveWindow.Hwnd, [ref] $procId)
    $scratch.Close(0)

    if ($Values -ne '') {
        $wanted = @($Values.Split(',') | ForEach-Object { $_.Trim() -eq 'True' })
        if ($wanted.Count -ne 5) { throw "expected five values, got $($wanted.Count)" }
        Set-Five $w $wanted
    }

    $lines = @("pid=$procId")
    foreach ($e in (Get-Five $w).GetEnumerator()) { $lines += "$($e.Key)=$($e.Value)" }
    Set-Content -LiteralPath $ResultPath -Value $lines -Encoding ascii
    Set-Content -LiteralPath $ReadyPath -Value 'ready' -Encoding ascii

    # Held open until the parent says otherwise. This is what makes the exit
    # ordering something the parent chooses rather than something it races.
    #
    # The work directory going away is the parent giving up on us -- it deletes
    # that directory last. Treated as a release rather than ignored, so a worker
    # whose parent has stopped waiting quits its Word through the finally below
    # instead of holding one forever. Measured: a run where the parent timed out
    # on a slow Word start left exactly this worker behind.
    $workDir = Split-Path -Parent $ReleasePath
    while (-not (Test-Path -LiteralPath $ReleasePath)) {
        if (-not (Test-Path -LiteralPath $workDir)) { break }
        Start-Sleep -Milliseconds 50
    }

    if ($RestoreValues -ne '') {
        $back = @($RestoreValues.Split(',') | ForEach-Object { $_.Trim() -eq 'True' })
        if ($back.Count -ne 5) { throw "expected five restore values, got $($back.Count)" }
        Set-Five $w $back
        # Read back rather than assumed: a restore that did not land looks exactly
        # like a restore that did, from the writing side.
        $lines = @()
        foreach ($e in (Get-Five $w).GetEnumerator()) { $lines += "restored_$($e.Key)=$($e.Value)" }
        Add-Content -LiteralPath $ResultPath -Value $lines
    }
} catch {
    Add-Content -LiteralPath $ResultPath -Value "error=$($_.Exception.Message.Split([char]10)[0])"
    Set-Content -LiteralPath $ReadyPath -Value 'error' -Encoding ascii
} finally {
    if ($null -ne $w) {
        # Quit(), never Quit(<arg>): under Windows PowerShell 5.1 the argument
        # form does not bind, throws, and leaves a WINWORD that process exit does
        # not reap (probe-quit0-leak.ps1). Reported rather than swallowed --
        # a swallowed quit failure is how this repo accumulated 14 stranded Words.
        try { $w.Quit() } catch { Add-Content -LiteralPath $ResultPath -Value "quitfailed=$($_.Exception.Message.Split([char]10)[0])" }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch { }
    }
}
`;

const workRoot = await mkdtemp(path.join(tmpdir(), "word-ac-conc-"));
const workerPath = path.join(workRoot, "word-worker.ps1");

let failures = 0;
let seq = 0;
/** Every Word pid this probe's workers created, so the finally can reap safely. */
const created = new Set();

const say = (line) => process.stdout.write(`${line}\n`);
const step = (label, ok, detail) => {
    if (!ok) failures += 1;
    say(`${ok ? "ok  " : "FAIL"}  ${label}${detail ? ` -- ${detail}` : ""}`);
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** Poll `probe` until it is true, or fail loudly. Never silently gives up. */
async function until(what, probe) {
    const deadline = Date.now() + DEADLINE_MS;
    for (;;) {
        if (await probe()) return;
        if (Date.now() > deadline) throw new Error(`timed out after ${DEADLINE_MS} ms waiting for ${what}`);
        await sleep(POLL_MS);
    }
}

const exists = async (file) =>
    await readFile(file, "utf8").then(
        () => true,
        () => false,
    );

/** Kills one pid, by number, name-checked. Only ever called on a worker's own Word. */
async function killPid(pid) {
    await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-Command",
        // Name-checked because pids are reused, and this probe runs alongside
        // whatever else the machine is doing.
        `$p = Get-Process -Id ${pid} -ErrorAction SilentlyContinue; ` +
            "if ($p -and ($p.ProcessName -eq 'WINWORD' -or $p.ProcessName -eq 'powershell')) { $p.Kill() }",
    ]).catch(() => {});
}

/**
 * Starts one Word in its own powershell.exe and waits until it is holding.
 *
 * `values` is null for a read-only instance, or the five booleans to write, in
 * NAMES order.
 *
 * Returns the worker's own reading of the five, the pid of the Word it created,
 * and the three ways to end it: `release()` for a clean Quit, `killShell()` and
 * `killWord()` for the two crash shapes.
 */
async function startWorker(values = null, restoreValues = null) {
    const n = (seq += 1);
    const resultPath = path.join(workRoot, `w${n}.txt`);
    const readyPath = path.join(workRoot, `w${n}.ready`);
    const releasePath = path.join(workRoot, `w${n}.release`);
    const encode = (v) => (v ? NAMES.map((name) => (v[name] ? "True" : "False")).join(",") : "");
    const written = encode(values);

    // Discrete argv, no command string: a work directory under
    // C:\Users\O'Brien\ would otherwise reach a PowerShell parser.
    const child = execFile("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        workerPath,
        "-ResultPath",
        resultPath,
        "-ReadyPath",
        readyPath,
        "-ReleasePath",
        releasePath,
        "-Values",
        written,
        ...(restoreValues ? ["-RestoreValues", encode(restoreValues)] : []),
    ]);
    // Without a listener a spawn failure raises an unhandled 'error' event and
    // takes the probe down with a stack that says nothing about which arm.
    child.on("error", (err) => say(`      worker ${n} spawn error: ${err.message}`));

    await until(`worker ${n} (${written || "read-only"}) to be holding a Word`, () => exists(readyPath));

    const text = await readFile(resultPath, "utf8");
    const observed = {};
    let wordPid = 0;
    for (const line of text.trim().split(/\r?\n/)) {
        const [key, value] = line.split("=");
        if (key === "pid") wordPid = Number(value);
        else if (NAMES.includes(key)) observed[key] = value.trim() === "True";
        else if (key === "error") throw new Error(`worker ${n} failed: ${value}`);
    }
    if (!wordPid) throw new Error(`worker ${n} did not report a Word pid: ${text}`);
    created.add(wordPid);

    return {
        n,
        wordPid,
        shellPid: child.pid,
        values: observed,
        async release() {
            await writeFile(releasePath, "go");
            await until(`worker ${n}'s Word (pid ${wordPid}) to exit`, async () => !(await wordPids()).includes(wordPid));
            // Whatever the worker read back after its own restore write. Empty
            // when it had no restore to do.
            const tail = await readFile(resultPath, "utf8").catch(() => "");
            const readBack = {};
            for (const line of tail.trim().split(/\r?\n/)) {
                const [key, value] = line.split("=");
                if (key?.startsWith("restored_")) readBack[key.slice("restored_".length)] = value.trim() === "True";
            }
            return readBack;
        },
        async killShell() {
            await killPid(child.pid);
        },
        async killWord() {
            await killPid(wordPid);
        },
        async waitWordGone() {
            await until(`worker ${n}'s Word (pid ${wordPid}) to leave the process list`, async () =>
                !(await wordPids()).includes(wordPid),
            );
        },
        /** Bounded, and returns the answer rather than throwing: "it did not" is data. */
        async wordGoneWithin(ms) {
            const deadline = Date.now() + ms;
            for (;;) {
                if (!(await wordPids()).includes(wordPid)) return true;
                if (Date.now() > deadline) return false;
                await sleep(POLL_MS);
            }
        },
    };
}

/**
 * Reads the five from a fresh Word, started only once every pid in `mustBeGone`
 * has left the process list.
 *
 * The wait is the discriminator. A reading taken while any writer is alive is
 * the measurement arm C got wrong, and no amount of care elsewhere recovers it.
 */
async function readPersisted(mustBeGone = []) {
    for (const pid of mustBeGone) {
        await until(`pid ${pid} to exit before the persisted value is read`, async () => !(await wordPids()).includes(pid));
    }
    const reader = await startWorker();
    const values = reader.values;
    await reader.release();
    return values;
}

const all = (values, expected) => NAMES.every((n) => values[n] === expected);
const show = (values) =>
    all(values, true) ? "all five ON" : all(values, false) ? "all five OFF" : NAMES.map((n) => `${n}=${values[n]}`).join(" ");

/** The five, all set to one value, in the shape startWorker wants. */
const uniform = (value) => Object.fromEntries(NAMES.map((n) => [n, value]));

/** Puts the store into a known state and confirms it from a fresh instance. */
async function seed(value) {
    const w = await startWorker(uniform(value));
    await w.release();
    const back = await readPersisted([w.wordPid]);
    if (!all(back, value)) throw new Error(`could not seed the store to ${value}: a fresh instance reads ${show(back)}`);
}

const findings = [];
/**
 * Records one arm's answer and prints it beside the arm, so a reading is legible
 * in the scroll as well as in the summary at the end.
 *
 * `did` is what the arm actually did, spelled per arm rather than assumed: arm 7's
 * writer restores before it quits, and a fixed "wrote OFF" line would misdescribe it.
 */
const record = (arm, question, did, reading) => {
    findings.push({ arm, question, reading });
    say(`  arm ${arm}: ${question}\n      ${did} -> a fresh instance reads ${reading}`);
};

const pidsBefore = await wordPids();
say(`WINWORD before: [${pidsBefore.join(", ")}]`);

let found = null;
try {
    await writeFile(workerPath, WORKER, "utf8");

    const first = await readPersisted();
    found = first;
    say(`found state: ${show(found)}\n`);

    // --- 1. control: a lone writer, quit cleanly ----------------------------
    say("arm 1 -- a lone writer that quits");
    await seed(true);
    {
        const w = await startWorker(uniform(false));
        await w.release();
        const after = await readPersisted([w.wordPid]);
        step("a lone writer's value persists once it has exited", all(after, false), show(after));
        record(1, "does a lone writer's value persist?", "seeded ON, writer wrote OFF and quit", show(after));
    }

    // --- 2. the writer's host is killed ------------------------------------
    //
    // This is the shape word-host.mjs produces: the client kills the host
    // process when the quit RPC misses its 20 s deadline. The arm originally
    // assumed the Word would then exit by itself on the refcount drop -- which
    // is what word-host.ps1's quit site describes, for a host that exits
    // NORMALLY after releasing its reference. The first run of this probe
    // measured that a KILLED host is not that case: pid 39280 was still in the
    // process list 180 s after its host died, and had to be terminated by hand.
    // So the arm measures both halves: what a fresh reader sees while that
    // orphan is still alive, and what it sees once the reaper has killed it.
    say("\narm 2 -- the writer's host is killed and its Word is left orphaned");
    await seed(true);
    {
        const w = await startWorker(uniform(false));
        await w.killShell();
        const selfExited = await w.wordGoneWithin(SELF_EXIT_MS);
        step(
            "measured whether an orphaned Word exits on its own",
            true,
            selfExited
                ? `it exited within ${SELF_EXIT_MS} ms`
                : `it was STILL ALIVE after ${SELF_EXIT_MS} ms -- the orphan ledger is load-bearing`,
        );

        if (!selfExited) {
            const whileAlive = await readPersisted();
            step(
                "measured what a reader sees while the orphan is still alive",
                true,
                all(whileAlive, true)
                    ? "ON -- the reader sees the pre-write value, as the concurrent-reader case predicts"
                    : `${show(whileAlive)}`,
            );
            record(
                "2a",
                "with the host killed and its Word still alive, what does a fresh reader see?",
                "seeded ON, writer wrote OFF, host killed, its Word still up",
                show(whileAlive),
            );
            await w.killWord();
        }

        await w.waitWordGone();
        const after = await readPersisted([w.wordPid]);
        step(
            "measured what a killed host leaves behind once its Word is gone",
            true,
            all(after, false) ? "OFF -- the suppression PERSISTS through a host crash" : `${show(after)} -- not persisted`,
        );
        record("2b", "once that orphan is gone, does the suppression it wrote persist?", "seeded ON, writer wrote OFF, then the orphan was killed", show(after));
    }

    // --- 3. the writer's Word is killed outright ----------------------------
    //
    // The orphan reaper's path: Clear-OrphanedWord finds a pid recorded by a
    // host that is gone and terminates it through Stop-VerifiedWord. The Word is
    // killed first here so the kill provably lands on a live instance rather
    // than racing the self-exit measured in arm 2.
    say("\narm 3 -- the writer's Word is killed outright");
    await seed(true);
    {
        const w = await startWorker(uniform(false));
        await w.killWord();
        await w.killShell();
        await w.waitWordGone();
        const after = await readPersisted([w.wordPid]);
        step(
            "measured what a killed Word leaves behind",
            true,
            all(after, false) ? "OFF -- the suppression PERSISTS through a kill" : `${show(after)} -- not persisted`,
        );
        record(3, "does a killed Word still persist what it wrote?", "seeded ON, writer wrote OFF, then was killed", show(after));
    }

    // --- 4. the user's Word exits AFTER ours --------------------------------
    say("\narm 4 -- a user's Word is alive across the write and exits after it");
    await seed(true);
    {
        const user = await startWorker();
        step("the user's Word read the seeded state", all(user.values, true), show(user.values));

        const w = await startWorker(uniform(false));
        await w.release();

        await user.release();
        const after = await readPersisted([w.wordPid, user.wordPid]);
        step(
            "measured which of the two exits decided the stored value",
            true,
            all(after, false)
                ? "OFF -- the writer's value survived the user's Word exiting after it"
                : `${show(after)} -- the user's Word OVERWROTE the writer on its way out`,
        );
        record(4, "user's Word exits after the writer -- who wins?", "seeded ON, user's Word up first, writer wrote OFF and quit, user quit last", show(after));
    }

    // --- 5. the user's Word exits BEFORE ours -------------------------------
    say("\narm 5 -- a user's Word is alive across the write and exits before it");
    await seed(true);
    {
        const user = await startWorker();
        step("the user's Word read the seeded state", all(user.values, true), show(user.values));

        const w = await startWorker(uniform(false));
        await user.release();
        await w.release();

        const after = await readPersisted([w.wordPid, user.wordPid]);
        step(
            "measured which of the two exits decided the stored value",
            true,
            all(after, false)
                ? "OFF -- the writer's value is what the user's next Word will see"
                : `${show(after)} -- the writer's value did not land`,
        );
        record(5, "writer exits after the user's Word -- who wins?", "seeded ON, user's Word up first and quit first, writer wrote OFF and quit last", show(after));
    }

    // --- 6. the real tool, under the same concurrency -----------------------
    //
    // Arms 1-5 use a bare writer, which is the primitive. This is the product:
    // create_document captures, suppresses, authors and restores inside one
    // call, with a user's Word alive across the whole of it and exiting after.
    say("\narm 6 -- create_document runs with a user's Word alive, which exits after it");
    await seed(true);
    {
        const user = await startWorker();
        step("the user's Word read the ON state we are protecting", all(user.values, true), show(user.values));

        const cache = new RenderCache({ cacheRoot: path.join(workRoot, "artifacts"), log: () => {} });
        let result;
        let ourPid = null;
        // WordHost calls this.onOwnedPid(pid), so replacing the property after
        // construction works and gets us the pid by attribution rather than by
        // differencing the process list.
        cache.host.onOwnedPid = (pid) => {
            ourPid = pid;
            created.add(pid);
        };
        try {
            result = await cache.createDocument(path.join(workRoot, "probe.docx"), {
                blocks: [{ kind: "paragraph", text: 'It said "quote" -- and it stayed.' }],
            });
        } finally {
            await cache.dispose?.().catch(() => {});
            // The product's own fallback, exercised here rather than trusted:
            // a no-op when the quit landed, and the only thing that ends Word
            // when it did not.
            cache.reapOwnedWord?.();
        }

        step(
            "the tool reports nothing about autocorrect",
            result.autoCorrect === undefined,
            JSON.stringify(result.autoCorrect ?? null),
        );

        // Our Word must be gone before the user's exits, or the ordering this
        // arm is named for is not the ordering that happened.
        if (ourPid) {
            await until(`our authoring Word (pid ${ourPid}) to exit`, async () => !(await wordPids()).includes(ourPid));
        }
        await user.release();

        const after = await readPersisted([user.wordPid, ...(ourPid ? [ourPid] : [])]);
        step("the user's next Word still reads all five ON", all(after, true), show(after));
        record(6, "the real tool, user's Word exiting after it", "seeded ON, user's Word up first, create_document ran, user quit last", show(after));
    }

    // --- 7. the ordering a correct save/restore can still lose to -----------
    //
    // Arms 4 and 5 put the user's Word in place BEFORE the suppression, so its
    // in-memory copy is the ON state we are protecting and its late exit is
    // harmless -- it writes back what we wanted anyway. Arm 6 is that same
    // benign ordering with the real tool.
    //
    // The ordering that looked dangerous is a user's Word that starts INSIDE
    // the suppression window. The hypothesis was: it reads the suppressed OFF,
    // our restore then puts ON back, and since the last exit decides the stored
    // value (arm 4) the user's Word writes our suppression back out after we
    // have already undone it.
    //
    // MEASURED, AND THE HYPOTHESIS IS FALSE. A Word started while another
    // instance is holding the suppressed values reads all five ON -- the
    // pre-write, persisted value. The live suppression in our instance is not
    // visible to it, which is the same observation as arm 2a from the other
    // side. So the value a user's Word carries to its own exit is the value we
    // are protecting, and there is no ordering in which it flushes our
    // suppression back out.
    //
    // Both halves are reported as measurements rather than asserted: this arm
    // was first written asserting the mid-suppression reader would see OFF, and
    // it went red against a correct product.
    say("\narm 7 -- a user's Word starts INSIDE the suppression window and exits after the restore");
    await seed(true);
    {
        const w = await startWorker(uniform(false), uniform(true));

        // Started while the suppression is live: a user opening Word during a
        // create_document call.
        const user = await startWorker();
        step(
            "measured what a user's Word started mid-suppression reads",
            true,
            all(user.values, false)
                ? "OFF -- it picked up the live suppression, so it can carry it to its own exit"
                : `${show(user.values)} -- the live suppression is NOT visible to it; it reads the persisted value`,
        );

        const readBack = await w.release();
        step(
            "the writer's own restore landed on its instance before it quit",
            NAMES.every((n) => readBack[n] === true),
            JSON.stringify(readBack),
        );

        await user.release();
        const after = await readPersisted([w.wordPid, user.wordPid]);
        step(
            "measured whether a correct restore survives that user's Word exiting last",
            true,
            all(after, true)
                ? "ON -- the restore survived"
                : `${show(after)} -- the restore was LOST: the user's Word wrote the suppressed values back out`,
        );
        record(7, "a user's Word opened mid-suppression, exiting after our restore -- does the restore survive?", "seeded ON, writer wrote OFF, user's Word opened, writer restored ON and quit, user quit last", show(after));
    }
} catch (err) {
    failures += 1;
    say(`FAIL  the probe raised: ${err.stack ?? err.message}`);
} finally {
    // Reap AFTER the restore, not before. An earlier version reaped first, and a
    // run where the restore's own worker timed out on a slow Word start left that
    // Word behind with nothing left to collect it -- the reap had already
    // happened. Anything the restore creates is in `created` too, so doing it in
    // this order covers both. Only pids this probe's own workers created are
    // touched, and only if they were not alive beforehand.
    if (found) {
        // Retried once. The failure that made this necessary was a Word taking
        // longer than the deadline to start on a machine this probe had already
        // put ~35 instances through -- a slow start, not a broken one, and
        // leaving the user's settings changed because of it is the exact outcome
        // this whole probe exists to prevent.
        for (let attempt = 1; attempt <= 2; attempt += 1) {
            try {
                const w = await startWorker(found);
                await w.release();
                const back = await readPersisted([w.wordPid]);
                const ok = NAMES.every((n) => back[n] === found[n]);
                step("the probe put the machine back the way it found it", ok, show(back));
                break;
            } catch (err) {
                if (attempt === 1) {
                    say(`      restore attempt 1 failed (${err.message}); retrying`);
                    continue;
                }
                failures += 1;
                say(`FAIL  could not restore the found state: ${err.message}`);
                say(`      it was ${show(found)} -- put it back by hand`);
            }
        }
    }

    for (const pid of created) {
        if (pidsBefore.includes(pid)) continue;
        if (!(await wordPids()).includes(pid)) continue;
        say(`      reaping this probe's own Word, pid ${pid}`);
        await killPid(pid);
    }

    // Deleted last, and deliberately: a worker still holding a Word reads this
    // directory's disappearance as its release signal and quits cleanly.
    await rm(workRoot, { recursive: true, force: true }).catch(() => {});
}

// The census is differenced, never "no WINWORD exists": other sessions and the
// developer's own Word are routinely alive here.
const pidsAfter = await wordPids();
const leaked = pidsAfter.filter((pid) => !pidsBefore.includes(pid));
say(`\nWINWORD after: [${pidsAfter.join(", ")}]`);
step("no WINWORD was left behind", leaked.length === 0, leaked.length ? `leaked ${leaked.join(", ")}` : "");

say("\nsummary");
for (const f of findings) say(`  arm ${f.arm}  ${f.question}\n           -> ${f.reading}`);

say(failures === 0 ? "\nall checks passed" : `\n${failures} check(s) failed`);
process.exitCode = failures === 0 ? 0 : 1;
