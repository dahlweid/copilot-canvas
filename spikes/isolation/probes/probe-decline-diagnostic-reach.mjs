// Does a decline diagnostic written by word-host.ps1 actually reach a reader?
//
// PR #36 added Write-HostDiagnostic so that a refusal to kill -- which means a
// leaked Word, a pid collision, or both -- is reported rather than swallowed.
// That is only a fix if something reads the channel it writes to. word-host.mjs
// buffers child stderr into a private ring and the buffer has exactly one
// reader, inside #onExit, where it decorates the rejection handed to calls that
// were still in flight. A clean dispose has none, and the orphan sweep runs at
// *startup*, so its diagnostics would sit in the buffer for the whole session
// and be dropped at the end of it.
//
// Reading the code says that. This measures it, with a control, because the
// static read cannot rule out Node surfacing child stderr some other way.
//
//   arm A (control): spawn word-host.ps1 directly and read its stderr.
//                    Proves the host emits the line at all -- without this a
//                    silent arm B is unattributable between "never written"
//                    and "written and lost".
//   arm B (subject): drive the same sweep through WordHost and record every
//                    line handed to the log callback the extension supplies.
//
// Both arms point at the SAME live WINWORD, recorded in a legacy single-field
// pid file so the sweep must decline. The Word at risk is one we started, so a
// probe bug damages nothing but this probe.
import { spawn } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const extRoot = path.resolve(here, "../../../.github/extensions/office-canvas");
const SCRIPT = path.join(extRoot, "src/word/word-host.ps1");
const { WordHost } = await import(pathToFileURL(path.join(extRoot, "src/word/word-host.mjs")).href);

const DECLINE = "refusing to reap";

async function wordCensus() {
    return await new Promise((resolve) => {
        const p = spawn("powershell.exe", [
            "-NoProfile", "-NonInteractive", "-Command",
            "(Get-Process -Name WINWORD -ErrorAction SilentlyContinue).Id -join ','",
        ]);
        let out = "";
        p.stdout.on("data", (c) => (out += c));
        p.on("close", () => resolve(out.trim() ? out.trim().split(",").map(Number) : []));
    });
}

// A dead-but-real host pid: this child has exited by the time its $PID reaches
// us, so the sweep treats the entry as orphaned rather than live.
async function deadPid() {
    return await new Promise((resolve) => {
        const p = spawn("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", "$PID"]);
        let out = "";
        p.stdout.on("data", (c) => (out += c));
        p.on("close", () => resolve(Number.parseInt(out.trim(), 10)));
    });
}

// The census travels with every figure: these numbers were not taken on a quiet
// machine and nothing here has measured what a quiet one would do.
console.log(`WINWORD census at start: ${(await wordCensus()).length}`);

const owner = new WordHost({ log: () => {} });
let ownedPid = null;
const dirs = [];
try {
    ({ ownedPid } = await owner.ping());
    if (!ownedPid) throw new Error("the host did not report a Word pid; the probe has nothing to point at");
    console.log(`live WINWORD used as the decline target: ${ownedPid}`);

    // ---- arm A: control -- spawn the host directly and read raw stderr.
    const dirA = await mkdtemp(path.join(tmpdir(), "decline-reach-a-"));
    dirs.push(dirA);
    await writeFile(path.join(dirA, `${await deadPid()}.pid`), `${ownedPid}`, "utf8");
    const rawStderr = await new Promise((resolve) => {
        const child = spawn("powershell.exe", [
            "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
            "-File", SCRIPT, "-PidDir", dirA,
        ], { stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
        let err = "";
        child.stderr.setEncoding("utf8");
        child.stderr.on("data", (c) => (err += c));
        child.stdout.on("data", () => {});
        child.stdin.write(JSON.stringify({ id: 1, cmd: "ping", args: {} }) + "\n");
        // The sweep runs at startup, before ping is answered; quit as soon as we
        // have given it time to run rather than waiting on Word.
        setTimeout(() => {
            child.stdin.write(JSON.stringify({ id: 2, cmd: "quit", args: {} }) + "\n");
        }, 4000);
        child.on("exit", () => resolve(err));
        setTimeout(() => { try { child.kill(); } catch { } }, 40000);
    });
    const armA = rawStderr.includes(DECLINE);
    console.log(`arm A (raw stderr from word-host.ps1): decline line present = ${armA}`);
    for (const line of rawStderr.split(/\r?\n/).filter((l) => l.trim())) console.log(`    A| ${line}`);

    // ---- arm B: subject -- the same sweep, through the client the extension uses.
    const dirB = await mkdtemp(path.join(tmpdir(), "decline-reach-b-"));
    dirs.push(dirB);
    await writeFile(path.join(dirB, `${await deadPid()}.pid`), `${ownedPid}`, "utf8");
    const logged = [];
    const reaper = new WordHost({ log: (m) => logged.push(m), pidDir: dirB });
    try {
        await reaper.ping();
    } finally {
        await reaper.dispose();
    }
    const armB = logged.some((l) => l.includes(DECLINE));
    console.log(`arm B (WordHost log callback):        decline line present = ${armB}`);
    for (const line of logged) console.log(`    B| ${line}`);

    console.log("");
    if (armA && !armB) {
        console.log("RESULT: the host writes the decline and the client drops it.");
        console.log("        Write-HostDiagnostic reports to a channel with no reader on a clean run.");
    } else if (armA && armB) {
        console.log("RESULT: the decline reaches the extension log. The channel is read.");
    } else if (!armA) {
        console.log("RESULT: arm A is silent, so the control failed: the host did not emit the line.");
        console.log("        Arm B says nothing about the transport until this is fixed.");
    }
} finally {
    await owner.dispose().catch(() => { });
    for (const d of dirs) await rm(d, { recursive: true, force: true }).catch(() => { });
    console.log(`WINWORD census at end: ${(await wordCensus()).length}`);
    // Deliberately not asserting the two census reads are equal: another session
    // may start or end a Word mid-run and an inequality has causes this probe
    // cannot distinguish. Printed at both ends so the conditions travel with the
    // figures instead of living in a commit message.
}
