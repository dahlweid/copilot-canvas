// Measures the ceiling on Stop-Word's exit wait.
//
// Context: #26 measured Quit()-to-exit at 2563-3593 ms and proposed raising the
// host's wait to 30 s, on the grounds that word-pids.mjs waits far longer. The
// question this probe answers is whether the host's wait is free to be chosen at
// all, because Stop-Word does not run in a vacuum -- it runs *inside* the `quit`
// JSON-RPC command, under a client-side timeout, and word-host.mjs kills the
// host process the moment that timeout expires (`#send`, "restarting host").
//
// So the wait has a hard ceiling: the client's quit timeout. Below it, Word gets
// its graceful exit. Above it, the client gives up first and kills the host
// mid-wait -- which destroys the graceful exit that raising the number was meant
// to protect. That is a claim about a system, not about Word, so it is measured
// here rather than argued.
//
// Method: run the real host script with its wait made unconditional (no early
// break) and bounded by the arm's value, so the arm's number is what the quit
// actually costs. Speak the wire protocol directly rather than through WordHost,
// so the client timeout is an input rather than a constant we inherit.
//
// Run:
//   node spikes/isolation/probes/probe-quit-rpc-ceiling.mjs

import { spawn, execFileSync } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { tmpdir } from "node:os";
import path from "node:path";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const WORD_DIR = path.resolve(HERE, "../../../.github/extensions/office-canvas/src/word");
const HOST_PS1 = path.join(WORD_DIR, "word-host.ps1");
const HOST_MJS = path.join(WORD_DIR, "word-host.mjs");

// An anchor that has moved must be a hard failure, not a skip: a probe that
// silently fails to apply its own setup reports on a system it never built.
function replaceOnce(text, anchor, replacement, what) {
    const hits = text.split(anchor).length - 1;
    if (hits !== 1) throw new Error(`anchor for ${what} matched ${hits} times, expected exactly 1`);
    return text.replace(anchor, replacement);
}

// Read the ceiling from the source rather than restating it. A restatement is a
// defect when nothing fails if the fact changes underneath it -- and this number
// is the entire subject of the probe.
const hostMjs = await readFile(HOST_MJS, "utf8");
const timeoutMatch = hostMjs.match(/#send\("quit",\s*\{\},\s*([\d_]+)\)/);
if (!timeoutMatch) throw new Error("could not read the quit timeout out of word-host.mjs -- has #send('quit', ...) moved?");
const QUIT_TIMEOUT_MS = Number.parseInt(timeoutMatch[1].replaceAll("_", ""), 10);

const hostPs1 = await readFile(HOST_PS1, "utf8");
const WAIT_ANCHOR = "while ($waited.ElapsedMilliseconds -lt 10000) {";
const BREAK_ANCHOR = "if ($null -eq $p -or $p.ProcessName -ne 'WINWORD') { break }";

function log(line) {
    process.stdout.write(`${line}\n`);
}

async function runArm(boundMs) {
    let text = replaceOnce(hostPs1, WAIT_ANCHOR, `while ($waited.ElapsedMilliseconds -lt ${boundMs}) {`, "the wait bound");
    // Removing the break is what makes the arm's number the cost of the quit
    // rather than the cost of Word's exit. Without it both arms finish in ~3 s
    // and the ceiling is never approached, so the probe would report "no
    // difference" while never having tested one.
    text = replaceOnce(text, BREAK_ANCHOR, "# probe: wait the full bound", "the early break");

    const dir = await mkdtemp(path.join(tmpdir(), "quit-ceiling-"));
    const script = path.join(dir, "word-host.ps1");
    await writeFile(script, text, "utf8");

    const child = spawn(
        "powershell.exe",
        ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script, "-PidDir", dir],
        { stdio: ["pipe", "pipe", "pipe"], windowsHide: true },
    );
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stderr.on("data", () => {});

    let buffer = "";
    const pending = new Map();
    child.stdout.on("data", (chunk) => {
        buffer += chunk;
        let i;
        while ((i = buffer.indexOf("\n")) >= 0) {
            const line = buffer.slice(0, i).trim();
            buffer = buffer.slice(i + 1);
            if (!line) continue;
            let msg;
            try {
                msg = JSON.parse(line);
            } catch {
                continue;
            }
            const entry = pending.get(msg.id);
            if (!entry) continue;
            pending.delete(msg.id);
            clearTimeout(entry.timer);
            entry.resolve(msg);
        }
    });

    let nextId = 1;
    const send = (cmd, timeoutMs) =>
        new Promise((resolve) => {
            const id = nextId++;
            const timer = setTimeout(() => {
                pending.delete(id);
                resolve({ timedOut: true });
            }, timeoutMs);
            pending.set(id, { resolve, timer });
            child.stdin.write(`${JSON.stringify({ id, cmd, args: {} })}\n`);
        });

    let ownedPid = null;
    try {
        const ping = await send("ping", 120_000);
        ownedPid = ping?.result?.ownedPid ?? null;
        if (!ownedPid) throw new Error("the host did not report an owned Word pid; the arm would prove nothing");

        const t0 = Date.now();
        const res = await send("quit", QUIT_TIMEOUT_MS);
        const elapsed = Date.now() - t0;

        if (res.timedOut) {
            // This is word-host.mjs's #kill() path: it kills the host here.
            child.kill();
            return { boundMs, ownedPid, elapsed, outcome: "CLIENT TIMED OUT -- host killed mid-wait" };
        }
        return { boundMs, ownedPid, elapsed, outcome: res.ok ? "quit returned cleanly" : `quit failed: ${res.error?.code}` };
    } finally {
        try {
            child.kill();
        } catch {}
        // Only ever the pid this arm minted and printed.
        if (ownedPid) {
            try {
                execFileSync("powershell.exe", [
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    `$p = Get-Process -Id ${ownedPid} -ErrorAction SilentlyContinue; ` +
                        `if ($null -ne $p -and $p.ProcessName -eq 'WINWORD') { $p.Kill(); 'probe killed ${ownedPid}' }`,
                ], { encoding: "utf8", stdio: ["ignore", "inherit", "ignore"] });
            } catch {}
        }
        await rm(dir, { recursive: true, force: true }).catch(() => {});
    }
}

log(`client quit timeout read from word-host.mjs: ${QUIT_TIMEOUT_MS} ms`);
log("");

for (const bound of [10_000, 30_000]) {
    const r = await runArm(bound);
    const side = bound < QUIT_TIMEOUT_MS ? "under" : "over";
    log(`bound ${String(bound).padStart(6)} ms (${side} the ${QUIT_TIMEOUT_MS} ms ceiling): quit took ${String(r.elapsed).padStart(6)} ms -- ${r.outcome}`);
}

log("");
log("A wait longer than the client's quit timeout does not buy Word more time to");
log("exit gracefully. It spends the difference getting the host killed.");
