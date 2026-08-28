// Node side of the Word bridge.
//
// Owns a single long-lived `powershell.exe` child running word-host.ps1, which
// in turn owns the hidden Word instance. Spawning PowerShell per operation
// would cost ~0.5-1s each; keeping it warm makes repeated queries cheap.
//
// Requests are correlated by id, so they may be pipelined; the PowerShell host
// processes one line at a time.

import { spawn, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import path from "node:path";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SCRIPT_PATH = path.join(HERE, "word-host.ps1");

const DEFAULT_TIMEOUT_MS = 120_000;
/** Starting Word cold, plus loading the PDF export engine, is genuinely slow. */
const STARTUP_TIMEOUT_MS = 180_000;

export class WordUnavailableError extends Error {
    constructor(message) {
        super(message);
        this.name = "WordUnavailableError";
        this.code = "word_unavailable";
    }
}

/**
 * Ends an orphaned Word instance we created.
 *
 * The PowerShell host normally quits Word itself, but if that process dies
 * without running its teardown (crash, external kill, timeout kill) the hidden
 * WINWORD is left behind with no owner. Synchronous on purpose so it can also
 * run from a `process.on("exit")` hook.
 *
 * The process-name check guards against killing an unrelated process that has
 * since inherited the recycled PID.
 */
export function reapWordProcess(pid) {
    if (!pid || process.platform !== "win32") return false;
    try {
        const listed = spawnSync("tasklist", ["/FI", `PID eq ${pid}`, "/FO", "CSV", "/NH"], {
            encoding: "utf8",
            windowsHide: true,
        });
        if (!/WINWORD\.EXE/i.test(listed.stdout ?? "")) return false;
        spawnSync("taskkill", ["/PID", String(pid), "/F", "/T"], { encoding: "utf8", windowsHide: true });
        return true;
    } catch {
        return false;
    }
}

export class WordHost {
    #child = null;
    #buffer = "";
    #stderr = "";
    #pending = new Map();
    #nextId = 1;
    #starting = null;
    #disposed = false;
    /** docId -> the args it was opened with, so we can replay after a restart. */
    #openArgs = new Map();

    constructor({ log = () => {}, onOwnedPid = () => {}, pidDir = null } = {}) {
        this.log = log;
        this.onOwnedPid = onOwnedPid;
        this.pidDir = pidDir;
        this.ownedPid = null;
        this.wordVersion = null;
    }

    get running() {
        return this.#child !== null && this.#child.exitCode === null;
    }

    async #ensureStarted() {
        if (this.#disposed) throw new WordUnavailableError("The Word host has been shut down.");
        if (this.running) return;
        if (this.#starting) return this.#starting;

        this.#starting = (async () => {
            this.#buffer = "";
            this.#stderr = "";

            const child = spawn(
                "powershell.exe",
                [
                    "-NoProfile",
                    "-NonInteractive",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-File",
                    SCRIPT_PATH,
                    ...(this.pidDir ? ["-PidDir", this.pidDir] : []),
                ],
                { stdio: ["pipe", "pipe", "pipe"], windowsHide: true },
            );

            child.stdout.setEncoding("utf8");
            child.stdout.on("data", (chunk) => this.#onStdout(chunk));
            child.stderr.setEncoding("utf8");
            child.stderr.on("data", (chunk) => {
                this.#stderr = (this.#stderr + chunk).slice(-4000);
            });
            child.on("exit", (code, signal) => this.#onExit(code, signal));
            child.on("error", (err) => {
                this.#rejectAll(new WordUnavailableError(`Could not start powershell.exe: ${err.message}`));
            });

            this.#child = child;
        })();

        try {
            await this.#starting;
        } finally {
            this.#starting = null;
        }
    }

    #onStdout(chunk) {
        this.#buffer += chunk;
        let index;
        while ((index = this.#buffer.indexOf("\n")) >= 0) {
            const line = this.#buffer.slice(0, index).trim();
            this.#buffer = this.#buffer.slice(index + 1);
            if (!line) continue;
            let message;
            try {
                message = JSON.parse(line);
            } catch {
                // The host must only ever write protocol frames to stdout. Anything
                // else is a bug worth seeing in the extension log.
                this.log(`word-host: unparsable stdout line: ${line.slice(0, 300)}`);
                continue;
            }
            const entry = this.#pending.get(message.id);
            if (!entry) continue;
            this.#pending.delete(message.id);
            clearTimeout(entry.timer);
            if (message.ok) entry.resolve(message.result);
            else {
                const err = new Error(message.error?.message ?? "Unknown Word error");
                err.code = message.error?.code ?? "word_error";
                entry.reject(err);
            }
        }
    }

    #onExit(code, signal) {
        const detail = this.#stderr.trim();
        this.#child = null;
        // The host owned the hidden Word instance. If it exited without running
        // its own teardown, that WINWORD process is orphaned -- end it here.
        const orphan = this.ownedPid;
        this.ownedPid = null;
        if (orphan && reapWordProcess(orphan)) {
            this.log(`word-host: reaped orphaned WINWORD ${orphan} after host exit`);
        }
        this.#rejectAll(
            new WordUnavailableError(
                `The Word host exited (code ${code}, signal ${signal}).${detail ? ` ${detail}` : ""}`,
            ),
        );
    }

    /** Ends the Word instance we own, if any. Safe to call from an exit hook. */
    reapOwnedWord() {
        const pid = this.ownedPid;
        this.ownedPid = null;
        return reapWordProcess(pid);
    }

    #rejectAll(error) {
        for (const [, entry] of this.#pending) {
            clearTimeout(entry.timer);
            entry.reject(error);
        }
        this.#pending.clear();
    }

    async #send(cmd, args, timeoutMs) {
        await this.#ensureStarted();
        const id = this.#nextId++;
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                this.#pending.delete(id);
                // A timeout usually means Word is stuck behind a modal dialog.
                // The process is unrecoverable from here, so tear it down; the
                // next request transparently starts a fresh one.
                this.log(`word-host: '${cmd}' timed out after ${timeoutMs}ms, restarting host`);
                this.#kill();
                const err = new Error(`Word did not respond to '${cmd}' within ${Math.round(timeoutMs / 1000)}s.`);
                err.code = "word_timeout";
                reject(err);
            }, timeoutMs);

            this.#pending.set(id, { resolve, reject, timer });
            try {
                this.#child.stdin.write(`${JSON.stringify({ id, cmd, args: args ?? {} })}\n`);
            } catch (err) {
                this.#pending.delete(id);
                clearTimeout(timer);
                reject(new WordUnavailableError(`Could not write to the Word host: ${err.message}`));
            }
        });
    }

    /**
     * Sends a command, transparently recovering from a host restart by
     * replaying the `open` that the command depends on.
     */
    async request(cmd, args = {}, { timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
        try {
            return await this.#send(cmd, args, timeoutMs);
        } catch (err) {
            const docId = args?.docId;
            const message = err.message ?? "";
            const recoverable =
                docId &&
                this.#openArgs.has(docId) &&
                cmd !== "open" &&
                (err.code === "word_unavailable" ||
                    err.code === "word_timeout" ||
                    /No open document with id/i.test(message) ||
                    // COM surfaces a dead Word as RPC_E_DISCONNECTED / 0x800706BA,
                    // localized, so match the HRESULT rather than the text.
                    /0x800706BA|0x80010108|RPC/i.test(message));
            if (!recoverable) throw err;

            this.log(`word-host: replaying open for '${docId}' after ${err.code ?? "error"}`);
            await this.#send("open", this.#openArgs.get(docId), STARTUP_TIMEOUT_MS);
            // A recovery means Word was replaced, so the pid we recorded is
            // stale; re-learn it or we would later reap the wrong process.
            try {
                await this.ping();
            } catch {
                /* non-fatal */
            }
            return await this.#send(cmd, args, timeoutMs);
        }
    }

    async ping() {
        const result = await this.#send("ping", {}, STARTUP_TIMEOUT_MS);
        this.ownedPid = result.ownedPid ?? null;
        this.wordVersion = result.wordVersion ?? null;
        if (this.ownedPid) this.onOwnedPid(this.ownedPid);
        return result;
    }

    async openDocument({ docId, path: docPath, workDir }) {
        const args = { docId, path: docPath, workDir };
        const result = await this.#send("open", args, STARTUP_TIMEOUT_MS);
        this.#openArgs.set(docId, args);
        // The ping may not have run yet; capture ownership as soon as we know it.
        if (!this.ownedPid) {
            try {
                await this.ping();
            } catch {
                /* non-fatal */
            }
        }
        return result;
    }

    exportPdf({ docId, out, from, to }) {
        return this.request("export", { docId, out, from, to }, { timeoutMs: STARTUP_TIMEOUT_MS });
    }

    /**
     * Reads the document's WordprocessingML into `out` and closes it again.
     *
     * Bounded by the startup timeout rather than the default one because it
     * includes a cold Word start. ADR 0005 requires every `Documents.Open` to
     * be timeout-bounded: detection and open are not atomic, so a file that
     * looked free can still be taken between the two, and an unbounded open
     * against a held file hangs forever.
     */
    structure({ docId, path: docPath, workDir, out }) {
        return this.request(
            "structure",
            { docId, path: docPath, workDir, out },
            { timeoutMs: STARTUP_TIMEOUT_MS },
        );
    }

    outline({ docId, limit }) {
        return this.request("outline", { docId, limit });
    }

    search({ docId, query, limit, matchCase, wholeWord }) {
        return this.request("search", { docId, query, limit, matchCase, wholeWord });
    }

    text({ docId, fromPage, toPage }) {
        return this.request("text", { docId, fromPage, toPage });
    }

    info({ docId }) {
        return this.request("info", { docId });
    }

    async closeDocument({ docId }) {
        this.#openArgs.delete(docId);
        if (!this.running) return { closed: true };
        try {
            return await this.#send("close", { docId }, 30_000);
        } catch {
            return { closed: false };
        }
    }

    #kill() {
        const child = this.#child;
        this.#child = null;
        if (child && child.exitCode === null) {
            try {
                child.kill();
            } catch {
                /* already gone */
            }
        }
    }

    /** Quits Word cleanly, then makes sure the process is gone. */
    async dispose() {
        if (this.#disposed) return;
        this.#openArgs.clear();
        if (!this.running) {
            this.#disposed = true;
            this.#kill();
            this.reapOwnedWord();
            return;
        }
        const child = this.#child;
        let quit = false;
        try {
            await this.#send("quit", {}, 20_000);
            quit = true;
        } catch (err) {
            this.log(`word-host: quit failed (${err.code ?? "error"}): ${err.message}`);
        }
        // A successful quit means the host already ended Word; forget the pid so
        // the exit handler cannot kill a process that has since reused it.
        if (quit) this.ownedPid = null;
        // Only now: `#send` refuses to run against a disposed host, so flipping
        // this any earlier would reject the very quit we just sent.
        this.#disposed = true;
        await new Promise((resolve) => {
            if (!child || child.exitCode !== null) return resolve();
            const timer = setTimeout(resolve, 5000);
            child.once("exit", () => {
                clearTimeout(timer);
                resolve();
            });
        });
        this.#kill();
    }
}
