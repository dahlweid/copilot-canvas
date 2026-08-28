// Live Word pixel-streaming spike -- measurement driver.
//
// Answers the question the plan poses: is streaming a live Word window good
// enough to replace the PDF renderer? Measures cold start, capture latency,
// scroll latency and frame size, and writes sample frames for visual comparison
// against the PDF render of the same document.
//
// Run: node .github/extensions/word-canvas/spikes/live-word/run-spike.mjs [docx]

import { spawn, execFile } from "node:child_process";
import { mkdir, writeFile, rm } from "node:fs/promises";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { fileURLToPath } from "node:url";
import { performance } from "node:perf_hooks";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const EXT_ROOT = path.resolve(HERE, "..", "..");
const OUT_DIR = path.join(HERE, "frames");

class LiveWordHost {
    #child;
    #pending = new Map();
    #buffer = "";
    #nextId = 1;

    constructor({ pidDir }) {
        this.#child = spawn(
            "powershell.exe",
            [
                "-NoProfile",
                "-NonInteractive",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                path.join(HERE, "live-word.ps1"),
                "-PidDir",
                pidDir,
            ],
            { stdio: ["pipe", "pipe", "pipe"], windowsHide: true },
        );
        this.#child.stdout.setEncoding("utf8");
        this.#child.stdout.on("data", (chunk) => this.#onData(chunk));
        this.#child.stderr.setEncoding("utf8");
        this.#child.stderr.on("data", (chunk) => process.stderr.write(`[ps] ${chunk}`));
        this.#child.on("exit", (code) => {
            for (const { reject } of this.#pending.values()) {
                reject(new Error(`host exited with code ${code}`));
            }
            this.#pending.clear();
        });
    }

    #onData(chunk) {
        this.#buffer += chunk;
        let index;
        while ((index = this.#buffer.indexOf("\n")) !== -1) {
            const line = this.#buffer.slice(0, index).trim();
            this.#buffer = this.#buffer.slice(index + 1);
            if (!line) continue;
            let message;
            try {
                message = JSON.parse(line);
            } catch {
                process.stderr.write(`[ps non-json] ${line}\n`);
                continue;
            }
            const waiter = this.#pending.get(message.id);
            if (!waiter) continue;
            this.#pending.delete(message.id);
            if (message.ok) waiter.resolve(message.result);
            else waiter.reject(new Error(message.error));
        }
    }

    request(cmd, args = {}, timeoutMs = 120_000) {
        const id = this.#nextId++;
        return new Promise((resolve, reject) => {
            const timer = setTimeout(() => {
                this.#pending.delete(id);
                reject(new Error(`${cmd} timed out after ${timeoutMs}ms`));
            }, timeoutMs);
            this.#pending.set(id, {
                resolve: (v) => {
                    clearTimeout(timer);
                    resolve(v);
                },
                reject: (e) => {
                    clearTimeout(timer);
                    reject(e);
                },
            });
            this.#child.stdin.write(`${JSON.stringify({ id, cmd, args })}\n`);
        });
    }

    async dispose() {
        try {
            await this.request("quit", {}, 15_000);
        } catch {
            /* the host may already be gone */
        }
        await new Promise((r) => setTimeout(r, 500));
        if (this.#child.exitCode === null) this.#child.kill();
    }
}

const stats = (values) => {
    const sorted = [...values].sort((a, b) => a - b);
    const at = (q) => sorted[Math.min(sorted.length - 1, Math.floor(q * sorted.length))];
    return {
        n: sorted.length,
        min: +sorted[0].toFixed(1),
        p50: +at(0.5).toFixed(1),
        p95: +at(0.95).toFixed(1),
        max: +sorted[sorted.length - 1].toFixed(1),
        mean: +(sorted.reduce((a, b) => a + b, 0) / sorted.length).toFixed(1),
    };
};

const wordPids = async () => {
    const { stdout } = await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-Command",
        "(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) -join ','",
    ]);
    return stdout.trim() ? stdout.trim().split(",").map(Number) : [];
};

// --- run ---------------------------------------------------------------------

const workRoot = await mkdtemp(path.join(tmpdir(), "word-live-spike-"));
let fixture = process.argv[2];
if (!fixture) {
    fixture = path.join(workRoot, "spike.docx");
    process.stderr.write("generating fixture...\n");
    await execFileAsync("powershell.exe", [
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        path.join(EXT_ROOT, "test", "make-fixture.ps1"),
        "-Out",
        fixture,
        "-Chapters",
        "4",
    ]);
}

await mkdir(OUT_DIR, { recursive: true });
const pidsBefore = await wordPids();
const report = { document: fixture, measurements: {} };
const host = new LiveWordHost({ pidDir: path.join(workRoot, "pids") });

try {
    process.stderr.write("starting live Word (window parked off-screen)...\n");
    const coldStart = performance.now();
    const started = await host.request("start", { path: fixture, width: 560, height: 950 });
    report.measurements.coldStartMs = Math.round(performance.now() - coldStart);
    report.window = started;
    process.stderr.write(
        `  started in ${report.measurements.coldStartMs}ms: pid ${started.pid}, ` +
            `${started.width}x${started.height}px, ${started.pageCount} pages\n`,
    );

    // 1. Capture throughput -- the number that decides whether this is viable.
    const captureMs = [];
    const encodeMs = [];
    const sizes = [];
    for (let i = 0; i < 60; i++) {
        const t0 = performance.now();
        const frame = await host.request("capture", { quality: 70 });
        captureMs.push(performance.now() - t0);
        encodeMs.push(frame.ms);
        sizes.push(frame.bytes);
    }
    report.measurements.captureRoundTripMs = stats(captureMs);
    report.measurements.captureInProcessMs = stats(encodeMs);
    report.measurements.frameBytes = stats(sizes);
    report.measurements.fps = +(1000 / report.measurements.captureRoundTripMs.mean).toFixed(1);
    process.stderr.write(
        `  capture: ${report.measurements.captureRoundTripMs.mean}ms mean ` +
            `(${report.measurements.fps} fps), ${Math.round(report.measurements.frameBytes.mean / 1024)} KB/frame\n`,
    );

    // 2. Scroll latency -- COM scroll plus the capture that has to follow it.
    const scrollMs = [];
    const scrollFrameMs = [];
    for (let i = 0; i < 20; i++) {
        const t0 = performance.now();
        await host.request("scroll", { down: 3 });
        scrollMs.push(performance.now() - t0);
        const t1 = performance.now();
        await host.request("capture", { quality: 70 });
        scrollFrameMs.push(performance.now() - t1);
    }
    report.measurements.scrollMs = stats(scrollMs);
    report.measurements.scrollToFrameMs = stats(scrollMs.map((v, i) => v + scrollFrameMs[i]));
    process.stderr.write(
        `  scroll+frame: ${report.measurements.scrollToFrameMs.mean}ms mean\n`,
    );

    // 3. Page jump latency.
    const gotoMs = [];
    for (const page of [1, 3, 5, 2, 4]) {
        const t0 = performance.now();
        await host.request("goto", { page });
        await host.request("capture", { quality: 70 });
        gotoMs.push(performance.now() - t0);
    }
    report.measurements.gotoPageMs = stats(gotoMs);

    // 4. Sample frames for visual comparison, with and without chrome cropped.
    await host.request("goto", { page: 1 });
    await host.request("capture", { quality: 90, out: path.join(OUT_DIR, "frame-full.jpg") });
    await host.request("crop", { top: 190, bottom: 30 });
    await host.request("capture", { quality: 90, out: path.join(OUT_DIR, "frame-cropped.jpg") });

    // 5. Resize behaviour -- does the layout follow the panel width?
    const resized = await host.request("resize", { width: 900, height: 1200 });
    await host.request("crop", { top: 0, bottom: 0 });
    await host.request("capture", { quality: 90, out: path.join(OUT_DIR, "frame-wide.jpg") });
    report.measurements.resize = { requested: { width: 900, height: 1200 }, actual: resized };

    const info = await host.request("info");
    report.measurements.zoomPercent = info.zoom;
} catch (err) {
    report.error = err.message;
    process.stderr.write(`SPIKE FAILED: ${err.stack ?? err.message}\n`);
} finally {
    await host.dispose();
}

await new Promise((r) => setTimeout(r, 1500));
report.leakedWordProcesses = (await wordPids()).filter((pid) => !pidsBefore.includes(pid));

await writeFile(path.join(OUT_DIR, "report.json"), JSON.stringify(report, null, 2));
process.stderr.write(`\n${JSON.stringify(report.measurements, null, 2)}\n`);
process.stderr.write(`frames and report written to ${OUT_DIR}\n`);
if (report.leakedWordProcesses.length) {
    process.stderr.write(`LEAKED Word processes: ${report.leakedWordProcesses.join(", ")}\n`);
}

await rm(workRoot, { recursive: true, force: true }).catch(() => {});
process.exit(report.error || report.leakedWordProcesses.length ? 1 : 0);
