// A stand-in for word-host.ps1 that speaks the JSON-RPC line protocol but never
// answers `create`, modelling a `SaveAs2` that has wedged (#96). WordHost talks
// to whatever child its `launch` seam spawns, so this lets the timeout, the
// `#kill` and the typed `word_timeout` reject be exercised without Word.
//
// It answers `ping` so `#ensureStarted`'s learn-the-pid step completes and the
// start resolves, then goes deaf to `create`: the frame is read and dropped, so
// the only thing that can end the wait is WordHost's own timer.
//
// The first CLI arg, if a number, is a millisecond delay applied to the `ping`
// reply. `#ensureStarted` awaits that `ping`, so the delay stands in for a slow
// cold Word start: it lets a test prove the operation window is derived *after*
// start (spent from the shared deadline), rather than snapshotted before it and
// composed on top (#128). Everything but `ping` still answers immediately.

import { createInterface } from "node:readline";

const pingDelayMs = Number(process.argv[2]) || 0;

const rl = createInterface({ input: process.stdin });

rl.on("line", (line) => {
    const text = line.trim();
    if (!text) return;
    let msg;
    try {
        msg = JSON.parse(text);
    } catch {
        return;
    }

    if (msg.cmd === "create") {
        // The wedge. Read and drop it; never write a reply frame.
        return;
    }

    const reply = () => process.stdout.write(`${JSON.stringify({ id: msg.id, ok: true, result: {} })}\n`);

    if (msg.cmd === "ping" && pingDelayMs > 0) {
        // A slow cold start: the parent's `#ensureStarted` blocks here.
        setTimeout(reply, pingDelayMs);
        return;
    }

    // Everything else answers immediately, so startup and any control command
    // resolve and only the wedged `create` is left hanging.
    reply();
});

// Keep the process alive until the parent kills it.
process.stdin.resume();
