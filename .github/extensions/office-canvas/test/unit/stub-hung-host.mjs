// A stand-in for word-host.ps1 that speaks the JSON-RPC line protocol but never
// answers `create`, modelling a `SaveAs2` that has wedged (#96). WordHost talks
// to whatever child its `launch` seam spawns, so this lets the timeout, the
// `#kill` and the typed `word_timeout` reject be exercised without Word.
//
// It answers `ping` so `#ensureStarted`'s learn-the-pid step completes and the
// start resolves, then goes deaf to `create`: the frame is read and dropped, so
// the only thing that can end the wait is WordHost's own timer.

import { createInterface } from "node:readline";

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

    // Everything else answers immediately, so startup and any control command
    // resolve and only the wedged `create` is left hanging.
    process.stdout.write(`${JSON.stringify({ id: msg.id, ok: true, result: {} })}\n`);
});

// Keep the process alive until the parent kills it.
process.stdin.resume();
