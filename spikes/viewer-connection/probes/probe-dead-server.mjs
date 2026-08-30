// Does a viewer panel ever stop claiming it is reconnecting?
//
// Issue #66 rests on a claim about the platform: `EventSource` retries a dead
// origin forever, so `readyState` sits at CONNECTING and a status keyed on it is
// permanent. That was read out of the specification and then observed in the
// product from the outside -- a panel stuck for ~90 minutes on a port nothing
// was listening on -- but never driven end to end, and the unit suite's fake
// source cannot establish it: a double that returns what its author expected is
// exactly this repo's named trap.
//
// So this drives the real module with a real `EventSource` against a real
// `ViewerInstance` that is then closed, and does it twice: once through the
// handler that shipped before #66, once through `monitorConnection`. The
// difference between those two runs is the fix, measured rather than reasoned.
//
// Run:
//   node --experimental-eventsource spikes/viewer-connection/probes/probe-dead-server.mjs
//
// No Word, no document, no Office of any kind: `ViewerInstance.start()` only
// listens, and the single cache touch on the SSE path is `state.wordVersion`,
// which the stub below answers. Nothing here opens a document.
//
// What this is NOT. Node's `EventSource` is undici's, not Chromium's. Same
// specification, different implementation, so this is evidence about the
// mechanism and not a measurement of the webview.
//
// This paragraph used to go on: "Nothing in this repo reaches the webview:
// `app.js` cannot be imported under Node -- it resolves pdf.js from an absolute
// `/vendor/` URL and dies there -- and no browser harness exists." That was true
// when this probe was written and is not any more. #76 and #85 built
// `test/unit/ui-harness.mjs`, and `test/unit/connection-lost.test.mjs` now
// drives the committed `app.js` against the committed markup and asserts the
// panel's own `#status` element goes to the terminal message. So the untested
// span has narrowed to Chromium itself: this probe owns the real-`EventSource`
// end, that test owns the rendering end, and neither claims a webview.

import { ViewerInstance } from "../../../.github/extensions/office-canvas/src/server.mjs";
import {
    monitorConnection,
    GRACE_MS,
    DEADLINE_MS,
    RECONNECTING,
    LOST,
} from "../../../.github/extensions/office-canvas/src/ui/connection-status.mjs";

if (typeof globalThis.EventSource !== "function") {
    console.error("EventSource is not defined. Re-run with --experimental-eventsource (Node >= 22.3).");
    process.exit(2);
}

const STATE = ["CONNECTING", "OPEN", "CLOSED"];
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** Long enough to watch several retry attempts land after the deadline passes. */
const OBSERVE_MS = DEADLINE_MS + 8_000;

/**
 * The handler as it shipped before #66, verbatim from `app.js:420`.
 *
 * Kept as the counterfactual arm. Without it this probe would show the fixed
 * build terminating and could not show that the old one does not, which is the
 * half that makes the measurement a comparison instead of a demonstration.
 */
function attachShippedHandler(source, setStatus) {
    source.onerror = () => {
        setTimeout(() => {
            if (source.readyState === EventSource.CONNECTING) setStatus(RECONNECTING, { busy: true });
        }, 1500);
    };
}

async function run(label, attach) {
    const instance = new ViewerInstance({
        // Never resolved beyond `wordVersion`: no document is opened, so no
        // render is ever requested and no Word process is started.
        cache: () => ({ wordVersion: "probe" }),
        instanceId: `probe-${label}`,
        workspacePath: process.cwd(),
    });

    const url = await instance.start();
    const port = new URL(url).port;

    let status = null;
    let statusChanges = 0;
    let errors = 0;
    let attached = false;

    const source = new EventSource(new URL("/events", url));
    // Attachment is read from this end. A socket count cannot tell "detached"
    // from "never rendered" -- it read zero for a healthy panel while #66 was
    // being investigated -- but `onopen` firing is the connection reporting
    // itself.
    source.addEventListener("open", () => {
        attached = true;
    });
    source.addEventListener("error", () => {
        errors += 1;
    });

    attach(source, (text, options = {}) => {
        status = { text, ...options };
        statusChanges += 1;
    });

    await sleep(700);
    const attachedBeforeClose = attached;
    const stateWhileAlive = STATE[source.readyState];

    // The same call a normal panel close makes: ends every SSE client, then
    // closes the server. After it the port is gone for good -- the extension
    // allocates a fresh one per process, which is why no retry can succeed.
    await instance.close();

    const samples = [];
    const started = Date.now();
    while (Date.now() - started < OBSERVE_MS) {
        await sleep(1_000);
        samples.push({
            at: Math.round((Date.now() - started) / 1000),
            readyState: STATE[source.readyState],
            errors,
            says: status?.text ?? null,
        });
    }

    source.close();

    return { label, port, attachedBeforeClose, stateWhileAlive, statusChanges, errors, samples, final: status };
}

const results = [];
results.push(await run("shipped", attachShippedHandler));
results.push(await run("fixed", (source, setStatus) => monitorConnection(source, { setStatus })));

for (const r of results) {
    console.log(`\n=== ${r.label} (port ${r.port}) ===`);
    console.log(`attached before close : ${r.attachedBeforeClose} (readyState ${r.stateWhileAlive})`);
    for (const s of r.samples) {
        console.log(
            `  t+${String(s.at).padStart(2)}s  ${s.readyState.padEnd(10)} errors=${String(s.errors).padEnd(3)} says: ${s.says ?? "(nothing)"}`,
        );
    }
    console.log(`final: ${JSON.stringify(r.final)}`);
}

const [shipped, fixed] = results;
const verdict = {
    grace_ms: GRACE_MS,
    deadline_ms: DEADLINE_MS,
    both_attached: shipped.attachedBeforeClose && fixed.attachedBeforeClose,
    retries_forever: shipped.errors > 1,
    shipped_still_claiming_recovery: shipped.final?.text === RECONNECTING,
    shipped_readyState_at_end: shipped.samples.at(-1)?.readyState,
    fixed_reached_terminal: fixed.final?.text === LOST,
    fixed_terminal_not_busy: fixed.final?.busy !== true,
};

console.log(`\n=== verdict ===\n${JSON.stringify(verdict, null, 2)}`);

// A probe that cannot fail measures nothing. These are the propositions the
// issue and the fix rest on; if any stops holding, this exits non-zero rather
// than printing a table nobody reads to the end.
const required = [
    "both_attached",
    "retries_forever",
    "shipped_still_claiming_recovery",
    "fixed_reached_terminal",
    "fixed_terminal_not_busy",
];
const broken = required.filter((k) => verdict[k] !== true);
if (broken.length) {
    console.error(`\nFAILED: ${broken.join(", ")}`);
    process.exit(1);
}
console.log("\nOK: the shipped handler never stops claiming a recovery; the fix terminates.");
