// Office-free coverage of the one bound that makes a wedged save (#96) survivable:
// `WordHost.#send` arms a timeout, and on expiry it kills the child and rejects
// with a typed `word_timeout`. Before this test that guarantee -- the timer, the
// `#kill`, the typed reject -- had zero coverage: `git grep word_timeout` found
// no test touching word-host.mjs, and the author-level test asserts only that a
// finite `timeoutMs` is *passed*, not that anything acts on it.
//
// A save that never answers is exactly a `SaveAs2` that has wedged behind a
// modal or a converter negotiation. The `launch` seam puts a child that speaks
// the JSON-RPC line protocol but drops `create` in Word's place, so the only
// thing that can end the wait is WordHost's own timer.
//
// Mutations this catches (results in the PR body). Note the failure shape: a
// broken bound makes this test HANG, not report a red assertion, and a hang is
// simply never a green pass -- that is the property that matters here.
//   - Delete the `setTimeout` block in `#send` (word-host.mjs) so the wedged
//     `create` is never bounded: `assert.rejects` never settles, so the test
//     never reaches a passing state. With a `--test-timeout` set the runner
//     kills it as a failure; with node:test's default (`Infinity`) it hangs
//     until the run is killed. Either way it cannot pass -- which is the point.
//   - Rename the host-side `timeoutMs` parameter so the value is ignored and the
//     bound silently reverts to a constant: the timer still fires here because
//     the timer lives on the Node side, so this stays green -- which is why the
//     companion assertion in document-author.test.mjs checks the value handed
//     down, and why the two tests are needed together.

import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { WordHost } from "../../src/word/word-host.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const STUB = path.join(here, "stub-hung-host.mjs");

/** A host whose child hangs on `create`, standing in for a wedged SaveAs2. */
function hungHost() {
    return new WordHost({ launch: { command: process.execPath, args: [STUB] } });
}

/**
 * A host whose child hangs on `create` and, additionally, takes `pingMs` to
 * answer `ping` -- the reply `#ensureStarted` awaits. That stands in for a slow
 * cold Word start, so a test can tell a deadline derived *after* start (the
 * #128 fix) from a window snapshotted before it. `log` is forwarded so a test
 * can read the derived window `#send` arms.
 */
function slowStartHungHost(pingMs, log = () => {}) {
    return new WordHost({ launch: { command: process.execPath, args: [STUB, String(pingMs)] }, log });
}

test("a save that never answers is bounded and rejects with a typed word_timeout", async () => {
    const host = hungHost();
    try {
        const started = Date.now();
        await assert.rejects(
            () => host.create({ path: "C:\\ignored.docx", blocks: [], timeoutMs: 200 }),
            (err) => {
                assert.equal(err.code, "word_timeout", `a wedged save rejected as ${err.code} rather than word_timeout`);
                return true;
            },
        );
        const elapsed = Date.now() - started;
        assert.ok(elapsed < 30_000, `the reject took ${elapsed}ms -- the bound did not fire`);
    } finally {
        await host.dispose().catch(() => {});
    }
});

test("the host is torn down after a save times out, so the next call starts fresh", async () => {
    const host = hungHost();
    try {
        await host.create({ path: "C:\\ignored.docx", blocks: [], timeoutMs: 200 }).then(
            () => assert.fail("the wedged save resolved"),
            (err) => assert.equal(err.code, "word_timeout"),
        );
        // #send's timeout path calls #kill(), which nulls the child. A host left
        // running behind a dead COM call is the failure mode the restart exists
        // to avoid: the next request must be able to start a fresh child.
        assert.equal(host.running, false, "the host was left running behind a wedged save");
    } finally {
        await host.dispose().catch(() => {});
    }
});

// #128: the cold Word start must be spent *from* the operation's budget, not
// added on top of it. `#send` awaits `#ensureStarted()` before arming its timer,
// so a window derived after that await shares one wall clock with the start,
// whereas a value snapshotted before it composes on top -- ~180s of start plus
// the whole window, against an error that quotes only the window.
//
// This is observable without Word: the stub delays its `ping` reply (the frame
// `#ensureStarted` awaits) by `pingMs`, then wedges `create`. The host is given
// a `deadline` whose entire budget is consumed by that delayed start, so a
// correctly-composed `#send` has ~0ms left and arms its timer at ~0ms. The
// companion assertion in document-author.test.mjs checks that the author hands
// down `deadline` rather than a snapshot, so the two together pin caller and host.
//
// This reads the *derived* timeout figure, not wall-clock elapsed, on purpose.
// A test that asserts "the reject came back under N ms" passes on a fast machine
// for the wrong reason and flakes on a slow one -- the clock-generosity class
// this repo keeps finding, and the class the #128 fix itself is about. `#send`
// logs the window it armed (`timed out after <armMs>ms`), which is exactly the
// value the derivation produces. So the assertion is on that number:
//   - Fix: the cold start (pingMs) outlasts the budget, so when `#send` derives
//     `deadline - Date.now()` *after* the start it is already <= 0, clamped to 0
//     -- it arms ~0 ms and the log says so.
//   - Snapshot mutant (derive `deadline - Date.now()` *before* the await): the
//     window is the full budget, snapshotted while the deadline was still in the
//     future, so it arms ~budgetMs.
// No wall clock is trusted: a slow box changes how long the start takes, not the
// derived figure, so this cannot pass or fail for a timing accident.
//
// Mutation that reddens this (result in the PR body): move the
// `timeoutMs`/`deadline` derivation in `#send` to before `await
// this.#ensureStarted()`. The armed figure jumps from ~0 to ~budgetMs and the
// assertion fails.
test("a slow start is spent from the deadline: the armed window is what is left, not the whole budget", async () => {
    const pingMs = 600;
    const budgetMs = 400; // shorter than the start, so the deadline expires mid-start

    let armedMs = null;
    const log = (line) => {
        const m = /timed out after (\d+)ms/.exec(line);
        if (m) armedMs = Number(m[1]);
    };
    const host = slowStartHungHost(pingMs, log);
    try {
        await assert.rejects(
            () => host.create({ path: "C:\\ignored.docx", blocks: [], deadline: Date.now() + budgetMs }),
            (err) => {
                assert.equal(err.code, "word_timeout", `rejected as ${err.code} rather than word_timeout`);
                return true;
            },
        );
        assert.notEqual(armedMs, null, "the host never logged an armed window -- the timeout path did not run");
        // The whole budget was spent by the cold start, so what is left to arm is
        // ~0. A snapshot taken before the start would arm the full budget. The
        // midpoint is the cleanest place to split them, and it is a claim about
        // the derived figure, not about elapsed time.
        assert.ok(
            armedMs < budgetMs / 2,
            `the host armed ${armedMs}ms; the start was added on top of the ${budgetMs}ms budget, not spent from it`,
        );
    } finally {
        await host.dispose().catch(() => {});
    }
});

// #128 gave `#send` three input shapes for one budget parameter (a bare number,
// `{ deadline }`, `{ timeoutMs }`). The hazard of that is a dropped budget: an
// object carrying neither key silently degrading to an unbounded or
// default-bounded call. `#send` guards it with a `TypeError` rather than a
// default. The public `create`/`edit`/`structure` wrappers each default their
// value so they cannot trip it -- the safe design -- so this exercises the guard
// through `request`, the one public method that forwards a caller's budget
// object verbatim, with the realistic slip of an empty object.
test("a budget object carrying neither a deadline nor a timeoutMs is rejected, not silently defaulted", async () => {
    const host = hungHost();
    try {
        await assert.rejects(
            () => host.request("info", { docId: "d" }, {}),
            (err) => {
                assert.ok(err instanceof TypeError, `a dropped budget surfaced as ${err?.name} rather than TypeError`);
                return true;
            },
        );
    } finally {
        await host.dispose().catch(() => {});
    }
});

