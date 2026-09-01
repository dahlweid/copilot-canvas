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
// Mutations this can go red for (results in the PR body):
//   - Delete the `setTimeout` block in `#send` (word-host.mjs) so the wedged
//     `create` is never bounded: this test's `assert.rejects` never settles and
//     the case fails on its own deadline rather than passing.
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
