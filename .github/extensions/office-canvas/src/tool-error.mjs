// The tool boundary's error shape, in its own module so it can be tested.
//
// This lived in `extension.mjs`, which cannot be imported from a test because
// it calls `joinSession()` at module scope. That was not a neutral detail: the
// one defect this function has caused twice is *invisible* from either side of
// it -- an error carrying rich details looks correct in the editor's own tests,
// and the agent receives a bare code, and nothing in between was assertable.
//
// Extracting it is the smallest change that lets a test stand where the agent
// stands. See CONTEXT.md: assert through the boundary the caller actually sees,
// not one layer beneath it.

/**
 * Reduces any internal error to what a tool caller receives.
 *
 * Deliberately narrow: `code`, `message` and `data` and nothing else. Anything
 * a thrower attaches as a top-level property is dropped here, silently and by
 * design -- internal errors carry references (snapshots, COM results, causes)
 * that have no meaning to an agent and should not be serialized to one.
 *
 * The corollary is the trap: a fact worth telling the agent must be on `data`.
 * `EditError`'s constructor is what guarantees that for the edit path.
 */
export function asToolError(err) {
    const code = err?.code ?? "word_error";
    const message = err?.message ?? "Unknown error";
    // Tools are not canvas actions, so `CanvasError` means nothing to a tool
    // caller. The code is folded into the message instead, because that is what
    // the agent actually reads when deciding what to do next.
    const wrapped = new Error(`${code}: ${message}`);
    wrapped.code = code;
    // Facts the host attached to the failure, such as `writable: false` on a
    // locked original. Dropping them here would leave the agent with a code and
    // no way to tell why.
    if (err?.data) wrapped.data = err.data;
    return wrapped;
}
