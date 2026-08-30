// What a tool failure carries to the agent.
//
// ## The channel, measured
//
// A **thrown** error carries nothing at all. Measured end-to-end through a live
// CLI session against a probe extension whose failure carried a distinctive
// marker in each of the three candidate places -- message, `code`, `data`. The
// agent received:
//
//     Failed to execute `probe_throw_coded` tool with arguments: {} due to
//     error: Error: Tool execution failed
//
// Not one marker survived. Two hops discard it, and either alone would be
// enough:
//
//   1. `copilot-sdk/extension.js`, `_executeToolAndRespond`, reduces a throw to
//      `error.message` and forwards that string as
//      `handlePendingToolCall({ requestId, error: message })`. `code`, `data`,
//      `name` and `stack` end here.
//   2. The runtime's `tools_handle_pending_tool_call` binds that message, logs
//      it at *warning* level, and then rejects with a string literal --
//      `rejectExternalTool(r, new Error("Tool execution failed"))`. The message
//      ends here.
//
// This module used to fold the code into the thrown message, on the reasoning
// that message text is the one thing that survives. Measured, it is not: hop 2
// discards the message too. That is why #45's symptom outlived a fix aimed
// directly at it -- the fix put the fact one layer above the discard.
//
// See `spikes/tool-errors/FINDINGS.md` for the run, the markers and the two
// source citations, with the version each was taken on.
//
// ## What does survive
//
// The **result**. In the same run, a handler that *returned* the SDK's
// documented `ToolResultObject` reached the agent as its `textResultForLlm`,
// verbatim and whole. So a fact worth telling the agent has to be in that one
// string. Not on `code`, not on `data`, not in the message of a throw.
//
// The corollary, and the trap this module now exists to close: **a tool handler
// must not throw.** `extension.mjs` wraps every registered tool so that it
// cannot, at the registration site rather than per handler, because a new tool
// added later would otherwise inherit the broken channel by simply not being
// told.
//
// ## Why the text has to say, itself, that it failed
//
// Also measured: the returned failure arrived with no failure framing of any
// kind. `resultType: "failure"` reached the agent looking exactly like an
// ordinary success, and the sibling `error` field did not arrive at all. So the
// banner on the first line is not decoration -- it is the only thing that tells
// the agent this is a failure rather than a result to act on.

/** Beyond this, details are stated as too large rather than pasted. */
const MAX_DETAIL_CHARS = 4000;

/**
 * Renders a failure's `data` bag into the one channel that reaches the agent.
 *
 * `data` used to be the place a fact was supposed to go. It reaches nobody, so
 * it is serialized into the text instead.
 */
function renderDetails(data) {
    if (!data || typeof data !== "object") return null;
    if (Object.keys(data).length === 0) return null;

    let json;
    try {
        json = JSON.stringify(data);
    } catch {
        // A circular or otherwise unserializable `data` must not turn a legible
        // failure back into an illegible one: a throw here lands in the SDK's
        // catch, which is exactly the channel that discards everything.
        return "(details could not be serialized)";
    }
    if (typeof json !== "string") return null;
    if (json.length <= MAX_DETAIL_CHARS) return json;
    // Deliberately not truncated-and-pasted: half a JSON object is not JSON,
    // and an agent parsing it gets a second failure whose cause is us.
    return `(${json.length} characters of details, too large to include)`;
}

/**
 * The failure a tool caller actually receives.
 *
 * @param {string} toolName Named in the banner so a failure is attributable
 *                          when several calls are in flight.
 * @param {unknown} err     Anything a handler threw.
 */
export function toolFailure(toolName, err) {
    const code = err?.code ?? "word_error";
    const message = err?.message ?? (typeof err === "string" && err ? err : "Unknown error");
    const summary = `${code}: ${message}`;

    // First line is the failure banner, because nothing else in what the agent
    // receives distinguishes this from a successful result.
    const lines = [`${toolName} failed.`, summary];
    const details = renderDetails(err?.data);
    if (details) lines.push(`Details: ${details}`);

    return {
        textResultForLlm: lines.join("\n"),
        resultType: "failure",
        // Set because the SDK documents it and the runtime preserves it for
        // whatever reads it. Measured not to reach the agent, so nothing here
        // depends on it -- the text above carries the whole story on its own.
        error: summary,
    };
}
