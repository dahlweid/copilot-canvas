# Tool errors — findings

Written for issue #45. One question, eight arms, and one result that contradicts
the issue's own hypothesis.

## The question

Every failure from this extension reached the agent as:

```
Error: Tool execution failed
```

The code, the message and the `data` bag were all gone. #45 established that the
error is *intact* in our code — `asToolError` built a rich error and threw it —
so the loss is downstream. The issue proposed the fix that follows from that:
since structured fields are discarded, fold the code into the **message text**,
which is the one thing that survives.

That is a hypothesis about a mechanism, and this repo's most expensive recurring
failure is a right conclusion resting on a wrong mechanism, because nothing fails
when that happens. `asToolError` **already** folded the code into the message —
`new Error(\`${code}: ${message}\`)` — and the symptom outlived it. So the
hypothesis was already falsified by the shipped code before it was proposed.

The real question, which nobody had asked: **which parts of a thrown error
actually survive the crossing, measured rather than reasoned about?**

## The probe

`probes/error-channel-probe/extension.mjs`. A session-scoped extension
registering five tools and one canvas with three actions. Every arm carries a
distinctive marker — `MARK-A1-CODE`, `MARK-B-TEXT`, and so on — so a result can
be *attributed* rather than recognised: a marker missing from what the agent
receives is a marker that channel dropped.

Run live, inside a real Copilot CLI session, by installing it under the session's
`extensions/` directory, reloading extensions, and calling each tool and action
as the agent. This is the only way to observe the far end: the discard happens
in the host, so nothing beneath it can see the result.

**No Word.** Every arm returns or throws immediately; no Office process starts.

## What it measured

Copilot CLI **1.0.80**, Windows, `win32-x64`.

### Tools

| Arm | Handler | What the agent received |
| --- | --- | --- |
| A1 | throws `Error` + `.code` + `.data` | `Error: Tool execution failed` — **no marker at all** |
| A2 | throws a bare `Error` | `Error: Tool execution failed` |
| B | returns `{textResultForLlm, resultType:"failure", error}` | `MARK-B-TEXT: …` **verbatim**; `MARK-B-ERROR` absent |
| B2 | same, no `error` field | `MARK-B2-TEXT` verbatim |
| C | returns a plain object | `{"code":"MARK-C-CODE","message":"MARK-C-MESSAGE"}` verbatim |

### Canvas actions

| Arm | Handler | What the agent received |
| --- | --- | --- |
| D | throws `CanvasError("MARK-D-CODE", "MARK-D-MESSAGE")` | `Canvas operation failed: Error: MARK-D-MESSAGE` — message survives, **code does not** |
| E | throws a bare `Error` | `Canvas operation failed: Error: MARK-E-MESSAGE` |
| F | returns a failure object | echoed as JSON |

## The answer

**A thrown tool error carries nothing. The only channel that crosses is a
returned `ToolResultObject`'s `textResultForLlm` string.**

Two hops discard it, and either alone would be enough. Both are outside this
repo; `probes/probe-bridge-discard.mjs` re-checks that they still read this way.

1. **The SDK**, `copilot-sdk/extension.js`, `_executeToolAndRespond`:

   ```js
   } catch (error) {
     const message = error instanceof Error ? error.message : String(error);
     await this.rpc.tools.handlePendingToolCall({ requestId, error: message });
   }
   ```

   `code`, `data`, `name` and `stack` end here. This is **one hop earlier than
   #45 states**, which matters: a fix aimed only at hop 2 would still lose the
   code.

2. **The runtime**, `%LOCALAPPDATA%\copilot\pkg\win32-x64\1.0.80\sdk\index.js`,
   `tools_handle_pending_tool_call`. Minified; reformatted, it reads:

   ```js
   let requestId = …, error = n.error, result = n.result;
   return error
     ? (S.warning(`External tool call ${requestId} failed: ${error}`),
        { success: this.rejectExternalTool(requestId, new Error("Tool execution failed")) })
     : …
   ```

   The message is logged at *warning* level and then thrown away; the rejection
   is built from a string literal. **The message ends here** — which is why
   folding the code into it changed nothing.

The same site is present in every CLI runtime installed on this machine —
1.0.71, 1.0.73, 1.0.78-2, 1.0.79-5, 1.0.79-9 and 1.0.80 — so this is not a
regression in one build. The minified identifiers differ between versions
(`rejectExternalTool(r,…)` in 1.0.80, `rejectExternalTool(e.requestId,…)` in
1.0.73), which is why the citation probe matches a pattern rather than a literal:
a literal taken from one version reports the discard as *absent* from another,
which would be a false all-clear on the exact question being asked.

Meanwhile the **result** path is preserved. The same SDK function, on the
success side, forwards a recognised `ToolResultObject` unchanged:

```js
if (rawResult == null)                  result = "";
else if (typeof rawResult === "string") result = rawResult;
else if (isToolResultObject(rawResult)) result = rawResult;
else                                    result = JSON.stringify(rawResult);
```

which is exactly what arms B, B2 and C measured from the far end.

## Confirmation after the fix

Same session, same CLI, through the real `office-canvas` tools and canvas:

```
read_document { path: "C:\definitely\not\here\nope.docx" }
  before:  Error: Tool execution failed
  after:   read_document failed.
           file_not_found: No such file: C:\definitely\not\here\nope.docx

read_document { path: "…\CONTEXT.md" }
  after:   read_document failed.
           unsupported_type: .md is not a Word document. Supported: .docx, .docm, .doc, .dotx, .rtf.

open_document (canvas action) { path: "CONTEXT.md" }
  before:  Canvas operation failed: Error: .md is not a Word document. …   (no code)
  after:   Canvas operation failed: Error: unsupported_type: .md is not a Word document. …
```

Incidental, measured on the way and recorded because it contradicts an
assumption a reader could reasonably carry over from #28: **canvas action input
*is* schema-validated by the host before dispatch.** Invoking `open_document`
with `{}` was refused with `Invalid input for action "open_document" … "path" is
a required property` and the handler never ran. #28 measured the opposite for
**tools** — `required`, `enum`, `type`, `minItems`, `maximum` and
`additionalProperties` were each violated and every one reached the handler. The
two surfaces differ; a tool handler is still the only defence it has.

## The handler is called with two arguments, not one

Found in review of the fix (PR #91, round 1), in the same function as hop 1 and
a few lines above it:

```js
const rawResult = await handler(args, {
  sessionId, toolCallId, toolName, arguments: args,
  availableTools, traceparent, tracestate
});
```

The first version of `reportingFailures` was written `async (args) =>
tool.handler(args)` and dropped that second argument for every registered tool
at once — including `traceparent`/`tracestate`, which are W3C trace context, so
the symptom would have been a missing span rather than an error.

Worth recording rather than quietly fixing, because of *why* it survived review
of a change whose entire subject was information dropped at a boundary. No tool
handler in the extension reads the context, so nothing was broken, no test could
go red, and the loss was invisible from inside a handler — the same shape as #45
itself, one level up. A wrapper is a boundary, and the arity of a boundary is
part of what crosses it.

The fix is variadic forwarding, which is also correct for an SDK that adds a
third argument later. `probes/probe-bridge-discard.mjs` now tracks this call
site alongside the two discards, because `src/tool-error.mjs` quotes it as its
reason.

## What this is *not*

Stated explicitly, because each is a claim this measurement does not license.

- **It does not show the failure is visible to the user as a failure.** Arm B set
  `resultType: "failure"` and the agent-facing result carried no failure framing
  of any kind — indistinguishable from a success. Whether the *UI* renders it
  differently is **unmeasured**. This is why the fix puts a `<tool> failed.`
  banner in the text itself rather than trusting the flag.
- **It does not show where the CLI's `S.warning` line goes.** #45 asks whether
  the real message reaches a log a user could read. No such line appeared in the
  extension log during any arm. Not found is not the same as not written.
- **It does not cover platforms other than `win32-x64`, or CLI versions other
  than those installed here.**
- **It says nothing about MCP tools**, which take a different path through the
  same runtime.

## What changed because of it

- `src/tool-error.mjs`: `asToolError` (which threw) replaced by `toolFailure`
  (which returns a `ToolResultObject`). The code, the message and the whole
  `data` bag are rendered into `textResultForLlm`.
- `src/tool-error.mjs`: `reportingFailures` wraps one tool so a throw becomes a
  returned failure, and `extension.mjs` applies it at the **registration site**
  to every tool. Per-handler `try/catch` was the previous shape and is what let
  one call site drift; a tool added later now inherits the legible channel by
  being registered at all. It forwards variadically — see the call-site section
  above for what the first version of it dropped.
- `extension.mjs`: canvas actions keep throwing, because measured, that works
  there — but the code is folded into the message, because measured, the `code`
  field does not cross.
- Several tests that asserted `wrapped.data?.x` "reaches the agent" were
  re-pointed at the text. They were green throughout the defect, which is the
  concrete cost of asserting one layer beneath the boundary.
