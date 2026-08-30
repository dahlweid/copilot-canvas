# Viewer connection — findings

Written for issue #66. One probe, one question.

## The question

`src/ui/app.js` announced `"Reconnecting…"` whenever the `EventSource` for
`/events` was in `CONNECTING`. The issue's claim is that this status can never
stop being wrong, because the render server's port is allocated per extension
process — so after a restart that origin is gone permanently, `EventSource`
retries it forever, and `readyState` never leaves `CONNECTING`.

That was **read out of the specification**, and separately **observed from the
outside**: a panel sat in that state for ~90 minutes against port 57318 while the
extension process was listening on zero sockets. Neither of those drives the
mechanism. The specification is not the implementation, and an outside
observation of a stuck panel does not show *which* of the states it was stuck in.

So: does a real `EventSource`, pointed at a really-closed `ViewerInstance`, keep
retrying and keep `readyState` at `CONNECTING` indefinitely?

## The probe

`probes/probe-dead-server.mjs`. Starts a real `ViewerInstance`, connects a real
`EventSource` (Node's, behind `--experimental-eventsource`), waits for it to
attach, then calls `instance.close()` — the same call a normal panel close makes
— and samples `readyState`, the error count and the status text once a second for
`DEADLINE_MS + 8s`.

It runs twice: once through the handler as it shipped before #66 (copied
verbatim), once through `monitorConnection`. Two arms, so the output is a
comparison rather than a demonstration.

**No Word.** `ViewerInstance.start()` only listens, and the single cache touch on
the SSE path is `state.wordVersion`, answered by a stub resolver. No document is
opened, so nothing starts an Office process.

**Attachment is read from the client end** — `onopen` having fired. This matters:
while #66 was being investigated, a count of established TCP connections was used
to conclude that a panel had detached, and that instrument turned out to be
worthless, because a *freshly opened, healthy, rendering* panel also showed zero
connections at 6 s and again at 26 s. The count cannot separate "detached" from
"never attached". `onopen` can, because the probe is the client.

## What it measured

Run on Node v24.18.0, Windows.

```
=== shipped (port 49153) ===
attached before close : true (readyState OPEN)
  t+ 1s  CONNECTING errors=1   says: (nothing)
  t+ 2s  CONNECTING errors=1   says: Reconnecting…
  ...
  t+18s  CONNECTING errors=7   says: Reconnecting…
final: {"text":"Reconnecting…","busy":true}

=== fixed (port 62038) ===
attached before close : true (readyState OPEN)
  t+ 2s  CONNECTING errors=1   says: Reconnecting…
  t+ 9s  CONNECTING errors=4   says: Reconnecting…
  t+10s  CLOSED     errors=4   says: Connection lost — this panel has stopped updating and cannot resume.
  ...
  t+18s  CLOSED     errors=4   says: Connection lost — this panel has stopped updating and cannot resume.
final: {"text":"Connection lost — this panel has stopped updating and cannot resume.","error":true}
```

1. **The issue's mechanism holds.** Against a closed server the connection sits
   in `CONNECTING`, not `CLOSED`, and errors keep arriving — 7 of them in 18 s,
   at no decreasing rate. There is no point at which the client gives up on its
   own. The `CONNECTING` branch is therefore permanently true and the status
   keyed on it permanently shown.
2. **The old handler never stops claiming a recovery.** It was still saying
   `"Reconnecting…"`, `busy: true`, at the last sample. Nothing in it can end
   that; the 90-minute observation is this, continued.
3. **The fix terminates, and does so by its own act.** At t+10s — `DEADLINE_MS`
   — the state is `CLOSED`, which the probe never asked for: the monitor called
   `source.close()` itself. The error count then stops rising, because there is
   nothing left retrying. That is why the wording is allowed to say *cannot
   resume*: it is a report of what the panel did to itself, not a prediction
   about a server the panel cannot see.
4. **The status is `error`, not `busy`.** No spinner over a dead panel.

The probe asserts all four and exits non-zero otherwise, so it can fail.

## What this is not

- **Not a measurement of the webview.** Node's `EventSource` is undici's, not
  Chromium's. Same specification, different implementation. This is strong
  evidence about *the mechanism the fix depends on*; it is not evidence about
  what the panel does in the product.
- **Not reachable from the unit suite.** CI runs `node --test` on Node 22 with
  no flags (`validate.yml`), and `EventSource` is flag-gated. A test needing a
  flag CI does not pass is a test CI does not run. Hence a spike.
- **Not evidence about recovery.** The probe shows the panel stopping cleanly.
  What, if anything, restores a stopped panel is untested — which is exactly why
  the shipped message names no remedy.
- **Not evidence about `ViewerInstance` reuse (#61).** This never exercises the
  instance map; it constructs an instance directly and closes it.

## The gap that remains

Nothing in this repo executes `src/ui/app.js`. Measured:

```
await import('./src/ui/app.js')
→ Error: Cannot find module 'C:\vendor\pdf.min.mjs'
  imported from …\src\ui\pdf-view.mjs
```

It dies at module resolution of pdf.js's absolute `/vendor/` specifier, before
any DOM access. Closing that link needs a loader stub plus a DOM stub, or a
browser-driven harness — and Chrome is not installed on this machine. So the
monitor's behaviour is measured here, `app.js`'s delegation to it is checked as
source text in `ui-contract.test.mjs`, and the webview itself is exercised by
nothing. Stated rather than papered over.
