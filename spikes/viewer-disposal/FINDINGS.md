# Viewer disposal: what a closed panel goes on doing

Issue: [#81](https://github.com/dahlweid/copilot-canvas/issues/81).
Probe: `probes/probe-close-during-retry.mjs`, with `probes/probe-copy-hammer.ps1`
and `probes/probe-strict-lock.ps1`.

## The question

`ViewerInstance.close()` does not cancel an in-flight `#autoRefresh()`. #80
measured that continuation and recorded it **inert**. That measurement used a
**fake** `RenderCache`.

Two things were therefore open:

1. Is the continuation actually inert, or was inertness a property of the fake?
2. Does it explain the user-visible failure — a panel re-opened just after an
   extension reload failing with *"Another process is holding demo.docx open
   more strictly than Word does"*, then succeeding unaided ~8 s later?

The suspected chain for (2): reload → `close()` → the loop continues for the
~6 s that `AUTO_REFRESH_DELAYS_MS = [500, 1500, 4000]` spans → `refresh()` →
`RenderCache` → `WordHost` **opens the document in Word** → the new panel's open
fails on that lock.

## Why a fake cache could not answer either question

`RenderCache.refresh()` reaches `WordHost`, which starts a PowerShell host that
drives Word over COM and opens the document. A fake has no file handle and no
process. It cannot open anything, so "nothing happened" was true of the fake and
said nothing about the real path. The consequence that mattered was invisible by
construction — not overlooked.

So the probe uses the **real** `RenderCache`, the real `WordHost`, real Word, the
real `FileWatcher` and a real generated `.docx`.

## What is injected, and what is not

Exactly one thing: `cache.refresh` is made to throw `file_locked` for its **first
two** calls, to arm the retry loop deterministically instead of depending on
winning a race. The third attempt — the one that lands ~2 s *after* `close()` —
runs the genuine `RenderCache.refresh`. **The decisive observation is made on
unmodified code.**

`WordHost.openDocument` is wrapped to count and timestamp entries; the wrapper
delegates to the real method.

The trigger is `utimes`, not a write. The watcher fingerprints `mtimeMs|size` and
the render key is `hash(path|mtimeMs|size)`, so moving mtime is a change by both
definitions and guarantees `RenderCache.open` cannot short-circuit, while leaving
the bytes a valid `.docx`. Deleting the file does not work as a trigger:
`FileWatcher.#settle` returns early on `previous === null`.

Arm 1 asserts the invariant we **want** (`0, 0`), not the behaviour we found. It
therefore fails on `c214e3f` as evidence and passes after the fix as a regression
witness. A probe asserting the defect would go green today and red the day
someone fixed it.

## Measured — before the fix, on `c214e3f`

```
WINWORD.EXE before: 14 (5132, 15296, 17084, 17444, 20808, 25280, 36160, 38884, 47628, 51640, 55528, 61388, 65448, 69544)
generating fixture...
        [cache] render: exported demo.docx in 7733ms
        [viewer] document changed on disk: demo.docx
        [viewer] auto-refresh failed (file_locked); retrying in 500ms
        [viewer] auto-refresh failed (file_locked); retrying in 1500ms
        [cache] render: exported demo.docx in 1147ms
  FAIL  nothing reaches cache.refresh or WordHost after close() returns
        work landed on a disposed viewer after close() returned
+ actual - expected

  {
+   refreshesAfterClose: 2,
+   wordHostOpensAfterClose: 1
-   refreshesAfterClose: 0,
-   wordHostOpensAfterClose: 0
  }

  ok    a WordHost open of a document does NOT make it refuse a concurrent Copy-Item
        3 WordHost opens of the original during the window
        Copy-Item: 137 attempts, 0 refused
  ok    positive control: the same counter does count a genuinely strict holder
        Copy-Item: 77 attempts, 43 refused (System.IO.IOException)
WINWORD.EXE after: 14 (5132, 15296, 17084, 17444, 20808, 25280, 36160, 38884, 47628, 51640, 55528, 61388, 65448, 69544)
  ok    this probe left no Word process behind
        host reported owning: 64872

4 arm(s), 1 failed
```

Run three times. Arm 1 gave `2` and `1` every time.

## Measured — after the fix

```
WINWORD.EXE before: 14 (5132, 15296, 17084, 17444, 20808, 25280, 36160, 38884, 47628, 51640, 55528, 61388, 65448, 69544)
generating fixture...
        [cache] render: exported demo.docx in 7437ms
        [viewer] document changed on disk: demo.docx
        [viewer] auto-refresh failed (file_locked); retrying in 500ms
        [viewer] watch: change was not consumed (The canvas was closed while a refresh was in flight.); it stays pending
  ok    nothing reaches cache.refresh or WordHost after close() returns
        close() returned in 1ms
        cache.refresh: 1 before close, 0 after it
        WordHost.openDocument: 1 before close, 0 after it
        viewer.status after the window: ready
        scripted failures left unconsumed: 1
  ok    a WordHost open of a document does NOT make it refuse a concurrent Copy-Item
        3 WordHost opens of the original during the window
        Copy-Item: 140 attempts, 0 refused
  ok    positive control: the same counter does count a genuinely strict holder
        Copy-Item: 82 attempts, 50 refused (System.IO.IOException)
WINWORD.EXE after: 14 (5132, 15296, 17084, 17444, 20808, 25280, 36160, 38884, 47628, 51640, 55528, 61388, 65448, 69544)
  ok    this probe left no Word process behind
        host reported owning: 63336

4 arm(s), 0 failed
```

## Conclusions

1. **The bug is real, and worse than #80 recorded.** After `close()` returned,
   **2 `cache.refresh` calls and 1 `WordHost.openDocument`** landed on a disposed
   instance, followed by a **1147 ms PDF export** (924 ms on an earlier run). A
   panel that no longer exists was opening a document in Word and exporting it.
   "Inert" was an artefact of the instrument.

2. **The route #80 predicted does not exist.** It attributed inertness partly to
   `refresh()` short-circuiting on `not_open`. `close()` never nulls `this.doc`,
   so that guard is never reached. Whatever inertness there was rested entirely
   on `#clients.clear()` making `#broadcast` write to an empty set.

3. **The lock hypothesis is refuted.** During 3 real `WordHost` opens of the
   document, a `Copy-Item` hammer made **137 attempts and was refused 0 times**.
   `Copy-Item` is not a proxy for the failing call — it *is* the call in
   `word-host.ps1`'s `Open-Doc` that raises "more strictly than Word does".

4. **The mechanism for (3) is in our own measured table.** `Open-Doc` copies the
   original (requesting read, granting `ReadWrite`) and opens **the copy**. It
   never holds the original strictly, so a continuation entering `WordHost`
   cannot produce the error the user saw. The hypothesis contradicted a table
   already in this repo.

5. **The zero in (3) is evidence of absence, not absence of evidence**, and only
   because of arm 3. The same hammer against the same file, with a
   `FileShare::None` handle held for 3 s, was refused **43 of 77** — every
   refusal a `System.IO.IOException`, the exact type `word-host.ps1` maps to
   `file_locked`. A hammer that is never refused proves nothing until it has been
   shown it *can* be refused.

6. **This bug does not explain the stranded `WINWORD.EXE` processes either.**
   `14 → 14` on every run, with the probe's own owned pid reaped. The
   continuation reused the live hidden Word rather than starting one.

7. **After the fix, arm 1 reads `0` and `0`**, `viewer.status` is `ready` rather
   than `error`, and one scripted failure is left unconsumed — the loop had work
   remaining and did not take it. The watcher logs the change as *not consumed*,
   so it stays pending rather than being recorded as handled.

## What this is not

- **It is not an explanation of the user's 8-second recovery.** Conclusions 3–5
  remove this bug as the cause. What did cause it is unmeasured, and nothing here
  bears on it.
- **It is not a measurement of the backoff's depth.** Nothing here says whether
  three attempts or four is right. The delays are used as shipped because the
  question was about the real ~6 s window.
- **It is not a claim that `close()` cancels work already inside Word.** The fix
  stops new work and refuses results. One already-started `cache.refresh` may
  still complete; see the PR for why awaiting it would be worse.
- **Arm 2 is a statement about `Open-Doc` as written**, not about Word in
  general. A future `Open-Doc` that opened the original directly would need this
  re-measured; arm 3 is the control that would then matter.
- **The `WINWORD.EXE` counts are a leak check, not a census.** 14 pre-existing
  processes were present throughout and are out of scope here.
