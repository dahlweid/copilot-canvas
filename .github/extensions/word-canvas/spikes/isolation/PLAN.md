# Live Office Streaming — Build Plan

**Status:** proposal, pending rubber-duck review
**Supersedes:** the verdict in `spikes/live-word/FINDINGS.md` ("PDF for v1, streaming is the path to editing")
**Scope:** Word first, Excel and PowerPoint by design

---

## 1. Why this plan exists

v1 (shipped, PR #1) renders a `.docx` by driving a hidden Word instance to export a PDF and
letting the canvas' native PDF viewer display it. It is page-accurate and robust, but it is
fundamentally a **snapshot pipeline**: every change costs a full re-export, the viewer's
scroll position is unreadable, and there is no path to editing that doesn't go through
"export again".

The alternative is to stream the pixels of a **live** Word window into the canvas. That was
spiked and rejected, primarily because a live Word window is a real window: it appears in the
taskbar and Alt+Tab, it can steal focus, and the user can click into it.

**The rejection no longer holds.** Direct measurement on this machine shows a Word instance
can be made completely invisible to the user *and still render and capture*, and can be driven
entirely through COM across that isolation boundary. That changes the trade-off enough to
justify rebuilding the renderer.

> **Read §11 and §12 first.** A third option — keep the static, colour-correct, accessible PDF
> as the *display* path and use the hidden live instance purely as an *edit* channel — was
> measured after this plan was drafted and **wins**. Single-page re-export costs 168 ms and is
> flat in page position, so it hides behind an optimistic overlay. An independent critique
> (§12) reached the same conclusion, and re-measurement showed §6 below is wrong. §11
> supersedes phases 3 onward; §4 (isolation) and §5 (COM edit channel) survive, because
> option C needs both.

---

## 2. What was measured (and what was wrong before)

Everything in this section is a direct observation on this machine (Windows 11, Microsoft 365
Word 16.0 German UI, 125% display scaling), not inference.

### 2.1 The decisive result: a non-input desktop still renders

Create a second **desktop object** with `CreateDesktop`, launch `WINWORD.EXE` with
`STARTUPINFO.lpDesktop` pointing at it, and **never call `SwitchDesktop`**. The window exists
on a desktop that is never shown.

| Observation | Result |
| --- | --- |
| Window found via `EnumDesktopWindows` | `OpusApp visible=True 3072x1240 'desktop-probe.docx - Word'` |
| `PrintWindow(PW_RENDERFULLCONTENT)` from the *default* desktop | `ok=True brightness=41.2 colours=304` — real render |
| Visual inspection of the saved JPEG | Full Word window: ribbon, document text, spell-check squiggles |
| Live re-render proof: change `View.Zoom` 112 → 77 via COM, recapture | New capture shows the new layout — **not a stale frame** |
| Sustained capture rate on that desktop | **61.6 ms/frame** (~16 fps) incl. JPEG encode |
| Same window on the input desktop (spike measurement) | 34.3 ms/frame (~29 fps) — ⚠️ **not comparable to the row above**; different region, dimensions and JPEG settings. Withdrawn, see §12.1 |
| `Documents.Open` via COM across the desktop boundary | **OK in 359 ms**, `Content.Information(4)` → 13 pages |
| `Window.GetPoint` on a **paragraph** range | `x=193 y=529 w=811 h=78` — a valid rectangle, but these are paragraph bounds, **not** a caret. Corrected in §12.2, where a collapsed caret measures `w=0` with per-glyph `x` precision |
| `Window.RangeFromPoint` at 3 different screen points (click-to-caret) | **`start=22`, `255`, `401`** — distinct, correctly ordered document positions |
| Full edit round trip: `RangeFromPoint` → `Select` → `TypeText` → `Undo` | **All OK** |

And the user-visible consequence, measured directly rather than assumed
(`probes/probe-visibility.ps1`):

| Check, run from the default desktop while Word runs on the hidden one | Result |
| --- | --- |
| Windows belonging to that PID on the hidden desktop | 18 (Word is fully there) |
| Windows that `EnumWindows` can see | **0** |
| Alt+Tab candidates (visible, unowned, not `WS_EX_TOOLWINDOW`) | **0** |
| `Shell.Application.Windows().Count` (the shell's own view) | **0** |
| `Process.MainWindowTitle` | **empty** — even .NET cannot find a main window |
| Process alive and holding the document | yes |

A window on a non-input desktop object has **no taskbar button, no Alt+Tab entry, no virtual
desktop presence, and cannot receive user input** — the shell (`explorer.exe`) runs on the
`Default` desktop and only enumerates windows on its own desktop. This is total isolation as
a *side effect of where the window lives*, not as a set of style tricks we have to maintain.

### 2.2 Cross-process `WS_EX_TOOLWINDOW` also works (fallback)

From a non-elevated process, on another process's `OpusApp` HWND: ex-style `0x00000110` →
`0x00000190` (TOOLWINDOW set, APPWINDOW cleared), `IsWindowVisible` still true, capture still
real (`brightness=45.6 colours=194`), and Word still responsive to COM afterwards. The
documented cross-process `SetWindowLong` restriction applies to `GWL_WNDPROC`/`DWL_DLGPROC`
(subclassing), not `GWL_EXSTYLE`.

The change must be bracketed `ShowWindow(SW_HIDE)` → `SetWindowLong` → `ShowWindow(SW_SHOWNOACTIVATE)`,
because taskbar/Alt+Tab membership is only re-evaluated on re-show.

### 2.3 Binding COM to *our* instance, never the user's

`AccessibleObjectFromWindow(hwnd_of__WwG, OBJID_NATIVEOM, IID_IDispatch)` returns the
`Word.Window` of **that specific process**, from which `.Application` gives that process's
object model. Measured: `Documents.Count == 1` (only our document).

This is a significant safety improvement over v1. `New-Object -ComObject Word.Application`
goes through the Running Object Table and may hand back the *user's* Word. The HWND-based
bind cannot: we launched the process, we know its PID, we find its window on our own desktop,
we bind to that. **The user's Word is unreachable by construction.**

### 2.4 Two prior claims that were wrong — and why that matters

1. **`FINDINGS.md` says the dark page was "an artefact of the zoom bug". It is not.** With the
   fit corrected, the page still renders dark (measured `pageBrightness ≈ 58–60`, and visually
   confirmed: dark page, light text). The cause is Office dark mode following the user's
   `UI Theme = 4` (Black). This is a **real fidelity defect of the streaming approach** and is
   tracked as open risk R1 below. `FINDINGS.md` must be corrected.

2. **The isolation research agent confidently asserted both §2.1 and §2.2 were impossible** —
   "kernel-enforced `ERROR_ACCESS_DENIED`", "DWM does not composite non-input desktops, capture
   returns black, do not pursue it" — with citations and a ranked recommendation table. Both
   are false on this machine. The report also warned that COM across desktops would hang on
   DDE paths; `Documents.Open` completed in 359 ms.

   The lesson is a process one and belongs in the plan: **for this problem class, plausible
   secondary sources are not evidence.** Every load-bearing platform claim in this plan must
   be backed by a probe in `spikes/`, and the build phases below are sequenced so that the
   riskiest claims are re-verified first, on more than one machine.

---

## 3. Architecture

```mermaid
flowchart TB
    subgraph hidden["Isolated desktop object 'CopilotOfficeRender' (never switched to)"]
        W[WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE<br/>launched with STARTUPINFO.lpDesktop]
    end
    subgraph host["Extension process"]
        H[office-host.ps1<br/>capture loop + COM driver]
        E[extension.mjs<br/>canvas + actions]
        S[server.mjs<br/>loopback HTTP]
    end
    C[Canvas iframe<br/>img + overlay]

    H -- CreateProcess/lpDesktop --> W
    H -- "AccessibleObjectFromWindow<br/>OBJID_NATIVEOM (COM, cross-desktop)" --> W
    W -- "PrintWindow PW_RENDERFULLCONTENT" --> H
    H -- "frames + geometry (JSON/stdio)" --> E
    E --> S
    S -- "MJPEG / SSE" --> C
    C -- "input intents (HTTP POST)" --> S
    S --> E
    E -- "COM commands" --> H
```

**The rule that keeps this honest: pixels flow out, intents flow in.** We never inject
Win32 input. Everything the user does in the canvas is translated into an object-model call.
That is what makes the isolated desktop viable (input can't cross it anyway), what makes
editing safe and undoable, and what generalises to Excel and PowerPoint.

### 3.1 Component layout

```
.github/extensions/word-canvas/          # renamed later to office-canvas
  extension.mjs
  src/
    isolation/
      desktop.ps1          # CreateDesktop, CreateProcess-on-desktop, teardown
      bind.ps1             # AccessibleObjectFromWindow -> app object model
    capture/
      capture.ps1          # PrintWindow loop, crop, encode, dirty-detect
    apps/
      app-word.ps1         # Word adapter: geometry, scroll, hit-test, edit verbs
      app-excel.ps1        # later
      app-ppt.ps1          # later
    office-host.ps1        # command loop; composes the above
    office-host.mjs        # Node bridge (spawn, correlate, restart, reap)
    server.mjs             # loopback: /stream, /api/*, /events
    ui/                    # <img> + caret/selection overlay + chrome
  spikes/
    live-word/             # existing spike (FINDINGS.md to be corrected)
    isolation/             # the probes from this investigation, promoted
```

The **app adapter** is the generalisation seam. Everything above it — desktop isolation,
capture, transport, overlay, canvas wiring — is app-agnostic. An adapter answers a fixed
question set: *what is the content window; what is the page/sheet/slide geometry; how do I
scroll; how do I map a point to a document position; how do I apply an edit; how do I know
something changed.*

---

## 4. Isolation strategy (ranked, with fallbacks)

| Rank | Strategy | Isolation | Renders? | Evidence |
| --- | --- | --- | --- | --- |
| **1** | Separate **desktop object**, never switched to | Total: no taskbar, no Alt+Tab, no virtual desktop, user input cannot reach it | **Yes**, ~16 fps | §2.1, measured |
| 2 | `WS_EX_TOOLWINDOW` + park off-screen at x=30000 | No taskbar/Alt+Tab; window still on the user's desktop and technically reachable | Yes, ~29 fps | §2.2, measured |
| 3 | Off-screen park only | Taskbar + Alt+Tab still show Word | Yes | measured |
| ✗ | `ShowWindow(SW_HIDE)` or `Application.Visible=false` | Total | **No — pure black** (`brightness=0.0 colours=1`) | measured |
| ✗ | Inactive **virtual** desktop | Total | Expected black (DWM cloaks); **not yet probed** | — |
| ✗ | Session 0 / service | Total | Prohibited by Microsoft policy (KB257757) | policy |

Ship **1 with automatic fallback to 2**. Both are already proven to render; the fallback
matters because rank 1 is the one most likely to behave differently on other hardware,
under RDP, or on a locked workstation (open question Q1).

Note the two mechanisms the user's UAC analogy conflates: the **secure desktop** (a desktop
object, which is what we're using, and which isolates by *location*) and **UIPI/integrity
levels** (which block lower-integrity processes from posting messages to higher-integrity
windows). UIPI is not useful here — it would work *against* us if we needed input injection.
We don't, which is why the desktop-object half of the analogy is the half that pays off.

---

## 5. Read-write design

**Yes, this approach enables read-write** — and it does so more safely than input injection
would. **This is measured, not designed on paper:** on the hidden desktop, `Window.GetPoint`
returned a valid rectangle (a *collapsed* caret measures `w=0` — see §12.2),
`Window.RangeFromPoint` mapped three different screen points
to three correct and correctly-ordered document positions, and a full
`RangeFromPoint → Select → TypeText → Undo` round trip completed successfully (§2.1).

That matters because the documentation warns `RangeFromPoint` needs *screen* coordinates of a
*currently rendered* region, and it was not obvious that a desktop nobody looks at would
satisfy either condition. It does: the hidden desktop has its own coordinate space, the `_WwG`
content window reported a normal rect (`L=7 T=178 R=976 B=1171`), and hit-testing against it
behaves exactly as it would on the visible desktop.

### 5.1 Why not input injection

`SendInput` targets the input desktop and cannot cross to our hidden desktop.
`PostMessage(WM_CHAR/WM_KEYDOWN)` can cross a desktop boundary only in limited ways and is
unreliable against Office's input handling (IME, autocorrect, and ribbon state all assume a
real focus/keyboard state). Injection also makes every keystroke a separate undo record and
gives us no way to know what actually happened.

### 5.2 The intent channel

The canvas captures user interaction and sends **document-level intents**:

| Canvas gesture | Intent | Word object model |
| --- | --- | --- |
| Click in page | `place-caret {x, y, page}` | `Window.RangeFromPoint(x, y)` → `Range.Select()` |
| Drag | `select-range {from, to}` | two `RangeFromPoint` calls → `Document.Range(a,b).Select()` |
| Type | `insert-text {text}` | `Selection.TypeText(text)` |
| Backspace / Delete | `delete {direction, unit}` | `Selection.Delete(unit, count)` |
| Enter | `insert-paragraph` | `Selection.TypeParagraph()` |
| Ctrl+B / I / U | `toggle-format {bold\|italic\|underline}` | `Selection.Font.Bold = -1` etc. |
| Ctrl+Z / Y | `undo` / `redo` | `Document.Undo()` / `Document.Redo()` |
| Scroll | `scroll {delta}` | `Window.SmallScroll` |
| Paste | `paste {text}` | `Selection.TypeText` (plain) — never the shared clipboard |

Coordinate mapping in both directions is available: `Window.RangeFromPoint(x, y)` maps a
screen point to a document position, and `Window.GetPoint(out x, out y, range)` maps a range
back to a screen rectangle. Since we know the window's screen rect and our crop offset, canvas
coordinates convert to window coordinates with a scale factor we already compute for DPI.

### 5.3 Caret and selection are drawn by us

Word's own caret will not blink reliably in a non-focused, non-input-desktop window, and a
captured caret would lag the keystroke by a frame. So the canvas **draws its own caret and
selection highlight** in an overlay, positioned from `Window.GetPoint`. This also makes typing
feel instant: the overlay advances optimistically on keypress, and the next frame confirms it.

Two constraints measured in §12.2 apply here:

- A **collapsed** range yields a true caret (`w=0`, per-glyph `x`), so caret placement is exact.
- A **multi-line** selection yields only a *bounding box*, not the selection's real shape. The
  overlay must therefore walk the selection line by line — one `GetPoint` per line — and draw a
  rectangle list. This is ordinary work, but it is not free and must be budgeted.

### 5.4 Batching and undo granularity

Typed characters are coalesced into `TypeText` calls on a short debounce (~120 ms) so a
sentence is a handful of COM calls, not one per character, and undo granularity stays close to
what a user expects. Every intent is applied to the **temp working copy**, exactly as v1 does;
"save" is an explicit action that copies back to the original path.

### 5.5 What read-write does *not* get us

IME composition, autocorrect-as-you-type, spell-check context menus, drag-and-drop, and
ribbon-driven commands are not reachable through this channel without significant extra work.
The honest framing is **"structured editing"**, not "Word in a panel". The plan targets: text
entry, deletion, selection, basic character/paragraph formatting, undo/redo, and find/replace.

---

## 6. Change notification — SUPERSEDED, this section is wrong

> **Do not implement this section.** Measurement (§12.1) showed the events it relies on either
> do not exist or do not fire in a headless host. The correct design is a polled `doc.Saved`
> dirty flag at 1.14 ms per read. Kept below only to record what was tried.

Capturing at 16–30 fps unconditionally is wasteful for a document that changes rarely. The
pipeline is demand-driven:

1. **Intent-triggered** — any intent we send is followed by a capture.
2. **Event-triggered** — Word's `Application` events (`WindowSelectionChange`,
   `DocumentChange`, `WindowScroll`) fire over COM and can be sinked from the host, which
   requests a frame.
3. **External-change-triggered** — the existing `watcher.mjs` (fs.watch + 300 ms debounce +
   settle polling, already shipped and proven to live-reload a regenerated document) still
   drives reload when a script rewrites the file.
4. **Idle** — no frames at all. A cheap dirty check (capture at low frequency, compare a hash
   of a downsampled frame) guards against missed events without a full-rate loop.

---

## 7. Generalising to Excel and PowerPoint

The isolation, capture, transport and overlay layers are app-agnostic. Each app supplies an
adapter. What differs:

| Concern | Word | Excel | PowerPoint |
| --- | --- | --- | --- |
| Content window class | `_WwG` | `EXCEL7` (child of `XLDESK`) | `mdiClass` / slide pane |
| Unit of navigation | page | sheet + visible range | slide |
| Geometry query | `Window.GetPoint`, `Information(wdActiveEndPageNumber)` | `Range.Left/Top/Width/Height`, `Window.VisibleRange` | `Shape.Left/Top`, `Slide.SlideIndex` |
| Point → position | `RangeFromPoint` | `Window.RangeFromPoint` | `Slide.Shapes.Range` hit-test (weaker) |
| Scroll | `SmallScroll` | `Window.ScrollRow/ScrollColumn` | `View.GotoSlide` |
| Edit verb | `Selection.TypeText` | `Range.Value2 = …` | `TextFrame.TextRange.Text = …` |
| "Page-accurate" means | print layout | **not** a page concept — the grid is the truth | slide canvas, fixed aspect |

Two consequences worth deciding up front:

- **Excel breaks the page metaphor.** The canvas chrome must be per-app, not one generic
  viewer with a page counter. Excel wants a sheet tab bar and a formula bar; PowerPoint wants
  a slide rail.
- **PowerPoint hit-testing is weaker** — there is no true `RangeFromPoint` equivalent for
  arbitrary points, so click-to-place-caret inside a text box needs a shape-bounds walk. Plan
  for PowerPoint editing to be shape-level first, text-level second.

Build order: **Word (read) → Word (write) → Excel (read) → PowerPoint (read) → write for
both.** Excel before PowerPoint because its object model hit-testing is stronger and it will
stress the adapter seam harder (no pages, virtualised grid), which is exactly what we want to
learn early.

---

## 8. Risks and open questions

| # | Risk | Impact | Mitigation / next step |
| --- | --- | --- | --- |
| **R1** | **Dark mode**: the streamed page follows the user's Office theme and renders dark. `ExportAsFixedFormat` (v1) never has this problem. | High — breaks "page-accurate" | Three candidates, in order: (a) find the per-user setting Word 365 uses for dark page and set it for our instance only — it is *not* `Common\UI Theme` (flipping that to 5 at launch did not change the render, measured), so it is likely a roamed cloud setting; (b) run our instance under a separate Windows user profile / isolated registry hive; (c) accept it and post-process the frame (invert luminance on the page region) — ugly and lossy. **Must be resolved before this replaces the PDF renderer.** |
| **R2** | Non-input-desktop rendering may not hold on other machines, under RDP, or when the workstation is locked | Fatal to rank-1 isolation | Partially closed — see the matrix below. **Locked workstation is the one that matters and is still untested.** Automatic fallback to rank 2 regardless |
| **R3** | 61.6 ms/frame on the hidden desktop is ~2× the input-desktop cost | Sluggish typing feel | Try `SetThreadDesktop` onto the hidden desktop before capturing (untried); crop to the content window; encode WebP/JPEG at lower quality; only capture on intent/event |
| **R4** | Empirical `0.92` fit factor in the current zoom maths is unexplained | Mis-scaled pages | Replace with page-rectangle detection: `Window.GetPoint` on a range at the top-left of the page body gives a real anchor; derive scale from that instead of guessing |
| **R5** | Office dialogs (activation, licensing, recovery, "document in use") appearing on a desktop nobody can see | Silent hang | Enumerate the hidden desktop's windows on a timer; any unexpected top-level window is detected, screenshotted into the log, and dismissed or surfaced to the user as an actionable error |
| **R6** | Orphaned `WINWORD.EXE` on a desktop that is never shown is *worse* than v1's orphans — invisible to the user | Resource leak, file locks | Reuse v1's PID-file reaping, plus close the desktop handle on exit; a desktop object is destroyed when its last window closes, which gives a second cleanup signal |
| **R7** | Supportability: Microsoft does not bless Office automation in unusual window/desktop contexts | Long-term fragility | Stay in an interactive user session (we do); never Session 0; keep rank-2 fallback which is an entirely ordinary window |
| **Q1** | ~~Does WGC work on a non-input desktop?~~ **Closed — WGC is unavailable to us regardless.** `GraphicsCaptureSession.IsBorderRequired = false` (the only way to suppress WGC's yellow capture border) requires a declared capability that is **only available to MSIX-packaged apps**. Our host is an unpackaged Node/PowerShell process, so the yellow border would be permanent and drawn around the captured content. `PrintWindow` draws no border. **`PrintWindow` stays.** | — | Revisit only if the host app is ever packaged |
| **Q2** | Does typing latency through the COM intent channel feel acceptable? | Decides whether write mode ships | Prototype gate in phase 4 — measure keypress → confirmed frame |
| **Q3** | `RangeFromPoint`/`GetPoint` need the target to be in the *rendered viewport*. Scrolling before hit-testing adds a COM round trip. | Latency spike on click-after-scroll | Keep a cached map of visible ranges per frame; only `ScrollIntoView` when the target is off-viewport |

### 8.1 Isolation robustness matrix (R2)

The single biggest threat to this plan is that all of §2 was measured on one machine in one
state. Progress so far:

| Variable | Status | Result |
| --- | --- | --- |
| Office **hardware acceleration ON** (the default; no `Graphics` registry key present) | **Tested** | `brightness=41.2 colours=304` — works |
| Office **hardware acceleration OFF** (`DisableHardwareAcceleration=1`) | **Tested** | `brightness=41.2 colours=304` — **identical**, works |
| **Workstation locked** | **Untested — highest priority** | Unknown. Research inference says the message-based render path *should* survive a lock because it does not depend on display output, but this is explicitly unverified and community reports of `PrintWindow` failing on locked workstations concern a different code path. **Must be tested manually; it is the most likely single point of failure in normal use, since users lock their machines constantly.** |
| Second machine / different GPU | Untested | — |
| RDP connected | Untested | Inference: likely works |
| RDP disconnected | Untested | Inference: high risk |
| Display sleep / power save | Untested | Related to the locked case |
| Different Office build / channel | Untested | — |

The acceleration result is the meaningful one so far: it was flagged as *the* critical
variable that could make the whole finding contingent on a user setting, and it turned out to
make no difference at all. That materially raises confidence in the mechanism. The locked-
workstation case replaces it as the top open risk.

---

## 9. Build phases

Each phase ends with a written, measured result. Phases 1–2 are pure de-risking: if they fail
we stop, and v1's PDF renderer remains the shipping renderer.

| Phase | Goal | Exit criterion |
| --- | --- | --- |
| **0. Correct the record** | Fix the wrong dark-page claim in `FINDINGS.md`; promote the four probes from `$TEMP` into `spikes/isolation/` | Findings match measurements; probes are reproducible from the repo |
| **1. Prove isolation holds** | Work the R2 matrix in §8.1: **locked workstation first** (hard gate), then a second machine, RDP connected + disconnected, display sleep | Locked-workstation behaviour known and handled. If a lock kills the render, the canvas must detect it and recover on unlock rather than showing a dead frame — that recovery path is part of the exit criterion, not a follow-up |
| **2. Solve dark mode (R1)** | Make the streamed page render in print colours without touching the user's settings | A capture whose page region is white on a machine with `UI Theme = 4` |
| **2b. Decide streaming vs option C** | Measure single-page `ExportAsFixedFormat` re-export latency (`From`/`To` page args) using v1's existing export path | **DONE — resolved in favour of option C.** 168 ms single-page, 183 ms edit-to-updated-page, flat in page position. See §11. Phases 3+ below describe the streaming build and are superseded unless option C fails |
| **3. Read-only streaming canvas** | Isolated desktop + capture + transport + crop/fit via page-rect detection (R4) + chrome | Side-by-side with the PDF canvas at equal fidelity; scroll stays put across reloads (the thing v1 cannot do) |
| **4. Write mode for Word** | Intent channel, overlay caret/selection, batched `TypeText`, undo/redo, explicit save | Type a paragraph, format it, undo it; measure keypress → frame (Q2) |
| **5. Excel adapter (read)** | Prove the adapter seam by making a second app work without touching the layers above it | Excel renders and scrolls with no changes to capture/transport/overlay |
| **6. PowerPoint adapter (read)** | Third app | As above |
| **7. Write for Excel/PowerPoint** | Cell edits, text-box edits | — |

**Kill criteria.** If phase 1 shows rank-1 isolation is machine-specific *and* rank-2's
taskbar suppression proves unreliable, or if phase 2 cannot solve dark mode without mutating
the user's Office configuration, streaming does not replace the PDF renderer. It could still
ship as an opt-in second canvas for editing, with PDF remaining the default viewer.

---

## 10. Relationship to the shipped v1

v1 stays. It is the fallback renderer when Office is an older build, when isolation fails, or
when the user prefers it — and it is the only renderer that gives free text selection, native
Ctrl+F and print. The streaming canvas is additive, and the two share the document lifecycle,
temp-copy safety, macro disabling, watcher, and process reaping that v1 already proved out.

---

## 11. The third option: static render + live edit channel (option C)

The plan above frames this as a binary — snapshot PDF (v1) versus pixel streaming. It is not.
The two capabilities that motivate streaming are **liveness** and **editing**, and only one of
them actually requires streaming pixels.

**Option C:** keep `ExportAsFixedFormat` → PDF as the *display* path, and use the hidden,
isolated live Word instance purely as the *edit and query* channel. Edits arrive as intents
over COM exactly as in §5; after an edit settles, re-export and refresh the PDF.

### What option C keeps that streaming throws away

| Property | Streaming | Option C |
| --- | --- | --- |
| Page colour correct under a dark Office theme | ✗ (risk R1, unsolved) | ✓ — `ExportAsFixedFormat` always emits print colours |
| Text selection and copy | ✗ — a bitmap has no text | ✓ native |
| Ctrl+F within the page | ✗ must be reimplemented via COM `Find` | ✓ native |
| Print | ✗ | ✓ native |
| Page thumbnails, zoom controls | ✗ must be rebuilt | ✓ native |
| **Screen-reader accessibility** | ✗ — an opaque bitmap | ✓ — tagged PDF text |
| Depends on non-input-desktop rendering (R2) | ✓ **fatal dependency** | ✗ — only needs COM, which is far better supported |
| Depends on capture frame rate (R3) | ✓ | ✗ |
| Depends on the 0.92 fit fudge (R4) | ✓ | ✗ |
| Editing via COM intents | ✓ | ✓ — **identical mechanism** |
| Live external-change reload | ✓ | ✓ — `watcher.mjs`, already shipped |
| Scroll position preserved across updates | ✓ inherently | ~ needs work: re-anchor the viewer after refresh |
| Update latency after an edit | ~35–60 ms | re-export cost |

The striking row is the dependency block. Option C **deletes risks R1, R2, R3 and R4 outright**
— including the two the plan itself calls the biggest threats — while keeping the editing
capability that was the entire reason to move away from v1. Everything measured in §2.1 about
COM working across the desktop boundary (`Documents.Open` in 359 ms, `RangeFromPoint`,
`GetPoint`, `TypeText`, `Undo`) applies unchanged to option C, because those are COM facts, not
rendering facts. Only the `PrintWindow` results become irrelevant.

### What option C costs

1. **Re-export latency after each edit.** This is the crux. v1 measured a full 13-page export
   at roughly 1–2 s, which is far too slow to sit between a keystroke and its echo.
   Mitigations, in order of promise:
   - `ExportAsFixedFormat` takes `From`/`To` page arguments, so re-export **only the edited
     page**, not the document. Needs measuring; likely an order of magnitude cheaper.
   - Echo typing **optimistically in the overlay** (§5.3 already draws its own caret and
     selection). The user sees their character immediately in the overlay layer; the
     re-exported page lands underneath a moment later. This is the same trick that makes
     streaming feel fast, applied to a static backing image.
   - Debounce the re-export to end-of-word or end-of-sentence rather than per keystroke.
2. **Scroll and zoom state across refresh.** Reloading the PDF resets the viewer. Needs an
   anchor-and-restore step (page number + offset), which streaming gets for free.
3. **It is not "Word in a panel".** Neither is streaming, given §5.5, so this is a smaller
   difference than it first appears.

### Honest assessment

**Phase 2b has now been measured** (`probes/probe-export.ps1`, 13-page document, this machine):

| Operation | Time |
| --- | --- |
| First export in a fresh Word process (one-off PDF-engine load) | 4547 ms |
| Full 13-page export | 664 ms |
| **Single page export, page 1** | **168 ms** |
| **Single page export, page 7** | **172 ms** — flat, does not scale with page position |
| **Edit + single-page re-export, end to end** | **183 ms mean, 168 ms min** |

Single-page re-export is roughly **4× cheaper than a full export and independent of where the
page sits in the document**, which is the property that matters: editing page 200 of a 300-page
document costs the same as editing page 1.

183 ms is comfortably inside what an optimistic overlay hides. The overlay draws the typed
character immediately (§5.3 already requires that machinery for streaming too), and the
re-exported page lands underneath within a fifth of a second. For comparison, streaming
updates in 61.6 ms — about 3× faster per update, but streaming must pay that cost on *every*
frame, whereas option C pays it only at debounce boundaries.

**The gate therefore resolves in favour of option C.** The plan's phases 3 onward should be
rewritten around static page rendering plus a COM edit channel, which:

- deletes risks **R1 (dark page), R2 (non-input-desktop rendering), R3 (frame rate) and
  R4 (the 0.92 fit fudge)** entirely;
- keeps every measured editing capability from §2.1, all of which are COM facts;
- preserves text selection, Ctrl+F, print, and **screen-reader accessibility**, none of which
  a bitmap can offer;
- reuses `render-cache.mjs` and `watcher.mjs` from the shipped v1 rather than replacing them.

What survives from the streaming research is the part that was actually load-bearing: the
**hidden-desktop isolation** (§4) and the **COM intent channel** (§5). Those are exactly what
option C needs. The pixel-capture work becomes a fallback for the one case option C cannot
serve — showing a live view of the user's *own* Word session — which is not a v2 requirement.

The remaining open risk for option C is **scroll and zoom restoration** across a PDF refresh
(§11 cost 2), which is ordinary engineering rather than a platform unknown.

### The one thing option C changes about the viewer

There is a tension in option C that has to be stated plainly, because it contradicts a v1
finding.

v1 concluded "no pdf.js needs to be vendored" — the canvas iframe renders PDF natively, giving
pagination, scroll, zoom, selection, Ctrl+F and print for free. That is correct **for read-only
display** and option C keeps it there.

It does not survive contact with editing. An edit mode needs three things the native PDF
viewer will not provide to the parent document:

1. the current scroll position, so a refresh can be restored rather than reset;
2. page geometry in client coordinates, so a click can be mapped to a document position via
   `RangeFromPoint` (§5.2);
3. the ability to replace **one** page in place without reloading the document and losing the
   viewer's state.

So option C splits the viewer in two:

| Mode | Renderer | Rationale |
| --- | --- | --- |
| Read-only display (v1 behaviour, unchanged) | native PDF viewer in an `<iframe>` | zero dependencies, everything free |
| Edit mode | **pdf.js**, canvas + text layer | we own scroll, geometry and per-page re-render |

pdf.js is the right choice rather than page bitmaps because it keeps a **text layer**, so
selection, in-page search and screen-reader accessibility survive into edit mode — which was
option C's main advantage over streaming in the first place. Rendering pages as plain images
would throw that away and land us back at streaming's weaknesses with none of its liveness.

**This is the one new dependency option C introduces**, and it should be treated as a real cost
in the phase plan: vendoring pdf.js, wiring its text layer to our overlay, and keeping the two
viewer modes from diverging. It is still a much smaller and better-understood cost than the
four platform risks option C removes.

---

## 12. Independent critique, and what measurement said about it

A rubber-duck critic reviewed §1–§10 (before §11 existed). Its headline verdict — *keep PDF as
the canonical read renderer, use a live hidden Word instance only as an edit channel* — is
**the same conclusion as option C**, reached independently. That convergence is the strongest
signal in this document.

Every disputed factual claim below was then re-measured rather than argued.

### 12.1 Confirmed by measurement — the plan was wrong

**Word has no content-change event.** (`probes/probe-events.ps1`)

| Event | Result |
| --- | --- |
| `WindowScroll` | **does not exist** |
| `DocumentChanged`, `ContentChange` | **do not exist** |
| `DocumentChange` | exists, but did **not** fire on a content edit |
| `WindowSelectionChange` | subscribed cleanly, did **not** fire on a programmatic edit or selection change |

§6's event-driven design was wrong. Note the honest caveat: a PowerShell host does not pump
window messages, so the two "did not fire" rows may reflect delivery rather than semantics.
**That distinction does not matter** — a headless host cannot depend on these either way.

**The replacement is a cheap polled dirty flag**, which measurement shows is entirely viable:

| Poll candidate | Cost per read | Verdict |
| --- | --- | --- |
| `doc.Saved` | **1.14 ms** | ✓ flips `True`→`False` on any edit, resettable, use this |
| `window.VerticalPercentScrolled` | 3.13 ms | ✓ cheap enough for scroll tracking |
| `doc.Content.Information(4)` (page count) | **36.5 ms** | ✗ far too expensive to poll; only after a settle |

At 10 Hz, `doc.Saved` costs about 1% of one core. §6 should be rewritten around this.

**Other critique findings accepted without further test**, because they are design gaps rather
than platform questions: no version/consistency protocol between pixels, geometry and intents
(a click on frame N must not edit state N+1); no revision-ownership or conflict policy for
write mode, where v1's "reopen from the original on crash" becomes a **data-loss path**; hidden
modal dialogs can deadlock the COM thread, so a separate watchdog and a kill-on-close **Job
Object** must replace PID-file cleanup; multi-document concurrency is undefined; the frame-rate
comparison in §2.1 (16 fps vs 29 fps) is **invalid** because the two measurements used
different regions, dimensions and JPEG settings, and is withdrawn.

### 12.2 Partly refuted by measurement

The critique said `probe-hittest.ps1` proved nothing about caret geometry because its `GetPoint`
call returned an 811×78 rectangle — a paragraph, not a caret. **It was right about my probe and
right that I reported it carelessly.** But the capability does exist; the earlier probe simply
measured the wrong thing. (`probes/probe-caret.ps1`)

| Range | `GetPoint` result | Reading |
| --- | --- | --- |
| collapsed at 120 | `x=4356 y=625 w=0 h=46` — **zero width** | a true zero-width caret |
| collapsed at 121 | `x=4367 y=625 w=0 h=46` | **11 px apart — per-glyph precision** |
| span 120–140 (one line) | `x=4356 w=213 h=46` | correct single-line extent |
| span 120–400 (multi-line) | `x=4224 w=811 h=235` | **a bounding box — critique correct** |

And the round trip is exact: caret at offset 120 → screen point → `RangeFromPoint` → **offset
120, zero characters of error**.

So: collapsed-caret placement and click-to-position are sound and precise. **Multiline
selection genuinely cannot be expressed as one rectangle** — the overlay must build a
per-line rectangle list, which is ordinary work but must be designed rather than assumed.

An API-shape correction worth recording: **`GetPoint` is a method of `Window`, not `Range`.**
Calling it on a `Range` fails with `DISP_E_UNKNOWNNAME`.

### 12.3 Where the critique and option C differ

The critique proposes streaming during an edit session. §11 measured single-page re-export at
**168 ms, flat in page position**, which the critique did not have. That number makes
streaming unnecessary even for editing: an optimistic overlay plus a 168 ms page refresh
delivers editing without inheriting the dark-page, non-input-desktop, frame-rate and page-fit
risks. Option C is therefore the stronger form of the critique's own recommendation.

### 12.4 Consequences for the plan

1. §6 (change notification) — **rewrite** around polled `doc.Saved`, not events.
2. §2.1 — **withdraw** the frame-rate comparison; **correct** the caret claim.
3. Add a **version/consistency protocol** to §5 before any coordinate intent is implemented.
4. Add a **revision-ownership and conflict policy** before any write path is implemented.
5. Replace PID-file cleanup with a **Job Object**; add a **dialog watchdog** in a separate process.
6. Scope the first build to **Word only**; treat Excel and PowerPoint as separate products
   pending their own geometry probes (§7's adapter seam is thinner than claimed).
7. Build a **disposable typing/caret/selection prototype before** the production renderer, so
   the editing feel is judged before anything durable is built on it.
