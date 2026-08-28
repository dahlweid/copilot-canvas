# Spike: live Word pixel-streaming vs. PDF export

**Question the plan asked:** is streaming a live, off-screen Word window into the canvas good
enough to replace the PDF renderer — now, or later for read-write mode?

**Verdict: no for v1, yes as the path to editing.** Fidelity is not the problem; it is
excellent. The problem is that streaming buys a live window we do not yet need, and pays for
it with a *visible* Word process, a continuous capture loop, a fiddly DPI correction, and the
loss of everything the native PDF viewer gives us for free (text selection, Ctrl+F, print,
zoom, page thumbnails).

Reproduce with:

```
node .github/extensions/word-canvas/spikes/live-word/run-spike.mjs [document.docx]
```

It writes `frames/report.json` plus three sample JPEGs. Everything below was measured on this
machine (Word 16.0, German UI, 125% display scaling) against a generated 7-page document.

## How it works

`live-word.ps1` is the same shape as the production `word-host.ps1` — a long-lived PowerShell
process speaking newline-delimited JSON over stdio, with the same PID-file ownership tracking
and kill-fallback cleanup — but it opens Word **visibly** and captures its window instead of
exporting a PDF.

- The window is parked at `x = 30000` with `SetWindowPos`, i.e. beyond the right edge of any
  plausible desktop, and shown with `SW_SHOWNOACTIVATE` so it never takes focus.
- Frames come from `PrintWindow` with `PW_RENDERFULLCONTENT`, encoded to JPEG by an embedded
  C# class (`Add-Type`). PowerShell alone is far too slow for a capture loop.
- Scrolling goes through COM (`Window.SmallScroll`), not synthetic input.
- `crop` trims the ribbon, which costs ~230 px of the captured bitmap.

`Application.Visible = $true` is **unavoidable**: `PrintWindow` returns an empty bitmap for a
hidden window. This is the single biggest behavioural difference from the shipped renderer,
which runs Word fully hidden.

## Measurements

| Metric | Result |
| --- | --- |
| Cold start (Word launch → first frame, 560×950) | 4.1 s |
| Capture round-trip (Node → PS → JPEG → Node) | 34.3 ms mean, 33 ms p50, 46 ms p95 |
| Capture in-process (PrintWindow + encode) | 29.2 ms mean |
| Achieved frame rate | **29 fps** |
| Frame size (560×950, JPEG q70) | 66 KB |
| `SmallScroll` COM call | 6.3 ms mean |
| Scroll → next frame | **37.5 ms mean**, 50 ms p95 |
| Page jump (`GoTo` + `ScrollIntoView` + frame) | 110 ms mean |
| Resize to 900×1200 | honoured exactly |
| Leaked `WINWORD.EXE` after run | none |

At ~30 fps with 38 ms scroll latency this feels like a live document, not a remote desktop.
Bandwidth is ~2 MB/s while actively scrolling and zero when idle (frames are only captured on
demand).

## Fidelity

**The pixels are correct.** They come from Word's own layout engine, so they are accurate by
construction — see `frames/frame-wide.jpg`, where the page renders white with black text and
correct pagination even though the surrounding Word chrome is in the user's dark Office theme.
Word's dark theme applies to the application chrome, not the page.

This corrected an earlier wrong conclusion in this spike: intermediate frames appeared to show
a *dark page*, which looked like a fatal fidelity divergence. It was an artefact of the zoom
bug below — the page was scaled past the edges of the captured bitmap, so the frame was filled
entirely with the dark area *around* the page. Once the fit was correct, the page appeared
white. Worth recording, because the wrong version of this finding would have killed the
approach for the wrong reason.

### The DPI trap

This cost the most time and is the main hidden complexity of the approach:

1. Word does **not** re-fit the page when the window is resized via `SetWindowPos`. The zoom
   must be reset explicitly (`View.Zoom.PageFit = wdPageFitBestFit`), otherwise the page keeps
   whatever width it had when the document opened and is clipped horizontally.
2. `PageFit` fits the page to the window rect, but Word then *paints* at the display scale
   factor. At 125% scaling the page overflows the captured bitmap by exactly 1.25×. The zoom
   has to be divided by the scale factor afterwards.
3. The scale factor cannot be read with `GetDeviceCaps` — this host process is DPI-unaware, so
   Windows virtualizes it to a flat 96 DPI and it always reports 1.0. It has to be read from
   *Word's* window with `GetDpiForWindow(hwnd)`, which reports 120 here.
4. Even then the fit overshoots by a few percent, because `PageFit` measures against a client
   width that excludes chrome the capture does include. The spike applies an empirical `0.92`.

Step 4 is the uncomfortable one: an empirical fudge factor is a bad foundation for a renderer
whose entire selling point is accuracy. Making this robust would mean detecting the page
rectangle in the captured bitmap and correcting from that, which is real work.

None of this exists on the PDF path, where the page geometry is defined by the PDF itself.

## Comparison

| | PDF export (shipped) | Live pixel streaming |
| --- | --- | --- |
| Page accuracy | Print-exact by construction | Print-exact, after DPI correction |
| Word window | Fully hidden | **Must be visible** (parked off-screen) |
| Text selection / copy | Yes, native | No — it is a bitmap |
| Ctrl+F in the panel | Yes, native | No (our own search only) |
| Print / save a copy | Yes, native | No |
| Zoom / page thumbnails | Native viewer chrome | Would have to be rebuilt |
| Cost when idle | Zero | Zero (frames on demand) |
| Cost when scrolling | Zero — the viewer scrolls locally | ~30 fps capture + encode loop |
| Latency to first paint | ~4.7 s cold, sub-second warm | 4.1 s cold |
| Reacting to an external edit | Re-export (~1 s warm) | Instant — same document object |
| Live caret / selection / editing | Impossible | Natural |
| Scroll position across reloads | Lost (viewer is not scriptable) | Fully under our control |
| DPI handling | None needed | Fiddly, partly empirical |

## Conclusion

For a **read-only** viewer the PDF path wins on every axis that matters. Streaming would trade
away text selection, native search, print and viewer chrome — and add a visible Word window —
in exchange for advantages (live caret, instant reaction to edits, controllable scroll
position) that a read-only viewer cannot use.

For a future **read-write** mode the calculus inverts. Editing needs a live caret, immediate
visual feedback, and preserved scroll position; re-exporting a PDF per keystroke is not
viable. At 29 fps and 38 ms scroll latency, streaming is comfortably fast enough to carry it,
and this spike shows the hard parts (off-screen placement, capture, COM-driven scrolling,
process ownership) all work.

**Recommendation:** keep the PDF renderer for read-only. Revisit this spike when read-write is
picked up, and budget for the page-rectangle detection that would replace the empirical zoom
correction.

## Deviation from the plan

The plan proposed shipping this as a second, experimental `word-live` canvas inside the
extension so both renderers could be opened side by side. That was not done, deliberately: it
would put a *visible* Word window and a continuous capture loop into an extension whose v1
promise is that Word stays hidden and untouched — a real hazard (a stray window can surface on
a multi-monitor setup) in exchange for a comparison this document already settles. The spike
remains fully runnable and reproducible as a standalone script. If a live side-by-side is
wanted later, `live-word.ps1` already exposes everything a canvas would need (`start`,
`resize`, `crop`, `capture`, `scroll`, `goto`, `zoom`, `info`, `stop`).

## Sample frames

- `frames/frame-full.jpg` — raw capture at panel size, ribbon included.
- `frames/frame-cropped.jpg` — ribbon cropped; this is what a canvas would display.
- `frames/frame-wide.jpg` — 900×1200, showing full chrome and a print-accurate white page.
