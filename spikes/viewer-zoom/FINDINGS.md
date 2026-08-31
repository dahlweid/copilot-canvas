# Zooming the PDF viewer — findings

Written for issue #106. One probe, `probes/probe-zoom-overflow.mjs`, four
questions. It needs Edge or Chrome; it needs no Word and no Office.

Measured on Windows, Node v24.18.0, Google Chrome (headless), against this
branch. The probe serves the shipped `src/ui/` over HTTP with the real vendored
pdf.js — worker reassembled from its committed parts by the product's own
`vendor-assets.mjs` — drives the browser over CDP, and reads
`getBoundingClientRect`, `clientWidth`, `scrollLeft` and `HTMLCanvasElement.width`
off the running product.

## Why a probe at all

Every one of the four questions below is a **layout or paint** result, and the
unit suite cannot answer any of them: `test/unit/ui-harness.mjs` is explicit that
it computes no styles and lays nothing out. Its stub elements have no
`getBoundingClientRect` and no scroll box, so the assertion "the left edge is
reachable" has no meaning there at all.

The document is built by the probe rather than committed: a minimal
three-page PDF, uncompressed, Helvetica, one line per page, with a real xref.
Nothing document-shaped is committed to this repo, and pdf.js parses it through
the same code path it uses on a Word export.

## The run

Panel 900px wide (viewer `clientWidth` 885 once its scrollbar is taken), then
narrowed to 520 (`clientWidth` 505) and widened back.

| step | scale | viewer | page css | canvas px | scrollW | scrollLeft floor | leftReach | rightReach | spans | layer w |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| open (fit-width) | 1.394 | 885 | 853 | 853 | 885 | 0 | +16 | +16 | 1 | 853 |
| zoom in ×1 | 1.742 | 885 | 1066 | 1066 | 1098 | 0 | +16 | −197 | 1 | 1066 |
| zoom in ×2 | 2.178 | 885 | 1332 | 1332 | 1364 | 0 | +16 | −463 | 1 | 1332 |
| zoom out ×1 | 1.742 | 885 | 1066 | 1066 | 1098 | 0 | +16 | −197 | 1 | 1066 |
| fit-height | 0.891 | 885 | 545 | 545 | 885 | 0 | +170 | +170 | 1 | 545 |
| fit-width | 1.394 | 885 | 853 | 853 | 885 | 0 | +16 | +16 | 1 | 853 |
| zoom out below fit | 0.892 | 885 | 545 | 545 | 885 | 0 | +170 | +170 | 1 | 545 |
| **fitted, panel → 520** | **0.773** | **505** | **473** | 473 | 505 | 0 | +16 | +16 | 1 | 473 |
| **zoomed 0.966, panel → 900** | **0.966** | **885** | **591** | 591 | 885 | 0 | +147 | +147 | 1 | 591 |

`leftReach` is the page's left edge in the viewer's own coordinates *after* the
viewer has been scrolled as far left as it goes. A negative number is content
that exists and cannot be reached. `rightReach` is the same distance on the
other side, and equal to `leftReach` is what centred means. `canvas px` is
`HTMLCanvasElement.width` — the bitmap's own pixel width, not its CSS width.

Both margins are measured from the viewer's **content** box, not the border box
`getBoundingClientRect()` returns. Measured the obvious way first and every row
whose page was narrower than the viewer came out 15px asymmetric — that is the
vertical scrollbar, which lies inside the border box and outside the content
box, not a layout fault. Worth recording because the naive measurement reports
a perfectly centred page as off-centre, and by a plausible-looking amount.

## Question 1 — is the left edge of a zoomed page reachable

Only because of one declaration. The control holds everything else fixed —
same document, same scale 2.178, same viewport — and puts `.pages` back to
centring alone:

| `.pages` | leftReach | scrollWidth | scrollLeft floor |
| --- | --- | --- | --- |
| `align-items: center` alone | **−223.5** | 1109 | 0 |
| `+ min-width: max-content` | **+16** | 1364 | 0 |

A flex line only grows to `max-content` because something asks it to. Without
that, a 1332px page centred on an 885px line hangs off **both** ends by 223.5px,
and the scroll range covers the end overflow only: `scrollLeft` is already at
its floor of 0 with a quarter of the page to the left of the viewport. The
scroll container does not know the content is there — `scrollWidth` reports
1109, not the 1364 the page actually occupies.

`min-width: max-content` makes the box at least as wide as its widest child, so
a wide page grows the scrollable area instead of overflowing it. Chosen over
`align-items: safe center`, which fixes the same thing but only where the `safe`
keyword is supported.

### And the narrow case still centres

The tempting fix is `align-items: flex-start`, which also cures the overflow —
and silently loses centring for every page narrower than the panel, which at a
default fit is most of them. That failure is invisible in the wide case, so the
probe measures both margins on every row rather than the left one alone:

| step | page css | viewer | leftReach | rightReach |
| --- | --- | --- | --- | --- |
| open (fit-width) | 853 | 885 | +16 | +16 |
| fit-height | 545 | 885 | +170 | +170 |
| zoom out below fit | 545 | 885 | +170 | +170 |
| zoomed 0.966, panel → 900 | 591 | 885 | +147 | +147 |

Equal to the pixel on all six narrow rows, at four different scales and two
panel widths, including one reached by zooming *out* well below the fit and one
reached by dragging the panel. `align-items: center` is untouched by this
change; the diff only adds a minimum width, which does nothing at all until a
child is wider than the line.

## Question 2 — do the bitmaps repaint, or only the boxes

`#renderPage` early-returns on `page.rendered`, so a re-scale that resizes the
boxes without clearing that flag leaves every bitmap at the scale of the *first*
paint. That state photographs perfectly: the page is the right size and the
image inside it is smooth, because the browser is scaling one bitmap.

Measured, `canvas px` tracks `page css` exactly at all eight steps, including
**two zoom-ins in a row** — 853 → 1066 → 1332 — which is the case a single zoom
cannot distinguish. The re-scale path per page cancels the render task, cancels
the text-layer task, resets `rendered`, re-measures, and repaints every page the
`IntersectionObserver` last reported as visible.

That last clause is load-bearing and is *not* obvious: an
`IntersectionObserver` does **not** re-fire on a re-scale. A page that was
intersecting and still is produces no entry, so nothing would ask it to repaint.
The view therefore records `page.visible` from each entry and drives the repaint
itself.

## Question 3 — does the text layer survive

pdf.js's own stylesheet sizes every span with
`round(down, var(--total-scale-factor) * Npx, var(--scale-round-x))`. If that
custom property stops being set the declaration is **invalid** — not wrong, not
stale — and the layer collapses to zero width, taking text selection and the
change overlay with it.

Measured, the layer keeps its spans and its width tracks the page at all eight
steps (853 / 1066 / 1332 / 1066 / 545 / 853 / 473 / 591). The re-scale sets
`--total-scale-factor` on the container before it re-measures a single page.
The guard on this is `TEXT_LAYER_CONTRACT` in `test/unit/mutate-pdfjs.ps1`.

## Question 4 — does a fit survive a panel resize

It does, and a hand-picked zoom is not overwritten by one.

- **Fit is sticky.** At fit-width, narrowing the panel from 885 to 505 took the
  page from 853 to **473** — exactly `505 − 32`, the viewer's client width less
  the 32px gutter. The fit was recomputed, not merely preserved as a number.
- **A chosen scale is not.** After pressing zoom-in the standing fit is
  cleared; widening the panel from 505 back to 885 left the scale at
  **0.966**, the value the user chose, rather than snapping back to fit.

This is what `#scaleObserver` now is. It was previously declared, never
assigned, never read and never disconnected, with no `ResizeObserver` anywhere
in the tree — the vestige of exactly this behaviour. It is now a real
`ResizeObserver` on the viewport, observing `box: "border-box"` because the
content box changes when a scrollbar appears and fit-height reads
`clientHeight`, which could otherwise oscillate. It is disconnected in
`#teardown`, which the unit suite asserts.

## Re-running

```
node spikes/viewer-zoom/probes/probe-zoom-overflow.mjs
```

Exits 0 when all six findings hold, 1 otherwise, so it is a regression check as
well as a record.
