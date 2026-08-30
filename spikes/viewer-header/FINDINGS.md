# The viewer's header marks — findings

Written for issue #87. One probe, `probes/probe-mark-alignment.mjs`, three
questions. It needs Edge or Chrome; it needs no Word and no Office.

Measured on Windows, Node v24.18.0, Microsoft Edge (headless), against
`6dc3369` for the "before" figures and against this branch for the "after".

## Why a probe at all

Both defects arrived as screenshots of the shipped panel, and the issue offered
a cause for the first one — *`align-self: center` centres the mark against a
flex line whose height is set by `.doc-meta` at 12px* — as an **explicit
hypothesis**, not a diagnosis. It cannot be settled by reading CSS: the flex
line's cross size depends on what the **grid** does to `.bar-main` before the
flex algorithm runs, and the `align-self` on `.word-mark` reads exactly the same
either way.

So the probe serves the shipped `src/ui/` over HTTP with stub `/api/state` and
`/api/word-icon` routes, drives headless Edge over CDP, and reads
`getBoundingClientRect` off the running product. It is a measurement of
Chromium, which is what the panel is.

## Question 1 — where is the mark, and why

Offsets are from the filename's **optical** centre — half the cap height above
the baseline — because that is what the eye compares a 16px glyph against, not
the line box. Positive means the mark sits **low**.

### Before (`6dc3369`)

Reproduced on demand rather than quoted: the probe carries a `before #87`
variant that undoes the one property the fix adds, so the "before" half of this
comparison can still be produced after the fix has landed. Its numbers match the
run against `6dc3369` exactly (+5 / +4.5 / +5).

| theme | variant | line h | actions h | mark centre | optical centre | offset |
| --- | --- | --- | --- | --- | --- | --- |
| default (14/20) | **before #87 (bar-main stretched)** | 30 | 30 | 23 | 18 | **+5** |
| default (14/20) | mark `align-self: baseline` | 21 | 30 | 20.5 | 23.5 | −3 |
| default (14/20) | mark `align-self: first baseline` | 21 | 30 | 20.5 | 23.5 | −3 |
| default (14/20) | **`.bar-main { align-self: center }`** | 20 | 30 | 23 | 23 | **0** |
| default (14/20) | bar-main centred + mark baseline | 21 | 30 | 20.5 | 23.5 | −3 |
| large (16/24) | **before #87** | 34 | 34 | 25 | 20.5 | **+4.5** |
| large (16/24) | mark `align-self: baseline` | 25 | 34 | 22.5 | 25 | −2.5 |
| large (16/24) | **`.bar-main { align-self: center }`** | 25 | 34 | 25 | 25 | **0** |
| meta absent | **before #87** | 30 | 30 | 23 | 18 | **+5** |
| meta absent | mark `align-self: baseline` | 21 | 30 | 20.5 | 23.5 | −3 |
| meta absent | **`.bar-main { align-self: center }`** | 20 | 30 | 23 | 23 | **0** |

### The hypothesis is refuted

`.doc-meta`'s 12px font is **not** the cause. The `meta absent` row is the
control: with `.doc-meta` emptied the offset is unchanged, still exactly +5px.
A 12px sibling could only ever have made the line *shorter* than the 20px name,
and the measured line was 30px — **taller** than both.

The real cause is one layer up. `.bar-main` is a **grid item**, and a grid item
stretches by default to the row's height. That row's height is set by the action
buttons in the *other* grid area: `actions h` is 30 in every row above, and the
stretched `.bar-main` measured 30 to match. A single-line flex container takes
its line's cross size from the container, so the line was 30px and
`align-self: center` on the mark centred it against **the buttons**, 5px below
the filename it is supposed to sit beside. The theme rows corroborate: at
16/24 the buttons grow to 34 and the offset moves with them, to +4.5.

### The fix, and the one that looks right and is not

The property that expresses "align to the name" is `align-self: center` on
**`.bar-main`** — it tells the grid not to stretch that area, so the flex line
collapses to its content (20px, the filename's own box) and the mark's existing
`align-self: center` centres on the name. Measured 0px in all three themes.

`align-self: baseline` on the mark is the obvious-looking fix and is wrong. A
replaced element's baseline is its bottom margin edge, so a 16px icon sits *on*
the text baseline and reads 3px **high** — and it makes the line taller (21px)
to accommodate the alignment, which the centred variants do not.

No margin, translate or offset appears in any variant. One tuned against one
font stack, one zoom level and one theme is wrong on the next; the table above
is what that claim rests on.

### After (this branch)

```
markTop 15, markHeight 16, markCentre 23
nameTop 13, nameHeight 20, nameBaseline 28, capHeight 10, opticalCentre 23
offsetFromOptical 0, offsetFromNameBox 0
barMainHeight 20, barActionsHeight 30
markAlignSelf "center", barMainAlignSelf "center"
```

`barMainHeight` 20 rather than 30 is the mechanism, visible directly: the bar's
main area no longer stretches to the buttons.

## Question 2 — `hidden` does not hide an `<svg>`

Not in the issue. Found while measuring the *Open in Word* button.

| element | interface | `hidden` accessor in prototype chain | attribute set | computed display | client rects |
| --- | --- | --- | --- | --- | --- |
| `svg` | `SVGSVGElement` | false | false | `inline` | 1 |
| `img` | `HTMLImageElement` | true | true | `none` | 0 |

`hidden` is an IDL attribute of **`HTMLElement`**. `SVGElement` does not inherit
it, so `svg.hidden = true` sets an expando property on the object, reflects no
attribute, and hides nothing.

**Consequence: the glyph swap never worked.** On a machine with Word the button
drew the brand mark *and* the arrow. That is the second half of what the
screenshot shows, and it is why the button measured 21px wider there.

`word-mark.test.mjs`'s "a loaded mark replaces the drawn glyph" passed
throughout, because it drove a stub object where `hidden` is a plain data
property. It is a clean instance of this repo's *test that cannot fail for the
reason its comment claims* class: the stub could not model the one platform
behaviour the assertion depended on. That is the strongest argument for
**removing** `showWordMark`'s `fallback` parameter rather than keeping it —
keeping it preserves a branch the product does not use, whose only covering test
cannot go red.

## Question 3 — is the button deterministic?

The button's appearance must not depend on whether this machine's Word yielded
an icon. Measured by loading the same page twice, once with `/api/word-icon`
answering with a PNG and once with it answering 404.

### Before (`6dc3369`)

Measured against the tree at `6dc3369`, before the markup changed. Unlike the
geometry above, this half is **not** reproducible from this branch: the second
mark is gone from `index.html`, and the probe reports what the markup contains
rather than injecting an element to measure.

| machine | width | height | visible children |
| --- | --- | --- | --- |
| Word installed | 136.34 | 30 | `img#openInWordMark,svg#openInWordGlyph` |
| no Word (404) | 115.34 | 30 | `svg#openInWordGlyph` |

identical: **no** — 21px of difference, decided by whether this machine's Word
yielded an icon.

### After (this branch)

| machine | width | height | visible children |
| --- | --- | --- | --- |
| Word installed | 115.34 | 30 | `svg#openInWordGlyph` |
| no Word (404) | 115.34 | 30 | `svg#openInWordGlyph` |

identical: **yes**

## Notes for anyone re-running it

- **No image is committed.** The stub `/api/word-icon` builds a 32×32 PNG in
  memory (CRC32 + `deflateSync` + hand-built IHDR/IDAT/IEND). The repo is public
  and the mark is Microsoft's; a tree-walking test in `ui-contract.test.mjs`
  asserts no `.png`/`.ico`/`.jpg`/`.gif`/`.bmp` anywhere.
- **Baseline is measured, not derived.** A zero-height
  `display: inline-block; vertical-align: baseline` span is appended to the
  name; its `top` *is* the baseline. Cap height comes from canvas
  `measureText("H").actualBoundingBoxAscent`. Neither is read out of a font
  table.
- **Edge is launched with discrete argv elements** via `spawn` — never a shell,
  never `cmd.exe`, never interpolated into a command string. This repo has been
  bitten twice by a path crossing a parser nobody accounted for.
- **No CDP dependency.** Node 22+ has a global `WebSocket`;
  `Target.createTarget` → `Target.attachToTarget { flatten: true }` is enough.
  The DevTools endpoint is printed on **stderr**.
- **Gotcha, hit and fixed:** the probe holds `/events` open to mimic the real
  SSE stream, and `server.close()` waits for it — the probe printed its table
  and then hung. `server.closeAllConnections()` must come first.

## What this does *not* measure

The unit suite executes `src/ui/app.js` under Node (#85), but that harness
stands in `fetch`, `EventSource`, `IntersectionObserver` and the DOM. It is
Node, not Chromium. Every Chromium claim on this page comes from the probe and
from nowhere else, and the panel was not verified end to end.

## The mutation harness, and one mutant that survived

`mutate.mjs` applies and reverses each of five named edits, so the assertions
guarding this change can be shown to go red for the right reason:

| mutant | file | what it does |
| --- | --- | --- |
| `stretch` | `app.css` | deletes `align-self: center` from `.bar-main` |
| `nudge` | `app.css` | adds `margin-top: 2px` to `.word-mark` |
| `second-mark` | `index.html` | puts a second Word mark back in the button |
| `rewire` | `app.js` | wires that mark up again, with a fallback |
| `fallback` | `word-mark.mjs` | makes `showWordMark` write to a second element |

`stretch` **survived** the first time it was run. The alignment test extracted
the `.bar-main` rule and matched `align-self:\s*center` against the raw text —
and the long comment inside that rule explains the fix by naming the
declaration, so the assertion matched the prose and passed against a rule that
no longer contained it. Comments are now stripped before the rule is extracted,
after which the mutant dies. A test that documents its subject in the text it
greps is a way for a guard to read as protection while protecting nothing, and
it is not specific to CSS.

Two smaller traps hit while running the matrix, both worth the next reader's
attention:

- **`git checkout --` is not a backup.** These files were staged, then edited
  again; restoring a mutation with `git checkout -- <file>` silently reverted
  the later edit too, and an unrelated test went red under a mutation that could
  not have caused it. The harness reverses its own edit instead.
- **`$args` is a PowerShell automatic variable.** A helper function declaring a
  parameter of that name ran `node --test` with no files at all: `tests 0`,
  exit 0, five mutants apparently surviving. It is the repo's known
  glob-matches-nothing trap wearing different clothes. Assert the test count is
  non-zero, always.
- **`\bhidden\b` matches inside `aria-hidden="true"`.** `-` is not a word
  character, so the obvious spelling reports every decorative icon in this
  markup as hidden. It made a correct `<svg>` fail a "ships visible" assertion.
  The `hidden` content attribute has to be matched standing alone.
- **Prose in markup is markup to a regex.** A comment inside the Open in Word
  button described the bug using a literal `<svg>`, and the repo's
  icons-are-decorative test — which scans each button's inner markup for
  `<svg …>` — matched the sentence and demanded `aria-hidden` on it.
