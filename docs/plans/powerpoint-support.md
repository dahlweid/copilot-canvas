# Plan: full PowerPoint support, display path first

Status: **proposal**. Nothing here has been filed as an issue. No product code
was written for it.

This document answers two questions the maintainer asked: *what exists today for
PowerPoint*, and *what is required to add full PowerPoint support*, with the
display/export path prioritised.

Every platform claim below is either backed by a probe that was actually run —
cited by path — or marked **unknown** and listed as a spike to run. Nothing was
executed to produce this document: no PowerPoint was started, no probe was run,
no Word was started, no integration suite was taken. The measurements are read
from `spikes/powerpoint/FINDINGS.md`, which is what a committed spike tree is
for.

---

## 1. What exists today

### 1.1 Shipped code: nothing

Searched `.github/extensions/office-canvas/` for `powerpoint`, `pptx`,
`POWERPNT`, `presentation` and `slide`, excluding `src/ui/vendor/`. In the
shipped extension — `extension.mjs` and everything under `src/` — there are
**zero matches**.

Concretely:

- `RenderCache.SUPPORTED` is `{.docx, .docm, .doc, .dotx, .rtf}`, and
  `render-cache.mjs` documents it as *the single source of truth* from which tool
  descriptions and the picker are derived. `.pptx` is not in it.
- One canvas is registered: `wordCanvas`, `id: "word-doc"`. `extension.mjs`
  exports `canvases: [wordCanvas]`.
- Four tools ship: `create_document`, `read_document`, `edit_document`,
  `revert_document`, plus seven canvas actions (`open_document`, `go_to_page`,
  `get_outline`, `search`, `get_text`, `refresh`, `get_info`).
- The only PowerPoint mention anywhere in the extension folder is a comment in
  `test/integration/create-smoke.mjs` contrasting Word's create semantics with
  PowerPoint's attach semantics.

**Correction to the brief.** The brief says v0.1 ships "Word support plus some
basic PowerPoint groundwork". There is no PowerPoint groundwork in the product.
What exists is PowerPoint **evidence**: a completed spike, and an ADR amendment
written from it. "We measured it" and "we shipped it" are exactly the two states
the brief warned against blurring, and the tree is entirely in the first.

### 1.2 What does exist

| Artefact | State |
| --- | --- |
| `spikes/powerpoint/FINDINGS.md` | Complete. Six questions answered, one correction recorded, one unresolved item. |
| `spikes/powerpoint/probes/` (12 probes) | Committed and re-runnable. Under repair for #139 — see §7.2. |
| `docs/adr/0003` amendment | Written. Confirms the rendering premise, rewrites the process-model premise. |
| Product code | None. |

`docs/adr/0003-one-extension-many-canvases.md` still says, in the present tense,
that "Word, Excel, and PowerPoint ship as a single extension that declares a
canvas per application". One canvas is declared. That sentence describes the
intended end state, not the tree, and it should be corrected when the second
canvas lands, or before.

---

## 2. The display path

This is the near-term target and the evidence for it is unusually strong.

### 2.1 Export to PDF and reuse the renderer — confirmed

`spikes/powerpoint/probes/probe-export.ps1` measured, on a generated 13-slide
deck at 960 x 540 pt:

- **13 PDF pages for 13 slides.** Exactly one page per slide.
- **`/MediaBox` matches `PageSetup.SlideWidth/Height` to 0.000 x 0.000 pt** on
  every page, read by scanning the PDF bytes directly with no PDF library.
- **94–110 ms per single-page export**, against Word's 168 ms. The whole deck
  exported in 815–865 ms (~64 ms/slide). Single-slide export is **flat**: slide 7
  costs what slide 1 costs, so a render-on-demand canvas does not pay for
  position in the deck.

The brief's hypothesis is therefore correct, and better than it guessed:
PowerPoint's page mapping is *exact* where Word's is merely adequate, and its
per-page cost is *lower*. There is no rendering reason to split the extension and
no new vendored asset is needed for display. **The existing pdf.js worker renders
PowerPoint output unchanged.**

### 2.2 Where it breaks

**Notes pages.** Exporting with `ppPrintOutputNotesPages` instead of
`ppPrintOutputSlides` also produces 13 pages — but at 540 x 720 pt, portrait. The
page *count* would not have caught it; the geometry check did. This is a one-line
trap in the export call, and the geometry assertion is the guard.
(`spikes/powerpoint/probes/probe-export.ps1`)

**`PrintRange` is not optional.** `$null`, `[Type]::Missing` and
`[Reflection.Missing]::Value` all fail — `NullReferenceException` for the first,
"Missing parameter does not have a default value" for the last. A real
`PrintOptions.Ranges` object is mandatory even when exporting the whole deck.
Word's equivalent argument is optional, so this is a direct port hazard.

**Animations, builds, transitions, hidden slides, speaker view.** *Unknown.* No
probe in `spikes/powerpoint/` touches any of them. A build-animated slide almost
certainly flattens to its final state in a fixed-format export, and a hidden
slide plausibly does not export at all — but neither is measured, and a deck
whose slide count and page count disagree because three slides were hidden would
break page-to-slide addressing silently. **Spike S1, §8.**

### 2.3 The cost figures describe the wrong instance

This is the most important thing I found that the brief did not anticipate.

`spikes/powerpoint/probes/probe-export.ps1` obtains its instance with
`New-OwnedPowerPoint`, which is `New-Object -ComObject PowerPoint.Application` —
the **COM-attach** path. Every export timing above (94–110 ms/slide,
815–865 ms/deck, ~3019–3494 ms cold start) was measured on an attached instance.

But §3 establishes that the host **cannot use** the COM-attach path, because on a
machine where the user has PowerPoint open, that instance *is* the user's. The
architecture requires the isolated separate-desktop instance. **No probe measured
export cost from an isolated instance.**

The only isolated-instance timing that exists is 12,144 ms / 12,146 ms to reach a
bound instance with a deck open, from the **control arm** of
`spikes/powerpoint/probes/probe-cross-instance-lock.ps1` — reproduced twice, 2 ms
apart. That is launch plus bind plus open, not export.

So the honest statement of the display budget today is: **first render ≈ 12 s
(measured, n=2), steady-state per-slide export ≈ 100 ms (measured, but on a
different instance type than we will use).** Whether the isolated instance
exports at the same rate is **unknown**. **Spike S2, §8.**

---

## 3. The host: single-instance changes everything

### 3.1 What is measured

`spikes/powerpoint/probes/probe-single-instance.ps1`, with Word as the control:

| Test | PowerPoint | Word (control) |
| --- | --- | --- |
| Two `New-Object` calls in one process | **attached**, 1 process | 2 processes |
| `New-Object` from a separate process | **attached**, 1 process | n/a |
| Launch `POWERPNT.EXE` on the same desktop | **exits**, hands off | n/a |

The Word control is what makes this a property of PowerPoint rather than of the
method. A second attach also saw and could name the deck the first had opened —
one object model, not two.

The escape works: `spikes/powerpoint/probes/probe-second-process.ps1` launched
`POWERPNT.EXE` into a separate window-station desktop and bound through
`AccessibleObjectFromWindow` / `OBJID_NATIVEOM`, the technique
`spikes/isolation/probes/probe-bind.ps1` already uses for Word. Proven isolated:
the bound instance had the deck open while the shared instance reported
`Presentations.Count = 0`. The bind target is **`mdiClass`**, not Word's `_WwG`;
`MsoCommandBar` and `RICHEDIT60W` also answer the binder but hand back a stub
with an empty `Version` and no `Presentations`. The test is "can it name the
deck", not "did it answer".

### 3.2 What a PowerPoint host can legitimately be

Three viable shapes. I recommend **B**.

**A — attach to the user's PowerPoint.** Rejected. `Quit()` closes their decks;
killing "our" pid kills theirs; `DisplayAlerts` and `AutomationSecurity` leak into
their session. `AutomationSecurity` is the sharper edge: the Word host
force-disables macros before opening anything because we render documents the
user may not have authored, and setting that on the user's instance changes the
security posture of *their* work. Not viable at any price.

**B — isolated instance on a separate desktop, always. (Recommended.)** Never
call `New-Object -ComObject PowerPoint.Application` in product code at all, so
the failure mode cannot occur. Cost: ~12 s to a bound instance. Everything the
Word host does — invisible instance, attribution, reap ledger, idle shutdown —
then has an analogue, but the attribution mechanism must change (§3.4).

**C — no PowerPoint at all for read and display.** Genuinely available for part
of the surface, and this is the strongest single finding in the spike for
sequencing purposes. See §4.

The recommendation is **B for export, C for structure**, which is the split §5
builds the tool surface on.

### 3.3 Must the host refuse to act when the user has PowerPoint open?

**No — and this is worth stating because the intuitive answer is yes.**

Under shape B we never attach, so the user's instance is irrelevant to *starting*
ours. What their instance can do is **hold a deck**, and that is a file-level
problem with a file-level answer, not a reason to refuse.

`spikes/powerpoint/probes/probe-cross-instance-lock.ps1` measured a genuinely
separate PowerPoint opening a deck the first instance holds: it **hung
>71,000 ms** and had to be killed by pid. It created its frame, created its
`mdiClass` window, and then `Presentations.Count` stayed at **0 forever**. No
dialog. No refusal. No read-only fallback. Identical in shape to the Word finding
that forced ADR 0005, and reproduced twice.

So the rule the Word host already follows applies unchanged and is the answer:
**never open the original; always open an unblocked temp copy** (ADR 0006). A
held deck is readable — `spikes/powerpoint/probes/probe-bulk-read.ps1` read the
zip of a held deck in 275 ms and notes explicitly that only writes are blocked —
so the copy succeeds and the isolated instance opens a file nobody holds.

One caveat, stated because it is inference and not measurement: the 275 ms figure
is a .NET zip read, not `Copy-Item`. That `Copy-Item` succeeds against a
PowerPoint-held deck is **inferred** from the same share-mode algebra measured for
Word in `spikes/isolation/probes/probe-fileshare-algebra.ps1`, not measured for
PowerPoint. **Spike S3, §8.** If PowerPoint turns out to hold a deck more strictly
than Word does, the display path for an open deck collapses to the OOXML zip and
the answer changes.

### 3.4 Lifecycle and teardown when quitting is never safe

`createIdleShutdown` in `src/word-lifecycle.mjs` is application-agnostic and
should be reused as-is. Its in-flight discipline — check the counter when *arming*
the timer and again when it **fires** — is the property that matters and is not
Word-specific.

What changes is `dispose`:

- **`Quit()` does not terminate `POWERPNT.EXE`** — 15/15 cycles left the process
  running. A kill is the fallback, not the exception.
- **A killed PowerPoint makes the *next* launch show a modal safe-mode prompt.**
  This silently broke the control arm of the lock probe until it was diagnosed. So
  teardown is: graceful `Quit()`, wait, kill only if the pid we created is still
  there, and be prepared to dismiss the prompt (`WM_COMMAND`/`IDNO`) on the next
  launch. Under shape B the pid is the one `CreateProcess` returned, so "the pid
  we created" is knowable without inference.
- Word's `Get-AttributedWordPid` derives the pid from `ActiveWindow.Hwnd` via
  `GetWindowThreadProcessId`, because differencing the pid set is **measured
  unsound** — `spikes/isolation/probes/probe-init-attribution.ps1` saw 2 new pids
  for 1 instance created, with the right one first only by luck. For PowerPoint
  under shape B the hwnd route is unnecessary: the separate-desktop launcher
  already owns a `CreateProcess` pid. But the *rule* transfers and hardens: **if
  the attributed pid is not one we created, refuse to kill and report
  "unattributed" rather than guessing.** For PowerPoint that refusal is not
  hygiene; a wrong guess destroys the user's unsaved work.
- The idle window should probably be longer than Word's `IDLE_SHUTDOWN_MS =
  60_000`, because re-acquiring costs ~12 s rather than ~1 s. That is a judgement
  call, not a measurement, and it should be made with S2's number in hand.

---

## 4. Structure reads need no PowerPoint at all

`spikes/powerpoint/probes/probe-bulk-read.ps1`:

| Arm | Method | Time |
| --- | --- | --- |
| A | naive per-shape COM walk, 39 shapes / 13 slides | 585 ms |
| B | one `Shapes.Range().TextFrame.TextRange.Text` per slide | 444 ms |
| C | **read the `.pptx` as a zip, parse `ppt/slides/*.xml`, no COM** | **82 ms** |
| C (held) | same, while PowerPoint holds the deck | 275 ms |

There is no `Content.WordOpenXML` equivalent in the PowerPoint object model, so
arm B is the best the object model offers and it is barely better than naive. The
real answer is that a `.pptx` *is* an OOXML zip.

Three consequences, in increasing order of importance:

1. It is 7.1x faster than the COM walk.
2. It works **while the deck is held open**, so it is not restricted to idle
   files.
3. It **needs no PowerPoint process**, so it costs nothing against the 12 s bind
   and is available before any instance exists.

Point 3 is what makes the sequencing in §9 possible: a read-only PowerPoint tool
surface can ship with **no PowerPoint host at all**, and the host is only needed
for pixels.

### 4.1 The join trap this creates

`spikes/powerpoint/probes/probe-localization.ps1` measured, on German Office:

| Shape | COM `Shape.Name` | stored `cNvPr/@name` | stored `p:ph/@type` |
| --- | --- | --- | --- |
| title | `Title 1` | `Titel 1` | `title` |
| body | `Text Placeholder 2` | `Textplatzhalter 2` | `body` |
| textbox | `TextBox 3` | `Textfeld 3` | *(none)* |

A control arm settles the mechanism: renaming a shape through COM round-trips to
disk verbatim, while untouched shapes stay German in the file. So `Shape.Name` is
a live property, and PowerPoint translates only *default* placeholder names into
English on read.

This lands directly on the recommended architecture. **A host that reads structure
from the zip and writes through COM sees a different name from each side,
silently.** And PowerPoint is *forgiving* about lookups — `Shapes.Item('Title 1')`
and `Shapes.Item('Titel 1')` both succeed, where German Word throws on
`Style = "Heading 1"`. Forgiving is worse: the failure does not announce itself.

Layout names are localized outright (`Titelfolie`, `Zwei Inhalte`,
`Abschnittsüberschrift`).

**Rule: never address a placeholder or layout by name.** Join on `p:ph/@type`,
`Shapes.Placeholders[i]`, or `PlaceholderFormat.Type` (1 = title, 2 = body) — all
three measured language-independent. This is the repo's existing "join on
structural keys, never on a display name" rule, and PowerPoint gives it a second,
sharper instance.

---

## 5. Tool surface

### 5.1 Addressing: is there a slide analogue of ADR 0006?

Partly, and it is *better* than Word's, which is the surprise.

ADR 0006 makes an address a **coordinate, not a handle** because Word exposes no
stable paragraph identity, so an address is derived from heading path plus text
plus occurrence index — and two mutations move it.

PowerPoint differs in one respect that matters: **slides carry a real identity.**
`Slide.SlideID` is documented as stable across reordering and insertion, and the
OOXML side has `p:sldId/@id`. Neither is measured in this spike — **Spike S4,
§8** — but if they behave as documented, slide-level addressing is a handle, not a
coordinate, and it survives insertion and reorder in a way no Word paragraph
address does.

Below the slide it degrades to Word's situation: a shape has `Shape.Name`, which
§4.1 proves is localized *and* mutable, and a placeholder index, which moves when
a placeholder is added. So:

- **Slide level: a handle** (pending S4). `slide_id`, plus a `slide_index` for
  human legibility that is explicitly documented as not an address.
- **Shape level: a coordinate**, addressed by `p:ph/@type` plus occurrence, on the
  same read-then-address contract and the same revision token as Word.

The revision token machinery (`src/revision-token.mjs`) is file-derived and
already application-agnostic. Reuse it.

### 5.2 Proposed tools

| Tool | Shape | Notes |
| --- | --- | --- |
| `read_presentation` | mirrors `read_document` | Paged like `read_document` (which defaults to 300 paragraphs); a deck is small enough that a slide-count default is more natural. Served entirely from the zip — **no PowerPoint process**. Returns slide id, index, `p:ph/@type` map, placeholder text, notes text, and a revision token. |
| `create_presentation` | mirrors `create_document` | Needs the host. Blocks on S2 and the whole of §3. Lowest priority of the four. |
| `edit_presentation` | **do not build a mirror of `edit_document`** | See §5.3. |
| `revert_presentation` | mirrors `revert_document` | Snapshot-and-rename is file-level and application-agnostic; `src/word/snapshots.mjs` is reusable almost verbatim. Only meaningful once something can write. |

Plus canvas actions: `open_presentation`, `go_to_slide`, `get_outline`, `search`,
`get_text`, `refresh`, `get_info` — the Word set with `page` becoming `slide`.

### 5.3 Editing: what is and is not feasible

The brief is right to be sceptical, and the evidence supports the scepticism for a
reason more specific than "slides are complicated".

**What genuinely works.** `spikes/powerpoint/probes/probe-lock.ps1` measured the
transient-lock round trip — open, edit, save, close — at **163–177 ms** warm, and
edit plus 1-slide re-export at 137–141 ms mean / 93 ms min. So a single scoped
mutation through COM is cheap and the ADR 0005 write model transfers unchanged.
Text replacement inside an identified placeholder is feasible today.

**Why a mirror of `edit_document` is still the wrong shape.** `edit_document`'s
five operations — replace text, insert paragraph before/after, delete paragraph,
set heading level — are meaningful because a Word document is a **linear sequence
of paragraphs** with an outline over it. That is what makes "insert a paragraph
after this one" a complete instruction. A slide is a **2-D canvas of positioned
shapes under a layout contract**. "Insert a shape after this one" has no meaning:
the missing information is geometry, and neither the agent nor the host can supply
it without either inheriting a layout placeholder or inventing coordinates.
Inventing coordinates produces decks that are subtly wrong in a way a structure
map cannot show and the canvas *can* — the worst combination available.

There is also a repo-standard argument. ADR 0006 refuses to batch edits because an
address is a coordinate and a batch API would imply a stability that does not
exist. The same reasoning refuses a generic slide-mutation API: it would imply a
positional model the format does not give us for free.

**Recommended instead**, in feasibility order:

1. `set_placeholder_text(slide_id, placeholder_type, text)` — feasible now.
   Addresses through the language-independent key, changes text only, changes no
   geometry. Measured cost ~170 ms round trip.
2. `add_slide(after_slide_id, layout, {placeholder: text})` — feasible, because a
   **layout supplies the geometry**. This is the one structural operation where
   the position question has an answer that is not invented. Blocked on S5 (layout
   addressing without names) and on the host.
3. `delete_slide(slide_id)`, `reorder_slide(slide_id, to_index)` — feasible and
   geometry-free. Both need S4 first, because both are exactly the operations that
   would move an index-based address.
4. Anything that positions a shape — **not planned**. Say so explicitly in the
   tool descriptions so the agent does not try to reach it by another route.

That is a *smaller* surface than Word's, deliberately, and it is the honest one.

---

## 6. The canvas

**A separate `powerpoint-deck` canvas, not a mode on `word-doc`.**

ADR 0003 already decided this and the reasoning survives contact with the spike:
the agent picks a canvas from its description, "Word document" is a clearer signal
than a generic canvas with a mode parameter, and the input schemas legitimately
differ (a page for Word, a slide for PowerPoint). `extension.mjs` exports
`canvases: [wordCanvas]` — an array — so a second entry is the shape the code
already anticipates.

What is shared, and should be extracted rather than copied: `src/server.mjs`,
`src/render-cache.mjs`, `src/render-cache-slot.mjs`, `src/watcher.mjs`,
`src/vendor-assets.mjs`, `src/revision-token.mjs`, `src/word-lifecycle.mjs`
(renamed), the whole of `src/ui/pdf-view.mjs`, and the vendored pdf.js. That is
the majority of the extension by bytes and by risk.

What differs in the UI: "page N of M" becomes "slide N of M", the outline is a
slide list rather than a heading tree, and `src/ui/word-mark.mjs` (Word's
layout-mark rendering) has no PowerPoint analogue.

Naming debt to settle early: `src/word-lifecycle.mjs`, `DocumentError`,
`word_error` as the default code in `src/tool-error.mjs`, and
`RenderCache.wordVersion` are all Word-shaped names on application-agnostic
machinery. `word_error` is the one that actually misleads — a PowerPoint failure
surfacing under that code names a cause the code cannot know, which is precisely
the class of defect this repo already treats as a real bug.

---

## 7. Error taxonomy

The repo's rule is **split where the platform distinguishes, stay collapsed where
it does not**. Applying it:

**`file_locked` and `permission_denied` stay split, unchanged.**
`spikes/powerpoint/probes/probe-lock.ps1` measured lock detection by attempting a
write handle at **0–1 ms in both directions and correct in both** (Word: 4 ms
held, 9 ms free). The exception types that separate the causes —
`System.IO.IOException` vs `System.UnauthorizedAccessException`, `EBUSY` vs
`EPERM` — are filesystem-level and application-independent. Nothing about
PowerPoint changes this.

**`file_locked` still provably never means "the user has it open in PowerPoint".**
The measured facts: a held deck is readable
(`spikes/powerpoint/probes/probe-bulk-read.ps1`, arm C held, 275 ms) while writes
to it fail (`spikes/powerpoint/probes/probe-lock.ps1` T2: direct overwrite fails,
write-temp-then-rename fails, exclusive write handle fails). So a *reader* of a
PowerPoint-held deck succeeds, exactly as with Word, and a message telling the
user to close PowerPoint would be wrong for the same reason. **Never advise
closing PowerPoint in an error message.** (One gap: that `Copy-Item` specifically
succeeds is inferred, not measured — S3.)

**`writable` stays a single collapsed flag.** Same reason as for Word: it comes
from taking a write handle, which folds sharing violation, ACL and the read-only
attribute into one observation. The platform does not distinguish there.

**New codes PowerPoint appears to need:**

| Code | Why | Evidence |
| --- | --- | --- |
| `powerpoint_unavailable` | the isolated launch or the `mdiClass` bind failed | `spikes/powerpoint/probes/probe-second-process.ps1` — the bind can return a stub with an empty `Version` and no `Presentations`, which must be a typed failure rather than a null dereference |
| `powerpoint_timeout` | the open never returned | `spikes/powerpoint/probes/probe-cross-instance-lock.ps1` — >71,000 ms, no dialog, external kill required |
| `bind_failed` | distinguishes "launched but could not bind" from "could not launch" | same probe; whether these need *different remediation* is not established, so **default to collapsed** until it is |

And one that must *not* be added: a code for "PowerPoint is already running".
Under shape B that is not a failure condition at all.

**Unresolved, and it belongs in this section.** An intermittent `0x800706BA` "RPC
server is unavailable" appeared across several probes, with matching
`POWERPNT.EXE` faults in the Windows Application log (`0xc0000005` in
`combase.dll` at offset `0x19e198`). Ruled out by probe: notes-page export
(`spikes/powerpoint/probes/probe-notes-control.ps1`, both arms died),
warm-vs-fresh instance reuse (`spikes/powerpoint/probes/probe-stability.ps1`,
0 failures in both arms on a clean run), idle self-termination
(`spikes/powerpoint/probes/probe-hide.ps1`), and the `Saved = $true` cast error
(`spikes/powerpoint/probes/probe-saved-flag.ps1`, deterministic and separate).
The remaining hypothesis is COM teardown racing PowerPoint's own shutdown. **It is
not characterised as a rate**, so it must not yet be used to argue for or against
anything — including for a retry policy. **Spike S6, §8.**

### 7.1 Word-shaped constants that will not port

`spikes/powerpoint/FINDINGS.md` collected these. Each is a silent failure, not a
loud one:

| Trap | Word | PowerPoint |
| --- | --- | --- |
| `DisplayAlerts` for "no alerts" | `0` (`wdAlertsNone`) | **`1`** (`ppAlertsNone`); `2` is *all*. Porting `0` selects an **undefined** value. |
| Export range | optional | **mandatory** `PrintOptions.Ranges` object |
| `Saved` flag | Boolean | **`MsoTriState`**. `$pres.Saved = $true` fails **8/8** with an invalid cast; `= -1` works **8/8** (`spikes/powerpoint/probes/probe-saved-flag.ps1`) |
| `Quit()` terminates the process | yes | **no — 15/15 cycles** |
| Bind target | `_WwG` | **`mdiClass`** |
| Second instance | new process | **attaches to the existing one** |

The `DisplayAlerts` one deserves emphasis because it is the exact shape the brief
predicted more of, and it is the worst kind. The Word host sets
`DisplayAlerts = wdAlertsNone` *before anything can open a document*, precisely
because "a modal prompt on a hidden Word is an unrecoverable hang, not an error"
(`src/word/word-host.ps1`). Porting the constant unchanged to PowerPoint sets an
undefined value on an invisible instance — so the failure mode of getting this
wrong is the very hang the setting exists to prevent.

I looked for more of this class and found none beyond the table. That is a
statement about what is *measured*, not about what exists.

### 7.2 About #139, and what the host must not inherit

The brief is accurate. `_common.ps1`'s `New-OwnedPowerPoint` snapshots `POWERPNT`
pids, calls `New-Object`, sleeps 300 ms, and takes the set difference as `Owned`.
`Close-OwnedPowerPoint` guards `Quit()` on `Owned.Count > 0` — the guard its own
SAFETY comment is about — but then unconditionally sweeps every pid in `Owned`
with `Stop-Process -Id $p -Force`.

The guard protects the wrong step. `Owned` is computed by differencing, and
differencing is **measured unsound** in this repo already: for Word,
`spikes/isolation/probes/probe-init-attribution.ps1` saw 2 new pids for 1 instance
created, with the correct one first by luck. For PowerPoint it is worse, because
`New-Object` normally creates *no* pid at all — so any `POWERPNT.EXE` the user
starts inside that 300 ms window is adopted as ours, `Quit()`-ed, and then
force-killed with their unsaved work in it. #139 and #136 are one defect class,
and PowerPoint is the instance of it where the blast radius is the user's data.

**The rule the host must carry:** the only sound PowerPoint pid is the one
`CreateProcess` returned to us in the separate-desktop launcher. Every other route
— differencing, `Get-Process`, enumeration — is a guess, and the correct response
to a guess is `$null` and a refusal to kill, exactly as `Get-AttributedWordPid`
already does for Word.

---

## 8. Spikes to run

Each needs a machine and the maintainer's consent. None may be run by a session
that has not been told PowerPoint is safe to start, and none should be run at all
until #139 lands.

| # | Question | Blocks | Priority |
| --- | --- | --- | --- |
| **S1** | Do hidden slides, builds/animations, transitions and speaker notes preserve the 1 slide to 1 page mapping? What does a build-animated slide flatten to? | W2 addressing | **High** — a count mismatch breaks addressing silently |
| **S2** | Export cost from an **isolated separate-desktop** instance: cold to first PDF, per-slide steady state, and whether the flatness holds. All existing numbers are from a COM-attached instance. | W4, idle-window choice | **High** — the display budget is unmeasured for the instance we must use |
| **S3** | Does `Copy-Item` succeed against a PowerPoint-held deck? Run the `spikes/isolation/probes/probe-fileshare-algebra.ps1` matrix against a PowerPoint holder to learn its access and share mode. | W1 read path, §7 taxonomy | **High** — currently inferred from Word |
| **S4** | Is `Slide.SlideID` / `p:sldId/@id` stable across insert, delete, reorder, save-and-reload, and a round trip through the user's PowerPoint? | W2, `delete_slide`, `reorder_slide` | **High** — decides handle vs coordinate |
| **S5** | Can a layout be selected language-independently (`CustomLayout` index, `SlideMaster.CustomLayouts`, or the layout part's `p:ph` set), given that layout *names* are localized? | `add_slide` | Medium |
| **S6** | Characterise `0x800706BA` as a **rate**, under the isolated-instance path specifically. Currently anecdotal and explicitly not usable as an argument. | retry policy, W4 | Medium — needed before shipping, not before starting |
| **S7** | Does `Presentations.Open(..., ReadOnly := msoTrue)` avoid the >71 s hang on a held deck? The measured hang used the isolated launcher's command-line open. | fallback path only | Low — unnecessary if S3 passes |
| **S8** | What does the extension actually cost after PowerPoint lands? Re-measure the envelope. | release gate | Low — see §10 |

**S1–S4 are the blocking set.** S5–S8 can run alongside implementation.

---

## 9. Proposed work items

Titles and one-paragraph scopes, sequenced. **Not filed** — for the coordinator to
gate.

### Can start immediately (no spike, no PowerPoint)

**W0 — De-Word the shared machinery, before there is a second consumer.** Rename
`src/word-lifecycle.mjs` to an application-neutral name, make
`RenderCache.wordVersion` application-tagged, and change `src/tool-error.mjs`'s
default code from `word_error` to something that does not assert an application.
Mostly rename, no new capability, cheap now and expensive after a second host
exists. The `word_error` default is the part that is a live correctness issue
rather than tidiness: a PowerPoint failure surfacing under that code names a cause
the code cannot know. *Depends on: nothing.*

**W1 — `read_presentation`, served entirely from the OOXML zip.** Parse
`ppt/slides/*.xml` and `ppt/notesSlides/*.xml` out of the `.pptx` and return a
structure map — slide index, slide id, `p:ph/@type` map, placeholder text, notes
text — under a revision token from the existing `src/revision-token.mjs`. Measured
at 82 ms cold and 275 ms against a held deck
(`spikes/powerpoint/probes/probe-bulk-read.ps1`), needs **no PowerPoint process
whatsoever**, and is therefore fully testable on `ubuntu-latest` against a
generated fixture. This is the largest capability increment available for the
least risk, and it ships before any host exists. Join keys are `p:ph/@type` and
placeholder index — never a name (§4.1). *Depends on: nothing. Confirmed by S4,
not blocked by it.*

**W3 — Correct ADR 0003's present tense, and record the host decision.** ADR 0003
states that Word, Excel and PowerPoint *ship* as one extension declaring a canvas
per application; one canvas ships. Either correct the tense or add a new ADR
recording the §3.2 decision — that the PowerPoint host is always an isolated
separate-desktop instance and that `New-Object -ComObject PowerPoint.Application`
is forbidden in product code. I lean to a new ADR, because "never attach" is a
decision with a rationale and a rejected alternative, which is what an ADR is for.
*Depends on: nothing.*

**W5 — Office-free source assertions for the PowerPoint constants.** Following
`test/unit/liveness-identity.test.mjs` and `test/unit/quit-argument.test.mjs`:
extract the constant values from the shipped PowerPoint host source and assert
them, so `ppAlertsNone = 1` cannot be silently reverted to Word's `0`, the
`Saved = -1` MsoTriState form cannot regress to `$true`, the export call cannot
lose its `PrintOptions.Ranges`, and no product file can contain
`New-Object -ComObject PowerPoint.Application`. Each is a defect a
licensed-PowerPoint test would catch only if someone ran it, and that a source
assertion catches in CI on `ubuntu-latest`. The last is the important one: it is a
machine-checkable statement of the §3.2 decision, and it can land with W0, before
the host it guards exists. *Depends on: W0 for the forbidden-`New-Object`
assertion; W4 for the rest.*

### Blocked on a spike

**W2 — Slide addressing and the read-then-address contract for decks.** Decide
handle-vs-coordinate at slide level from S4's result, define the shape address
(`p:ph/@type` plus occurrence), and document the caching rule the way ADR 0006
does for Word. If S4 shows `SlideID` stable, this is a materially better contract
than Word's and should be written up as such rather than quietly mirrored.
*Depends on: S4. Extends W1.*

**W4 — The isolated PowerPoint host.** A separate-desktop launcher plus `mdiClass`
`OBJID_NATIVEOM` bind, modelled on `spikes/isolation/probes/probe-bind.ps1` and
`spikes/powerpoint/probes/probe-second-process.ps1`; `ppAlertsNone = 1`; macros
force-disabled; export via `ExportAsFixedFormat` with a real `PrintOptions.Ranges`;
a geometry assertion on `/MediaBox` so a notes-page regression cannot pass;
teardown by graceful `Quit()` then kill-by-created-pid only, with
safe-mode-prompt dismissal on the next launch; `createIdleShutdown` reused
unchanged with a longer window. The hardest item here and the one with the most
ways to be quietly wrong. *Depends on: S2, S3, W0.*

**W6 — The `powerpoint-deck` canvas.** A second entry in `extension.mjs`'s
`canvases` array reusing `src/ui/pdf-view.mjs` and the vendored pdf.js unchanged,
with slide-shaped wording, a slide-list outline, and no `word-mark` analogue.
Small once W4 exists, and near-zero risk to the rendering path because nothing
about the renderer changes. *Depends on: W4, W2.*

**W7 — `set_placeholder_text`, the one editing operation the evidence supports.**
Open a temp copy in the isolated instance, address the placeholder by
`p:ph/@type`, replace its text, save, close — measured at 163–177 ms warm
(`spikes/powerpoint/probes/probe-lock.ps1`) — under the ADR 0005 transient-lock
model and the ADR 0006 revision token, both of which transfer unchanged. Ship with
`revert_presentation` (snapshot-and-rename, from `src/word/snapshots.mjs`) or not
at all. Explicitly **not** a general slide-mutation API; the tool description
should say what is out of scope so the agent stops looking. *Depends on: W4, W2.*

**W8 — `add_slide`, `delete_slide`, `reorder_slide`.** The three structural
operations with a defensible answer to the geometry question — `add_slide` because
a layout supplies it, the other two because they need none. All three move
index-based addresses, so all three need W2 settled first. *Depends on: S4, S5,
W7.*

**W9 — `create_presentation`.** The `create_document` analogue. Lowest value of
the set — the stated priority is display — and the highest exposure to §7.1, since
it touches `Saved`, `SaveAs`, alerts and layouts at once. *Depends on: W4, W8.*

### Suggested order

```
now:    W0, W1, W3      (parallel; none needs PowerPoint)
then:   S1 S2 S3 S4     (need a machine and consent; S3 and S4 are cheap)
then:   W2, W4          (W4 is the long pole)
then:   W5, W6          (display path complete at W6)
then:   S5, W7, W8
later:  S6, W9, S7, S8
```

The display path — the stated priority — is complete at **W6**, and W1 delivers a
real read capability before any of the risk in W4 is taken.

---

## 10. The size envelope

Measured on this checkout (CRLF working tree), 118 files:

| Area | Bytes |
| --- | --- |
| `src/ui/vendor/` (pdf.js plus the three worker parts) | 1,732,797 |
| `src/word/` | 285,560 |
| rest of `src/` | 203,632 |
| `test/unit/` | 731,354 |
| `test/integration/` | 215,929 |
| root (`extension.mjs`, manifest) | 41,215 |
| **total** | **3,210,487** |

Against `tools/validate-extensions.mjs`'s limits: **1,789,513 bytes of headroom**
to the 5,000,000 hard cap, and **789,513** before the warning band opens.

**Two corrections to the brief here.** First, the warning threshold is not a
picked 4,000,000 — it is `MAX_TOTAL_BYTES - MAX_FILE_BYTES`, derived so that
"below this, no single legal file can carry the total past the envelope; above it,
the very next one can". The value is right; the framing matters, because the file
explains at length why a third standalone number was removed. Second, the brief's
"recently measured around 3,120 KB" is ~16 KB from what this checkout holds, which
is the *expected* behaviour and the reason #82 removed the routine byte figure
from the validator: two clones of one commit measured ~23 KB apart on line endings
alone. The figure describes a checkout, not a tree, and should not be quoted
between sessions as if it identified one. My 3,210,487 carries the same caveat.

**The envelope is not the binding constraint for PowerPoint.** No new vendored
asset is required: the display path reuses the existing pdf.js worker byte for
byte, because `ExportAsFixedFormat` produces an ordinary PDF. Realistic additions
are a `src/powerpoint/` host of the same order as `src/word/` (~285 KB, and it
should be smaller — no structure-map machinery, since W1 lives outside it) plus
tests. Even a pessimistic 600 KB leaves ~1.19 MB.

What *would* bind is a second vendored renderer, and nothing in the evidence calls
for one. Re-measure at release (S8), and note that `test/` counts against the
envelope too — 29.5% of the tree today, which is the first place to look if the
number ever gets tight.

---

## 11. Risks and unknowns, ranked

| # | Risk | Status | Note |
| --- | --- | --- | --- |
| 1 | Attaching to the user's PowerPoint destroys their work | **measured** | `spikes/powerpoint/probes/probe-single-instance.ps1`, with a Word control. Mitigated by W0/W5 making `New-Object -ComObject PowerPoint.Application` a CI failure rather than a rule people remember. |
| 2 | Opening a held deck hangs forever with no dialog | **measured** | >71,000 ms, twice, `spikes/powerpoint/probes/probe-cross-instance-lock.ps1`. Mitigated by the temp-copy model — *if* S3 confirms the copy succeeds. |
| 3 | Export cost from an isolated instance | **unknown** | Every published figure is from a COM-attached instance (§2.3). S2. |
| 4 | Hidden slides or animations breaking the 1:1 page mapping | **unknown** | Nothing in the spike touches them. S1. |
| 5 | `Copy-Item` against a PowerPoint-held deck | **inferred** | From Word's share-mode algebra, not measured for PowerPoint. S3. |
| 6 | COM/file name divergence corrupting a join | **measured** | `spikes/powerpoint/probes/probe-localization.ps1`. Made *more* dangerous by PowerPoint accepting both names. Mitigated by never joining on a name. |
| 7 | `ppAlertsNone = 1` ported as Word's `0` | **measured** | Sets an undefined value on an invisible instance; the failure mode is the hang the setting prevents. W5 asserts it in CI. |
| 8 | ~12 s first-render stall | **measured** (n=2, 2 ms apart) | Mitigated by W1 serving structure with no host, and by starting the host ahead of time rather than lazily. |
| 9 | `0x800706BA` RPC failures | **unknown rate** | Real (Application-log faults), four causes ruled out by probe, not characterised. Must not be used as an argument until S6. |
| 10 | `SlideID` stability | **unknown** | Decides handle vs coordinate. S4. |
| 11 | Killed PowerPoint gives a modal safe-mode prompt on the next launch | **measured** | Broke a probe's control arm before it was diagnosed. Teardown must handle it. |
| 12 | Size envelope | **measured, ample** | 1,789,513 bytes free; no new vendored asset needed. |

---

## 12. What contradicts the brief

Listed plainly, since the coordinator asked to be corrected rather than agreed
with.

1. **"v0.1 ships Word support plus some basic PowerPoint groundwork."** No
   PowerPoint groundwork ships. Zero matches for `powerpoint`/`pptx`/`slide`/
   `presentation` in `extension.mjs` or non-vendor `src/`. `SUPPORTED` is
   Word-only; one canvas is declared. What exists is a completed spike and an ADR
   amendment — evidence, not product.

2. **The PowerPoint export timings do not describe the instance the architecture
   requires.** `spikes/powerpoint/probes/probe-export.ps1` used
   `New-OwnedPowerPoint`, i.e. COM attach. The 94–110 ms/slide and ~3.0–3.5 s
   cold-start figures are attach-path numbers. Isolated-instance export cost is
   unmeasured. This is the gap most likely to move the display-path estimate, and
   the brief did not anticipate it.

3. **`tools/validate-extensions.mjs` does not "warn above 4,000,000 bytes" as a
   chosen threshold** — it warns above `MAX_TOTAL_BYTES - MAX_FILE_BYTES`, which
   evaluates to 4,000,000 by derivation. The value is right; the framing is the
   point of that code.

4. **The ~3,120 KB figure should not be relayed between sessions.** #82 removed
   exactly that number from the validator's routine output because it measures a
   checkout, not a tree. This checkout measures 3,210,487 bytes.

5. **"Does one slide reliably yield one page?" is already answered — yes, exactly,
   with 0.000 pt geometry error** — but the *notes-page* trap the brief guessed at
   is real and sharper than expected: notes export also yields 13 pages, at a
   different geometry. Page count alone does not catch it; the geometry assertion
   does, and it should ship as an assertion rather than a comment.

6. **"Whether the host must refuse to act when the user has PowerPoint open" has a
   negative answer.** Under the never-attach design the user's instance is
   irrelevant to starting ours, and a held *deck* is a file-level problem the
   existing temp-copy model already solves. Refusing would be a self-inflicted
   limitation.

7. **`file_locked` needs no PowerPoint-specific rework at all.** The brief asked
   what PowerPoint distinguishes; the answer is "nothing new". Lock detection is
   0–1 ms and correct in both directions, and the exception types that split
   `file_locked` from `permission_denied` are filesystem-level. The taxonomy change
   PowerPoint forces is at the *host* layer (`powerpoint_unavailable`,
   `powerpoint_timeout`), not the file layer.

8. **Slide addressing may be genuinely better than Word's, not merely analogous.**
   ADR 0006's "coordinate, not a handle" is forced by Word having no stable
   paragraph identity. Slides plausibly *do* have one. Unmeasured — S4 — but if it
   holds, mirroring Word's weaker contract would be a mistake.

9. **#139 and #136 are one defect class, not two bugs.** Both are pid ownership by
   differencing, which `spikes/isolation/probes/probe-init-attribution.ps1` already
   measured unsound for Word. PowerPoint is the instance where the consequence is
   the user's unsaved work rather than a leaked process. Worth fixing as a class,
   and worth carrying into the host as a rule (§7.2) rather than only into the
   probes.
