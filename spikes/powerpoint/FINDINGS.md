# Spike: does PowerPoint behave the way Word does?

`docs/adr/0003-one-extension-many-canvases.md` commits the architecture to one
extension serving Word **and** PowerPoint, on the reasoning that both export to
PDF and can therefore share a rendering pipeline and a canvas. Excel was dropped
from that claim once its export was understood. Every measurement in this repo
came from Word; PowerPoint had never been probed. This spike probes it.

Environment: PowerPoint 16.0 build 20326, **German UI** (`LanguageID(UI)` = 1031),
Windows 11. Fixture is a generated 13-slide deck at 960 x 540 pt, matching the
13-page Word fixture so the numbers line up.

---

## Verdict

**ADR 0003's generalisation survives, but only after one of its premises is
rewritten.**

Rendering — the reason the two applications were grouped together — holds
completely. `ExportAsFixedFormat` gives exactly one PDF page per slide, page
geometry matches the slide size to 0.000 pt, and PowerPoint is *cheaper* than
Word per page. There is no rendering reason to split the extension, and no
reason to drop PowerPoint the way Excel was dropped.

What does not hold is the process model. **PowerPoint is single-instance.**
`New-Object -ComObject PowerPoint.Application` does not start a PowerPoint; it
attaches to whatever PowerPoint is already running, including the user's. So the
"hidden instance we own" that the design assumes does not exist by that route.
It does exist by another route — the separate-desktop launch that
`spikes/isolation` already uses for Word — but that is a different, more
expensive mechanism, and the ADR should say so rather than implying the two
applications are automated the same way.

Everything else in the Word findings transferred, including the one nobody
wanted: PowerPoint blocks on a held file exactly as Word does.

---

## Answers

### 1. Does `ExportAsFixedFormat` on a `.pptx` give one page per slide, page-accurately?

**Yes, exactly.** (`probe-export.ps1`)

| Check | Result |
| --- | --- |
| Pages produced for a 13-slide deck | **13** |
| `/MediaBox` vs `PageSetup.SlideWidth/Height` | **0.000 x 0.000 pt** delta on every page |

The PDF was inspected by parsing `/MediaBox` out of the file directly, with no
PDF library, so the number is the geometry the renderer actually wrote.

This is the gate question and it passes cleanly. One caveat worth knowing:
exporting with `ppPrintOutputNotesPages` instead of `ppPrintOutputSlides` also
produces 13 pages, but at 540 x 720 pt — portrait notes pages, not slides. The
page *count* alone would not have caught that; the geometry check did.

**Trap:** the `PrintRange` argument is not optional. `$null`,
`[Type]::Missing` and `[Reflection.Missing]::Value` all fail — the first with a
`NullReferenceException`, the last with "Missing parameter does not have a
default value". A real `PrintOptions.Ranges` object is mandatory even when
exporting the whole deck:

```powershell
$pres.PrintOptions.Ranges.ClearAll()
$null = $pres.PrintOptions.Ranges.Add($from, $to)
$pres.PrintOptions.RangeType = 4          # ppPrintSlideRange
$pres.ExportAsFixedFormat($out, 2, 2, 0, 1, 1, $pres.PrintOptions.Ranges, ...)
```

### 2. Does PowerPoint block on a held file the way Word does?

**Yes. Identically, and just as badly.** (`probe-cross-instance-lock.ps1`)

This answer is a correction. An earlier probe in this spike concluded "no
blocking" — and it was wrong, because its "second instance" was the *same*
instance re-opening its own file. Once `probe-single-instance.ps1` exposed that,
the test was redone against a genuinely separate process.

| Arm | Target | Result |
| --- | --- | --- |
| B  control | a **free** deck | bound in **12 144 ms**, 13 slides, read-write |
| A  test | the deck the first instance **holds** | **hung >71 000 ms**, killed externally |

Reproduced twice, with the control landing within 2 ms across runs (12 144 ms and
12 146 ms) and the test arm hanging both times.

In the test arm PowerPoint started, created its frame, created its `mdiClass`
document window — and then `Presentations.Count` stayed at **0** forever. It
showed **no dialog at all**. It did not fail, refuse, or fall back to read-only.
It sat there until it was killed by pid.

This is the same shape as the Word finding that forced `docs/adr/0005`: block
silently, no alert, external kill required. **The transient-lock model transfers
to PowerPoint unchanged, and for the same reason.**

External writers are blocked while a deck is held, too (`probe-lock.ps1`):
direct overwrite fails, write-temp-then-rename fails, and taking an exclusive
write handle fails. Lock detection by attempting a write handle costs **0-1 ms**
in both directions and is correct in both (Word: 4 ms held / 9 ms free).

### 3. Is there a bulk structural read, like `Content.WordOpenXML`?

**Not in the object model — but there is something better outside it.**
(`probe-bulk-read.ps1`)

| Arm | Method | Time |
| --- | --- | --- |
| A | naive per-shape COM walk, 39 shapes over 13 slides | 585 ms |
| B | one `Shapes.Range().TextFrame.TextRange.Text` per slide | 444 ms (1.3x) |
| C | read the `.pptx` as a zip and parse `ppt/slides/*.xml`, **no COM** | **82 ms** (7.1x) |
| C (held) | same, while PowerPoint holds the deck open | 275 ms |

PowerPoint has no `Content.WordOpenXML` equivalent, so arm B is the best the
object model offers and it is barely an improvement — the round trips are in the
shape traversal, not the text fetch. The real answer is that a `.pptx` *is* an
OOXML zip, so the structural read needs no PowerPoint at all.

Two things make arm C more than a micro-optimisation:

- It works **while PowerPoint holds the deck open** (275 ms; the file is
  readable, only writes are blocked), so it is not restricted to idle documents.
- It yields `p:ph/@type` — `title`, `ctrTitle`, `subTitle`, `body` — which the
  COM path does not expose as cleanly, and which is not localized. See question 6.

Word's contrast was 3724 ms naive vs 289 ms bulk, a 13x defect. PowerPoint's
naive walk is much less catastrophic in absolute terms, because a deck has tens
of shapes where a document has hundreds of paragraphs. The trap is the same
shape; it is just smaller here.

### 4. Can PowerPoint run genuinely hidden?

**Yes — but not by the route the design assumes.** (`probe-hide.ps1`,
`probe-single-instance.ps1`, `probe-second-process.ps1`)

Taken alone, the visibility results are good:

- `Application.Visible = msoFalse` **throws** ("Hiding the application window is
  not allowed") — but `Visible` already *reads* `msoFalse` for a COM-created
  instance, so the assignment is unnecessary rather than impossible.
- `Presentations.Open(..., WithWindow := msoFalse)` gives **0 document windows**
  and **0 visible top-level windows** (11 hidden ones).
- Both windowless and windowed instances survived 45 s idle without
  self-terminating.

Re-measured on the isolated `CreateProcess` instance (`probe-hide.ps1`, the H1/H3
writes #141 restores), the picture is confirmed and sharpened. The write is still
refused — `Application.Visible = msoFalse` comes back a `COMException`, reported
by exception *type* since the message is German on this machine. But the isolated
instance, launched as a real process on a private desktop, *reads* `Visible =
msoTrue` (`-1`), where the COM-created instance above reads `msoFalse` — so on the
isolated route the assignment is neither unnecessary nor possible, and H1 offers
no way to hide. The documented fallback H3 does work: `Application.WindowState =
ppWindowMinimized` (`2`) is **accepted** on the isolated instance (read back `3` →
`2`). The probe captures the original `WindowState` before writing — and writes
only if that capture succeeded, so a restore is always possible; a failed capture
skips the write rather than landing one it cannot undo — then restores it on the
happy path *and*, demonstrated with an injected fault (`-InjectFailure`), in a
`finally` that runs on the failure path. The "restore outstanding" state is a
dedicated flag, not a `$null` sentinel, because `$null` would conflate "never
captured" with "already restored" and let an uncaptured write proceed. "Written
and never restored" onto an instance was the defect #141 corrects, and a restore
that only runs when nothing threw — or one landed with nothing to put back — is
that defect wearing a `finally`.

The problem is upstream of visibility. **There is only ever one PowerPoint.**

| Test | PowerPoint | Word (control) |
| --- | --- | --- |
| Two `New-Object` calls in one process | **attached**, 1 process | 2 processes |
| `New-Object` from a separate process | **attached**, 1 process | n/a |
| Launch `POWERPNT.EXE` on the same desktop | **exits**, hands off | n/a |

The Word control is what makes this a statement about PowerPoint rather than
about the method: the identical test produced two independent `WINWORD.EXE`
processes. A second call also saw the deck the first had opened, and could name
it — one object model, not two.

The consequences are concrete. A COM-created "hidden instance" is the user's
PowerPoint: `Quit()` would close their decks, killing "our" pid kills theirs,
and application-level settings such as `DisplayAlerts` leak into their session.

**There is one escape, and it works** (`probe-second-process.ps1`): launch
`POWERPNT.EXE` into a separate window-station desktop and bind through
`AccessibleObjectFromWindow` / `OBJID_NATIVEOM` — the same technique
`spikes/isolation/probes/probe-bind.ps1` uses for Word. Proven isolated: the
bound instance had `deck.pptx` open while the shared instance reported
`Presentations.Count = 0`.

The bind target differs from Word's. Word binds on the document surface `_WwG`;
PowerPoint binds on **`mdiClass`**. All 23 descendant windows of `PPTFrameClass`
were offered to the binder; `MsoCommandBar` and `RICHEDIT60W` also answer, but
their `.Application` is a stub with an empty `Version` and no `Presentations`.
Only `mdiClass` hands back an object that can name the open deck. "It answered
the binder" is not the test; "it can name the deck" is.

### 4a. Does attribution need a window? `Application.HWND` and `ActiveWindow`

Two claims had been imported to PowerPoint rather than measured on it, and both
touch how a PowerPoint COM object is attributed to a pid. `probe-app-hwnd.ps1`
measures them on the isolated instance.

- **`Application.HWND` exists on PowerPoint but does not marshal through the
  `OBJID_NATIVEOM` bind.** Reflection resolves the member: `InvokeMember('HWND',
  GetProperty)` comes back a `COMException`, not a `MissingMemberException`, so
  the IDispatch name is real — consistent with the aside in the Word probe
  `spikes/isolation/probes/probe-word-ownership-hwnd.ps1` that PowerPoint has an
  `Application.Hwnd` and Word does not. But read through the bound object,
  `$app.HWND` reads **null through the bind** — no usable handle is marshalled —
  and so does the `Hwnd` of a live `ActiveWindow`. (Whether a non-terminating
  error accompanies that null is deliberately not asserted: the probe runs under
  `$ErrorActionPreference = 'Continue'`, so such an error would never reach a
  `catch`, and `$Error` growth is not a reliable detector. That last point is
  measured only for the **managed** .NET property adapter, in
  `probe-processname-after-exit.ps1`, which converts a throw to `$null` without
  `$Error` growing; whether the **COM/IDispatch** adapter does the same is
  unmeasured in this tree. So the null is reported as an observation, with no
  claim about what, if anything, accompanies it.)
  The real `PPTFrameClass` window handle *does* exist and *does* attribute our
  `CreateProcess` pid through `GetWindowThreadProcessId` — so attribution comes
  from the external window enumeration the isolated route already performs, not
  from an HWND read on the automation object. Because the member resolves, the
  Word probe's aside is **not** falsified and is left unchanged; the nuance is
  that the property reads null through this binding, not that it is absent from
  the product.
- **PowerPoint's `ActiveWindow` returns `$null` — not a window, and with no
  caught exception — when no presentation is open.** With `Presentations.Count =
  0` and the process alive, `$app.ActiveWindow` read `$null` in **5 of 5 fresh
  isolated instances** (`null/window/threw/inconclusive = 5/0/0/0` — the tally the
  probe now commits to its own output). The arm loops over **cycles, not reads**:
  each cycle is a fresh `Start-IsolatedPowerPoint` → close the deck → read → sweep,
  so the result is reproduced *across processes* rather than sampled once from one
  — a stack of reads of a single instance would only show that instance is stable,
  which is not the claim. A `$null` is counted as a windowless read only when the
  cycle clears **both** prerequisites: the process is alive on **both** sides of the
  read (so a `$null` from an RCW whose process exited mid-call is not miscounted),
  **and** the cycle reached a verified `Presentations.Count = 0`. The deck-closing
  `Close()` sits in a swallowing `catch`, so it can silently leave a deck; a `$null`
  read whose deckless state was not established would assert a cause the code never
  proved, so it is scored INCONCLUSIVE rather than counted. All 5 cycles cleared both
  gates (each printed `pres=0`), so every counted `$null` is earned, not merely
  observed. This cross-instance tally
  agrees with, and supersedes, an earlier three-run reading. The reading rests on
  an explicit `$null` check that tells a returned `$null` apart from a returned
  window; an earlier reading that reported a live window was an artifact of a
  branch that never drew that distinction: `$null.Hwnd` does not throw in this
  shell (no `Set-StrictMode` under `spikes/powerpoint/`), so a `$null` printed
  identically to a window. The
  Word claim — that `ActiveWindow` *throws* without a document — is true and
  measured for *Word* (`spikes/isolation`), where it is a design constraint.
  PowerPoint does the third thing: it does not hand back a usable window and no
  exception is caught, it yields `$null`. (As above, "no exception caught" is the
  measured claim, not "no exception" — the probe runs under `Continue`, so a
  non-terminating error would never reach a `catch`.) Importing the Word result to
  PowerPoint would be wrong.

Neither result changes the isolation design, because attribution (which pid is
this COM object) is not exclusivity (is it safe to write to it). The isolated
`CreateProcess` route is required because PowerPoint is single-instance and
`New-Object` attaches — measured in `probe-single-instance.ps1` — regardless of
how `HWND` or `ActiveWindow` behave.

### 5. Export cost per slide and per deck; cold start

(`probe-export.ps1`, `probe-stability.ps1`)

| Measurement | PowerPoint | Word |
| --- | --- | --- |
| Single-page export | **94-110 ms** | 168 ms |
| Whole 13-page/slide export | 815-865 ms (~64 ms/slide) | 664 ms |
| Cold start, fresh instance to first PDF | ~3019-3494 ms | not comparable (see below) |
| Open + edit + save + close, warm | **163-177 ms** | not comparable (see below) |
| Edit + 1-slide re-export | 137-141 ms mean, 93 ms min | n/a |

The two Word timings for cold start and open-edit-save-close are deliberately
not quoted. They came from separate Word probes, not measurements made alongside
these PowerPoint runs, so neither supports a controlled comparison. The omitted
cold-start result was a single measurement of one-off engine load on Word's
first export in a fresh process, not process creation.

PowerPoint cold start decomposes as roughly 1000 ms to create the COM object,
250 ms to open the deck, and 1562-2017 ms for the first export.

Single-slide export is flat: slide 7 costs the same as slide 1, so a
render-on-demand canvas does not pay for position in the deck.

### 6. Are layout and style names localized?

**Yes — and with a twist Word does not have.** (`probe-localization.ps1`)

Layout names are localized outright: `Titelfolie`, `Titel und Inhalt`,
`Abschnittsüberschrift`, `Zwei Inhalte`, `Vergleich`, `Nur Titel`.

The twist is shape names, where **COM and the file on disk disagree**:

| Shape | COM `Shape.Name` | stored `cNvPr/@name` | stored `p:ph/@type` |
| --- | --- | --- | --- |
| title | `Title 1` | `Titel 1` | `title` |
| body | `Text Placeholder 2` | `Textplatzhalter 2` | `body` |
| textbox | `TextBox 3` | `Textfeld 3` | *(none)* |

A control arm settles what is happening: renaming a shape through COM to
`ZZ_probe_custom_name` round-trips to disk verbatim, while the untouched shapes
stay German in the file. So the COM `Name` is a live property, and PowerPoint is
translating only the *default* placeholder names into English on read.

This matters precisely because question 3 recommends the OOXML path for
structural reads. **A host that reads structure from the zip and writes through
COM will see different shape names from each side, silently.** The join key must
be `p:ph/@type` (stable) or the placeholder index, never the name.

Unlike Word, PowerPoint is forgiving about lookups — `Shapes.Item('Title 1')`
and `Shapes.Item('Titel 1')` both succeed, where `Selection.Style = "Heading 1"`
throws in German Word. That is worse, not better: the failure does not announce
itself.

The language-independent alternatives all work: `Shapes.Placeholders[i]`,
`PlaceholderFormat.Type` (an enum: 1 = title, 2 = body), and `p:ph/@type` in the
XML. The rule is simply **never address a placeholder or layout by name**.

---

## Traps found, for whoever ports the Word host

| Trap | Word | PowerPoint |
| --- | --- | --- |
| `DisplayAlerts` value for "no alerts" | `0` (`wdAlertsNone`) | **`1`** (`ppAlertsNone`); `2` is *all*. Porting `0` selects an undefined value. |
| `PrintRange` / export range | optional | **mandatory** `PrintOptions.Ranges` object |
| `Saved` flag | Boolean | **`MsoTriState`**. `$pres.Saved = $true` fails **8/8** with an invalid cast; `= -1` works **8/8** (`probe-saved-flag.ps1`) |
| `Quit()` terminates the process | yes | **no — 15/15 cycles** left `POWERPNT.EXE` running |
| Bind target for an isolated instance | `_WwG` | **`mdiClass`** |
| Second instance | new process | **attaches to the existing one** |

And one that only shows up in cleanup: because `Quit()` does not terminate,
killing is the fallback — and a killed PowerPoint makes the **next** launch show
a modal safe-mode prompt ("PowerPoint konnte beim letzten Mal nicht gestartet
werden"). This silently broke the control arm of the lock probe until it was
diagnosed. Graceful `Quit()` first, kill only as a fallback, and dismiss the
prompt with `WM_COMMAND`/`IDNO` if it appears.

---

## Where this leaves ADR 0003

The decision stands. The reasoning behind one of its premises does not.

ADR 0003 groups Word and PowerPoint because both export to PDF. That is
**confirmed** — more strongly than for Word, since PowerPoint's page-per-slide
mapping is exact and its per-page cost is lower. Excel was dropped because its
export does not have this property; PowerPoint should not follow it.

The amendment is to the process model, which the ADR treats as shared and is
not:

> **Proposed amendment to ADR 0003.** The claim "one hidden Office instance per
> application, created via COM" holds for Word and **not** for PowerPoint.
> PowerPoint registers a single automation server per session:
> `New-Object -ComObject PowerPoint.Application` attaches to the user's running
> PowerPoint, and launching `POWERPNT.EXE` on the current desktop hands off and
> exits. A private instance requires the separate-desktop launch and
> `OBJID_NATIVEOM` bind already used in `spikes/isolation` — binding on
> `mdiClass` rather than Word's `_WwG`. The ADR should state that isolation is
> obtained by *different mechanisms* per application, and that the PowerPoint
> mechanism costs ~12 s to reach a bound instance versus ~1 s for a COM attach.

Two consequences worth pulling out, because they are not obvious from the
amendment alone:

- **Never call `Quit()` on a COM-attached PowerPoint, and never kill its pid.**
  Both destroy the user's session. This is not a hygiene rule for probes; it is a
  correctness rule for the host.
- The ~12 s bind cost means a PowerPoint canvas cannot open an isolated instance
  lazily on first interaction without a visible stall. It should either be
  started ahead of time or the first render should come from the OOXML/PDF path.

`docs/adr/0005`'s transient-lock model transfers **unchanged**. Question 2 was
the one place where a materially better story was plausible, and the evidence
went the other way.

---

## Unresolved

An intermittent `0x800706BA` "RPC server is unavailable" appeared across several
probes. It is real, not a scripting artefact: the Windows Application log carries
matching `POWERPNT.EXE` faults, exception `0xc0000005` in `combase.dll` at offset
`0x19e198`. Ruled out by probe: notes-page export (`probe-notes-control.ps1`,
both arms died), warm-vs-fresh instance reuse (`probe-stability.ps1`, 0 failures
in both arms on a clean run), idle self-termination (`probe-hide.ps1`), and the
`Saved = $true` cast error (`probe-saved-flag.ps1`, a deterministic and separate
fault). The remaining hypothesis is COM teardown racing PowerPoint's own
shutdown, which the single-instance finding makes more plausible but does not
establish. **It is not characterised as a rate**, so it should not yet be used to
argue for or against anything.

---

## Probes

Run `make-fixture.ps1` first; everything else is independent and re-runnable.

| Probe | Question |
| --- | --- |
| `make-fixture.ps1` | generates the 13-slide fixture |
| `probe-export.ps1` | Q1, Q5 — page-per-slide, geometry, export cost |
| `probe-single-instance.ps1` | Q4 — single vs multi instance, with a Word control |
| `probe-second-process.ps1` | Q4 — can a private process exist, and can we bind to it |
| `probe-cross-instance-lock.ps1` | Q2 — blocking on a held file, with a free-file control |
| `probe-lock.ps1` | Q2 — external writers, lock detection cost |
| `probe-bulk-read.ps1` | Q3 — COM walk vs OOXML zip |
| `probe-localization.ps1` | Q6 — layout, placeholder and shape names |
| `probe-hide.ps1` | Q4 — visibility (H1/H3 on the isolated route, restored) and idle survival |
| `probe-app-hwnd.ps1` | attribution — `Application.HWND` and `ActiveWindow` on PowerPoint (isolated route) |
| `probe-saved-flag.ps1` | the `Saved` MsoTriState trap |
| `probe-stability.ps1`, `probe-notes-control.ps1` | crash investigation, both negative; FRESH arm re-measures `Quit()` reaping on the isolated route |

Process hygiene, corrected by issue #139. These probes used to snapshot the
`POWERPNT.EXE` pid set and treat the difference as processes they had created —
then `Quit()` and force-kill it. That is unsound in both halves: `New-Object`
attaches rather than starts (see the single-instance finding above), and
differencing over-reports (`probe-init-attribution.ps1`, 2 new pids for 1
instance). The census is now a **report only**; no COM-obtained PowerPoint is
quit or killed; the only kills are on `CreateProcess` pids through
`Stop-VerifiedPpt`, which verifies process name and recorded `StartTime` before
acting. The probes therefore no longer terminate a PowerPoint they did not
start — but they may leave one running, and they still write `DisplayAlerts` on
an instance they may have attached to. See `README.md`, *Process safety*.

`probe-stability.ps1` was affected as an instrument. Its FRESH arm was defined as
"open, export, close, **quit**" and its 15/15 figure above counted cycles where
`Quit()` failed to reap the process, with a kill between cycles making the next
one fresh. #139 removed both the quit and that kill, so for a time the arm could
not establish a new process per cycle and could not reproduce that measurement.
The original figures were measured on a clean machine, where the census genuinely
was empty, and they stand.

#142 rebuilt the arm onto the `CreateProcess` route in `_isolated.ps1`. Each
cycle now launches through `Start-IsolatedPowerPoint` — a pid the kernel handed
us, provably not the user's — exports, closes the deck, calls `Quit()`, and polls
whether the process reaps. The between-cycle sweep is `Stop-IsolatedPowerPoint`,
which routes any survivor through `Stop-VerifiedPpt` — a kill gated on the
recorded pid **and** `StartTime` that declines rather than guessing — never a
census difference. Re-measured this way the arm reproduces the original exactly,
and neither figure is discarded:

| Measurement | Conditions | `Quit()` failed to reap |
| --- | --- | --- |
| original | `New-Object` per cycle, clean machine, empty census, kill between cycles | **15/15** |
| #142 re-measure | `CreateProcess` isolated instance per cycle, verified between-cycle sweep | **15/15** (0 reaped, 0 threw, 0 self-exited — all 15 reached `Quit()`, it returned normally each time, the process survived each time; exports 15/15 clean, mean ~2.7 s; no OS fault logged; POWERPNT census empty before and after) |

The two agree: the original was sound as taken, and the redesigned arm confirms
it on an instance whose ownership is not in doubt, with a sweep that left nothing
running.
