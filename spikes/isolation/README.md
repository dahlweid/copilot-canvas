# Isolation and editing spike

Research for a possible v2 of the Word canvas: can a live Word instance be made **completely
invisible** to the user while still rendering, and can it be driven well enough to support
editing?

`PLAN.md`, the 2,059-line write-up, has been **removed** (#103). It was the build plan for
the design this spike falsified, and the product went another way. What it had measured is
summarised below; the file itself is recoverable from git history at `v0.1.0` (`93c3536`).

## Conclusion in one line

Full-window pixel streaming works and is invisible, but a **static colour-correct PDF for
display plus a hidden live Word instance as an edit channel** (option C) beats it, because
single-page re-export costs 168 ms and streaming loses text selection, Ctrl+F, print and
screen-reader accessibility.

## What the plan had measured

The numbers that outlived it. Where a probe is named below it still ships and can be re-run.
The two lock rows and the addressing row name none: the scripts that produced them were cited
by the plan and by nothing else, so they were removed with it. Their conclusions are what
survive, in ADR 0005 and in `word-host.ps1`; the scripts are in git history at `v0.1.0`.

| Finding | Where it lives now |
| --- | --- |
| Single-page PDF re-export costs **168 ms** and is flat in page position; a full 13-page export costs 664 ms (`probe-export.ps1`) | the conclusion above, and the render cache |
| Word raises **no** content-change event; polling `doc.Saved` costs 1.14 ms (`probe-events.ps1`) | the file watcher |
| Holding the original document open **blocks**: all three external write patterns fail with sharing violations, a second Word opening the same path hangs indefinitely with `DisplayAlerts` already off, and `~$` appears in the user's folder | ADR 0005 |
| So the lock must be transient, and "already open" must be detected without calling `Documents.Open` — by taking a write handle, 4 ms when held, 9 ms when free | ADR 0005, `Test-FileWritable` |
| A structural read must not walk paragraphs: on 219 paragraphs, per-paragraph `Range.Text`/`OutlineLevel` cost **3724 ms** against **289 ms** for one `Content.WordOpenXML` that returns strictly more | inlined at `Cmd-Structure` in `word-host.ps1` |
| The revision token, a SHA-256 prefix computed in 3 ms, changes on our own save and on external regeneration but **not** on a save with nothing dirty | inlined in `revision-token.mjs` |
| Style IDs inside the markup are **localized** — the first heading came back as `Überschrift1`, not `Heading1` | ADR 0006, and the structure parser |
| "Unreadable" is two conditions and both runtimes tell them apart by type — `EBUSY`/`IOException` against `EPERM`/`UnauthorizedAccessException` (`probe-errno-mapping.mjs`) | ADR 0006, `document-reader.mjs` |
| `Quit(<arg>)` does not bind under Windows PowerShell 5.1: it throws, the swallowing catch hides it, and process exit does not reap the survivor (`probe-quit0-leak.ps1`) | `quit-argument.test.mjs`, which pins it in the tree |

Two figures are deliberately **not** carried over: the 4547 ms first export in a fresh Word
process and the 228 ms open-edit-save-close round trip. Both are activation-inclusive —
they measure a cold process, not the operation — and quoting them as operation costs is the
mistake #35 exists to correct.

## Probes

Each script is self-verifying, self-cleaning, and prints its own measurements. They snapshot
`WINWORD` PIDs before starting and only ever kill processes they created, so they are safe to
run with Word already open.

| Probe | Question it answers |
| --- | --- |
| `probe-desktop.ps1` | Does a Word window on a `CreateDesktop` desktop that is never switched to still **render**? (yes) |
| `probe-visibility.ps1` | Is it genuinely invisible — taskbar, Alt+Tab, `EnumWindows`, shell? (yes: 18 windows, 0 visible) |
| `probe-hide.ps1` | Does cross-process `WS_EX_TOOLWINDOW` work as a fallback? (yes) |
| `probe-bind.ps1` | Can COM bind to *our* instance only, never the user's? (yes, via `OBJID_NATIVEOM`) |
| `probe-bench.ps1` | Capture frame rate, cross-desktop `Documents.Open`, dark-theme flip |
| `probe-darkmode.ps1` | Where does Word's dark page setting live? (**unresolved**) |
| `probe-hittest.ps1` | Click→position→type→undo round trip on the hidden desktop (works) |
| `probe-caret.ps1` | Is collapsed-**caret** geometry available, and does it round-trip? (yes, `w=0`, 0 chars of error) |
| `probe-events.ps1` | Does Word raise a content-change event? (**no** — poll `doc.Saved` at 1.14 ms instead) |
| `probe-export.ps1` | What does a single-page PDF re-export cost? (**168 ms, flat in page position**) |

## Running them

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\probes\probe-export.ps1
```

Most probes need a `.docx` fixture at `$env:TEMP\desktop-probe.docx`. Any document works; the
measurements quoted above used a 13-page one. Copy one there first:

```powershell
Copy-Item .\path\to\any.docx "$env:TEMP\desktop-probe.docx"
```

## Why the probes exist at all

Three separate confident, well-cited claims — two from research agents and one from this
repository's own earlier `spikes/live-word/FINDINGS.md` — turned out to be **wrong** when
measured. Notably: that cross-process `SetWindowLong` could not work, that a non-input desktop
could not render, and that the streamed page was colour-accurate.

Treat secondary sources here as hypotheses. Every claim above that still matters has a probe
next to it that produced the number.
