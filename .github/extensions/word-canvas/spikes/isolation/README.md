# Isolation and editing spike

Research for a possible v2 of the Word canvas: can a live Word instance be made **completely
invisible** to the user while still rendering, and can it be driven well enough to support
editing?

`PLAN.md` is the write-up. Read §1's callout first — it points at the conclusion, which is
**not** the one the plan starts out arguing for.

## Conclusion in one line

Full-window pixel streaming works and is invisible, but a **static colour-correct PDF for
display plus a hidden live Word instance as an edit channel** (option C, §11) beats it, because
single-page re-export costs 168 ms and streaming loses text selection, Ctrl+F, print and
screen-reader accessibility.

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
measurements quoted in `PLAN.md` used a 13-page one. Copy one there first:

```powershell
Copy-Item .\path\to\any.docx "$env:TEMP\desktop-probe.docx"
```

## Why the probes exist at all

Three separate confident, well-cited claims — two from research agents and one from this
repository's own earlier `spikes/live-word/FINDINGS.md` — turned out to be **wrong** when
measured. Notably: that cross-process `SetWindowLong` could not work, that a non-input desktop
could not render, and that the streamed page was colour-accurate.

Treat secondary sources here as hypotheses. If a claim in `PLAN.md` matters, there is a probe
next to it that produced the number.
