# PowerPoint probes

Evidence for issue #11: *Probe PowerPoint before assuming Word's findings
generalise.*

**Read [`FINDINGS.md`](FINDINGS.md) for the answers and the numbers.** This file
just says how to run the probes.

## Running them

```powershell
cd spikes\powerpoint
powershell -NoProfile -ExecutionPolicy Bypass -File .\probes\make-fixture.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\probes\probe-export.ps1
```

`make-fixture.ps1` generates `.fixtures\deck.pptx`, a 13-slide deck at
960 x 540 pt chosen to match the 13-page Word fixture so the numbers compare.
The fixture is generated rather than committed — `.fixtures/` is gitignored, and
no binary belongs in this repo. Every other probe takes an optional `-Fixture`
path and otherwise finds it there.

Each probe is independent and re-runnable. There is no ordering requirement
beyond generating the fixture first.

| Probe | Question it answers |
| --- | --- |
| `make-fixture.ps1` | — generates the fixture |
| `probe-export.ps1` | Q1, Q5 — one page per slide? page-accurate? what does it cost? |
| `probe-single-instance.ps1` | Q4 — is PowerPoint single-instance? (Word control) |
| `probe-second-process.ps1` | Q4 — can a private process exist, and can we bind to it? |
| `probe-cross-instance-lock.ps1` | Q2 — does it block on a held file? (free-file control) |
| `probe-lock.ps1` | Q2 — external writers, lock detection cost. **T3/T3c invalid, see its header** |
| `probe-bulk-read.ps1` | Q3 — COM walk vs reading the OOXML zip |
| `probe-localization.ps1` | Q6 — layout, placeholder and shape names on German Office |
| `probe-hide.ps1` | Q4 — visibility, windowless open, idle survival |
| `probe-saved-flag.ps1` | the `Saved` MsoTriState trap |
| `probe-stability.ps1` | crash investigation — fresh vs warm instances |
| `probe-notes-control.ps1` | crash investigation — notes-page export, exonerated |

`probe-cross-instance-lock.ps1` takes ~2.5 minutes: one arm is a deliberate hang
under a 70 s timeout.

## Shared files

- `_common.ps1` — output helpers, the census/report split described below, the
  one sanctioned kill (`Stop-VerifiedPpt`), and `/MediaBox` parsing that reads
  PDF page geometry without a PDF library.
- `_isolated.ps1` — launches a genuinely private PowerPoint on its own window
  station desktop and binds to it through `OBJID_NATIVEOM`. Needed because
  `New-Object -ComObject PowerPoint.Application` attaches to the user's
  PowerPoint instead of starting one.

## Process safety

These probes used to claim they "only ever kill the difference" between a
`POWERPNT` pid census taken before and after attaching. That claim was wrong in
the one direction that matters, and issue #139 removed it:

- `New-Object -ComObject PowerPoint.Application` **attaches**. The instance the
  probe holds is routinely the user's, with their unsaved decks in it.
- Differencing a census over-reports — `probe-init-attribution.ps1` measured 2
  new pids for 1 instance created — and is non-empty by construction whenever
  anyone else starts PowerPoint during the census window.

So the old guard `if ($ctx.Owned.Count -gt 0) { Quit; Stop-Process }` passed
*precisely* in the race it existed to catch. What the probes do now:

- **The census is a report, never an authorisation.** An empty difference is
  evidence we started nothing; a non-empty difference is evidence of nothing at
  all. `NewPids` is printed, never acted on.
- **A COM-obtained PowerPoint is never quit and never killed.** No signal at
  this layer can establish that we created it, so the honest classification is
  *unproven* and the honest action is to release the RCW and report.
- **The only kills are on pids returned by `CreateProcess`**, through
  `Stop-VerifiedPpt`, which re-checks `ProcessName` and the `StartTime` recorded
  at launch and returns `declined:<reason>` rather than guessing.
- Each probe closes only the presentations it opened itself, in a `finally`, and
  never enumerates `Presentations`.
- Fixtures are copied into `%TEMP%` per run, so two probes never contend for the
  same path.

What that buys, and what it does not. These probes no longer terminate or quit a
PowerPoint they did not start. They **may still leave one running** — `Quit()`
is measured not to reap `POWERPNT` (15/15 cycles) and the fallback kill on an
attached instance is gone — and they **may still change application-level
settings** on an instance they attached to: `DisplayAlerts` is set on every
instance obtained through `New-PowerPointInstance`, which is a write into the
user's live session if that is whose instance it is.

A killed PowerPoint makes the *next* launch show a modal safe-mode prompt.
`_isolated.ps1` quits gracefully first — its instance came from `CreateProcess`,
so it genuinely owns it — and dismisses the prompt with `WM_COMMAND`/`IDNO` if it
appears anyway.
