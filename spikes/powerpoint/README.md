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

- `_common.ps1` — output helpers, pid-ownership tracking, and `/MediaBox`
  parsing that reads PDF page geometry without a PDF library.
- `_isolated.ps1` — launches a genuinely private PowerPoint on its own window
  station desktop and binds to it through `OBJID_NATIVEOM`. Needed because
  `New-Object -ComObject PowerPoint.Application` attaches to the user's
  PowerPoint instead of starting one.

## Process safety

These probes were written to run alongside other sessions driving Office.

- Every probe snapshots `POWERPNT.EXE` and `WINWORD.EXE` pids before it starts
  and only ever kills the difference.
- `Close-OwnedPowerPoint` calls `Quit()` **only** when a new pid was actually
  created. Since PowerPoint is single-instance, quitting an attached instance
  would close the user's decks.
- `_isolated.ps1` only ever touches the pid `CreateProcess` returned to it.
- Fixtures are copied into `%TEMP%` per run, so two probes never contend for the
  same path.

One consequence is worth knowing before running these by hand: `Quit()` does not
terminate `POWERPNT.EXE` (15/15 cycles), so a kill is the fallback — and a killed
PowerPoint makes the *next* launch show a modal safe-mode prompt.
`_isolated.ps1` quits gracefully first and dismisses the prompt with
`WM_COMMAND`/`IDNO` if it appears anyway.
