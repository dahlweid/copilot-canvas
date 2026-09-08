# Review guidance for this repository

This repository drives **installed Microsoft Office applications through COM** to
read, author and edit documents, and displays the result in a Copilot canvas.
Almost every unusual thing in this codebase exists because a probe measured the
platform behaving in a way the obvious code does not survive.

Read `CONTEXT.md` for vocabulary and `docs/adr/` for decisions. The probes that
established each fact are committed under `spikes/` and are re-runnable.

## The standard this repo holds

**A claim about platform behaviour must be backed by a probe that was actually
run.** If you flag something and the code cites a measurement, treat the
measurement as the stronger evidence and say so rather than restating the
general rule. If you believe a cited measurement does not support the conclusion
drawn from it, that is a high-value finding — say it plainly.

**A measurement that can come back empty must say how much it matched.** The
probe that found nothing and the probe that ran wrong produce the same output —
silence — and silence reads as agreement with whatever you already expected.
This has bitten in four separate tools: a `Select-String` for
`^# (tests|pass|fail)` over test output, which matched nothing because the
summary prefix depends on **both** the Node version and whether stdout is a
terminal; a `.Contains('on')` against `gh` output that searched for a whole
array element rather than a substring, because PowerShell line-splits
native-command output; a regex anchor carried across an **engine boundary**,
our harnesses straddling two — a PowerShell harness applying anchors, driving
a Node suite whose tests apply their own; and a `Select-Object -First 3` over
a diff that showed three comment lines and hid the fourth, which was the one
that mattered. None of them errored. All of them looked like measurements.

The first is the sharpest, because nobody wrote anything careless. Measured,
`node --test` with no `--test-reporter`:

|              | stdout piped       | real console       |
| ------------ | ------------------ | ------------------ |
| Node 22.23.2 | `# tests 1` (TAP)  | `ℹ tests 1` (spec) |
| Node 24.18.0 | `ℹ tests 1` (spec) | `ℹ tests 1` (spec) |

Node 22 picks its default from `process.stdout.isTTY` when the test runner's
reporter utilities load; Node 24 has no choice left to make. So the prefix
depends on the version **and** on the terminal, and neither alone accounts for
it. An earlier draft of this rule said it was "not TTY detection", on the
strength of a probe that set `isTTY` from a preload and saw no change. That
probe could not have shown anything: the default is a module-level `const`,
already evaluated by the time a preload assigns to it, so the observation was
the one the hypothesis predicted either way. Only a different class of input
settles it — a real console, where Node 22 emits spec while the same version
piped emits TAP. `validate.yml` pins Node 22 and passes no reporter flag, and
CI pipes, so `^# tests` matches today, matches nothing the day CI moves to 24,
and the run stays green either way. A probe that needs a stable prefix should
name its reporter.

The other two carry the same shape in miniature. `.Contains('on')` is `False`
on the `System.Object[]` that two lines of native-command output produce and
`True` on the `System.String` that one line produces, so the filter's meaning
depends on how much output there was — and it starts working as the output
shrinks, which is a false all-clear on exactly the small input someone reduces
to when isolating a problem. And `(?m)^alpha$` under .NET matches LF input but
not CRLF, because .NET puts `$` only before `\n` — a statement about newline
handling, which does not deny the independent end-of-input position — while
`/^alpha$/m` in JavaScript matches both, because ECMAScript puts it before any
LineTerminator and `\r` is one. A mechanism verified in one engine is not
verified in the other.

So before trusting a result, ask which two states the check has to tell apart.
**A check whose value does not differ between the state it accepts and the
state it rejects is not a check.** That question picks the instrument where a
fixed recipe does not, and the headline above is the commonest case of it
rather than the whole of it:

- To separate a **broken instrument** from a real absence, give the probe a
  **positive control** — an input you know matches.
- To separate **truncation** from completeness, assert the **count** against
  the whole input, not merely that some rows came back.
- To separate "**absent**" from "present in a form my pattern does not
  describe", you need a second, wider view of the same input. A control
  cannot do it.

That third one is the limit of the positive control, and it was measured
rather than assumed. With the filter above and a control line of `# tests 1`,
a typo'd filter reads control 0 and is caught — but Node 24 output, which
*does* carry a summary, and a run with no summary at all **both** read control
1 and data 0. A control certifies the instrument against an input you chose,
so what it certifies is your model of the data; where that model is the thing
that is wrong, the control passes and the zero still misleads. A wider matcher
separates those two, 3 against 0. This arrived twice — once on purpose, and
once when the matcher written to measure the table above scored zero against
real console output that contained what it sought, because the control began
at a line start and the data did not.

Where it does apply, the control is not ceremony. The probe first written to
check that anchor claim returned `false` for every input, including the one
that had to be `true`, because a JavaScript program was interpolated into a
PowerShell double-quoted string: `\` is not the escape character there, a
backtick is, so Node received `^alpha\$` — an escaped literal dollar sign —
and searched for text that occurs nowhere. The failure agreed with the
hypothesis, so nothing in the output invited doubt; that is a broken
instrument, and the control caught it. Two refinements, each of which the bare
rule misses:

- **The instrument must not change between calibration and use.** Calibrating
  "expect 14" with a filter that excludes JSDoc `*` lines and then checking with
  one that excludes only `//` is not a filter matching nothing: the count is
  real and the filter is real, and it is the *pair* that is incoherent. Verify
  with the filter that set the expectation, or re-derive the expectation.
- **Verify a restoration against the bytes you saved — not against a diff, and
  not against a green baseline.** A mutation harness that restores in the same
  shell it mutates in is one killed process away from leaving the tree
  mutated, and a diff read through a filter is this class eating its own
  cleanup check. But a passing baseline is no better: a mutant that survives
  on an untested boundary leaves the suite green, so that check reads the same
  whether the restoration worked or not. Compare the restored files against
  the captured pre-mutation bytes, unfiltered and untruncated, and keep the
  baseline beside it as a *behavioural* check. Capture those bytes yourself —
  `git HEAD` does not know about edits that were already in the tree.

What makes this class expensive is not the wrong answer, it is where the wrong
answer sends you: its output is attributed to the code under test first. A tree
left mutated by a killed run surfaced as a test timing out on correct code, and
the code was distrusted for a while before the tooling was.

Expect the criterion to outrun both of those remedies. The preload probe above
reported a count and would have passed a positive control: it produced a
plausible, non-empty, correct-looking result whose value was identical whether
or not `isTTY` mattered. Neither counting nor controlling reaches that. When
the check you hold cannot differ across the two states you are trying to
separate, what you need is a different check — often a different class of
input altogether — and not a better-tuned version of the one you have.

Finally, hold the examples to the rule. Every instance above arrived with a
plausible cause already attached, and several were stated more confidently than
they had been measured — the anchor claim was true of PowerShell and false of
Node, and the test-output prefix was blamed first on interactive-versus-CI and
then on the Node version alone, when it is both. Each was on its way into this
file, carrying authority and with nothing downstream positioned to check it,
and each survived only because somebody re-measured their own claim. An example
is a claim, and an example inside a rule about unverified measurement is a
claim under the brightest light there is.

## Deliberate decisions that look like defects

Please do not report these as bugs on their own. Do report a place where the
code **fails to follow** one of them.

| Looks wrong | Why it is right |
| --- | --- |
| Numeric `wd*` constants instead of named styles | Word's UI here is German. Both `Range.Style = 'Heading 1'` **and** the OOXML style id `'berschrift1'` throw. Only the numeric constant (`-1 - level`) or assigning another paragraph's `Style` **object** works. Naming a style is the bug. |
| Style ids copied verbatim, never constructed or matched | Word mints ids from the *localized* style name with non-ASCII dropped, so the id is `berschrift1` — not `Heading1`, not `Überschrift1`. Any construction or comparison of style ids is wrong. |
| `edit_document` refuses to batch edits | An address is a **coordinate, not a handle**. Deleting one of a duplicate-text group silently renumbers its successors; renaming a heading moves every address beneath it. A batch API would imply a stability that does not exist. One edit per read is the contract (ADR 0006). |
| Generous timeouts and long polling deadlines | `Documents.Open` on a locked file, or one carrying mark-of-the-web, **hangs indefinitely** rather than failing. Word's process teardown also contends on per-user state, so its tail is load-dependent and is *not* bounded by measurements taken on an idle machine. Deadlines here are deliberately generous; polling makes that free on success. |
| `file_locked` returned for a file that is not open in Word | Word takes a handle with **write access** granting **`FileShare::Read`**, while our readers request *read* access and grant `ReadWrite` — compatible, so `Copy-Item` and Node both read and copy a document open in Word without complaint. Measured, that case *succeeds*. `file_locked` means *held more strictly than Word holds it*, and "the user has it open in Word" is provably never the cause. Do not "correct" this to the intuitive reading, and never advise closing Word in its message. |
| A reader that grants `FileShare::Read` is a bug, not a safer choice | A caller's `FileShare` is what it grants to **others**, so granting only `Read` refuses to let anyone else write — which conflicts with the write handle Word holds. Measured: such a reader gets a sharing violation on a document `Copy-Item` handles fine. Readers of a possibly-open document must grant `ReadWrite`. |
| `file_locked` and `permission_denied` split, while `writable` stays a single flag | Both follow one rule: **split where the platform distinguishes, stay collapsed where it does not.** A read handle cleanly separates the causes (`EBUSY` vs `EPERM`; `IOException` vs `UnauthorizedAccessException`) *and* they need different remediation — a lock may clear on its own and is worth retrying, an ACL will not. `Test-FileWritable` takes a *write* handle, which genuinely folds sharing violation, ACL and read-only into one observation, so it is reported as a fact beside a typed code and no cause is ever inferred from it. This is not inconsistency. |
| Protected View used instead of clearing the zone marker | `Unblock-File` would silently delete a security marker from a user's file. The Protected View path keeps `Zone.Identifier` intact (ADR 0007). |
| Vendored pdf.js worker committed as three split parts | Extension install enforces **1,000,000 bytes per file, 5,000,000 total** (decimal). The worker is 1,262,398 bytes. Repo-folder install never runs a packager, so the split must exist in the committed tree. |
| `"version": 1` in `copilot-extension.json` | That field is the *manifest format* version and is parsed as a `u32`. A semver string there makes the extension uninstallable. The product version lives in `productVersion`. |
| `console.log` avoided outside `src/ui/` | Anything on stdout corrupts the JSON-RPC channel. Under `src/ui/` it is fine — that code runs in the iframe. |
| Integration tests that cannot run in CI | They require an installed, licensed Word. The Office-free unit tests run on `ubuntu-latest`; the integration suites are a local gate by necessity, not by neglect. |

## Where the real bugs in this codebase live

These are the failure modes that have actually bitten us. Scrutiny here is
welcome and has repeatedly found genuine defects.

- **Tearing down Word while work is in flight.** Idle shutdown must check the
  in-flight counter both when *arming* the timer and again when the timer
  **fires** — a timer armed while idle cannot see work that starts afterwards.
  Any new disposal path (canvas close, cache eviction, error handling) must ask
  whether work is outstanding. The one deliberate exception is `shutdown()` on
  SIGINT/SIGTERM, where the process is going away regardless.
- **Character arithmetic on a Word range.** Inside a table, a cell paragraph's
  `Range.Text` ends with `\r` plus `chr(7)` — two characters in the string — but
  `End - Start` counts the end-of-cell mark as **one** position. Trims must be
  derived from position spans, never from string length.
- **Joining COM-side and file-side identity on anything translatable.** The
  application reports localized names while the file stores something else.
  Join on structural keys, never on a display name. The same trap applies to
  errors: discriminate on the **exception type**, never on its message, which
  is localized.
- **An error message that names a cause the code cannot know.** Several bugs
  here have been a correct code carrying a message that asserted the wrong
  reason. If the message names a cause, that cause must be one the code
  actually distinguished.
- **A test that cannot fail for the reason its comment claims.** Node opens
  files with `FILE_SHARE_READ|WRITE|DELETE`, so `readFile`/`writeFile` proves
  nothing about exclusivity; a leak assertion that sleeps a fixed interval
  passes on slack. Assertions must be able to fail.
- **Leaking a Word or PowerPoint process**, or killing one we did not start.
  PowerPoint is **single-instance**: a COM-attached instance *is the user's*,
  so quitting or killing it destroys their work.
- **Unbounded `Documents.Open`**, or any COM call that can hang, without a
  timeout.
- **Caching an address across an edit.**
- **Assuming a read returned the whole document.** Reads are paged and default
  to 300 paragraphs.
- **A path crossing a parser nobody accounted for.** This has been found twice,
  in different parsers, so treat it as a class rather than two sites. A path
  interpolated into a PowerShell single-quoted literal breaks on an apostrophe
  (`C:\Users\O'Brien\…`). A path handed to `cmd.exe` breaks on `&`, `^` or a
  matched `%VAR%` pair — Node quotes an argv element only when it holds a space,
  tab or quote, and `cmd` parses whatever is left: measured, 3 of 9 ordinary
  filenames corrupted, including `%PATH%.docx` expanding into the path.
  Note the two sets barely overlap, so enumerating dangerous characters per
  site is the mistake. **Prefer removing the parser to escaping it**: pass
  values as discrete argv elements (`powershell.exe -File script.ps1 -Param
  value`, `explorer.exe <path>`), never interpolated into a command string.
  Flag any new `-Command` with an interpolated value, any `shell: true`, and any
  `cmd.exe /c`. A `node -e "<program>"` whose program text is assembled by
  interpolation belongs in that list for the *opposite* reason to the entries
  above it. Measured: a regex written `/^alpha\$/m` inside a PowerShell
  double-quoted string reached Node as `^alpha\$`, an escaped literal — and the
  bare `$` arrives intact, because `$/` is not a valid variable start. Nothing
  was corrupted in transit; the parser's presence induced a defence that was
  unnecessary and destructive. The remedy is stronger for that, not weaker:
  removing the parser stops a live one mangling a correct value **and** removes
  the impulse to defend against it. Write the program to a file and pass the
  path as argv.

## Conventions

- Entry point must stay `.github/extensions/<name>/extension.mjs`; no
  `package.json` and no `node_modules` — the SDK is auto-resolved, and the
  folder must run exactly as committed.
- Tests are split into `test/unit/` (Office-free, run in CI) and
  `test/integration/` (`*-smoke.mjs`, need real Office).
- Shared test helpers such as `test/integration/word-pids.mjs` must be
  **imported**, not copied.
- **Quote the claim you rely on; never point at a bare `file.ext:NN`
  coordinate.** A coordinate hides the claim, so when it rots a reader lands on
  unrelated code and files the miss as a cosmetic nit, closing the question over
  whatever error sat underneath (#141, #168; ADR 0009). This has two forms. **In
  issue and pull-request text**, quote the line or the claim it stands for; a
  coordinate may sit *beside* the quote but never replace it. **In a committed
  file**, do not write a coordinate at all — reference code by name and file,
  `Set-ParagraphText (word-host.ps1)`, and note this covers a bare `:NN` into the
  file's own body. Three exceptions: a coordinate pinned to a commit (`… as of
  4abf952`), which cannot rot; one inside a verbatim probe transcript, because a
  recorded run is evidence and evidence is not edited to satisfy a gate; and one
  quoted as the subject under discussion. The gate adds a fourth pass — vendored
  and binary files — which is a limit of its scope, not an exception anyone may
  invoke. Only the in-tree half is gated: nothing
  reads tracker text, so that half stays convention, while
  `check-citation-lines.mjs` now **rejects** a coordinate in any tracked line it
  reads rather than validating it. Read its green message as it states itself —
  for fenced blocks it counts the lines it read past and how many were
  coordinate-shaped; for exempt files it counts the files and the
  coordinate-shaped lines inside them; vendored and extension-skipped binary
  files it never opens, so there it can only count files. A pin is resolved
  against git, so a SHA-shaped string naming no commit is rejected rather than
  waved past.

## If you are coordinating work rather than doing it

A coordination session dispatches work sessions, routes their pull requests
through review, relays findings and decides when things merge. If that is your
role, these are binding, and each exists because ignoring it burned real time.

**Do not review code yourself. Route it to an independent sub-agent.** Reading
a session's diff, forming a verdict on it, and then commissioning a review is
marking your own work and hiring a witness. It also makes you the least
independent reader in the loop while consuming the coordination capacity that
is your actual job. Write the brief — the brief is where your judgement
belongs, and a good one carries what the reviewer could not otherwise know:
which claims are load-bearing, what earlier rounds already settled, and what
must **not** be re-reviewed. Then let the sub-agent do the reading. Corollary:
if you have no verdict, you have nothing to withhold — do not stage
"independent" agreement with a conclusion you already reached.

**Reviews run as independent sub-agents, and a different model each round.**
Rounds have repeatedly overturned each other here; a second round on the same
model mostly agrees with the first. Feed round 1's dismissed "cosmetic" findings
to round 2 explicitly — an overturned Low has twice been the real defect.

**Do not block.** Process messages from child sessions as they arrive rather
than queueing them behind your own work; a coordinator that serialises is a
coordinator that has stopped coordinating. In particular, do not generalise one
session's narrow request into a global freeze. When a session asks for a quiet
machine it is asking about a specific interference — usually **Word**, because
process-leak assertions difference pids across a window. Office-free work does
not perturb that. Dispatch it, with the constraint written into the kickoff
("start no Word, PowerPoint or `test/integration/` suite"), not assumed.

**Your own corrections are claims.** A correction you relay carries authority
and lands where nothing checks it: every downstream gate compares the diff to
the brief, and the brief is the thing that was wrong. Verify a correction before
sending it, and treat a session that adopts your wording *without* checking the
source as exhibiting the defect rather than complying. A session that reads the
probe and refuses your framing is doing the right thing — say so plainly.

**An issue's *remedy* is as checkable as its *diagnosis*, and gets checked far
less.** Diagnoses get scrutiny because they make a claim about the code;
remedies read as intent and slip through. Check the proposed fix against the
same evidence bar as the reported bug **before dispatch** — afterwards the
session inherits the error and ships it under review cover.
