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
  `cmd.exe /c`.

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
  file's own body. Two exceptions: a coordinate pinned to a commit (`… as of
  4abf952`), which cannot rot; one inside a verbatim probe transcript, because a
  recorded run is evidence and evidence is not edited to satisfy a gate; and one
  quoted as the subject under discussion. Only the in-tree half is gated: nothing
  reads tracker text, so that half stays convention, while
  `check-citation-lines.mjs` now **rejects** a coordinate in any tracked line it
  reads rather than validating it. Read its green message as it states itself —
  it counts out the fenced, exempt and vendored lines it never read.

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
