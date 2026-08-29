# Copilot Canvas — Office documents

Agent-callable tools that read, author, and edit Word and PowerPoint documents
by driving the installed Office application through COM, plus a side-panel
canvas that displays the result page-accurately.

The tools are the product. The canvas is how the user watches the agent work.
Excel is deliberately out of scope for now; see ADR 0003.

## Language

### The product

**Document tool**:
An agent-callable tool that reads, authors, or edits an Office document through
the installed application's COM API. Registered in `tools[]`, callable with no
canvas open.
_Avoid_: canvas action (a different, display-scoped mechanism), plugin, command

**Canvas**:
The side-panel surface that displays a rendered document. Display only — it is
never an editing surface, and it holds no authority over the document.
_Avoid_: viewer, panel, editor

### Files and authority

**Original**:
The document file at the path the user named. There is only ever this one file
— the agent edits it in place rather than working on a duplicate.
_Avoid_: source, target, real file

**Operation**:
One agent edit, bounded by open, change, save, close. The unit of locking, of
undo, and of verification. Measured at roughly 230 ms.
_Avoid_: transaction, command

**Snapshot**:
A copy of the original taken before an operation, kept as the revert point.
Undo restores a snapshot; it does not use Word's in-process undo stack, which
does not survive the close.
_Avoid_: backup, working copy

**Lock window**:
The span within an operation during which the file is open and therefore
unwritable by anyone else. Kept as short as possible, because while it is open
no script can regenerate the document and no other Office process can open it.
A second process attempting to open a held file does not fail — it **blocks
indefinitely with no dialog**, measured for both Word and PowerPoint, and needs
an external kill. That is why every open is timeout-bounded and why a lock is
detected by probing for a write handle rather than by attempting an open. The
same silent-hang shape occurs on a file carrying a mark-of-the-web; see ADR 0005
and ADR 0007.

**Hidden instance**:
The Office process the extension drives, `Visible = false`, which no human ever
interacts with. **Word only.** Word is multi-instance, so each COM activation
starts a `WINWORD.EXE` that is genuinely ours to quit and, if necessary, kill.
_Avoid_: background Word, server instance

**Attached instance**:
What a COM activation gives for PowerPoint, which registers one automation
server per session: the object returned *is the user's PowerPoint* if one is
running. It is not hidden and it is not ours. `Quit()` on it closes the user's
decks and killing its pid kills their session, so neither is ever permitted.
Distinguishing this from a hidden instance is a correctness rule, not hygiene.
See ADR 0003.

**Isolated instance**:
A genuinely private Office process obtained by launching on a separate desktop
and binding through `OBJID_NATIVEOM` — the only way to get one for PowerPoint.
Costs roughly 12 s to reach a bound instance against 1 s for a COM attach, which
is why it cannot be opened lazily on first interaction.

**User instance**:
The user's own visible Office application, opened by "Edit in Word". Outside
our control, and a second writer we must detect rather than collide with.
Launched over plain argv (`explorer.exe <path>`), never through `cmd.exe`,
because a shell parse corrupts ordinary filenames — measured, 3 of 9, including
one that launches a *different* path silently.

**Handoff**:
Transfer of write authority between the agent and the user instance. Exactly
one of them holds it at any moment.

### Failure

Failures carry a typed code, and there are **two surfaces** that emit them. A
layer branching on a code has to know which one it is talking to, because the
two sets barely overlap.

These lists are the codes **worth branching on**, not an inventory. Hand-kept
inventories drift — this section has already been wrong twice that way, and no
single grep finds every site (some codes sit on a continuation line of a
multi-line call). The source is authoritative: extension codes are the first
argument to `new CanvasError(...)` or `new DocumentError(...)`, host codes are
the second argument to `Send-Fail` in `word-host.ps1`.

**Extension surface** — raised by a tool, a canvas action or the viewer's HTTP
routes, mostly before any Word process is involved: `invalid_path`,
`unsupported_type`, `file_not_found`, `not_open`, `invalid_query`,
`invalid_json`, `payload_too_large`, `canvas_not_open`, `no_document`,
`page_out_of_range`, `internal`, `not_found`, and `word_canvas_error` as the
fallback when a canvas action's Word call fails with nothing more specific.

**Word host surface** — the PowerShell host over the JSON protocol. Today it
emits **two** codes: `word_error` for any COM failure and `unknown_command` for
a `cmd` it does not recognise, the latter being a protocol bug rather than
anything a layer handles. Separately, `word_unavailable` and `word_timeout` are
minted **Node-side** in `word-host.mjs` when the host dies or overruns, so they
reach a caller without the host ever being asked.

The typed set arrives with `read_document` and is **not on `main` yet**:
`file_locked`, `permission_denied`, `document_unreadable`, `copy_failed`,
`write_failed`, `no_such_document`, `invalid_request`,
`document_changed_during_read` — alongside `file_not_found`, `word_unavailable`,
`word_timeout` and `page_out_of_range`, which already exist, and `word_error`,
which stays as the fallback rather than being replaced.

`file_not_found` is the only code common to both surfaces today, and it means the
same thing in each. Everything else is surface-specific.

Most of the host's codes mean what they say. Two do not — `file_locked` and
`permission_denied` — and a third trap is not a code at all: `writable` is a
field reported *beside* a code, which is why it is described here with them.
Every layer must get all three right, because the code is what an agent branches
on and the message is what a user acts on.

**`file_locked`**:
The original is held **more strictly than Word itself holds it**. Measured: Word
takes a handle with **write access** granting **`FileShare::Read`**, while
`Copy-Item` and Node both request *read* access and grant `ReadWrite` — a
compatible pair, so either opens a document the user has open in Word without
complaint. That case *succeeds*.
"The user has it open in Word" is therefore provably **never** the cause of a
read failure, and a message saying so names the one condition incapable of
producing it. Worth retrying, because the holder may let go.

The mechanism matters, not just the outcome: a caller's `FileShare` value is
what it grants to *others*, so a reader asking for `FileShare::Read` is refusing
to let anyone else write, which conflicts with the write access Word already
holds. Measured against a real open document, such a reader gets a sharing
violation on a file `Copy-Item` copies fine. **A reader must itself grant
`ReadWrite`.** `spikes/isolation/probes/probe-fileshare-algebra.ps1` runs five
readers against four holders and against real Word, and is the whole evidence
trail: its `read, grants Read` row separates the candidate *access* mechanisms,
and its `write, grants ReadWrite` row is what finally measured the *share* half.
Every reader that asks only for read access is blind to the share mode — which
is how this claim was revised twice with nothing going red.
_Avoid_: in use, open elsewhere, busy

**`permission_denied`**:
The filesystem refused, by ACL. A separate code from `file_locked` not merely
because the platform distinguishes them — `EBUSY`/`IOException` against
`EPERM`/`UnauthorizedAccessException`, discriminated on exception **type**, never
message, because the machine is German — but because the remediation differs: a
lock may clear on its own, a denied ACL will not. Never retry it.

**`writable`**:
A fact reported beside a code, never a cause inferred from one. It comes from
taking a *write* handle, which cannot separate a sharing violation from a
denying ACL from a read-only attribute, so it stays a single collapsed flag.
The rule both halves share: **split where the platform distinguishes, stay
collapsed where it does not.**

Two non-failures worth naming, because both sound like failures and neither is.
On Windows/NTFS, which is the only platform the host runs on: the **read-only
attribute** does not block reading, and neither does `chmod`, because Node's
`chmod` on Windows does nothing but toggle that same attribute. (On POSIX
`chmod` genuinely does remove read permission — the claim here is a Windows one,
and it is worth stating that explicitly because this repo's unit tests run on
Linux.) No layer needs a branch for either.

### Rendering

**Render**:
Producing page images through the application's fixed-format export — the same
pipeline the app uses to print. Feeds both the canvas and the agent's own
verification.

**Page-accurate**:
Laid out by the Office application itself, so pagination, fonts, and spacing
match what that application would print. The property that ruled out HTML
conversion.
_Avoid_: WYSIWYG, faithful, high-fidelity

**Verification**:
The agent re-rendering after an edit and receiving the page back as an image,
so it can see what it changed rather than assume.

**Auto-refresh**:
The canvas reloading when the file changes on disk, including when a script
replaces it.
_Avoid_: live (it meant pixel streaming during the spike, and now means this)

### Addressing

**Structure map**:
What a read returns — the document's paragraphs with their text, style and
heading path, each carrying an ID minted for that read. Obtained in one
`WordOpenXML` call and parsed outside Word, never by walking paragraphs.
_Avoid_: outline, index, model

**Address**:
An ID from a structure map, identifying the paragraph an edit applies to.
Derived from heading path, text and occurrence index, because Word exposes no
stable paragraph identity of its own. **A coordinate, not a handle**: valid
within one read-then-edit cycle and never cached across an edit. Deleting one of
several identically-worded paragraphs renumbers its successors, so a stale
address stays *valid* while pointing at different content; renaming a heading
moves every address beneath it.
_Avoid_: selector, locator, range

**Translated identity**:
Any name that Office localizes, which therefore differs between what COM reports
and what the file stores. PowerPoint reports a shape as `Title 1` while the
`.pptx` stores `Titel 1`; Word mints style ids from the localized name with
non-ASCII dropped, so Heading 1 is `berschrift1` on disk. Two applications, two
mechanisms, one rule: **never join COM-side and file-side identity on anything a
human could have translated.** Join on structural keys — `p:ph/@type` for
placeholders — and carry style ids verbatim rather than constructing or matching
them.
_Avoid_: display name, style name (as identifiers)

**Revision token**:
A hash of the file, returned with a structure map and required by any edit. A
mismatch means someone else changed the document, and the edit is refused
rather than applied to the wrong place.
_Avoid_: version, etag, checksum

**Intent**:
What the agent submits to change or author a document — headings, paragraphs,
tables and lists — which the extension translates into COM calls. Deliberately
not OOXML, and deliberately not prose.
_Avoid_: command, patch, script

### Review

Every PR gets a Copilot review. Reply to every comment, push, re-request,
repeat. The coordinator requests and merges; the owning session replies and
pushes, so two worktrees never touch one branch.

**Read the review body, not just the thread list.** Findings arrive in two
shapes and only one of them is a thread. **Suppressed** findings — "previously
missed, in code that hasn't changed" — appear in the body alone: no inline
thread, no resolve button, no reply affordance, and setting `in_reply_to` to
one of their comment IDs fails with a 422. So they sit outside "address every
comment, reply, repeat" entirely, and a PR can show zero unresolved threads
while carrying the only finding that gates its merge. Reply to those with a
top-level PR comment. Three of the best findings on #12 arrived this way,
including the one that turned out to have three sites rather than one.

**Do not filter review content by author login.** The reviewer renders as a
different string on each endpoint that reports it:

| endpoint | login |
| --- | --- |
| REST `/pulls/{n}/reviews` | `copilot-pull-request-reviewer[bot]` |
| REST `/pulls/{n}/comments` | **`Copilot`** |
| GraphQL `reviewThreads.nodes.comments.nodes.author.login` | `copilot-pull-request-reviewer` |

A filter correct for two of the three returns **empty** on the third, and an
empty result is indistinguishable from a clean round — so the mistake reports
*good news*. It has already produced two wrong conclusions here: a review round
believed missing on #16, and later that same PR's round four believed to have no
inline findings when it had two. Select comments by `created_at` against the
review's `submitted_at` instead, and cross-check the body's *comments generated*
count against the number retrieved.

**A review record can be an empty stub, and an empty stub is indistinguishable
from a clean round.** Requesting a review sometimes creates a stub within
seconds, with the substantive review arriving minutes later against the same
commit. Measured twice on #42:

| record | submitted | commit | body |
| --- | --- | --- | --- |
| `5057359653` | 08:14:49Z | `b6ebdb5` | **empty** |
| `5057365906` | 08:17:58Z | `b6ebdb5` | the finding |
| `5057371102` | 08:21:27Z | `bac53f3` | **empty** |
| `5057375339` | 08:24:08Z | `bac53f3` | the finding |

**"Sometimes" is doing real work in that sentence, and it is a correction to the
first version of this rule**, which said a request *produces* two records. The
very next request — round 1 of the PR that added this paragraph — produced **one**
record, substantive, 82 s after the request, with no stub at all. Two of two on
one PR and zero of one on the next: what makes the difference is unknown, so the
timing is not predictable and must not be planned around.

What survives is the part that does not depend on the stub appearing. Inside that
window both obvious questions return a wrong answer: reading the *latest* review
yields a record with no findings, and `requested_reviewers` has **already
emptied**, so "is a review pending?" reads as *no*. Counting records to derive
the round number also double-counts whenever a stub does appear, which walks a PR
into the six-round cap early.

Discriminate on the body, not on recency, not on the reviewer list, and not on
the record count: every review actually performed here opens with a state
headline (🟢/🟡) and closes with a *comments generated* count. An empty body
means the review has not happened yet.

**Six rounds, then merge.** This is a hard cap, not a target. At round six the
PR merges with any remaining comments declined explicitly in a reply. The one
thing that may go past six is a **correctness or data-loss defect**, and it buys
exactly one more round — it does not reset the counter.

**Effort follows the same curve: Balanced for rounds 1–4, Lite for 5–6.** The
deep pass is worth paying for while the diff is still novel; by round five the
reviewer is re-reading code it has already cleared, and the findings have
consistently been documentation and contract precision rather than defects.
Escalate back to Balanced mid-loop only when a round lands a **substantive code
change** — new logic, a changed contract, a fix touching more than the site it
was aimed at — because that is new material the deep pass has never seen.
A one-line comment or message fix is not that, and re-reviewing it deeply has
never once paid.

**Effort cannot be set programmatically**, so the curve's upper half is
aspirational and every round is Lite in practice. It is a control on the PR's
Reviewers panel and nothing else: measured across five routes, including a REST
field that is accepted and then silently ignored. Because a request can look
like it succeeded, **read the effort level printed in every review body** rather
than assuming the one you asked for took effect.

The cap exists because the loop has no natural end. The reviewer re-scans
unchanged code every round and surfaces different **suppressed** findings each
time, so "0 new comments" never means "clear" — on #12, rounds four and five
both reported zero new comments and both carried real suppressed findings. A
reviewer that keeps finding things is not evidence the file is getting better.

What the rounds were actually worth, measured on the first two PRs to run the
full loop:

| rounds | what they produced |
| --- | --- |
| 1–3 | the only findings that changed shipping code — including the path-quoting comment that exposed `cmd.exe /c start` corrupting 3 of 9 ordinary filenames |
| 4–6 | documentation and contract precision: stale comments, miscounts, a schema bound the runtime did not enforce |
| 6+ | not observed to produce anything, on a diff of +3860/−107 across 29 files |

Declining is a real outcome and must be recorded as one: reply saying why, so
the next engineer sees a decision rather than an omission. Replies are for the
human record — Copilot cannot read them.

**Lite is a floor.** A clean Lite pass means the cheap checks passed, not that
the code is right — an independent deeper read of #16 found a critical
data-loss defect that six Lite rounds elsewhere never approached. Spending
rounds 5–6 at Lite is only defensible because the deep pass ran at 1–4; it is
not a claim that Lite suffices.

**Mutation-check every test before you rely on it: reintroduce the defect, run,
confirm red.** A test that cannot fail is worse than no test, because it reads
as evidence. This is not a caution, it is a step — two of #12's four
test-related findings were tests that could never fail, and *both were written
by someone who had just been told about that exact failure mode*, one of them
in the same session in which they had criticised another test for it. The
tautology is easy to write and invisible on inspection: asserting that every
entry in `SUPPORTED` appears in a list built from `SUPPORTED` passes with the
defect fully reintroduced. Knowing about the failure mode demonstrably does not
prevent it; running the check does.

The same rule catches the case where the machine, not the test, is providing the
coverage. A green suite says nothing when **the environment cannot produce the
input**: this account generates no apostrophe in a path, the Linux runner
generates no exclusive lock, an idle machine generates no contended shutdown.
That code is untested, not correct, and the suite disguises it. The tell is when
a test's coverage depends on a property of the machine rather than on the test.

**Assert through the boundary the caller actually sees, not one layer beneath
it.** A test placed below a boundary will confirm a property no caller can
observe. `edit_document`'s recovery path set `snapshot`, `rolledBack` and
`documentUnchanged` on its error, and its unit tests asserted all three and
passed — while `asToolError` forwarded only `code`, `message` and `data`, so
every one of those facts was dropped at the tool boundary and the agent saw a
bare `word_timeout` with nothing about the state of the user's document. The
requirement was to keep the bytes *and* tell the caller they exist; the tests
could only see the first half.

An independent deep review missed it too. It traced all three branches of the
recovery function and correctly cleared it — no path discards the pre-edit
bytes — because **a reviewer reasoning about a function does not reason about
the boundary that function's output has to cross.** That blind spot is
structural, not a matter of effort level, so it is not something a better review
fixes. Test at the surface a caller consumes.

**Neither side of a boundary can be inferred from the other, and a fix on one
side makes the other look covered.** The Word host set `[Console]::OutputEncoding`
to UTF-8 and never set `InputEncoding`. Reads were therefore perfect — a German
document came back with every umlaut intact — while every non-ASCII character
the agent *wrote* was decoded as the OEM codepage and saved to disk as mojibake:
`Grüße` became `Gr├╝├ƒe` in `word/document.xml`, with `U+00FC` absent. The
existing line was not an oversight; it was **evidence that someone had
recognised the boundary needed an encoding**, and that is precisely what stopped
anyone looking at its other half.

Nothing went red, and the reason is worth stating exactly, because the obvious
explanation is wrong. It is *not* that the write direction is untested — edits
are written and read back end to end. It is that **every fixture in the
repository is ASCII**, including the line labelled `Umlaut check:`, which
contains `Grosse Aepfel, Strasse, Muenchen` and not one umlaut. A test named for
the property it does not exercise is worse than no test, because the next person
greps for the property, finds it, and stops. One non-ASCII round-trip assertion
goes red on its first run.

The same shape holds across surfaces, not just directions. A declared
`minimum`/`maximum` is **not enforced for tools** and **is enforced for canvas
actions**. Both were measured first-hand, because neither predicts the other.

The tool half was measured with a purpose-built probe tool that **does not ship
here**: it declared an integer bound, and every violation reached its handler
unchanged, including a string where an integer was declared. That is provenance,
not something to look up — no tool schema in this repo declares the bound it was
measured against, so do not go hunting for one. The canvas half is checkable in
place: `get_outline` declares `limit` as `1`–`2000` in `extension.mjs`, and both
`0` and `99999` are refused before the handler runs, naming the field, the value
and the violated bound.

The consequence is asymmetric, and that is the part that changes code. For
**tools** the runtime must validate its own arguments, so `normalizeReadArgs`
rejecting `limit` below `1` or above `MAX_READ_LIMIT` is load-bearing rather
than belt-and-braces. For **canvas actions** a handler-side check of a declared
bound is dead code. Carrying one surface's result to the other already turned a
real finding into a wrong one here: the two `word-host.ps1` handlers that take
`limit` — `Cmd-Outline` and `Cmd-Search` — apply a default and then coerce with
`[int]`, performing no bounds check whatever, and were reported as further sites
of the declared-but-unenforced-bound defect. They are not. The bound is enforced
above them, so the absent check is correct rather than missing. Measure each
direction and each surface separately, or say you have not.

**Name the discriminating case before you trust the probe.** `spikes/isolation/PLAN.md`
§19 states the rule and the failure that produced it: a probe on which every case
agrees has measured nothing, and the fix is not more care reading the output but
identifying up front the one case whose result differs between the candidate
mechanisms. Four readers of a Word-held file all requested `ReadWrite`, so all
four succeeded under either candidate holder, and the lock model was revised
twice with nothing ever going red.

This generalises past probes to ordinary fixtures. The layout-mark normalizer
covers four marks and only **two of them discriminate**: `w:tab` passes by luck
because both sides happen to emit `\t`, and `w:softHyphen` passes because both
sides drop it — both would pass against a completely broken mapping. `w:br` and
`w:noBreakHyphen` carry the entire measurement. Which cases discriminate is not
recoverable from a green run, so record it in the fixture where the next reader
will find it.

**A claim about state is a measurement, and it carries the instant it was
taken.** Cross-session messages queue and have arrived here **~11 hours** late,
and the PR-state notification attached to one is frozen at send time along with
it. A correction arrived asserting that #22 was open with `main` at #21; #22 had
merged ten hours earlier. It was accurate when written. The claim it corrected
had been premature when made and had become true by the time it was read — both
parties right about different instants, which no amount of care about *facts*
prevents. Re-read state immediately before acting on it, and never carry a
status out of a message and into a mutation.

The same applies to git. "Is my work merged?" cannot be answered by
`git log main..branch`, which cannot distinguish squashed-and-merged from new,
**nor** by `git diff main branch` once the branch is behind, because `main`'s
newer commits then read as changes the branch removed. Both misled here in the
same direction, within one session: two commits looked like stranded work worth
rescuing, and a reset would have looked like destroying it. Rebasing produced a
conflict whose resolution showed `main` already held the same rule in better
wording — the commits were redundant, not stranded. Only searching `main` for
the content itself distinguished the two, and the two look identical from every
count and every diff.

**A mutation gate can report on a mutant it never applied.** Three independent
mechanisms turned up here, one per layer session, which is what makes this a
class rather than three bugs. In all three the run completes and prints a verdict,
so nothing about the result says the instrument misfired. **None of this tooling
is on `main` yet** — the runners arrive with #26 and #43 — so the file names below
are provenance, not paths a reader of `main` can open.

- **The anchor moved.** The mutator patches by matching a snippet of source. That
  snippet can match zero times after a refactor, or more than once, and a run that
  patched nothing — or patched a different line — still produces a verdict for the
  line you named. Two outcomes cannot express this, so use four: `KILLED`,
  `SURVIVED`, `MISSING` when the anchor matched **zero** times, and `AMBIGUOUS`
  when it matched **more than once**. Abort on any match count you did not predict
  in advance, not only on a count of zero — a mutant landing on the second of two
  matching sites can fail an unrelated test and be recorded as `KILLED`, which is
  the one direction of this defect that reports good news.
- **The runner never started.** `mutate-create.ps1` (#26) resolves paths relative
  to the extension root and exits 1 when invoked from the repo root. It fails
  *loudly* and is still dangerous, because "gate ran, gate red" and "gate never
  ran" are the same exit code. This is the `MISSING` category one level up, aimed
  at the runner instead of a mutant, and only reading the output caught it.
- **The mutation was inert.** A guard asserting that `.gitattributes` marks the
  vendored pdf.js parts `-text` (#43) was mutated by moving that file aside. Once
  `.gitattributes` was itself committed, git resolves attributes **from the index**
  when the working-tree file is absent, so the mutation stopped disabling anything
  and both mutants went from `KILLED` to `SURVIVED` with no change to the test.
  Generalises past this one file: **mutating a tracked config file by moving or
  editing the working copy is inert the moment that file is committed.**

The remedy common to all three is to treat the mutant itself as the thing under
test: assert that the mutation was applied, and that the *specific* test you
expect to go red is the one that did. A gate that only counts failures cannot
tell a killed mutant from an unrelated breakage.

**An instrument that measures correctly and reports illegibly is still a wrong
answer.** The autocorrect probe on #26 (not on `main`, so it is named here as
provenance only) runs each arm in a child
process and captures stdout to a file. The child writes **CP850** and the parent
reads it back as **Windows-1252**, so `(c) → ©` printed as a cedilla and the
curly quotes best-fitted back to `"`. The arm that exists to catch Word's
substitutions was printing lines marked `[REWRITTEN]` whose *asked* and *got*
were character-for-character identical on screen — the substitution rendered as a
no-op. The comparisons themselves are `-ceq` against the live COM string and
never touch the console, so nothing measured was ever wrong; only the report was,
and two comments in the tree had been written from that report.

Two things follow. **Treat "a PowerShell child-process boundary silently
transcodes" as the class**, of which #40's stdin defect is one instance and this
one — output rather than input, probe rather than host — is another; they were
found independently and neither instrument would have caught the other. And
**prose about a measurement drifts from the measurement, because nothing checks
prose**: the same PR found a comment promising a bait character that its own
probe had never rewritten, and both findings on #42 were the wording asserting
behaviour the code did not have. Print evidence escaped (`<U+XXXX>`) rather than
rendered, and re-run rather than recall when writing about what a probe found.

