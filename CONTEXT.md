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

**How a code reaches the agent at all — the two channels differ, measured.** A
tool failure must be **returned**, not thrown: a thrown error crosses the host as
the bare string `Tool execution failed`, with the code, the message and the
`data` bag all discarded (#45, `spikes/tool-errors/`). `src/tool-error.mjs`'s
`toolFailure` renders all three into `textResultForLlm`, which is the one thing
that survives, and `extension.mjs` applies it at the tool **registration site**
so no handler can opt out. A **canvas action** is the opposite: its thrown
message survives but its `code` field does not, so the code is folded into the
message there instead. Anything a layer records outside those two paths is
visible to a test and to nobody else.

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
`ReadWrite`.** Two probes carry the evidence, and they were built in that order
for a reason: `spikes/isolation/probes/probe-fileshare-algebra.ps1` runs five
readers against four holders and against real Word — its `read, grants Read`
row measures the holder's *access*, and its `write, grants ReadWrite` row, the
last to be added, is what finally measured the *share* half.
`spikes/isolation/probes/probe-share-vs-access.ps1` reaches the same answer from
three holders that differ by one property each, so the two agree by different
constructions rather than by sharing a mistake.
Every reader that asks only for read access is blind to the share mode — which
is how this claim was revised twice with nothing going red. Re-run against real
Word: its row is byte-identical to the synthetic `write access, grants Read`
holder and differs from `write access, grants ReadWrite` in exactly the share
column, which is the cell that took three attempts to put on the table.
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

### Lifecycle

**The slot**:
The extension holds at most one `RenderCache`, and therefore at most one hidden
Word, in a single slot (`src/render-cache-slot.mjs`). Callers ask the slot; it
builds lazily and hands back what it has. Nobody else may hold a cache.

**The teardown window**:
A disposal is **not instantaneous**. `WordHost.dispose` sends `quit` under a 20 s
timeout and then waits up to 5 s more for the child to exit — a ~25 s ceiling,
derived from those constants rather than measured. `#disposed` is set *between*
the two, right after the quit returns; it cannot be set earlier, because `#send`
refuses to run against a disposed host and would reject the very quit being sent.
So for the tail of that window the host answers every request with `The Word host
has been shut down.` Before that — while the quit is still in flight — a caller
is not answered at all: its command goes to a host sitting inside `Stop-Word`,
and is rejected only when the child exits, with the same `word_unavailable` code
but a different message, `The Word host exited (code N, signal S).` It cannot
respawn Word, because `dispose()` clears `#openArgs` on entry and that makes
`request()`'s replay path unreachable. **One window, two failures** — do not
treat the shutdown sentence as the whole of it.

Two rules follow, and #61 was both of them being broken at once:

- **Clear the slot synchronously, before awaiting the teardown.** Emptying it
  afterwards leaves the doomed cache reachable for the whole window, so a caller
  arriving in the tail is handed a corpse instead of a fresh host. Await the
  teardown by all means — just not before the swap.
- **Resolve the cache per use; never capture one.** This is what turned a ~25 s
  window into a process-lifetime failure: `ViewerInstance` kept its constructor
  reference, and `open` reuses an existing instance for the same id by design
  ("rehydrate, reload, focus"), so a panel created inside the window never
  consulted the slot again. Measured: the surface recovered only on
  `extensions_reload`, which clears it because the instance map dies with the
  process.

Fixing the slot alone would also cure the stickiness, since a panel could then
only ever capture a live cache. Resolving per use is what makes that property
**structural** instead of incidental, so a disposal path added later cannot
strand a panel by forgetting the first rule.

The idle timer is a standing suspect here and was **not** the cause of #61: it is
gated on `isDisplaying()` and canvases were open. The disposal that matters is
the one the last panel's `onClose` runs.

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

Every PR gets a code review, to a hard cap of **two rounds** (below): reply to
every finding, push, re-review, repeat. Prefer GitHub's Copilot reviewer; when
it cannot run, run a `code-review` sub-agent locally under the same rules rather
than skipping the review. The coordinator requests or runs it and merges; the
owning session replies and pushes, so two worktrees never touch one branch.

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

**An empty-bodied review record is somebody's *reply*, not a pending review.**
`/pulls/{n}/reviews` mixes actual reviews with one record per reply posted to a
review thread. The reply-records carry an empty body and are authored by whoever
replied — which for this repo's loop is *us*, once per round:

| PR | records | authored by reviewer | authored by us, empty |
| --- | --- | --- | --- |
| #42 | 5 | 3 (bodies 1889, 473, 384) | 2 |
| #47 | 4 | 3 (bodies 2095, 370, **120**) | 1 |

The 120-byte record on #47 is the billing phantom, and it is counted here as
reviewer-authored because that is what it is — a phantom is produced by the
reviewer, not by a reply. An earlier version of this row read `2 … | 2`, which
double-counted a reply and hid the phantom; the row is re-derived from
`/pulls/47/reviews` rather than reasoned about.

**A correction is a claim, and inherits the burden of one.** A reviewer found
the inconsistency above and proposed the row should read `5 | 3 | 2`. The
finding was right and the proposed fix was wrong: querying the endpoint returns
**four** records — `2095`, `370` and `120` from the Bot, one empty from a User —
so the truth is `4 | 3 | 1`. Adopting the correction unmeasured would have
replaced one wrong row with another and retired the finding that could have
caught it. **Measure the fix, not just the defect** — and note this happened
three times in one afternoon, each time with a true conclusion resting on a
mechanism nobody had checked: a codepage named but never read (`chcp` said 850,
not 437), a suppression flag reporting that assignments did not throw, and this
file's own justification for `assertNoLeakedWord` citing attribution when the
work is done by a timeout. **The conclusion surviving is what stops anyone
looking at the mechanism.**

Three consequences. **Counting records to derive the round number over-counts by
exactly the number of replies you have posted** — deterministic, not a race, and
it walks a PR into the round cap early. **Reading "the latest review" can
return your own reply**: empty body, no findings, indistinguishable from a clean
round. And `requested_reviewers` empties as soon as the review is delivered, so
it cannot answer "has this round come back?" either.

Select on author *and* body: a review is authored by the reviewer, opens with a
state headline (🟢/🟡) and closes with a *comments generated* count.

**There are three producers of review records, not two, and the cheap first cut
is `user.type`.** Measured across #36, #43 and #47:

| producer | author | `user.type` | body | `Review details` | owns an `in_reply_to` comment |
| --- | --- | --- | --- | --- | --- |
| genuine round | the reviewer | **Bot** | 370–7135 | **yes** | no |
| billing phantom | the reviewer | **Bot** | exactly 120 | no | no |
| reply wrapper | whoever replied | **User** | empty | no | **yes, exactly one each** |

`Bot` removes every reply wrapper before the body is inspected at all; the
`Review details` block then splits a genuine round from a phantom among what
remains. The block is a sound **negative** test — no record lacking it is a round
— and its discriminating case passes: #47's round-2 approval is only 370 bytes,
the shortest genuine review observed, and still carries the block. But it cannot
say what a non-round *is*, and the two kinds of non-round mean opposite things:
one is the reviewer failing to run, the other is you talking.

The reply wrapper is identified structurally rather than by shape or timing —
each such record owns exactly one comment, and that comment has `in_reply_to`
set. Shape alone would have called #43's two pre-outage wrappers phantoms, which
is the sufficient-cause trap one paragraph up.

**Operationally: a counter derived from record counts is already wrong, in the
direction that costs rounds.** #36 has 7 records and 2 rounds, #47 has 4 and 2,
#43 has 3 and 1. Five of #36's seven are not rounds and come from two different
producers — and four of those five were produced by *replying to the review*,
which is exactly the action a round exists to invite. The counter would advance
every time someone answers.

One drift worth naming, since it is this file's subject: the check-run annotation
says *"recent **account** payments have failed"* while the review-record body
says *"recent **GitHub Actions** payments have failed"*. Different strings from
different surfaces, describing one condition. **Match on record shape, never on
the sentence.**

**And that closing count is load-bearing, because a review that never ran is
also recorded as a review.** Measured on #36 and #47 simultaneously: both carried
a third record authored by `copilot-pull-request-reviewer[bot]`, `state:
COMMENTED`, at the correct head commit, with zero inline comments and this body
in full:

> The job was not started because recent GitHub Actions payments have failed or
> your spending limit needs to be increased.

Every discriminator above passes it. Right author, non-empty body, right commit,
arrived promptly after the request. It is not a reply-record and it is not a
stale round. **It is a billing failure wearing the shape of a clean review**, and
counting it would consume one of the two rounds and could merge a PR on the
strength of a review that never executed.

What separates it is the one thing the reviewer produces and our query cannot:
the `Review details` block — *files reviewed*, *comments generated*, *review
effort level*. A genuine review always closes with it; this record has no such
block at all. So the rule is sharper than "cross-check the count": **the
count's absence is the tell, not its value.** A round with no `comments
generated` figure did not happen, whatever else the record says.

This is the general rule of this section arriving at its own subject matter. A
query that reports good news must be verified against a number the query did not
produce — and the reason that works is precisely that the failure modes cannot
forge it. Zero findings is what a healthy PR, a reply-record, and an unbilled
job all look like.

**Re-running discriminates flaky from deterministic; it does not discriminate
environment from tree.** Diagnosing the same outage from the other end, #36 saw
`validate` go red on a comment-only change, re-ran the identical commit, saw it
fail again, and concluded the tree had caused it. The reasoning applied this
repo's own rule — *a static diff cannot alternate* — outside its range. A
billing block reproduces perfectly on an unchanged tree, so non-alternation
narrows the cause to {tree fault, deterministic environment} and does not select
between them. What settled it was a **control** (the last green run reported nine
steps, so "zero steps" was a signal and not an artefact of the query) and then
**asking the system what happened** — the check-run annotations endpoint states
the cause in one sentence, for less than either re-run cost. **Read the
annotation before inferring from the outcome.**

The completing half is how to establish *scope*, because a single blocked run is
equally consistent with "my branch is broken". #43 built the discriminating table
rather than trusting one data point: the twelve most recent runs repo-wide,
**5/5 blocked at ≥09:02 and 7/7 clean at ≤08:57, across three branches and two
workflows**, with the annotation confirmed on each blocked run individually
rather than inferred from its conclusion. Varying branch and workflow is what
turns "my PR is red" into "the account is blocked" — the same design rule as
everywhere else here, since a set of runs that all agree has measured nothing.
**And note what makes the annotation unavoidable: a billing block and a genuine
test failure are both `conclusion: failure`.** The field everyone reads cannot
separate them.

**And the same run list will report that the outage has changed shape, if you
let an absence speak.** Hours into it I re-checked, saw nothing newer than
11:06, and reported that runs had stopped being *created* at all — a worsening,
"a change from failing in 2 s". Nothing had changed. Every run this repo
produces is `event=pull_request` or a push to `main`; the workflow has no
schedule and no other trigger, so nothing but a push can create a row. The gap
was the gap between my own pushes, and each run had in fact been created within
five seconds of its commit (`13:12:08+02:00` → `11:12:13Z`). The shape was
constant throughout: row created promptly, job never started, `steps: 0`,
cause on the check-run annotation.

That is this section's own failure in a quieter field. I modelled a mechanism —
an escalating block — from `created_at` alone, while `event` sat unread in the
same objects and the push history sat in `git log`. **An absence in a list is
evidence about the trigger before it is evidence about the system**, and here I
owned the trigger. One field settles it: *what would have had to happen for a
row to appear?*

**Then I transported that finding across six pull requests without checking it
on any of them.** Having established the phantom record on #36 and #47, I
broadcast *"round N is not consumed"* to every session in the stack. It is true
only where a review was requested **after** the outage began, which was those two
PRs and no others. Reconciled by counting genuine reviews per PR:

| PR | genuine rounds | phantom |
| --- | --- | --- |
| #26 | **4** | none |
| #34 | 2 | none |
| #36 | 2 | 1 |
| #43 | **1, genuine and worked** | none |
| #46 | **0 — never reviewed at all** | none |
| #47 | 2 | 1 |

Two of those were live errors, not pedantry: #26 was two rounds further into the
cap in force at the time than anyone believed, and #43's round 1 was real and
had already been answered while I was telling its session the round had not run.
**A per-artefact observation is not a property of the system**, and the
correction
arrived within the hour of writing that a summary is what discards a scope — the
recommendation being itself a summary, which is the shape rather than the
exception. What caught it was a layer session **applying the discriminator to its
own PR instead of accepting the claim**, which is the only reason any of this
gets caught at all.

The failure mode this outage produces is the one worth remembering: a job that
never started is, from outside, indistinguishable from a job still running.
Waiting is the natural response and it never terminates, with every dependent
session parked behind it.

Note this pulls the opposite way from the rule above it, and both are true. Do
not use login as a **filter** you trust to be complete, because the reviewer
renders as three different strings; do read login as a **discriminator**, because
a record's author is what tells you which kind of record it is.

**This paragraph was wrong twice before it was right, both times for the same
reason.** It first claimed a review request *produces* two records, the second
empty; corrected to "sometimes produces" when the next request produced one; and
was only correct once someone read `user.login`, which had been sitting in the
same JSON object throughout. Two rounds of theorising about GitHub's timing, from
`submitted_at` and body length, when the field naming the cause was already in
hand. **Before modelling a mechanism, spend the remaining fields.**

**Two rounds, then merge.** This is a hard cap, not a target. At round two the
PR merges with any remaining comments declined explicitly in a reply. The one
thing that may go past two is a **correctness or data-loss defect**, and it buys
exactly one more round — it does not reset the counter.

**And if GitHub's reviewer cannot run, run the review locally instead of
skipping it.** The billing outage proved the delegation is the fragile part, not
the review: a `code-review` sub-agent reads the same diff from the same tree
with no external dependency and no per-run cost. **Same rules, same cap** — two
rounds, read every finding, decide, reply, push, repeat, decline explicitly in
writing. The only thing that changes is who produced the findings, and that goes
in the squash body so the record says which reviewer ran.

The division of labour is unchanged and is what keeps two worktrees off one
branch: **the coordinator runs the review, the owning session addresses it and
pushes.** A local review is read-only, so the coordinator may fetch and diff
another session's branch without owning it — reviewing is not mutating. Never
let a session be the sole reviewer of its own diff; self-review re-reads the
author's intent, which is the blind spot the loop exists to cover.

### Write the squash message; never let the merge capture the PR body

**A PR body is current, a commit message is historical, and only one of those is
the right home for a claim meant to outlive the conversation.** A body is mutable
right up to the instant of the squash, and the squash can capture it at any
moment without warning — so its content at merge time is a race, not a decision.

Measured, on #26, by the session that lost the race: the squash landed at
`16:53:05Z` and their correction to the body was submitted at `16:53:31Z`, **26
seconds late**. The replaced body still headlined *"Per-process. Not per-user,
not persisted"* — the exact claim retracted everywhere else on that branch —
together with the concurrent-read table presented as evidence and three stale
counts. Had the merge taken the body, the most durable copy of a retracted claim
in this repo would have been the one commit message nobody ever revisits, while
every other site carried the correction.

**Three drafts of this section made three different false claims about GitHub's
body retention, so the fourth does not make one.** In order: that a replaced body
"cannot be re-measured" (false — prior versions are retained); then that GitHub
"retains every prior version" recoverable from a named node (false twice over —
the cap is 100 revisions, revisions can be deleted, and the node I cited holds
the body *after* that edit, not before); then that recovering one "takes a
deliberate query against a surface almost nobody visits" (false — the web UI
shows an **edited** marker with a revision picker to any reader).

**The claim is cut rather than corrected a fourth time.** The rule stands without
it: what the squash captures is a race, which the timestamps above establish on
their own. A mechanism that has been wrong three times and cannot change the
practice is not worth a fourth attempt — the same reason this repo declines
probes that cannot change a decision.

**And the way the second error survived my own check is the part to keep.** I
verified the reviewer's finding by grepping the retained nodes for
`per-process|not persisted` and got a hit on **all five**, which I read as
confirmation. Those words appear in the *retraction* as well as in the claim
being retracted, so the pattern could not distinguish the two states it was cited
for. Every arm agreed and it measured nothing — this file's own oldest rule,
broken by me while checking somebody else's report of a false claim. The
discriminating test was one line away: compare each node's `diff` against the
current body, which shows immediately that the cited node *is* the current body.

The margin is not the point and the rule does not depend on it being small. What
makes a commit message the right home is one structural property: it is **written
once, at the merge, and carries that timestamp**. Verified here — the squash
commit's own date matches `merged_at` to the second. Not *immutable*: history
gets rewritten in this repo routinely, by amend, rebase and force-push, and the
guidance a few sections down is about exactly that. The narrower true claim is
that a commit message is not editable **in place** the way a body is: the SHA is
a hash of the message, so changing one produces a different commit.

**It is not a warrant that anyone verified anything.** A maintainer merging
someone else's PR with a message assembled from a review they did not run
produces an artifact of identical form and none of the standing — and a reader
can see that a message is a commit message without learning whether its author
ran the gate. So the second half is a **stated practice, not an inferred
property**: the coordinator writes the squash message, citing what was verified
and where. Recording "written by the verifier" as the *reason* a commit message
is trustworthy would be this repo's own named failure — a record asserting a
cause the mechanism producing it never distinguished.

The same mutability cuts the other way, which is why the rule is not "bodies are
untrustworthy": it is precisely what let that session correct #26's body *after*
the merge, so the PR page now carries the retraction. One property, both signs.

Two properties of the local reviewer that the GitHub one does not have, both
worth exploiting. It takes an **explicit diff range**, so a review can be scoped
to exactly the commits that will land rather than to whatever the PR page thinks
the base is. And it has **no suppressed-findings channel** — every finding
arrives in the transcript, so the "read the body, not just the thread list" trap
above simply does not exist. What it loses is the second opinion of a different
model on a different substrate; treat a clean local round as evidence, not as
proof, exactly as with a clean Lite round.

What does **not** relax under either reviewer is the local evidence: validator,
citations, units and the relevant integration gates, on the head that will
actually land.

**The cut from six rounds to two was a throughput decision, and it has a
measured cost worth stating rather than burying.** The table below — taken
before the cut — puts the only findings that changed *shipping code* in rounds
one to three. Two rounds therefore gives up a band that has historically paid.
The cost is real, and it is not the same claim as "rounds three onward were
worthless".

**Effort: Balanced for both rounds.** The old curve (Balanced 1–4, Lite 5–6)
existed because late rounds re-read code already cleared. With two rounds there
is no late half — both rounds see novel material, so neither is a candidate for
the cheap pass.

**Effort cannot be set programmatically**, so this is aspirational and every
round is Lite in practice. It is a control on the PR's Reviewers panel and
nothing else: measured across five routes, including a REST field that is
accepted and then silently ignored. Because a request can look like it
succeeded, **read the effort level printed in every review body** rather than
assuming the one you asked for took effect.

The cap exists because the loop has no natural end. The reviewer re-scans
unchanged code every round and surfaces different **suppressed** findings each
time, so "0 new comments" never means "clear" — on #12, rounds four and five
both reported zero new comments and both carried real suppressed findings. A
reviewer that keeps finding things is not evidence the file is getting better.

What the rounds were worth, measured on the first two PRs to run the old
six-round loop. This table is why the cut to two costs something, and it is kept
in its original units rather than restated in the new ones:

| rounds | what they produced |
| --- | --- |
| 1–3 | the only findings that changed shipping code — including the path-quoting comment that exposed `cmd.exe /c start` corrupting 3 of 9 ordinary filenames |
| 4–6 | documentation and contract precision: stale comments, miscounts, a schema bound the runtime did not enforce |
| 6+ | not observed to produce anything, on a diff of +3860/−107 across 29 files |

Read it in the direction that costs, not the one that comforts: it retires
rounds four onward, and it does **not** retire round three.

Declining is a real outcome and must be recorded as one: reply saying why, so
the next engineer sees a decision rather than an omission. Replies are for the
human record — Copilot cannot read them.

**Lite is a floor.** A clean Lite pass means the cheap checks passed, not that
the code is right — an independent deeper read of #16 found a critical
data-loss defect that six Lite rounds elsewhere never approached. That result
gets sharper under a two-round cap, not softer: there are now fewer passes
behind a merge, all of them Lite in practice, and one of them may be absent
entirely when billing is down. **The local gates are what carry a merge; the
review is what improves it.** Treat a clean review as the weaker of the two
signals, and never as a substitute for running the suite.

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

**A green CI tick on this repo has never covered a single line of Word's
behaviour, and cannot on a runner without Office.** This is structural, not an
outage. `.github/workflows/validate.yml` runs `ubuntu-latest`, and its own
header states the constraint:

> Everything here must run without Office installed. The unit suite under
> .github/extensions/office-canvas/test/unit/ is Office-free and runs below; the
> integration suite drives Word through powershell.exe and cannot run on a
> hosted runner — it stays a local gate until a self-hosted Windows runner with
> an Office licence exists.

Three validation steps run behind that tick, after a checkout and a Node setup:
packaging invariants (`validate-extensions.mjs`), citation resolution
(`check-citations.mjs`), and the **Office-free** unit tests.

**Be exact about what the third one covers, because it is not nothing.** The
unit suite checks our *policy*, and checks the part that was actually wrong:
`word-pids.test.mjs` runs on the hosted runner under names like "the kill is
name-checked, because pids are reused" and "a pid that merely appeared during the
run is never killed". That file also states its own limit better than a summary
can — process listing and killing are both injected, so what is asserted is
"which pids are eligible to be killed — rather than any Windows behaviour".

The line falls between **the decisions we make** and **what Word does when we act
on them**. CI checks the first. It has never observed the second: whether
`Quit()` releases the process, what a handle pins, whether a setting persists
into the user's next session, what a save or a share mode does to a file on disk.
Those rest on the **local** gate and the cited probes, and no tick has ever been
evidence for one of them.

**The billing outage makes this urgent, and not in the direction it looks.** Red
is *not* handled carefully here — #36 re-ran a billing-blocked job and concluded
the **tree** had broken it (above). What red earns is not care but
*investigation*: it is conspicuous, so someone eventually asks. Green asks
nothing of anyone. When billing resumes the ticks return meaning what they always
meant, which was never Word's behaviour — and by then the merges that carried a
documented red will look like the risky ones. **Documenting only the red leaves
the more misleading half standing.**

The general rule, which outlives this repo's outage: **a check's name is not its
scope.** `validate` sounds total and is not. Read what the workflow runs before
treating a tick as evidence for anything, and when a PR's behaviour is covered
only locally, say so in the PR body — the reader six months out has the tick and
no way to know what it excluded.

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

**And if that surface cannot be imported, the unimportable surface is the
defect — not a constraint to design the test around.** The rule above was
written from this instance and was not available to follow *in* it:
`asToolError` lived in `extension.mjs`, which calls `joinSession()` at module
scope, so a test importing it would start a session. There was nowhere to stand
where the agent stands. The result was not that the boundary went untested —
it was that the rule quietly degraded into its own failure mode, and assertions
landed one layer beneath the boundary because that was the only layer
reachable. The defect then recurred **after being found**: three fields were
patched at a single call site, while the constructor that dropped them served
21 raise sites in that file alone, every one of them losing whatever it
carried — including both revision tokens. The repair was to set `data` once in
the constructor, and to move the boundary into `src/tool-error.mjs` unchanged
and test it there.

The tell is a module-scope side effect. **A module that connects, joins or
spawns at import time cannot be asserted through**, and every test written
against it will be a test of something else.

The clause has a range, and it is worth stating now rather than after someone
carries it past it. The same discard happens once more at the CLI bridge, where
a tool failure reaches the agent as a bare `Tool execution failed` with the code
and message dropped — the identical defect, one boundary further out, on a
boundary that is **not ours and cannot be extracted**. "Extract it" is no
remedy there; the only move left is a probe that observes what the agent
actually received. So the clause governs boundaries we own, and says nothing
about the ones we merely call.

**Neither side of a boundary can be inferred from the other, and a fix on one
side makes the other look covered.** The Word host set `[Console]::OutputEncoding`
to UTF-8 and never set `InputEncoding`. Reads were therefore perfect — a German
document came back with every umlaut intact — while every non-ASCII character
the agent *wrote* was decoded as the OEM codepage and saved to disk as mojibake:
`Grüße` became `Gr├╝├ƒe` in `word/document.xml`, with `U+00FC` absent. The
existing line was not an oversight; it was **evidence that someone had
recognised the boundary needed an encoding**, and that is precisely what stopped
anyone looking at its other half.

**The same defect has a zero-distance form: a comment asserting that the far side
works, written from the near side by someone who could not observe it.** #36's
diagnostic writer carried *"diagnostics go to stderr, which `word-host.mjs`
captures and surfaces."* Measured with a control: the host emitted the decline,
and the extension's log callback received **nothing at all** — stderr was
buffered into a ring with exactly one consumer, in `#onExit`, which decorates
rejections for in-flight calls, and a clean dispose has none. The fact and the
code contradicting it were the same sentence, so nothing had to drift for it to
be wrong. This is why a fix that exists is worse than no fix when it makes the
neighbouring half look covered.

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
taken.** Cross-session messages queue and have arrived here **~12 hours** late,
and the PR-state notification attached to one is frozen at send time along with
it. A correction arrived asserting that #22 was open with `main` at #21, noting
it was the third time the merge had been relayed, and correctly declining to
rebase on that basis.

The timestamps are worth stating exactly, because the obvious reading of this
is wrong. The message's state refresh is stamped `23:37:42Z`; **#22 merged at
`23:45:48Z`, eight minutes later**, and it was read some twelve hours after
that. So the correction was not a stale claim written carelessly — it was
**accurate when written and went stale in the queue**, while the claim it
corrected had been premature when made and became true shortly after. Both
parties right about different instants, which no amount of care about *facts*
prevents, and which re-checking at the writing end would not have caught
either.

The transportable half: **a check's conclusion travels and its timestamp does
not.** *"#22 is still open"* reads as a property of #22; it was an observation
with an expiry, and the expiry is what gets dropped — the same scope-discarding
this file keeps finding, with time as the discarded dimension rather than
applicability. Which is also why the SHA discipline here works and nothing
equivalent has ever worked for status: **a claim about the tree can be pinned by
a commit, and a claim about a pull request's state cannot be pinned by anything
a message is able to carry.** The two are not interchangeable evidence, however
alike they look side by side in one report.

So: re-read state immediately before acting on it, and never carry a status out
of a message and into a mutation. Not an argument for re-verifying everything —
narrower than that. A claim about *mutable remote state* either carries when it
was observed or it is not a claim.

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

**Check a session for unpushed work before archiving it — nothing on the remote
side can.** The tempting check is to compare the PR's recorded `headRefOid`
against `git rev-parse origin/<branch>`, and it is unsound: it compares two
*remote* facts to each other and is silent about the machine. The counterexample
appeared in this repo's own working tree while the paragraph recommending it was
being written:

```
PR #56 headRefOid               3f6b813
origin/dahlweid-verbose-tribble 3f6b813   <- equal
HEAD                            097cf50   <- one commit, unpushed
```

`git rev-list --count origin/<branch>..HEAD` answers it, as do `git status -sb`
and `git cherry -v`. Run one of them *before* the archive, not after.

**A tree-equality check was drafted here and cut.** A squash does carry the tip's
tree when `main` has not advanced — measured, `5ef2403` and `f8d48fa` both
`1330d69` — but that instance held only because the tip was itself a merge
commit squashed twelve minutes later. Where `main` advances and the squash must
three-way merge, the trees differ with nothing stranded, and the check raises a
false alarm. One agreeing instance measured nothing about the case that matters,
so the rule is not stated rather than stated with conditions that were never
probed.

**The reason to run any of this:** a session being archived says nothing about
whether its work is finished. Three were archived here within one session's span
— that one was complete, while **#29 still had 7 argument-form `.Quit()` call
sites across 6 files on `main`**, excluding comments and the probe that
deliberately measures the throw. The archival looked identical in both cases.

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
  the one direction of this defect that reports good news. **It has a real
  instance** (#26): a mutant scored `KILLED` by two assertions that both read
  *"the bait paragraph is missing"*, so the assertion actually under test never
  ran and the kill was attributed to assertions that did not execute. It
  surfaced in a
  *hand-run* check, which is the one context with no runner to classify it.

  **And its cause was reported as two-way when the evidence showed three**, which
  is the more useful half. The first explanation — the bait corrupted crossing
  stdin (#40) — was present, sufficient, and stopped the search. Nobody asked what
  *else* produces that message. A **successful** substitution does: the test
  located each bait by its own first twelve characters, and for all three baits
  that prefix spans the substitution site (`He said "hel`, `A dash -- li`,
  `Copyright (c`). Word rewriting the bait destroys the locator, so the two
  assertions that name the defect were unreachable **for the case they exist to
  catch**, on any Word in any locale — independently of the encoding defect and
  of the arms that rewrite nothing. **A sufficient cause is not an exclusive
  one**, and a first explanation that fits is exactly what stops the second from
  being looked for. Fixed by anchoring on an autocorrect-invariant token, with the
  anchor shape enforced mechanically so a weak anchor is a hard failure rather
  than a silent narrowing.
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

The remedy common to these three is to treat the mutant itself as the thing
under test: assert that the mutation was applied, and that the *specific* test
you expect to go red is the one that did. A gate that only counts failures
cannot tell a killed mutant from an unrelated breakage.

**And a surviving mutant is not always a missing test.** Sometimes it is two
branches being proved by one shared assertion: deleting a null-identity guard
reddened nothing because the test asserted a *shared* outcome — the Word is still
alive — which a second, independent guard satisfied on its own (#36). Both
branches were covered by an assertion that could not say which one had done the
work. The remedy there is **attribution**, not a stronger assertion. Note also
what did *not* catch it: the two-independent-records discriminator, because there
was no second record of the quantity to disagree with the first.

**The same shape occurs outside the mutation gate, and the instrument is
whatever the step actually depends on.** Three more, each of which completed and
reported:

- **The harness could not run the subject.** Word probes launched through
  `Start-Job` wedged: every arm that reached `Document.SaveAs2` hung
  indefinitely, twice, costing two full probe runs before the harness rather
  than Word was suspected. The obvious cause was apartment state, and a control
  excluded it — the identical body in a real `powershell.exe` returned from
  `SaveAs2` in **149 ms under `-STA` and 138 ms under `-MTA`**, so a real MTA
  process saves fine and the job runspace's thread is what wedges. Scope, since
  it is easy to carry this too far: `SaveAs2` is what was measured, and *"a COM
  call that pumps a message loop wedges there"* is the inferred mechanism, not a
  call-by-call result. Four probes in this repo still use `Start-Job`; they are
  **unaudited, not known-broken**. The working rule is that a probe here starts
  a real process with discrete argv.
- **The guard could not fire.** A BOM self-check (#43) had two conditions, and
  the first — `"$([char]0x2014)".Length -ne 1` — builds the character from a
  *number*, so it is 1 however the file was decoded, and measured 1 in both
  arms. The entire guard was the second condition. Nothing was broken, and the
  hazard is the ordering: **the dead condition came first and read as the
  primary check**, so anyone simplifying that line would have kept the inert one
  and the guard would have died silently. It is the inverse of two live guards
  masking a weak test — one live and one dead, with the dead one the more
  plausible-looking of the pair.
- **The range never reached git.** A runbook step said *"count the commits you
  expect"* — `git rev-list --count $fp..HEAD`. In PowerShell that returns **0**
  on a branch with commits to count, and **exits 0 without a word**. Reproduced
  here: a variable expansion abutting `..` is split, so git receives **two**
  arguments, `<sha>` and `..HEAD`. The wreckage is still valid git —
  `<sha> ^HEAD HEAD` — and since a fork point is by definition reachable from
  `HEAD`, **0 is the honest answer to the question that was actually asked.**
  Two things make this worse than an ordinary quoting bug. `${fp}..HEAD` splits
  too, so the reflex fix does nothing; and `main..$br` does **not** split, so
  the form everyone types daily is safe and the hazard only appears once the
  variable moves to the **left** of the range — which is precisely what a reader
  does when substituting a shell variable into a documented
  `<old-head>..HEAD`. Quote the whole range, and state the expected count next
  to the command so a `0` reads as wrong rather than as *already done*.

**And when a convention gets rediscovered, that is evidence the first guard was
unfindable.** `param([string] $Doc)` type-constrains that name for the whole
scope, and PowerShell variable names are case-insensitive — so a later
`$doc = $w.Documents.Add()` silently **coerces the Document to a string**
instead of failing. `$doc.Content` then reads `$null`, and the error surfaces
two lines further on as *"the property Text was not found on this object"*,
naming a line that is entirely correct. Same class as the rest of this section:
the failure names a cause the code never distinguished.

This has now been hit independently twice, and it was **already** guarded the
first time — `make-fixture.ps1:92` carries a comment refusing `$table` for
exactly this reason. A guard that exists as a comment at one call site is
findable only by someone already reading that site, which is no one who needs
it. Rediscovery is the signal worth acting on, not the fix.

Auditing for it is cheaper than it looks, and the mechanical check is the one to
reach for: **does any asserted outcome have more than one producer?** Counting
raise sites per error code is a grep and needs no judgement, where "is this test
attributing correctly" needs plenty. Where every code has exactly one raise site,
the shape is impossible by construction.

**"The remedy is cheaper than the measurement" assumes you already know which
remedy you need.** Recorded because it was wrong here, in coordination advice I
gave. Faced with a TOCTOU between an identity check and a kill, I recommended
skipping the measurement and going straight to a P/Invoke that opens one handle
and does both through it — sound, and it would have been carried. #36 measured
the cheap half first, on the grounds that it decides whether the P/Invoke is
needed at all:

| arm | `Kill()` on a process allowed to exit between the calls |
| --- | --- |
| one `Process` object, `.Handle` never touched | **access denied** — it re-opened the pid |
| one `Process` object, `.Handle` touched first | "the process has exited" — it used the retained handle |

Two different failures from the same sequence, which is the discriminator:
`Kill()` performs its own `OpenProcess` unless a handle is pinned. So the
reviewer's proposed remedy — capture a single `Process` object — does **not**
close the window, and the real fix is `$null = $p.Handle`: one line, no
`Add-Type`, because Windows will not recycle a pid while a handle is open. The
measurement was cheaper than the remedy I recommended *and* found a cheaper
remedy than either candidate. Skip a measurement to save effort only when the
remedies are already known to be the same.

**When no reachable bound lies outside the distribution, stop looking for a
number.** The `Quit()`-to-exit spread is 779–3702 ms across three sessions, with
one Word observed surviving a 30 s poll — and the RPC layer imposes a **20 s**
ceiling above which the client kills the host outright, so a 30 s deadline is not
merely generous but unreachable: it expires by being killed mid-wait. There is no
value that makes teardown a guarantee. The deadline is therefore a **budget whose
expiry must be safe**, not a bound that must hold — here, falling through to an
identity-proving kill. Nudging such a number upward is the move that looks
measured and still straddles.

**And that fix is only half a fix without the one above it.** A budget whose
expiry is safe is safe only if the kill it falls through to cannot hit a
recycled pid — so pinning the handle is not a separate tidy-up, it is the
precondition that makes safe expiry actually safe. Adopt them together; read
separately, each looks optional.

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
rendered.

**"Re-run rather than recall" only covers half of it.** Over-generalised prose
has a measurement behind it, so re-running at the new scope settles it. Prose
that was never measured has nothing to re-run, and the failure is upstream: not
noticing the sentence made a checkable claim at all. `check-citations` covers
only the easy half — it verifies that a probe *reference* resolves, and says
nothing about whether the surrounding sentence is what the probe found; the bait
list that started this had no citation, so nothing was even eligible. The second
half is caught by asking *"is this a measurement or a memory?"*, which is a
habit and not a gate. And a guard that cannot fire today is still worth keeping
if it is labelled as one — the label is the part that has to be true.

**There is a third kind, and neither remedy reaches it: a sound measurement
whose recommendation ranges past what was measured.** The `Quit()`-to-exit probe
measured teardown correctly, twice, with its conditions recorded beside the
figures — and then concluded that the variance "argues for a deadline far
outside the distribution", which ranges over a **client-side timeout the probe
never observed and had no way to observe**. Re-running does not catch it: the
measurement re-runs green. Noticing does not catch it either: the prose does cite
a real measurement, and cites it accurately. The failure is at the edge of scope
rather than in the number or in the memory. Two of the three corrections against
that probe are now about its prose rather than its figures, which is the pattern
and not a coincidence. **Measuring soundly does not license recommending past the
edge of what was measured, and nothing about a correct measurement flags when its
prose has stepped over that edge.**

**What travels is rarely the measurement; it is a one-line summary of the
measurement, and a summary is exactly what discards a scope.** #36 declined a
`ReleaseComObject` finding partly on the grounds that this repo had measured
`ReleaseComObject` *causing* a leak. Reading the cited source instead of the
citation: #16 round 5 measured it on a **Protected View window** RCW, **across
operations** of a long-lived host, and explicitly refused to pin a mechanism
beyond Protected View's uninstrumented second Word. It had been applied to an
**`Application`** RCW in a script that exits milliseconds later — different
object, different lifecycle, named mechanism not in play. Nothing about the
original note was wrong; it is correctly scoped. The summary of it was not, and
the summary is what shipped.

**A cheap discriminator for this, worth reaching for first: does the general
claim convict your own codebase?** One grep settled it — `Stop-Word`, in the very
file doing the citing, calls `ReleaseComObject($script:App)` itself. If the #16
result generalised, it would condemn the host's own teardown inside the PR that
exists to fix that teardown. **A general claim that would convict the code
around it is not a general claim.**

**The cheapest instruments fail this way too, and they are the ones nobody
checks.** Two within ten minutes on #26, and both recur here: a summary grep
written as `^# (tests|pass|fail)` against a reporter that emits `ℹ tests 284`
**matched nothing and exited 0** — a check that printed no verdict and read as a
pass; and `tools/validate-extension.mjs` for `validate-extensions.mjs` **exited
1**, which reads as "validator red" and means "gate never ran". That is the
runner-cwd trap again with one letter instead of one directory. I hit the same
grep twice in one session after documenting it, and then posted this very
paragraph to a PR with `gh api -f "body=@file"` — where `-f` means *literal* and
only `-F` reads the file. It exited 0 and returned a plausible comment URL for a
comment whose entire body was a 41-character Windows path. **Assert that the
check produced a verdict, not merely that it exited well** — a run that matched
nothing and a run that passed are the same observable, and a call that succeeded
at posting the wrong bytes looks exactly like one that worked.

**The sharpest instance of that shape is the test runner itself, and the exit
status is no help.** Measured here on Node v24.18.0, same tree, same moment:

| invocation | output | exit |
| --- | --- | --- |
| `node --test "**/test/unit/*.test.mjs"` | 8 lines, `tests 0 … pass 0 … fail 0` | **0** |
| `node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"` | `tests 258 … pass 258 … fail 0` | 0 |
| `node --test <directory>` | `tests 1 … pass 0 … fail 1` | 1 |

The first form is the dangerous one: it ran **nothing** and said so only in the
body, never in the status, so a `&&` chain, a CI step or a person reading the
last line all see a pass. The third is its inverse — a `MODULE_NOT_FOUND`
wearing the shape of a test failure, which sends you looking for a broken test
that does not exist. **Assert on the `pass` count, never on the exit code**, and
treat `tests 0` as red.

**A figure another session hands you is a memory, not a measurement — including
when the session is you.** I reported this trap as *"zero output lines"*. The
session I gave it to declined to write it down without reproducing it, and was
right: it is eight lines, not zero, which is the difference between "the
instrument is silent" and "the instrument is answering a question nobody
asked". The claim survived only because someone refused to transport it. This
sits one level above *"is this a measurement or a memory?"*: applied to received
claims, the answer is **always memory**, whatever it cost the sender to obtain.
The coordinator is the worst offender by construction, because broadcasting is
the job.

**A gate run on a head that cannot merge measures a tree that will never
exist.** Merge-readiness is a *precondition* of running the gates, not a check
performed after they come back green. I asked a session for full gates on a head
that was `CONFLICTING` at the moment I asked — and the conflict had been caused
by my own merge of another PR, reported in the same message as an aside, four
paragraphs below the words *"Verified: MERGEABLE"*. Had the session complied,
every number would have been truthful and about nothing. Re-read
`mergeable`/`mergeStateStatus` immediately before asking, and again immediately
before merging; `gh pr view` lags, so cross-check `headRefOid` against
`git ls-remote`.

**A control proves an instrument can see. It does not prove it is pointed at the
right thing.** The `ReleaseComObject` probe A/B'd one operation, found the file
free in 16 ms either way, and carried exactly the control this file demands — 
leave the document open, confirm the file is *still held* at 30 s — which
passed, demonstrating the instrument detects retention. Every reading was true.
The damage was on an axis the probe could not have seen: across *operations* of
a long-lived host, adding the release scored **19/20 with 7 WINWORD left alive**
against **20/20 and zero** without it. The control answered "can this detect
retention?"; the honest question was "can this detect retention **across
operations**?", and nothing in the setup forced anyone to notice those differ.
This is the `FileShare` lesson one level down, and it was walked into by the
same person hours after writing the probe that exposed it. **State what the
control ranges over, in the same breath as the control.**

**A fork point is a property of history, not of a branch, so record it as a
literal SHA and never re-derive it.** The rule it replaces was mine and sounded
stronger: *confirm the fork point with `merge-base` rather than trusting a
number I gave you*. `merge-base` answers "most recent common ancestor of these
two branches". A rebase target answers "where was this branch forked". Those
coincided for six consecutive checks on #43 and decoupled the moment the lower
branch rebased — silently, with a plausible SHA and exit 0. Measured across the
three endpoints anyone would reach for:

| `git rebase --onto <target> <endpoint>` | commits replayed |
| --- | --- |
| literal `cee2da6` — the recorded fork point | **9**, correct |
| `4da84c2` — what `merge-base` returns *now* | 14 |
| `dahlweid-create-document-tool` — the branch name | 14 |

The five extras are the lower branch's own pre-rebase commits, which is the
exact pile `--onto` exists to prevent; it returns through a different door the
moment the endpoint is named indirectly. **Re-deriving is not a stronger check
than remembering — it is a different question that agrees until someone
rewrites history.** What caught it was the cheap guard: *count the commits you
expect*. 14 against 9, no error and no warning anywhere.

**And never rebase onto a branch that is about to be squash-merged.** Squash
replaces N commits with one new commit that is none of them, so a branch rebased
onto the pre-squash head carries ancestry that `main` will never have, and its
next rebase must replay commits already present as content. Rebase onto `main`
after the squash lands, or let the forge's cascading merge retarget the PR.

**A review scoped to a superseded head reviews a tree that will not ship, and
whether that matters is decided by ancestry, not by apology.** Both local rounds
on #46 were scoped to `6fdd8d40`; the head moved to `645e072` mid-review. The
question is not "did it move" but "how":

```
git merge-base --is-ancestor <reviewed-head> <current-head>
```

Exit **0** means the new head *contains* the reviewed one — an addition, so the
completed review still holds for everything it covered, and only the delta needs
attention. Exit **1** means a rebase, so ancestry can no longer vouch for the
content and the verdict is **undecided rather than void** — see *What survives a
rebase is decided by what the artefact measures* below, which is the check to
run next. On #46 it was 0 and one commit of comments; on #26 the same check
against `771a776` returned 1, and there the artefact was a set of gate numbers
over code that #33 had rewritten, so they had to be re-requested rather than
re-quoted. **Run this check before deciding a stale review or a stale gate
result still counts**, and re-scope any round still running instead of letting
it finish against the old range.

**A reviewer is thorough within the frame it is given and does not interrogate
the frame.** First evidence on replacing the forge's reviewer with a local one:
round 1 on #46 spent ~15 minutes and 50+ tool calls, named a concrete mutation
for every new assertion, byte-scanned all nine changed files for stray non-ASCII
literals, recomputed the cited offsets against the head blob, and checked that
an assertion's snippet window actually reaches the character it claims to test —
attacking test *vacuity*, this repo's most recurrent defect class. It returned
clean. It also confirmed a mojibake signature as "arithmetically correct under
CP437" — correct arithmetic on a codepage nobody had measured. The **author**
then ran `chcp`, found 850, and retracted the claim, showing the assertion
survives because both codepages map `0xC3` to `U+251C`. The reviewer checked the
arithmetic and never asked whether the premise had been measured. That is the
`FileShare` lesson a third time, now with a reviewer in the loop: **a conclusion
can be right while its mechanism is wrong, and nothing fails, because the
observable is identical either way.**

The operational consequence, given a two-round cap: **on an unchanged tree a
second identical pass re-reads the same diff.** Round 2 must change the angle —
a different model, aimed at what round 1's framing structurally could not see —
or it is a formality. On #46, round 1 proved `[Console]::InputEncoding` is set
at line 60 and precedes first use; it never asked whether that assignment can
*throw*, which on Windows it can when stdin is a redirected pipe, which is
exactly how this host is spawned. Same line, same diff, question round 1 was
never pointed at.

**Do not ask for a WINWORD census delta. It is below this machine's noise
floor.** The gate was specified here and was wrong. Asked for it on #26, L3 got
`14 -> 15` on one run and `14 -> 14 with identical pid sets` on an earlier run
of the same suites on the same tree — two answers, one tree, which is the tell.
So they sampled the process table every 15 s for 75 s **while running no Word
code at all**:

| Z | count | appeared | disappeared |
| --- | --- | --- | --- |
| 12:38:57 | 14 | — | baseline |
| 12:39:12 | 15 | 50496 | |
| 12:39:27 | 16 | 27928 | |
| 12:39:42 | 16 | 44828 | 27928 |
| 12:39:57 | 14 | | 44828, 50496 |
| 12:40:12 | 14 | | |

Four WINWORDs born and three died in 75 idle seconds, amplitude **±2** around a
baseline of 14. A ±1 delta across a multi-minute run therefore resolves nothing,
and — the half that matters — **it would equally fail to show a real leak.** A
census that does not move is one draw from a distribution whose amplitude
exceeds the signal, and it had been read as a positive result when it is a
non-event. **Ask for the per-suite leak assertion `assertNoLeakedWord` in
`word-pids.mjs` by name and pass count** — but for the right reason, which is
not the one first written here.

**The reason that assertion beats a census is its 90 s polling deadline, not
attribution.** The justification originally recorded — *"it diffs over seconds
and attributes by the pid the host itself started, so other sessions' churn
cannot enter it"* — is wrong in both clauses, and was corrected by the session
that was asked to run it. `pidsBefore` is captured before the first test and
asserted after the last, so the window is ~1.5–2.5 minutes, the **same order as
the census**. And `newWordPids` (`word-pids.mjs:59-61`) is a pure set difference
with no ownership predicate; the `ledger` only splits the failure *message* into
owned versus unattributed, as its own docstring says: *"Both still fail."* The
text even names "another session's" as a possible cause of a red.

Issue #25 later made attribution sound **at its source** — `Initialize-Word`
reads its own window handle instead of differencing pids — and it is worth being
explicit that this does **not** revive the claim above. What changed is the
meaning of a ledger entry, and therefore what `killOwnedWord` is entitled to
kill. `assertNoLeakedWord` still differences, deliberately: it has to catch the
leaks attribution *cannot* see, such as the second WINWORD that
`ProtectedViewWindows.Open` spawns without the bridge ever holding a handle to
it. So the mechanism below is unchanged, and the corrected sentence stays
corrected.

What actually does the work is `timeoutMs = 90000` at `word-pids.mjs:110-119`:
it polls and fails only if a pid is **still alive at the deadline**, so a
foreign Word that exits during the poll cannot fail it, whereas a single-sample
census counts it. That is the mechanism, and it is worth stating what it is
*not*: the deadline does not outlast foreign churn in general. The file's own
comment at `:95-102` says exit latency **is not bounded** by the idle
measurements, that Word's shutdown contends on per-user state, and that one run
had a Word survive a **30 s** poll and exit on its own. 90 s is generous — free
on green runs, so there is no reason to shave it — not proven sufficient. The
distinction is load-bearing: the assertion is not immune to other sessions, so
a residual false red is possible — that is #37 — and anyone shortening that
deadline or widening the capture window on the strength of "background churn
cannot enter it" would be acting on a mechanism that does not exist. The
property that matters holds either way, and it is the one the census lacked:
**it can go red for a real leak**, because a Word we minted and dropped is
still alive at 90 s.

**An API that answers "was your instruction accepted" is not answering "is the
world now in that state".** Third instance in this repo, each in different
clothes: `Quit()` returns ~3.1–3.7 s before `WINWORD` exits;
`ReleaseComObject` returns `0` while the release is what leaks the process; and
`Suppress-AutoCorrect` in #26 returned `suppressed = $true` meaning **five
property assignments did not throw**, never reading one value back. Treat any
success flag derived from "the call did not throw" as unmeasured until something
reads the state back.

**A claim discharged onto a neighbouring test is an unchecked citation.**
`create-smoke.mjs:98` told the reader that a green bait proves nothing and that
*"the evidence for that is the separate settings read-back check"*, naming a
line — and that line asserted the flag, not the state, so it never carried the
claim. `check-citations` guards `probe-*.{ps1,mjs}` paths only; **prose pointing
at another test is unguarded and cannot be mutation-tested.** What made it
invisible is worth the whole entry: the bait arms rewrite 0 of 6 with every
setting ON, so they are inert here by construction, and the read-back stays
green because a flag reads back what was written. **Two green checks, neither
able to go red for that defect, each looking like the other's backstop.**

**A merge gate measures what it measures, and "the headline claim is unproven"
is not in it.** #26 reached the point of merge with gates green at the exact
head, four review rounds, past the cap, `MERGEABLE` confirmed from both `gh` and
`git ls-remote`. Every reading was true. The squash was issued **twice** and
refused both times — GraphQL and REST alike — solely because the PR is position
1 of a stack and stacked PRs require `POST /pulls/{n}/merge-async`. The owning
session found the defect independently, minutes later. **What prevented a PR
titled "autocorrect suppressed" from landing with nothing able to prove
suppression was an unrelated API constraint about stacks**, not any gate in this
file. Before merging, ask what assertion would go red if the PR's *title* were
false; a full green board does not answer that question and will not raise it.


### The autocorrect claim was wrong in the other direction too

The read-back defect above was the smaller half. The recorded finding that the
autocorrect settings are **per-process, not per-user and not persisted** is
false. Measured sequentially: set a value, **quit the writer**, start a fresh
instance, read -- the value is there. All five, every time.

**The original probe could not have found this, and it looked like a clean
result.** It switched the settings off on instance A and read them from a second
process B *while A was still alive*. Two observations, and only two: a
concurrent reader sees the old value, and a reader started **after the writer
quits** sees the new one. Neither locates the flush in time — B may have cached
its own copy at startup, so "persisted late" and "persisted early but read from
a cache" are not distinguished, and nothing here needs them to be. What the pair
does establish is the thing that matters: **per-process isolation and
persistence are the same observation to a concurrent reader.** The instrument
could not distinguish the two states it was cited as distinguishing. That is the
`FileShare` share-column defect again: a probe blind to the property its own
construction puts on the near side, every arm agreeing. The discriminator is one
line -- quit the writer before reading -- and it was missing because holding A
open is the natural way to write that test.

The consequence was not theoretical. A fresh Word on this machine read all five
settings **off**, and `Suppress-AutoCorrect` ran on every authoring run, which
is a sufficient cause. Stated at the width of the evidence: this is the value on
**this profile**, before and after, measured from a fresh instance. Word's
shipped default for these was not measured — that would need a reset profile —
so "we turned them off" rests on the run history plus the observed state, not on
a platform-wide default. Verified independently from the coordinator with a
read-only check -- no assignments in it, so it could not perturb what it
measured.

Three rules came out of it, and all three still hold as rules even though the
code they were written for is gone.

**Cleanup belongs inside the operation that made the mess, not in the teardown
that may be skipped.** Save-and-restore at quit was the obvious repair and is
defeated by our own measured kill path: `#send("quit", {}, 20_000)` expiring
kills the host, so the restore is protected by exactly the code least likely to
run. A repair a known failure path skips is not a repair. Suppress, author and
restore inside the single operation, in a `finally`, while the COM object is
still held.

**Restore to the state you found, never to defaults.** The true original was
never recorded, so restoring "Word's factory values" would be a second
unrequested write dressed as a repair. Found-state is auditable; a guessed
original is not, and it silently destroys the evidence of what we did.

**A feature whose success state is "nothing happened" needs its instruments
proven against a forced failure, every time** -- because the passing case and
the broken case are the same observation. Both autocorrect checks that turned
out to be unable to go red were in this one area, and that is why: suppression
working and suppression silently failing look identical unless something is
made to fail on purpose.

### Two Words alive at once: whichever exits cleanly last decides the value

The retraction above left one thing open, and #51 named it: capture-and-restore
is a save/restore against a store that a *second* Word is also holding a copy
of. If the flush is last-writer-wins, a correct restore can still lose to a
user's Word flushing its own stale copy afterwards. Measured rather than
reasoned about, in `spikes/isolation/probes/probe-autocorrect-concurrency.mjs`
-- seven arms, each seeding the store to all-five-ON, and each reading the
result from a **fresh** instance started only after every Word it created had
left the process list.

| what the arm did | a fresh instance afterwards reads |
| --- | --- |
| lone writer wrote OFF and quit (the control) | **OFF** |
| host killed, its Word still up | **ON** |
| that orphan then killed | **ON** |
| writer's Word killed outright | **ON** |
| user's Word up first, writer quit, user quit last | **ON** |
| user's Word up first and quit first, writer quit last | **OFF** |
| `create_document`, user's Word quitting after it | **ON** |
| user's Word opened *mid-suppression*, quit after our restore | **ON** |

Three answers, only one of which was the question.

**Last clean exit wins.** Arms 4 and 5 are the same two instances in the two
orders and they disagree, which is the whole finding: the value a user's next
Word sees is the one held by whichever instance exited cleanly last. So
capture-and-restore was never a guarantee here. It was a best effort against
shared state, and the case it loses -- a user's Word that was already open and
exits after us -- is the ordinary one, not a corner.

**A killed Word persists nothing.** Our restore lived in a `finally` that a host
crash skips, so the obvious repair was a durable ledger of priors, written
before the suppression and replayed at the next start. It was never needed: what
a killed Word was holding is discarded, so a crash mid-authoring left the user's
stored settings exactly as they were. The crash path was safe because it never
reached the store, not because anything cleaned up after it. This contradicted
the brief that commissioned the work, which named the kill path as the residual
hole no restore could cover.

**A concurrent reader sees the store, not the other instance.** Arm 7 was
written expecting the dangerous ordering -- a user opening Word *during* our
suppression, reading our OFF, and carrying it out past our restore -- and it
went red against a correct product, because that Word reads the stored ON
instead. The live value in our instance is not visible to it.

No mechanism is claimed by any of this. Every row is a reading, and the arms
that would distinguish "flushed at exit" from "flushed earlier and read from a
cache" were not run because nothing turns on the answer. Asserting the flush as
though it had been measured is the exact error that produced the retraction
above, and it is easy to make twice.

Also measured, incidentally: an orphaned Word whose host was **killed** does not
exit on its own. It was still in the process list 45 s later and had to be
reaped by hand. This does not contradict the account at `word-host.ps1`'s quit
site, which concerns a host that *exited normally* after releasing its COM
reference -- that Word did go. The distinction is the kill: a terminated process
runs no release. The pid ledger is load-bearing on the kill path, not
belt-and-braces.

### The strongest fix was to stop writing the settings at all

Everything above is about managing a hazard. The measurement that removed it is
`spikes/isolation/probes/probe-autocorrect-necessity.ps1`, and its result is
that the suppression bought nothing on any path this host uses.

**Autocorrect and autoformat-as-you-type are typing features, and this host
never types.** Every character reaches a document through `Set-ParagraphText` ->
`Range.Text` assignment; there is no `Selection.TypeText` anywhere in
`word-host.ps1`. Eight baits, all five settings ON, inserted through exactly
that path: all eight came back verbatim. `Content.AutoFormat()` in the same
process rewrote three of them, so the instrument can see a rewrite.

**The scepticism that had to be answered first was about the baits, not the
result.** The earlier "0 of 6 rewritten with every setting on" read stronger
than it was, because its positive control only rewrote 3 of 6 and the probe
itself raised the reason: a German Word ships a German AutoCorrect list, and
`teh` -> `the` is an English entry. Enumerating `AutoCorrect.Entries` on this
machine settled it -- **402 entries, German, and `teh` is not among them**. That
bait was inert. It measured nothing, which is the same instrument defect as the
retraction above, found in the same place twice. The replacement baits were
drawn from the machine's own list (`(c)` -> `©`, `abeiten` -> `arbeiten`,
`adnere` -> `andere`, `adneren` -> `anderen`) so they are provably live here.

**Coverage, stated per feature, because the five are not one mechanism.**

| setting | covered by a live bait? |
| --- | --- |
| `AutoCorrect.ReplaceText` | **yes** — four triggers read out of this machine's own list |
| `Options.AutoFormatAsYouTypeReplaceQuotes` | **yes** — `Content.AutoFormat()` rewrote its bait in the same run |
| `Options.AutoFormatAsYouTypeReplaceSymbols` | **yes** — same |
| `AutoCorrect.CorrectSentenceCaps` | **no** |
| `AutoCorrect.CorrectInitialCaps` | **no** |

The last two are keystroke handlers with no programmatic trigger of any kind, so
their baits went in verbatim but nothing available here can show they *could*
have been rewritten. That is a gap in the evidence and is written down as one
rather than argued away.

**What actually settles it is about this codebase rather than about Word.**
`Cmd-Edit` has never suppressed anything. Every `edit_document` writes through
the identical `Range.Text` assignment with all five settings ON, and
`test/integration/edit-smoke.mjs` already asserts the text lands verbatim on
that path. The repo already depended on this measurement everywhere except one
function. Suppressing in `Cmd-Create` alone was half-bought insurance whose only
distinctive effect was editing the user's Word.

So `Disable-AutoCorrect`, `Restore-AutoCorrect`, the `$acState` plumbing and the
`autoCorrect` result field are all gone, and the safety rule in the file header
went back to the simple form it should never have left: **these settings are
never written, on any instance.** The guard is
`test/unit/autocorrect-not-suppressed.test.mjs`, which is a source assertion and
therefore has to prove its own detectors can fire -- each one is run against a
synthetic source containing the thing it looks for, in the same test, because
"this string is absent" is the exact shape that passes vacuously.

**The general lesson is the one worth carrying.** A safety mechanism protecting
against a risk that was never measured is not free: this one bought nothing and
cost shared mutable state with the user's own application, plus two retracted
claims and four rounds of work. Before hardening a mitigation, measure whether
the thing it mitigates can happen on the paths you actually use.

### What survives a rebase is decided by what the artefact measures

The ancestry check above is a **sufficient condition, not a necessary one**, and
reading it as necessary retires evidence that is still good. `merge-base
--is-ancestor` answers whether the reviewed commits are *contained* in the new
head — nothing about trees. Exit **0** settles it cheaply: the reviewed content
is present verbatim, so the verdict holds for what it covered. Exit **1** does
not settle anything; it means the question is open and has to be answered by
looking at the artefact.

The sharper rule for that case: **a review of content survives when the rebase
does not touch the hunks it read; a gate does not survive when the code it
exercises was rewritten.** Same pending rebase, opposite conclusions for a
content review and a Word-teardown gate, and the discriminator is whether the
thing under test changed — not whether the ancestry did.

With one caveat that keeps this honest: unchanged hunks are not sufficient on
their own. A review of code whose callers, invariants or dependencies moved
underneath it is stale even though its own lines are byte-identical, which is
why "did the subject change" is a judgement about the artefact's *reach* rather
than a `git diff` on its line ranges.

### A test double that cannot produce the state under test

The probe rule has a unit-test form, found in this repo's own viewer code: a
fake cache returned `changed: true` unconditionally, which made an early-return
branch unreachable -- including from the test named after the bug that branch
causes. And an assertion naming exactly the right invariant,
`owners.length === text.length`, could not fail because every fixture was ASCII,
where code points and UTF-16 code units coincide. **The arms agreed by
construction, in the fixtures rather than in the harness.** When a test double
returns a constant, ask which real state it has made unreachable.

### A correction is a claim, and a restated number has no instrument

Three instrument failures in one session, all self-caught, all worth keeping:
a verification step searching for `[regex]::Escape`-d backslashes with
`-SimpleMatch` and reporting four good fixes as **LOST**; a file count restated
in prose as `30` when it was 29, minutes after being measured; and a PR body
overwritten by composing destructive input inline. **Prepare the artefact, then
submit it** -- write the payload to a file and pass it, never compose it in the
call that destroys the old value.

Two more from the coordinator, in one hour, both **ad-hoc verification steps
that returned a plausible wrong number rather than failing**:

| check | reported | actual | why |
| --- | --- | --- | --- |
| `git show HEAD:f \| Out-String`, counting `` `r `` | CR=1357 | **CR=0** | the PowerShell text pipeline re-inserts CRLF on the way out, so the number was the *line count* wearing the label of a defect |
| `Select-String -Context 3,3`, then projecting `.Context` in a pipeline | the block appears twice | **once** | one `MatchInfo` is emitted per matching **line**; the default formatter merges overlapping windows, but a manual projection emits one block each, so an overlapping region is duplicated |

The second is the more dangerous, because a duplicated block is exactly what a
genuinely duplicated code site looks like -- it was read as "the comment appears
at two sites" when it appears once, at `structure-map.mjs:266`. **The display
de-duplicates and the object model does not.** Running the same command
interactively shows one clean merged block and hides the defect completely; only
a script that projects `.Context` itself ever sees double. Measured on a
purpose-built fixture: two alternatives matching on two adjacent lines yield
**2** `MatchInfo` objects with overlapping windows, which `Out-String` renders as
**one** block and
`$_.Context.PreContext + $_.Line + $_.Context.PostContext` renders as **two**.
Count occurrences with a plain loop over the lines instead.

Both errors were caught only because the number was load-bearing enough to
re-measure with a second instrument, which is the whole rule: **a verification
step is evidence code, so it needs stricter discrimination than production
code, not looser.** And the second one twice over -- the first correction to it
named the wrong mechanism (*"re-prints once per alternative"*), was right about
the count, and had to be measured again.

### When a claim names a mechanism, read the source that would have to be true

The cheapest check in this file, and the one that caught the most in a single
afternoon. Not skepticism, not a re-run -- **one lookup.** A claim that names a
mechanism is making a checkable statement about code or about the platform, and
that statement is usually one file away.

Four instances in one afternoon, across four sessions:

| claim | the lookup | what it said |
| --- | --- | --- |
| "the mojibake is CP437 arithmetic" | run `chcp` | 850 |
| "suppression is verified" | read the function | five assignments, no read-back |
| "the leak assertion attributes by owner" | `word-pids.mjs:59-61` | pure set difference |
| "the settings are per-process" | quit the writer, then read | they persist |

**In every one the conclusion was already right, which is exactly why nothing
ever failed and why only reading the mechanism could have caught it.** A wrong
mechanism under a right conclusion is invisible to every test by construction --
the observable is identical, so no assertion discriminates. The defect surfaces
only when someone acts on the mechanism rather than the conclusion: hardening a
share mode, shaving a deadline, or trusting a flag to mean what it says.

The corollary is that **a correction inherits the burden of a claim.** Each of
the four above was found while correcting something else, and two of the
corrections were themselves overstated on their first attempt -- a proposed
review-record row that was wrong in a different way than the one it fixed, and a
replacement account of the leak deadline claiming it "outlasts foreign churn"
when the cited file says the tail is unbounded and load-dependent. Correcting a
claim puts you in exactly the state that produced it: confident, and one lookup
short.

### A comment that instructs a caller who does not exist yet

A description of what code does rots *passively*: something changes, the
sentence goes stale, and the next reader meets it alongside the code it
describes. An instruction aimed at a future layer rots **aimed**. It keeps
pointing, and what it points at is a caller nobody has written -- so there is no
behaviour anywhere that can contradict it.

Measured instance, both versions read from the tree rather than from a report.
`structure-map.mjs` shipped in L1 (`4abf952`) carrying, at line 266:

```
// The localized id, kept because an edit that wants to reapply this
// style must use it -- the English name would throw.
```

Every clause of that is a claim about the write path, and the write path did not
exist yet. When it was built, the remedy it prescribes was measured not to work:
the localized id throws as well. The replacement on `main` states what was
measured instead -- assigning a style *by* this id throws, the English
`Heading 1` throws too, the write side names no style at all and uses numeric
`wd*` constants or another paragraph's Style object -- and the field is now
labelled "reported for identification, not for reapplication."

**The half that could rot was the half nothing could test.** The comment makes
two statements, and they are not alike. *"The localized id is kept"* describes
the field the line sits on, and L1's own tests assert exactly that -- `styleId`
carried verbatim rather than constructed. *"An edit that wants to reapply this
style must use it"* addresses a caller that did not exist, and no assertion in
the repo could go red for it, then or later. One comment, one covered clause and
one uncoverable one -- and the uncoverable one is the one that was wrong.

What history shows next is co-location, and only that: the correction landed in
`4da84c2`, the same commit that built the write path, on the line L2 had to
touch to record its own measurement. It does not show *why*, and no repository
artefact can -- there is no record of the comment being routed for review, but
an absent record is not evidence of absence. The narrow reading is the useful
one anyway: **the comment was fixed by someone working on that line, not by
anyone checking comments.** Nothing here routes a comment to the layer it
addresses; on the one instance on record, co-location is what corrected it.

So: **a comment addressed to a future layer should say what was measured and
when, never what to do.** The honest form of the original is *"the English name
throws (measured); the localized id is what the file contains"* -- which would
have been **incomplete** when the write path measured further, rather than
wrong. Incomplete invites the next measurement; wrong redirects it.

This is the discharge rule turned forwards in time. Asserting a property of a
neighbour discharges a claim onto code you are not testing; instructing a future
caller discharges it onto code that does not exist, which is strictly worse,
because the neighbour can at least be read.
