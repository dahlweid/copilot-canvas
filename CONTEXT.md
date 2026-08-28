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
takes `FileShare::Read`, so a document the user has open in Word can still be
read and copied — that case *succeeds*. "The user has it open in Word" is
therefore provably **never** the cause, and a message saying so names the one
condition incapable of producing the error. Worth retrying, because the holder
may let go.
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
thread, no resolve button, no reply affordance, and `in_reply_to` on their id
returns 422. So they sit outside "address every comment, reply, repeat"
entirely, and a PR can show zero unresolved threads while carrying the only
finding that gates its merge. Reply to those with a top-level PR comment.
Three of the best findings on #12 arrived this way, including the one that
turned out to have three sites rather than one.

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

**Today the ladder's top rung is unreachable, so every round is Lite.** Effort
level is a UI-only control on the PR's Reviewers panel; it cannot be set from
the CLI or the REST API. Every review body states the level it ran at — always
read it rather than assuming the requested level took effect.

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

**Lite is a floor.** Effort level is UI-only and cannot be set from the CLI;
every review is labelled with its level, so check rather than assume. A clean
Lite pass means the cheap checks passed, not that the code is right — an
independent deeper pass over #16 found a critical data-loss defect that six Lite
rounds elsewhere never approached.

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

