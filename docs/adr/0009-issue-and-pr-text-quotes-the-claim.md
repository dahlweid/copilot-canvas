# A reference names what it relies on, not a line number

A **positional coordinate** — `file.ext:NN`, `file.ext:NN-NN`, or a bare `:NN`
naming a line of the file it is written in — is not a reference. It hides the
claim it stands for, and an unrelated edit relocates that claim without touching
it. Two rules follow, one per surface:

- **Issue and pull-request text quotes the claim** it rests on — the line, or
  the claim the line stands for. A coordinate may accompany a quote; it may not
  replace one. This is #160(b).
- **Committed files carry no positional coordinate at all.** Reference code
  indirectly, by name and file: `Set-ParagraphText (word-host.ps1)`. See
  *The rule in the tree* below for the exact form and the two exceptions.

The convention exists because a coordinate **hides the content of the claim it
stands for**. A reader who follows a number lands on whatever is at that offset;
if the number has rotted, they land on unrelated code and read the miss as a
bookkeeping nit — "stale line number", a Low — and stop. A rotted coordinate
therefore *reads as cosmetic and closes the question*, which is worse than an
ordinary stale comment: it can sit on top of a substantive error and conceal it,
because nobody who classified it as cosmetic looked underneath.

## The evidence: #141

This is not hypothetical. Issue #141's note to whoever took it cited
`probe-single-instance.ps1:167-170` for a claim it described as PowerPoint's
`ActiveWindow` throwing without a presentation. Two things were wrong, and the
coordinate was how they lasted — but not in the way the first correction on #141
supposed, and the true sequence is the sharper argument.

The coordinate **resolved correctly when it was written.** On `main`'s tree at
the moment #141 was filed (`fed5270`), `:167-170` covers the cited claim:
`ActiveWindow throws without a document, so taking it would mean adding one, and
S3 measures what bare New-Object does`. What broke it was an **unrelated edit to
the same file**: about seventeen minutes after filing, a merge (#139, `ff31de9`)
inserted a census/`Quit`-gate comment above that passage. The claim itself was
not touched — the four lines are byte-identical before and after — but the
insertion pushed them from `:167-170` down to `:176-179`, leaving `:167-170` on
the inserted census text. A later change (#144, `a8a9b7c`) deleted the line
outright, so at `main` today the range points at unrelated text and `ActiveWindow`
is gone from the file entirely. The claim was relocated by an insertion above it,
without a change to its own text.

That is the point, and this repo already made it once, on the record. ADR 0006
rejected raw paragraph indices for a document address, and the reason it gives is
this one exactly:

> Raw paragraph indices were rejected because any insertion shifts every
> subsequent address.

`:167-170` is a raw index, and an insertion above it shifted the claim it named —
the failure 0006 designed the whole read-then-address scheme to avoid. So
verifying `:167-170` at the instant you cite it is not worthless: it buys
correctness *at that instant*. What it does not buy is any durability past it,
because the next commit to that file can relocate the claim without touching your
issue. That is what makes verification-at-writing an insufficient defence against
relocation and a quote an effective one: a quote cannot be relocated by an edit
elsewhere; it either still matches the file or visibly does not.

The disanalogy with ADR 0006 is the part that must not be skipped, because it is
where the tracker is *worse*. In-tree, 0006's addresses move too — but **not
silently**: an edit "move[s] the revision token, and the next read is forced", so
staleness there is machinery-caught and refusable. Tracker text has no revision
token, and this ADR's own limits below record that no check can read it. So the
in-tree fragility was solved by a mechanism the tracker cannot have. A quote is
not a substitute for that mechanism — it is a weaker, manual mitigation, and the
difference in guarantee is the point. The token acts without a reader: it
*detects* the drift and *refuses* the edit. A quote does neither; it only makes
the drift **apparent** to a reader who looks. What the quote buys, with no token
available, is that the relied-on claim is carried inline, so a mismatch against
the file is self-evident to anyone who checks rather than hidden behind a number
that still resolves.

And underneath the coordinate sat the second defect, the one that actually
mattered: the claim was a **Word** measurement wearing a PowerPoint label. PR
#157 measured the PowerPoint behaviour directly — its `ActiveWindow` returns
`$null` with **no exception caught** (the probe runs under
`$ErrorActionPreference = 'Continue'`, so a non-terminating COM error is not
ruled out; the measured claim is the narrower "none caught"), where the imported
Word result had it throwing. The correction now recorded on #141 — including a
correction to that correction's own account of *when* the range broke — sets this
out. A bare `:167-170` gave a reader nothing to weigh that mislabel against: once
broken, following it lands on census text that reads as an unrelated stale
reference, not as a contradiction of the sentence citing it. This is a property
of the notation, not a charge against any reader, and the ADR does not claim to
know which tree the note's author was looking at — only what the trees show.

A quote would have broken this where the coordinate could not. Writing
`"ActiveWindow throws without a document"` puts the word *document* in front of a
reader of a *PowerPoint* issue, where a bare `:167-170` conceals it. The reader
does not have to resolve a number and cross-check an arm of a probe to notice the
mismatch; the mismatch is on the page. Quoting the claim is the only one of the
two #160 remedies that would have caught #141: both #160(a) gates read
*committed files*, and #141 is tracker text, which neither reaches.

## The rule in the tree

It once said here that positional citations in committed comments and docs were
fine, because `tools/check-citations.mjs` and `tools/check-citation-lines.mjs`
gate them. That exemption was wrong, and the counter-example is in this
repository's own safety-critical code. The rule is now the same on both
surfaces.

**In a committed file, reference code by name, not by position.** Write
`Set-ParagraphText (word-host.ps1)`, or `the Quit(0) gate
(probe-single-instance.ps1)`. Where the thing has no name, say what it is: `the
census helper in _common.ps1`. This covers every tracked file — source,
comments, docs, probes, tests, workflows — and it covers a bare `:NN` into the
file's own body exactly as it covers a cross-file coordinate.

Three exceptions, all narrow, all because the objection above does not apply:

- **A coordinate pinned to a commit.** `structure-map.mjs:266` *as of*
  `4abf952` cannot rot — the SHA fixes the tree the number indexes. This is why
  the evidence section above can name `probe-single-instance.ps1:167-170` at
  `fed5270`: without this exception, this ADR could not state its own evidence.
- **A coordinate inside a verbatim transcript** — the captured output of a probe
  run, in a fenced block. The reasoning is different from the two around it and
  is the reason this exception cannot be dropped: **a recorded run is evidence,
  and evidence is not edited to satisfy a gate.** This repo's standard is that a
  claim about platform behaviour is backed by a probe that was actually run;
  rewriting a coordinate inside that probe's printed output destroys the thing
  the standard rests on, and leaves a transcript that no longer matches what the
  probe would print if re-run.

  The worked case is `spikes/word-icon/FINDINGS.md`, whose arm-3 output reads
  `"icon" appears in [generated\rpc.d.ts:18, generated\session-events.d.ts:13]`.
  Those two coordinates index **Copilot CLI 1.0.80's generated typings** — an
  external artifact, named at its version on the line above — not this tree,
  which contains no `generated/` directory at all.

  It is tempting to fold this into the pin exception, since `1.0.80` fixes the
  artifact as surely as a SHA fixes a tree. Resist it: `check-citation-lines.mjs`
  skips fenced blocks *wholesale* and does not look for a version, so a fence
  holding an unpinned coordinate into this repo is skipped too. The exception
  protects the recording. It is not a claim that a coordinate inside a fence
  still resolves, and the gate's success message says so and counts the lines it
  skipped rather than reporting a clean tree.
- **A coordinate that is the subject under discussion,** rather than a
  reference being followed — this document analysing a rotted citation, and
  `check-citation-lines.mjs` and its unit test carrying rot-shaped fixtures to
  assert on. Naming a string is not pointing at a line.

### Why the tree was never the safe half

`word-host.ps1` argues that Word's autocorrect cannot fire because the host
never types. It is a safety argument, and it rests on four coordinates:

> every character of document text goes in through the single function
> `Set-ParagraphText (:1531)`, whose only write is `(Get-TextRange $para).Text =
> $text (:1532)`. […] the only other `.Text =` assignments in this file are
> `Find.Text (:1250, :1323)`

All four are wrong, measured at `5d81ffc`. `Set-ParagraphText` and its write sit
some 170 lines further down than claimed, and the two `Find.Text` assignments
roughly 175 lines down. What the cited lines actually hold there is `$fromPage =
0`, an `if` on `$a.fromPage`, `name = $item.Name` and `$thrown.data.writable =
$writable` — real, plausible, unrelated lines of the same file. This is #168,
and the exact offsets are deliberately not repeated here: they would rot the way
the ones above them did. Re-derive them by searching for the names.

`check-citation-lines.mjs`, **as it stood when this was written**, exits
successfully on this file, and it is worth being exact about what that does and
does not mean, because the obvious reading is not the true one and the true one
is worse. (This analysis is what retired that version of the tool; the last
section records what replaced it.)

The gate did not examine these four and judge them acceptable. **It never saw
them at all.** The tool does name individual citations when it has a finding —
on failure it prints each offending one with its category and the file it was
written in — so the silence here is not a lenient verdict, it is the absence of
any verdict. Its matcher requires a filename with a recognised extension before
the colon —

> `/(?:[A-Za-z0-9._-]+[/\\])*[A-Za-z0-9._-]+\.(?:mjs|ps1|js|ts|md):\d+(?:-\d+)?/g`

— and these four are **bare** self-references, `(:1531)`, with no filename in
front. Run that matcher over the comment and it returns no matches. Green here
is silence, not a verdict.

That is the first failure, and it compounds: a bare self-reference is
simultaneously the **most** fragile form, because it rots on any insertion above
it in its own file and that is the most frequent edit a file receives, and the
**least** checkable form, because a matcher needs a filename to recognise a
citation at all. The weakest notation is the one nothing can watch.

The second failure is that qualifying them would not have helped. Written as
`word-host.ps1:1531`, all four would be matched — and all four would still pass,
because the file has 2300 lines and every one of 1531, 1532, 1250 and 1323 names
a line that exists, in a file that exists, in range. The gate's own header says
why it cannot do better:

> A coordinate that still resolves to real code while no longer meaning anything
> is invisible to any gate.

So the exemption rested on a gate that misses this twice over: it cannot see the
form actually used, and it would pass the qualified form too. It is structurally
unable to catch the failure that matters, and here it is failing to catch it
underneath a safety claim. That is #141's shape, in the tree, on shipped code.

## Two limits, stated so this is not oversold

Both are load-bearing, and a version of this convention that hid either would be
worse than none.

**Only the in-tree half is enforced.** No CI check reads issue or PR text — no
check in this repo can, because those bodies live on GitHub, not in the tree.
That half stays convention.

The in-tree half is different in kind, because a ban on a *syntax* is decidable
where a check on whether a coordinate still *means* what it claimed is not, and
it is now gated: `check-citation-lines.mjs` was inverted from *validate* to
*reject*. It no longer asks whether a coordinate resolves — the analysis above is
why that question was worth nothing — it asks whether one is present, and fails
the build when one is. The three exceptions above are the gate's three
allowances: a coordinate followed *immediately* by an "as of `<sha>`" pin passes
and is counted, fenced transcripts are skipped, and a short exempt list covers
the files that carry rot-shaped fixtures as their subject matter — the checker,
its unit test, and this document.

Read a green run as the gate itself states it, which is narrower than "this tree
complies". It certifies that no line **it read** carries a coordinate, and it
prints what it did not read: the fenced lines, the exempt files, the vendored
and binary ones. That wording is not caution for its own sake. The gate shipped
saying "no positional coordinate in any tracked file" while two sat unread in
`spikes/word-icon/FINDINGS.md`, and a review caught it — an overstatement is
worse than a narrow claim precisely because it is what stops the next person
looking.

**Dropping the number does not make a reference correct.** A name survives
insertion, which a coordinate does not, and `check-citations.mjs` still guards
a `probe-*` filename against dangling. But a renamed function or a moved file
breaks a name too. The claim is narrow and matches what was measured: the rot in
this tree is overwhelmingly *relocation by an edit elsewhere*, and a name is
immune to exactly that — not to everything.