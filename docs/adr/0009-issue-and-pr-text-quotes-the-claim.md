# Issue and pull-request text quotes the claim it relies on, not a coordinate

When issue or pull-request text rests on something a file says, it **quotes that
thing** — the line, or the claim the line stands for — rather than pointing at a
bare `file.ext:NN` coordinate. A coordinate may accompany a quote; it may not
replace one. This is #160(b).

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
the moment #141 was filed (`fed5270`), `:167-170` is exactly the passage cited:
`ActiveWindow throws without a document, so taking it would mean adding one, and
S3 measures what bare New-Object does`. What broke it was an **unrelated edit to
the same file**: about seventeen minutes after filing, a merge (#139, `ff31de9`)
inserted a census/`Quit`-gate comment above that passage, pushing the claim down
to `:176-177` and leaving `:167-170` pointing at the inserted census text. A
later change (#144, `a8a9b7c`) deleted the line outright, so at `main` today the
range points at unrelated text and `ActiveWindow` is gone from the file
entirely. The coordinate was checked, was right, and was broken before any reader
of the issue reached it — by someone editing a part of the file that had nothing
to do with the claim.

That is the point, and it is the same one ADR 0006 makes about addresses: a
coordinate is a position, and positions move under edits you do not control.
Verifying `:167-170` at the instant you cite it buys nothing, because the next
commit to that file can relocate the claim without touching your issue. A quote
cannot be relocated by an edit elsewhere; it either still matches the file or
visibly does not.

And underneath the coordinate sat the second defect, the one that actually
mattered: the claim was a **Word** measurement wearing a PowerPoint label —
verified in PR #157, which measured PowerPoint's `ActiveWindow` returning `$null`
rather than throwing. The correction now recorded on #141 (including a correction
to that correction's own account of *when* the range broke) sets this out. A bare
`:167-170` gave a reader nothing to weigh that mislabel against: following it,
once broken, lands on census text that reads as an unrelated stale reference, not
as a contradiction of the sentence citing it. This is a property of the notation,
not a charge against any reader — and the ADR does not claim to know which tree
the note's author was looking at, only what the trees show.

A quote would have broken this where the coordinate could not. Writing
`"ActiveWindow throws without a document"` puts the word *document* in front of a
reader of a *PowerPoint* issue, where a bare `:167-170` conceals it. The reader
does not have to resolve a number and cross-check an arm of a probe to notice the
mismatch; the mismatch is on the page. Quoting the claim is the only one of the
two #160 remedies that would have caught #141 — see the limits below.

## What this does and does not forbid

It targets **issue and pull-request text**, the two surfaces where nothing in
this repo gates a coordinate.

It does **not** forbid positional citations in committed code comments and docs.
Those are legitimate, they are this repo's own style, and they are gated by
#160(a): `tools/check-citations.mjs` guards `probe-*` filenames, and
`tools/check-citation-lines.mjs` guards positional `:NN` coordinates for the
failures a machine can decide without knowing what the citation meant. A rule
that banned coordinates everywhere would contradict the tooling this repo just
added. Where a coordinate genuinely helps in issue text, the convention is that
it *accompanies* a quote, never stands in for one.

## Two limits, stated so this is not oversold

Both are load-bearing, and a version of this convention that hid either would be
worse than none.

**Nothing enforces this. It is a convention, not a gate.** No CI check reads
issue-tracker text — no check in this repo can, because issue and PR bodies live
on GitHub, not in the tree. Do not read this record, or the rule in
`.github/copilot-instructions.md`, as something CI verifies. It changes how a
writer writes; it is not mechanically checked.

**The #160(a) gates would not have caught #141, and this convention is the only
part of #160 that would have.** `check-citations.mjs` gates *committed files* and
is *filename-only*; `check-citation-lines.mjs` gates *committed* positional
coordinates. #141 is issue-tracker text, which neither reaches. The two gates
catch the in-tree class — the rotted `CONTEXT.md` coordinates #160 was filed
over — and nothing on the tracker. Conversely this convention reaches the
tracker and is unenforced. Each part covers exactly what the other cannot, and
neither covers everything.
