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
`probe-single-instance.ps1:167-170` for a claim about PowerPoint's
`ActiveWindow`. Two things were wrong at once, and the coordinate hid both. The
range was wrong **when it was written** — `:167-170` held WINWORD-census text,
not the claim — and underneath it the claim itself was a **Word** measurement
wearing a PowerPoint label. That premise survived filing, implementation, and
three review rounds. Round 2 spotted the citation, classified it as a rotted
reference, and stopped there.

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
