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
