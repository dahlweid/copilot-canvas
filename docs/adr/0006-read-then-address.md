# The agent addresses the document by reading it first, under a revision token

The agent does not describe edit locations in prose or by raw paragraph number.
A read tool returns a **structure map** — the document's paragraphs with an ID
minted for that read — together with a **revision token** derived from the file.
Edits cite an ID and present the token. If the file changed underneath, the
token no longer matches and the edit is refused with an instruction to re-read.

This is optimistic concurrency, and it is forced by the write model. Because we
hold the document open only for the length of one operation (ADR 0005), we have
no lock spanning read and write and therefore cannot assume anything stayed
where we found it. A script may have regenerated the file; the user may have
typed in their own Word. The token turns that from a silent corruption into a
detectable, refusable condition.

Raw paragraph indices were rejected because any insertion shifts every
subsequent address. Injected bookmarks or content controls were rejected
because they modify the user's document to serve our bookkeeping.

## Consequences

**Structural reads must use `Content.WordOpenXML`, parsed outside Word.**
Walking `Document.Paragraphs` and touching properties per paragraph measured
3724 ms for 219 paragraphs, because each property is a cross-process call and
the cost scales with length. One `WordOpenXML` call plus local parsing measured
289 ms for the same document and returns more: text, style, and full markup.
A per-paragraph property walk is a defect, not a slow path.

IDs are derived from **heading path, paragraph text, and an occurrence index**,
because Word exposes no stable paragraph identity. In `demo.docx` text alone was
already unique across 219 paragraphs, so the occurrence index did nothing — but
documents with repeated cells, boilerplate, or many empty paragraphs will
collide, so it stays. A fixture with verbatim-repeated paragraphs and repeated
empty paragraphs now proves it disambiguates; the counter is scoped to the
heading path, so identical text under two different headings is already distinct
at occurrence 1.

Reads happen against **an unblocked temp copy, not the original**. This is
stronger than closing the document promptly: the original is never held at all,
so a script may regenerate it and the canvas may auto-refresh while a read is in
flight. It also sidesteps mark-of-the-web, whose only other fix — `Unblock-File`
on the original — would modify the user's document in order to read it. The cost
is that the copy and the file could diverge, so the revision token is taken on
both sides of the read and a divergence is retried once and then refused rather
than papered over.

Style IDs in the markup are **localized, and mangled on top of that**. The first
heading here is not `Heading1` — but it is not `Überschrift1` either. Word mints
the ID from the localized style name ("Überschrift 1") and drops the characters
outside ASCII on the way, so what is actually in the file is **`berschrift1`**
(verified by unzipping a Word-authored `.docx` and reading `word/styles.xml`).
A style ID is therefore carried, never matched: resolution goes through the
styles part, whose `w:name` holds the canonical built-in name ("heading 1") in
any UI language, and falls back to `w:outlineLvl` for genuinely custom styles.
Matching English style names is the same trap that already breaks
`Selection.Style = "Heading 1"` on this machine; matching the *obvious* German
spelling would have failed just as silently.

An edit response returns a fresh token, so the agent is not forced to re-read
after its own edits — only after someone else's. A save with nothing dirty does
not change the token, so it does not churn on inspection.

**An address is a coordinate, not a handle, and two mutations move it.**
Derived addressing buys stability against the things that churn for no reason —
re-saving, run splitting, per-session `w:rsid*` stamps, and paragraphs moving
within their section — and the same content authored in two separate Word runs
yields byte-different files but *identical* addresses. It cannot survive changes
to what the address is made of. Two cases, both proven by test:

- Deleting one member of a duplicate-text group **renumbers its successors**,
  because the occurrence index counts from the top of the heading path.
- Renaming a heading **moves every address beneath it**, plus the heading's own.

Neither is silent: both are edits, so both move the revision token, and the next
read is forced. The rule this imposes on callers is that an address may be used
within one read-then-edit cycle and must not be cached across an edit.
