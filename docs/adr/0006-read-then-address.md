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
collide, so it stays.

Style IDs in the markup are **localized** — the first heading here is
`Überschrift1`, not `Heading1` — so style resolution must go through the styles
part or fall back to `w:outlineLvl`. Matching English style names is the same
trap that already breaks `Selection.Style = "Heading 1"` on this machine.

An edit response returns a fresh token, so the agent is not forced to re-read
after its own edits — only after someone else's. A save with nothing dirty does
not change the token, so it does not churn on inspection.
