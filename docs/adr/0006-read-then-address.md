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

### Failures are typed, because the caller has to branch on them

Every host failure originally collapsed to a single `word_error`. That is
adequate for one consumer that only reports the message and hopeless for a
consumer that must decide what to do: "the file is locked", "the file is gone",
"the document is corrupt" and "the disk is full" all demand different
responses, and telling them apart meant matching `$_.Exception.Message` — which
is **localized**, so the matching works on an English machine and silently stops
working on this German one.

The codes a caller may branch on:

| code                        | meaning                                                    |
| --------------------------- | ---------------------------------------------------------- |
| `file_not_found`            | nothing at that path                                        |
| `file_locked`               | another process holds it *more strictly than Word does*     |
| `permission_denied`         | we are not allowed to read it; no other process involved    |
| `document_unreadable`       | Word opened it and could not make sense of it               |
| `copy_failed`               | the working copy could not be made                          |
| `write_failed`              | the output could not be written                             |
| `no_such_document`          | the document id is not registered (host restarted)          |
| `word_unavailable`          | Word could not be started or died                           |
| `word_timeout`              | an operation exceeded its budget                            |
| `page_out_of_range`         | a page was asked for beyond the end                         |
| `invalid_request`           | the arguments do not describe an operation                  |
| `document_changed_during_read` | the file moved under the read; the map would mis-address |

`file_locked` deserves care. Word opens a document with `FileShare::Read`, so a
document **open in Word can still be read and copied** — which is the only
reason a copy-based read works against a document the user is looking at. So
`file_locked` does not mean "open in Word"; it means a stricter holder, and a
message that says "close it in Word" would usually be wrong.

`file_locked` and `permission_denied` are separated on measurement, not on
intuition, and they are separated because they need different remediation — a
lock may clear on its own and is worth retrying, a permission will not and is
not. What each cause actually produces:

| cause                              | `Copy-Item` (host)                  | Node stream |
| ---------------------------------- | ----------------------------------- | ----------- |
| `FileShare::None` (exclusive lock) | `System.IO.IOException`             | `EBUSY`     |
| ACL denying read                   | `System.UnauthorizedAccessException`| `EPERM`     |
| `FileShare::Read` (Word's own)     | **succeeds**                        | **succeeds**|
| read-only attribute                | **succeeds**                        | **succeeds**|

The bottom two rows are the ones worth remembering. A document open in Word is
readable, and the read-only attribute does not block reading at all — so
neither reaches these codes, and neither is a failure.

The host branches on the exception **type**, never the message. Messages are
localized, so message-matching works on an English machine and silently stops
working on this German one; that trap has already cost this project twice.
Note also that `writable = $false` on its own does not imply a lock:
`Test-FileWritable` takes a *write* handle, which the read-only attribute and a
write-denying ACL also refuse.

Two shapes follow from this and are worth stating because both were originally
backwards:

- **`writable` is reported on failure, not only on success.** The case where the
  caller most needs to know the original is locked is the case where the
  operation failed.
- **Filesystem errnos are translated at the boundary.** The revision token is
  the first thing a read touches on the original — before Word is started at
  all — so an exclusively held file produced a raw `EBUSY` that escaped the
  entire typed vocabulary. An untyped errno is a missing code, not a detail.
