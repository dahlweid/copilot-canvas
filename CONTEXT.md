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
no script can regenerate the document and no other Word can open it.

**Hidden instance**:
The extension's own Office process, `Visible = false`, which no human ever
interacts with. One per application.
_Avoid_: background Word, server instance

**User instance**:
The user's own visible Office application, opened by "Edit in Word". Outside
our control, and a second writer we must detect rather than collide with.

**Handoff**:
Transfer of write authority between the agent and the user instance. Exactly
one of them holds it at any moment.

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
stable paragraph identity of its own.
_Avoid_: selector, locator, range

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
