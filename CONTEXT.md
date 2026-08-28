# Copilot Canvas — Office documents

Agent-callable tools that read, author, and edit Word, Excel, and PowerPoint
documents by driving the installed Office application through COM, plus a
side-panel canvas that displays the result page-accurately.

The tools are the product. The canvas is how the user watches the agent work.

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
The document file at the path the user named. The only file the user cares
about.
_Avoid_: source, target, real file

**Working copy**:
The extension's private duplicate of the original, opened by the hidden
instance. Agent edits land here.
_Avoid_: temp file, scratch copy

**Apply**:
Promoting the working copy back onto the original. The only operation that
modifies the user's file.
_Avoid_: save, commit, sync

**Hidden instance**:
The extension's own Office process, `Visible = false`, which no human ever
interacts with. One per application.
_Avoid_: background Word, server instance

**User instance**:
The user's own visible Office application, opened by "Edit in Word". Outside
our control.

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
