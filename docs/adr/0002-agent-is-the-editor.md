# The agent is the editor; the canvas displays and never edits

The document operations are exposed as agent-callable tools in `tools[]`. The
canvas is a display surface with no editing affordances: the user reads it, and
changes it by asking Copilot. This is deliberate, and it is the opposite of what
a reader will expect from something that looks like a document viewer.

The reason is that it removes an entire category of machinery rather than
merely deferring it. A human editing in the canvas would need coordinate
mapping from click to document position, a version-consistency protocol between
the canvas and the document, a typing overlay to hide round-trip latency, and
caret rendering. An agent editing through COM needs none of these, and the
168 ms render budget stops being a latency constraint because nobody is waiting
on a keystroke.

It also makes the tool useful in a way a viewer is not: authoring and editing
`.docx` through Word's own object model removes the usual detour through
third-party document libraries, so the output is what Word itself would
produce.

## Consequences

Tools, not canvas actions, are the editing surface — and this is forced, not
stylistic. A canvas action handler returns `unknown` and has no channel for
binary content, whereas a tool result carries `binaryResultsForLlm` as
`ToolBinaryResult[]`. We verified by probe that an image returned this way
genuinely reaches the model: a tool returned four randomly-chosen quadrant
colours that were deliberately withheld from the text result, and the model
named all four correctly. That is what makes the agent's verification loop —
edit, re-render, look at the page — possible at all. Had we built editing as
canvas actions, the agent could never have seen its own work.

The user still needs a way to edit by hand, so the canvas offers "Edit in
Word", which opens the document in their own visible Office application. That
introduces a second writer and therefore requires a handoff rule; see the write
model ADR.
