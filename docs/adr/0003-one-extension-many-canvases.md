# One extension hosting several canvases, not one extension per application

Word, Excel, and PowerPoint ship as a single extension that declares a canvas
per application (`canvases?: Canvas[]` accepts an array), rather than three
independent extensions.

The motivation is a runtime constraint, not tidiness. Extensions may not carry
a `package.json` or `node_modules`, and `install_extension` copies exactly one
folder which must run as committed. There is therefore no way for three
extension folders to share a library: three extensions would mean three
vendored copies of pdf.js, three copies of the render cache, server, and
watcher, three installs for the user, and three sets of drift. Gist sharing
caps at roughly 5 MB and refuses binary files, so the duplication is not merely
wasteful — it approaches a hard limit.

Separate canvas ids within that one extension are kept because the agent picks
a canvas from its description, and "Word document" is a clearer signal than a
generic canvas with a mode parameter. Their input schemas also legitimately
differ: a page number for Word, a slide for PowerPoint, a sheet and range for
Excel.

## Consequences

The shared pdf.js rendering surface is a good fit for Word and PowerPoint,
where pagination is inherent to the format. It is a poor fit for Excel, where
pages are invented by print setup and a PDF cannot show a formula at all. That
does not undermine this decision — because the canvases are already separate,
Excel can adopt a different view without disturbing the others. Only the
packaging is shared, and that is the part the runtime forces us to share.
