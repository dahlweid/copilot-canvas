// Probe: what accessibility information does bundled pdf.js actually expose,
// and how does that compare with the host's native PDF plugin?
//
// ADR 0004 chose pdf.js over the native viewer. An early draft of that ADR
// claimed accessibility parity; the claim was withdrawn because nothing had been
// tested. This probe exists to replace the assumption with a measurement -- and,
// as it turns out, to establish that one half of the comparison cannot be made
// at all.
//
// ## What can be measured, and what cannot
//
// Three things are measurable and are measured here:
//
//   1. Is the PDF Word produced actually tagged? (`/MarkInfo << /Marked true >>`
//      and a `/StructTreeRoot`.) Read straight out of the file, so this half
//      needs neither pdf.js nor a browser.
//   2. Does pdf.js surface that structure through `page.getStructTree()`, with
//      roles that correspond to the document's headings?
//   3. Does the text layer's DOM order match the document's reading order?
//
// One thing is **not** measurable, and this is the finding rather than a gap in
// the probe: the native plugin renders in its own process and exposes no DOM.
// That is the very fact ADR 0004 rests on. There is no way to enumerate its
// accessibility tree from the embedding page, so a like-for-like DOM comparison
// against pdf.js is structurally impossible -- not merely unimplemented. Any
// "parity" claim would have to come from driving a real screen reader against
// both, which is not something this repo can do in a probe.
//
// So the honest output is an absolute statement of what pdf.js provides, plus a
// named gap: assistive-technology output is not measured by anyone here.
//
// ## Usage
//
//   node spikes/pdfjs/probes/probe-accessibility.mjs <exported.pdf>
//
// Part 1 prints immediately. Parts 2 and 3 need a browser, so the probe then
// serves a page -- open the printed URL in the app's webview (the same harness
// spikes/pdfjs/probes/probe-range-requests.mjs uses) and the results are printed
// back here.

import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const VENDOR = path.resolve(HERE, "..", "..", "..", ".github", "extensions", "office-canvas", "src", "ui", "vendor");

const pdfPath = process.argv[2];
if (!pdfPath) {
    console.error("usage: node probe-accessibility.mjs <exported.pdf>");
    process.exit(2);
}
const pdf = await readFile(pdfPath);

// --- part 1: is the file tagged? --------------------------------------------
//
// Read as bytes rather than through pdf.js on purpose. pdf.js reporting a
// structure tree would not distinguish "Word tagged the file" from "pdf.js
// synthesised something", and it is the export setting that is in question.

const latin1 = pdf.toString("latin1");
const marks = {
    "/MarkInfo": latin1.includes("/MarkInfo"),
    "/Marked true": /\/Marked\s+true/.test(latin1),
    "/StructTreeRoot": latin1.includes("/StructTreeRoot"),
    "/Lang": /\/Lang\s*\(/.test(latin1),
};

// Whether the raw scan can see structure elements at all.
//
// Measured, and it is the reason no role table is printed here: Word writes the
// structure tree into compressed object streams (`/ObjStm`), so `/StructElem`
// and the role names are inside a Flate-compressed blob and simply are not
// present as bytes in the file. A first version of this probe counted role
// tokens with a regex over the raw bytes and reported `{"/L": 1}` for a document
// with five pages of headings, tables and lists -- not a finding about the
// document but an artefact of scanning compressed data, with the single hit a
// coincidental byte sequence. Absence here means nothing whatever, so the roles
// are read through pdf.js in part 2 instead, where they have been decompressed.
const objectStreams = (latin1.match(/\/Type\s*\/ObjStm/g) ?? []).length;

console.log(`file: ${pdfPath} (${pdf.length} bytes)`);
console.log("\n--- part 1: tagging, read from the file itself ---");
for (const [key, value] of Object.entries(marks)) console.log(`  ${key.padEnd(16)} ${value}`);
console.log(`  compressed object streams: ${objectStreams}`);
console.log(
    objectStreams > 0
        ? "  -> structure elements are inside those streams; a raw scan cannot enumerate roles.\n" +
              "     Part 2 reads them through pdf.js, which decompresses them."
        : "  -> no object streams; structure elements would be visible as raw bytes.",
);
if (!marks["/StructTreeRoot"]) {
    console.log("  NOTE: untagged. Parts 2 and 3 will measure what pdf.js does with an untagged file,");
    console.log("        which is a different question from the one ADR 0004 asks.");
}

// --- parts 2 and 3: what pdf.js exposes -------------------------------------

const main = await readFile(path.join(VENDOR, "pdf.min.mjs"));
const workerParts = JSON.parse(await readFile(path.join(VENDOR, "pdfjs.manifest.json"), "utf8")).worker.parts;
const worker = Buffer.concat(await Promise.all(workerParts.map((part) => readFile(path.join(VENDOR, part.name)))));
const textLayerCss = await readFile(path.join(VENDOR, "pdf-text-layer.css"));

const js = { "content-type": "text/javascript; charset=utf-8" };

const PAGE = `<!doctype html><meta charset="utf-8"><title>accessibility probe</title>
<link rel="stylesheet" href="/vendor/pdf-text-layer.css">
<style>.page{position:relative;--scale-round-x:1px;--scale-round-y:1px;margin:8px}</style>
<body><div id="out">running…</div><div class="page" id="page"></div>
<script type="module">
const out = document.getElementById("out");
const report = async (data) => {
    out.textContent = JSON.stringify(data, null, 2);
    await fetch("/report", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(data) });
};
try {
    const pdfjs = await import("/vendor/pdf.min.mjs");
    pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.mjs";
    const doc = await pdfjs.getDocument({ url: "/doc.pdf" }).promise;
    const page = await doc.getPage(1);

    // Part 2: the structure tree, as pdf.js hands it over.
    let structTree = null;
    let structError = null;
    try {
        structTree = await page.getStructTree();
    } catch (e) {
        structError = String(e && e.message || e);
    }
    const roles = [];
    let contentRefs = 0;
    const walk = (node, depth) => {
        if (!node) return;
        if (node.role) roles.push({ depth, role: node.role });
        // Leaves are \`{ type: "content", id: "p0_mc0" }\` -- the marked-content
        // ids that tie a structure node to a run of text on the page. They carry
        // no \`role\`, so a walk that only records roled nodes never sees them and
        // would report "no refs" about a tree that is full of them.
        if (node.type === "content" && node.id) contentRefs += 1;
        for (const child of node.children ?? []) walk(child, depth + 1);
    };
    walk(structTree, 0);

    // Part 3: the text layer's DOM order, against the extraction order.
    const viewport = page.getViewport({ scale: 1 });
    const host = document.getElementById("page");
    host.style.width = viewport.width + "px";
    host.style.height = viewport.height + "px";
    const layer = document.createElement("div");
    layer.className = "textLayer";
    host.append(layer);
    const textLayer = new pdfjs.TextLayer({ textContentSource: await page.getTextContent(), container: layer, viewport });
    await textLayer.render();

    const spans = [...layer.querySelectorAll("span")];
    const domOrder = spans.map((s) => s.textContent).filter((t) => t.trim());
    const contentOrder = (await page.getTextContent()).items.map((i) => i.str).filter((t) => t.trim());

    // Does the layer carry anything an assistive technology could use beyond the
    // text itself? Measured, not assumed: role/aria attributes on the spans.
    const ariaAttrs = new Set();
    let spansWithIds = 0;
    for (const span of spans) {
        for (const attr of span.getAttributeNames()) {
            if (attr === "role" || attr.startsWith("aria-")) ariaAttrs.add(attr);
        }
        if (span.id) spansWithIds += 1;
    }

    await report({
        ok: true,
        pdfjsVersion: pdfjs.version,
        pages: doc.numPages,
        structTreePresent: Boolean(structTree),
        structError,
        roleCount: roles.length,
        rolesTop: roles.slice(0, 40),
        distinctRoles: [...new Set(roles.map((r) => r.role))],
        contentRefs,
        // Whether the spans we mount can be tied back to those refs. pdf.js's own
        // viewer wires this up separately; the question is what *our* text layer
        // does, which is a different thing from what the structure tree contains.
        spansWithMarkedContentIds: spansWithIds,
        textLayerSpans: spans.length,
        readingOrderMatches: JSON.stringify(domOrder) === JSON.stringify(contentOrder),
        domOrderFirst: domOrder.slice(0, 8),
        contentOrderFirst: contentOrder.slice(0, 8),
        textLayerAriaAttributes: [...ariaAttrs],
        // The half that cannot be measured from here, recorded so the report
        // says so rather than staying silent about it.
        nativePluginComparable: false,
        nativePluginNote: "the native plugin exposes no DOM to the embedding page; see ADR 0004",
    });
} catch (e) {
    await report({ ok: false, error: String(e && e.message || e) });
}
</script></body>`;

const server = createServer(async (req, res) => {
    const url = req.url.split("?")[0];
    if (url === "/vendor/pdf.worker.min.mjs") {
        res.writeHead(200, js);
        return res.end(worker);
    }
    if (url === "/vendor/pdf.min.mjs") {
        res.writeHead(200, js);
        return res.end(main);
    }
    if (url === "/vendor/pdf-text-layer.css") {
        res.writeHead(200, { "content-type": "text/css; charset=utf-8" });
        return res.end(textLayerCss);
    }
    if (url === "/doc.pdf") {
        res.writeHead(200, { "content-type": "application/pdf", "content-length": pdf.length, "accept-ranges": "bytes" });
        return res.end(pdf);
    }
    if (url === "/report" && req.method === "POST") {
        const chunks = [];
        for await (const chunk of req) chunks.push(chunk);
        console.log("\n--- parts 2 and 3: what pdf.js exposes ---");
        console.log(Buffer.concat(chunks).toString("utf8"));
        console.log("\nDone. Ctrl-C to stop.");
        res.writeHead(204);
        return res.end();
    }
    if (url === "/") {
        res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
        return res.end(PAGE);
    }
    res.writeHead(404);
    res.end("nf");
});

await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
console.log(`\nOpen this in the app's webview: http://127.0.0.1:${server.address().port}/`);
