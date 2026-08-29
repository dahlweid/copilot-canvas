// Exports a .docx to PDF through the extension's own render pipeline, so a
// probe measures the artefact we actually ship rather than one produced another
// way. Writes the PDF beside the document and prints its path.
//
//   node spikes/pdfjs/probes/export-fixture-pdf.mjs <document.docx> <out.pdf>
//
// Needs an installed, licensed Word.

import { copyFile, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const SRC = path.resolve(HERE, "..", "..", "..", ".github", "extensions", "office-canvas", "src");
const { RenderCache } = await import(pathToFileURL(path.join(SRC, "render-cache.mjs")).href);

const [docPath, outPath] = process.argv.slice(2);
if (!docPath || !outPath) {
    console.error("usage: node export-fixture-pdf.mjs <document.docx> <out.pdf>");
    process.exit(2);
}

const cacheRoot = await mkdtemp(path.join(tmpdir(), "export-fixture-"));
const cache = new RenderCache({ cacheRoot, log: (m) => console.error(`[cache] ${m}`) });
try {
    const info = await cache.open(path.resolve(docPath));
    const pdf = await cache.pdf(path.resolve(docPath));
    await copyFile(pdf.file ?? pdf, path.resolve(outPath));
    console.log(`pages=${info.pageCount} -> ${path.resolve(outPath)}`);
} finally {
    // `dispose()`, not `shutdown()`. An optional call to a method that does not
    // exist is a silent no-op: the first version of this script wrote
    // `cache.shutdown?.()`, which left Word holding the working copy and turned
    // the cleanup into an EBUSY -- with a leaked WINWORD.EXE joining the
    // population every other session is already contending with.
    await cache.dispose();
    await rm(cacheRoot, { recursive: true, force: true });
}
