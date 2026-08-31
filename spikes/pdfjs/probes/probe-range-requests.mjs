// Probe: what Range requests does pdf.js actually issue, and does the
// repo's parseByteRange (server.mjs, PR #5) handle every form?
//
// Issue #9 claims L4 is parseByteRange's "first real consumer". This checks
// that claim against a live pdf.js instead of assuming it.
import { readFile } from "node:fs/promises";
import { createServer } from "node:http";
import { join } from "node:path";
import { pathToFileURL } from "node:url";

const DIR = join(import.meta.dirname, "worker-split");
const REPO_SERVER = join(
  import.meta.dirname,
  "../../..",
  ".github/extensions/office-canvas/src/server.mjs",
);

// No try/catch. The second of this probe's two questions is answered entirely
// by this import, so a recovery here would print `parseByteRange -> n/a` for
// every range, exit 0, and read as a measured absence rather than as a probe
// that never ran. Let it throw: an unhandled rejection here exits non-zero.
const { parseByteRange } = await import(pathToFileURL(REPO_SERVER).href);
if (typeof parseByteRange !== "function") {
  throw new Error(`${REPO_SERVER} exports no parseByteRange function`);
}
console.log(`parseByteRange imported from ${REPO_SERVER}`);

const worker = await readFile(join(DIR, "pdf.worker.min.mjs"));
const main = await readFile(join(DIR, "pdf.min.mjs"));
const pdf = await readFile(process.argv[2]);

const seen = [];

const server = createServer(async (req, res) => {
  const url = req.url.split("?")[0];
  if (url === "/vendor/pdf.worker.min.mjs") {
    res.writeHead(200, { "content-type": "text/javascript; charset=utf-8" });
    return res.end(worker);
  }
  if (url === "/vendor/pdf.min.mjs") {
    res.writeHead(200, { "content-type": "text/javascript; charset=utf-8" });
    return res.end(main);
  }
  if (url === "/doc.pdf") {
    const range = req.headers.range;
    seen.push({ method: req.method, range: range ?? null });
    if (req.method === "HEAD") {
      res.writeHead(200, {
        "content-type": "application/pdf",
        "content-length": pdf.length,
        "accept-ranges": "bytes",
      });
      return res.end();
    }
    if (range) {
      // deliberately naive parse here; the point is to RECORD what arrives
      const m = /^bytes=(\d*)-(\d*)$/.exec(range.trim());
      if (m) {
        let start, end;
        if (m[1] === "") {
          const suffix = Number(m[2]);
          start = Math.max(0, pdf.length - suffix);
          end = pdf.length - 1;
        } else {
          start = Number(m[1]);
          end = m[2] === "" ? pdf.length - 1 : Number(m[2]);
        }
        const slice = pdf.subarray(start, end + 1);
        res.writeHead(206, {
          "content-type": "application/pdf",
          "content-range": `bytes ${start}-${end}/${pdf.length}`,
          "content-length": slice.length,
          "accept-ranges": "bytes",
        });
        return res.end(slice);
      }
    }
    res.writeHead(200, {
      "content-type": "application/pdf",
      "content-length": pdf.length,
      "accept-ranges": "bytes",
    });
    return res.end(pdf);
  }
  if (url === "/report") {
    res.writeHead(200, { "content-type": "application/json" });
    return res.end(JSON.stringify(seen));
  }
  if (url === "/") {
    const html = `<!doctype html><meta charset="utf-8"><title>range probe</title>
<body><div id="out">loading</div>
<script type="module">
const out = document.getElementById("out");
try {
  const pdfjs = await import("/vendor/pdf.min.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.mjs";
  // force the ranged path: small chunks so pdf.js must ask for several
  const doc = await pdfjs.getDocument({
    url: "/doc.pdf",
    disableRange: false,
    disableStream: true,
    disableAutoFetch: true,
    rangeChunkSize: 8192,
  }).promise;
  const n = doc.numPages;
  const last = await doc.getPage(n);
  const tc = await last.getTextContent();
  out.textContent = "PROBE_DONE " + JSON.stringify({ ok:true, pages:n, lastPageItems: tc.items.length });
} catch (e) {
  out.textContent = "PROBE_DONE " + JSON.stringify({ ok:false, error:String(e&&e.message||e) });
}
</script></body>`;
    res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    return res.end(html);
  }
  res.writeHead(404);
  res.end("nf");
});

await new Promise((r) => server.listen(0, "127.0.0.1", r));
console.log(`SERVER_PORT=${server.address().port}`);
console.log(`pdf: ${pdf.length} bytes`);

process.on("SIGINT", () => process.exit(0));
setInterval(() => {
  if (seen.length) {
    const ranges = seen.filter((s) => s.range).map((s) => s.range);
    const uniq = [...new Set(ranges)];
    console.log(`\n--- ${seen.length} requests, ${uniq.length} distinct range forms ---`);
    for (const r of uniq.slice(0, 12)) {
      let parsed;
      try {
        parsed = JSON.stringify(parseByteRange(r, pdf.length));
      } catch (e) {
        parsed = "THREW: " + e.message;
      }
      console.log(`  ${r.padEnd(28)} parseByteRange -> ${parsed}`);
    }
    const suffix = uniq.filter((r) => /^bytes=-\d+$/.test(r));
    console.log(`suffix-form ranges (bytes=-N): ${suffix.length ? suffix.join(", ") : "NONE"}`);
    seen.length = 0;
  }
}, 3000);
