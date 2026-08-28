// Probe: can a >1,000,000-byte pdf.js worker be shipped as UTF-8-valid parts
// and reassembled by a server route, such that pdf.js still works?
//
// This validates the design handed to issue #9 before any code is written.
// Run: node probe-worker-split.mjs
import { createWriteStream } from "node:fs";
import { mkdir, readFile, writeFile, stat } from "node:fs/promises";
import { createHash } from "node:crypto";
import { join } from "node:path";
import { createServer } from "node:http";

const VER = "6.2.108";
const CDN = `https://cdn.jsdelivr.net/npm/pdfjs-dist@${VER}/build`;
const DIR = join(import.meta.dirname, "worker-split");
const PER_FILE_CAP = 1_000_000; // decimal, measured
const TOTAL_CAP = 5_000_000;

const sha = (b) => createHash("sha256").update(b).digest("hex");

async function fetchTo(url, dest) {
  try {
    await stat(dest);
    return await readFile(dest); // cached
  } catch {}
  const res = await fetch(url);
  if (!res.ok) throw new Error(`${url} -> ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  await writeFile(dest, buf);
  return buf;
}

// Split into <=cap chunks so every part is independently valid UTF-8.
// The real constraint is NOT "boundary byte < 0x80" — it is that the boundary
// must not fall INSIDE a multi-byte sequence, i.e. the byte at the split point
// must not be a continuation byte (0x80-0xBF). Prefer a newline when one is
// reachable, so the parts stay diffable.
function splitSafe(buf, cap) {
  const parts = [];
  let start = 0;
  while (start < buf.length) {
    let end = Math.min(start + cap, buf.length);
    if (end < buf.length) {
      const nl = buf.lastIndexOf(0x0a, end - 1);
      if (nl > start && end - nl < 65536) {
        end = nl + 1; // split just after a newline
      } else {
        while (end > start && (buf[end] & 0xc0) === 0x80) end--;
        if (end === start) throw new Error("no safe boundary found");
      }
    }
    parts.push(buf.subarray(start, end));
    start = end;
  }
  return parts;
}

function isValidUtf8(buf) {
  const dec = new TextDecoder("utf-8", { fatal: true });
  try { dec.decode(buf); return true; } catch { return false; }
}

const results = [];
const ok = (name, pass, detail) => {
  results.push({ name, pass, detail });
  console.log(`${pass ? "PASS" : "FAIL"}  ${name}${detail ? " — " + detail : ""}`);
};

await mkdir(DIR, { recursive: true });

console.log("--- fetching pdf.js ---");
const worker = await fetchTo(`${CDN}/pdf.worker.min.mjs`, join(DIR, "pdf.worker.min.mjs"));
const main = await fetchTo(`${CDN}/pdf.min.mjs`, join(DIR, "pdf.min.mjs"));
console.log(`worker ${worker.length} bytes, main ${main.length} bytes`);

ok("worker exceeds the per-file cap (i.e. splitting is genuinely required)",
   worker.length > PER_FILE_CAP, `${worker.length} > ${PER_FILE_CAP}`);

// Aim for parts comfortably under the cap.
const target = 600_000;
const parts = splitSafe(worker, target);
ok("worker splits into >1 part", parts.length > 1, `${parts.length} parts`);
ok("every part is under the per-file cap",
   parts.every((p) => p.length <= PER_FILE_CAP),
   parts.map((p) => p.length).join(" + "));
ok("every part is independently valid UTF-8", parts.every(isValidUtf8));

const rejoined = Buffer.concat(parts);
ok("concatenation is byte-exact", sha(rejoined) === sha(worker),
   `sha256 ${sha(worker).slice(0, 16)}…`);

const total = worker.length + main.length;
ok("total stays under the total cap", total < TOTAL_CAP,
   `${total} < ${TOTAL_CAP}`);

for (let i = 0; i < parts.length; i++) {
  await writeFile(join(DIR, `pdf.worker.min.mjs.part${i}`), parts[i]);
}

// --- serve: parts concatenated behind ONE route, as server.mjs would ---
const pdfPath = process.argv[2];
const pdfBytes = pdfPath ? await readFile(pdfPath) : null;

const server = createServer(async (req, res) => {
  const url = req.url.split("?")[0];
  if (url === "/vendor/pdf.worker.min.mjs") {
    const chunks = [];
    for (let i = 0; i < parts.length; i++) {
      chunks.push(await readFile(join(DIR, `pdf.worker.min.mjs.part${i}`)));
    }
    const body = Buffer.concat(chunks);
    res.writeHead(200, { "content-type": "text/javascript; charset=utf-8", "content-length": body.length });
    return res.end(body);
  }
  if (url === "/vendor/pdf.min.mjs") {
    res.writeHead(200, { "content-type": "text/javascript; charset=utf-8", "content-length": main.length });
    return res.end(main);
  }
  if (url === "/doc.pdf" && pdfBytes) {
    res.writeHead(200, { "content-type": "application/pdf", "content-length": pdfBytes.length });
    return res.end(pdfBytes);
  }
  if (url === "/") {
    const html = `<!doctype html><meta charset="utf-8"><title>probe</title>
<body><div id="out">loading</div><canvas id="c"></canvas>
<script type="module">
const out = document.getElementById("out");
try {
  const pdfjs = await import("/vendor/pdf.min.mjs");
  pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.mjs";
  const doc = await pdfjs.getDocument({ url: "/doc.pdf" }).promise;
  const page = await doc.getPage(1);
  const tc = await page.getTextContent();
  const text = tc.items.map(i => i.str).join("").slice(0, 120);
  const vp = page.getViewport({ scale: 1 });
  const c = document.getElementById("c");
  c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
  await page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
  const px = c.getContext("2d").getImageData(0,0,c.width,c.height).data;
  let nonWhite = 0;
  for (let i=0;i<px.length;i+=4) if (px[i]<250||px[i+1]<250||px[i+2]<250) nonWhite++;
  window.__probe = { ok:true, pages: doc.numPages, w: c.width, h: c.height, nonWhite, text };
  out.textContent = "PROBE_DONE " + JSON.stringify(window.__probe);
} catch (e) {
  window.__probe = { ok:false, error: String(e && e.message || e) };
  out.textContent = "PROBE_DONE " + JSON.stringify(window.__probe);
}
</script></body>`;
    res.writeHead(200, { "content-type": "text/html; charset=utf-8" });
    return res.end(html);
  }
  res.writeHead(404); res.end("nf");
});

await new Promise((r) => server.listen(0, "127.0.0.1", r));
const port = server.address().port;
console.log(`\nSERVER_PORT=${port}`);
console.log(`open http://127.0.0.1:${port}/`);

// verify the served worker matches the original, independent of the browser
const served = Buffer.from(await (await fetch(`http://127.0.0.1:${port}/vendor/pdf.worker.min.mjs`)).arrayBuffer());
ok("served route reassembles to the original worker byte-for-byte",
   sha(served) === sha(worker), `${served.length} bytes`);

console.log("\nsummary: " + results.filter(r => r.pass).length + "/" + results.length + " passed");
if (!process.env.PROBE_KEEP_ALIVE) { server.close(); }
else { console.log("holding server open for browser probe…"); }
