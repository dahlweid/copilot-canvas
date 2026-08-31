// Probe: what Range requests does pdf.js actually issue, and does the
// repo's parseByteRange (server.mjs, PR #5) handle every form?
//
// Issue #9 claims L4 is parseByteRange's "first real consumer". This checks
// that claim against a live pdf.js instead of assuming it.
//
// Run: node probe-range-requests.mjs <some.pdf>, then open the printed URL.
// The probe exits by itself once the page reports back:
//
//   0  pdf.js finished; the range table is a measurement.
//   1  pdf.js failed in the browser; the range table is inconclusive.
//   2  the page never reported back at all (never opened, closed early, or
//      the in-page script died somewhere its catch could not see).
//
// Those three exits, and the `/done` route they hang off, exist because the
// probe's headline finding is a *negative* -- "suffix-form ranges: NONE" --
// and that negative is cited in shipped source (server.mjs, parseByteRange).
// The in-page script used to report its outcome only into the DOM, which
// nothing read, while the report tick keyed on the requests pdf.js had *made*.
// So pdf.js failing after its first ranged read still printed a whole-looking
// table ending in `NONE`, and SIGINT exited 0 either way: a swallowed error
// and a genuine absence were the same bytes on stdout. Measured, that is not
// hypothetical -- serving 500 to the second range request reproduces a report
// byte-identical on its verdict line to a healthy run (issue #109).
//
// The rule this file now follows: the verdict line is printed only on evidence
// that the measurement happened. Everything else prints INCONCLUSIVE and exits
// non-zero, so a broken instrument cannot be mistaken for a clean negative.
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
// How much of `seen` the live tail has already printed. The tail used to
// truncate `seen` instead, which is why the final report could not be built
// from the whole run.
let reported = 0;
// The page's own outcome, or null while it has not reported. `null` is the
// state that must never print a verdict.
let result = null;

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
  if (url === "/done" && req.method === "POST") {
    const chunks = [];
    let size = 0;
    for await (const c of req) {
      size += c.length;
      if (size > 64 * 1024) break;
      chunks.push(c);
    }
    let posted;
    try {
      posted = JSON.parse(Buffer.concat(chunks).toString("utf8"));
    } catch (e) {
      posted = { ok: false, error: `unparseable /done body: ${e.message}` };
    }
    res.writeHead(204);
    res.end();
    // First report wins. A reload would append to `seen` from a second load,
    // so accepting a later one would mix two runs into one table.
    if (result === null) {
      result = posted;
      // Small grace so any /doc.pdf request still in flight when the page
      // resolved lands in `seen` before the table is built from it.
      setTimeout(() => finish(), 250);
    }
    return;
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
// Report to the server, not just to the DOM. Nothing reads #out, so a result
// left only there is a result the transcript cannot see.
const report = async (r) => {
  out.textContent = "PROBE_DONE " + JSON.stringify(r);
  await fetch("/done", {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify(r),
  });
};
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
  await report({ ok:true, pages:n, lastPageItems: tc.items.length });
} catch (e) {
  await report({ ok:false, error:String(e&&e.message||e) });
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
console.log(`open http://127.0.0.1:${server.address().port}/`);

// Print at most this many distinct forms in one block. Not load-bearing --
// the suffix scan below always runs over the whole set -- but a silent cap
// under-reports a run with more forms than this, so the overflow is stated.
const SHOW = 12;

function printForms(uniq) {
  for (const r of uniq.slice(0, SHOW)) {
    let parsed;
    try {
      parsed = JSON.stringify(parseByteRange(r, pdf.length));
    } catch (e) {
      parsed = "THREW: " + e.message;
    }
    console.log(`  ${r.padEnd(28)} parseByteRange -> ${parsed}`);
  }
  if (uniq.length > SHOW) {
    console.log(`  … ${uniq.length - SHOW} further distinct range form(s) not shown (cap ${SHOW})`);
  }
}

let finished = false;
let tail = null;
function finish() {
  if (finished) return;
  finished = true;
  clearInterval(tail);

  const ranges = seen.filter((s) => s.range).map((s) => s.range);
  const uniq = [...new Set(ranges)];
  console.log(`\n=== FINAL REPORT: ${seen.length} /doc.pdf request(s), ${uniq.length} distinct range form(s) ===`);
  printForms(uniq);

  const suffix = uniq.filter((r) => /^bytes=-\d+$/.test(r));
  const sawSuffix = suffix.length ? suffix.join(", ") : "none seen";
  let code;
  if (result?.ok === true) {
    console.log(`browser: OK — ${result.pages} pages, ${result.lastPageItems} text items on the last page`);
    // The only path that may print the bare negative. The two below reach their
    // line without evidence that pdf.js finished, so for them an empty `suffix`
    // means "we did not see one", not "pdf.js does not send one".
    console.log(`suffix-form ranges (bytes=-N): ${suffix.length ? suffix.join(", ") : "NONE"}`);
    code = 0;
  } else if (result) {
    console.log(`browser: FAILED — ${result.error ?? JSON.stringify(result)}`);
    console.log(`suffix-form ranges (bytes=-N): INCONCLUSIVE (${sawSuffix}) — pdf.js did not finish, so absence here is not evidence of absence.`);
    code = 1;
  } else {
    console.log("browser: NO RESULT — the page never reported to /done (never opened, closed early, or the in-page script died before its catch).");
    console.log(`suffix-form ranges (bytes=-N): INCONCLUSIVE (${sawSuffix}) — nothing confirmed pdf.js ran at all, so absence here is not evidence of absence.`);
    code = 2;
  }

  // Set the code and let the loop drain rather than calling process.exit here:
  // stdout to a pipe is asynchronous on Windows, and exiting from inside the
  // callback that just wrote the verdict can truncate it -- which would be this
  // same bug in a new place. The unref'd timer is the backstop if some handle
  // outlives the close, by which time the writes have flushed.
  process.exitCode = code;
  server.closeAllConnections?.();
  server.close();
  setTimeout(() => process.exit(code), 2000).unref();
}

process.on("SIGINT", () => finish());

tail = setInterval(() => {
  // Live tail only: progress while the human gets the page open. It prints no
  // verdict, because at this point nothing has established that pdf.js
  // finished -- `seen` is filled by the requests pdf.js *made*.
  const fresh = seen.slice(reported);
  if (fresh.length) {
    reported = seen.length;
    const uniq = [...new Set(fresh.filter((s) => s.range).map((s) => s.range))];
    console.log(`\n--- ${fresh.length} new request(s), ${uniq.length} distinct range form(s) so far ---`);
    printForms(uniq);
  }
}, 3000);
