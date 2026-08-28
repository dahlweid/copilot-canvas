// Throwaway probe: does Word's fixed-format export embed its fonts?
// If it always does, pdf.js never needs its binary standard-font files, and
// constraint C4 (the gist channel refuses binaries) does not block ADR 0004.
// Runs on a cached PDF, so it does not start Word and cannot disturb the
// sibling sessions currently driving it.

import { readFileSync } from "node:fs";
import zlib from "node:zlib";

const file = process.argv[2];
const buf = readFileSync(file);

// PDF 1.5+ hides most objects inside compressed object streams, so scanning the
// raw bytes alone would under-report. Inflate every stream we can find first.
let text = buf.toString("latin1");
const chunks = [text];
const re = /stream\r?\n/g;
let m;
let inflated = 0;
while ((m = re.exec(text)) !== null) {
    const start = m.index + m[0].length;
    const end = text.indexOf("endstream", start);
    if (end < 0) continue;
    try {
        const out = zlib.inflateSync(buf.subarray(start, end));
        chunks.push(out.toString("latin1"));
        inflated++;
    } catch {
        // not a flate stream, or not independently decodable — skip
    }
}
const all = chunks.join("\n");

const count = (needle) => all.split(needle).length - 1;

const embedded = count("/FontFile") + count("/FontFile2") + count("/FontFile3");
const baseFonts = [...all.matchAll(/\/BaseFont\s*\/([A-Za-z0-9+#,\-_.]+)/g)].map((x) => x[1]);
const unique = [...new Set(baseFonts)];

// A subset tag is six capitals then '+', e.g. ABCDEE+Calibri. Word emits one
// per embedded subset, so it is corroborating evidence independent of /FontFile.
const subset = unique.filter((f) => /^[A-Z]{6}\+/.test(f));
const nonSubset = unique.filter((f) => !/^[A-Z]{6}\+/.test(f));

// The 14 standard fonts are exactly the ones a viewer must supply itself.
const STANDARD = /^(Helvetica|Courier|Times|Symbol|ZapfDingbats|Arial)([-,].*)?$/;
const needStandard = nonSubset.filter((f) => STANDARD.test(f));

console.log(`file            ${file}`);
console.log(`size            ${buf.length} bytes`);
console.log(`streams inflated ${inflated}`);
console.log(`/FontFile*      ${embedded}`);
console.log(`unique BaseFont ${unique.length}`);
console.log(`  subset-tagged (embedded): ${subset.join(", ") || "(none)"}`);
console.log(`  not subset-tagged:        ${nonSubset.join(", ") || "(none)"}`);
console.log(`standard-14 referenced:     ${needStandard.join(", ") || "(none)"}`);
console.log(`CID / Type0 fonts:          ${count("/Type0")}`);
console.log(`CMap references (non-Identity): ${count("/Encoding /") - count("/Encoding /Identity")}`);
console.log("");
console.log(
    embedded > 0 && needStandard.length === 0
        ? "=> every font is embedded; pdf.js standard font data would not be consulted for THIS document"
        : "=> NOT all fonts embedded; pdf.js would need its standard font data (binary, blocked by C4)",
);
