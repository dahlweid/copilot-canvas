// Throwaway probe: do any installed fonts forbid embedding?
//
// If Word cannot embed a font, its PDF export references it without a /FontFile
// and pdf.js falls back to its standard font data -- which is binary, and C4
// refuses binary files. That is the one remaining way the font question could
// still bite issue #9.
//
// fsType lives in the OS/2 table (uint16 at offset 8). Bit 1 (0x0002) means
// "restricted licence: no embedding". Reads font files directly, so it starts
// no Office process and cannot disturb the sibling sessions.

import { readdirSync, openSync, readSync, closeSync, statSync } from "node:fs";
import path from "node:path";

const DIR = process.argv[2] ?? "C:\\Windows\\Fonts";

function u16(b, o) { return (b[o] << 8) | b[o + 1]; }
function u32(b, o) { return ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0; }

function read(fd, len, pos) {
    const b = Buffer.alloc(len);
    readSync(fd, b, 0, len, pos);
    return b;
}

/** fsType values for every font face in one file (a .ttc holds several). */
function fsTypesOf(file) {
    const fd = openSync(file, "r");
    try {
        const head = read(fd, 12, 0);
        const offsets = [];
        if (head.toString("latin1", 0, 4) === "ttcf") {
            const n = u32(head, 8);
            const dir = read(fd, 4 * n, 12);
            for (let i = 0; i < n; i++) offsets.push(u32(dir, i * 4));
        } else {
            offsets.push(0);
        }
        const out = [];
        for (const base of offsets) {
            const h = read(fd, 12, base);
            const numTables = u16(h, 4);
            if (numTables === 0 || numTables > 512) continue;
            const dir = read(fd, 16 * numTables, base + 12);
            for (let i = 0; i < numTables; i++) {
                const tag = dir.toString("latin1", i * 16, i * 16 + 4);
                if (tag !== "OS/2") continue;
                const off = u32(dir, i * 16 + 8);
                out.push(u16(read(fd, 12, off), 8));
                break;
            }
        }
        return out;
    } finally {
        closeSync(fd);
    }
}

const RESTRICTED = 0x0002;
const NO_SUBSET = 0x0200;

let scanned = 0, failed = 0;
const restricted = [], noSubset = [];

for (const name of readdirSync(DIR)) {
    const file = path.join(DIR, name);
    if (!/\.(ttf|otf|ttc)$/i.test(name)) continue;
    try {
        if (!statSync(file).isFile()) continue;
        for (const fsType of fsTypesOf(file)) {
            scanned++;
            // fsType is a bitfield, but 0x0002 is exclusive: when set, no other
            // embedding bit is meaningful.
            if (fsType & RESTRICTED) restricted.push(`${name} (fsType=0x${fsType.toString(16)})`);
            if (fsType & NO_SUBSET) noSubset.push(`${name} (fsType=0x${fsType.toString(16)})`);
        }
    } catch {
        failed++;
    }
}

console.log(`scanned          ${scanned} faces in ${DIR}`);
console.log(`unreadable       ${failed}`);
console.log(`no-embed (0x0002) ${restricted.length}`);
for (const r of restricted.slice(0, 20)) console.log(`   ${r}`);
console.log(`no-subset (0x0200) ${noSubset.length}`);
for (const r of noSubset.slice(0, 10)) console.log(`   ${r}`);
console.log("");
console.log(
    restricted.length === 0
        ? "=> no installed font forbids embedding; the pdf.js standard-font fallback cannot fire\n   from locally-authored documents on this machine. Residual risk is a document carrying\n   a restricted font from elsewhere."
        : "=> restricted fonts exist; use one of the above as the test case for issue #9.",
);
