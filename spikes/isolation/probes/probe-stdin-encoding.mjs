/*
 * probe-stdin-encoding.mjs -- what happens to non-ASCII text on its way into a
 * document, and whether any assertion could have seen it.
 *
 * Context: issue #40 reports that word-host.ps1 sets [Console]::OutputEncoding
 * to UTF-8 and never sets [Console]::InputEncoding, so Node's UTF-8 JSON is
 * decoded by Windows PowerShell 5.1 as the OEM codepage. create_document
 * authors text from a spec, so this is its default case, not an edge case.
 *
 * THIS PROBE DOES NOT FIX ANYTHING. The encoding preamble is owned elsewhere.
 * It exists because three claims are being made about #26's test coverage and
 * this repo does not accept a claim about platform behaviour that is not backed
 * by a probe someone actually ran.
 *
 * It asserts three things, and exits 1 if the machine disagrees with any:
 *
 *   1. WRITE IS LOSSY.   Text sent through create arrives in word/document.xml
 *                        as OEM-decoded UTF-8 bytes: 'ue' -> U+251C U+255D.
 *                        Read straight out of the zip, so it is the file that
 *                        is wrong, not some rendering of it.
 *
 *   2. READ IS FAITHFUL. read_document returns those same characters, exactly.
 *                        This is the claim worth being careful about, because
 *                        "the read direction is clean" is easy to hear as "the
 *                        two bugs cancel and it looks fine". They do not
 *                        cancel. The read being faithful is precisely what
 *                        makes the corruption VISIBLE end to end -- so a single
 *                        non-ASCII round-trip assertion would have caught #40
 *                        outright. Nothing hid this; nobody looked.
 *
 *   3. IT COMPOUNDS.     Sending the mojibake back in does not return it
 *                        unchanged -- it is corrupted again. So the reason a
 *                        non-ASCII paragraph is uneditable is NOT that both
 *                        sides crossed the same boundary (that would MATCH).
 *                        It is that expectedText is corrupted a second time on
 *                        the way in and can never equal what is stored.
 *
 * Needs a real, licensed Word. Starts and stops its own instance.
 *
 * Run: node spikes/isolation/probes/probe-stdin-encoding.mjs
 */

import { mkdtemp, readFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { inflateRawSync } from "node:zlib";

import { WordHost } from "../../../.github/extensions/office-canvas/src/word/word-host.mjs";
import { DocumentReader } from "../../../.github/extensions/office-canvas/src/word/document-reader.mjs";

// Written as escapes on purpose. Windows PowerShell reads a BOM-less UTF-8
// script as ANSI, and this repo keeps its probes ASCII-only so a file cannot
// change meaning by being opened. The same discipline here means the literal
// below cannot itself be the thing that got mangled.
const SENT = "Gr\u00fc\u00dfe aus M\u00fcnchen"; // Gruesse aus Muenchen
const EXPECT_ON_DISK = "Gr\u251c\u255d\u251c\u0192e aus M\u251c\u255dnchen";
const ASCII_CONTROL = "Pure ASCII control line.";

const failures = [];
function check(label, actual, expected) {
    const ok = actual === expected;
    console.log(`${ok ? "ok  " : "FAIL"}  ${label}`);
    if (!ok) {
        console.log(`        expected: ${JSON.stringify(expected)}`);
        console.log(`        actual  : ${JSON.stringify(actual)}`);
        failures.push(label);
    }
}

/**
 * Reads one entry out of a zip without a zip library and without a subprocess.
 *
 * A .docx is a zip, and the obvious route to its bytes is Expand-Archive or a
 * shell one-liner. Both put a document path through a command-line parser, and
 * this repo has measured that parser corrupting ordinary filenames. There is no
 * parser to escape if there is no command line, so this walks the central
 * directory itself.
 */
function readZipEntry(buf, wanted) {
    let eocd = -1;
    for (let i = buf.length - 22; i >= 0; i--) {
        if (buf.readUInt32LE(i) === 0x06054b50) {
            eocd = i;
            break;
        }
    }
    if (eocd < 0) throw new Error("no end-of-central-directory record");

    const count = buf.readUInt16LE(eocd + 10);
    let p = buf.readUInt32LE(eocd + 16);
    for (let i = 0; i < count; i++) {
        if (buf.readUInt32LE(p) !== 0x02014b50) throw new Error("bad central directory entry");
        const method = buf.readUInt16LE(p + 10);
        const compressedSize = buf.readUInt32LE(p + 20);
        const nameLen = buf.readUInt16LE(p + 28);
        const extraLen = buf.readUInt16LE(p + 30);
        const commentLen = buf.readUInt16LE(p + 32);
        const localOffset = buf.readUInt32LE(p + 42);
        const name = buf.toString("utf8", p + 46, p + 46 + nameLen);

        if (name === wanted) {
            const lNameLen = buf.readUInt16LE(localOffset + 26);
            const lExtraLen = buf.readUInt16LE(localOffset + 28);
            const start = localOffset + 30 + lNameLen + lExtraLen;
            const raw = buf.subarray(start, start + compressedSize);
            const bytes = method === 0 ? raw : inflateRawSync(raw);
            return bytes.toString("utf8");
        }
        p += 46 + nameLen + extraLen + commentLen;
    }
    throw new Error(`entry not found: ${wanted}`);
}

/** Concatenates every <w:t> run, which is where Word splits a corrupted word. */
function paragraphTexts(xml) {
    return [...xml.matchAll(/<w:p[ >][\s\S]*?<\/w:p>/g)].map((m) =>
        [...m[0].matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((t) => t[1]).join(""),
    );
}

const docDir = await mkdtemp(path.join(tmpdir(), "probe-stdin-enc-"));
const workRoot = await mkdtemp(path.join(tmpdir(), "probe-stdin-enc-work-"));
const doc = path.join(docDir, "umlaut.docx");
const echo = path.join(docDir, "echo.docx");

const host = new WordHost({ log: () => {} });
try {
    // ---- 1. write is lossy, measured in the file itself -----------------
    const created = await host.create({
        path: doc,
        blocks: [
            { kind: "paragraph", text: SENT },
            { kind: "paragraph", text: ASCII_CONTROL },
        ],
    });
    check("create reports success", created.status, "created");

    const onDisk = paragraphTexts(readZipEntry(await readFile(doc), "word/document.xml"));
    check("non-ASCII is mojibake in word/document.xml", onDisk[0], EXPECT_ON_DISK);
    check("ASCII on the same write is untouched", onDisk[1], ASCII_CONTROL);

    // ---- 2. the read direction is faithful, so it is all visible ---------
    const reader = new DocumentReader({ host, workRoot, log: () => {} });
    const back = await reader.read(doc, { limit: 10 });
    const read0 = back.paragraphs?.[0]?.text ?? "";
    check("read_document returns exactly what is on disk", read0, EXPECT_ON_DISK);
    check("so a round-trip assertion would go red", read0 === SENT, false);

    // ---- 3. it compounds, so no expectedText can ever match --------------
    const echoed = await host.create({
        path: echo,
        blocks: [{ kind: "paragraph", text: EXPECT_ON_DISK }],
    });
    check("echo create reports success", echoed.status, "created");
    const echoBack = await reader.read(echo, { limit: 5 });
    const echo0 = echoBack.paragraphs?.[0]?.text ?? "";
    check("sending the mojibake back in is not idempotent", echo0 === EXPECT_ON_DISK, false);
} finally {
    await host.request("quit", {}).catch(() => {});
    await host.dispose?.().catch?.(() => {});
}

console.log(
    failures.length === 0
        ? "\nVERDICT: write corrupts non-ASCII, read is faithful, corruption compounds."
        : `\nVERDICT: ${failures.length} assertion(s) disagreed with this machine.`,
);
process.exit(failures.length === 0 ? 0 : 1);
