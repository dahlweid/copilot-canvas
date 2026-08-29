/*
 * probe-stdin-encoding.mjs -- what happens to non-ASCII text on its way into a
 * document, and whether any assertion could have seen it.
 *
 * Context: issue #40 reported that word-host.ps1 set [Console]::OutputEncoding
 * to UTF-8 and never set [Console]::InputEncoding, so Node's UTF-8 JSON was
 * decoded by Windows PowerShell 5.1 as the OEM codepage. create_document
 * authors text from a spec, so this was its default case, not an edge case.
 *
 * THE FIX HAS LANDED (#46, both encoding lines are now set), AND THESE
 * EXPECTATIONS HAVE BEEN INVERTED TO MATCH. They previously pinned the broken
 * behaviour on purpose; they now specify the correct behaviour, so this file is
 * a regression guard for #40 rather than a description of it.
 *
 * What it asserts, exiting 1 if the machine disagrees with any:
 *
 *   1. WRITE IS LOSSLESS. Text sent through create arrives in
 *                        word/document.xml unchanged. Read straight out of the
 *                        zip, so it is the file being checked, not some
 *                        rendering of it. The old corrupted form is asserted
 *                        ABSENT by name, so a regression cannot pass as merely
 *                        "different".
 *
 *   2. READ IS FAITHFUL. read_document returns the same characters. This
 *                        direction was never broken, and it is why the defect
 *                        was visible end to end once anyone looked.
 *
 *   3. IT ROUND-TRIPS.   Feeding the value back in returns it unchanged.
 *
 *   4. WHICH ARM DISCRIMINATES, and this is the one worth keeping. A
 *                        write-then-read-back assertion exercises exactly ONE
 *                        crossing, so under #40 it could pass while the
 *                        document was permanently uneditable. The discriminating
 *                        case is to read the value back and feed it in as
 *                        `expectedText`, because the defect lived in the SECOND
 *                        crossing. That is asserted here to succeed, with an
 *                        ASCII paragraph in the same document as the control:
 *                        without it, a red here would be equally consistent
 *                        with the edit path simply being broken.
 *
 * ON THE INVERSION ITSELF, recorded because the rule that governed it is worth
 * more than the diff. #46's author left the standing instruction: invert in the
 * same commit that picks up the fix, and treat any expectation still red
 * afterwards as a SECOND, SEPARATE defect -- an expected red is the perfect
 * hiding place for an unexpected one.
 *
 * On the rebase this went red five times, and four were the direct inversions
 * above. The fifth -- "both paragraphs located for the edit arm", which expected
 * `true` and got `false` -- had the wrong SHAPE for an inversion, so it was
 * traced rather than assumed: the lookup at what was then line 184 found the
 * paragraph by its MOJIBAKE text, which the fix means no longer exists. Same
 * inversion, reached through a constant rather than an expectation. Confirmed by
 * reading the lookup, not inferred from it having failed alongside the others.
 * No residual defect.
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
// What #40 used to produce. Kept so the regression can be named rather than
// merely differed from: "not the old mojibake" is a weaker claim than "equal to
// what was sent", so both are asserted.
const WAS_CORRUPTED_TO = "Gr\u251c\u255d\u251c\u0192e aus M\u251c\u255dnchen";
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
    check("non-ASCII survives into word/document.xml", onDisk[0], SENT);
    check("and is not the corruption #40 produced", onDisk[0] === WAS_CORRUPTED_TO, false);
    check("ASCII on the same write is untouched", onDisk[1], ASCII_CONTROL);

    // ---- 2. the read direction is faithful, as it always was -------------
    const reader = new DocumentReader({ host, workRoot, log: () => {} });
    const back = await reader.read(doc, { limit: 10 });
    const read0 = back.paragraphs?.[0]?.text ?? "";
    check("read_document returns exactly what is on disk", read0, SENT);
    check("so a round-trip assertion is green", read0 === SENT, true);

    // ---- 3. it round-trips, so expectedText can match --------------------
    const echoed = await host.create({
        path: echo,
        blocks: [{ kind: "paragraph", text: SENT }],
    });
    check("echo create reports success", echoed.status, "created");
    const echoBack = await reader.read(echo, { limit: 5 });
    const echo0 = echoBack.paragraphs?.[0]?.text ?? "";
    check("sending the value back in is idempotent", echo0, SENT);

    // ---- 4. the arm that discriminates: read back, then edit -------------
    // A single-crossing round trip could not see #40. Feeding the value straight
    // back as `expectedText` can, because that is the second crossing -- so this
    // is the arm that actually guards the fix. The ASCII paragraph is the
    // control: without it, a red here would be equally consistent with the edit
    // path simply being broken.
    const forEdit = await reader.read(doc, { limit: 10 });
    const moji = forEdit.paragraphs.find((p) => p.text === SENT);
    const ascii = forEdit.paragraphs.find((p) => p.text === ASCII_CONTROL);
    check("both paragraphs located for the edit arm", Boolean(moji && ascii), true);

    if (moji && ascii) {
        const mojiEdit = await host.edit({
            path: doc,
            wordIndex: moji.wordIndex,
            expectedText: moji.text,
            op: "replace_text",
            text: "replaced",
            headingLevel: null,
        });
        check("text read back CAN address the paragraph it came from", mojiEdit.status, "edited");

        // Re-read before the second edit. Under #40 the first edit changed
        // nothing, so one read served both arms; now that it succeeds it is a
        // mutation, and an address is a coordinate rather than a handle (ADR
        // 0006). Reusing the earlier read here would be the exact cached-address
        // defect this repo keeps finding, hidden inside a probe about encoding.
        const afterEdit = await reader.read(doc, { limit: 10 });
        const ascii2 = afterEdit.paragraphs.find((p) => p.text === ASCII_CONTROL);
        check("the ASCII control is still locatable after the first edit", Boolean(ascii2), true);

        if (ascii2) {
            const asciiEdit = await host.edit({
                path: doc,
                wordIndex: ascii2.wordIndex,
                expectedText: ascii2.text,
                op: "replace_text",
                text: "replaced",
                headingLevel: null,
            });
            check("control: an ASCII paragraph in the same document edits", asciiEdit.status, "edited");
        }
    }
} finally {
    await host.request("quit", {}).catch(() => {});
    await host.dispose?.().catch?.(() => {});
}

console.log(
    failures.length === 0
        ? "\nVERDICT: non-ASCII survives the write, the read, and a second crossing."
        : `\nVERDICT: ${failures.length} assertion(s) disagreed with this machine.`,
);
process.exit(failures.length === 0 ? 0 : 1);
