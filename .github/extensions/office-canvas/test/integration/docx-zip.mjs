// Reads one part out of a .docx, without Word and without our own reader.
//
// The point is the independence. A round-trip test that reads its result back
// through `readStructure` cannot distinguish "the text is intact" from "the
// text is corrupted and our reader corrupts it back the same way" -- and issue
// #40 was exactly a boundary where one direction was wrong, so a matching bug
// on both sides is not a hypothetical. This walks the zip's central directory
// and inflates the entry itself, so the only thing between the assertion and
// the bytes Word wrote is `zlib`.
//
// Deliberately minimal: enough of the format for OOXML packages Word writes,
// and an explicit failure for anything else rather than a silent wrong answer.

import { readFile } from "node:fs/promises";
import { inflateRawSync } from "node:zlib";

const EOCD_SIGNATURE = 0x06054b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;
const STORED = 0;
const DEFLATED = 8;
/** A size or offset of all-ones means the real value lives in a zip64 field. */
const ZIP64_SENTINEL = 0xffffffff;

/** Finds the end-of-central-directory record, which is at the end but not at a fixed offset. */
function findEndOfCentralDirectory(buf) {
    // The record is 22 bytes plus a comment of up to 65535; Word writes none,
    // but scanning back over the whole legal range costs nothing here.
    const earliest = Math.max(0, buf.length - 22 - 0xffff);
    for (let at = buf.length - 22; at >= earliest; at--) {
        if (buf.readUInt32LE(at) === EOCD_SIGNATURE) return at;
    }
    throw new Error("not a zip archive: no end-of-central-directory record");
}

/**
 * Returns the raw bytes of one entry in a zip archive.
 *
 * @param {Buffer} buf the whole archive
 * @param {string} entryName e.g. "word/document.xml"
 */
export function zipEntryBytes(buf, entryName) {
    const eocd = findEndOfCentralDirectory(buf);
    const entryCount = buf.readUInt16LE(eocd + 10);
    let at = buf.readUInt32LE(eocd + 16);
    if (at === ZIP64_SENTINEL) throw new Error("zip64 archive: not supported by this helper");

    for (let i = 0; i < entryCount; i++) {
        if (buf.readUInt32LE(at) !== CENTRAL_SIGNATURE) {
            throw new Error(`malformed central directory at entry ${i}`);
        }
        const method = buf.readUInt16LE(at + 10);
        const compressedSize = buf.readUInt32LE(at + 20);
        const nameLength = buf.readUInt16LE(at + 28);
        const extraLength = buf.readUInt16LE(at + 30);
        const commentLength = buf.readUInt16LE(at + 32);
        const localOffset = buf.readUInt32LE(at + 42);
        // Entry names in a .docx are ASCII, so the encoding of *this* read is
        // not a variable -- which matters, since it is a boundary of the kind
        // under test.
        const name = buf.toString("latin1", at + 46, at + 46 + nameLength);

        if (name === entryName) {
            if (compressedSize === ZIP64_SENTINEL || localOffset === ZIP64_SENTINEL) {
                throw new Error(`zip64 entry ${entryName}: not supported by this helper`);
            }
            if (buf.readUInt32LE(localOffset) !== LOCAL_SIGNATURE) {
                throw new Error(`malformed local header for ${entryName}`);
            }
            // The local header repeats the name and extra fields with lengths
            // of its own, which are not required to match the central copy.
            const dataAt =
                localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28);
            const raw = buf.subarray(dataAt, dataAt + compressedSize);
            if (method === STORED) return Buffer.from(raw);
            if (method === DEFLATED) return inflateRawSync(raw);
            throw new Error(`unsupported compression method ${method} for ${entryName}`);
        }

        at += 46 + nameLength + extraLength + commentLength;
    }

    throw new Error(`no ${entryName} in the package`);
}

/**
 * Reads `word/document.xml` out of a .docx and decodes it as UTF-8.
 *
 * UTF-8 is not a guess: the part carries an XML declaration saying so, and that
 * declaration is asserted here rather than trusted, because the whole value of
 * this helper is that it does not quietly reinterpret bytes.
 */
export async function documentXml(docxPath) {
    const bytes = zipEntryBytes(await readFile(docxPath), "word/document.xml");
    const xml = bytes.toString("utf8");
    if (!/^<\?xml[^?]*encoding="UTF-8"/i.test(xml)) {
        throw new Error(`word/document.xml does not declare UTF-8: ${xml.slice(0, 80)}`);
    }
    return xml;
}

/** Renders a string as `U+XXXX` codepoints, so a failure prints what it saw. */
export const codepoints = (text) =>
    [...text].map((ch) => "U+" + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0")).join(" ");

/** The five named entities XML defines, plus numeric character references. */
function decodeEntities(text) {
    return text
        .replace(/&#x([0-9a-f]+);/gi, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
        .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        // Last, or an escaped entity would be decoded twice.
        .replace(/&amp;/g, "&");
}

/**
 * The document's text, as the concatenation of its `w:t` runs.
 *
 * Concatenated rather than matched run by run because Word splits a run
 * wherever it likes -- proofing state, revision ids, a language boundary -- so
 * a search for a phrase inside a single `w:t` element can fail on a document
 * that is perfectly correct. Numeric character references are decoded for the
 * same reason: a non-ASCII character Word chose to escape is still that
 * character, and a test that missed it would report corruption that is not
 * there.
 */
export const documentPlainText = (xml) =>
    [...xml.matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)].map((m) => decodeEntities(m[1])).join("");
