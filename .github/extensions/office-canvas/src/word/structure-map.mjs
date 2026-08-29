// Structure map: the document's paragraphs, addressed.
//
// Pure. Takes the WordprocessingML that `Range.WordOpenXML` produced and turns
// it into the map an agent addresses edits against. Nothing here touches Word,
// COM or the filesystem, which is what lets it be tested on a hosted runner.
//
// Two things this file exists to get right, both of which have already bitten:
//
//   * **Style names are localized, and style ids are worse than localized.** On
//     this German Word, `Selection.Style = "Heading 1"` throws — and the id in
//     the file is not `Überschrift1` either. Word mints the id from the
//     localized name and strips the non-ASCII characters, so what it actually
//     writes is `berschrift1`. A style id is therefore carried, never matched.
//     Resolution goes through the styles part, whose `w:name` holds the
//     canonical built-in name ("heading 1") regardless of UI language, and
//     falls back to `w:outlineLvl` for styles that carry a genuinely localized
//     name.
//
//   * **Word exposes no stable paragraph identity.** Addresses are derived from
//     heading path + text + occurrence index. The occurrence index stays even
//     when text looks unique, because the document that proved the scheme had
//     no duplicate text at all and so never tested it.

import { createHash } from "node:crypto";
import { attr, childNamed, childrenNamed, findFirst, parseXml, textOf } from "./ooxml.mjs";

export class StructureError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "StructureError";
        this.code = code;
    }
}

/** `w:name` for a built-in heading style, in any UI language. */
const BUILTIN_HEADING_NAME = /^heading\s*([1-9])$/i;
/** `w:outlineLvl` is 0-based, and 9 means body text rather than heading 10. */
const BODY_TEXT_OUTLINE_LEVEL = 9;

const ADDRESS_PREFIX = "p:";
const ADDRESS_HEX_LENGTH = 12;

/** Markup that carries no paragraph text, or text we deliberately exclude. */
const NON_TEXT = new Set([
    "pPr", // paragraph properties, including numbering
    "rPr", // run properties
    "sectPr",
    "instrText", // field instructions, not the field result
    "delText", // text a tracked change deleted
    "delInstrText",
    // A text box is anchored *inside* a body paragraph's run, but Word keeps it
    // in a separate story: `Document.Paragraphs` never returns its paragraphs,
    // and the anchor paragraph's `Range.Text` does not contain its text. They
    // therefore get no address of their own, and folding their text into the
    // anchor would be worse than omitting it -- the anchor's address would move
    // whenever unrelated text-box content changed, and its reported `text` would
    // match no single editable range, so an edit replacing it would destroy the
    // text box. Measured on a fixture with a text box in a table cell: without
    // this the map reported "TEXTBOX PARAGRAPH ONEcell 11" where Word reports
    // "cell 11". Text-box content is out of scope for addressing rather than
    // half in it. Covers `w:txbxContent` under both the DrawingML and the
    // legacy VML spelling.
    "txbxContent",
]);

/**
 * Collapses whitespace and trims. Word splits a sentence across runs for
 * reasons invisible to the reader (spell-check state, a stray formatting
 * mark), so an address keyed on raw text would move when nothing did.
 */
export function normalizeText(value) {
    return String(value ?? "")
        .replace(/\s+/gu, " ")
        .trim();
}

const intOrNull = (value) => {
    if (value === null || value === undefined || value === "") return null;
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : null;
};

/**
 * Mints the address for a paragraph.
 *
 * A hash rather than a readable composite, because the composite is unbounded
 * in length and the structure map carries `headingPath`, `text` and
 * `occurrence` beside it — so an address stays verifiable without being long.
 * Deterministic: the same tuple always produces the same address, which is what
 * lets a later read resolve an address minted by an earlier one.
 */
export function mintAddress({ headingPath = [], text = "", occurrence = 1 } = {}) {
    const key = [
        headingPath.map(normalizeText).join("\u001f"),
        normalizeText(text),
        String(occurrence),
    ].join("\u0000");
    return ADDRESS_PREFIX + createHash("sha256").update(key, "utf8").digest("hex").slice(0, ADDRESS_HEX_LENGTH);
}

/**
 * Reads the styles part into `styleId -> { styleId, name, headingLevel }`.
 *
 * `headingLevel` is resolved once here, following `w:basedOn` where needed, so
 * a paragraph lookup is a map hit rather than a walk.
 */
export function resolveStyles(root) {
    const stylesPart = findFirst(root, "styles");
    const table = new Map();
    if (!stylesPart) return table;

    const raw = new Map();
    for (const style of childrenNamed(stylesPart, "style")) {
        const styleId = attr(style, "styleId") ?? attr(childNamed(style, "name"), "val");
        if (!styleId) continue;
        raw.set(styleId, {
            styleId,
            type: attr(style, "type"),
            name: attr(childNamed(style, "name"), "val"),
            basedOn: attr(childNamed(style, "basedOn"), "val"),
            outlineLevel: intOrNull(attr(childNamed(childNamed(style, "pPr"), "outlineLvl"), "val")),
        });
    }

    const headingLevelOf = (styleId, seen) => {
        const style = raw.get(styleId);
        if (!style || seen.has(styleId)) return null;
        seen.add(styleId);

        const builtin = BUILTIN_HEADING_NAME.exec(style.name ?? "");
        if (builtin) return Number(builtin[1]);

        if (style.outlineLevel !== null) {
            return style.outlineLevel >= BODY_TEXT_OUTLINE_LEVEL ? null : style.outlineLevel + 1;
        }
        // A custom heading style usually inherits its outline level rather than
        // restating it, so the chain has to be followed to the end.
        return style.basedOn ? headingLevelOf(style.basedOn, seen) : null;
    };

    for (const [styleId, style] of raw) {
        table.set(styleId, {
            styleId,
            name: style.name ?? styleId,
            headingLevel: headingLevelOf(styleId, new Set()),
        });
    }
    return table;
}

/** Picks the branch of an `mc:AlternateContent` a Word reader would render. */
function alternateContentBranch(node) {
    return childNamed(node, "Choice") ?? childNamed(node, "Fallback");
}

/**
 * Every paragraph in the body, in document order, including those inside
 * tables. Paragraphs never nest, so a `w:p` is never descended into.
 *
 * Each entry also carries `wordIndex`: the position this paragraph occupies in
 * Word's own `Document.Paragraphs` collection. That is what lets an address
 * minted here be turned back into a COM range without a per-paragraph walk.
 *
 * The two collections are *not* the same sequence, which is the whole reason
 * this is computed rather than assumed. Measured on a fixture with one 2x2
 * table and one text box: the markup has 8 paragraphs, `Document.Paragraphs`
 * has 10. The difference is exactly the **end-of-row marks** — Word counts the
 * mark that terminates each `w:tr` as a paragraph of its own, and there is no
 * `w:p` for it in the markup. A simpler fixture with no tables agreed 42 to 42,
 * which is precisely why agreement on a simple document proves nothing.
 */
function collectParagraphs(body) {
    const found = [];
    let wordIndex = 0;
    const visit = (node, inTable) => {
        for (const child of node.children) {
            if (typeof child === "string") continue;
            if (child.local === "p") {
                wordIndex += 1;
                found.push({ node: child, inTable, wordIndex });
            } else if (child.local === "AlternateContent") {
                const branch = alternateContentBranch(child);
                if (branch) visit(branch, inTable);
            } else if (child.local === "tr") {
                visit(child, true);
                wordIndex += 1; // the end-of-row mark
            } else {
                visit(child, inTable || child.local === "tbl");
            }
        }
    };
    visit(body, false);
    return found;
}

/** The text a reader sees, with layout marks turned into their characters. */
export function paragraphText(paragraph) {
    const parts = [];
    const visit = (node) => {
        for (const child of node.children) {
            if (typeof child === "string") continue;
            const local = child.local;
            if (NON_TEXT.has(local)) continue;
            if (local === "t") parts.push(textOf(child));
            else if (local === "tab") parts.push("\t");
            else if (local === "br" || local === "cr") parts.push("\n");
            else if (local === "noBreakHyphen") parts.push("-");
            else if (local === "softHyphen") continue;
            else if (local === "AlternateContent") {
                const branch = alternateContentBranch(child);
                if (branch) visit(branch);
            } else visit(child);
        }
    };
    visit(paragraph);
    return parts.join("");
}

/**
 * Maintains the stack of enclosing heading texts.
 *
 * A heading's own path is its *ancestors*, not itself, so an address under
 * "Chapter 1 > Section 1.1" reads the way a person would describe the location.
 */
function pushHeading(stack, level, text) {
    if (stack.length > level - 1) stack.length = level - 1;
    while (stack.length < level - 1) stack.push("");
    stack.push(text);
}

/**
 * Builds the structure map.
 *
 * Accepts either the Flat OPC package `Range.WordOpenXML` returns or a bare
 * `w:document`; without a styles part, heading levels fall back to whatever the
 * paragraphs state directly.
 *
 * `limit`/`offset` page the returned paragraphs only. Addresses are always
 * minted across the whole document, so paging never changes an address.
 */
export function buildStructureMap(xml, { limit = 0, offset = 0 } = {}) {
    const root = parseXml(xml);
    const body = findFirst(root, "body");
    if (!body) {
        throw new StructureError("no_document_body", "The markup contains no document body.");
    }

    const styles = resolveStyles(root);
    const headingStack = [];
    const occurrences = new Map();
    const paragraphs = [];

    collectParagraphs(body).forEach(({ node, inTable, wordIndex }, position) => {
        const properties = childNamed(node, "pPr");
        const styleId = attr(childNamed(properties, "pStyle"), "val");
        const style = styleId ? (styles.get(styleId) ?? null) : null;

        // A direct w:outlineLvl overrides the style, which is how a document
        // promotes or demotes a single paragraph without a new style.
        const direct = intOrNull(attr(childNamed(properties, "outlineLvl"), "val"));
        const headingLevel =
            direct !== null
                ? direct >= BODY_TEXT_OUTLINE_LEVEL
                    ? null
                    : direct + 1
                : (style?.headingLevel ?? null);

        const text = paragraphText(node);
        const normalized = normalizeText(text);
        const headingPath = headingStack.slice(0, headingLevel ? headingLevel - 1 : headingStack.length);

        const occurrenceKey = `${headingPath.map(normalizeText).join("\u001f")}\u0000${normalized}`;
        const occurrence = (occurrences.get(occurrenceKey) ?? 0) + 1;
        occurrences.set(occurrenceKey, occurrence);

        const numbering = childNamed(properties, "numPr");
        const list = numbering
            ? {
                  numId: attr(childNamed(numbering, "numId"), "val"),
                  level: intOrNull(attr(childNamed(numbering, "ilvl"), "val")) ?? 0,
              }
            : null;

        paragraphs.push({
            address: mintAddress({ headingPath, text, occurrence }),
            index: position + 1,
            // Where this paragraph sits in Word's own `Document.Paragraphs`.
            // Not the same as `index` once a table is involved; an edit
            // addresses Word through this, never through `index`.
            wordIndex,
            text,
            // Word's own id for the style, verbatim: `berschrift1` on this
            // German Word, because Word mints the id from the localized name
            // and drops the non-ASCII characters.
            //
            // Reported for identification, not for reapplication. Assigning a
            // style *by* this id throws -- measured, and recorded next to
            // `Set-ParagraphHeadingLevel` in `word-host.ps1`, where the English
            // `Heading 1` throws too. The write side therefore names no style
            // at all: it assigns numeric `wd*` constants or another paragraph's
            // Style *object*, and `edit-intent.mjs` rejects `styleId` as an
            // input outright.
            styleId: styleId ?? null,
            styleName: style?.name ?? styleId ?? null,
            headingLevel,
            headingPath,
            occurrence,
            inTable,
            list,
        });

        if (headingLevel) pushHeading(headingStack, headingLevel, normalized);
    });

    const start = Math.max(0, Math.trunc(offset) || 0);
    const count = limit > 0 ? Math.trunc(limit) : paragraphs.length;
    const page = paragraphs.slice(start, start + count);

    return {
        paragraphCount: paragraphs.length,
        headingCount: paragraphs.filter((p) => p.headingLevel !== null).length,
        styleCount: styles.size,
        offset: start,
        returned: page.length,
        truncated: start + page.length < paragraphs.length,
        paragraphs: page,
    };
}
