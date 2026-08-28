// The structure map: parsing, style resolution, address minting. Office-free.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";

import { buildStructureMap, mintAddress, normalizeText, paragraphText, resolveStyles, StructureError } from "../../src/word/structure-map.mjs";
import { parseXml } from "../../src/word/ooxml.mjs";
import { bareDocument, flatOpc, GERMAN_STYLES, paragraph, style, table } from "./word-fixtures.mjs";

const map = (body, options) => buildStructureMap(flatOpc(body), options);
const textsOf = (result) => result.paragraphs.map((p) => p.text);
const addressesOf = (result) => result.paragraphs.map((p) => p.address);

// --- style resolution: the localization trap --------------------------------

test("a German heading style id resolves through the canonical name in w:name", () => {
    // The measured trap: on this German Word the id is `berschrift1` -- Word
    // mints the id from "Überschrift 1" and drops the non-ASCII character, so
    // neither "Heading 1" nor "Überschrift1" matches it. `w:name` holds
    // "heading 1" whatever the UI language is.
    const result = map([paragraph("Kapitel Eins", { styleId: "berschrift1" })]);
    const [p] = result.paragraphs;
    assert.equal(p.styleId, "berschrift1", "the id is carried verbatim, for an edit to reapply");
    assert.equal(p.styleName, "heading 1");
    assert.equal(p.headingLevel, 1);
});

test("a style id is carried, never interpreted, whatever characters it holds", () => {
    // The id is opaque: resolution goes through w:name and w:outlineLvl, so an
    // id with an umlaut, without one, or with neither works the same way.
    const styles = [
        style({ styleId: "Überschrift1", name: "heading 1" }),
        style({ styleId: "s7", name: "heading 2" }),
    ];
    const result = buildStructureMap(
        flatOpc([paragraph("Mit Umlaut", { styleId: "Überschrift1" }), paragraph("Opaque", { styleId: "s7" })], styles),
    );
    assert.equal(result.paragraphs[0].styleId, "Überschrift1");
    assert.equal(result.paragraphs[0].headingLevel, 1);
    assert.equal(result.paragraphs[1].styleId, "s7");
    assert.equal(result.paragraphs[1].headingLevel, 2);
});

test("English style ids are never matched, so a heading with no styles part is not guessed", () => {
    // Matching the id would appear to work on an English machine and fail on
    // this one. Better to resolve nothing than to resolve by language.
    const result = buildStructureMap(bareDocument([paragraph("Chapter One", { styleId: "Heading1" })]));
    assert.equal(result.styleCount, 0);
    assert.equal(result.paragraphs[0].headingLevel, null);
    assert.equal(result.paragraphs[0].styleId, "Heading1");
});

test("a style whose name is localized too is resolved by w:outlineLvl", () => {
    const result = map([paragraph("Ein Kapitel", { styleId: "Kapitelüberschrift" })]);
    assert.equal(result.paragraphs[0].styleName, "Kapitelüberschrift");
    assert.equal(result.paragraphs[0].headingLevel, 2, "w:outlineLvl 1 is heading level 2");
});

test("a style inherits its outline level through w:basedOn", () => {
    const result = map([paragraph("Unterkapitel", { styleId: "Unterkapitel" })]);
    assert.equal(result.paragraphs[0].headingLevel, 2);
});

test("a w:basedOn cycle terminates instead of hanging", () => {
    const styles = [style({ styleId: "A", name: "A", basedOn: "B" }), style({ styleId: "B", name: "B", basedOn: "A" })];
    const styleTable = resolveStyles(parseXml(flatOpc([paragraph("x")], styles)));
    assert.equal(styleTable.get("A").headingLevel, null);
    assert.equal(styleTable.get("B").headingLevel, null);
});

test("a direct w:outlineLvl on the paragraph overrides its style", () => {
    const result = map([paragraph("Promoted", { styleId: "berschrift1", outlineLevel: 2 })]);
    assert.equal(result.paragraphs[0].headingLevel, 3);
});

test("w:outlineLvl 9 is body text, not heading 10", () => {
    const result = map([paragraph("Not a heading", { styleId: "berschrift1", outlineLevel: 9 })]);
    assert.equal(result.paragraphs[0].headingLevel, null);
});

test("the style table counts every style, not just the ones in use", () => {
    assert.equal(map([paragraph("x")]).styleCount, GERMAN_STYLES.length);
});

// --- paragraph text ---------------------------------------------------------

test("text is joined across runs, because Word splits them invisibly", () => {
    const raw = `<w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>`;
    assert.equal(map([paragraph(null, { raw })]).paragraphs[0].text, "Hello world");
});

test("tabs, breaks and non-breaking hyphens become their characters", () => {
    const raw = `<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:noBreakHyphen/><w:t>d</w:t><w:softHyphen/></w:r>`;
    assert.equal(map([paragraph(null, { raw })]).paragraphs[0].text, "a\tb\nc-d");
});

test("field instructions and tracked deletions are excluded, insertions are not", () => {
    const raw = [
        `<w:r><w:instrText> PAGE </w:instrText></w:r>`,
        `<w:del><w:r><w:delText>gone</w:delText></w:r></w:del>`,
        `<w:ins><w:r><w:t>kept</w:t></w:r></w:ins>`,
    ].join("");
    assert.equal(map([paragraph(null, { raw })]).paragraphs[0].text, "kept");
});

test("paragraph properties never leak into the text", () => {
    // w:pStyle carries `berschrift1` in an attribute; a naive textOf on the
    // paragraph would still be fine, but w:numPr and friends can hold text.
    assert.equal(map([paragraph("Title", { styleId: "berschrift1" })]).paragraphs[0].text, "Title");
});

test("an mc:AlternateContent contributes its Choice branch once, not both branches", () => {
    const raw = `<mc:AlternateContent><mc:Choice Requires="wps"><w:r><w:t>modern</w:t></w:r></mc:Choice><mc:Fallback><w:r><w:t>legacy</w:t></w:r></mc:Fallback></mc:AlternateContent>`;
    assert.equal(map([paragraph(null, { raw })]).paragraphs[0].text, "modern");
});

test("paragraphText is exported and works on a parsed paragraph directly", () => {
    const root = parseXml(paragraph("standalone"));
    assert.equal(paragraphText(root.children[0]), "standalone");
});

// --- heading paths ----------------------------------------------------------

test("a heading's own path is its ancestors, not itself", () => {
    const result = map([
        paragraph("Chapter", { styleId: "berschrift1" }),
        paragraph("Section", { styleId: "berschrift2" }),
        paragraph("Body", { styleId: "Standard" }),
    ]);
    assert.deepEqual(result.paragraphs[0].headingPath, []);
    assert.deepEqual(result.paragraphs[1].headingPath, ["Chapter"]);
    assert.deepEqual(result.paragraphs[2].headingPath, ["Chapter", "Section"]);
});

test("a new heading at the same level replaces its predecessor's subtree", () => {
    const result = map([
        paragraph("One", { styleId: "berschrift1" }),
        paragraph("Sub", { styleId: "berschrift2" }),
        paragraph("Two", { styleId: "berschrift1" }),
        paragraph("Body", { styleId: "Standard" }),
    ]);
    assert.deepEqual(result.paragraphs[3].headingPath, ["Two"]);
});

test("a skipped heading level is padded so a path index still maps to a heading level", () => {
    const result = map([
        paragraph("Chapter", { styleId: "berschrift1" }),
        paragraph("Deep", { styleId: "berschrift1", outlineLevel: 2 }),
        paragraph("Body"),
    ]);
    // The heading's own path is the ancestors that actually exist -- padding is
    // not invented for a paragraph that has no missing ancestor above it.
    assert.equal(result.paragraphs[1].headingLevel, 3);
    assert.deepEqual(result.paragraphs[1].headingPath, ["Chapter"]);
    // Below it, the absent level 2 is padded, so index 2 is still the level-3
    // heading rather than sliding up to where the level-2 heading would be.
    assert.deepEqual(result.paragraphs[2].headingPath, ["Chapter", "", "Deep"]);
});

test("headingCount counts headings and paragraphCount counts everything", () => {
    const result = map([
        paragraph("Chapter", { styleId: "berschrift1" }),
        paragraph("Body"),
        paragraph(""),
    ]);
    assert.equal(result.paragraphCount, 3);
    assert.equal(result.headingCount, 1);
});

// --- addresses and the occurrence index -------------------------------------
//
// The reason the occurrence index is kept even when text looks unique: the
// document that proved the scheme had no duplicate paragraph text at all, so
// uniqueness was never actually tested. These fixtures test it.

test("duplicate text under the same heading is disambiguated by the occurrence index", () => {
    const result = map([
        paragraph("Chapter", { styleId: "berschrift1" }),
        paragraph("See the appendix."),
        paragraph("Something else."),
        paragraph("See the appendix."),
    ]);
    const [, first, , second] = result.paragraphs;
    assert.equal(first.text, second.text);
    assert.deepEqual(first.headingPath, second.headingPath);
    assert.equal(first.occurrence, 1);
    assert.equal(second.occurrence, 2);
    assert.notEqual(first.address, second.address, "identical text under one heading must not share an address");
});

test("every address in a document with heavy duplication is unique", () => {
    const body = [paragraph("Chapter", { styleId: "berschrift1" })];
    for (let i = 0; i < 25; i++) body.push(paragraph("Repeated line."), paragraph(""));
    const result = map(body);
    assert.equal(new Set(addressesOf(result)).size, result.paragraphCount);
});

test("empty paragraphs are addressable and distinct from each other", () => {
    const result = map([paragraph(""), paragraph(""), paragraph("")]);
    assert.deepEqual(
        result.paragraphs.map((p) => p.occurrence),
        [1, 2, 3],
    );
    assert.equal(new Set(addressesOf(result)).size, 3);
});

test("the same text under different headings gets different addresses at occurrence 1", () => {
    const result = map([
        paragraph("Alpha", { styleId: "berschrift1" }),
        paragraph("Shared line."),
        paragraph("Beta", { styleId: "berschrift1" }),
        paragraph("Shared line."),
    ]);
    const [, underAlpha, , underBeta] = result.paragraphs;
    assert.equal(underAlpha.occurrence, 1);
    assert.equal(underBeta.occurrence, 1, "the occurrence counter is scoped to the heading path");
    assert.notEqual(underAlpha.address, underBeta.address);
});

test("addresses are deterministic across separate reads of identical markup", () => {
    const body = [paragraph("Chapter", { styleId: "berschrift1" }), paragraph("Body."), paragraph("Body.")];
    assert.deepEqual(addressesOf(map(body)), addressesOf(map(body)));
});

test("an address is stable under invisible run splitting", () => {
    const whole = map([paragraph("Hello world")]);
    const split = map([paragraph(null, { raw: `<w:r><w:t xml:space="preserve">Hello </w:t></w:r><w:r><w:t>world</w:t></w:r>` })]);
    assert.equal(whole.paragraphs[0].address, split.paragraphs[0].address);
});

test("an address moves when the text actually changes", () => {
    assert.notEqual(map([paragraph("before")]).paragraphs[0].address, map([paragraph("after")]).paragraphs[0].address);
});

test("mintAddress is prefixed, fixed-length and total on empty input", () => {
    const address = mintAddress({});
    assert.match(address, /^p:[0-9a-f]{12}$/);
    assert.equal(mintAddress({ headingPath: [], text: "", occurrence: 1 }), address);
});

test("mintAddress separates its fields so a shifted boundary is a different address", () => {
    // "ab" + "" must not collide with "a" + "b".
    assert.notEqual(mintAddress({ headingPath: ["ab"], text: "" }), mintAddress({ headingPath: ["a"], text: "b" }));
    assert.notEqual(mintAddress({ headingPath: ["a", "b"] }), mintAddress({ headingPath: ["a b"] }));
});

test("normalizeText collapses the whitespace an address must not depend on", () => {
    assert.equal(normalizeText("  a \t b \n c  "), "a b c");
    assert.equal(normalizeText(null), "");
    assert.equal(normalizeText(undefined), "");
});

// --- tables, lists, paging --------------------------------------------------

test("paragraphs inside a table are included and flagged", () => {
    const result = map([paragraph("Before"), table(["Cell one", "Cell two"]), paragraph("After")]);
    assert.deepEqual(textsOf(result), ["Before", "Cell one", "Cell two", "After"]);
    assert.deepEqual(
        result.paragraphs.map((p) => p.inTable),
        [false, true, true, false],
    );
});

test("numbering is reported so a list item is recognisable", () => {
    const result = map([paragraph("Item", { numbering: { level: 1, numId: "3" } }), paragraph("Plain")]);
    assert.deepEqual(result.paragraphs[0].list, { numId: "3", level: 1 });
    assert.equal(result.paragraphs[1].list, null);
});

test("index is 1-based and continuous across tables", () => {
    const result = map([paragraph("a"), table(["b"]), paragraph("c")]);
    assert.deepEqual(
        result.paragraphs.map((p) => p.index),
        [1, 2, 3],
    );
});

test("paging returns a window without changing any address", () => {
    const body = Array.from({ length: 10 }, (_, i) => paragraph(`Line ${i}`));
    const all = map(body);
    const page = map(body, { limit: 3, offset: 4 });

    assert.equal(page.paragraphCount, 10, "the count describes the document, not the page");
    assert.equal(page.returned, 3);
    assert.equal(page.offset, 4);
    assert.equal(page.truncated, true);
    assert.deepEqual(addressesOf(page), addressesOf(all).slice(4, 7));
    assert.deepEqual(
        page.paragraphs.map((p) => p.index),
        [5, 6, 7],
    );
});

test("a page that reaches the end is not truncated", () => {
    const body = Array.from({ length: 4 }, (_, i) => paragraph(`Line ${i}`));
    assert.equal(map(body, { limit: 10, offset: 2 }).truncated, false);
    assert.equal(map(body).truncated, false);
});

test("an offset past the end returns nothing rather than throwing", () => {
    const result = map([paragraph("only")], { offset: 99 });
    assert.equal(result.returned, 0);
    assert.deepEqual(result.paragraphs, []);
});

// --- input handling ---------------------------------------------------------

test("a bare w:document is accepted as well as a Flat OPC package", () => {
    const result = buildStructureMap(bareDocument([paragraph("Plain")]));
    assert.equal(result.paragraphCount, 1);
    assert.equal(result.styleCount, 0, "a bare document has no styles part");
});

test("markup with no body is a typed error", () => {
    assert.throws(
        () => buildStructureMap(`<w:styles xmlns:w="x"/>`),
        (err) => err instanceof StructureError && err.code === "no_document_body",
    );
});

test("the section properties trailing the body are not a paragraph", () => {
    assert.equal(map([paragraph("only")]).paragraphCount, 1);
});

// --- address stability under mutation ---------------------------------------
//
// The assumption every layer above this one depends on: an address minted by
// one read still names the same paragraph on the next read. What follows pins
// both halves of that -- the edits an address survives, and the edits it does
// not. The ones it does not survive are not silent: any edit moves the file's
// revision token, so the agent is told to re-read rather than left addressing a
// document that has moved underneath it.

const H1 = (text) => paragraph(text, { styleId: "berschrift1" });

// Deliberately duplicate-heavy: "Boilerplate." appears under two headings and
// twice under one of them, which is the case demo.docx never had.
const BASE = [
    H1("Intro"),
    paragraph("Boilerplate."),
    paragraph("Only in the intro."),
    H1("Terms"),
    paragraph("Boilerplate."),
    paragraph("Boilerplate."),
    paragraph("Unique tail."),
];

/** address -> "headingPath|text#occurrence", for comparing two reads. */
const addressBook = (body) =>
    new Map(map(body).paragraphs.map((p) => [p.address, `${p.headingPath.join(">")}|${p.text}#${p.occurrence}`]));

/** Addresses present in `before` that are gone in `after`. */
const lost = (before, after) => [...before.keys()].filter((a) => !after.has(a));

test("appending a paragraph at the end leaves every existing address alone", () => {
    const after = addressBook([...BASE, paragraph("Appended.")]);
    assert.deepEqual(lost(addressBook(BASE), after), []);
});

test("an edit in one section does not move addresses in another", () => {
    // No positional index anywhere in the key, so an insertion above cannot
    // shift what is below -- the failure mode that ruled out paragraph numbers.
    const before = addressBook(BASE);
    const after = addressBook([
        H1("Intro"),
        paragraph("Boilerplate."),
        paragraph("Inserted right here."),
        paragraph("Only in the intro."),
        H1("Terms"),
        paragraph("Boilerplate."),
        paragraph("Boilerplate."),
        paragraph("Unique tail."),
    ]);
    assert.deepEqual(lost(before, after), [], "an insertion under Intro moved an address under Terms");
});

test("rewriting one paragraph moves only its own address", () => {
    const before = addressBook(BASE);
    const after = addressBook(BASE.map((p) => (p === BASE[6] ? paragraph("Rewritten tail.") : p)));
    const gone = lost(before, after).map((a) => before.get(a));
    assert.deepEqual(gone, ["Terms|Unique tail.#1"]);
});

test("reordering paragraphs within a section leaves their addresses alone", () => {
    // Position is not part of the key, so a paragraph that moves within its
    // section keeps its address. An edit citing it still lands correctly.
    const before = addressBook(BASE);
    const after = addressBook([
        H1("Intro"),
        paragraph("Only in the intro."),
        paragraph("Boilerplate."),
        H1("Terms"),
        paragraph("Unique tail."),
        paragraph("Boilerplate."),
        paragraph("Boilerplate."),
    ]);
    assert.deepEqual(lost(before, after), []);
});

test("appending to a duplicate group leaves the existing members alone", () => {
    // Growth at the end of a run of identical text is safe: occurrences 1 and 2
    // keep their addresses and the newcomer takes 3.
    const before = addressBook(BASE);
    const after = addressBook([...BASE.slice(0, 6), paragraph("Boilerplate."), ...BASE.slice(6)]);
    assert.deepEqual(lost(before, after), []);
    assert.equal([...after.values()].filter((v) => v.startsWith("Terms|Boilerplate.")).length, 3);
});

test("KNOWN LIMIT: deleting a member of a duplicate group renumbers the ones after it", () => {
    // The cost of an occurrence index. Deleting the *first* "Boilerplate."
    // under Terms makes the second one occurrence 1, so its address changes
    // even though its text and heading did not.
    //
    // This is safe rather than silent: the deletion moved the file's revision
    // token, so an edit citing the stale address is refused and the agent
    // re-reads. It is not safe to cache an address across an edit to a
    // duplicate group, and layer 2 should not try.
    const before = addressBook(BASE);
    const after = addressBook(BASE.filter((_, i) => i !== 4));
    assert.deepEqual(
        lost(before, after).map((a) => before.get(a)),
        ["Terms|Boilerplate.#2"],
        "expected exactly the trailing duplicate to be renumbered",
    );
});

test("KNOWN LIMIT: renaming a heading moves every address beneath it", () => {
    // The heading path is part of the key, which is what makes an address
    // readable and scopes the occurrence counter -- and it means a retitled
    // section invalidates its contents. Addresses outside it are untouched.
    const before = addressBook(BASE);
    const after = addressBook(BASE.map((p) => (p === BASE[3] ? H1("Conditions") : p)));
    assert.deepEqual(
        lost(before, after)
            .map((a) => before.get(a))
            .sort(),
        ["Terms|Boilerplate.#1", "Terms|Boilerplate.#2", "Terms|Unique tail.#1", "|Terms#1"].sort(),
        "the retitled heading and its contents, and nothing else",
    );
});

test("an address does not depend on the paragraphs Word split it into", () => {
    // Word re-splits runs on save for reasons invisible to the reader. If an
    // address moved with them, every address would expire on every save.
    const split = [
        H1("Intro"),
        paragraph(null, { raw: `<w:r><w:t xml:space="preserve">Boiler</w:t></w:r><w:r><w:t>plate.</w:t></w:r>` }),
        paragraph("Only in the intro."),
        ...BASE.slice(3),
    ];
    assert.deepEqual(lost(addressBook(BASE), addressBook(split)), []);
});
