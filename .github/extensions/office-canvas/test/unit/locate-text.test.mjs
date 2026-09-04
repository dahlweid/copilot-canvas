// The change overlay's locator: given the text items pdf.js extracts from a
// page, which items cover the paragraph the last edit left behind?
//
// Office-free by construction. No PDF fixture can live in the extension folder
// -- it is binary, so C4 refuses it, and importing one from `spikes/` would
// break C3 -- which is what forced the locator to be a pure function over item
// arrays. The items here are the shape `page.getTextContent()` returns: `str`
// plus `hasEOL`.

import { test } from "node:test";
import assert from "node:assert/strict";

import { buildPageText, findOccurrences, locateText, MIN_PARTIAL_CHARS, normalizeText } from "../../src/ui/locate-text.mjs";
import { normalizeText as mapNormalizeText } from "../../src/word/structure-map.mjs";

/** `getTextContent()` item shape, with the two fields the locator reads. */
const item = (str, hasEOL = false) => ({ str, hasEOL, transform: [1, 0, 0, 1, 0, 0], width: str.length, height: 12 });

/** A page whose every item ends a line, which is what a Word export produces. */
const lines = (...strs) => strs.map((str) => item(str, true));

test("the locator normalizes text exactly as the structure map does", () => {
    // Derived, not restated. One side normalizes the text stored in the change
    // record and the other normalizes the page; if they ever diverge the overlay
    // silently stops appearing, with no error anywhere.
    const corpus = [
        "plain",
        "  leading and trailing  ",
        "internal   run",
        "tab\tseparated",
        "line\nbreak",
        "crlf\r\nbreak",
        "non\u00a0breaking",
        "\u2003em space",
        "",
        "   ",
        null,
        undefined,
        42,
    ];
    for (const value of corpus) {
        assert.equal(normalizeText(value), mapNormalizeText(value), `disagreement on ${JSON.stringify(value)}`);
    }
});

test("a paragraph on one line is located over exactly its item", () => {
    const items = lines("Chapter One", "The quick brown fox.", "Another line.");
    const found = locateText(items, "The quick brown fox.");
    assert.equal(found.status, "located");
    assert.deepEqual(found.range, { startItem: 1, endItem: 1 });
});

test("a match that starts and ends mid-item still names the covering items", () => {
    // Word splits a line into runs for reasons invisible to the reader, so a
    // paragraph routinely spans several items and rarely aligns to their edges.
    const items = [item("Intro. The quick"), item(" brown"), item(" fox. Outro."), item("", true)];
    const found = locateText(items, "The quick brown fox.");
    assert.equal(found.status, "located");
    assert.deepEqual(found.range, { startItem: 0, endItem: 2 });
});

test("a line break between items reads as a space, not as nothing", () => {
    // The discriminating case for hasEOL. Without it the page text is
    // "The quickbrown fox." and the paragraph is simply never found -- so the
    // overlay would degrade to a page marker on every multi-line paragraph and
    // look like a locator that merely works less often than hoped.
    const items = lines("The quick", "brown fox.");
    const found = locateText(items, "The quick brown fox.");
    assert.equal(found.status, "located");
    assert.deepEqual(found.range, { startItem: 0, endItem: 1 });
});

test("an end-of-line marker with no text of its own still separates the lines", () => {
    // pdf.js emits zero-width items carrying only `hasEOL`. Reading `str` and
    // ignoring `hasEOL` on those loses the break.
    const items = [item("The quick"), item("", true), item("brown fox.")];
    assert.equal(buildPageText(items).text, "The quick brown fox.");
});

test("the character-to-item map survives a character that is two UTF-16 units", () => {
    // The discriminating case, and the reason the ASCII fixture above cannot
    // stand in for it. `for...of` over a string yields *code points*, so an
    // emoji is one iteration but two units in `text`. Push one owner per
    // iteration and the map shifts left by one per astral character -- silently,
    // because `indexOf` counts units and the drift only shows up as a quad drawn
    // over the wrong item. An ASCII fixture cannot observe it: the two units of
    // measure coincide.
    const items = [item("\u{1F600}\u{1F600}\u{1F600} Titel", true), item("AAAA", true), item("Der Zielabsatz.", true)];
    const { text, owners } = buildPageText(items);
    assert.equal(owners.length, text.length, "one owner per UTF-16 unit, not per code point");

    const located = locateText(items, "AAAA");
    assert.equal(located.status, "located");
    assert.deepEqual(
        located.range,
        { startItem: 1, endItem: 1 },
        "the match is wholly inside item 1; a shifted map spills the range into item 2",
    );
});

test("the character-to-item map stays the same length as the text it describes", () => {
    // Every collapsed space, and the trailing space a final hasEOL leaves, must
    // drop an owner too. A drift here does not fail loudly: it silently shifts
    // every quad after the drift onto the wrong item.
    const items = [item("  padded   text  ", true), item("second line", true)];
    const { text, owners } = buildPageText(items);
    assert.equal(owners.length, text.length);
    assert.equal(text, "padded text second line");
    assert.ok(!text.endsWith(" "), "a trailing hasEOL must not leave a space on the end");
});

test("whitespace differences between the record and the page do not defeat a match", () => {
    const items = lines("The   quick\tbrown", "fox.");
    const found = locateText(items, "  The quick brown fox.  ");
    assert.equal(found.status, "located");
});

test("text that occurs twice on a page is refused, not guessed", () => {
    // The overlay must never assert a position the code did not determine. The
    // structure map's occurrence index counts within a heading, not within a
    // page, so there is nothing here to break the tie with.
    const items = lines("Total: 12", "Something else", "Total: 12");
    const found = locateText(items, "Total: 12");
    assert.equal(found.status, "ambiguous");
    assert.equal(found.occurrences, 2);
    assert.equal(found.range, null, "an ambiguous result must carry no range to draw");
});

test("overlapping occurrences count as ambiguity too", () => {
    const items = lines("aaaa");
    const found = locateText(items, "aaa");
    assert.equal(found.status, "ambiguous");
    assert.equal(found.occurrences, 2);
});

test("text that is not on the page is not found", () => {
    const found = locateText(lines("Chapter One", "Body text."), "A paragraph from somewhere else entirely.");
    assert.equal(found.status, "not_found");
    assert.equal(found.range, null);
});

test("a paragraph whose tail spilled onto this page matches at the page start", () => {
    const target = "This paragraph is long enough to cross a page break cleanly.";
    const tail = "enough to cross a page break cleanly.";
    const found = locateText(lines(tail, "The next paragraph."), target);
    assert.equal(found.status, "partial");
    assert.equal(found.where, "start");
    assert.deepEqual(found.range, { startItem: 0, endItem: 0 });
});

test("a paragraph whose head is on this page matches at the page end", () => {
    const target = "This paragraph is long enough to cross a page break cleanly.";
    const head = "This paragraph is long enough to cross";
    const found = locateText(lines("Earlier text.", head), target);
    assert.equal(found.status, "partial");
    assert.equal(found.where, "end");
    assert.deepEqual(found.range, { startItem: 1, endItem: 1 });
});

test("a partial match shorter than the confidence floor is refused", () => {
    // The discriminating case for the floor. "The paragraph begins" is a real
    // prefix of the target and sits at the very end of the page, so every
    // structural condition for a partial match holds -- only its length does
    // not. Without the floor this draws a box on a coincidence.
    const target = "The paragraph begins here and continues onto the following page.";
    const shortPrefix = target.slice(0, MIN_PARTIAL_CHARS - 1);
    assert.ok(shortPrefix.length > 0);
    const found = locateText(lines("Earlier text.", shortPrefix), target);
    assert.equal(found.status, "not_found");

    // ...and one character more is accepted, so the floor is the thing deciding.
    const atFloor = target.slice(0, MIN_PARTIAL_CHARS);
    const accepted = locateText(lines("Earlier text.", atFloor), target);
    assert.equal(accepted.status, "partial");
});

test("a partial match must sit at a page edge, not float in the middle", () => {
    // Both branches need their own fixture. A prefix floating in the middle is
    // refused by the `endsWith` branch and a suffix floating in the middle by
    // the `startsWith` branch -- so a single fixture leaves one of the two
    // guards untested, and a mutation of that branch survives.
    const target = "The paragraph begins here and continues onto the following page.";
    const prefix = target.slice(0, 40);
    const suffix = target.slice(-40);

    assert.equal(locateText(lines("Earlier text.", prefix, "Later text."), target).status, "not_found");
    assert.equal(locateText(lines("Earlier text.", suffix, "Later text."), target).status, "not_found");
});

// --- narrowing to the changed span (#166) -----------------------------------

test("a span narrows the box to the items holding the words that changed", () => {
    const items = lines("Chapter One", "The quick brown fox", "jumps over the lazy dog.", "Another paragraph.");
    const target = "The quick brown fox jumps over the lazy dog.";
    const wide = locateText(items, target);
    assert.deepEqual(wide.range, { startItem: 1, endItem: 2 }, "the unnarrowed match must span both lines");

    const found = locateText(items, target, { span: "lazy dog." });
    assert.equal(found.status, "located", "narrowing must not change what was matched");
    assert.equal(found.matched, target, "`matched` names the match, not the narrowing");
    assert.equal(found.narrowed, true);
    assert.deepEqual(found.range, { startItem: 2, endItem: 2 });
});

test("a span is searched only inside the match, never across the page", () => {
    // The discriminating case, and the reason the span is passed into the
    // locator rather than searched for on its own: "brown fox" also sits in an
    // unrelated paragraph earlier on the page. Searching the page for it would
    // be ambiguous at best and would box the wrong paragraph at worst.
    const items = lines("A brown fox appeared.", "The quick brown fox", "jumps over the lazy dog.");
    const found = locateText(items, "The quick brown fox jumps over the lazy dog.", { span: "brown fox" });
    assert.equal(found.status, "located");
    assert.equal(found.narrowed, true);
    assert.deepEqual(found.range, { startItem: 1, endItem: 1 }, "the box must be inside the located paragraph");
});

test("a span occurring twice inside the paragraph leaves the whole paragraph boxed", () => {
    // Ambiguity inside the match is the one case narrowing cannot resolve, and
    // the answer is the box that was there before -- never a first-match guess.
    const items = lines("Total: 12 and", "again Total: 12 here.");
    const target = "Total: 12 and again Total: 12 here.";
    const found = locateText(items, target, { span: "Total: 12" });
    assert.equal(found.status, "located");
    assert.equal(found.narrowed, false);
    assert.deepEqual(found.range, { startItem: 0, endItem: 1 });
});

test("a span that is not in the match at all leaves the match standing", () => {
    // Reachable whenever the page and the record disagree about the text --
    // a ligature, a smart quote, a glyph the export renders differently. The
    // paragraph still matched, so the reader still gets a box.
    const items = lines("The quick brown fox.");
    const found = locateText(items, "The quick brown fox.", { span: "silver badger" });
    assert.equal(found.status, "located");
    assert.equal(found.narrowed, false);
    assert.deepEqual(found.range, { startItem: 0, endItem: 0 });
});

test("a span equal to the whole match is not treated as a narrowing", () => {
    const items = lines("The quick brown fox.");
    const found = locateText(items, "The quick brown fox.", { span: "The quick brown fox." });
    assert.equal(found.narrowed, false);
});

test("no span leaves every result exactly as it was", () => {
    // The feature is additive. A record built before spans existed, or one whose
    // span could not be derived, must produce the pre-#166 answer.
    const items = lines("Chapter One", "The quick brown fox.");
    for (const span of [null, undefined, "", "   "]) {
        const found = locateText(items, "The quick brown fox.", { span });
        assert.equal(found.status, "located");
        assert.equal(found.narrowed, false, `${JSON.stringify(span)} should narrow nothing`);
        assert.deepEqual(found.range, { startItem: 1, endItem: 1 });
    }
});

test("a span is normalized before it is searched for", () => {
    const items = lines("The quick brown fox", "jumps over the lazy dog.");
    const found = locateText(items, "The quick brown fox jumps over the lazy dog.", { span: "  lazy   dog.  " });
    assert.equal(found.narrowed, true);
    assert.deepEqual(found.range, { startItem: 1, endItem: 1 });
});

test("a span narrows a page-straddling partial match too", () => {
    // A paragraph that crosses a page break still deserves the narrower box on
    // whichever page holds the changed words -- and must not claim one on the
    // page that does not.
    const target = "This paragraph is long enough to cross a page break cleanly and then some.";
    const head = "This paragraph is long enough to cross";
    const items = lines("Earlier text.", head);

    const inside = locateText(items, target, { span: "long enough" });
    assert.equal(inside.status, "partial");
    assert.equal(inside.narrowed, true);
    assert.deepEqual(inside.range, { startItem: 1, endItem: 1 });

    // The changed words are on the *other* page, so this page keeps the fragment
    // it matched rather than being narrowed to something it does not hold.
    const elsewhere = locateText(items, target, { span: "and then some." });
    assert.equal(elsewhere.status, "partial");
    assert.equal(elsewhere.narrowed, false);
});

test("an empty target is reported as empty rather than matching everything", () => {
    // "" is a substring of every string, so a locator that did not special-case
    // it would report a confident match at offset 0 on every page.
    for (const value of ["", "   ", null, undefined]) {
        const found = locateText(lines("Some text."), value);
        assert.equal(found.status, "empty", `${JSON.stringify(value)} should be empty`);
        assert.equal(found.range, null);
    }
});

test("a page with no text at all is not found rather than throwing", () => {
    for (const items of [[], [item("")], [item("   ", true)], null, undefined]) {
        assert.equal(locateText(items, "anything").status, "not_found");
    }
});

test("findOccurrences answers nothing for an empty needle", () => {
    assert.deepEqual(findOccurrences("abc", ""), []);
    assert.deepEqual(findOccurrences("abcabc", "abc"), [0, 3]);
});
