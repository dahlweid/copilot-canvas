// What the overlay is told about the operation that just ran.
//
// Office-free: the record is built from an `edit_document` result, which is a
// plain object. The results here are shaped from the ones `document-editor.mjs`
// returns, so a change to that shape shows up as a failure rather than as an
// overlay that quietly stops appearing.

import { test } from "node:test";
import assert from "node:assert/strict";

import { changedSpan, changeRecordFrom } from "../../src/change-record.mjs";
import { OPERATION_NAMES } from "../../src/word/edit-intent.mjs";

const CLOCK = () => "2024-01-01T00:00:00.000Z";

/** An `edit_document` result, with the fields the record reads. */
const result = (over = {}) => ({
    applied: { op: "replace_text", description: "Replaced the paragraph text", ...over.applied },
    page: 3,
    paragraph: { address: "p:0123456789ab", text: "The replacement text.", style: "Standard", ...over.paragraph },
    document: { revisionToken: "sha256:abcdef0123456789", ...over.document },
    ...over,
});

test("a record is built for every operation the editor can apply", () => {
    // Derived from the editor's own list rather than a copy of it. A new
    // operation arriving with no thought given to what the overlay should draw
    // fails here instead of silently producing a box on the wrong thing.
    for (const op of OPERATION_NAMES) {
        const record = changeRecordFrom(result({ applied: { op, description: op } }), { now: CLOCK });
        assert.ok(record, `${op} produced no record`);
        assert.equal(record.op, op);
    }
});

test("a replacement carries the post-edit text and the page Word reported", () => {
    const record = changeRecordFrom(result(), { now: CLOCK });
    assert.deepEqual(record, {
        op: "replace_text",
        page: 3,
        text: "The replacement text.",
        span: null,
        locatable: true,
        at: "2024-01-01T00:00:00.000Z",
    });
});

test("the record carries no description, because the editor's names an address", () => {
    // `applied.description` reads "replace the text of p:8f957157e47d". It is
    // written for the agent, and it was rendered into the change banner until a
    // run against a real document put the address on screen. The viewer words the
    // change from `op` instead, so the field must not travel -- a record that
    // carries it is one field away from being shown again.
    const record = changeRecordFrom(
        result({ applied: { op: "replace_text", description: "replace the text of p:8f957157e47d" } }),
        { now: CLOCK },
    );
    assert.equal(record.description, undefined);
    assert.ok(
        !JSON.stringify(record).includes("p:"),
        `the record carries an address-bearing string: ${JSON.stringify(record)}`,
    );
});

test("the record carries no address, because an address cannot survive the edit", () => {
    // ADR 0006: an address is a coordinate valid inside one read-then-edit
    // cycle. The result the record is built from *does* contain one, so this
    // asserts a deliberate omission rather than an absence that happened.
    const source = result();
    assert.ok(source.paragraph.address, "the source result must carry an address for this to mean anything");
    const record = changeRecordFrom(source, { now: CLOCK });
    assert.ok(!("address" in record), "the record must not carry an address");
    assert.ok(
        !JSON.stringify(record).includes(source.paragraph.address),
        "the address must not reach the viewer under any key",
    );
});

test("the record carries no revision token, which is not the invalidator", () => {
    // Invalidation is on the render key, which moves on any write. A token
    // beside it would be a second rule that could disagree -- a byte-identical
    // regeneration leaves the token alone while still producing a new render.
    const source = result();
    assert.ok(source.document.revisionToken, "the source result must carry a token for this to mean anything");
    const record = changeRecordFrom(source, { now: CLOCK });
    assert.ok(!JSON.stringify(record).includes(source.document.revisionToken));
});

test("a deletion is recorded but is not locatable", () => {
    // Nothing is left on the page to find. Anchoring to whatever now follows the
    // gap would assert a position the code never determined.
    const record = changeRecordFrom(
        result({ applied: { op: "delete_paragraph", description: "Deleted the paragraph" } }),
        { now: CLOCK },
    );
    assert.equal(record.op, "delete_paragraph");
    assert.equal(record.locatable, false);
    assert.equal(record.text, null, "an unlocatable record must carry no text to search for");
    assert.equal(record.page, 3, "the page is still known, so the marker can still be placed");
});

test("an empty paragraph is not locatable either", () => {
    // A heading demoted to body text can leave an empty paragraph, and the empty
    // string matches at offset zero of every page.
    for (const text of ["", "   ", "\n\t", null, undefined]) {
        const record = changeRecordFrom(result({ paragraph: { text } }), { now: CLOCK });
        assert.equal(record.locatable, false, `${JSON.stringify(text)} should not be locatable`);
        assert.equal(record.text, null);
    }
});

test("the recorded text is normalized, so it can match the page", () => {
    const record = changeRecordFrom(result({ paragraph: { text: "  The   replacement\ttext.  " } }), { now: CLOCK });
    assert.equal(record.text, "The replacement text.");
});

test("a page Word could not report leaves the page null rather than guessing 1", () => {
    // Page 1 is a real page. Defaulting to it would put a marker on a page the
    // edit may not have touched, and nothing downstream could tell the
    // difference between a known page 1 and an unknown one.
    for (const page of [null, undefined, 0, -2, Number.NaN, "3", Infinity]) {
        const record = changeRecordFrom(result({ page }), { now: CLOCK });
        assert.equal(record.page, null, `page ${JSON.stringify(page)} should not be trusted`);
    }
});

test("a fractional page is floored rather than passed through", () => {
    assert.equal(changeRecordFrom(result({ page: 3.7 }), { now: CLOCK }).page, 3);
});

test("a result that applied nothing produces no record at all", () => {
    // Returning a record the viewer would have to know to ignore is how an
    // overlay ends up drawn for an operation that never happened.
    for (const empty of [null, undefined, {}, { applied: null }, { applied: {} }, { applied: { op: "" } }]) {
        assert.equal(changeRecordFrom(empty, { now: CLOCK }), null, `${JSON.stringify(empty)} should produce no record`);
    }
});

test("the timestamp comes from the injected clock", () => {
    const record = changeRecordFrom(result(), { now: () => "2020-02-02T02:02:02.000Z" });
    assert.equal(record.at, "2020-02-02T02:02:02.000Z");
});

test("the default clock produces a parseable ISO timestamp", () => {
    const record = changeRecordFrom(result());
    assert.ok(Number.isFinite(Date.parse(record.at)));
});

// --- the changed span (#166) ------------------------------------------------
//
// The record narrows the highlight to the words that changed. Every way of
// failing to derive a span must land on `null`, which is the whole-paragraph
// highlight that was there before -- so each of these is a branch with a test,
// not a fallthrough.

/** A result carrying both sides of the edit, as `document-editor.mjs` returns. */
const replacement = (previousText, text) =>
    result({ previousText, paragraph: { text }, applied: { op: "replace_text", description: "replaced" } });

test("a replacement narrows to the words that changed", () => {
    // The issue's own case: a sentence inserted into a long paragraph. The
    // record still carries the whole paragraph as the join key.
    const before = "Die Architektur ist modular. Sie skaliert gut. Das Team ist klein.";
    const after = "Die Architektur ist modular. Hallo Markus. Sie skaliert gut. Das Team ist klein.";
    const record = changeRecordFrom(replacement(before, after), { now: CLOCK });
    assert.equal(record.text, after);
    assert.equal(record.span, "Hallo Markus.");
});

test("the span is widened to whole words, never left mid-word", () => {
    // A raw prefix/suffix diff of "GitHub" -> "GitLab" is "L", which is both
    // meaningless to a reader and far likelier to occur twice on the page than
    // the word is -- and occurring twice is what makes it unlocatable.
    const record = changeRecordFrom(replacement("We use GitHub daily.", "We use GitLab daily."), { now: CLOCK });
    assert.equal(record.span, "GitLab");
});

test("the span is derived from normalized text on both sides", () => {
    // The overlay searches a normalized page. A span carrying a tab could never
    // be found there, and the failure would be a highlight that silently never
    // narrows rather than an error.
    const record = changeRecordFrom(replacement("One\ttwo three.", "One  two   four.  "), { now: CLOCK });
    assert.equal(record.span, "four.");
    assert.ok(record.text.includes(record.span), "the span must be findable inside the record's own text");
});

test("a pure deletion inside a paragraph narrows to nothing", () => {
    // The words that changed were removed, so there is no text left on the page
    // to draw a box around. The paragraph is still locatable and still marked.
    const record = changeRecordFrom(replacement("One two three four.", "One four."), { now: CLOCK });
    assert.equal(record.span, null);
    assert.equal(record.locatable, true);
    assert.equal(record.text, "One four.");
});

test("a rewrite sharing no prefix or suffix carries no span", () => {
    // The derived span would be the whole paragraph, so carrying it would put a
    // second copy of the same string in the record and imply a narrowing that
    // did not happen.
    const record = changeRecordFrom(replacement("Alpha beta gamma.", "Delta epsilon zeta."), { now: CLOCK });
    assert.equal(record.span, null);
});

test("text equal after normalization carries no span", () => {
    // Word can rewrite a paragraph into something this normalization cannot
    // tell apart from the old one. There is no narrower box to draw honestly.
    const record = changeRecordFrom(replacement("The  same   text.", "The same text."), { now: CLOCK });
    assert.equal(record.span, null);
});

test("an edit with no prior text carries no span", () => {
    // The editor supplies `previousText` only where the paragraph Word touched
    // is the one the address named. Absence must mean "no narrowing", never
    // "everything changed".
    for (const previousText of [null, undefined, "", "   "]) {
        const record = changeRecordFrom(result({ previousText }), { now: CLOCK });
        assert.equal(record.span, null, `${JSON.stringify(previousText)} should derive no span`);
        assert.equal(record.text, "The replacement text.", "the paragraph is still the join key");
    }
});

test("an unlocatable record carries no span either", () => {
    // A deletion has no text on the page, so it cannot have a narrower part of
    // one. Deriving a span here would be a value nothing could ever use.
    const record = changeRecordFrom(
        result({
            previousText: "One two three four.",
            applied: { op: "delete_paragraph", description: "deleted" },
        }),
        { now: CLOCK },
    );
    assert.equal(record.locatable, false);
    assert.equal(record.span, null);
});

test("the span is always a substring of the text the record carries", () => {
    // The span is searched inside the paragraph's own match. A span that is not
    // part of that text could never be found there, so this is the invariant the
    // locator's narrowing depends on -- checked over the whole corpus rather
    // than at one example.
    const pairs = [
        ["One two three.", "One two and a half three."],
        ["One two three.", "One 2 three."],
        ["Ein Satz über Straßen.", "Ein Satz über Strassen."],
        ["a b c d e f g", "a b c X e f g"],
        ["Trailing words here.", "Trailing words here. And more."],
        ["Leading words here.", "More. Leading words here."],
    ];
    for (const [before, after] of pairs) {
        const record = changeRecordFrom(replacement(before, after), { now: CLOCK });
        if (record.span === null) continue;
        assert.ok(
            record.text.includes(record.span),
            `${JSON.stringify(record.span)} is not inside ${JSON.stringify(record.text)}`,
        );
        assert.notEqual(record.span, record.text, "a span equal to the whole text is not a narrowing");
    }
});

test("the span derivation is a pure function of the two strings", () => {
    // Exported so the branch table can be read without building a result around
    // each case, and so a caller cannot reach the narrowing without the two
    // texts it is derived from.
    assert.equal(changedSpan("", "anything"), null);
    assert.equal(changedSpan("anything", ""), null);
    assert.equal(changedSpan("same", "same"), null);
    assert.equal(changedSpan("One two three.", "One 2 three."), "2");
});

test("a shared boundary between prefix and suffix does not widen the span", () => {
    // Inserting a sentence puts ". " both at the end of the common prefix and
    // inside the common suffix. A diff that refuses that overlap -- as the
    // host's must, since it decomposes the paragraph to assign a range -- cuts
    // the suffix short and pulls the following word into the box. This one
    // names a span rather than a decomposition, so it does not have to.
    const before = "Die Architektur ist modular. Sie skaliert gut.";
    const after = "Die Architektur ist modular. Hallo Markus. Sie skaliert gut.";
    assert.equal(changedSpan(before, after), "Hallo Markus.");
});

test("every operation the editor can apply has wording of its own", async () => {
    // The viewer words a change from `op`. A table with a generic fallback means
    // a newly added operation reaches the reader as "Changed" and nothing goes
    // red -- so the check is against the editor's list, not against a copy.
    const { describeChange, DESCRIBED_OPS, GENERIC_PHRASE } = await import("../../src/ui/change-wording.mjs");

    const missing = OPERATION_NAMES.filter((op) => describeChange({ op }) === GENERIC_PHRASE);
    assert.deepEqual(missing, [], `these operations reach the reader as "${GENERIC_PHRASE}": ${missing.join(", ")}`);

    // And the other direction: wording for an operation that no longer exists is
    // dead text nobody will notice is dead.
    const orphaned = DESCRIBED_OPS.filter((op) => !OPERATION_NAMES.includes(op));
    assert.deepEqual(orphaned, [], `wording exists for operations the editor cannot apply: ${orphaned.join(", ")}`);
});
