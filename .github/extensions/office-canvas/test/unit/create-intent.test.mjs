// Office-free tests for the create_document spec validator.
//
// The spec is the only thing standing between an agent's JSON and a document
// Word will author. Everything it lets through becomes a file on disk, and the
// failure this file exists to prevent is the quiet one: a spec that is *partly*
// understood, producing a document that differs from what was asked for with
// nothing anywhere reporting it. So the validator refuses rather than coerces,
// and these tests are mostly about what it refuses.

import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
    BLOCK_HELP,
    BLOCK_KINDS,
    BLOCKS,
    fieldUsage,
    MAX_BLOCKS,
    MAX_HEADING_LEVEL,
    MAX_LIST_ITEMS,
    MAX_TABLE_COLUMNS,
    MAX_TABLE_ROWS,
    MAX_TEXT_LENGTH,
    MIN_BLOCK_HEADING_LEVEL,
    validateSpec,
} from "../../src/word/create-intent.mjs";
import * as created from "../../src/word/create-intent.mjs";
import { MIN_HEADING_LEVEL } from "../../src/word/edit-intent.mjs";

const rejects = (spec, code) => {
    const err = (() => {
        try {
            validateSpec(spec);
            return null;
        } catch (e) {
            return e;
        }
    })();
    assert.ok(err, `expected ${code} but the spec was accepted: ${JSON.stringify(spec)}`);
    assert.equal(err.code, code, `wrong code for ${JSON.stringify(spec)}: ${err.message}`);
    return err;
};

const para = (text = "body") => ({ kind: "paragraph", text });

test("a spec of every block kind round-trips normalized", () => {
    const spec = validateSpec({
        blocks: [
            { kind: "heading", level: 1, text: "Title" },
            { kind: "paragraph", text: "Body." },
            { kind: "list", items: ["one", "two"] },
            { kind: "table", rows: [["a", "b"]] },
        ],
    });

    // Optional flags are materialized rather than left undefined, so the host
    // never has to decide what an absent field means.
    assert.equal(spec.blocks[2].ordered, false);
    assert.equal(spec.blocks[3].headerRow, false);
    assert.deepEqual(spec.blocks[3].rows, [["a", "b"]]);
});

test("every kind in BLOCKS is implemented", () => {
    // The `default` arm of normalizeBlock is unreachable by design. This is what
    // makes that true: adding a key to BLOCKS without teaching normalizeBlock
    // about it fails here rather than silently dropping the block from an
    // authored document.
    for (const kind of BLOCK_KINDS) {
        const minimal = {
            heading: { kind, level: MIN_BLOCK_HEADING_LEVEL, text: "x" },
            paragraph: { kind, text: "x" },
            list: { kind, items: ["x"] },
            table: { kind, rows: [["x"]] },
        }[kind];
        assert.ok(minimal, `BLOCKS declares '${kind}' but this test has no example of it`);
        const spec = validateSpec({ blocks: [minimal] });
        assert.equal(spec.blocks[0].kind, kind);
    }
});

test("an unknown kind is refused and the message lists the real ones", () => {
    const err = rejects({ blocks: [{ kind: "quote", text: "x" }] }, "unknown_block_kind");
    for (const kind of BLOCK_KINDS) {
        assert.ok(err.message.includes(kind), `the message does not mention '${kind}'`);
    }
});

test("a field that belongs to another kind is refused, not ignored", () => {
    // The important half. Accepting `level` on a paragraph and dropping it would
    // produce body text where a heading was asked for, with no error raised --
    // which is the failure this whole layer exists to make impossible.
    rejects({ blocks: [{ kind: "paragraph", text: "x", level: 2 }] }, "invalid_block");
    rejects({ blocks: [{ kind: "list", items: ["x"], headerRow: true }] }, "invalid_block");
    rejects({ blocks: [{ kind: "heading", level: 1, text: "x", ordered: true }] }, "invalid_block");
});

test("a missing required field is refused", () => {
    for (const kind of BLOCK_KINDS) {
        for (const field of BLOCKS[kind].required) {
            const block = {
                heading: { kind, level: 1, text: "x" },
                paragraph: { kind, text: "x" },
                list: { kind, items: ["x"] },
                table: { kind, rows: [["x"]] },
            }[kind];
            delete block[field];
            const err = rejects({ blocks: [block] }, "invalid_block");
            assert.ok(err.message.includes(field), `the message does not name the missing field '${field}'`);
        }
    }
});

test("heading levels are bounded by the same constants edit_document uses", () => {
    assert.equal(MIN_BLOCK_HEADING_LEVEL, MIN_HEADING_LEVEL + 1);
    rejects({ blocks: [{ kind: "heading", level: MIN_HEADING_LEVEL, text: "x" }] }, "invalid_heading_level");
    rejects({ blocks: [{ kind: "heading", level: MAX_HEADING_LEVEL + 1, text: "x" }] }, "invalid_heading_level");
    rejects({ blocks: [{ kind: "heading", level: 1.5, text: "x" }] }, "invalid_heading_level");
    assert.equal(validateSpec({ blocks: [{ kind: "heading", level: MAX_HEADING_LEVEL, text: "x" }] }).blocks[0].level, MAX_HEADING_LEVEL);
});

test("a line break in any text is refused, wherever it appears", () => {
    // A paragraph break inside a block's text turns one paragraph into several,
    // which shifts every address after it. The rule is imported from
    // edit-intent.mjs rather than restated, and it has to reach the nested text
    // too -- a list item and a table cell are paragraphs like any other.
    rejects({ blocks: [{ kind: "paragraph", text: "a\nb" }] }, "invalid_text");
    rejects({ blocks: [{ kind: "heading", level: 1, text: "a\r\nb" }] }, "invalid_text");
    rejects({ blocks: [{ kind: "list", items: ["ok", "a\nb"] }] }, "invalid_text");
    rejects({ blocks: [{ kind: "table", rows: [["ok", "a\nb"]] }] }, "invalid_text");
});

test("a refusal says which block it was", () => {
    // A 500-block spec with one bad line is unactionable without the position.
    const err = rejects({ blocks: [para(), para(), { kind: "paragraph", text: "a\nb" }] }, "invalid_text");
    assert.match(err.message, /block 2/, `the message does not locate the block: ${err.message}`);
});

test("the text length bound comes from edit-intent, not a second copy", () => {
    const long = "x".repeat(MAX_TEXT_LENGTH + 1);
    rejects({ blocks: [{ kind: "paragraph", text: long }] }, "invalid_text");
    rejects({ blocks: [{ kind: "table", rows: [[long]] }] }, "invalid_text");
    assert.equal(validateSpec({ blocks: [para("x".repeat(MAX_TEXT_LENGTH))] }).blocks[0].text.length, MAX_TEXT_LENGTH);
});

test("a ragged table is refused rather than squared off", () => {
    // Word's Tables.Add takes one row count and one column count, so a ragged
    // spec has no faithful rendering. Padding it would produce a table the
    // caller did not ask for and cannot tell apart from one it did.
    const err = rejects({ blocks: [{ kind: "table", rows: [["a", "b"], ["c"]] }] }, "invalid_block");
    assert.match(err.message, /rectangular/);
});

test("empty collections are refused", () => {
    rejects({ blocks: [] }, "invalid_spec");
    rejects({ blocks: [{ kind: "list", items: [] }] }, "invalid_block");
    rejects({ blocks: [{ kind: "table", rows: [] }] }, "invalid_block");
    rejects({ blocks: [{ kind: "table", rows: [[]] }] }, "invalid_block");
});

test("the size ceilings are enforced at exactly the exported values", () => {
    // Each of these is also what the tool schema declares. The pairing is the
    // point: a declared bound the runtime does not enforce is a promise the
    // contract cannot keep, and an enforced bound the schema does not declare is
    // a refusal the model cannot anticipate.
    assert.ok(validateSpec({ blocks: Array.from({ length: MAX_BLOCKS }, () => para()) }));
    rejects({ blocks: Array.from({ length: MAX_BLOCKS + 1 }, () => para()) }, "invalid_spec");

    const items = (n) => ({ kind: "list", items: Array.from({ length: n }, (_, i) => `item ${i}`) });
    assert.ok(validateSpec({ blocks: [items(MAX_LIST_ITEMS)] }));
    rejects({ blocks: [items(MAX_LIST_ITEMS + 1)] }, "invalid_block");

    const rows = (r, c) => ({ kind: "table", rows: Array.from({ length: r }, () => Array.from({ length: c }, () => "x")) });
    assert.ok(validateSpec({ blocks: [rows(MAX_TABLE_ROWS, MAX_TABLE_COLUMNS)] }));
    rejects({ blocks: [rows(MAX_TABLE_ROWS + 1, 1)] }, "invalid_block");
    rejects({ blocks: [rows(1, MAX_TABLE_COLUMNS + 1)] }, "invalid_block");
});

test("a spec-level field nobody implements is refused, not dropped", () => {
    // `title` was in an earlier draft of this module, accepted by the validator
    // and ignored by the host -- a spec field that vanished silently, which is
    // the same defect as a half-applied block. Refusing is what makes "accepted"
    // mean "authored".
    rejects({ blocks: [para()], title: "Report" }, "invalid_spec");
    rejects({ blocks: [para()], footer: "x" }, "invalid_spec");
});

test("no predicted paragraph count is reported", () => {
    // A prediction used to be returned and it was wrong by construction: it
    // counted paragraphs the way Word's COM object model does (a row-end mark
    // per table row) while every number the caller can act on is OOXML-derived,
    // where row-end marks are not paragraphs. Measured on the smoke fixture: the
    // prediction said 18, the document says 14.
    //
    // This is a pin, not a preference. Re-exporting a `paragraphsIn` would put
    // two coordinate systems back in one result object with nothing marking
    // which is which.
    assert.equal(Object.keys(created).includes("paragraphsIn"), false);
    assert.equal(typeof created.validateSpec, "function");
});

// --- derive, don't restate ---------------------------------------------------
//
// The rule this repo has broken in three consecutive pull requests: a
// description, help string or comment writes down a bound a constant already
// defines, the constant moves, and the model is told something false with
// nothing able to go red.
//
// L1 wrote a test for exactly this that was a tautology -- it compared the
// rendered string against the constant, which is true whether the string derives
// or is a coincidentally-correct literal. So these look for the *defect*: a
// numeric literal where an interpolation belongs.

const here = path.dirname(fileURLToPath(import.meta.url));
const extensionSource = await readFile(path.join(here, "..", "..", "extension.mjs"), "utf8");

const createToolSource = () => {
    const start = extensionSource.indexOf("const createBlockSchema");
    const end = extensionSource.indexOf("\nconst editDocumentTool", start);
    assert.ok(start > 0 && end > start, "could not find the create_document tool in extension.mjs");
    return extensionSource.slice(start, end);
};

test("BLOCK_HELP names every kind and renders the heading range from the constants", async () => {
    for (const kind of BLOCK_KINDS) assert.ok(BLOCK_HELP.includes(kind), `BLOCK_HELP omits '${kind}'`);
    assert.ok(
        BLOCK_HELP.includes(`${MIN_BLOCK_HEADING_LEVEL}–${MAX_HEADING_LEVEL}`),
        `BLOCK_HELP does not render the heading range: ${BLOCK_HELP}`,
    );

    // The line above is true of a hardcoded "1–9" as well, which is precisely
    // the tautology L1 shipped: an assertion that passes with the defect
    // reinstated. So also look at how the string is built, where a literal is
    // visible and an interpolation is not.
    const source = await readFile(path.join(here, "..", "..", "src", "word", "create-intent.mjs"), "utf8");
    const helpSource = source.slice(source.indexOf("export const BLOCK_HELP"), source.indexOf('].join(" ")', source.indexOf("export const BLOCK_HELP")));
    assert.doesNotMatch(helpSource, /\d/, `BLOCK_HELP must interpolate its bounds, not write them down: ${helpSource}`);
});

test("the create_document schema declares no bound as a literal", () => {
    // Scoped to this tool's source rather than the whole file: other schemas
    // have their own bounds and their own tests, and a file-wide rule would
    // claim a surface this module does not own.
    //
    // This used to exempt 0 and 1, on the reasoning that a floor of one is the
    // definition of non-empty rather than a tunable bound. A review found the
    // cost of that: `items`, `rows` and each row carried no `minItems` at all
    // while the validator refused every empty one, and the exemption is why no
    // assertion here could notice. The floors are constants now
    // (MIN_LIST_ITEMS and friends) and the exemption is gone, so the schema and
    // the validator read one symbol per bound in both directions.
    const literals = [...createToolSource().matchAll(/\b(?:minimum|maximum|maxItems|minItems):\s*(-?\d+)/g)];
    const offending = literals.map((m) => m[0]);
    assert.deepEqual(
        offending,
        [],
        `bounds must be interpolated from src/word/create-intent.mjs, not written down: ${offending.join(", ")}`,
    );
});

test("every floor the validator enforces is also declared by the schema", () => {
    // The defect this pins, in the shape it shipped in: schema-valid and always
    // rejected at runtime. A model cannot learn a rule the schema does not
    // state; it only ever meets it as a failure. So each floor is checked on
    // both sides -- enforced at exactly the exported value, and declared.
    rejects({ blocks: [] }, "invalid_spec");
    rejects({ blocks: [{ kind: "list", items: [] }] }, "invalid_block");
    rejects({ blocks: [{ kind: "table", rows: [] }] }, "invalid_block");
    rejects({ blocks: [{ kind: "table", rows: [[]] }] }, "invalid_block");

    const source = createToolSource();
    for (const name of ["MIN_BLOCKS", "MIN_LIST_ITEMS", "MIN_TABLE_ROWS", "MIN_TABLE_COLUMNS"]) {
        assert.match(source, new RegExp(`minItems:\\s*${name}\\b`), `the schema does not declare ${name}`);
    }
});

test("the schema's per-kind clauses are derived from BLOCKS, not written alongside them", () => {
    // `required` in the schema holds only `kind`, because JSON Schema cannot
    // make it conditional on `kind` here. That pushes the per-kind rule into
    // each field's description -- which was hand-written ("heading only:",
    // "list only:") and so was a second copy of BLOCKS in the one place no test
    // looked and the model did.
    const source = createToolSource();
    for (const field of ["level", "text", "ordered", "items", "rows", "headerRow"]) {
        assert.match(source, new RegExp(`\\$\\{fieldUsage\\("${field}"\\)\\}`), `${field}'s description does not derive its kinds`);
    }
    assert.doesNotMatch(source, /\b(?:heading|paragraph|list|table) only:/, "a per-kind clause is spelled out rather than derived");

    // ...and the derivation actually says the right thing, since the two
    // assertions above are equally true of a fieldUsage that returned "".
    assert.match(fieldUsage("level"), /required by heading\./);
    assert.match(fieldUsage("text"), /required by heading and paragraph\./);
    assert.match(fieldUsage("ordered"), /optional on list\./);
    assert.match(fieldUsage("headerRow"), /optional on table\./);
});

test("fieldUsage refuses a field no kind takes", () => {
    // The teeth on the derivation. Renaming a field in BLOCKS without renaming
    // it in the schema now fails at extension load, rather than shipping a
    // schema that advertises a field the validator refuses as unsupported.
    assert.throws(() => fieldUsage("caption"), /no block kind takes/);
});

test("the create_document description does not claim autocorrect is always off", () => {
    // Suppression is conditional and deliberately so: on a Word this tool
    // started it is switched off, but on one the host attached to it is left
    // alone, because those are the user's settings and AutoCorrect is
    // per-application (measured -- toggled on a hidden instance and read back
    // as unchanged from a second process). The description asserted the
    // unconditional version, which is a promise the code does not keep on the
    // attached path. It must point at the reported outcome instead.
    const source = createToolSource();
    assert.match(source, /autoCorrect/, "the description does not name the field reporting what happened");
    assert.doesNotMatch(
        source,
        /autocorrect is switched off on the instance/i,
        "the description asserts suppression unconditionally",
    );
});

test("the create_document schema derives its block kinds from BLOCK_KINDS", () => {
    assert.match(createToolSource(), /enum:\s*BLOCK_KINDS/, "the kind enum must come from the constant");
    // ...and the prose alongside it, which is the copy that actually drifted the
    // last three times, because a schema `enum` is at least machine-checked.
    assert.match(createToolSource(), /\$\{BLOCK_HELP\}/, "the block help must be interpolated, not spelled out");
    for (const kind of BLOCK_KINDS) {
        assert.doesNotMatch(
            createToolSource(),
            new RegExp(`enum:\\s*\\[[^\\]]*"${kind}"`),
            "the kind list must not be spelled out in the schema",
        );
    }
});

test("the create_document path description derives the creatable extensions", () => {
    assert.match(createToolSource(), /\$\{creatableList\(\)\}/, "the extension list must be interpolated");
    assert.doesNotMatch(createToolSource(), /\.docx\b/, "the extension list must not be spelled out in the description");
});
