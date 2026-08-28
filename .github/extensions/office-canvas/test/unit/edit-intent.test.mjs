// Office-free tests for the edit intent.
//
// The point of a closed operation set is that a bad intent is refused before a
// snapshot is taken and before Word is started, so these tests are the first
// line of the edit path even though they never touch a document.

import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
    describeIntent,
    EditIntentError,
    MAX_HEADING_LEVEL,
    MIN_HEADING_LEVEL,
    OPERATION_NAMES,
    validateIntent,
} from "../../src/word/edit-intent.mjs";
import { mintAddress } from "../../src/word/structure-map.mjs";

const ADDRESS = mintAddress({ headingPath: ["Intro"], text: "Hello world", occurrence: 1 });

const codeOf = (fn) => {
    try {
        fn();
    } catch (err) {
        assert.ok(err instanceof EditIntentError, `expected an EditIntentError, got ${err}`);
        return err.code;
    }
    return null;
};

test("accepts an address minted by mintAddress", () => {
    // Pins the validator's pattern to the real minting function. A change to
    // the address format that this file did not follow fails here rather than
    // at the first edit against a real document.
    const intent = validateIntent({ op: "replace_text", address: ADDRESS, text: "New text" });
    assert.equal(intent.address, ADDRESS);
    assert.equal(intent.op, "replace_text");
    assert.equal(intent.text, "New text");
});

test("every operation name is accepted with its required fields", () => {
    const required = {
        replace_text: { text: "x" },
        insert_paragraph_after: { text: "x" },
        insert_paragraph_before: { text: "x" },
        delete_paragraph: {},
        set_heading_level: { headingLevel: 2 },
    };
    for (const op of OPERATION_NAMES) {
        const intent = validateIntent({ op, address: ADDRESS, ...required[op] });
        assert.equal(intent.op, op);
        assert.ok(describeIntent(intent).length > 0);
    }
});

test("rejects an unknown operation", () => {
    assert.equal(codeOf(() => validateIntent({ op: "rewrite_everything", address: ADDRESS })), "unknown_operation");
    assert.equal(codeOf(() => validateIntent({ address: ADDRESS })), "unknown_operation");
});

test("rejects anything that is not a paragraph address", () => {
    const bad = ["", "p:123", "p:0123456789AB", "0123456789ab", "paragraph 4", null, 7, `${ADDRESS}x`];
    for (const address of bad) {
        assert.equal(
            codeOf(() => validateIntent({ op: "delete_paragraph", address })),
            "invalid_address",
            `expected ${JSON.stringify(address)} to be refused`,
        );
    }
});

test("refuses line breaks in paragraph text", () => {
    // A paragraph break inside the text silently turns one addressed paragraph
    // into two, moving every address after it. Refusing is the whole reason
    // this check exists.
    for (const text of ["one\ntwo", "one\r\ntwo", "one\vtwo", "one\ftwo"]) {
        assert.equal(codeOf(() => validateIntent({ op: "replace_text", address: ADDRESS, text })), "invalid_text");
    }
});

test("accepts tabs and unicode in paragraph text", () => {
    const text = "Spalte\tWert — Überschrift ✓";
    assert.equal(validateIntent({ op: "replace_text", address: ADDRESS, text }).text, text);
});

test("rejects missing or oversized text", () => {
    assert.equal(codeOf(() => validateIntent({ op: "replace_text", address: ADDRESS })), "invalid_text");
    assert.equal(codeOf(() => validateIntent({ op: "replace_text", address: ADDRESS, text: 42 })), "invalid_text");
    assert.equal(
        codeOf(() => validateIntent({ op: "replace_text", address: ADDRESS, text: "x".repeat(50_001) })),
        "invalid_text",
    );
});

test("rejects a field an operation does not take", () => {
    assert.equal(
        codeOf(() => validateIntent({ op: "delete_paragraph", address: ADDRESS, text: "x" })),
        "invalid_intent",
    );
    assert.equal(
        codeOf(() => validateIntent({ op: "replace_text", address: ADDRESS, text: "x", headingLevel: 1 })),
        "invalid_intent",
    );
    assert.equal(
        codeOf(() => validateIntent({ op: "delete_paragraph", address: ADDRESS, styleId: "berschrift1" })),
        "invalid_intent",
    );
});

test("heading level is 0 through 9, integers only", () => {
    for (const level of [0, 1, 5, 9]) {
        assert.equal(validateIntent({ op: "set_heading_level", address: ADDRESS, headingLevel: level }).headingLevel, level);
    }
    for (const level of [-1, 10, 1.5, "1", null, undefined]) {
        assert.equal(
            codeOf(() => validateIntent({ op: "set_heading_level", address: ADDRESS, headingLevel: level })),
            "invalid_heading_level",
            `expected level ${JSON.stringify(level)} to be refused`,
        );
    }
});

test("heading level is optional on an insert", () => {
    const inherited = validateIntent({ op: "insert_paragraph_after", address: ADDRESS, text: "x" });
    assert.equal("headingLevel" in inherited, false);

    const explicit = validateIntent({ op: "insert_paragraph_after", address: ADDRESS, text: "x", headingLevel: 3 });
    assert.equal(explicit.headingLevel, 3);
});

test("no operation accepts a style name or id", () => {
    // The write side must never name a style: `Range.Style` takes neither the
    // OOXML style id Word writes ("berschrift1") nor the English name
    // ("Heading 1"), and both throw on a localized Word. Only a numeric
    // headingLevel gets in.
    for (const op of OPERATION_NAMES) {
        for (const field of ["style", "styleId", "styleName"]) {
            assert.equal(
                codeOf(() => validateIntent({ op, address: ADDRESS, [field]: "berschrift1" })),
                "invalid_intent",
                `${op} must not accept ${field}`,
            );
        }
    }
});

test("rejects a non-object intent", () => {
    for (const input of [null, undefined, "replace_text", 3, []]) {
        assert.ok(["invalid_intent", "unknown_operation"].includes(codeOf(() => validateIntent(input))));
    }
});

// --- The declared bound and the enforced bound are one constant -------------
//
// L1 hit the drift version of this on `limit`: the tool schema advertised one
// bound while the runtime enforced another, so the agent was told a value was
// legal and then refused for using it. The fix there was to export the
// constant and consume it in the schema; these two tests are what keep that
// true here, and they are deliberately different in kind.

const here = path.dirname(fileURLToPath(import.meta.url));
const extensionSource = await readFile(path.join(here, "..", "..", "extension.mjs"), "utf8");

test("the schema's heading-level bound is derived, not restated", () => {
    // Asserting the rendered bound equals the constant would be a tautology
    // once it derives. So look for the defect itself -- a numeric literal in
    // the `headingLevel` block -- the way supported-types.test.mjs does.
    const block = extensionSource.match(/headingLevel: \{[\s\S]*?\n {12}\}/)?.[0];
    assert.ok(block, "could not find the headingLevel schema block in extension.mjs");

    assert.doesNotMatch(
        block,
        /(minimum|maximum): *\d/,
        "the headingLevel schema states a numeric bound instead of deriving it from edit-intent.mjs",
    );
    assert.match(block, /minimum: MIN_HEADING_LEVEL/);
    assert.match(block, /maximum: MAX_HEADING_LEVEL/);

    // The description is the part the model actually reads, and it carried its
    // own hardcoded copy of the same numbers.
    assert.doesNotMatch(
        block,
        /Heading level: *\d/,
        "the headingLevel description spells out the bound instead of deriving it",
    );
});

test("validateIntent enforces exactly the bound the schema declares", () => {
    // The pair, not either half: a shared constant makes the two *agree*, but
    // only this asserts the runtime actually refuses what the schema forbids.
    const at = (headingLevel) => () =>
        validateIntent({ op: "set_heading_level", address: "p:0123456789ab", headingLevel });

    assert.equal(at(MIN_HEADING_LEVEL)().headingLevel, MIN_HEADING_LEVEL);
    assert.equal(at(MAX_HEADING_LEVEL)().headingLevel, MAX_HEADING_LEVEL);

    assert.equal(codeOf(at(MIN_HEADING_LEVEL - 1)), "invalid_heading_level");
    assert.equal(codeOf(at(MAX_HEADING_LEVEL + 1)), "invalid_heading_level");
});
