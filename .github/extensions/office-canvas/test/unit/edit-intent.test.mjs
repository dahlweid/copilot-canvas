// Office-free tests for the edit intent.
//
// The point of a closed operation set is that a bad intent is refused before a
// snapshot is taken and before Word is started, so these tests are the first
// line of the edit path even though they never touch a document.

import test from "node:test";
import assert from "node:assert/strict";

import {
    describeIntent,
    EditIntentError,
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
