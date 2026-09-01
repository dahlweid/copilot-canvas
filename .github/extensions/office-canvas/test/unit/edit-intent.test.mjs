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
    fieldRequirementHelp,
    HEADING_LEVEL_HELP,
    OPERATION_HELP,
    MAX_HEADING_LEVEL,
    MIN_HEADING_LEVEL,
    OPERATION_NAMES,
    TEXT_HELP,
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

test("the help text documents every operation, and only real ones", () => {
    // Generated from the same table that validates, so this cannot drift -- but
    // assert it anyway, because "cannot drift" is a property of the current
    // implementation and this is the property callers depend on.
    for (const name of OPERATION_NAMES) {
        assert.ok(OPERATION_HELP.includes(`${name} —`), `${name} is not documented in the help text`);
    }
    const documented = [...OPERATION_HELP.matchAll(/([a-z_]+) —/g)].map((m) => m[1]);
    assert.deepEqual(documented.sort(), [...OPERATION_NAMES].sort(), "the help text names an operation that does not exist");
});

test("the heading range the help advertises is the range the validator enforces", () => {
    // Read as a caller reads it: the numbers come out of the rendered text, not
    // out of the constants. Asserting `help.includes(MAX_HEADING_LEVEL)` would
    // build the expectation from the same constant that built the text and pass
    // for any value of it -- the tautology `supported-types.test.mjs` documents.
    const found = OPERATION_HELP.match(/heading \((\d+)[–-](\d+)\) or body text \((\d+)\)/);
    assert.ok(found, `no heading range in the help text: ${OPERATION_HELP}`);

    const [lowest, highest, body] = [Number(found[1]), Number(found[2]), Number(found[3])];
    const at = (headingLevel) => () =>
        validateIntent({ op: "set_heading_level", address: "p:0123456789ab", headingLevel });

    for (const level of [body, lowest, highest]) {
        assert.equal(at(level)().headingLevel, level, `the help advertises ${level} but the validator refuses it`);
    }
    assert.equal(codeOf(at(highest + 1)), "invalid_heading_level", `the validator accepts ${highest + 1}, which the help does not advertise`);
    assert.equal(codeOf(at(body - 1)), "invalid_heading_level", `the validator accepts ${body - 1}, which the help does not advertise`);
});
// The `text` and `headingLevel` schema descriptions, read the way a caller
// reads them.
//
// These parse the rendered sentence and then put the validator through what it
// claims, rather than comparing the string against the table that built it --
// the tautology `CONTEXT.md` documents, which passes with the defect fully
// reintroduced. Here a mislabelled clause (an operation advertised as taking a
// field it refuses, or the phrasing for two levels swapped) goes red, because
// the expectation comes out of the prose and the answer comes out of
// `validateIntent`. An operation named in the prose that does not exist goes
// red too: the validator answers `unknown_operation`, which is not the code any
// branch below expects.

const namesIn = (clause) =>
    clause
        .split(/,| and /)
        .map((name) => name.trim())
        .filter(Boolean);

const requirementsFrom = (help) => {
    const clause = (pattern) => {
        const found = help.match(pattern);
        return found ? namesIn(found[1]) : [];
    };
    return {
        required: clause(/Required by ([^;.]+)/),
        optional: clause(/optional on ([^;.]+)/),
        refused: clause(/refused by ([^;.]+)/),
    };
};

const textRules = requirementsFrom(TEXT_HELP);
const headingRules = requirementsFrom(HEADING_LEVEL_HELP);

/** The smallest intent the `text` prose says is complete apart from `headingLevel`. */
const baseIntent = (op) =>
    textRules.required.includes(op) ? { op, address: ADDRESS, text: "Hello" } : { op, address: ADDRESS };

test("the `text` rule the schema advertises is the rule the validator enforces", () => {
    assert.ok(textRules.required.length > 0, `no "Required by" clause in: ${TEXT_HELP}`);
    assert.ok(textRules.refused.length > 0, `no "refused by" clause in: ${TEXT_HELP}`);

    for (const op of textRules.required) {
        assert.equal(validateIntent({ op, address: ADDRESS, text: "Hello" }).text, "Hello");
        assert.equal(
            codeOf(() => validateIntent({ op, address: ADDRESS })),
            "invalid_text",
            `the schema says ${op} requires \`text\`, but the validator accepts it without one`,
        );
    }

    for (const op of textRules.optional) {
        assert.equal(validateIntent({ op, address: ADDRESS, text: "Hello" }).text, "Hello");
        assert.equal(
            validateIntent({ op, address: ADDRESS }).text,
            undefined,
            `the schema says \`text\` is optional on ${op}, but the validator will not go without it`,
        );
    }

    for (const op of textRules.refused) {
        assert.equal(
            codeOf(() => validateIntent({ op, address: ADDRESS, text: "Hello" })),
            "invalid_intent",
            `the schema says ${op} refuses \`text\`, but the validator took one`,
        );
    }
});

test("the `headingLevel` rule the schema advertises is the rule the validator enforces", () => {
    assert.ok(headingRules.required.length > 0, `no "Required by" clause in: ${HEADING_LEVEL_HELP}`);
    assert.ok(headingRules.optional.length > 0, `no "optional on" clause in: ${HEADING_LEVEL_HELP}`);
    assert.ok(headingRules.refused.length > 0, `no "refused by" clause in: ${HEADING_LEVEL_HELP}`);

    for (const op of headingRules.required) {
        assert.equal(validateIntent({ ...baseIntent(op), headingLevel: 1 }).headingLevel, 1);
        assert.equal(
            codeOf(() => validateIntent(baseIntent(op))),
            "invalid_heading_level",
            `the schema says ${op} requires \`headingLevel\`, but the validator accepts it without one`,
        );
    }

    for (const op of headingRules.optional) {
        assert.equal(validateIntent({ ...baseIntent(op), headingLevel: 1 }).headingLevel, 1);
        assert.equal(
            validateIntent(baseIntent(op)).headingLevel,
            undefined,
            `the schema says \`headingLevel\` is optional on ${op}, but the validator will not go without it`,
        );
    }

    for (const op of headingRules.refused) {
        assert.equal(
            codeOf(() => validateIntent({ ...baseIntent(op), headingLevel: 1 })),
            "invalid_intent",
            `the schema says ${op} refuses \`headingLevel\`, but the validator took one`,
        );
    }
});

test("every operation reaches both schema descriptions", () => {
    // The falsifiable half of the derivation. The two tests above pin the *rule*
    // for every operation the prose names, so they cannot notice one the prose
    // fails to name -- measured, not assumed: making `fieldRequirementHelp` skip
    // `insert_paragraph_before` dropped it from both descriptions and left the
    // suite at 18 pass, 0 fail. This one goes red on that mutation.
    //
    // It is not the tautology `CONTEXT.md` documents. There the compared list is
    // built inside the test from the same constant by the same path, so no
    // change to the code under test can separate them. Here the expectation is
    // `OPERATION_NAMES` and the subject is the string `fieldRequirementHelp`
    // produced, with the function under test in between -- which is exactly
    // where an omission would be introduced.
    for (const [label, help] of [
        ["TEXT_HELP", TEXT_HELP],
        ["HEADING_LEVEL_HELP", HEADING_LEVEL_HELP],
    ]) {
        const rules = requirementsFrom(help);
        const named = [...rules.required, ...rules.optional, ...rules.refused].sort();
        assert.deepEqual(
            named,
            [...OPERATION_NAMES].sort(),
            `${label} does not account for every operation exactly once: ${help}`,
        );
    }
});

test("the table carries no requirement level the derivation cannot render", () => {
    // Importing this module runs `fieldRequirementHelp` for both fields, so an
    // unrecognised level throws before anything can read a description that
    // silently dropped an operation. Exercised through a field the table does
    // not carry, because a level that is merely absent looks the same to the
    // guard as one that is misspelled.
    assert.throws(() => fieldRequirementHelp("noSuchField"), /not a requirement level/);
    for (const field of ["text", "headingLevel"]) {
        assert.ok(fieldRequirementHelp(field).endsWith("."), `${field} renders no sentence`);
    }
});
