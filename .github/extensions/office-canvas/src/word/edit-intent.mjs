// The typed edit intent.
//
// ADR 0002 makes the agent the editor, which is exactly why this is a closed
// set of named operations rather than a free-form payload. Two shapes were
// deliberately rejected:
//
//   * **OOXML.** Handing Word markup straight through would make every caller
//     responsible for schema correctness and would let a malformed fragment
//     corrupt a document the agent cannot then read back to diagnose.
//   * **Prose.** "Make the second paragraph friendlier" is not verifiable. An
//     operation that cannot be checked against the document afterwards cannot
//     be reverted with any confidence either.
//
// Styles are the other reason this file exists. It is measured on this machine
// that `Range.Style` accepts neither the OOXML style id nor the English style
// name:
//
//   | `Range.Style = 'berschrift1'` (the id in the file) | throws |
//   | `Range.Style = 'Heading 1'`                        | throws |
//   | `Range.Style = <Style object of another paragraph>`| works  |
//   | `Range.Style = -2` (wdStyleHeading1)               | works  |
//
// So the write side never names a style at all. An intent carries a numeric
// `headingLevel`, which maps to the language-independent `wd*` constants, and
// an insert with no `headingLevel` inherits the reference paragraph's `Style`
// *object*. Carrying the style id verbatim remains right for addressing and
// reading; it is not a thing you can write back.

/** Matches `mintAddress` output. Pinned against it by unit test, not by eye. */
const ADDRESS_PATTERN = /^p:[0-9a-f]{12}$/;

/** Generous, but not unbounded: a single paragraph is not a document. */
export const MAX_TEXT_LENGTH = 50_000;

/**
 * Exported because `edit_document`'s schema declares this same bound. L1 hit
 * the drift version of this with `limit` -- a schema that advertised one bound
 * while the runtime enforced another -- and fixed it by making the declared
 * and the enforced bound one constant. Same shape here: change it once.
 */
export const MIN_HEADING_LEVEL = 0;
export const MAX_HEADING_LEVEL = 9;

export class EditIntentError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "EditIntentError";
        this.code = code;
    }
}

/**
 * The operations, and what each one requires.
 *
 * `text` and `headingLevel` are listed per operation rather than validated ad
 * hoc so that an operation cannot quietly accept a field it then ignores.
 */
export const OPERATIONS = {
    replace_text: {
        text: "required",
        headingLevel: "rejected",
        help: "rewrite the paragraph's text, keeping its style",
    },
    insert_paragraph_after: {
        text: "required",
        headingLevel: "optional",
        help: "add a new paragraph after it",
    },
    insert_paragraph_before: {
        text: "required",
        headingLevel: "optional",
        help: "add a new paragraph before it",
    },
    delete_paragraph: {
        text: "rejected",
        headingLevel: "rejected",
        help: "remove it",
    },
    set_heading_level: {
        text: "rejected",
        headingLevel: "required",
        help: `make it a heading (${MIN_HEADING_LEVEL + 1}–${MAX_HEADING_LEVEL}) or body text (${MIN_HEADING_LEVEL})`,
    },
};

export const OPERATION_NAMES = Object.keys(OPERATIONS);

/**
 * The caller-facing description of the operations, generated from the same
 * table that validates them.
 *
 * It lives here rather than in `extension.mjs` for two reasons. It cannot then
 * document an operation that does not exist, or omit one that does — the two
 * lists are one list. And it is importable, so a test can assert on the text a
 * caller actually reads; the version in `extension.mjs` was module-scope in a
 * file that calls `joinSession()` on import and so could only ever be checked
 * by reading its own source as a string.
 */
export const OPERATION_HELP = OPERATION_NAMES.map((name) => `${name} — ${OPERATIONS[name].help}.`).join(" ");

/**
 * Text a paragraph may be given.
 *
 * Line breaks are refused rather than translated. A paragraph break inside the
 * text would silently turn one addressed paragraph into several, and every
 * address after it in the document would shift — which is precisely the thing
 * the address model cannot absorb. Several paragraphs is several operations,
 * each against a fresh read.
 *
 * Exported because `create_document` writes paragraphs too, and the rule is the
 * same one for the same reason. A second copy of it would be a second place for
 * the bound and the wording to drift apart.
 */
export function requireText(value) {
    if (typeof value !== "string") {
        throw new EditIntentError("invalid_text", "`text` must be a string.");
    }
    if (value.length > MAX_TEXT_LENGTH) {
        throw new EditIntentError(
            "invalid_text",
            `\`text\` is ${value.length} characters; the limit for one paragraph is ${MAX_TEXT_LENGTH}.`,
        );
    }
    if (/[\r\n\v\f\u0007]/.test(value)) {
        throw new EditIntentError(
            "invalid_text",
            "`text` may not contain line or paragraph breaks. One operation edits one paragraph; insert further paragraphs with further operations, each against a fresh read.",
        );
    }
    return value;
}

function requireHeadingLevel(value) {
    if (!Number.isInteger(value) || value < MIN_HEADING_LEVEL || value > MAX_HEADING_LEVEL) {
        throw new EditIntentError(
            "invalid_heading_level",
            `\`headingLevel\` must be an integer from ${MIN_HEADING_LEVEL} (body text) to ${MAX_HEADING_LEVEL}, got ${JSON.stringify(value)}.`,
        );
    }
    return value;
}

/**
 * Validates an intent and returns it normalized.
 *
 * Throws rather than returning a result object: an invalid intent must never
 * reach the point where a snapshot is taken, and a thrown typed error is
 * harder to ignore than a flag.
 */
export function validateIntent(input) {
    if (!input || typeof input !== "object" || Array.isArray(input)) {
        throw new EditIntentError("invalid_intent", "An edit intent must be an object.");
    }

    const { op, address, text, headingLevel, ...rest } = input;

    if (typeof op !== "string" || !Object.hasOwn(OPERATIONS, op)) {
        throw new EditIntentError(
            "unknown_operation",
            `Unknown operation ${JSON.stringify(op)}. Supported: ${OPERATION_NAMES.join(", ")}.`,
        );
    }
    const shape = OPERATIONS[op];

    const extra = Object.keys(rest);
    if (extra.length > 0) {
        throw new EditIntentError(
            "invalid_intent",
            `Unsupported field${extra.length === 1 ? "" : "s"} for ${op}: ${extra.join(", ")}.`,
        );
    }

    if (typeof address !== "string" || !ADDRESS_PATTERN.test(address)) {
        throw new EditIntentError(
            "invalid_address",
            `\`address\` must be a paragraph address from read_document, like "p:0123456789ab", got ${JSON.stringify(address)}.`,
        );
    }

    const normalized = { op, address };

    if (shape.text === "required") normalized.text = requireText(text);
    else if (text !== undefined) {
        throw new EditIntentError("invalid_intent", `${op} takes no \`text\`.`);
    }

    if (shape.headingLevel === "required") normalized.headingLevel = requireHeadingLevel(headingLevel);
    else if (shape.headingLevel === "optional" && headingLevel !== undefined) {
        normalized.headingLevel = requireHeadingLevel(headingLevel);
    } else if (shape.headingLevel === "rejected" && headingLevel !== undefined) {
        throw new EditIntentError("invalid_intent", `${op} takes no \`headingLevel\`.`);
    }

    return normalized;
}

/** A short, human-readable description, used in snapshot manifests and logs. */
export function describeIntent(intent) {
    switch (intent.op) {
        case "replace_text":
            return `replace the text of ${intent.address}`;
        case "insert_paragraph_after":
            return `insert a paragraph after ${intent.address}`;
        case "insert_paragraph_before":
            return `insert a paragraph before ${intent.address}`;
        case "delete_paragraph":
            return `delete ${intent.address}`;
        case "set_heading_level":
            return intent.headingLevel === 0
                ? `make ${intent.address} body text`
                : `make ${intent.address} a level ${intent.headingLevel} heading`;
        default:
            return intent.op;
    }
}
