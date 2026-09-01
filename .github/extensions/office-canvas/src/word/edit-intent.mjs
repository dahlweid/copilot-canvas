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
 *
 * `help` and `describe` are the two things an operation says about itself:
 * `help` is what the schema tells the model beforehand, `describe` is what a
 * snapshot manifest and a log line record afterwards. Both sit on the entry so
 * that adding an operation is one edit in one place. `describe` was a `switch`
 * at the foot of this file ending in `default: return intent.op` — a second,
 * silent list of the same names (issue #131).
 *
 * A describer is only ever called for its own operation and so never branches
 * on one; the single conditional here is `set_heading_level`'s, and it is on
 * the *value*, not the operation. That is the whole reason these are functions
 * rather than strings: level 0 is body text, not a level 0 heading.
 */
export const OPERATIONS = {
    replace_text: {
        text: "required",
        headingLevel: "rejected",
        help: "rewrite the paragraph's text, keeping its style",
        describe: ({ address }) => `replace the text of ${address}`,
    },
    insert_paragraph_after: {
        text: "required",
        headingLevel: "optional",
        help: "add a new paragraph after it",
        describe: ({ address }) => `insert a paragraph after ${address}`,
    },
    insert_paragraph_before: {
        text: "required",
        headingLevel: "optional",
        help: "add a new paragraph before it",
        describe: ({ address }) => `insert a paragraph before ${address}`,
    },
    delete_paragraph: {
        text: "rejected",
        headingLevel: "rejected",
        help: "remove it",
        describe: ({ address }) => `delete ${address}`,
    },
    set_heading_level: {
        text: "rejected",
        headingLevel: "required",
        help: `make it a heading (${MIN_HEADING_LEVEL + 1}–${MAX_HEADING_LEVEL}) or body text (${MIN_HEADING_LEVEL})`,
        describe: ({ address, headingLevel }) =>
            headingLevel === MIN_HEADING_LEVEL
                ? `make ${address} body text`
                : `make ${address} a level ${headingLevel} heading`,
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

const REQUIREMENT_LEVELS = ["required", "optional", "rejected"];

function joinNames(names) {
    if (names.length <= 2) return names.join(" and ");
    return `${names.slice(0, -1).join(", ")} and ${names.at(-1)}`;
}

/**
 * Which operations require, allow and refuse one of the per-operation fields,
 * phrased for a caller.
 *
 * `edit_document`'s tool schema used to hand-write this same split as prose in
 * the `text` and `headingLevel` descriptions, which made the schema a second
 * copy of the table above and left the two free to drift — the restatement this
 * repo has corrected four times before (issue #28). The reasoning is
 * `OPERATION_HELP`'s: the rule the validator enforces and the rule the model
 * reads are one rule, so there is one place to change it.
 *
 * Exported for one test only: an unrecognised level throws here, so importing
 * this module type-checks every level in the table, and a guard nothing
 * exercises is a guard that may not work.
 */
export function fieldRequirementHelp(field) {
    const byLevel = new Map(REQUIREMENT_LEVELS.map((level) => [level, []]));

    for (const name of OPERATION_NAMES) {
        const level = OPERATIONS[name][field];
        const names = byLevel.get(level);
        if (!names) {
            throw new Error(`${name}.${field} is ${JSON.stringify(level)}, which is not a requirement level.`);
        }
        names.push(name);
    }

    const phrases = {
        required: (names) => `Required by ${joinNames(names)}`,
        optional: (names) => `optional on ${joinNames(names)}`,
        rejected: (names) => `refused by ${joinNames(names)}`,
    };

    return `${REQUIREMENT_LEVELS.filter((level) => byLevel.get(level).length > 0)
        .map((level) => phrases[level](byLevel.get(level)))
        .join("; ")}.`;
}

/**
 * The `text` and `headingLevel` descriptions `edit_document`'s schema carries.
 *
 * They live beside the table for the same reason `OPERATION_HELP` does: they
 * are generated from it, and they are importable, so a test can assert on the
 * text a caller actually reads rather than on `extension.mjs`'s source.
 */
export const TEXT_HELP = `The new text. ${fieldRequirementHelp("text")} One paragraph: line breaks are refused, because a second paragraph would move the addresses after it.`;

export const HEADING_LEVEL_HELP = `Heading level: ${MIN_HEADING_LEVEL + 1}–${MAX_HEADING_LEVEL} for a heading, ${MIN_HEADING_LEVEL} for body text. ${fieldRequirementHelp(
    "headingLevel",
)} An insert given no \`headingLevel\` follows the style Word would use itself.`;

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

    // Both fields read the same three levels, and both handle all three, even
    // though no operation marks `text` optional today. The asymmetry was real
    // and measured: `fieldRequirementHelp` renders "optional on X" for whichever
    // field the table says it about, so a `text: "optional"` operation would
    // have shipped a schema description the validator then refused — a promise
    // nothing keeps, in a string written for the model to act on.
    if (shape.text === "required") normalized.text = requireText(text);
    else if (shape.text === "optional" && text !== undefined) {
        normalized.text = requireText(text);
    } else if (shape.text === "rejected" && text !== undefined) {
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

/**
 * The describer for an operation, or a throw.
 *
 * The failure this refuses to be quiet about is issue #131's. `describeIntent`
 * used to end in `default: return intent.op`, so an operation added to the
 * table above and not to that switch described itself as its own bare name.
 * That was not a visibly broken string — nothing downstream could tell it from
 * a deliberate one — and it dropped the *address*, which for this string's two
 * consumers is the part carrying the information. A revert manifest naming the
 * operation and not the paragraph it was applied to would have been a recovery
 * path with the recovery taken out.
 *
 * Which is why this throws instead of degrading: with no fallback left, an
 * operation carrying no prose never reaches a manifest at all. The message says
 * that and no more. Predicting the old behaviour here would be this repo's own
 * defect — an error naming a consequence the code has just made impossible.
 */
function describerFor(op) {
    const describe = OPERATIONS[op]?.describe;
    if (typeof describe !== "function") {
        throw new Error(
            `Operation ${JSON.stringify(op)} has no \`describe\` in OPERATIONS, so it is refused rather than ` +
                `described. Snapshot manifests and log lines are written from that string and \`revert_document\` ` +
                `reads the manifest back, so every operation needs prose of its own naming the paragraph it was ` +
                `applied to.`,
        );
    }
    return describe;
}

// Runs at import, the way `fieldRequirementHelp`'s level check does: an
// operation added to the table with no prose of its own fails before the
// extension loads, rather than at the moment a manifest is being written for
// an edit that is about to happen.
for (const name of OPERATION_NAMES) describerFor(name);

/** A short, human-readable description, used in snapshot manifests and logs. */
export function describeIntent(intent) {
    return describerFor(intent.op)(intent);
}
