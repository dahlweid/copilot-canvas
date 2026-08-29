// The typed document spec that `create_document` authors from.
//
// This is the authoring counterpart to `edit-intent.mjs`, and it deliberately
// reuses that file's vocabulary rather than inventing a parallel one: issue #8
// requires that a document can be created whole and then edited incrementally,
// which only holds if "a paragraph" and "a heading level" mean the same thing on
// both sides. Every rule those two share is imported from there, never restated.
//
// Why blocks rather than paragraphs
// ---------------------------------
// An edit intent addresses one paragraph because an address is a coordinate into
// a document that already exists. Nothing exists yet here, so there is nothing to
// address; a spec is an ordered list of *blocks*, and a block may expand to many
// paragraphs. A list of four items is four paragraphs, and a 3x2 table is nine.
// Keeping the block as the unit is what lets the caller say "a bulleted list"
// without knowing how Word counts it.
//
// Why the shapes are closed
// -------------------------
// Every block kind names exactly the fields it takes, and an unknown field is an
// error rather than something quietly ignored. A spec that is silently
// half-applied produces a document that differs from what was asked for with no
// error raised — which is the same failure mode as autocorrect, and it is the
// one this layer exists to eliminate.

import {
    EditIntentError,
    MAX_HEADING_LEVEL,
    MAX_TEXT_LENGTH,
    MIN_HEADING_LEVEL,
    requireText,
} from "./edit-intent.mjs";

export { EditIntentError, MAX_HEADING_LEVEL, MAX_TEXT_LENGTH, MIN_HEADING_LEVEL };

/**
 * The lowest level a `heading` block may carry.
 *
 * Derived, not written down: `edit_document` uses level 0 to mean "body text",
 * and a heading block whose level said "not a heading" would be a paragraph
 * block spelled a second way. So the floor here is one above that floor, and it
 * moves if that one does.
 */
export const MIN_BLOCK_HEADING_LEVEL = MIN_HEADING_LEVEL + 1;

/**
 * Ceilings.
 *
 * Each of these is exported and consumed by the tool schema, so the bound the
 * model is told and the bound the runtime enforces are one value. This repo has
 * shipped the drifted version of that three times in three pull requests — a
 * description hardcoding a range a constant already defined, the constant
 * moving, and the model being told something false with nothing going red. The
 * fix is not vigilance; it is that there is only one number.
 */
export const MAX_BLOCKS = 500;
export const MAX_LIST_ITEMS = 500;
export const MAX_TABLE_ROWS = 200;
export const MAX_TABLE_COLUMNS = 40;

/**
 * Floors.
 *
 * Constants for the same reason the ceilings are, and named because the schema
 * did not declare them: a review found `items`, `rows` and each row declared
 * with no `minItems` while the validator refused every empty one. Schema-valid
 * and always rejected at runtime is the worst shape a contract can have — the
 * model cannot learn the rule from the schema and only meets it as a failure.
 *
 * They are all 1 and unlikely to move, which is not the point. The point is
 * that the declared floor and the enforced floor are one value, so they cannot
 * disagree again the way they just did.
 */
export const MIN_LIST_ITEMS = 1;
export const MIN_TABLE_ROWS = 1;
export const MIN_TABLE_COLUMNS = 1;
export const MIN_BLOCKS = 1;

/**
 * The block kinds and the fields each one takes.
 *
 * A single table rather than a switch, so a kind cannot accept a field it then
 * ignores and so `BLOCK_KINDS` cannot fall out of step with what is implemented.
 */
export const BLOCKS = {
    heading: { required: ["level", "text"], optional: [] },
    paragraph: { required: ["text"], optional: [] },
    list: { required: ["items"], optional: ["ordered"] },
    table: { required: ["rows"], optional: ["headerRow"] },
};

export const BLOCK_KINDS = Object.keys(BLOCKS);

/** One line of prose per kind, for the tool description. Derived, not restated. */
export const BLOCK_HELP = [
    `heading — a heading at \`level\` ${MIN_BLOCK_HEADING_LEVEL}–${MAX_HEADING_LEVEL}.`,
    "paragraph — one paragraph of body text.",
    "list — `items`, bulleted, or numbered when `ordered` is true.",
    "table — `rows`, a rectangular array of cell strings, with an optional `headerRow`.",
].join(" ");

function andList(names) {
    if (names.length <= 1) return names.join("");
    return `${names.slice(0, -1).join(", ")} and ${names[names.length - 1]}`;
}

/**
 * Which kinds take a field, and whether they require it — read off `BLOCKS`.
 *
 * The tool schema cannot express "required, but only when `kind` is heading",
 * so `required` there holds only `kind` and the per-kind rule has to live in
 * each field's description. `edit_document` reached the same place and solved it
 * the same way, over seven review rounds.
 *
 * The part worth keeping is that this is *derived*. Those clauses were written
 * by hand first — "heading only:", "list only:" — which is a second copy of
 * `BLOCKS` in prose, in the one place no test looks and the model does. Moving
 * a field between kinds would have left the description describing the old
 * shape, and the review that caught the missing floors caught this too.
 *
 * Throws for a field no kind takes, which makes a rename fail at extension load
 * rather than ship a schema advertising a field the validator refuses.
 */
export function fieldUsage(field) {
    const required = BLOCK_KINDS.filter((kind) => BLOCKS[kind].required.includes(field));
    const optional = BLOCK_KINDS.filter((kind) => BLOCKS[kind].optional.includes(field));
    const parts = [];
    if (required.length > 0) parts.push(`required by ${andList(required)}`);
    if (optional.length > 0) parts.push(`optional on ${andList(optional)}`);
    if (parts.length === 0) {
        throw new Error(`no block kind takes \`${field}\`; BLOCKS and the tool schema have diverged.`);
    }
    return `${parts.join("; ")}.`;
}

function fail(code, message) {
    throw new EditIntentError(code, message);
}

function requireBlockHeadingLevel(value, where) {
    if (!Number.isInteger(value) || value < MIN_BLOCK_HEADING_LEVEL || value > MAX_HEADING_LEVEL) {
        fail(
            "invalid_heading_level",
            `${where}: \`level\` must be an integer from ${MIN_BLOCK_HEADING_LEVEL} to ${MAX_HEADING_LEVEL}, got ${JSON.stringify(value)}.`,
        );
    }
    return value;
}

/**
 * Applies the shared paragraph-text rule and re-labels the failure.
 *
 * `requireText` reports which rule was broken but not which block broke it, and
 * a spec of 500 blocks with one bad line is unactionable without the position.
 */
function requireBlockText(value, where) {
    try {
        return requireText(value);
    } catch (err) {
        if (err instanceof EditIntentError) fail(err.code, `${where}: ${err.message}`);
        throw err;
    }
}

function requireFields(block, kind, index) {
    const shape = BLOCKS[kind];
    const allowed = new Set([...shape.required, ...shape.optional, "kind"]);
    const extra = Object.keys(block).filter((key) => !allowed.has(key));
    if (extra.length > 0) {
        fail(
            "invalid_block",
            `block ${index} (${kind}): unsupported field${extra.length === 1 ? "" : "s"} ${extra.join(", ")}. ${kind} takes ${shape.required.concat(shape.optional).join(", ")}.`,
        );
    }
    for (const field of shape.required) {
        if (block[field] === undefined) {
            fail("invalid_block", `block ${index} (${kind}): \`${field}\` is required.`);
        }
    }
}

function normalizeBlock(block, index) {
    const where = `block ${index}`;
    if (!block || typeof block !== "object" || Array.isArray(block)) {
        fail("invalid_block", `${where}: each block must be an object.`);
    }

    const kind = block.kind;
    if (typeof kind !== "string" || !Object.hasOwn(BLOCKS, kind)) {
        fail("unknown_block_kind", `${where}: unknown kind ${JSON.stringify(kind)}. Supported: ${BLOCK_KINDS.join(", ")}.`);
    }
    requireFields(block, kind, index);
    const label = `${where} (${kind})`;

    switch (kind) {
        case "heading":
            return {
                kind,
                level: requireBlockHeadingLevel(block.level, label),
                text: requireBlockText(block.text, label),
            };

        case "paragraph":
            return { kind, text: requireBlockText(block.text, label) };

        case "list": {
            if (!Array.isArray(block.items)) fail("invalid_block", `${label}: \`items\` must be an array of strings.`);
            if (block.items.length < MIN_LIST_ITEMS) {
                fail("invalid_block", `${label}: \`items\` needs at least ${MIN_LIST_ITEMS}, got ${block.items.length}.`);
            }
            if (block.items.length > MAX_LIST_ITEMS) {
                fail("invalid_block", `${label}: ${block.items.length} items; the limit is ${MAX_LIST_ITEMS}.`);
            }
            if (block.ordered !== undefined && typeof block.ordered !== "boolean") {
                fail("invalid_block", `${label}: \`ordered\` must be a boolean.`);
            }
            return {
                kind,
                ordered: block.ordered === true,
                items: block.items.map((item, i) => requireBlockText(item, `${label} item ${i}`)),
            };
        }

        case "table": {
            if (!Array.isArray(block.rows)) fail("invalid_block", `${label}: \`rows\` must be an array of arrays.`);
            if (block.rows.length < MIN_TABLE_ROWS) {
                fail("invalid_block", `${label}: \`rows\` needs at least ${MIN_TABLE_ROWS}, got ${block.rows.length}.`);
            }
            if (block.rows.length > MAX_TABLE_ROWS) {
                fail("invalid_block", `${label}: ${block.rows.length} rows; the limit is ${MAX_TABLE_ROWS}.`);
            }
            if (block.headerRow !== undefined && typeof block.headerRow !== "boolean") {
                fail("invalid_block", `${label}: \`headerRow\` must be a boolean.`);
            }

            const width = Array.isArray(block.rows[0]) ? block.rows[0].length : -1;
            if (width < MIN_TABLE_COLUMNS) {
                fail("invalid_block", `${label}: each row must be an array of at least ${MIN_TABLE_COLUMNS} cell string(s).`);
            }
            if (width > MAX_TABLE_COLUMNS) {
                fail("invalid_block", `${label}: ${width} columns; the limit is ${MAX_TABLE_COLUMNS}.`);
            }

            const rows = block.rows.map((row, r) => {
                if (!Array.isArray(row)) fail("invalid_block", `${label} row ${r}: must be an array of cell strings.`);
                // Word's Tables.Add takes one row count and one column count, so
                // a ragged spec has no faithful rendering. Refusing is the only
                // outcome that cannot silently produce a different document.
                if (row.length !== width) {
                    fail(
                        "invalid_block",
                        `${label} row ${r}: has ${row.length} cells but row 0 has ${width}. A table must be rectangular.`,
                    );
                }
                return row.map((cell, c) => requireBlockText(cell, `${label} row ${r} cell ${c}`));
            });

            return { kind, headerRow: block.headerRow === true, rows };
        }

        default:
            // Unreachable: `kind` was checked against BLOCKS above. Present so
            // that adding a key to BLOCKS without implementing it fails loudly
            // rather than producing a document missing that block.
            return fail("unknown_block_kind", `${label}: no translation implemented.`);
    }
}

/**
 * Validates a document spec and returns it normalized.
 *
 * Throws rather than returning a result object, for the reason `validateIntent`
 * does: an invalid spec must never reach the point where a file is created, and
 * a half-written document is worse than no document.
 */
export function validateSpec(input) {
    if (!input || typeof input !== "object" || Array.isArray(input)) {
        fail("invalid_spec", "A document spec must be an object with a `blocks` array.");
    }

    const { blocks, ...rest } = input;

    const extra = Object.keys(rest);
    if (extra.length > 0) {
        fail("invalid_spec", `Unsupported field${extra.length === 1 ? "" : "s"}: ${extra.join(", ")}.`);
    }

    if (!Array.isArray(blocks)) fail("invalid_spec", "`blocks` must be an array.");
    if (blocks.length < MIN_BLOCKS) {
        fail("invalid_spec", "`blocks` may not be empty; there would be nothing to author.");
    }
    if (blocks.length > MAX_BLOCKS) {
        fail("invalid_spec", `${blocks.length} blocks; the limit for one document is ${MAX_BLOCKS}.`);
    }

    return { blocks: blocks.map((block, index) => normalizeBlock(block, index)) };
}

/** A short, human-readable description, for logs and snapshot manifests. */
export function describeSpec(spec) {
    const counts = new Map();
    for (const block of spec.blocks) counts.set(block.kind, (counts.get(block.kind) ?? 0) + 1);
    const parts = [...counts].map(([kind, n]) => `${n} ${kind}${n === 1 ? "" : "s"}`);
    return `a document of ${parts.join(", ")}`;
}
