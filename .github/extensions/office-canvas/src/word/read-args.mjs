// Paging arguments for a structure read, validated at the tool boundary.
//
// `buildStructureMap` treats `limit = 0` as "return everything" -- a reasonable
// internal convention, and the reason this module exists. The tool handler used
// to apply its default with `??`, which only substitutes for null/undefined, so
// a caller asking for `limit: 0` passed straight through to that convention and
// received *every* paragraph.
//
// The caller here is a language model, so 0 is a plausible input rather than a
// hostile one: a remaining-count that has reached zero, or a guess that 0 means
// "no preference". And the outcome was not an error -- it was the whole document
// silently dumped into the agent's context, which is the exact thing the default
// exists to prevent. The defaulting mechanism defeated itself precisely at the
// boundary it was added for.
//
// We reject rather than clamp. Silently substituting a different number is how a
// caller ends up believing it received a whole document that it did not, which
// is the same failure mode paging already has to be careful about.
//
// The ceiling is enforced here too, and the first version of this module did not
// enforce it -- on the argument that `slice` already answers an over-large limit
// correctly, so rejecting one would refuse a request we can satisfy. That missed
// what makes the bound different from an arbitrary number: the tool schema
// *declares* `maximum: 5000`, which makes it a promise to the caller rather than
// an internal preference. An unenforced declared bound gives the contract two
// answers depending on whether the host pre-validates -- either the caller is
// rejected upstream in a shape that is not our typed `invalid_request`, or the
// declaration is simply false. That is the same defect as `limit: 0`, one level
// up: a constraint the runtime does not enforce. The floor was fixed and the
// ceiling left.
//
// So both bounds are checked here, and MAX_READ_LIMIT is exported for the schema
// to declare, so the two cannot state different numbers.
//
// Checked at runtime rather than left to the declared schema because nothing in
// this repo demonstrates that the extension host validates schemas before
// dispatch. That may well be true; it is simply not something we have measured,
// and enforcing it ourselves makes the answer the same either way.

/** Paragraphs returned when the caller expresses no preference. */
export const DEFAULT_READ_LIMIT = 300;

/**
 * Largest page a caller may ask for. Declared by the tool schema as `maximum`
 * and enforced below; the schema imports this so the two cannot drift.
 */
export const MAX_READ_LIMIT = 5000;

export class ReadArgsError extends Error {
    constructor(message) {
        super(message);
        this.name = "ReadArgsError";
        this.code = "invalid_request";
    }
}

const describe = (value) => {
    if (typeof value === "string") return `a string (${JSON.stringify(value)})`;
    if (typeof value === "number") return String(value);
    if (value === null) return "null";
    return `a ${typeof value}`;
};

/**
 * `Number.isInteger` is the whole type check: it is false for strings, booleans,
 * null, NaN, Infinity and fractions alike, so there is no separate typeof guard.
 */
const requireInteger = (value, field, { minimum, maximum = null, hint }) => {
    if (!Number.isInteger(value)) {
        throw new ReadArgsError(`\`${field}\` must be an integer, not ${describe(value)}. ${hint}`);
    }
    if (value < minimum) {
        throw new ReadArgsError(`\`${field}\` must be ${minimum} or greater, not ${value}. ${hint}`);
    }
    if (maximum !== null && value > maximum) {
        throw new ReadArgsError(
            `\`${field}\` must be ${maximum} or less, not ${value}. Page the document instead: addresses are minted across the whole of it, so paging never changes one.`,
        );
    }
    return value;
};

/**
 * Resolves `limit`/`offset` for a structure read, or throws `invalid_request`.
 *
 * Absent (`undefined` or `null`) means "no preference" and takes the default.
 * Anything present must be an integer within the bounds the tool schema
 * declares: `limit` from 1 to MAX_READ_LIMIT, `offset` from 0. In particular
 * `limit: 0` is rejected rather than treated as "no preference" -- downstream it
 * would mean "every paragraph", which is what this module exists to prevent.
 */
export function normalizeReadArgs(args, { defaultLimit = DEFAULT_READ_LIMIT, maxLimit = MAX_READ_LIMIT } = {}) {
    const source = args ?? {};
    const limit = source.limit ?? null;
    const offset = source.offset ?? null;

    return {
        limit:
            limit === null
                ? defaultLimit
                : requireInteger(limit, "limit", {
                      minimum: 1,
                      maximum: maxLimit,
                      hint: `Omit \`limit\` for the default of ${defaultLimit}.`,
                  }),
        offset:
            offset === null
                ? 0
                : requireInteger(offset, "offset", {
                      minimum: 0,
                      hint: "Omit `offset` to start at the beginning.",
                  }),
    };
}
