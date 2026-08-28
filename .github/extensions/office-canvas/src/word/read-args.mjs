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
// Deliberately NOT enforced here: the schema's `maximum`. An over-large limit is
// already answered correctly -- `slice` clamps, and the response reports
// `returned` and `truncated` honestly -- so rejecting it would refuse a request
// we can satisfy perfectly. Validate where the behaviour is wrong, not wherever
// a bound happens to be written down.
//
// This is checked here rather than left to the declared JSON schema because
// nothing in this repo demonstrates that the extension host validates schemas
// before dispatch. That may well be true; it is simply not something we have
// measured, and a guard that costs three comparisons is cheaper than the
// experiment.

/** Paragraphs returned when the caller expresses no preference. */
export const DEFAULT_READ_LIMIT = 300;

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
const requireInteger = (value, field, minimum, hint) => {
    if (!Number.isInteger(value)) {
        throw new ReadArgsError(`\`${field}\` must be an integer, not ${describe(value)}. ${hint}`);
    }
    if (value < minimum) {
        throw new ReadArgsError(`\`${field}\` must be ${minimum} or greater, not ${value}. ${hint}`);
    }
    return value;
};

/**
 * Resolves `limit`/`offset` for a structure read, or throws `invalid_request`.
 *
 * Absent (`undefined` or `null`) means "no preference" and takes the default.
 * Anything present must be a valid integer -- including 0, which is a real
 * request for zero paragraphs and not a way of asking for all of them.
 */
export function normalizeReadArgs(args, { defaultLimit = DEFAULT_READ_LIMIT } = {}) {
    const source = args ?? {};
    const limit = source.limit ?? null;
    const offset = source.offset ?? null;

    return {
        limit:
            limit === null
                ? defaultLimit
                : requireInteger(limit, "limit", 1, `Omit \`limit\` for the default of ${defaultLimit}.`),
        offset: offset === null ? 0 : requireInteger(offset, "offset", 0, "Omit `offset` to start at the beginning."),
    };
}
