// Byte-range parsing for the PDF endpoint.
//
// Office-free: pure function, no Word, no filesystem, no network.
// Run: node --test .github/extensions/office-canvas/test/unit/

import test from "node:test";
import assert from "node:assert/strict";
import { parseByteRange } from "../../src/server.mjs";

const SIZE = 1000;

test("no Range header means send the whole entity", () => {
    assert.equal(parseByteRange(undefined, SIZE), null);
    assert.equal(parseByteRange("", SIZE), null);
    assert.equal(parseByteRange("   ", SIZE), null);
});

test("a closed range is returned inclusive", () => {
    assert.deepEqual(parseByteRange("bytes=0-99", SIZE), { start: 0, end: 99 });
    assert.deepEqual(parseByteRange("bytes=100-199", SIZE), { start: 100, end: 199 });
});

test("an open-ended range runs to the last byte", () => {
    assert.deepEqual(parseByteRange("bytes=900-", SIZE), { start: 900, end: 999 });
});

test("an end past the entity is clamped, not rejected", () => {
    // A reader that asks for more than exists gets what exists; only a *start*
    // past the end is unsatisfiable.
    assert.deepEqual(parseByteRange("bytes=990-5000", SIZE), { start: 990, end: 999 });
});

test("a suffix range returns the final N bytes", () => {
    // The regression this file exists for. `bytes=-500` is not `0-500`: PDF
    // readers fetch the trailer this way, and serving the head of the file
    // returns bytes that parse as garbage rather than failing loudly.
    assert.deepEqual(parseByteRange("bytes=-500", SIZE), { start: 500, end: 999 });
    assert.deepEqual(parseByteRange("bytes=-1", SIZE), { start: 999, end: 999 });
});

test("a suffix larger than the entity yields the whole entity", () => {
    assert.deepEqual(parseByteRange("bytes=-5000", SIZE), { start: 0, end: 999 });
});

test("a start at or past the end is unsatisfiable", () => {
    assert.equal(parseByteRange("bytes=1000-", SIZE), "unsatisfiable");
    assert.equal(parseByteRange("bytes=1000-1200", SIZE), "unsatisfiable");
    assert.equal(parseByteRange("bytes=5000-6000", SIZE), "unsatisfiable");
});

test("a backwards range is unsatisfiable", () => {
    assert.equal(parseByteRange("bytes=200-100", SIZE), "unsatisfiable");
});

test("a zero-length suffix is unsatisfiable", () => {
    // RFC 9110: `-0` cannot be satisfied. The previous implementation served
    // byte 0 for it.
    assert.equal(parseByteRange("bytes=-0", SIZE), "unsatisfiable");
});

test("an empty entity cannot satisfy any range", () => {
    assert.equal(parseByteRange("bytes=0-", 0), "unsatisfiable");
    assert.equal(parseByteRange("bytes=-10", 0), "unsatisfiable");
});

test("malformed or multi-range headers fall back to the whole entity", () => {
    // Ignoring what we cannot honour is required; answering it wrongly is not.
    for (const header of [
        "bytes=abc-def",
        "bytes=0-99, 200-299",
        "items=0-99",
        "bytes=",
        "0-99",
        "bytes=-",
    ]) {
        assert.equal(parseByteRange(header, SIZE), null, `expected passthrough for ${header}`);
    }
});

test("the whole entity can be requested explicitly", () => {
    assert.deepEqual(parseByteRange("bytes=0-999", SIZE), { start: 0, end: 999 });
    assert.deepEqual(parseByteRange("bytes=0-", SIZE), { start: 0, end: 999 });
});
