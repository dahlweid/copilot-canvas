// Getting the Word mark off this machine's own Word, and coping when there is
// none (#68).
//
// Office-free by construction: the extractor is injected, because CI has neither
// Word nor PowerShell and the parts worth measuring -- the caching, the refusal
// of a non-PNG, the failure path -- are the same on every machine.

import { test } from "node:test";
import assert from "node:assert/strict";

import { createWordIconSource, decodeIcon } from "../../src/word/word-icon.mjs";

/** A minimal but real PNG: the 8-byte signature and a little body. */
const PNG = Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), Buffer.from("IHDR-ish")]);
const PNG_B64 = PNG.toString("base64");

test("the icon is extracted once, however many times it is asked for", async () => {
    // The requirement in #68's own words -- extract once, cache in memory, do
    // not re-extract per request. A route serving a 32px decoration must not
    // spawn PowerShell each time a panel opens.
    let runs = 0;
    const source = createWordIconSource({
        run: async () => {
            runs += 1;
            return PNG_B64;
        },
    });

    const first = await source.get();
    const second = await source.get();

    assert.equal(runs, 1, `expected one extraction, got ${runs}`);
    assert.deepEqual(first.buffer, PNG);
    assert.equal(second.etag, first.etag, "the same icon produced two different ETags");
});

test("two requests arriving together share one extraction", async () => {
    // The promise is memoized rather than the value, which is the difference
    // between one spawn and two when a panel makes both requests at once --
    // which it does: the bar wires two placements on the same tick.
    let runs = 0;
    const source = createWordIconSource({
        run: async () => {
            runs += 1;
            await new Promise((r) => setTimeout(r, 5));
            return PNG_B64;
        },
    });

    await Promise.all([source.get(), source.get(), source.get()]);

    assert.equal(runs, 1, `three concurrent callers caused ${runs} extractions`);
});

test("a machine with no Word is remembered as such, not asked again", async () => {
    // The expensive mistake this guards: memoizing only success would re-spawn
    // PowerShell on every request forever, on exactly the machines that can
    // never answer.
    let runs = 0;
    const source = createWordIconSource({
        run: async () => {
            runs += 1;
            return null;
        },
    });

    assert.equal(await source.get(), null);
    assert.equal(await source.get(), null);
    assert.equal(runs, 1, `a missing Word was re-extracted ${runs} times`);
});

test("an extractor that throws yields no icon rather than an error", async () => {
    // Nothing waits on this and nobody should be told. A rejection here would
    // reach the route as a 500 and, through it, the user.
    const source = createWordIconSource({
        run: async () => {
            throw new Error("powershell.exe is not on this machine");
        },
    });

    assert.equal(await source.get(), null);
});

test("output that is not a PNG is refused", async () => {
    // `-NoProfile` silences a profile, but a machine policy banner or a
    // transcription notice can still land on stdout ahead of the payload, and
    // base64 will happily decode it. Serving that as image/png would draw a
    // broken image, which is worse than drawing none.
    assert.equal(decodeIcon(Buffer.from("Transcript started, output file is ...").toString("base64")), null);
    assert.equal(decodeIcon(""), null);
    assert.equal(decodeIcon("   "), null);
    assert.equal(decodeIcon(null), null);
    // The signature alone is not an image either.
    assert.equal(decodeIcon(PNG.subarray(0, 8).toString("base64")), null);

    const good = decodeIcon(PNG_B64);
    assert.ok(good, "a real PNG was refused");
    assert.match(good.etag, /^"[0-9a-f]{32}"$/, "the ETag is not a quoted digest");
});

test("different icons get different ETags", async () => {
    // The guard on the caching test above: an ETag that is constant would make
    // "the same icon produced the same ETag" pass while measuring nothing.
    const other = Buffer.concat([PNG, Buffer.from("more")]);

    assert.notEqual(decodeIcon(PNG_B64).etag, decodeIcon(other.toString("base64")).etag);
});

test("a timeout leaves the caller with no icon rather than a hung request", async () => {
    // `ExtractAssociatedIcon` reads an executable's resources and does not
    // launch Word, but it is still a spawned process on a machine that may be
    // busy, and the panel behind this request is already drawn.
    const source = createWordIconSource({
        timeoutMs: 5,
        run: ({ signal }) =>
            new Promise((resolve) => {
                signal.addEventListener("abort", () => resolve(null), { once: true });
            }),
    });

    assert.equal(await source.get(), null);
});
