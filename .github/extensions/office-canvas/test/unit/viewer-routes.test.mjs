// The viewer's own HTTP routes, as a client actually receives them.
//
// Office-free: the instance is stood up with a stub cache, because none of the
// routes asserted here reach Word. What this layer can see and a source-reading
// test cannot is the *served* answer -- a route left registered replies whatever
// the switch statement looks like to a reader.

import { test, after } from "node:test";
import assert from "node:assert/strict";

import { ViewerInstance } from "../../src/server.mjs";

/** Enough of a RenderCache for routes that never reach Word. */
const stubCache = { wordVersion: "stub", async open() {}, async pdf() {}, async refresh() {} };

const servers = [];
after(async () => {
    for (const server of servers) await server.close();
});

async function viewer(options = {}) {
    const instance = new ViewerInstance({ cache: stubCache, instanceId: "test", workspacePath: null, ...options });
    servers.push(instance);
    return { base: await instance.start(), instance };
}

test("the picker's browse route is gone", async () => {
    // #69. `/api/browse` listed the workspace and the recents, and existed only
    // to fill the picker. Both went with it.
    const { base } = await viewer({ workspacePath: process.cwd() });
    const res = await fetch(new URL("/api/browse", base));

    assert.equal(res.status, 404);
    assert.equal((await res.json()).error.code, "not_found");
});

test("a route that is still wired answers, so the 404 above means something", async () => {
    // The guard on the test above: if the server had stopped serving anything
    // at all -- failed to start, or answered 404 to everything -- that
    // assertion would pass while measuring nothing.
    const { base } = await viewer();
    const res = await fetch(new URL("/api/state", base));

    assert.equal(res.status, 200);
    assert.equal((await res.json()).doc, null);
});

// --- The Word mark over HTTP (#68) -------------------------------------------
//
// The icon source is injected throughout, so none of this needs Word, or
// PowerShell, or Windows. What it does need is the real route, because the
// caching contract lives in headers.

/** An icon source that answers with fixed bytes, standing in for the extractor. */
const iconSource = (icon) => ({ async get() { return icon; } });

const PNG = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x01, 0x02]);
const ICON = { buffer: PNG, etag: '"0123456789abcdef0123456789abcdef"' };

test("the Word mark is served as a PNG a browser can cache", async () => {
    const { base } = await viewer({ wordIcon: iconSource(ICON) });
    const res = await fetch(new URL("/api/word-icon", base));

    assert.equal(res.status, 200);
    assert.equal(res.headers.get("content-type"), "image/png");
    assert.equal(res.headers.get("etag"), ICON.etag);
    // Revalidated rather than pinned: the URL carries no version, so a
    // far-future expiry would hold a browser to whichever Word was installed
    // the first time it asked.
    assert.match(res.headers.get("cache-control"), /must-revalidate/);
    assert.deepEqual(Buffer.from(await res.arrayBuffer()), PNG);
});

test("a browser that already has the mark is told so rather than sent it again", async () => {
    const { base } = await viewer({ wordIcon: iconSource(ICON) });
    const res = await fetch(new URL("/api/word-icon", base), { headers: { "If-None-Match": ICON.etag } });

    assert.equal(res.status, 304);
    assert.equal((await res.arrayBuffer()).byteLength, 0);
});

test("a stale ETag gets the mark, so the 304 above is conditional on something", async () => {
    const { base } = await viewer({ wordIcon: iconSource(ICON) });
    const res = await fetch(new URL("/api/word-icon", base), { headers: { "If-None-Match": '"not-this-one"' } });

    assert.equal(res.status, 200);
    assert.equal((await res.arrayBuffer()).byteLength, PNG.length);
});

test("a machine with no Word answers 404 and says nothing else", async () => {
    // The degradation contract. The <img> errors, the bar keeps its drawn
    // glyph, and no status, banner or log line reaches the user -- a missing
    // decoration that announced itself would be worse than the missing
    // decoration. A 500 here would be an error the panel had to show.
    const logged = [];
    const { base } = await viewer({ wordIcon: iconSource(null), log: (line) => logged.push(line) });
    const res = await fetch(new URL("/api/word-icon", base));

    assert.equal(res.status, 404);
    assert.equal((await res.json()).error.code, "not_found");
    assert.deepEqual(logged, [], `a missing icon was reported: ${logged.join(" | ")}`);
});
