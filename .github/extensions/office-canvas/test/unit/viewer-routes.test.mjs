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
