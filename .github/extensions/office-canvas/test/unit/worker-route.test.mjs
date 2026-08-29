// The /vendor/ routes: what the iframe actually receives.
//
// Office-free. The viewer instance is stood up with a stub cache because none of
// these routes touch Word -- they read committed files off disk, which is the
// whole reason this layer can be tested on a hosted runner.

import { test, after } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { cp, mkdtemp, readdir, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { ViewerInstance } from "../../src/server.mjs";
import { readManifest, resetWorkerCache, VENDOR_DIR } from "../../src/vendor-assets.mjs";

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");

/** Enough of a RenderCache for routes that never reach Word. */
const stubCache = { wordVersion: "stub", async open() {}, async pdf() {}, async refresh() {} };

const servers = [];
const staged = [];
after(async () => {
    for (const server of servers) await server.close();
    for (const dir of staged) await rm(dir, { recursive: true, force: true });
    resetWorkerCache();
});

async function viewer(options = {}) {
    const instance = new ViewerInstance({ cache: stubCache, instanceId: "test", workspacePath: null, ...options });
    servers.push(instance);
    return { base: await instance.start(), instance };
}

/** A throwaway copy of the committed vendor directory, safe to damage. */
async function stageVendor() {
    const dir = await mkdtemp(path.join(tmpdir(), "vendor-route-"));
    staged.push(dir);
    await cp(VENDOR_DIR, dir, { recursive: true });
    return dir;
}

/** Fetches and returns the response plus its body as a Buffer. */
async function get(base, route, init) {
    const res = await fetch(new URL(route, base), init);
    return { res, body: Buffer.from(await res.arrayBuffer()) };
}

test("the worker route serves the exact bytes the manifest describes", async () => {
    const { base } = await viewer();
    const manifest = await readManifest();
    const { res, body } = await get(base, "/vendor/pdf.worker.min.mjs");

    assert.equal(res.status, 200);
    assert.equal(body.byteLength, manifest.worker.bytes);
    assert.equal(sha256(body), manifest.worker.sha256);
    assert.equal(Number(res.headers.get("content-length")), manifest.worker.bytes);
    assert.match(res.headers.get("content-type"), /javascript/);
});

test("the worker is larger than any single file the installer would accept", async () => {
    // The reason this route exists at all. If the worker ever fits in one file
    // the split is dead weight -- and if this assertion is what tells us, it
    // will say so rather than the split silently continuing to be carried.
    const manifest = await readManifest();
    assert.ok(
        manifest.worker.bytes > 1_000_000,
        "the worker now fits the per-file cap; the split and this route can go",
    );
    for (const part of manifest.worker.parts) {
        assert.ok(part.bytes <= 1_000_000, `${part.name} does not fit the cap`);
    }
});

test("the worker route answers HEAD with the length and no body", async () => {
    const { base } = await viewer();
    const manifest = await readManifest();
    const { res, body } = await get(base, "/vendor/pdf.worker.min.mjs", { method: "HEAD" });
    assert.equal(res.status, 200);
    assert.equal(Number(res.headers.get("content-length")), manifest.worker.bytes);
    assert.equal(body.byteLength, 0);
});

test("a matching If-None-Match gets a 304 rather than 1.26 MB again", async () => {
    const { base } = await viewer();
    const first = await get(base, "/vendor/pdf.worker.min.mjs");
    const etag = first.res.headers.get("etag");
    assert.ok(etag, "the worker response must carry an ETag to revalidate against");

    const { res, body } = await get(base, "/vendor/pdf.worker.min.mjs", { headers: { "If-None-Match": etag } });
    assert.equal(res.status, 304);
    assert.equal(body.byteLength, 0);
});

test("a stale ETag gets the body, not a 304", async () => {
    // The discriminating case. A route that answered 304 to everything would
    // pass the test above and serve last version's worker forever.
    const { base } = await viewer();
    const { res, body } = await get(base, "/vendor/pdf.worker.min.mjs", {
        headers: { "If-None-Match": '"not-the-current-worker"' },
    });
    assert.equal(res.status, 200);
    assert.ok(body.byteLength > 0);
});

test("the worker is not cached immutably, because its URL carries no version", async () => {
    const { base } = await viewer();
    const { res } = await get(base, "/vendor/pdf.worker.min.mjs");
    const cacheControl = res.headers.get("cache-control") ?? "";
    assert.ok(
        !cacheControl.includes("immutable"),
        "an immutable worker at a versionless URL outlives the pdf.min.mjs it was built with",
    );
    assert.match(cacheControl, /must-revalidate/);
});

test("a vendor read failure is a typed 500, not a truncated 200", async () => {
    // The failure this route is shaped around: parts streamed as they are read
    // would put a `200` and two thirds of a file on the wire before the third
    // read failed, and the client would receive a syntactically ruined worker
    // with no indication anything went wrong.
    //
    // Injected against a staged copy of the vendor directory rather than the
    // committed one, so the test cannot leave the working tree broken for
    // whatever runs next.
    const staged = await stageVendor();
    const manifest = JSON.parse(await readFile(path.join(staged, "pdfjs.manifest.json"), "utf8"));
    await rm(path.join(staged, manifest.worker.parts.at(-1).name));

    const { base } = await viewer({ vendorDir: staged });
    const { res, body } = await get(base, "/vendor/pdf.worker.min.mjs");

    assert.equal(res.status, 500);
    const payload = JSON.parse(body.toString("utf8"));
    assert.equal(payload.error.code, "vendor_incomplete");
    assert.match(payload.error.message, /vendor-pdfjs/, "the message must name the script that fixes it");
});

test("a missing static vendor file is a typed 500, not an empty 200", async () => {
    const staged = await stageVendor();
    await rm(path.join(staged, "pdf-text-layer.css"));

    const { base } = await viewer({ vendorDir: staged });
    const { res, body } = await get(base, "/vendor/pdf-text-layer.css");

    assert.equal(res.status, 500);
    assert.equal(JSON.parse(body.toString("utf8")).error.code, "vendor_missing");
});

test("only the allowlisted vendor files are reachable", async () => {
    const { base } = await viewer();
    const manifest = await readManifest();

    // A part on its own is a fragment of a program. Serving one would make a
    // truncated worker look like a working route.
    for (const part of manifest.worker.parts) {
        const { res } = await get(base, `/vendor/${part.name}`);
        assert.equal(res.status, 404, `${part.name} must not be individually reachable`);
    }
    const { res: manifestRes } = await get(base, "/vendor/pdfjs.manifest.json");
    assert.equal(manifestRes.status, 404);
});

test("a traversal in the vendor path cannot escape the vendor directory", async () => {
    const { base } = await viewer();
    // Encoded so the URL parser does not normalise it away before we see it.
    const { res } = await get(base, "/vendor/..%2F..%2Fextension.mjs");
    assert.equal(res.status, 404);
});

test("pdf.min.mjs is served as JavaScript, byte for byte", async () => {
    const { base } = await viewer();
    const { res, body } = await get(base, "/vendor/pdf.min.mjs");
    assert.equal(res.status, 200);
    assert.match(res.headers.get("content-type"), /javascript/);
    assert.deepEqual(body, await readFile(path.join(VENDOR_DIR, "pdf.min.mjs")));
});

test("the text-layer stylesheet is served as CSS", async () => {
    // Served as text/css or the browser drops it, and the text layer then sits
    // at its intrinsic size with every span at the wrong offset -- while the
    // page image still renders correctly.
    const { base } = await viewer();
    const { res, body } = await get(base, "/vendor/pdf-text-layer.css");
    assert.equal(res.status, 200);
    assert.match(res.headers.get("content-type"), /^text\/css/);
    assert.deepEqual(body, await readFile(path.join(VENDOR_DIR, "pdf-text-layer.css")));
});

test("every absolute-URL import the UI makes is a route the server serves", async () => {
    // C3's import check only looks at relative specifiers, so `/vendor/...` is
    // invisible to it -- a rename would produce a blank canvas with every gate
    // still green. This closes that gap by asking the server itself.
    const { base } = await viewer();
    const uiDir = path.join(VENDOR_DIR, "..");
    // Read the directory rather than name the files: a new UI module importing a
    // vendored asset is the case a fixed list would miss, and it is exactly the
    // case this test exists for.
    const sources = (await readdir(uiDir, { withFileTypes: true }))
        .filter((entry) => entry.isFile() && /\.(mjs|js|html|css)$/.test(entry.name))
        .map((entry) => entry.name);
    assert.ok(sources.length >= 4, `expected UI sources to enumerate, saw ${sources.join(", ")}`);
    const referenced = new Set();
    for (const name of sources) {
        const text = await readFile(path.join(uiDir, name), "utf8");
        for (const match of text.matchAll(/["'(](\/vendor\/[A-Za-z0-9._-]+)["')]/g)) {
            referenced.add(match[1]);
        }
    }
    assert.ok(referenced.size >= 3, `expected the viewer to reference the vendored files, saw ${[...referenced]}`);
    for (const route of referenced) {
        const { res } = await get(base, route, { method: "HEAD" });
        assert.equal(res.status, 200, `${route} is imported by the UI but the server does not serve it`);
    }
});
