// A missing or blank `path` argument, at the boundary the agent actually sees.
// Office-free.
//
// ## What this pins
//
// Every tool declares `required: ["path"]`, and the host enforces no schema
// keyword before dispatch -- measured in #28: `required`, `enum`, `type`,
// `minItems`, `maximum` and `additionalProperties` were each violated and every
// one reached the handler. So the declaration is documentation for the model and
// the handler is the only defence.
//
// The handler did refuse a missing `path`, by accident. `resolveInputPath`
// called `path.isAbsolute(undefined)`, and Node threw about its own argument
// contract. Measured through these very handlers before the fix, identically
// for all four tools:
//
//     code=ERR_INVALID_ARG_TYPE
//     message="ERR_INVALID_ARG_TYPE: The \"path\" argument must be of type string.
//              Received undefined"
//
// That message cannot name the tool's parameter, cannot say what to send
// instead, and is not a code any layer would branch on. #38.
//
// ## Why every assertion here pins a code *and* a message
//
// The pre-fix build already refuses these calls. A test asserting only that the
// handler rejects passes identically against the broken and the fixed code and
// measures nothing at all. So each assertion names the typed code and the
// sentence, and additionally asserts the Node code is *absent*.
//
// ## Mutations used to confirm these can go red
//
//   A. extension.mjs, `resolveInputPath` -> restore the body from main
//      (`if (path.isAbsolute(input) || ...)` with no guard). This is the defect
//      itself. Measured over the whole unit suite: 435 pass, 8 fail -- the eight
//      failures are every test in this file except the picker one, and no
//      pre-existing test moved, which is the concrete form of "nothing on main
//      pins this".
//   B. extension.mjs, `resolveInputPath` -> keep the guard but throw
//      `invalid_request` instead of `invalid_path`. Measured, this file: 3 pass,
//      6 fail. This is what pins the code rather than the fact of a typed
//      throw; the two survivors reach the second throw, which mutation B leaves
//      alone.
//   C. extension.mjs, the canvas `open` -> `if (ctx.input)` instead of
//      `if (ctx.input?.path)`, which is how refusing a blank path could
//      plausibly break the document picker. Measured, this file: 7 pass, 2 fail.
//
// The picker test is the one that stays green under A, and deliberately so: it
// does not test the fix, it guards the regression the fix could cause, and C is
// its proof. Every other test here is red under A.

import test, { after, before } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { loadExtension } from "./extension-stubs.mjs";

let home;
let sdk;
let canvas;

const toolNamed = (name) => {
    const tool = sdk.joined.tools.find((t) => t.name === name);
    assert.ok(tool, `${name} must be registered`);
    return tool;
};

before(async () => {
    // artifactsRoot() is read when a cache is constructed; keep it out of the
    // user's own COPILOT_HOME even though nothing here should build one.
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-path-args-"));
    process.env.COPILOT_HOME = home;
    sdk = await loadExtension();
    canvas = sdk.joined.canvases[0];
    assert.ok(canvas?.open, "the extension must have registered the word-doc canvas");
});

const openPanels = new Set();

after(async () => {
    for (const instanceId of openPanels) await canvas.onClose({ instanceId }).catch(() => {});
    if (home) await rm(home, { recursive: true, force: true });
});

async function openPanel(instanceId, input) {
    // Registered before the await, so a rejection still leaves a panel `after()`
    // closes -- an unclosed panel holds a listening server and the run never
    // exits.
    openPanels.add(instanceId);
    return canvas.open({ instanceId, input });
}

/** Asserts a tool refusal: the typed code, the sentence, and no Node code. */
function assertTypedRefusal(err, expected) {
    assert.equal(err.code, "invalid_path", `expected invalid_path, got ${err.code}`);
    assert.equal(err.message, `invalid_path: ${expected}`);
    assert.doesNotMatch(err.message, /ERR_INVALID_ARG_TYPE/);
    return true;
}

const REQUIRED = "`path` is required: give an absolute path to the document, or one relative to the workspace.";

const notAString = (received) =>
    "`path` must be a non-empty string — an absolute path to the document, or one relative to the workspace. " +
    `Received ${received}.`;

// Otherwise-valid arguments for each tool, with `path` left out. Everything else
// is present so the refusal cannot be about a different field.
const TOOLS_WITHOUT_PATH = [
    ["read_document", { limit: 10, offset: 0 }],
    ["create_document", { blocks: [{ kind: "paragraph", text: "Hello." }] }],
    ["edit_document", { op: "replace_text", address: "p:0123456789ab", revisionToken: "abc", text: "Hello." }],
    ["revert_document", { revisionToken: "abc" }],
];

for (const [name, args] of TOOLS_WITHOUT_PATH) {
    test(`${name} refuses a missing path with a typed invalid_path`, async () => {
        await assert.rejects(
            () => toolNamed(name).handler(args),
            (err) => assertTypedRefusal(err, REQUIRED),
        );
    });
}

test("a null path is the same refusal as an absent one", async () => {
    // `args.path === null` is what an agent emitting a JSON null produces, and
    // it reached `path.isAbsolute` exactly as `undefined` did.
    await assert.rejects(
        () => toolNamed("read_document").handler({ path: null }),
        (err) => assertTypedRefusal(err, REQUIRED),
    );
});

test("a path of the wrong type names the type it got", async () => {
    await assert.rejects(
        () => toolNamed("read_document").handler({ path: 42 }),
        (err) => assertTypedRefusal(err, notAString("a number")),
    );
});

test("an empty path is refused rather than resolved to the workspace directory", async () => {
    // `path.join(workspacePath, "")` is the workspace directory -- measured --
    // so a blank path resolved to a folder and was refused, if at all, further
    // down and under a code naming something else.
    for (const blank of ["", "   ", "\t"]) {
        await assert.rejects(
            () => toolNamed("read_document").handler({ path: blank }),
            (err) => assertTypedRefusal(err, notAString("an empty string")),
            `expected a typed refusal for ${JSON.stringify(blank)}`,
        );
    }
});

// --- the canvas, which shares `resolveInputPath` ---------------------------

test("opening the canvas with no path still opens the document picker", async () => {
    // The one thing refusing a blank path could plausibly break: `open` tests
    // `ctx.input?.path` for truthiness first, so an omitted path must never
    // reach the guard.
    const result = await openPanel("picker", {});
    assert.equal(result.title, "Word document");
    assert.equal(result.status, "No document open");
    assert.match(result.url, /^http:\/\/127\.0\.0\.1:\d+\//);
});

test("the open_document action refuses a missing path with a typed invalid_path", async () => {
    await openPanel("action-panel", {});
    const action = canvas.actions.find((a) => a.name === "open_document");
    assert.ok(action, "open_document must be registered");

    await assert.rejects(
        () => action.handler({ instanceId: "action-panel", input: {} }),
        (err) => {
            // A canvas action carries the code as a field rather than folded
            // into the message, so this is the same refusal in the other shape.
            assert.equal(err.name, "CanvasError");
            assert.equal(err.code, "invalid_path");
            assert.equal(err.message, REQUIRED);
            assert.doesNotMatch(err.message, /ERR_INVALID_ARG_TYPE/);
            return true;
        },
    );
});
