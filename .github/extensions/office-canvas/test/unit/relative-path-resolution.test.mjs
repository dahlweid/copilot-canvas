// Relative paths, at the boundary the agent and the canvas actually see.
// Office-free -- the resolution decision is made before any Word work, which is
// why #158 reproduces with no Word running.
//
// ## What this pins (#158)
//
// `resolveInputPath` turns the caller's `path` into the path Word is handed. A
// relative path is only meaningful against a root. The original shape was:
//
//     if (path.isAbsolute(input) || !session?.workspacePath) return normalizeDocPath(input);
//     return normalizeDocPath(path.join(session.workspacePath, input));
//
// Two things were wrong with it. The `|| !session?.workspacePath` arm read as
// "decline" but did not: it fell through to `normalizeDocPath(input)`, whose
// `path.resolve` (render-cache.mjs) is **cwd-relative**, so a workspace-less
// relative path was silently resolved against the extension process's working
// directory -- a root the caller never named. And even on its happy arm it
// joined against `session.workspacePath`, which in a CLI session is the
// *session-state* directory, not the project root a caller means by a relative
// path. Both roots are wrong; only an absolute path worked.
//
// The fix (option 2) resolves a relative path against the session's working
// directory, read live from the host via `session.rpc.metadata.snapshot()` --
// the field measured to be the worktree/project root. There is no fallback: if
// the snapshot cannot be read, or reports no usable absolute working directory,
// the relative path is *refused* with a typed `working_directory_unavailable`,
// never resolved against cwd or `workspacePath`. An absolute path never consults
// the snapshot at all.
//
// ## Why a tool assertion reads a returned value and a canvas one a throw
//
// The two channels differ and were measured separately (#45), the same way
// tool-path-args.test.mjs pins them: a tool that throws reaches the agent as the
// bare `Tool execution failed`, so its refusal has to be *returned*; a canvas
// action's thrown message survives but its `code` field does not, so its refusal
// is thrown with the code folded into the message.
//
// ## Negative controls (each must go red for the reason its test names)
//
// Three independent breakages, each caught by a different assertion here:
//
//   * **Revert to cwd resolution** (drop the snapshot read, fall through to
//     `normalizeDocPath(input)` for a relative path). The "resolves against the
//     working directory" tests then observe the path joined onto *cwd*, not the
//     reported working directory, so the exact-path assertion flips green->red;
//     and the two refusal groups stop refusing and answer instead, so their
//     `assertTypedRefusal`/`assert.rejects` flip. This is the #158 defect itself,
//     caught in the act rather than argued about.
//   * **Add a fallback on snapshot failure** (e.g. `process.cwd()` or
//     `workspacePath`). Only the snapshot-rejected group flips refusal->success,
//     which is condition 1 (no silent fallback) failing.
//   * **Remove the field guard** (let a non-absolute `workingDirectory` through).
//     Only the field-guard group flips: `path.join("", REL)` reaches
//     `normalizeDocPath` and resolves cwd-relative, a success where a refusal
//     was required.
//
// The refusal groups depend on the stub cache implementing create/edit/revert as
// well as open/read (stub-render-cache.mjs). Without those three methods the
// mutating tools would go red under a broken resolver for an *unrelated* reason
// -- a missing-method "is not a function", not a wrong-root resolution -- which
// is red for the wrong reason, the exact thing a control must exclude (round 1
// of #161 caught precisely this gap). Confirmed by running each breakage: the
// named group goes red via the described flip, restored to green by the fix.
// Recorded in the PR.

import test, { after, before, beforeEach } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { loadExtension } from "./extension-stubs.mjs";
import { built, normalizeDocPath } from "./stub-render-cache.mjs";

let home;
let sdk;
let canvas;

const toolNamed = (name) => {
    const tool = sdk.joined.tools.find((t) => t.name === name);
    assert.ok(tool, `${name} must be registered`);
    return tool;
};

// A relative path, given with forward slashes so `path.isAbsolute` is false on
// every platform the unit suite runs on.
const REL = "sub/probe-158.docx";

// The two refusal sentences resolveInputPath emits, both under the single
// `working_directory_unavailable` code and discriminated structurally (the
// snapshot rejected, versus its field was unusable), never on the message text.
const CODE = "working_directory_unavailable";
const snapshotUnreadable = (rel) =>
    `Cannot resolve the relative path ${JSON.stringify(rel)}: the session's working directory could not be read. ` +
    `Give an absolute path instead.`;
const noUsableWorkingDir = (rel) =>
    `Cannot resolve the relative path ${JSON.stringify(rel)}: the session reported no usable working directory. ` +
    `Give an absolute path instead.`;

/** Every document path any stub cache has been asked to open or read. */
const allDocPaths = () => built.flatMap((c) => c.docPaths);

before(async () => {
    // artifactsRoot() is read when a cache is constructed; keep it out of the
    // user's own COPILOT_HOME even though nothing here should build a real one.
    home = await mkdtemp(path.join(tmpdir(), "office-canvas-relpath-"));
    process.env.COPILOT_HOME = home;
    sdk = await loadExtension();
    canvas = sdk.joined.canvases[0];
    assert.ok(canvas?.open, "the extension must have registered the word-doc canvas");
});

// Each test declares the snapshot state it needs. Reset to the defaults between
// tests -- no workspace, no working directory, no snapshot error -- rather than
// leaking one test's state into the next.
beforeEach(() => {
    sdk.setWorkspacePath(null);
    sdk.resetMetadata();
});

const openPanels = new Set();

after(async () => {
    for (const instanceId of openPanels) await canvas.onClose({ instanceId }).catch(() => {});
    sdk.setWorkspacePath(null);
    sdk.resetMetadata();
    if (home) await rm(home, { recursive: true, force: true });
});

async function openPanel(instanceId, input) {
    openPanels.add(instanceId);
    return canvas.open({ instanceId, input });
}

/** Asserts a tool refusal carrying a typed code and its sentence, no Node code. */
function assertTypedRefusal(result, code, sentence) {
    assert.equal(
        result?.resultType,
        "failure",
        `expected a failure result, got ${JSON.stringify(result)?.slice(0, 160)}`,
    );
    const text = result.textResultForLlm;
    assert.match(text, new RegExp(code), `expected ${code} in:\n${text}`);
    assert.ok(text.includes(sentence), `expected the sentence in:\n${text}`);
    assert.doesNotMatch(text, /ERR_INVALID_ARG_TYPE/);
}

// Otherwise-valid arguments for each tool, so the only thing under test is how
// the relative `path` is resolved (or refused).
const TOOLS_WITH_RELATIVE = [
    ["read_document", { path: REL, limit: 10, offset: 0 }],
    ["create_document", { path: REL, blocks: [{ kind: "paragraph", text: "Hello." }] }],
    ["edit_document", { path: REL, op: "replace_text", address: "p:0123456789ab", revisionToken: "abc", text: "Hi." }],
    ["revert_document", { path: REL, revisionToken: "abc" }],
];

// --- Resolution against the working directory (the #158 core) -----------------

for (const [name, args] of TOOLS_WITH_RELATIVE) {
    test(`${name} resolves a relative path against the session's working directory`, async () => {
        sdk.setWorkingDirectory(home);
        const before = allDocPaths().length;

        const result = await toolNamed(name).handler(args);

        assert.notEqual(
            result?.resultType,
            "failure",
            `expected success, got ${JSON.stringify(result)?.slice(0, 160)}`,
        );
        const resolved = allDocPaths().slice(before);
        assert.equal(resolved.length, 1, "exactly one document path should have reached the cache");
        // The working directory joined with the relative input -- and nothing
        // cwd-ish. Reverting to cwd resolution makes this the wrong path.
        assert.equal(resolved[0], normalizeDocPath(path.join(home, REL)));
    });
}

// --- Refusal when the snapshot cannot be read (no fallback, condition 1) -------

for (const [name, args] of TOOLS_WITH_RELATIVE) {
    test(`${name} refuses a relative path when the metadata snapshot fails, without resolving it`, async () => {
        sdk.failMetadataSnapshot();
        const before = allDocPaths().length;

        assertTypedRefusal(await toolNamed(name).handler(args), CODE, snapshotUnreadable(REL));
        // Structural, not incidental: the path was never handed to a cache at
        // all, so it cannot have been resolved against cwd (or anything else).
        assert.deepEqual(
            allDocPaths().slice(before),
            [],
            "a declined relative path must never reach the render cache",
        );
    });
}

// --- Refusal when the snapshot reports no usable working directory (field guard) -

for (const [name, args] of TOOLS_WITH_RELATIVE) {
    test(`${name} refuses a relative path when the working directory is not an absolute string`, async () => {
        // A snapshot that answers, but with an empty (non-absolute) directory:
        // without the field guard this reaches `path.join("", REL)` and resolves
        // cwd-relative -- the #158 bug through a different door.
        sdk.setWorkingDirectory("");
        const before = allDocPaths().length;

        assertTypedRefusal(await toolNamed(name).handler(args), CODE, noUsableWorkingDir(REL));
        assert.deepEqual(
            allDocPaths().slice(before),
            [],
            "a declined relative path must never reach the render cache",
        );
    });
}

// --- The canvas open action: same three scenarios, on the throw channel -------

test("the open_document action resolves a relative path against the working directory", async () => {
    sdk.setWorkingDirectory(home);
    await openPanel("relpath-ok", {});
    const action = canvas.actions.find((a) => a.name === "open_document");
    assert.ok(action, "open_document must be registered");
    const before = allDocPaths().length;

    await action.handler({ instanceId: "relpath-ok", input: { path: REL } });

    const resolved = allDocPaths().slice(before);
    assert.equal(resolved.length, 1, "exactly one document path should have reached the cache");
    assert.equal(resolved[0], normalizeDocPath(path.join(home, REL)));
});

test("the open_document action refuses a relative path when the metadata snapshot fails", async () => {
    sdk.failMetadataSnapshot();
    await openPanel("relpath-fail", {});
    const action = canvas.actions.find((a) => a.name === "open_document");
    assert.ok(action, "open_document must be registered");
    const before = allDocPaths().length;

    await assert.rejects(
        () => action.handler({ instanceId: "relpath-fail", input: { path: REL } }),
        (err) => {
            // The canvas channel: measured, a thrown message survives but the
            // `code` field does not, so the code is folded into the message.
            assert.equal(err.name, "CanvasError");
            assert.equal(err.code, CODE);
            assert.equal(err.message, `${CODE}: ${snapshotUnreadable(REL)}`);
            assert.doesNotMatch(err.message, /ERR_INVALID_ARG_TYPE/);
            return true;
        },
    );
    assert.deepEqual(allDocPaths().slice(before), [], "a declined relative path must never reach the render cache");
});

test("the open_document action refuses a relative path when the working directory is unusable", async () => {
    sdk.setWorkingDirectory("");
    await openPanel("relpath-guard", {});
    const action = canvas.actions.find((a) => a.name === "open_document");
    assert.ok(action, "open_document must be registered");
    const before = allDocPaths().length;

    await assert.rejects(
        () => action.handler({ instanceId: "relpath-guard", input: { path: REL } }),
        (err) => {
            assert.equal(err.name, "CanvasError");
            assert.equal(err.code, CODE);
            assert.equal(err.message, `${CODE}: ${noUsableWorkingDir(REL)}`);
            return true;
        },
    );
    assert.deepEqual(allDocPaths().slice(before), [], "a declined relative path must never reach the render cache");
});

// --- Absolute paths never consult the snapshot --------------------------------

test("an absolute path is passed through unchanged, even when the snapshot would fail", async () => {
    const absolute = path.join(home, "abs-probe-158.docx");
    assert.ok(path.isAbsolute(absolute));

    // Snapshot set to reject: an absolute path must resolve anyway, proving it
    // never reads the working directory. A resolver that consulted the snapshot
    // for every path would refuse here.
    sdk.failMetadataSnapshot();
    let before = allDocPaths().length;
    let result = await toolNamed("read_document").handler({ path: absolute });
    assert.notEqual(result?.resultType, "failure", "an absolute path must resolve even when the snapshot fails");
    assert.equal(allDocPaths().slice(before)[0], normalizeDocPath(absolute));

    // With a working directory set, the absolute path is still itself, not
    // joined onto it.
    sdk.setWorkingDirectory(home);
    before = allDocPaths().length;
    result = await toolNamed("read_document").handler({ path: absolute });
    assert.notEqual(result?.resultType, "failure", "an absolute path must resolve with a working directory set");
    assert.equal(allDocPaths().slice(before)[0], normalizeDocPath(absolute));
});
