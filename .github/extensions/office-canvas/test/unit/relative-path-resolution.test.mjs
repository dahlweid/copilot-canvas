// Relative paths, at the boundary the agent and the canvas actually see.
// Office-free -- the resolution decision is made before any Word work, which is
// why #158 reproduces with no Word running.
//
// ## What this pins (#158)
//
// `resolveInputPath` turns the caller's `path` into the path Word is handed. A
// relative path is only meaningful against a root, and the only root the
// extension has is `session.workspacePath`. The prior shape was:
//
//     if (path.isAbsolute(input) || !session?.workspacePath) return normalizeDocPath(input);
//     return normalizeDocPath(path.join(session.workspacePath, input));
//
// The `|| !session?.workspacePath` arm reads as "decline", but it does not
// decline: it falls through to `normalizeDocPath(input)`, whose `path.resolve`
// (render-cache.mjs) is **cwd-relative**. So a relative path with no
// workspacePath was silently resolved against the extension process's working
// directory -- a root the caller never named, and neither the workspace nor an
// absolute path. The measured symptom was a lookup under the session-state
// directory (the issue's reproduction), not the workspace root.
//
// The fix splits that arm off into an honest refusal: relative input with no
// workspacePath throws a typed `workspace_unavailable` naming the one cause the
// branch established -- no root to resolve against -- and saying an absolute
// path is required. It is structurally incapable of resolving against cwd,
// because the only path that now reaches `normalizeDocPath` from a relative
// input is one joined onto a workspace root; the cwd branch is gone.
//
// ## Why a tool assertion reads a returned value and a canvas one a throw
//
// The two channels differ and were measured separately (#45), the same way
// tool-path-args.test.mjs pins them: a tool that throws reaches the agent as the
// bare `Tool execution failed`, so its refusal has to be *returned*; a canvas
// action's thrown message survives but its `code` field does not, so its refusal
// is thrown with the code folded into the message.
//
// ## Negative control (must be able to go red)
//
// Reverting the fix -- restoring `|| !session?.workspacePath` to the first line
// so a workspace-less relative path resolves against cwd again -- flips the
// falsy-branch tests here from a returned/thrown `workspace_unavailable` refusal
// to a *success* (the stub cache resolves and answers). failure -> success and
// throw -> resolve are both hard flips, so these tests cannot pass against the
// #158 defect. Confirmed by running exactly that reversion; recorded in the PR.

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

const workspaceUnavailable = (rel) =>
    `Cannot resolve the relative path ${JSON.stringify(rel)}: no workspace directory is available in this session, ` +
    `so there is no root to resolve it against. Give an absolute path instead.`;

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

// Each test declares the workspace it needs; the default is the no-workspace
// case (#158), so reset to it between tests rather than leaking one test's
// workspacePath into the next.
beforeEach(() => {
    sdk.setWorkspacePath(null);
});

const openPanels = new Set();

after(async () => {
    for (const instanceId of openPanels) await canvas.onClose({ instanceId }).catch(() => {});
    sdk.setWorkspacePath(null);
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

// Otherwise-valid arguments for each tool, so the only thing wrong is the
// relative `path` arriving with no workspace to resolve it against.
const TOOLS_WITH_RELATIVE = [
    ["read_document", { path: REL, limit: 10, offset: 0 }],
    ["create_document", { path: REL, blocks: [{ kind: "paragraph", text: "Hello." }] }],
    ["edit_document", { path: REL, op: "replace_text", address: "p:0123456789ab", revisionToken: "abc", text: "Hi." }],
    ["revert_document", { path: REL, revisionToken: "abc" }],
];

for (const [name, args] of TOOLS_WITH_RELATIVE) {
    test(`${name} refuses a relative path when no workspace is set, without resolving it`, async () => {
        const before = allDocPaths().length;
        assertTypedRefusal(await toolNamed(name).handler(args), "workspace_unavailable", workspaceUnavailable(REL));
        // Structural, not incidental: the path was never handed to a cache at
        // all, so it cannot have been resolved against cwd (or anything else).
        assert.deepEqual(
            allDocPaths().slice(before),
            [],
            "a declined relative path must never reach the render cache",
        );
    });
}

test("the open_document action refuses a relative path when no workspace is set", async () => {
    await openPanel("relpath-action", {});
    const action = canvas.actions.find((a) => a.name === "open_document");
    assert.ok(action, "open_document must be registered");
    const before = allDocPaths().length;

    await assert.rejects(
        () => action.handler({ instanceId: "relpath-action", input: { path: REL } }),
        (err) => {
            // The canvas channel: measured, a thrown message survives but the
            // `code` field does not, so the code is folded into the message.
            assert.equal(err.name, "CanvasError");
            assert.equal(err.code, "workspace_unavailable");
            assert.equal(err.message, `workspace_unavailable: ${workspaceUnavailable(REL)}`);
            assert.doesNotMatch(err.message, /ERR_INVALID_ARG_TYPE/);
            return true;
        },
    );
    assert.deepEqual(allDocPaths().slice(before), [], "a declined relative path must never reach the render cache");
});

test("a relative path is resolved against the workspace when one is set", async () => {
    sdk.setWorkspacePath(home);
    const before = allDocPaths().length;

    const result = await toolNamed("read_document").handler({ path: REL });

    assert.notEqual(result?.resultType, "failure", `expected success, got ${JSON.stringify(result)?.slice(0, 160)}`);
    const resolved = allDocPaths().slice(before);
    assert.equal(resolved.length, 1, "exactly one document path should have reached the cache");
    // The workspace root joined with the relative input -- and nothing cwd-ish.
    assert.equal(resolved[0], normalizeDocPath(path.join(home, REL)));
});

test("an absolute path is passed through unchanged whether or not a workspace is set", async () => {
    const absolute = path.join(home, "abs-probe-158.docx");
    assert.ok(path.isAbsolute(absolute));

    // No workspace: the guard must bite only relatives, never an absolute path.
    let before = allDocPaths().length;
    let result = await toolNamed("read_document").handler({ path: absolute });
    assert.notEqual(result?.resultType, "failure", "an absolute path must resolve with no workspace set");
    assert.equal(allDocPaths().slice(before)[0], normalizeDocPath(absolute));

    // With a workspace: the absolute path is still itself, not joined onto it.
    sdk.setWorkspacePath(home);
    before = allDocPaths().length;
    result = await toolNamed("read_document").handler({ path: absolute });
    assert.notEqual(result?.resultType, "failure", "an absolute path must resolve with a workspace set");
    assert.equal(allDocPaths().slice(before)[0], normalizeDocPath(absolute));
});
