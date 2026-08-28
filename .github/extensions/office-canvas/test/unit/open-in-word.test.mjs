// "Edit in Word" must hand the document path to the launcher with no shell
// parser in between.
//
// Office-free, and deliberately platform-independent: the point is the *shape*
// of the spawn, which is asserted through an injected spawnFn rather than by
// launching anything. The measurement that motivates it is Windows-only
// (spikes/isolation/probes/probe-open-in-word-quoting.mjs), but the regression
// it guards against is a source change, so it belongs on the hosted runner.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { ViewerInstance } from "../../src/server.mjs";

// Filenames a `cmd.exe` parse mangles, plus benign controls. The first three
// are the hostile cases: `&` truncates the path and runs the tail as a command,
// `^` is eaten as cmd's escape character, and a matched `%…%` pair is expanded.
// None has a space, which is what keeps Node from quoting them in the first
// place.
//
// The last two are controls, and they earn their place: they survive a cmd
// parse, because Node quotes them. They are here so the assertion covers the
// paths the old implementation handled correctly as well as the ones it broke —
// a fix that only repaired the hostile cases would be a regression for these.
const LAUNCH_PATHS = [
    "C:\\Docs\\R&D.docx",
    "C:\\Docs\\caret^v.docx",
    "C:\\Docs\\%PATH%.docx",
    "C:\\Docs\\with space.docx",
    "C:\\Docs\\it's mine.docx",
];

const instanceWith = (docPath, calls) => {
    const instance = new ViewerInstance({
        cache: null,
        instanceId: "test",
        workspacePath: null,
        spawnFn: (command, args, options) => {
            calls.push({ command, args, options });
            return { unref() {} };
        },
    });
    instance.doc = { path: docPath };
    return instance;
};

test("the path is passed as one argv element, never interpolated into a command", () => {
    for (const docPath of LAUNCH_PATHS) {
        const calls = [];
        instanceWith(docPath, calls).openInWord();

        assert.equal(calls.length, 1, `one spawn for ${docPath}`);
        const { args } = calls[0];
        assert.deepEqual(args, [docPath], `${docPath} must arrive whole and alone`);
    }
});

test("no cmd.exe, and no shell option, on the launch path", () => {
    const calls = [];
    instanceWith("C:\\Docs\\R&D.docx", calls).openInWord();

    const { command, options } = calls[0];
    // `cmd.exe /c start` was the original implementation and is the defect.
    assert.notEqual(command.toLowerCase(), "cmd.exe");
    assert.equal(options.shell ?? false, false, "a shell reintroduces the parser this avoids");
});

test("the launcher is detached, so closing the canvas does not close Word", () => {
    const calls = [];
    instanceWith("C:\\Docs\\plain.docx", calls).openInWord();

    assert.equal(calls[0].options.detached, true);
    assert.equal(calls[0].options.stdio, "ignore");
});

test("launching with no document open reports it rather than spawning", () => {
    const calls = [];
    const instance = new ViewerInstance({
        cache: null,
        instanceId: "test",
        workspacePath: null,
        spawnFn: (...a) => {
            calls.push(a);
            return { unref() {} };
        },
    });

    assert.throws(() => instance.openInWord(), /No document is open/);
    assert.equal(calls.length, 0);
});
