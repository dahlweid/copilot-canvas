// The supported-extension set is stated in one place and derived everywhere
// else. These tests are what keep that true.
//
// The drift they exist to catch was real and silent: `SUPPORTED` gained
// `.dotx`, but two tool parameter descriptions and the picker's own set were
// hardcoded copies that did not. The tool descriptions are the contract a
// language model reads, so the agent declined `.dotx` documents the code would
// have opened -- plausibly, by telling the user templates are unsupported, with
// no error and nothing to grep for.
//
// `extension.mjs` cannot be imported here: it calls `joinSession()` at module
// scope. So its tool descriptions are checked by reading the source, which is
// enough to assert the property that matters -- that the list is interpolated
// rather than spelled out.

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { SUPPORTED, supportedList, requireSupported } from "../../src/render-cache.mjs";
import { DOC_EXTENSIONS } from "../../src/server.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));
const extensionSource = await readFile(path.join(here, "..", "..", "extension.mjs"), "utf8");

test("every supported extension appears in the tools' path descriptions", () => {
    // The assertion the reviewer asked for -- but stated over the *rendered*
    // description, not over `supportedList()`. Asserting that supportedList()
    // contains SUPPORTED is a tautology: it is built from it, so the test
    // passes even when a description is hardcoded and wrong. Verified by
    // reintroducing the defect and watching this fail.
    const descriptions = [...extensionSource.matchAll(/path to a Word document \((.*?)\)\./g)].map((m) => m[1]);
    assert.ok(descriptions.length >= 2, `expected both path descriptions, found ${descriptions.length}`);

    for (const description of descriptions) {
        // A derived description is a template expression; render it the way
        // the SDK will, so the assertion is about what the model finally reads.
        const rendered = description.replace("${supportedList()}", supportedList());
        for (const ext of SUPPORTED) {
            assert.ok(rendered.includes(ext), `${ext} is missing from a tool description: "${rendered}"`);
        }
    }
});

test("the tool descriptions are derived, not restated", () => {
    // A hardcoded list is the defect itself, so look for one rather than for
    // its current contents: any literal that spells out two adjacent
    // extensions has stopped deriving.
    const hardcoded = extensionSource.match(/\.docx,\s*\.docm/g) ?? [];
    assert.deepEqual(hardcoded, [], "a tool description spells out the extension list instead of deriving it");

    const derived = extensionSource.match(/path to a Word document \(\$\{supportedList\(\)\}\)/g) ?? [];
    const anyDescription = extensionSource.match(/path to a Word document \(/g) ?? [];
    assert.ok(anyDescription.length >= 2, `expected both path descriptions, found ${anyDescription.length}`);
    assert.equal(
        derived.length,
        anyDescription.length,
        `${anyDescription.length - derived.length} path description(s) are not derived from SUPPORTED`,
    );
});

test("the picker offers exactly what the cache will open", () => {
    // These drifted apart once: a .dotx in the workspace was hidden from the
    // picker while open_document on the same path succeeded.
    assert.deepEqual([...DOC_EXTENSIONS].sort(), [...SUPPORTED].sort());
});

test("the unsupported-type error names the same set", () => {
    try {
        requireSupported("notes.txt");
        assert.fail("expected .txt to be rejected");
    } catch (err) {
        for (const ext of SUPPORTED) {
            assert.ok(err.message.includes(ext), `${ext} is missing from the rejection message`);
        }
    }
});

test("a template is accepted, since the cache supports it", () => {
    // The extension whose absence from the descriptions started this.
    assert.ok(requireSupported("template.dotx").endsWith(".dotx"));
});
