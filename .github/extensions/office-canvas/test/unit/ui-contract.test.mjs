// The UI's element contract.
//
// `app.js` reaches for its elements by id through a one-line `$` helper, so a
// renamed or missing id yields `null` and the failure surfaces later, somewhere
// else, as a TypeError on a property access -- or not at all, when the element is
// only touched on a branch nobody took during a smoke run. The change bar is
// exactly that shape: three ids, two of them behind click handlers that a
// rendering test never fires.
//
// Nothing else can see this. The validator parses modules, not markup; C3 checks
// import specifiers; the integration smokes drive the HTTP API rather than the
// DOM. This asks the two files whether they still agree.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const UI = path.resolve(HERE, "..", "..", "src", "ui");

const read = (name) => readFile(path.join(UI, name), "utf8");

test("every id app.js looks up exists in index.html", async () => {
    const [script, markup] = await Promise.all([read("app.js"), read("index.html")]);

    // Both sides are extracted rather than listed: a list here would be a third
    // copy of the contract, free to drift from the two it describes.
    const wanted = [...script.matchAll(/\$\("([^"]+)"\)/g)].map((m) => m[1]);
    const present = new Set([...markup.matchAll(/\bid="([^"]+)"/g)].map((m) => m[1]));

    assert.ok(wanted.length > 10, `expected app.js to look up ids, found ${wanted.length}`);

    const missing = [...new Set(wanted)].filter((id) => !present.has(id)).sort();
    assert.deepEqual(missing, [], `app.js looks up ids that index.html does not define: ${missing.join(", ")}`);
});

test("the change bar's ids are among them", async () => {    // A guard against the extraction quietly matching nothing. If the regex above
    // stopped finding lookups, `missing` would be empty and the test would pass
    // while measuring an empty set -- so name the ids this layer added and
    // require them to have been seen.
    const script = await read("app.js");
    const wanted = new Set([...script.matchAll(/\$\("([^"]+)"\)/g)].map((m) => m[1]));

    for (const id of ["changeBar", "changeText", "jumpToChange", "dismissChange"]) {
        assert.ok(wanted.has(id), `app.js no longer looks up ${id}`);
    }
});

test("hidden elements are hidden with a declaration that outranks a component's own display", async () => {
    // The UI hides things by setting the `hidden` attribute, and every element it
    // hides is a flex container. The UA stylesheet's `[hidden] { display: none }`
    // loses to any author `display`, so before this rule existed `#sidebar` and
    // `#changeBar` both computed to `display: flex` while carrying `hidden` --
    // measured in the running viewer, along with client rects to prove they were
    // laid out and not merely styled.
    //
    // `!important` is load-bearing, not decoration: `[hidden]` and `.sidebar` have
    // equal specificity, so without it source order decides and the component
    // wins. Measured by mutating the live rule -- dropping `!important` returned
    // `flex`, restoring it returned `none`.
    //
    // What this pins is the rule's presence and its priority. It cannot see a
    // stylesheet loaded after app.css, and there is no second stylesheet today;
    // the DOM-level check belongs to the smoke suite.
    const css = await read("app.css");
    const rule = css.match(/\[hidden\]\s*\{([^}]*)\}/);
    assert.ok(rule, "app.css defines no [hidden] rule, so `hidden` is advisory only");
    assert.match(rule[1], /display\s*:\s*none\s*!important/, "[hidden] must beat a component's own display");
});
