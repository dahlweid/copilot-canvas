// What the bar shows while the Word mark is loading, and if it never does (#68).
//
// The elements are stubs rather than a DOM: what is being measured is the order
// of the swap and which element is left visible, and that is a property of the
// module, not of a browser. `app.js`, which calls it, is executed by
// `app.test.mjs` since #76 -- but not this part of it: `showWordMark` is fired
// at load and deliberately neither awaited nor reported, so the panel cannot
// observe it and neither can a test driving the panel.

import { test } from "node:test";
import assert from "node:assert/strict";

import { showWordMark } from "../../src/ui/word-mark.mjs";

/** Just enough of an element: records listeners so a load or error can be fired. */
function stubElement() {
    const listeners = new Map();
    return {
        hidden: false,
        src: null,
        addEventListener(type, fn) {
            listeners.set(type, fn);
        },
        fire(type) {
            const fn = listeners.get(type);
            assert.ok(fn, `nothing is listening for ${type}`);
            fn();
        },
        listens: (type) => listeners.has(type),
    };
}

test("the mark stays hidden until it has actually loaded", async () => {
    // An <img> whose source 404s draws the browser's broken-image affordance.
    // Setting `hidden` only after a failure would flash that box on every
    // machine without Word, which is most of them.
    const img = stubElement();
    const fallback = stubElement();

    showWordMark({ img, fallback });

    assert.equal(img.hidden, true, "the mark is visible before it has loaded");
    assert.equal(fallback.hidden, false, "the drawn glyph was hidden before its replacement arrived");
    assert.equal(img.src, "/api/word-icon");
});

test("a loaded mark replaces the drawn glyph", async () => {
    const img = stubElement();
    const fallback = stubElement();

    showWordMark({ img, fallback });
    img.fire("load");

    assert.equal(img.hidden, false);
    assert.equal(fallback.hidden, true, "both the mark and the glyph are showing");
});

test("a machine with no Word keeps the glyph the markup shipped", async () => {
    // The degradation #68 asks for: nothing breaks and nothing is said. The
    // route answers 404, the image errors, and the bar looks exactly as it did
    // before this feature existed.
    const img = stubElement();
    const fallback = stubElement();

    showWordMark({ img, fallback });
    img.fire("error");

    assert.equal(img.hidden, true);
    assert.equal(fallback.hidden, false, "a failed mark left the button with no icon at all");
});

test("a placement with no glyph behind it simply shows nothing", async () => {
    // The document name had no icon before #68, so there is nothing to fall back
    // to and nothing to restore. A missing decoration is the whole failure.
    const img = stubElement();

    const visible = showWordMark({ img });
    img.fire("error");

    assert.equal(img.hidden, true);
    assert.equal(visible, img, "a placement with no fallback reported one anyway");
});

test("both outcomes are wired before the request is made", async () => {
    // Assigning `src` first would be a race on a cached response: the browser
    // can fire `load` synchronously, before a listener attached afterwards
    // exists, leaving the mark permanently hidden on exactly the second visit.
    const order = [];
    const img = {
        hidden: false,
        addEventListener: (type) => order.push(type),
        set src(_value) {
            order.push("src");
        },
    };

    showWordMark({ img });

    assert.deepEqual(order, ["load", "error", "src"], "the source was set before both handlers existed");
});
