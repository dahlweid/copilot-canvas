// What the bar shows while the Word mark is loading, and if it never does (#68).
//
// The element is a stub rather than a DOM: what is being measured is the order
// of the wiring and the state the image is left in, and that is a property of
// the module, not of a browser. `app.js`, which calls it, is executed by
// `app.test.mjs` since #76 -- but not this part of it: `showWordMark` is fired
// at load and deliberately neither awaited nor reported, so the panel cannot
// observe it and neither can a test driving the panel.
//
// ## Three tests were deleted here, and their subject with them (#87)
//
// "a loaded mark replaces the drawn glyph", "a machine with no Word keeps the
// glyph the markup shipped" and "a placement with no glyph behind it simply
// shows nothing" all covered the `fallback` parameter, which is gone: the Open
// in Word button was its only caller and now draws its own glyph
// unconditionally. They are deleted rather than rewritten because there is no
// longer anything for them to be about (precedent: #69).
//
// The first two are worth a sentence more, because they were passing on a claim
// the browser does not honour. The fallback was an `<svg>`, and `hidden` is an
// IDL attribute of `HTMLElement` that `SVGElement` does not inherit -- so
// `fallback.hidden = true` hid nothing in Chromium, while passing here against a
// stub on which `hidden` is a plain data property. Measured by
// spikes/viewer-header/probes/probe-mark-alignment.mjs. A test that cannot fail
// for the reason its comment gives is the failure this repo files; deleting
// these is the fix, not the cover-up.

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

    showWordMark({ img });

    assert.equal(img.hidden, true, "the mark is visible before it has loaded");
    assert.equal(img.src, "/api/word-icon");
});

test("a loaded mark is revealed", async () => {
    const img = stubElement();

    showWordMark({ img });
    img.fire("load");

    assert.equal(img.hidden, false, "the mark loaded and stayed hidden");
});

test("a machine with no Word shows nothing at all", async () => {
    // The degradation #68 asks for: nothing breaks and nothing is said. The
    // route answers 404, the image errors, and the bar looks exactly as it did
    // before this feature existed.
    const img = stubElement();

    showWordMark({ img });
    img.fire("error");

    assert.equal(img.hidden, true, "a failed mark left a broken image in the bar");
});

test("the mark takes no second element to hide", async () => {
    // The guard on #87's removal. `showWordMark` touched a `fallback` element,
    // and the one it was given was an <svg>, which `hidden` does not hide. It is
    // now given nothing to touch: whatever it is passed, the only element it
    // may write to is the image.
    const img = stubElement();
    const other = stubElement();

    showWordMark({ img, fallback: other, src: "/api/word-icon" });
    img.fire("load");

    assert.equal(other.hidden, false, "showWordMark still writes `hidden` to a second element");
    assert.equal(other.listens("load"), false, "showWordMark still wires a second element");
    assert.equal(other.src, null, "showWordMark still sets a source on a second element");
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
