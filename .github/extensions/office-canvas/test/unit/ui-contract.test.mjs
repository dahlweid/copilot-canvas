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

// --- The chrome's naming and its buttons -------------------------------------
//
// The header bar names the *document*; the product names itself elsewhere. That
// split is the whole point of issue #59, and it is spread across three files
// with nothing but convention holding it together, so pin it here.

/** Pull every `<button>` out of the markup, tag open to tag close. */
const buttonsIn = (markup) =>
    [...markup.matchAll(/<button\b([^>]*)>([\s\S]*?)<\/button>/g)].map((m) => ({
        attrs: m[1],
        inner: m[2],
        id: m[1].match(/\bid="([^"]+)"/)?.[1] ?? null,
        // An accessible name is the visible text once decorative children are
        // gone, or an explicit aria-label. `<svg>` contributes nothing to the
        // name -- which is exactly why an icon-only button needs the label.
        text: m[2].replace(/<svg[\s\S]*?<\/svg>/g, "").replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim(),
        ariaLabel: m[1].match(/\baria-label="([^"]+)"/)?.[1] ?? null,
        title: m[1].match(/\btitle="([^"]+)"/)?.[1] ?? null,
    }));

test("the product name says the same thing in all three places that carry it", async () => {
    // `<title>` is the browser/tab name, the picker `<h1>` is the no-document
    // headline, and `PRODUCT` in app.js rebuilds `document.title` once a
    // document is open. Three copies of one string: left unpinned, a rename
    // lands in one or two of them and the product answers to two names.
    const [script, markup] = await Promise.all([read("app.js"), read("index.html")]);

    const title = markup.match(/<title>([^<]*)<\/title>/)?.[1]?.trim();
    const h1 = markup.match(/<h1\b[^>]*\bid="productName"[^>]*>([^<]*)<\/h1>/)?.[1]?.trim();
    const constant = script.match(/const\s+PRODUCT\s*=\s*"([^"]*)"/)?.[1];

    assert.ok(title, "index.html has no <title>");
    assert.ok(h1, 'index.html has no <h1 id="productName">, so the picker names nothing');
    assert.ok(constant, "app.js no longer defines PRODUCT, so document.title cannot name the product");
    assert.equal(h1, title, "the picker headline and <title> disagree about the product's name");
    assert.equal(constant, title, "app.js PRODUCT and <title> disagree about the product's name");
});

test("the header bar names the document, not the product", async () => {
    // The bug this fixes: `docName` preferred the docx metadata `title`, so a
    // fixture whose properties said "Word Canvas Fixture" made the viewer look
    // like a product of that name. `doc.name` is the filename.
    const script = await read("app.js");
    const assignment = script.match(/el\.docName\.textContent\s*=\s*([^;]+);/)?.[1]?.trim();

    assert.ok(assignment, "app.js never sets docName.textContent");
    assert.equal(assignment, "doc.name", "the header bar must show the filename, not a metadata title");
});

test("the Change... button is gone from the markup and from the wiring", async () => {
    // Removing only the markup leaves `$("changeDoc")` returning null and the
    // listener attach throwing on load -- which takes the whole UI with it. Both
    // sides have to go, so ask both files.
    const [script, markup] = await Promise.all([read("app.js"), read("index.html")]);

    assert.ok(!markup.includes("changeDoc"), "index.html still carries the changeDoc button");
    assert.ok(!script.includes("changeDoc"), "app.js still references changeDoc, which is now null");
});

test("every button has an accessible name", async () => {
    const markup = await read("index.html");
    const buttons = buttonsIn(markup);

    // A guard against the extraction matching nothing: an empty set would pass
    // every assertion below while measuring no button at all.
    assert.ok(buttons.length >= 6, `expected the markup to define buttons, found ${buttons.length}`);

    for (const b of buttons) {
        const named = b.text || b.ariaLabel;
        assert.ok(named, `button ${b.id ?? b.attrs.trim()} has no accessible name`);
        // Icon-only is allowed, but then the label carries the name and the
        // title carries the same thing for a sighted pointer user.
        if (!b.text) {
            assert.ok(b.ariaLabel, `icon-only button ${b.id} needs aria-label`);
            assert.ok(b.title, `icon-only button ${b.id} needs title`);
        }
    }
});

test("button icons are decorative and stay out of the accessibility tree", async () => {
    // The icons repeat the label they sit beside. Left exposed, a screen reader
    // reads the button twice; left focusable, an SVG becomes a tab stop in IE-
    // derived engines. Both are one attribute each.
    const markup = await read("index.html");
    const buttons = buttonsIn(markup);
    const withIcons = buttons.filter((b) => b.inner.includes("<svg"));

    assert.ok(withIcons.length >= 5, `expected the toolbar and change bar to carry icons, found ${withIcons.length}`);

    for (const b of withIcons) {
        for (const svg of b.inner.match(/<svg\b[^>]*>/g) ?? []) {
            assert.match(svg, /aria-hidden="true"/, `the icon in ${b.id} is not aria-hidden`);
            assert.match(svg, /focusable="false"/, `the icon in ${b.id} is focusable`);
        }
    }
});

test("keyboard focus stays visible", async () => {
    // The restyle replaced the buttons' backgrounds and borders. A hover-only
    // affordance would leave a keyboard user with nothing, and `outline: none`
    // is the classic way that happens by accident while chasing a cleaner look.
    const css = await read("app.css");
    const rule = css.match(/:focus-visible\s*\{([^}]*)\}/);

    assert.ok(rule, "app.css defines no :focus-visible rule");
    assert.match(rule[1], /outline\s*:\s*\d/, ":focus-visible must draw an outline with a width");
    assert.ok(!/outline\s*:\s*(none|0)\b/.test(rule[1]), ":focus-visible removes the outline");
});

