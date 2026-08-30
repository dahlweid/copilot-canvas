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
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const UI = path.resolve(HERE, "..", "..", "src", "ui");
const SRC = path.resolve(HERE, "..", "..", "src");
const EXTENSION = path.resolve(HERE, "..", "..", "extension.mjs");

const read = (name) => readFile(path.join(UI, name), "utf8");
const readSrc = (name) => readFile(path.join(SRC, name), "utf8");
const readExtension = () => readFile(EXTENSION, "utf8");

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
    // every assertion below while measuring no button at all. Five, since #69
    // retired the picker's Open button along with the rest of that screen.
    assert.ok(buttons.length >= 5, `expected the markup to define buttons, found ${buttons.length}`);

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

// --- Who owns the connection status -------------------------------------------
//
// #66: `onerror` claimed "Reconnecting…" for as long as `readyState` was
// CONNECTING, and `EventSource` keeps it there forever, so the claim was
// permanent and false. The decision now lives in `connection-status.mjs`, where
// `connection-status.test.mjs` can execute it.
//
// **What these two can and cannot fail on**, stated because a guard that reads
// like protection and is not is this repo's most expensive habit. They fail on a
// delegation that was removed and on an inline status that came back -- the
// regression actually at issue. They *cannot* fail on a monitor that is
// imported, called, and then never reached at runtime, because nothing **here**
// executes `app.js`. That used to be forced -- `app.js` could not be imported
// under Node, failing at module resolution of `pdf-view.mjs`'s absolute
// `/vendor/pdf.min.mjs` before any DOM access -- and since #76 it is a property
// of this file only: `app.test.mjs` does execute `app.js`, against a stub DOM.
//
// That blind spot is no longer unattended, and the handover is measured rather
// than asserted. `test/unit/connection-lost.test.mjs` drives `app.js` until its
// `#status` element carries the terminal message. Replacing the call below with
// `monitorConnection(source, { setStatus: () => {} })` -- still a call, still
// matching the regex -- leaves **this file green** and turns that one red.
//
// Neither reaches a real `EventSource`, so the reconnect scheduling is still
// measured by `spikes/viewer-connection/probes/probe-dead-server.mjs` instead,
// against a real EventSource and a really-closed server.

test("app.js delegates its connection status rather than deciding it inline", async () => {
    const script = await read("app.js");

    assert.match(
        script,
        /import\s*\{[^}]*\bmonitorConnection\b[^}]*\}\s*from\s*"\.\/connection-status\.mjs"/,
        "app.js no longer imports the connection monitor",
    );
    assert.match(script, /monitorConnection\s*\(\s*source\s*,/, "app.js imports the monitor but never calls it");
});

test("app.js carries no connection status of its own", async () => {
    // Comments are stripped first, and that is not a convenience: this file's
    // own explanation of the fix quotes the very strings being banned, and the
    // comment left in `app.js` at the delegation site does too. Matching raw
    // source would make the guard fail on prose describing the fix while
    // passing on prose that happened to avoid the words -- i.e. it would be
    // measuring vocabulary, not behaviour. The `:` exclusion keeps `http://`
    // and `file://` from eating the rest of their line.
    const script = (await read("app.js")).replace(/\/\*[\s\S]*?\*\//g, "").replace(/(^|[^:])\/\/.*$/gm, "$1");

    // These strings live in the module now, so their presence here means a
    // second opinion about the connection was reintroduced beside the first.
    assert.doesNotMatch(script, /Reconnecting/, "app.js is announcing a reconnection again");
    assert.doesNotMatch(script, /readyState/, "app.js is branching on readyState again");
    // Both handlers belong to the monitor. An assignment here silently replaces
    // whichever one it names -- the monitor would still be wired and simply
    // never called.
    assert.doesNotMatch(script, /source\s*\.\s*on(error|open)\s*=/, "app.js overwrites a handler the monitor owns");
});

// --- The no-document state (#69) ---------------------------------------------
//
// The picker was the last screen that asked the user to choose a document, which
// is the one job the canvas gives to Copilot. Deleting it is only half the
// change: the state it occupied is reachable (`path` is optional in the canvas
// input schema), so what stands there now has to tell a first-time reader what
// to do instead.

test("the document picker is gone from the markup, the wiring and the API", async () => {
    // Same shape as the Change... removal above: markup alone is not enough,
    // because a leftover `$("pathForm")` returns null and the listener attach
    // throws on load, taking the whole UI with it.
    const [script, markup, server] = await Promise.all([read("app.js"), read("index.html"), readSrc("server.mjs")]);

    for (const id of ["pathForm", "pathInput", "pickerError", "recents", "recentsBlock", "workspaceDocs"]) {
        assert.ok(!markup.includes(id), `index.html still carries the picker's ${id}`);
        assert.ok(!script.includes(id), `app.js still references ${id}, which is now null`);
    }
    assert.ok(!script.includes("/api/browse"), "app.js still fetches the picker's browse endpoint");
    assert.ok(!server.includes("/api/browse"), "server.mjs still routes the picker's browse endpoint");
});

test("the no-document state invites no input", async () => {
    // The property, not the current markup: any form control on this screen is
    // the picker returning under another name. `<input>`, `<textarea>` and
    // `<select>` are the three ways a user types or picks, and a `<form>` is the
    // thing that would submit one.
    const markup = await read("index.html");
    const empty = markup.match(/<section class="empty"[\s\S]*?<\/section>/)?.[0];

    assert.ok(empty, "index.html has no empty state, so a canvas opened with no path shows nothing");
    for (const tag of ["<form", "<input", "<textarea", "<select"]) {
        assert.ok(!empty.includes(tag), `the no-document state offers a ${tag}> to operate`);
    }
});

test("the no-document state says who opens a document and how", async () => {
    // What replaces the picker has one job: a reader who opens a fresh canvas
    // must learn that they ask Copilot, and what Copilot will use. Both halves
    // are asserted, because either alone leaves the screen a dead end -- "ask
    // Copilot" with no verb, or an action name with nobody to address.
    const markup = await read("index.html");
    const empty = markup.match(/<section class="empty"[\s\S]*?<\/section>/)?.[0] ?? "";

    assert.match(empty, /\bCopilot\b/, "the no-document state never names Copilot");
    assert.match(empty, /open_document/, "the no-document state never names the action that opens a document");
    assert.match(
        empty,
        /No document is open/i,
        "the no-document state never says that no document is open",
    );
});

test("the canvas description no longer promises a document picker", async () => {
    // The schema still makes `path` optional, and should: the panel must be
    // able to exist empty. What had to change is the sentence that told the
    // agent -- and through it the user -- to expect a picker there.
    const extension = await readExtension();
    assert.ok(!/document picker/i.test(extension), "extension.mjs still promises a document picker");
    assert.match(
        extension,
        /Omit to open the canvas empty/,
        "the optional `path` no longer says what omitting it does",
    );
});

// --- Where the absolute path went (#71) --------------------------------------
//
// The path had a full-width row of its own whose tooltip repeated its text
// verbatim. Removing the row is the easy half; the hard half is that the path
// must not become pointer-only, which is what a bare `title` would make it.

test("the absolute-path row is gone from the markup, the wiring and the layout", async () => {
    // Three files, because a leftover in any one of them is a different bug: a
    // stale `$("docPath")` returns null and the assignment in `render()` throws,
    // taking the UI down; a stale grid area leaves a blank strip under the bar.
    const [script, markup, css] = await Promise.all([read("app.js"), read("index.html"), read("app.css")]);

    assert.ok(!markup.includes("docPath"), "index.html still carries the absolute-path row");
    assert.ok(!script.includes("docPath"), "app.js still references docPath, which is now null");
    assert.ok(!css.includes(".bar-path"), "app.css still styles the row that no longer exists");

    const bar = css.match(/\.bar\s*\{([^}]*)\}/);
    assert.ok(bar, "app.css defines no .bar rule");
    assert.ok(!/\bpath\b/.test(bar[1]), ".bar still reserves a grid area for the path row");
});

test("the path is reachable without a pointer", async () => {
    // The accessibility half of #71, and the reason a tooltip alone was not the
    // answer: `title` is not keyboard-reachable and is announced inconsistently.
    // What replaces the row has to be a control -- a tab stop with a name --
    // and `app.js` has to put the real path into that name rather than leaving
    // the static placeholder standing.
    const [script, markup] = await Promise.all([read("app.js"), read("index.html")]);
    const copy = buttonsIn(markup).find((b) => b.id === "copyPath");

    assert.ok(copy, "the markup defines no control carrying the path");
    assert.ok(copy.ariaLabel, "the path control has no accessible name");
    assert.ok(copy.title, "the path control has no tooltip for a pointer user");
    assert.match(
        script,
        /el\.copyPath\.setAttribute\("aria-label"/,
        "app.js never puts the path into the control's name",
    );
    assert.match(script, /el\.docName\.title\s*=/, "app.js never puts the path into the name's tooltip");
});

test("app.js delegates what the bar says about a document rather than deciding it inline", async () => {
    // Same reasoning as the connection status above: `app.js` cannot be
    // imported under Node, so a decision left in it has no reachable test. This
    // fails on a delegation that was removed and on the old inline assignment
    // coming back. It *cannot* fail on a wrong answer from the module -- that is
    // `doc-identity.test.mjs`'s job, and it executes the code.
    const script = await read("app.js");

    assert.match(
        script,
        /import\s*\{[^}]*\bdescribeDocument\b[^}]*\}\s*from\s*"\.\/doc-identity\.mjs"/,
        "app.js no longer imports the document description",
    );
    assert.match(script, /describeDocument\s*\(\s*doc\s*\)/, "app.js imports the description but never calls it");
});

// --- The Word mark (#68) -----------------------------------------------------
//
// The hard constraint is what is *absent*: no icon file may be committed. The
// repo is public, so shipping Microsoft's mark would be redistributing it, and
// the open icon sets have dropped the family for trademark reasons. The bytes
// come from the user's own Word at runtime or they do not come at all.

test("no icon is committed anywhere in the extension", async () => {
    // Stated as a property of the tree rather than of the files this change
    // happened to add: the constraint is "never a resource in the sources", so
    // it has to fail on a .png or .ico appearing anywhere, by any later change.
    const root = path.join(EXTENSION, "..");
    const found = [];
    const walk = async (dir) => {
        for (const entry of await readdir(dir, { withFileTypes: true })) {
            if (entry.name === "node_modules" || entry.name === ".git") continue;
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) await walk(full);
            else if (/\.(png|ico)$/i.test(entry.name)) found.push(path.relative(root, full));
        }
    };
    await walk(root);

    assert.deepEqual(found, [], `an icon was committed: ${found.join(", ")}`);
});

test("the Word mark is fetched at runtime, from a route the server actually has", async () => {
    const [markup, css, server] = await Promise.all([read("index.html"), read("app.css"), readSrc("server.mjs")]);

    assert.match(markup, /class="word-mark"/, "the markup has nowhere to put the mark");
    assert.match(css, /\.word-mark\s*\{/, "the mark is unsized, so it would draw at whatever the extraction produced");
    assert.match(server, /"GET \/api\/word-icon"/, "the markup fetches a route the server does not serve");
});

test("the mark ships with no source, so a machine without Word draws nothing broken", async () => {
    // The single most likely regression here: giving the <img> a `src` in the
    // markup. It looks tidier and it puts a broken-image box in the bar of
    // every machine that has no Word, which is most of them. The source is set
    // by `word-mark.mjs`, after which a successful load reveals the image.
    const markup = await read("index.html");
    const marks = [...markup.matchAll(/<img\b[^>]*class="word-mark"[^>]*>/g)].map((m) => m[0]);

    assert.ok(marks.length >= 2, `expected the name and the button to carry a mark, found ${marks.length}`);
    for (const mark of marks) {
        assert.ok(!/\bsrc=/.test(mark), `a word-mark ships with a src: ${mark}`);
        assert.match(mark, /\bhidden\b/, `a word-mark is visible before it has loaded: ${mark}`);
        // Decorative, exactly like the glyphs it replaces: the button's text is
        // its name and the document's name is beside it.
        assert.match(mark, /alt=""/, `a word-mark claims an accessible name: ${mark}`);
    }
});

test("app.js delegates the mark's fallback rather than deciding it inline", async () => {
    // The failure path is the whole point of #68's "degrade gracefully", and it
    // has branches, so it lives where a test can execute it.
    const script = await read("app.js");

    assert.match(
        script,
        /import\s*\{[^}]*\bshowWordMark\b[^}]*\}\s*from\s*"\.\/word-mark\.mjs"/,
        "app.js no longer imports the mark loader",
    );
    assert.match(script, /showWordMark\s*\(/, "app.js imports the loader but never calls it");
    // The button keeps a drawn glyph to fall back to; the name never had one.
    assert.match(script, /fallback:\s*el\.openInWordGlyph/, "the button has no glyph to fall back to");
});

