// `src/ui/app.js`, executed.
//
// Until #76 this file could not exist. `app.js` was unimportable under Node for
// two reasons, and the second only appears once the first is gone: `pdf-view.mjs`
// imports pdf.js from an absolute `/vendor/` URL, and `app.js` builds its element
// map at module scope, so the DOM is touched before any function runs.
// `ui-harness.mjs` answers both -- a resolve hook for the specifier, a small
// stand-in for the DOM -- without changing a line of the viewer.
//
// What is asserted here is behaviour the source-text checks in
// `ui-contract.test.mjs` cannot reach: those fail on a changed spelling, these
// fail on a wrong answer. Every one was made to go red by mutating the line it
// covers; the mutants are `test/unit/mutate-webview.ps1`.
//
// **What this still does not cover.** It runs `app.js`'s logic under Node, not
// under Chromium. `fetch`, `EventSource`, `IntersectionObserver` and the DOM are
// all stand-ins, so anywhere the browser's implementation differs from the model
// here, this suite is silent -- `EventSource` above all, where undici and
// Chromium are different implementations of one spec. The runtime end of that is
// measured by `spikes/viewer-connection/probes/probe-dead-server.mjs`.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";

import { loadApp, respondWith } from "./ui-harness.mjs";
import { textItem } from "./stub-pdfjs.mjs";

const PRODUCT = "Word Document Viewer";

const DOC = {
    key: "k1",
    name: "report.docx",
    path: "C:\\docs\\report.docx",
    pageCount: 3,
    // Deliberately under a thousand: `render()` formats this with
    // `toLocaleString()`, whose separator is the runner's locale and not a
    // property of `app.js`. Measured -- 1234 reads "1,234" on an en-US runner
    // and "1.234" on this machine's de-DE one, which would make the assertion a
    // statement about where it ran.
    wordCount: 900,
    title: "Quarterly Report",
    author: "A. Author",
};

const READY = {
    status: "ready",
    error: null,
    doc: DOC,
    lastPage: 1,
    wordVersion: "16.0.18025",
    pdfUrl: "/pdf/k1.pdf",
    change: null,
};

/** A change record shaped as `change-record.mjs` mints them. */
const change = (overrides = {}) => ({
    op: "replace_text",
    page: 2,
    text: "The quick brown fox.",
    locatable: true,
    at: 1_700_000_000_000,
    ...overrides,
});

/** The entry buttons a panel is showing, in order. */
const entries = (panel) => panel.children.filter((child) => child.classList?.contains("entry"));

/** The panel's heading line, which every populated panel opens with. */
const panelTitle = (panel) => panel.children[0]?.textContent;

test("the bar names the document and the tab names both", async (t) => {
    const app = await loadApp({ state: READY });
    t.after(app.restore);

    assert.equal(app.el("docName").textContent, "report.docx", "the bar is not naming the document");
    assert.equal(
        app.el("docMeta").textContent,
        "3 pages · 900 words · Quarterly Report · A. Author",
        "the meta line lost or reordered a property",
    );
    // The header names the *document*, so the tab is the only place the product
    // names itself while one is open (#59).
    assert.equal(app.document.title, `report.docx — ${PRODUCT}`);
    assert.equal(app.el("bar").hidden, false);
    assert.equal(app.el("viewer").classList.contains("visible"), true);
    assert.equal(app.el("emptyState").classList.contains("visible"), false);
    assert.equal(
        app.el("engineNote").textContent,
        "Rendered by Microsoft Word 16.0.18025 for page-accurate layout.",
        "the engine note did not take the version the host reported",
    );
});

test("with no document the panel shows the empty state and names itself", async (t) => {
    const app = await loadApp({ state: { status: "idle", doc: null, pdfUrl: null, change: null } });
    t.after(app.restore);

    assert.equal(app.el("emptyState").classList.contains("visible"), true);
    assert.equal(app.el("bar").hidden, true, "the header bar is showing with no document to name");
    assert.equal(app.el("viewer").classList.contains("visible"), false);
    assert.equal(app.el("sidebar").hidden, true);
    assert.equal(app.document.title, PRODUCT);
    assert.equal(app.pdfjs.opened.length, 0, "a document was fetched with none open");
});

test("a dismissed change stays dismissed when the same change is pushed again", async (t) => {
    // The defect this guards is written into `app.js`'s own comment: the state
    // arrives over SSE, so **every frame carries a fresh object**. Holding the
    // dismissed record by identity would un-dismiss the same change on the very
    // next push, which for a panel receiving periodic state is immediately.
    const app = await loadApp({ state: { ...READY, change: change() } });
    t.after(app.restore);

    assert.equal(app.el("changeBar").hidden, false);
    assert.equal(app.el("changeText").textContent, "Text replaced — page 2");

    app.el("dismissChange").dispatch("click");
    assert.equal(app.el("changeBar").hidden, true, "dismissing left the banner up");

    // A fresh object carrying the same change -- what the next state push is.
    app.source.emit("state", { ...READY, change: change() });
    await app.settle();
    assert.equal(app.el("changeBar").hidden, true, "the dismissed change came back on the next state push");

    // A genuinely different change must still announce itself.
    app.source.emit("state", { ...READY, change: change({ at: 1_700_000_009_999, text: "Something else." }) });
    await app.settle();
    assert.equal(app.el("changeBar").hidden, false, "a new change stayed hidden behind an old dismissal");
    assert.equal(app.el("changeText").textContent, "Text replaced — page 2");
});

test("the banner and the marker on the page are shown and dismissed together", async (t) => {
    // `showChange` drives both from one place, and its comment says why: a page
    // carrying a highlight with nothing explaining it is a state the reader
    // cannot act on. Nothing asserted that until this test -- measured, the
    // whole suite stayed green with `view.setChange(record)` deleted.
    const app = await loadApp({
        state: { ...READY, change: change() },
        pdfDocuments: {
            "/pdf/k1.pdf": {
                pageItems: [[], [textItem("The quick brown fox.")], []],
            },
        },
    });
    t.after(app.restore);

    // Page 2 has to have painted before it has any text to be marked on.
    app.observers.at(-1).enter(app.el("pages").children[1]);
    await app.settle();

    const overlay = (page) => page.children.find((child) => child.classList.contains("overlay"));
    const marked = overlay(app.el("pages").children[1]);

    assert.equal(app.el("changeBar").hidden, false);
    assert.equal(marked.hidden, false, "the banner named a change the page did not mark");
    assert.equal(marked.children.at(-1).textContent, "Text replaced");

    app.el("dismissChange").dispatch("click");
    await app.settle();

    assert.equal(app.el("changeBar").hidden, true);
    assert.equal(marked.hidden, true, "dismissing the banner left the highlight with nothing explaining it");
    assert.equal(marked.children.length, 0);
});

test("a change with no page offers nothing to jump to", async (t) => {
    // `describeChangeBanner` decides the wording; what is asserted here is that
    // `app.js` acts on `jumpable` -- an earlier version offered "Show me" beside
    // "page not reported" and scrolled nowhere.
    const app = await loadApp({ state: { ...READY, change: change({ page: null, locatable: false, text: null }) } });
    t.after(app.restore);

    assert.equal(app.el("changeBar").hidden, false);
    assert.equal(app.el("changeText").textContent, "Text replaced — page not reported");
    assert.equal(app.el("jumpToChange").hidden, true, "a change with no page still offered to show it");
});

test("'Show me' scrolls to the change's page and tells the host where it went", async (t) => {
    const app = await loadApp({
        state: { ...READY, change: change({ page: 2 }) },
        pdfDocuments: { "/pdf/k1.pdf": { pages: 3 } },
    });
    t.after(app.restore);

    assert.equal(app.el("jumpToChange").hidden, false);
    app.el("jumpToChange").dispatch("click");
    await app.settle();

    const pageRoots = app.el("pages").children;
    assert.equal(pageRoots.length, 3, "the view did not lay out the document's pages");
    assert.equal(pageRoots[1].scrolledIntoView, 1, "'Show me' did not scroll to the changed page");
    assert.deepEqual(
        app.callsTo("/api/page").map((call) => call.body),
        [{ page: 2 }],
        "the host was not told which page the reader is on",
    );
});

test("an unchanged state push does not re-parse the document", async (t) => {
    // The guard is `loadedPdfUrl === state.pdfUrl`. Without it every unrelated
    // push -- and they arrive for reasons that have nothing to do with the
    // content -- tears the document down and re-parses it.
    const app = await loadApp({
        state: READY,
        pdfDocuments: { "/pdf/k1.pdf": { pages: 3 }, "/pdf/k2.pdf": { pages: 4 } },
    });
    t.after(app.restore);

    assert.deepEqual(app.pdfjs.opened.map((call) => call.url), ["/pdf/k1.pdf"]);

    app.source.emit("state", { ...READY });
    await app.settle();
    assert.deepEqual(
        app.pdfjs.opened.map((call) => call.url),
        ["/pdf/k1.pdf"],
        "an unchanged state push re-fetched and re-parsed the same PDF",
    );

    // The URL carries the render key, so it moves exactly when the document does.
    app.source.emit("state", { ...READY, pdfUrl: "/pdf/k2.pdf", doc: { ...DOC, key: "k2" } });
    await app.settle();
    assert.deepEqual(
        app.pdfjs.opened.map((call) => call.url),
        ["/pdf/k1.pdf", "/pdf/k2.pdf"],
        "a new render was not loaded",
    );
});

test("copy puts on the clipboard exactly what the button announces", async (t) => {
    const app = await loadApp({ state: READY });
    t.after(app.restore);

    const copy = app.el("copyPath");
    assert.equal(copy.title, "C:\\docs\\report.docx");
    assert.equal(copy.getAttribute("aria-label"), "Copy full path: C:\\docs\\report.docx");
    assert.equal(copy.disabled, false);

    copy.dispatch("click");
    await app.settle();

    assert.deepEqual(app.clipboardWrites, [copy.title], "what was copied is not what the control said it would copy");
    assert.equal(app.el("status").textContent, "Full path copied.");
    assert.equal(app.el("status").classList.contains("error"), false);
    assert.equal(app.el("status").hidden, false);
});

test("a clipboard that is not there is reported, not claimed as a copy", async (t) => {
    // `navigator.clipboard` is simply absent outside a secure context, so the
    // call throws a TypeError rather than rejecting -- and the panel is not
    // always in one.
    const app = await loadApp({ state: READY, clipboard: { absent: true } });
    t.after(app.restore);

    app.el("copyPath").dispatch("click");
    await app.settle();

    assert.deepEqual(app.clipboardWrites, []);
    assert.equal(app.el("status").textContent, "Could not copy the path. It is in this button's tooltip.");
    assert.equal(app.el("status").classList.contains("error"), true, "a failed copy was not marked as a failure");
});

test("Enter searches at once, and a result jumps to its page", async (t) => {
    const app = await loadApp({
        state: READY,
        pdfDocuments: { "/pdf/k1.pdf": { pages: 3 } },
        routes: {
            "/api/search": ({ searchParams }) => ({
                hits: [
                    { page: 2, snippet: `…${searchParams.get("q")} here…` },
                    { page: 3, snippet: "…and here…" },
                ],
                truncated: false,
            }),
        },
    });
    t.after(app.restore);

    const input = app.el("searchInput");
    input.value = "widget";
    input.dispatch("input");
    await app.settle();
    assert.equal(app.callsTo("/api/search").length, 0, "typing searched immediately instead of waiting");

    input.dispatch("keydown", { key: "Enter" });
    await app.settle();

    assert.deepEqual(
        app.callsTo("/api/search").map((call) => call.query),
        [{ q: "widget", limit: "100" }],
        "Enter did not run the search it was holding",
    );

    const results = app.el("searchResults");
    assert.equal(results.hidden, false);
    assert.equal(panelTitle(results), "2 matches");
    assert.deepEqual(
        entries(results).map((entry) => entry.textContent),
        ["…widget here…2", "…and here…3"],
    );

    entries(results)[0].dispatch("click");
    await app.settle();
    assert.deepEqual(
        app.callsTo("/api/page").map((call) => call.body),
        [{ page: 2 }],
        "a search result did not move the reader to its page",
    );
    assert.equal(app.el("pages").children[1].scrolledIntoView, 1);
});

test("the outline is fetched once per document, not once per state push", async (t) => {
    const app = await loadApp({
        state: READY,
        routes: {
            "/api/outline": {
                headings: [
                    { text: "Introduction", page: 1, level: 1 },
                    { text: "Method", page: 2, level: 2 },
                ],
            },
        },
    });
    t.after(app.restore);

    app.el("toggleSidebar").dispatch("click");
    await app.settle();

    assert.equal(app.el("sidebar").dataset.open, "true");
    assert.equal(app.el("sidebar").hidden, false);
    assert.equal(app.el("toggleSidebar").getAttribute("aria-pressed"), "true");
    assert.equal(app.el("searchInput").focused, 1, "opening the sidebar did not put the caret in the search box");

    assert.equal(app.callsTo("/api/outline").length, 1);
    assert.equal(panelTitle(app.el("outline")), "Outline");
    assert.deepEqual(
        entries(app.el("outline")).map((entry) => [entry.textContent, entry.dataset.level]),
        [
            ["Introduction1", "1"],
            ["Method2", "2"],
        ],
    );

    // The same document again. `outlineLoadedFor` is what stops a re-fetch.
    app.source.emit("state", { ...READY });
    await app.settle();
    assert.equal(app.callsTo("/api/outline").length, 1, "an unchanged state push re-fetched the outline");

    // A different document must reload it.
    app.source.emit("state", { ...READY, doc: { ...DOC, key: "k2" }, pdfUrl: "/pdf/k2.pdf" });
    await app.settle();
    assert.equal(app.callsTo("/api/outline").length, 2, "a new document kept the previous document's outline");
});

test("closing and re-opening the sidebar shows the outline it already has", async (t) => {
    // The toggle calls `loadOutline` on every open, so `outlineLoadedFor` is the
    // only thing between a reader who likes the sidebar and a request per click.
    const app = await loadApp({
        state: READY,
        routes: { "/api/outline": { headings: [{ text: "Introduction", page: 1, level: 1 }] } },
    });
    t.after(app.restore);

    app.el("toggleSidebar").dispatch("click");
    await app.settle();
    app.el("toggleSidebar").dispatch("click");
    await app.settle();
    assert.equal(app.el("sidebar").hidden, true, "the sidebar did not close");

    app.el("toggleSidebar").dispatch("click");
    await app.settle();

    assert.equal(app.callsTo("/api/outline").length, 1, "re-opening the sidebar re-fetched an outline it had");
    assert.equal(panelTitle(app.el("outline")), "Outline");
    assert.equal(entries(app.el("outline")).length, 1, "the outline was emptied by the round trip");
});

test("a document with no headings says so rather than showing an empty panel", async (t) => {
    const app = await loadApp({ state: READY, routes: { "/api/outline": { headings: [] } } });
    t.after(app.restore);

    app.el("toggleSidebar").dispatch("click");
    await app.settle();

    assert.equal(panelTitle(app.el("outline")), "No headings in this document");
    assert.equal(entries(app.el("outline")).length, 0);
});

test("Reload forces a re-render, and a refusal is shown with the host's own wording", async (t) => {
    const app = await loadApp({ state: READY });
    t.after(app.restore);

    app.el("reload").dispatch("click");
    await app.settle();

    assert.deepEqual(
        app.callsTo("/api/refresh").map((call) => [call.method, call.body]),
        [["POST", { force: true }]],
        "Reload did not force a re-render",
    );
    assert.equal(app.el("status").textContent, "Reloaded from disk.");
    assert.equal(app.el("status").classList.contains("error"), false);
});

test("an error body is reported by its message, not by its status code", async (t) => {
    // `api()` reads `body.error.message`, falling back to the status. A panel
    // that showed "Request failed (500)" for a typed, worded failure would be
    // discarding the only part a reader can act on.
    const app = await loadApp({
        state: READY,
        routes: {
            "/api/open-in-word": respondWith(500, {
                error: { code: "file_locked", message: "report.docx is held by another program." },
            }),
        },
    });
    t.after(app.restore);

    app.el("openInWord").dispatch("click");
    await app.settle();

    assert.equal(app.el("status").textContent, "report.docx is held by another program.");
    assert.equal(app.el("status").classList.contains("error"), true);
});

test("a host-side error state is shown, and no document is drawn over it", async (t) => {
    // A document whose refresh failed terminally keeps its last-known `doc` and
    // a `pdfUrl` naming a render that may be gone. `render()` refuses to draw
    // it; `applyState` used to load it anyway on the first state a panel sees,
    // and a failed load then replaced the host's typed message with "Could not
    // display the document" -- an error about something never on screen, in
    // place of the one the reader could act on. Found by this file.
    const app = await loadApp({
        state: { ...READY, status: "error", error: { code: "file_locked", message: "Could not open the document." } },
    });
    t.after(app.restore);

    assert.equal(app.el("status").textContent, "Could not open the document.");
    assert.equal(app.el("status").classList.contains("error"), true);
    assert.equal(app.el("viewer").classList.contains("visible"), false, "the viewer was shown for a failed open");
    assert.equal(app.pdfjs.opened.length, 0, "a render was loaded for a document that failed to open");
});

test("clearing the error draws the document that was withheld", async (t) => {
    // The other half: refusing to draw an errored document must not leave the
    // panel unable to draw it once the error clears.
    const app = await loadApp({
        state: { ...READY, status: "error", error: { code: "file_locked", message: "Held by another program." } },
        pdfDocuments: { "/pdf/k1.pdf": { pages: 3 } },
    });
    t.after(app.restore);

    assert.equal(app.pdfjs.opened.length, 0);

    app.source.emit("state", READY);
    await app.settle();

    assert.deepEqual(app.pdfjs.opened.map((call) => call.url), ["/pdf/k1.pdf"], "the document stayed undrawn");
    assert.equal(app.el("viewer").classList.contains("visible"), true);
    assert.equal(app.el("status").hidden, true, "the cleared error is still on screen");
});
