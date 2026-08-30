// `src/ui/pdf-view.mjs`, executed.
//
// This is the file the extract-the-logic workaround could never reach. Its whole
// job is to drive the thing that would not load -- pdf.js behind an absolute
// `/vendor/` specifier -- so the decisions worth testing were moved out into
// `change-plan.mjs` and `locate-text.mjs` and what was left, the projection into
// the DOM and the lifetime of a document, had no reachable assertion at all
// (#76). `ui-harness.mjs` resolves the specifier to `stub-pdfjs.mjs` and
// supplies a `document`; the view under test is the committed one.
//
// The failures aimed at here are the ones this repo has actually paid for:
// leaking a document across a reload, and a render that outlives the document
// it belongs to. Neither is visible from outside the class, so both are asserted
// through what reaches pdf.js and what reaches the container.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";

import { loadPdfView } from "./ui-harness.mjs";

const PARAGRAPH = "The quick brown fox jumps over it.";

/** A page's three layers, found by class rather than by position. */
function partsOf(root) {
    const find = (name) => root.children.find((child) => child.classList?.contains(name));
    return { canvas: find("page-canvas"), textLayer: find("textLayer"), overlay: find("overlay") };
}

const boxes = (overlay) => overlay.children.filter((child) => child.classList.contains("change-box"));
const badge = (overlay) => overlay.children.find((child) => child.classList.contains("change-badge"));

/** A change record as `change-record.mjs` mints them. */
const change = (overrides = {}) => ({
    op: "replace_text",
    page: 2,
    text: PARAGRAPH,
    locatable: true,
    at: 1_700_000_000_000,
    ...overrides,
});

test("loading a second document destroys the first and replaces its pages", async (t) => {
    // The leak this guards is a worker-side handle per document. Dropping the
    // previous one rather than destroying it would accumulate for every refresh
    // -- and a refresh follows every edit -- for as long as the canvas is open.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pages: 3 });
    pdfjs.serve("/pdf/b.pdf", { pages: 2 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(container.children.length, 3);
    assert.equal(view.pageCount, 3);

    await view.load("/pdf/b.pdf");

    assert.equal(pdfjs.opened[0].doc.destroyed, 1, "the previous document was dropped rather than destroyed");
    assert.equal(view.pageCount, 2);
    assert.equal(container.children.length, 2, "the previous document's pages were left in the container");
    assert.equal(harness.observers[0].disconnected, 1, "the previous document's observer is still watching");
});

test("a load overtaken by a newer one destroys its own document instead of installing it", async (t) => {
    // The `#generation` guard. Two refreshes in quick succession is the ordinary
    // way to reach this: the first document arrives late, and without the guard
    // it would overwrite the newer one -- and never be destroyed.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/slow.pdf", { pages: 3, deferred: true });
    pdfjs.serve("/pdf/fast.pdf", { pages: 2 });

    const view = new PdfView(container);
    const slow = view.load("/pdf/slow.pdf");
    const fast = view.load("/pdf/fast.pdf");
    await fast;

    pdfjs.opened[0].resolve();
    await slow;

    assert.equal(pdfjs.opened[0].url, "/pdf/slow.pdf");
    assert.equal(pdfjs.opened[0].doc.destroyed, 1, "the superseded document was left alive");
    assert.equal(view.pageCount, 2, "the superseded document overwrote the current one");
    assert.equal(container.children.length, 2);
});

test("goToPage clamps to the document rather than scrolling nowhere", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pages: 3 });
    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");

    view.goToPage(99);
    view.goToPage(0);
    view.goToPage(2);

    assert.deepEqual(
        container.children.map((root) => root.scrolledIntoView),
        [1, 1, 1],
        "a page number outside the document did not clamp to an end",
    );
});

test("the pages are fitted to the container's width, and the text layer is told the scale", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    // 600pt wide pages, 32px of chrome: 300/600 halves them.
    container.clientWidth = 332;
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");

    // pdf.js's own text layer reads this custom property; a text layer scaled
    // differently from the canvas under it is selectable text that does not sit
    // on the glyphs.
    assert.equal(container.style.getPropertyValue("--total-scale-factor"), "0.5");
    assert.equal(container.children[0].style.width, "300px");
    assert.equal(container.children[0].style.height, "400px");
});

test("a page paints only once it is scrolled near, and gets a text layer of its own", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pages: 3 });
    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");

    const doc = pdfjs.opened[0].doc;
    assert.equal(doc.pages.get(1).renders, 0, "a page was painted before anything looked at it");

    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    assert.equal(doc.pages.get(2).renders, 1);
    assert.equal(doc.pages.get(1).renders, 0, "an unseen page was painted too");
    assert.equal(pdfjs.TextLayer.instances.length, 1);
    assert.equal(
        pdfjs.TextLayer.instances[0].options.container,
        partsOf(container.children[1]).textLayer,
        "the text layer was built over the wrong page",
    );

    // Painting twice would stack a second text layer over the first.
    harness.observers[0].enter(container.children[1]);
    await harness.settle();
    assert.equal(doc.pages.get(2).renders, 1, "a page already painted was painted again");
});

test("a page painted after the change arrived is still marked", async (t) => {
    // The lazy-render case, which is the *normal* one for a change below the
    // fold: the record is set while the page is an empty box with no text to
    // search. `#renderPage` asks the planner whether the page it just painted is
    // a candidate; a private answer to that question was wrong once already.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    // Deliberately not 1:1. At scale 1 the viewport's x term reduces to the
    // item's own translate, so a box asserted there is consistent with the
    // matrix multiply never happening -- measured: replacing pdf.js's formula
    // with the identity left this assertion green. Half scale separates them.
    container.clientWidth = 332;
    pdfjs.serve("/pdf/a.pdf", {
        width: 600,
        height: 800,
        pageItems: [[pdfjs.textItem("Something else entirely.")], [pdfjs.textItem(PARAGRAPH, { x: 40, y: 700 })]],
    });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    view.setChange(change());

    const overlay = partsOf(container.children[1]).overlay;
    assert.equal(overlay.hidden, true, "an unpainted page was marked from text it could not have read");

    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    assert.equal(overlay.hidden, false, "the page painted and the marker never appeared");
    assert.equal(container.children[1].classList.contains("has-change"), true);
    assert.equal(badge(overlay).textContent, "Text replaced");

    // Geometry, from the item transform through pdf.js's own viewport flip.
    assert.equal(boxes(overlay).length, 1);
    assert.deepEqual(
        ["left", "top", "width", "height"].map((side) => boxes(overlay)[0].style[side]),
        // The item sits at (40, 700) in y-up user space on an 800pt page, 12pt
        // tall and 204pt wide. At half scale the viewport flip puts its baseline
        // at y = 400 - 350 = 50, and the box's top 6 above that.
        ["20px", "44px", "102px", "6px"],
    );

    assert.equal(partsOf(container.children[0]).overlay.hidden, true, "an unrelated page was marked too");
});

test("a page named by the plan but with the text nowhere on it is marked without a box", async (t) => {
    // `planChangeMarks` returns the reported page with `found: null` when every
    // candidate was searchable and none held the text. Dropping that page
    // because its quads came out empty would quietly reinstate the miss the plan
    // just avoided, so it is marked -- and the badge has to say that it is the
    // page and not the words.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", {
        pageItems: [[pdfjs.textItem("Page one text.")], [pdfjs.textItem("Page two text.")]],
    });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[0], container.children[1]);
    await harness.settle();

    view.setChange(change({ text: "A paragraph on neither page." }));

    const overlay = partsOf(container.children[1]).overlay;
    assert.equal(overlay.hidden, false, "the reported page went unmarked");
    assert.equal(boxes(overlay).length, 0, "a box was drawn over text that was never located");
    assert.equal(badge(overlay).textContent, "Text replaced — on this page");
});

test("clearing the change removes every marker", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pageItems: [[], [pdfjs.textItem(PARAGRAPH)]] });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    view.setChange(change());
    const overlay = partsOf(container.children[1]).overlay;
    assert.equal(overlay.hidden, false);

    view.setChange(null);

    assert.equal(overlay.hidden, true, "dismissing the change left its marker on the page");
    assert.equal(overlay.children.length, 0, "the marker's boxes and badge outlived it");
    assert.equal(container.children[1].classList.contains("has-change"), false);
});

test("destroy releases the document, the observer and the page proxies", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pages: 2 });
    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    await view.destroy();

    const doc = pdfjs.opened[0].doc;
    assert.equal(doc.destroyed, 1, "the document outlived the view");
    assert.equal(harness.observers[0].disconnected, 1, "the observer is still holding the removed pages");
    assert.deepEqual([...doc.pages.values()].map((page) => page.cleanups), [1, 1]);
    assert.equal(container.children.length, 0);
    assert.equal(view.pageCount, 0);
});

test("the page reported as visible is the last one whose top has passed, within 8px", async (t) => {
    // The report is what a re-opened canvas restores position from, so a page
    // number that never moves is a bug the reader meets as "it forgot where I
    // was". The tolerance is asserted from both sides: a rule with a boundary
    // that no test crosses is a rule no test has read.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    pdfjs.serve("/pdf/a.pdf", { pages: 3 });
    const reported = [];
    const view = new PdfView(container, { onPageChange: (page) => reported.push(page) });
    await view.load("/pdf/a.pdf");

    container.rect = { top: 0, left: 0, width: 600, height: 800 };
    const tops = (values) => {
        for (const [index, top] of values.entries()) container.children[index].rect = { top, left: 0 };
    };

    // Page 1 filling the viewport; pages 2 and 3 below it.
    tops([0, 820, 1640]);
    harness.observers[0].enter(container.children[0]);
    await harness.settle();

    // Scrolled until page 2's top is exactly at the tolerance.
    tops([-812, 8, 828]);
    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    // One pixel short of it: still page 1.
    tops([-811, 9, 829]);
    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    assert.deepEqual(reported, [1, 2, 1], "the visible page was misreported");
});
