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
    assert.equal(
        harness.resizeObservers[0].disconnected,
        1,
        "the resize watcher outlived the document it was refitting",
    );
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

// --- zoom and fit (#106) ---------------------------------------------------

/**
 * A view whose pages sit inside a `.viewer` of the given size.
 *
 * The fits measure the *viewport*, not the container they write into: `.pages`
 * is `min-width: max-content`, so it is as wide as its widest page once a page
 * overflows, and a fit-width taken from it would fit the pages to themselves.
 * A viewer is needed to assert the branch the panel actually runs.
 */
function inViewer(harness, { width = 632, height = 832 } = {}) {
    const viewer = new harness.StubElement("div");
    viewer.className = "viewer";
    viewer.clientWidth = width;
    viewer.clientHeight = height;
    viewer.append(harness.container);
    return viewer;
}

test("a re-scale repaints every visible page at the new scale rather than resizing its box", async (t) => {
    // #106's first trap, and the one a screenshot of a fresh load cannot catch.
    // `#renderPage` early-returns on `page.rendered`, so a re-scale that resizes
    // the boxes without clearing the flag leaves each bitmap rasterised at the
    // scale it was first painted at. It looks right once and wrong on the second
    // zoom, which is why the scale of each paint is asserted and not the count.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    const viewer = inViewer(harness, { width: 632 });
    pdfjs.serve("/pdf/a.pdf", { pages: 3, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(view.scale, 1, "600pt pages in a 632px viewer, less 32px of padding, is scale 1");

    harness.observers[0].enter(container.children[0]);
    await harness.settle();

    const page1 = pdfjs.opened[0].doc.pages.get(1);
    assert.deepEqual(page1.renderScales, [1]);

    view.zoomIn();
    await harness.settle();

    assert.equal(view.scale, 1.25);
    assert.deepEqual(page1.renderScales, [1, 1.25], "the box was resized but the bitmap was not repainted");
    assert.equal(page1.renderCancels, 1, "the superseded paint was left to finish into the new boxes");
    assert.equal(container.children[0].style.width, "750px", "the box did not follow the new scale");

    // Twice. The single-zoom case passes even when `rendered` is never cleared,
    // because the first paint of a page that had never painted is the same
    // whether the flag was reset or was still false.
    view.zoomIn();
    await harness.settle();
    assert.deepEqual(page1.renderScales, [1, 1.25, 1.5625]);
});

test("a re-scale leaves the pages nobody is looking at unpainted", async (t) => {
    // The repaint is driven from the re-scale, not from the observer, because an
    // `IntersectionObserver` reports a *change* of intersection and a page that
    // was on screen before and after has not changed. Driving it from there must
    // not cost the laziness: on a 13-page document only what is on screen paints.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness);
    pdfjs.serve("/pdf/a.pdf", { pages: 3, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[0]);
    await harness.settle();

    view.zoomIn();
    await harness.settle();

    const pages = pdfjs.opened[0].doc.pages;
    assert.equal(pages.get(1).renders, 2);
    assert.equal(pages.get(2)?.renders ?? 0, 0, "a page nobody had scrolled to was painted by the zoom");
    assert.equal(pages.get(3)?.renders ?? 0, 0);
});

test("a page that has scrolled out is not repainted by a re-scale", async (t) => {
    // `page.visible` is what the observer last said, so it has to be cleared when
    // a page leaves. Recording only arrivals would make every page ever seen
    // repaint on every zoom, which on a long document is the whole file.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness);
    pdfjs.serve("/pdf/a.pdf", { pages: 3, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[0]);
    await harness.settle();
    harness.observers[0].leave(container.children[0]);
    harness.observers[0].enter(container.children[1]);
    await harness.settle();

    view.zoomIn();
    await harness.settle();

    const pages = pdfjs.opened[0].doc.pages;
    assert.equal(pages.get(2).renders, 2, "the page on screen was not repainted");
    assert.equal(pages.get(1).renders, 1, "a page that had scrolled away was repainted anyway");
});

test("--total-scale-factor follows every re-scale", async (t) => {
    // Non-negotiable, and guarded upstream by TEXT_LAYER_CONTRACT: pdf.js's own
    // text-layer stylesheet sizes every span with
    // `round(down, var(--total-scale-factor) * Npx, ...)`. With the property
    // stale the text sits off its glyphs; with it missing the declaration is
    // invalid and the whole layer collapses, taking selection and the change
    // overlay with it.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 632, height: 432 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    const factor = () => container.style.getPropertyValue("--total-scale-factor");
    assert.equal(factor(), "1");

    view.zoomIn();
    assert.equal(factor(), "1.25");

    view.zoomOut();
    assert.equal(factor(), "1");

    view.fitHeight();
    assert.equal(view.scale, 0.5, "fit-height did not change the scale, so this asserts nothing");
    assert.equal(factor(), "0.5");
});

test("zoom clamps at both ends, and the view says which presses would do nothing", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 632 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(view.canZoomIn, true);
    assert.equal(view.canZoomOut, true);

    for (let press = 0; press < 40; press++) view.zoomIn();
    assert.equal(view.scale, 4, "zoom ran past the top of the range");
    assert.equal(view.canZoomIn, false);

    for (let press = 0; press < 40; press++) view.zoomOut();
    assert.equal(view.scale, 0.2, "zoom ran past the bottom of the range");
    assert.equal(view.canZoomOut, false);
    assert.equal(view.canZoomIn, true);
});

test("fit-width and fit-height measure the viewport, not the pages they are sizing", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    // A viewer 432px wide and 232px tall, less 32px of `.pages` padding: 400px
    // across 800px of page height, and 200px down 800pt of it.
    inViewer(harness, { width: 432, height: 232 });
    pdfjs.serve("/pdf/a.pdf", { pages: 2, width: 800, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(view.scale, 0.5, "fit-width did not fit the page across the viewer");
    assert.equal(view.fitMode, "width");

    // The container is now wider than the viewer, which is what `min-width:
    // max-content` does. A fit measuring it would grow without bound.
    container.clientWidth = 4000;

    view.fitHeight();
    assert.equal(view.scale, 0.25, "fit-height did not fit a whole page down the viewer");
    assert.equal(view.fitMode, "height");

    view.fitWidth();
    assert.equal(view.scale, 0.5, "the fit measured the pages it had just resized");
});

test("a panel resize refits a standing fit and leaves a hand-picked zoom alone", async (t) => {
    // What `#scaleObserver` is for (#106). A fit is a standing instruction --
    // "keep this fitted" -- and a fit computed once at load is only ever a fit to
    // whatever the panel happened to be when the document opened. An explicit
    // zoom is the opposite: a resize must not discard it.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    const viewer = inViewer(harness, { width: 632, height: 832 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(view.scale, 1);

    const resize = harness.resizeObservers[0];
    assert.equal(resize.targets[0], viewer, "the watcher is on the pages, which resize because of it");
    assert.deepEqual(
        resize.options,
        { box: "border-box" },
        "watching the content box lets a scrollbar appearing shrink the pages that removed it",
    );

    viewer.clientWidth = 332;
    resize.fire();
    assert.equal(view.scale, 0.5, "the fit did not survive the panel resize");
    assert.equal(container.children[0].style.width, "300px");

    view.zoomIn();
    assert.equal(view.fitMode, null, "zooming by hand left a fit standing");
    const picked = view.scale;

    viewer.clientWidth = 632;
    resize.fire();
    assert.equal(view.scale, picked, "a resize discarded the scale the reader had chosen");
});

test("a resize that cannot be measured leaves the scale where it is", async (t) => {
    // A panel reporting no width is a panel that has not been laid out -- hidden,
    // or mid-open. Treating that as "fit to nothing" would snap the reader's
    // document to scale 1 for a reason they cannot see.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    const viewer = inViewer(harness, { width: 332 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    assert.equal(view.scale, 0.5);

    viewer.clientWidth = 0;
    harness.resizeObservers[0].fire();

    assert.equal(view.scale, 0.5, "an unmeasurable panel snapped the document to a scale nobody asked for");
});

test("a re-scale keeps the reader on the page they were reading", async (t) => {
    // Every box moves, so scroll position means something different afterwards.
    // Without the anchor a zoom on page 9 of 13 lands wherever the old offset
    // now points, which is a different page.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness);
    pdfjs.serve("/pdf/a.pdf", { pages: 3, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");

    container.rect = { top: 0, left: 0, width: 600, height: 800 };
    container.children[0].rect = { top: -1620, left: 0 };
    container.children[1].rect = { top: -810, left: 0 };
    container.children[2].rect = { top: 0, left: 0 };

    view.zoomIn();
    await harness.settle();

    assert.equal(container.children[2].scrolledIntoView, 1, "the zoom did not return to the page being read");
    assert.equal(container.children[0].scrolledIntoView, 0);
});

test("a re-scale re-places the change overlay instead of leaving it at the old scale", async (t) => {
    // The overlay is drawn in page coordinates, so its boxes are as scale-bound
    // as the bitmap under them. A zoom that moved the glyphs and not the
    // highlight would point at the wrong words.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 632 });
    pdfjs.serve("/pdf/a.pdf", {
        width: 600,
        height: 800,
        pageItems: [[pdfjs.textItem(PARAGRAPH, { x: 40, y: 700 })]],
    });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[0]);
    await harness.settle();
    view.setChange(change({ page: 1 }));

    const overlay = partsOf(container.children[0]).overlay;
    const before = boxes(overlay)[0];
    assert.ok(before, "the change was never marked, so this asserts nothing about scaling it");
    const width = Number.parseFloat(before.style.width);

    view.zoomIn();
    await harness.settle();

    const after = boxes(overlay)[0];
    assert.ok(after, "the zoom dropped the change marker");
    assert.equal(
        Number.parseFloat(after.style.width).toFixed(4),
        (width * 1.25).toFixed(4),
        "the marker stayed at the scale it was drawn at",
    );
});

test("a re-scale clears the text layer instead of stacking a second one over it", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 632 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    harness.observers[0].enter(container.children[0]);
    await harness.settle();

    view.zoomIn();
    await harness.settle();

    assert.equal(pdfjs.TextLayer.instances.length, 2, "the page was not given a text layer at the new scale");
    assert.equal(pdfjs.TextLayer.instances[0].cancelled, 1, "the old text layer was left running");
    assert.equal(
        pdfjs.TextLayer.instances[1].options.viewport.scale,
        1.25,
        "the new text layer was built on the old scale's viewport",
    );
    assert.equal(
        partsOf(container.children[0]).textLayer,
        pdfjs.TextLayer.instances[1].options.container,
        "the replacement layer was built somewhere else",
    );
});

test("a zoom during a load does not destroy the document being loaded", async (t) => {
    // `#epoch` exists because `#generation` cannot do this job: bumping the
    // generation is how a load says "the document being opened is superseded",
    // and a zoom that bumped it would make an in-flight load destroy its own
    // document and leave the panel empty.
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 632 });
    pdfjs.serve("/pdf/slow.pdf", { pages: 2, width: 600, height: 800, deferred: true });

    const view = new PdfView(container);
    const loading = view.load("/pdf/slow.pdf");
    await harness.settle();

    // The zoom cluster is live while a document is opening, so this is a press a
    // reader can actually make. If a re-scale bumped `#generation`, the load's
    // own guard would fire on its return and destroy the document it had just
    // opened -- leaving an empty panel and no error.
    view.zoomIn();
    view.fitWidth();
    pdfjs.opened[0].resolve();
    await loading;

    assert.equal(pdfjs.opened[0].doc.destroyed, 0, "the load destroyed the document it had just opened");
    assert.equal(view.pageCount, 2);
});

test("opening a document returns to fit-width, whatever the last one was left at", async (t) => {
    const harness = await loadPdfView();
    t.after(harness.restore);
    const { PdfView, pdfjs, container } = harness;

    inViewer(harness, { width: 332 });
    pdfjs.serve("/pdf/a.pdf", { pages: 1, width: 600, height: 800 });
    pdfjs.serve("/pdf/b.pdf", { pages: 1, width: 600, height: 800 });

    const view = new PdfView(container);
    await view.load("/pdf/a.pdf");
    view.zoomIn();
    assert.equal(view.fitMode, null);

    await view.load("/pdf/b.pdf");

    assert.equal(view.fitMode, "width", "the new document opened at the previous one's hand-picked zoom");
    assert.equal(view.scale, 0.5);
});
