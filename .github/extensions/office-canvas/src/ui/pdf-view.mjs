// The pdf.js rendering surface: a scrolling list of pages, each a canvas plus a
// selectable text layer, with the change overlay drawn on top.
//
// Runs in the canvas iframe, so `console` here is fine -- the JSON-RPC channel
// this would corrupt belongs to the host process, not to this document.
//
// ## Why not the host's PDF plugin
//
// ADR 0004. The plugin renders into a replaced element whose content is not in
// our DOM, so nothing can be drawn over it and nothing about it can be measured.
// Highlighting what the agent just changed -- the entire point of the canvas --
// is impossible there and straightforward here.
//
// ## Rendering is lazy, and that is affordable because the export is fast
//
// Pages start as empty boxes sized from `getViewport`, which needs no
// rasterisation, and are painted when they scroll near the viewport. The whole
// document is one PDF fetched once: measured, pdf.js issues a single unranged
// GET and zero Range requests, and forcing ranged mode cost 14 round trips to
// save nothing on files this size.
//
// ## Scale, and what a *fit* is (#106)
//
// The scale is one number for the whole document, and it is either a standing
// fit -- recomputed whenever the panel changes size -- or a scale the reader
// picked with a zoom button. `#fitMode` is what tells those apart; without it a
// resize would either discard the reader's zoom or leave a fit fitted to a panel
// width that no longer exists.
//
// A re-scale is not a resize of the boxes. Every page is reset and repainted,
// because a canvas holds a bitmap rasterised at one scale and stretching it is
// exactly the soft text this viewer exists not to have.

import * as pdfjs from "/vendor/pdf.min.mjs";
import { candidatePages, planChangeMarks } from "./change-plan.mjs";
import { describeChange } from "./change-wording.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.mjs";

/** How far outside the viewport a page is painted, as a fraction of its height. */
const RENDER_AHEAD = "150% 0px";

/** The zoom range, and the factor one press of a zoom button moves by. */
const MIN_SCALE = 0.2;
const MAX_SCALE = 4;
const ZOOM_STEP = 1.25;

/**
 * `.pages`'s own padding, on both axes, which a fit has to leave room for.
 *
 * Named because both fits now read it and the number belongs to `app.css`. A fit
 * that ignored it would size the page to the box it sits in rather than to the
 * space inside that box, and fit-width would produce a horizontal scrollbar --
 * which is the one thing fit-width is for not having.
 */
const PAGE_GUTTER = 32;

const clampScale = (value) => Math.max(MIN_SCALE, Math.min(MAX_SCALE, value));

export class PdfView {
    #container;
    #doc = null;
    #pages = [];
    #observer = null;
    /**
     * Watches the panel so a fit stays fitted.
     *
     * This field was declared and never assigned until #106. A fit computed once
     * at load is "fit to whatever the width was when the document opened", which
     * is not what the word means the moment the panel is dragged.
     */
    #scaleObserver = null;
    #onPageChange;
    /**
     * Announces a scale the caller did not ask for.
     *
     * A resize refits without anyone pressing a button, and the clamp bounds
     * gate whether those buttons are *enabled*, so a caller that only resynced
     * after its own presses would go stale: fit-width in a very narrow panel
     * clamps to `MIN_SCALE` and correctly disables zoom-out, then widening the
     * panel refits above the clamp with no press to resync from.
     */
    #onScaleChange;
    #scale = 1;
    /**
     * Which fit the reader is in, or `null` once they have zoomed by hand.
     *
     * A fit is a *standing* instruction -- recompute on every resize -- and an
     * explicit zoom is a one-off. Holding only the number could not tell the two
     * apart, so a resize would either always refit (discarding a zoom) or never
     * refit (leaving a fit stale).
     */
    #fitMode = "width";
    #change = null;
    /** Bumped on every load so a render that outlives its document can bail. */
    #generation = 0;
    /**
     * Bumped whenever what a paint would produce changes -- a load *or* a
     * re-scale.
     *
     * `#generation` cannot serve here: a re-scale must not make an in-flight
     * `load()` destroy its own document, which is what bumping that field means.
     * And a paint that survived a re-scale is just as wrong as one that survived
     * a document: it draws the old scale's bitmap into the new scale's box, and
     * builds a text layer on a viewport nothing else is using any more.
     */
    #epoch = 0;

    constructor(container, { onPageChange = () => {}, onScaleChange = () => {} } = {}) {
        this.#container = container;
        this.#onPageChange = onPageChange;
        this.#onScaleChange = onScaleChange;
    }

    get pageCount() {
        return this.#doc?.numPages ?? 0;
    }

    get scale() {
        return this.#scale;
    }

    /** `"width"`, `"height"`, or `null` when the reader has zoomed by hand. */
    get fitMode() {
        return this.#fitMode;
    }

    get canZoomIn() {
        return this.#scale < MAX_SCALE;
    }

    get canZoomOut() {
        return this.#scale > MIN_SCALE;
    }

    /**
     * Loads a PDF, replacing whatever was showing.
     *
     * The previous document is destroyed rather than dropped: pdf.js holds a
     * worker-side handle per document, and leaking those across the refresh that
     * follows every edit would accumulate for as long as the canvas is open.
     */
    async load(url) {
        const generation = ++this.#generation;
        this.#epoch++;
        await this.#teardown();

        const task = pdfjs.getDocument({ url });
        const doc = await task.promise;
        if (generation !== this.#generation) {
            await doc.destroy();
            return;
        }
        this.#doc = doc;

        // A document opens fitted to the width of the panel it opens in, and
        // stays fitted until the reader says otherwise.
        this.#fitMode = "width";
        this.#scale = this.#fitScaleFor(await doc.getPage(1), "width") ?? 1;
        this.#container.style.setProperty("--total-scale-factor", String(this.#scale));

        for (let number = 1; number <= doc.numPages; number++) {
            const page = this.#createPagePlaceholder(number);
            this.#pages.push(page);
            this.#container.append(page.root);
        }
        await Promise.all(this.#pages.map((page) => this.#measure(page)));

        this.#observe();
        this.#applyChange();
    }

    /** Scrolls a page into view. Replaces v1's reload-the-whole-document hack. */
    goToPage(number) {
        const page = this.#pages[Math.max(1, Math.min(number, this.#pages.length)) - 1];
        if (!page) return;
        page.root.scrollIntoView({ block: "start", behavior: "auto" });
    }

    /**
     * Draws the document at `value`, clamped, and leaves whatever fit was on.
     *
     * The clamp is the same 0.2--4 the fit has always applied, so a zoom cannot
     * reach a scale a fit could not.
     */
    setScale(value) {
        this.#fitMode = null;
        this.#rescale(value);
    }

    zoomIn() {
        this.setScale(this.#scale * ZOOM_STEP);
    }

    zoomOut() {
        this.setScale(this.#scale / ZOOM_STEP);
    }

    /** Fits the page width to the panel, and keeps it fitted as the panel moves. */
    fitWidth() {
        this.#fitMode = "width";
        this.#refit();
    }

    /** Fits a whole page into the panel's height, and keeps it fitted. */
    fitHeight() {
        this.#fitMode = "height";
        this.#refit();
    }

    /**
     * Shows -- or clears -- the marker for what the last operation changed.
     *
     * Accepts a record, never an address. An address is a coordinate into one
     * read (ADR 0006); by the time a page paints, the read that minted it is
     * over. The record carries text and a page number instead, both of which
     * describe the document this PDF was exported from.
     */
    setChange(record) {
        this.#change = record ?? null;
        this.#applyChange();
    }

    async destroy() {
        this.#generation++;
        this.#epoch++;
        await this.#teardown();
    }

    // --- internals -----------------------------------------------------------

    async #teardown() {
        this.#observer?.disconnect();
        this.#observer = null;
        this.#scaleObserver?.disconnect();
        this.#scaleObserver = null;
        for (const page of this.#pages) {
            page.renderTask?.cancel();
            page.textLayer?.cancel();
            page.proxy?.cleanup();
        }
        this.#pages = [];
        this.#container.replaceChildren();
        const doc = this.#doc;
        this.#doc = null;
        if (doc) await doc.destroy().catch(() => {});
    }

    /**
     * The scrolling box the pages sit in, or `null` when there is none.
     *
     * `.pages` is the container this view writes into; `.viewer` is what the
     * reader actually sees through, and it is the box a fit must be fitted to.
     * Measuring the container instead would make fit-width self-referential --
     * `.pages` is as wide as its widest page once a page overflows, so a
     * fit-width taken from it would fit the pages to themselves.
     */
    #viewport() {
        return this.#container.closest(".viewer");
    }

    /**
     * The scale that fits `proxy` to the panel, or `null` when it cannot be
     * measured.
     *
     * `null` rather than 1: a panel reporting no size is a panel that has not
     * been laid out yet -- hidden, or mid-open -- and snapping a reader's zoom
     * to 1 because of it would be a visible jump with no cause. `load()` has
     * nothing to keep, so it reads `null` as 1; a later fit reads it as "leave
     * this alone".
     */
    #fitScaleFor(proxy, mode) {
        const unscaled = proxy.getViewport({ scale: 1 });
        const box = this.#viewport() ?? this.#container;
        const available = (mode === "height" ? box.clientHeight : box.clientWidth) - PAGE_GUTTER;
        if (!(available > 0)) return null;
        return clampScale(available / (mode === "height" ? unscaled.height : unscaled.width));
    }

    /** Recomputes the standing fit, if there is one, against the panel as it is now. */
    #refit() {
        if (!this.#doc || !this.#fitMode) return;
        const proxy = this.#pages[0]?.proxy;
        if (!proxy) return;
        const scale = this.#fitScaleFor(proxy, this.#fitMode);
        if (scale !== null) this.#rescale(scale);
    }

    /**
     * Redraws the whole document at a new scale.
     *
     * Every page is *reset*, not merely resized. `#renderPage` early-returns on
     * `page.rendered`, so resizing the boxes alone leaves each bitmap at the
     * scale it was painted at -- which looks right on a fresh load and wrong the
     * moment the reader zooms a second time.
     *
     * The repaint is driven from here rather than left to the
     * `IntersectionObserver`, because that observer reports a *change* of
     * intersection. A page that was on screen before the re-scale and is still
     * on screen after it has not changed, so no entry is delivered and nothing
     * would repaint it. `page.visible` is what the observer last said, and it is
     * the only record of that.
     */
    #rescale(value) {
        const scale = clampScale(value);
        if (!Number.isFinite(scale) || scale === this.#scale || this.#pages.length === 0) return;

        // The page under the top of the viewport, which is where the reader is
        // reading. Taken before the boxes move, restored after.
        const anchor = this.#visiblePage();

        this.#scale = scale;
        this.#epoch++;
        // Non-negotiable. pdf.js's own text-layer stylesheet sizes every span
        // with `round(down, var(--total-scale-factor) * Npx, ...)`; with the
        // property missing the declaration is invalid and the layer collapses,
        // taking text selection and the change overlay with it.
        this.#container.style.setProperty("--total-scale-factor", String(scale));

        for (const page of this.#pages) {
            this.#resetPage(page);
            this.#sizeToScale(page);
        }

        this.#applyChange();
        this.goToPage(anchor);
        for (const page of this.#pages) {
            if (page.visible) this.#renderPage(page).catch(() => {});
        }

        // Every scale change in the class funnels through here, including the
        // ones nobody pressed a button for, so this is the one place that can
        // tell a caller its zoom controls are out of date.
        this.#onScaleChange(scale);
    }

    /** Returns a painted page to an empty box, cancelling whatever is in flight. */
    #resetPage(page) {
        page.renderTask?.cancel();
        page.renderTask = null;
        page.textLayer?.cancel();
        page.textLayer = null;
        page.textLayerDiv.replaceChildren();
        page.rendered = false;
        // `page.items` is kept on purpose: the extracted text is the same text
        // at any scale, and the overlay needs it to re-place its boxes before
        // the page has finished repainting.
    }

    #createPagePlaceholder(number) {
        const root = document.createElement("div");
        root.className = "page";
        root.dataset.page = String(number);

        const canvas = document.createElement("canvas");
        canvas.className = "page-canvas";

        const textLayerDiv = document.createElement("div");
        textLayerDiv.className = "textLayer";

        const overlay = document.createElement("div");
        overlay.className = "overlay";
        overlay.hidden = true;

        root.append(canvas, textLayerDiv, overlay);
        return {
            number,
            root,
            canvas,
            textLayerDiv,
            overlay,
            proxy: null,
            rendered: false,
            /** What the `IntersectionObserver` last said about this page. */
            visible: false,
        };
    }

    /** Sizes the placeholder without painting it. */
    async #measure(page) {
        const proxy = await this.#doc.getPage(page.number);
        page.proxy = proxy;
        this.#sizeToScale(page);
    }

    /** The half of `#measure` a re-scale can do without fetching anything. */
    #sizeToScale(page) {
        const viewport = page.proxy.getViewport({ scale: this.#scale });
        page.viewport = viewport;
        page.root.style.width = `${Math.floor(viewport.width)}px`;
        page.root.style.height = `${Math.floor(viewport.height)}px`;
    }

    #observe() {
        this.#observer = new IntersectionObserver(
            (entries) => {
                for (const entry of entries) {
                    const page = this.#pages[Number(entry.target.dataset.page) - 1];
                    if (!page) continue;
                    page.visible = entry.isIntersecting;
                    if (entry.isIntersecting) this.#renderPage(page).catch(() => {});
                }
                this.#reportVisiblePage();
            },
            { root: this.#viewport(), rootMargin: RENDER_AHEAD },
        );
        for (const page of this.#pages) this.#observer.observe(page.root);

        // What makes a fit a fit. Observed on the **border** box: the content
        // box shrinks when a scrollbar appears, and a fit that shrank the pages
        // because a scrollbar appeared could remove the scrollbar, grow them
        // back, and oscillate. The border box does not move when a scrollbar
        // does, so the loop has no way to start.
        this.#scaleObserver = new ResizeObserver(() => this.#refit());
        this.#scaleObserver.observe(this.#viewport() ?? this.#container, { box: "border-box" });
    }

    #reportVisiblePage() {
        this.#onPageChange(this.#visiblePage());
    }

    /** The last page whose top has passed the top of the viewport, within 8px. */
    #visiblePage() {
        const viewerTop = this.#container.getBoundingClientRect().top;
        let best = 1;
        for (const page of this.#pages) {
            if (page.root.getBoundingClientRect().top - viewerTop <= 8) best = page.number;
        }
        return best;
    }

    async #renderPage(page) {
        if (page.rendered || !page.proxy) return;
        page.rendered = true;
        const epoch = this.#epoch;

        // Paint at device resolution. Without this the canvas is upscaled by the
        // browser and the text is visibly soft, which on a page-accurate viewer
        // reads as a rendering bug rather than a display setting.
        const ratio = pdfjs.OutputScale.pixelRatio;
        const viewport = page.viewport;
        page.canvas.width = Math.floor(viewport.width * ratio);
        page.canvas.height = Math.floor(viewport.height * ratio);
        page.canvas.style.width = `${Math.floor(viewport.width)}px`;
        page.canvas.style.height = `${Math.floor(viewport.height)}px`;

        const context = page.canvas.getContext("2d", { alpha: false });
        // Held locally as well as on the page. A re-scale that lands mid-paint
        // replaces `page.renderTask`, and awaiting the field rather than the
        // task would then await the *replacement* -- so this paint would carry
        // on believing it had finished its own.
        const renderTask = page.proxy.render({
            canvasContext: context,
            viewport,
            transform: ratio === 1 ? null : [ratio, 0, 0, ratio, 0, 0],
        });
        page.renderTask = renderTask;
        try {
            await renderTask.promise;
        } catch (err) {
            if (err?.name !== "RenderingCancelledException") throw err;
            return;
        }
        if (epoch !== this.#epoch) return;

        const textContent = await page.proxy.getTextContent();
        if (epoch !== this.#epoch) return;
        page.items = textContent.items;

        // Cleared here as well as in `#resetPage`: a cancelled text layer can
        // have appended spans before it stopped, and they would otherwise sit
        // under the new ones at the old scale.
        page.textLayerDiv.replaceChildren();
        const textLayer = new pdfjs.TextLayer({
            textContentSource: textContent,
            container: page.textLayerDiv,
            viewport,
        });
        page.textLayer = textLayer;
        await textLayer.render();
        if (epoch !== this.#epoch) return;

        // A page can be painted after the record arrived -- lazily, that is the
        // normal case for a change below the fold. Ask the planner which pages
        // matter rather than deciding here: the straddle page is a candidate
        // too, and a private answer to that question was wrong.
        if (this.#change && candidatePages(this.#change).includes(page.number)) {
            this.#applyChange();
        }
    }

    // --- the change overlay --------------------------------------------------

    /**
     * Marks the change on whichever pages can currently answer for it.
     *
     * The join is page + text, both structural. It is never an address, and it
     * is never a stored geometry: the record describes a document state, the
     * PDF was exported from that same state, and the text is looked up afresh
     * each time a page paints.
     */
    #applyChange() {
        for (const page of this.#pages) {
            page.overlay.hidden = true;
            page.overlay.replaceChildren();
            page.root.classList.remove("has-change");
        }
        if (!this.#change) return;

        // The policy lives in `change-plan.mjs` so it can be tested; what is left
        // here is projection into the DOM. A page the plan names is marked even
        // when its quads come out empty -- that is a page the text was found on,
        // and dropping it would quietly reinstate the miss the plan just avoided.
        for (const mark of planChangeMarks(this.#change, this.#pages)) {
            const page = this.#pages[mark.number - 1];
            if (!page) continue;
            this.#markPage(page, this.#change, mark.found ? this.#quadsFor(page, mark.found) : []);
        }
    }

    #quadsFor(page, found) {
        const rects = [];
        const { startItem, endItem } = found.range ?? {};
        if (!Number.isInteger(startItem) || !Number.isInteger(endItem)) return rects;
        for (let index = startItem; index <= endItem; index++) {
            const item = page.items[index];
            if (!item?.transform) continue;
            const [, , , , x, y] = pdfjs.Util.transform(page.viewport.transform, item.transform);
            const height = Math.hypot(item.transform[2], item.transform[3]) * this.#scale;
            const width = (item.width ?? 0) * this.#scale;
            if (!(width > 0) || !(height > 0)) continue;
            rects.push({ left: x, top: y - height, width, height });
        }
        return rects;
    }

    #markPage(page, record, quads) {
        page.root.classList.add("has-change");
        page.overlay.hidden = false;
        for (const rect of quads) {
            const box = document.createElement("div");
            box.className = "change-box";
            box.style.left = `${rect.left}px`;
            box.style.top = `${rect.top}px`;
            box.style.width = `${rect.width}px`;
            box.style.height = `${rect.height}px`;
            page.overlay.append(box);
        }
        const badge = document.createElement("div");
        badge.className = "change-badge";
        badge.textContent = quads.length ? describeChange(record) : `${describeChange(record)} — on this page`;
        page.overlay.append(badge);
    }
}
