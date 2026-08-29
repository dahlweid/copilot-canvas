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

import * as pdfjs from "/vendor/pdf.min.mjs";
import { planChangeMarks } from "./change-plan.mjs";
import { describeChange } from "./change-wording.mjs";

pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdf.worker.min.mjs";

/** How far outside the viewport a page is painted, as a fraction of its height. */
const RENDER_AHEAD = "150% 0px";

export class PdfView {
    #container;
    #doc = null;
    #pages = [];
    #observer = null;
    #scaleObserver = null;
    #onPageChange;
    #scale = 1;
    #change = null;
    /** Bumped on every load so a render that outlives its document can bail. */
    #generation = 0;

    constructor(container, { onPageChange = () => {} } = {}) {
        this.#container = container;
        this.#onPageChange = onPageChange;
    }

    get pageCount() {
        return this.#doc?.numPages ?? 0;
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
        await this.#teardown();

        const task = pdfjs.getDocument({ url });
        const doc = await task.promise;
        if (generation !== this.#generation) {
            await doc.destroy();
            return;
        }
        this.#doc = doc;

        this.#scale = this.#fitWidthScale(await doc.getPage(1));
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
        await this.#teardown();
    }

    // --- internals -----------------------------------------------------------

    async #teardown() {
        this.#observer?.disconnect();
        this.#observer = null;
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

    #fitWidthScale(page) {
        const unscaled = page.getViewport({ scale: 1 });
        const available = this.#container.clientWidth - 32;
        if (!(available > 0)) return 1;
        return Math.max(0.2, Math.min(4, available / unscaled.width));
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
        return { number, root, canvas, textLayerDiv, overlay, proxy: null, rendered: false };
    }

    /** Sizes the placeholder without painting it. */
    async #measure(page) {
        const proxy = await this.#doc.getPage(page.number);
        page.proxy = proxy;
        const viewport = proxy.getViewport({ scale: this.#scale });
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
                    if (entry.isIntersecting) this.#renderPage(page).catch(() => {});
                }
                this.#reportVisiblePage();
            },
            { root: this.#container.closest(".viewer") ?? null, rootMargin: RENDER_AHEAD },
        );
        for (const page of this.#pages) this.#observer.observe(page.root);
    }

    #reportVisiblePage() {
        const viewerTop = this.#container.getBoundingClientRect().top;
        let best = 1;
        for (const page of this.#pages) {
            if (page.root.getBoundingClientRect().top - viewerTop <= 8) best = page.number;
        }
        this.#onPageChange(best);
    }

    async #renderPage(page) {
        if (page.rendered || !page.proxy) return;
        page.rendered = true;
        const generation = this.#generation;

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
        page.renderTask = page.proxy.render({
            canvasContext: context,
            viewport,
            transform: ratio === 1 ? null : [ratio, 0, 0, ratio, 0, 0],
        });
        try {
            await page.renderTask.promise;
        } catch (err) {
            if (err?.name !== "RenderingCancelledException") throw err;
            return;
        }
        if (generation !== this.#generation) return;

        const textContent = await page.proxy.getTextContent();
        if (generation !== this.#generation) return;
        page.items = textContent.items;

        page.textLayer = new pdfjs.TextLayer({
            textContentSource: textContent,
            container: page.textLayerDiv,
            viewport,
        });
        await page.textLayer.render();

        // A page can be painted after the record arrived -- lazily, that is the
        // normal case for a change below the fold.
        if (this.#change?.page === page.number) this.#applyChange();
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
