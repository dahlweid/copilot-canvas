// The pdf.js stand-in that `/vendor/pdf.min.mjs` resolves to under Node.
//
// `pdf-view.mjs` imports pdf.js from an absolute URL because the browser loads
// it over the render server. That specifier is correct for the product and
// meaningless on a filesystem -- under Node it resolves to `C:\vendor\` and
// raises ERR_MODULE_NOT_FOUND -- so `ui-harness.mjs` redirects it here with
// `module.registerHooks`, the same technique `extension-stubs.mjs` uses for the
// SDK. The viewer is not modified to suit the harness (#76 rules that out
// explicitly).
//
// The surface implemented is exactly what `pdf-view.mjs` reaches for:
// `GlobalWorkerOptions`, `getDocument`, `OutputScale.pixelRatio`, `TextLayer`
// and `Util.transform`. Nothing decodes a PDF; a document is described by the
// test that asks for it.

export const GlobalWorkerOptions = { workerSrc: null };

/** Device pixel ratio. 1 keeps canvas arithmetic in the tests readable. */
export const OutputScale = { pixelRatio: 1 };

export const Util = {
    /**
     * Matrix multiply, copied from pdf.js's own `Util.transform`.
     *
     * The overlay's quads come out of this, so an invented formula here would
     * make every geometry assertion a statement about this file rather than
     * about `pdf-view.mjs`.
     */
    transform(m, t) {
        return [
            m[0] * t[0] + m[2] * t[1],
            m[1] * t[0] + m[3] * t[1],
            m[0] * t[2] + m[2] * t[3],
            m[1] * t[2] + m[3] * t[3],
            m[0] * t[4] + m[2] * t[5] + m[4],
            m[1] * t[4] + m[3] * t[5] + m[5],
        ];
    },
};

export class TextLayer {
    static instances = [];

    constructor(options) {
        this.options = options;
        this.rendered = 0;
        this.cancelled = 0;
        TextLayer.instances.push(this);
    }

    async render() {
        this.rendered += 1;
    }

    cancel() {
        this.cancelled += 1;
    }
}

/** What each URL should hand back, keyed by the URL the view asks for. */
const specs = new Map();

/** Every `getDocument` call, in order. */
export const opened = [];

/** The document served for a URL no test described. */
const DEFAULT_SPEC = { pages: 1 };

/**
 * Describes the document a URL yields.
 *
 * `pages` is a page count, or `pageItems` an array of pdf.js text-item arrays --
 * one per page -- for a test that needs the overlay to find something. `deferred`
 * withholds the document until the recorded call is resolved by hand, which is
 * how the load-generation guard is exercised.
 */
export function serve(url, spec) {
    specs.set(url, spec);
}

export function reset() {
    specs.clear();
    opened.length = 0;
    TextLayer.instances.length = 0;
    GlobalWorkerOptions.workerSrc = null;
    OutputScale.pixelRatio = 1;
}

export function getDocument({ url }) {
    const spec = specs.get(url) ?? DEFAULT_SPEC;
    const doc = makeDocument(spec);
    const call = { url, doc, resolve: null };
    opened.push(call);

    const promise = spec.deferred
        ? new Promise((resolve, reject) => {
              call.resolve = () => resolve(doc);
              call.reject = reject;
          })
        : Promise.resolve(doc);

    return { promise, destroy: async () => {} };
}

function makeDocument(spec) {
    const pageItems = spec.pageItems ?? Array.from({ length: spec.pages ?? 1 }, () => []);
    const pages = new Map();

    const doc = {
        numPages: pageItems.length,
        destroyed: 0,
        async getPage(number) {
            if (!pages.has(number)) pages.set(number, makePage(number, pageItems[number - 1] ?? [], spec));
            return pages.get(number);
        },
        async destroy() {
            doc.destroyed += 1;
        },
        /** The proxies handed out so far, for a test that wants to inspect one. */
        pages,
    };
    return doc;
}

function makePage(number, items, spec) {
    const width = spec.width ?? 600;
    const height = spec.height ?? 800;
    return {
        pageNumber: number,
        renders: 0,
        cleanups: 0,
        getViewport({ scale }) {
            return {
                width: width * scale,
                height: height * scale,
                // PDF user space is y-up and the viewport flips it, exactly as
                // pdf.js does. The overlay's `top` depends on this sign.
                transform: [scale, 0, 0, -scale, 0, height * scale],
                scale,
            };
        },
        render(options) {
            this.renders += 1;
            this.lastRender = options;
            return { promise: Promise.resolve(), cancel() {} };
        },
        async getTextContent() {
            return { items };
        },
        cleanup() {
            this.cleanups += 1;
        },
    };
}

/** A pdf.js text item, with the fields `locate-text.mjs` and the overlay read. */
export function textItem(str, { x = 40, y = 700, size = 12, width = null, hasEOL = false } = {}) {
    return {
        str,
        hasEOL,
        width: width ?? str.length * (size * 0.5),
        height: size,
        transform: [size, 0, 0, size, x, y],
    };
}
