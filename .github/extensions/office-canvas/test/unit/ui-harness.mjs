// Makes the committed viewer modules importable from a test, and nothing more.
//
// Two blockers stand between `src/ui/app.js` and a Node import, and the second
// is only visible once the first is gone (#76 measured both):
//
//   1. `pdf-view.mjs` imports pdf.js from an absolute `/vendor/` URL, which is
//      right for a browser loading it over the render server and resolves to
//      `C:\vendor\pdf.min.mjs` under Node. `module.registerHooks` redirects it
//      to `stub-pdfjs.mjs` -- the same technique, for the same reason, as
//      `extension-stubs.mjs` uses on the SDK specifier.
//   2. `app.js` builds its element map at **module scope**, so `document` is
//      touched before any function is called. Removing blocker 1 alone gets a
//      `document is not defined` and no further. That is what `ui-dom.mjs` is
//      for.
//
// The viewer is not modified to suit this: no specifier is rewritten and no
// export is added. What runs under the assertions is the committed code.
//
// Importing this module installs the resolve hook (`registerHooks` is
// synchronous and in-thread, so `node --test` needs no extra flag); the browser
// globals are installed per test by `loadApp`, which returns the handle that
// removes them again.

import { readFile } from "node:fs/promises";
import { registerHooks } from "node:module";

import { createDocument, elementIdsFromMarkup, StubElement } from "./ui-dom.mjs";

const VENDOR_PREFIX = "/vendor/";

export const stubPdfjsUrl = new URL("./stub-pdfjs.mjs", import.meta.url).href;

/** The `/vendor/` specifiers that have a stand-in. Anything else must be noticed. */
const VENDOR_STUBS = new Map([["/vendor/pdf.min.mjs", stubPdfjsUrl]]);

registerHooks({
    resolve(specifier, context, nextResolve) {
        if (specifier.startsWith(VENDOR_PREFIX)) {
            const url = VENDOR_STUBS.get(specifier);
            // A new vendored import would otherwise resolve against the
            // filesystem root and fail somewhere less obvious, or -- worse --
            // find something. Say which specifier has no stand-in.
            if (!url) throw new Error(`no test stand-in for the vendored specifier '${specifier}'`);
            return { url, shortCircuit: true };
        }
        return nextResolve(specifier, context);
    },
});

const APP_URL = new URL("../../src/ui/app.js", import.meta.url).href;
const INDEX_URL = new URL("../../src/ui/index.html", import.meta.url);

/**
 * Lets the module graph settle.
 *
 * `setImmediate` runs after the microtask queue drains, so each turn clears
 * every promise chain that was ready; several turns clear chains that schedule
 * further work, which the state-then-render-then-load path does.
 */
export async function settle(turns = 8) {
    for (let i = 0; i < turns; i += 1) await new Promise((resolve) => setImmediate(resolve));
}

/** Replaces globals, and hands back the undo. */
function installGlobals(values) {
    const saved = Object.keys(values).map((name) => [name, Object.getOwnPropertyDescriptor(globalThis, name)]);
    for (const [name, value] of Object.entries(values)) {
        // `navigator` is an accessor with no setter on modern Node, so a plain
        // assignment throws in a module. Define, do not assign.
        Object.defineProperty(globalThis, name, { value, writable: true, configurable: true, enumerable: true });
    }
    return () => {
        for (const [name, descriptor] of saved) {
            if (descriptor) Object.defineProperty(globalThis, name, descriptor);
            else delete globalThis[name];
        }
    };
}

/** The `EventSource` the panel opens, with the server end under test control. */
class StubEventSource {
    static instances = [];

    constructor(url) {
        this.url = url;
        this.readyState = 0;
        this.closed = 0;
        this.onopen = null;
        this.onerror = null;
        this.handlers = new Map();
        StubEventSource.instances.push(this);
    }

    addEventListener(type, handler) {
        const list = this.handlers.get(type) ?? [];
        list.push(handler);
        this.handlers.set(type, list);
    }

    close() {
        this.closed += 1;
        this.readyState = 2;
    }

    /** Delivers a named event, JSON-encoded exactly as the server sends it. */
    emit(type, data) {
        const list = this.handlers.get(type) ?? [];
        if (list.length === 0) throw new Error(`the panel is not listening for the '${type}' event`);
        for (const handler of list) handler({ type, data: JSON.stringify(data) });
    }
}

class StubIntersectionObserver {
    static instances = [];

    constructor(callback, options) {
        this.callback = callback;
        this.options = options;
        this.targets = [];
        this.disconnected = 0;
        StubIntersectionObserver.instances.push(this);
    }

    observe(target) {
        this.targets.push(target);
    }

    unobserve(target) {
        this.targets = this.targets.filter((element) => element !== target);
    }

    disconnect() {
        this.disconnected += 1;
    }

    /** Reports the given elements as visible, as a scroll would. */
    enter(...elements) {
        this.callback(elements.map((target) => ({ target, isIntersecting: true })));
    }
}

const RESPONSE = Symbol("stub response");

/**
 * A route answering with a status other than 200.
 *
 * Wrapped rather than sniffed: a viewer state object carries its own `status`
 * field ("ready", "opening", "error"), so reading `status` off a plain body to
 * decide the HTTP code would misread every state push as an HTTP status.
 */
export function respondWith(status, body = {}) {
    return { [RESPONSE]: true, status, body };
}

/** Cache-busts the import so each test gets a fresh copy of app.js's state. */
let instances = 0;

/**
 * Imports the committed `app.js` against a stub DOM and a stub server.
 *
 * `state` is what `GET /api/state` answers. `routes` overrides any path, keyed
 * by pathname, with a body, a `({ pathname, searchParams }) => body` function,
 * or a `respondWith(...)`; anything not named answers `{}` with 200.
 * `pdfDocuments` maps a PDF URL to a `stub-pdfjs.mjs` document spec.
 */
export async function loadApp({ state = null, routes = {}, pdfDocuments = {}, clipboard = {} } = {}) {
    const markup = await readFile(INDEX_URL, "utf8");
    const { document, byId } = createDocument(elementIdsFromMarkup(markup));

    const calls = [];
    const clipboardWrites = [];

    const fetchStub = async (path, options = {}) => {
        const { pathname, searchParams } = new URL(path, "http://viewer.test");
        calls.push({
            path,
            pathname,
            query: Object.fromEntries(searchParams),
            method: options.method ?? "GET",
            body: options.body ? JSON.parse(options.body) : null,
        });

        const route = Object.hasOwn(routes, pathname)
            ? routes[pathname]
            : pathname === "/api/state"
              ? state
              : {};
        const answer = typeof route === "function" ? route({ pathname, searchParams }) : route;
        const wrapped = answer !== null && typeof answer === "object" && answer[RESPONSE];
        const status = wrapped ? answer.status : 200;
        const body = wrapped ? answer.body : (answer ?? {});

        return { ok: status >= 200 && status < 300, status, json: async () => body };
    };

    StubEventSource.instances.length = 0;
    StubIntersectionObserver.instances.length = 0;

    const navigatorStub = {};
    if (!clipboard.absent) {
        navigatorStub.clipboard = {
            writeText: async (text) => {
                clipboardWrites.push(text);
                if (clipboard.fails) throw clipboard.fails;
            },
        };
    }

    const restore = installGlobals({
        document,
        fetch: fetchStub,
        navigator: navigatorStub,
        EventSource: StubEventSource,
        IntersectionObserver: StubIntersectionObserver,
    });

    const pdfjs = await import(stubPdfjsUrl);
    pdfjs.reset();
    for (const [url, spec] of Object.entries(pdfDocuments)) pdfjs.serve(url, spec);

    instances += 1;
    await import(`${APP_URL}?instance=${instances}`);
    await settle();

    return {
        document,
        /** The element carrying this id, asserting it is one the markup ships. */
        el(id) {
            const element = byId.get(id);
            if (!element) throw new Error(`index.html declares no element with id '${id}'`);
            return element;
        },
        calls,
        callsTo: (pathname) => calls.filter((call) => call.pathname === pathname),
        clipboardWrites,
        pdfjs,
        get source() {
            return StubEventSource.instances.at(-1);
        },
        get observers() {
            return StubIntersectionObserver.instances;
        },
        settle,
        restore,
    };
}

/**
 * Imports `pdf-view.mjs` on its own, with the globals it needs.
 *
 * Separate from `loadApp` because the view is worth driving directly: through
 * `app.js` it is reachable only by whatever the panel happens to ask it for.
 */
export async function loadPdfView() {
    const { document } = createDocument([]);
    StubIntersectionObserver.instances.length = 0;

    const restore = installGlobals({ document, IntersectionObserver: StubIntersectionObserver });
    const pdfjs = await import(stubPdfjsUrl);
    pdfjs.reset();
    const { PdfView } = await import(new URL("../../src/ui/pdf-view.mjs", import.meta.url).href);

    return {
        PdfView,
        document,
        pdfjs,
        container: new StubElement("div"),
        get observers() {
            return StubIntersectionObserver.instances;
        },
        settle,
        restore,
    };
}
