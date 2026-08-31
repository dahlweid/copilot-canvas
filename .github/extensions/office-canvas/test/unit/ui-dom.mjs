// A DOM small enough to read in one sitting, and no larger.
//
// This exists so `app.js` and `pdf-view.mjs` can be imported under Node (#76).
// It is **not** a browser and does not try to be one: it parses no HTML, lays
// nothing out, computes no styles and dispatches nothing to a parent. What it
// implements is the set of members the two viewer modules actually touch, which
// is small -- element lookup by id, children, text, classes, `dataset`,
// attributes and listeners.
//
// ## Why not jsdom
//
// The extension folder ships exactly as committed, with no `package.json` and no
// `node_modules` (C2), so a runtime dependency is impossible and a dev-only one
// would change how this repo is built and run. That is the constraint. The
// measurement that says a hand-rolled stand-in is *sufficient* rather than
// merely permitted is in `app.test.mjs` and `pdf-view.test.mjs`: every assertion
// there runs against this file, and each was made to fail by mutating the code
// it covers.
//
// ## Two deliberate refusals
//
// `innerHTML` accepts only the empty string, and `dispatch` throws when nothing
// is listening. Both convert a silent no-op into a failure: the first would
// otherwise let markup be assigned and quietly vanish, the second would let a
// removed handler pass as a test that fired an event at nobody.

/** `Element.classList`, over a set of names. */
class StubClassList {
    #names = new Set();

    add(...names) {
        for (const name of names) if (name) this.#names.add(name);
    }

    remove(...names) {
        for (const name of names) this.#names.delete(name);
    }

    contains(name) {
        return this.#names.has(name);
    }

    toggle(name, force) {
        const on = force === undefined ? !this.#names.has(name) : Boolean(force);
        if (on) this.#names.add(name);
        else this.#names.delete(name);
        return on;
    }

    get value() {
        return [...this.#names].join(" ");
    }

    toString() {
        return this.value;
    }

    /** Used by the `className` setter; not part of the DOM interface. */
    reset(value) {
        this.#names = new Set(String(value ?? "").split(/\s+/u).filter(Boolean));
    }
}

/** `element.style`: a bag of properties that also answers `setProperty`. */
class StubStyle {
    setProperty(name, value) {
        this[name] = String(value);
    }

    getPropertyValue(name) {
        return this[name] ?? "";
    }
}

/** A text node. Present because `setStatus` appends one beside a spinner span. */
export class StubText {
    constructor(data) {
        this.nodeValue = String(data);
    }

    get textContent() {
        return this.nodeValue;
    }

    set textContent(value) {
        this.nodeValue = String(value);
    }
}

export class StubElement {
    constructor(tagName = "div") {
        this.tagName = String(tagName).toUpperCase();
        this.id = "";
        this.hidden = false;
        this.title = "";
        this.value = "";
        this.disabled = false;
        this.type = "";
        this.alt = "";
        this.src = null;
        this.children = [];
        this.classList = new StubClassList();
        this.dataset = {};
        this.style = new StubStyle();
        this.attributes = new Map();
        this.listeners = new Map();
        /** Set by `append`, so `closest` has a chain to walk. */
        this.parentElement = null;
        /** What `getBoundingClientRect` answers. Tests set it when it matters. */
        this.rect = { top: 0, left: 0, width: 0, height: 0 };
        this.scrolledIntoView = 0;
        this.focused = 0;
        /** pdf.js paints into this; the tests only need it to exist. */
        this.width = 0;
        this.height = 0;
    }

    get className() {
        return this.classList.value;
    }

    set className(value) {
        this.classList.reset(value);
    }

    get textContent() {
        return this.children.map((child) => child.textContent).join("");
    }

    set textContent(value) {
        this.children = [new StubText(value)];
    }

    get innerHTML() {
        return this.textContent;
    }

    set innerHTML(value) {
        if (value !== "") {
            throw new Error("the stub DOM parses no HTML; only `innerHTML = \"\"` (a clear) is supported");
        }
        this.children = [];
    }

    append(...nodes) {
        for (const node of nodes) {
            const child = typeof node === "string" ? new StubText(node) : node;
            if (child instanceof StubElement) child.parentElement = this;
            this.children.push(child);
        }
    }

    replaceChildren(...nodes) {
        this.children = [];
        this.append(...nodes);
    }

    setAttribute(name, value) {
        this.attributes.set(name, String(value));
    }

    getAttribute(name) {
        return this.attributes.has(name) ? this.attributes.get(name) : null;
    }

    removeAttribute(name) {
        this.attributes.delete(name);
    }

    addEventListener(type, handler) {
        const list = this.listeners.get(type) ?? [];
        list.push(handler);
        this.listeners.set(type, list);
    }

    removeEventListener(type, handler) {
        const list = (this.listeners.get(type) ?? []).filter((fn) => fn !== handler);
        this.listeners.set(type, list);
    }

    listens(type) {
        return (this.listeners.get(type) ?? []).length > 0;
    }

    /**
     * Fires an event at this element only. There is no bubbling and no default
     * action -- both would be a claim about a browser this file cannot make.
     *
     * Throws when nothing is listening, so a handler that was removed fails here
     * rather than passing as an event delivered to nobody.
     */
    dispatch(type, init = {}) {
        const list = this.listeners.get(type) ?? [];
        if (list.length === 0) {
            throw new Error(`nothing is listening for '${type}' on #${this.id || this.tagName}`);
        }
        let defaultPrevented = false;
        const event = { type, target: this, preventDefault: () => (defaultPrevented = true), ...init };
        for (const handler of list) handler(event);
        return { defaultPrevented };
    }

    /**
     * Walks up the `append` chain for a `.class`, `#id` or tag selector.
     *
     * Returned null unconditionally until #106, which was fine while nothing
     * asked -- but the viewer's fit modes measure `container.closest(".viewer")`,
     * and a `closest` that cannot find anything would have made every fit fall
     * back to its no-ancestor branch and asserted nothing about the branch that
     * actually runs in the panel.
     */
    closest(selector) {
        const matches = (element) => {
            if (selector.startsWith(".")) return element.classList.contains(selector.slice(1));
            if (selector.startsWith("#")) return element.id === selector.slice(1);
            return element.tagName === selector.toUpperCase();
        };
        for (let node = this; node; node = node.parentElement) {
            if (matches(node)) return node;
        }
        return null;
    }

    scrollIntoView() {
        this.scrolledIntoView += 1;
    }

    focus() {
        this.focused += 1;
    }

    getBoundingClientRect() {
        return { ...this.rect };
    }

    /** pdf.js asks a canvas for one; nothing here draws. */
    getContext() {
        this.context ??= { canvas: this };
        return this.context;
    }

    /** Every descendant, for a test that wants to look inside a rendered panel. */
    descendants() {
        const out = [];
        for (const child of this.children) {
            if (!(child instanceof StubElement)) continue;
            out.push(child, ...child.descendants());
        }
        return out;
    }
}

/**
 * Every id the shipped markup declares.
 *
 * Read from `index.html` rather than listed here on purpose. `app.js` looks its
 * elements up by id and gets `null` for one the markup does not have, so a
 * hand-written list would let this harness invent an element the product does
 * not ship and hide exactly the drift `ui-contract.test.mjs` guards.
 */
export function elementIdsFromMarkup(html) {
    return [...String(html).matchAll(/\sid="([^"]+)"/gu)].map((match) => match[1]);
}

/** A `document` carrying one element per id, plus the two factories app.js uses. */
export function createDocument(ids) {
    const byId = new Map();
    const document = {
        title: "",
        createElement: (tag) => new StubElement(tag),
        createTextNode: (text) => new StubText(text),
        getElementById: (id) => byId.get(id) ?? null,
    };
    for (const id of ids) {
        const element = new StubElement("div");
        element.id = id;
        byId.set(id, element);
    }
    return { document, byId };
}
