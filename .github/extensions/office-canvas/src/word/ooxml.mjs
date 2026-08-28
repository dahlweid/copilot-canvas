// A minimal XML reader for WordprocessingML, with no dependencies.
//
// C2 forbids a package.json in an extension folder, so there is no npm parser
// to reach for. This is deliberately not a general XML implementation: it reads
// the Flat OPC package that `Range.WordOpenXML` produces, which Word emits with
// well-formed, non-validating markup and no external entities.
//
// Elements are matched by **local name**, ignoring the prefix. Word always
// writes `w:`, `pkg:` and `mc:`, but a prefix is a serialization detail and
// binding it into every lookup would make the parser fail on markup that is
// still perfectly valid. The cost is that two namespaces sharing a local name
// would be conflated; inside a Word package none of the names we look at do.

export class XmlError extends Error {
    constructor(message) {
        super(message);
        this.name = "XmlError";
        this.code = "invalid_xml";
    }
}

const NAMED_ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };

/** Strips a namespace prefix: `w:pStyle` -> `pStyle`. */
export function localNameOf(name) {
    const colon = name.indexOf(":");
    return colon < 0 ? name : name.slice(colon + 1);
}

export function decodeEntities(text) {
    if (!text.includes("&")) return text;
    return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|[a-zA-Z]+);/g, (whole, body) => {
        if (body[0] === "#") {
            const code = body[1] === "x" || body[1] === "X" ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
            return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
        }
        return NAMED_ENTITIES[body] ?? whole;
    });
}

/**
 * Finds the `>` that closes a tag, skipping any inside quoted attribute values.
 * `indexOf(">")` is wrong here: an attribute may legitimately contain one.
 */
function findTagEnd(source, from) {
    let quote = null;
    for (let i = from; i < source.length; i++) {
        const char = source[i];
        if (quote) {
            if (char === quote) quote = null;
        } else if (char === '"' || char === "'") {
            quote = char;
        } else if (char === ">") {
            return i;
        }
    }
    return -1;
}

const WHITESPACE = new Set([" ", "\t", "\r", "\n"]);

function parseAttributes(source, start, end) {
    const attrs = {};
    let i = start;
    while (i < end) {
        while (i < end && WHITESPACE.has(source[i])) i++;
        if (i >= end) break;

        const nameStart = i;
        while (i < end && !WHITESPACE.has(source[i]) && source[i] !== "=" && source[i] !== "/") i++;
        const name = source.slice(nameStart, i);
        if (!name) {
            i++;
            continue;
        }

        while (i < end && WHITESPACE.has(source[i])) i++;
        if (source[i] !== "=") {
            attrs[name] = "";
            continue;
        }
        i++;
        while (i < end && WHITESPACE.has(source[i])) i++;

        const quote = source[i];
        if (quote === '"' || quote === "'") {
            const close = source.indexOf(quote, i + 1);
            const stop = close < 0 ? end : close;
            attrs[name] = decodeEntities(source.slice(i + 1, stop));
            i = stop + 1;
        } else {
            const valueStart = i;
            while (i < end && !WHITESPACE.has(source[i])) i++;
            attrs[name] = decodeEntities(source.slice(valueStart, i));
        }
    }
    return attrs;
}

function element(name, attrs) {
    return { name, local: localNameOf(name), attrs, children: [] };
}

/**
 * Parses an XML document into a tree. Children are either element objects
 * (`{ name, local, attrs, children }`) or plain strings for text.
 *
 * Returns a synthetic `#document` root, so a document with a leading processing
 * instruction or comment needs no special handling by callers.
 */
export function parseXml(source) {
    if (typeof source !== "string" || source.trim() === "") {
        throw new XmlError("Expected XML text, got nothing.");
    }

    const root = element("#document", {});
    const stack = [root];
    let i = 0;

    while (i < source.length) {
        const open = source.indexOf("<", i);
        if (open < 0) {
            const tail = source.slice(i);
            if (tail) stack[stack.length - 1].children.push(decodeEntities(tail));
            break;
        }
        if (open > i) stack[stack.length - 1].children.push(decodeEntities(source.slice(i, open)));

        if (source.startsWith("<!--", open)) {
            const close = source.indexOf("-->", open + 4);
            i = close < 0 ? source.length : close + 3;
            continue;
        }
        if (source.startsWith("<![CDATA[", open)) {
            const close = source.indexOf("]]>", open + 9);
            const stop = close < 0 ? source.length : close;
            stack[stack.length - 1].children.push(source.slice(open + 9, stop));
            i = close < 0 ? source.length : close + 3;
            continue;
        }
        if (source.startsWith("<?", open)) {
            const close = source.indexOf("?>", open + 2);
            i = close < 0 ? source.length : close + 2;
            continue;
        }
        if (source.startsWith("<!", open)) {
            const close = findTagEnd(source, open + 2);
            i = close < 0 ? source.length : close + 1;
            continue;
        }

        const close = findTagEnd(source, open + 1);
        if (close < 0) throw new XmlError(`Unterminated tag at offset ${open}.`);

        if (source[open + 1] === "/") {
            const name = source.slice(open + 2, close).trim();
            // Word does not emit mismatched tags; unwinding to the matching
            // element rather than popping blindly keeps a malformed fragment
            // from silently reparenting everything after it.
            for (let depth = stack.length - 1; depth > 0; depth--) {
                if (stack[depth].name === name) {
                    stack.length = depth;
                    break;
                }
            }
            i = close + 1;
            continue;
        }

        const selfClosing = source[close - 1] === "/";
        const inner = source.slice(open + 1, selfClosing ? close - 1 : close);
        let nameEnd = 0;
        while (nameEnd < inner.length && !WHITESPACE.has(inner[nameEnd])) nameEnd++;
        const name = inner.slice(0, nameEnd);
        if (!name) throw new XmlError(`Empty tag name at offset ${open}.`);

        const node = element(name, parseAttributes(inner, nameEnd, inner.length));
        stack[stack.length - 1].children.push(node);
        if (!selfClosing) stack.push(node);
        i = close + 1;
    }

    // Truncated markup is reachable in practice: the host streams the package
    // to a file, and a write cut short by a full disk or a killed host leaves a
    // prefix that parses perfectly until it simply stops. Without this check the
    // parser returns a tree that looks whole -- the paragraph being written when
    // the file ended becomes a phantom empty paragraph, and gets a minted
    // address like any other. Refusing is the only honest answer.
    if (stack.length !== 1) {
        const unclosed = stack
            .slice(1)
            .map((node) => node.name)
            .join(" > ");
        throw new XmlError(`Markup ended while still inside <${unclosed}>; the document markup is truncated.`);
    }

    return root;
}

/** Attribute lookup by local name, so `w:val` and `val` both resolve. */
export function attr(node, name) {
    if (!node) return null;
    if (node.attrs[name] !== undefined) return node.attrs[name];
    for (const key of Object.keys(node.attrs)) {
        if (localNameOf(key) === name) return node.attrs[key];
    }
    return null;
}

export function childNamed(node, local) {
    if (!node) return null;
    for (const child of node.children) {
        if (typeof child !== "string" && child.local === local) return child;
    }
    return null;
}

export function childrenNamed(node, local) {
    if (!node) return [];
    return node.children.filter((child) => typeof child !== "string" && child.local === local);
}

/** Depth-first search for the first descendant (or self) with this local name. */
export function findFirst(node, local) {
    if (!node) return null;
    if (node.local === local) return node;
    for (const child of node.children) {
        if (typeof child === "string") continue;
        const hit = findFirst(child, local);
        if (hit) return hit;
    }
    return null;
}

/** Concatenated text of a node and everything under it. */
export function textOf(node) {
    if (!node) return "";
    let out = "";
    for (const child of node.children) {
        out += typeof child === "string" ? child : textOf(child);
    }
    return out;
}
