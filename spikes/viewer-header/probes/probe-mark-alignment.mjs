// Where the Word mark actually sits in the viewer header, measured in a real
// browser (#87).
//
// The defect arrived as a screenshot and the issue offered a cause as an
// explicit hypothesis. A hypothesis about flex cross-sizing cannot be settled by
// reading CSS: the answer depends on what the *grid* does to `.bar-main` before
// the flex algorithm ever runs. So this drives the shipped `src/ui/` in headless
// Chromium over CDP and reads `getBoundingClientRect` off the running product.
//
// ## Why not a stand-in
//
// `test/unit/ui-harness.mjs` can import `app.js` under Node (#76), but that
// harness computes no styles and lays nothing out -- it is explicit about being
// "not a browser". Every number below is a layout result, so a stand-in could
// only restate the assumption being tested.
//
// ## No committed image
//
// `/api/word-icon` is answered with a PNG built in memory here. The repo is
// public and the real mark is Microsoft's (#68), so nothing image-shaped may be
// committed -- `ui-contract.test.mjs` walks the tree asserting exactly that.
//
// ## Run
//
//   node spikes/viewer-header/probes/probe-mark-alignment.mjs
//
// Needs Edge or Chrome on PATH-ish (looked up below) and no Office at all.

import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { deflateSync } from "node:zlib";

const UI_DIR = fileURLToPath(new URL("../../../.github/extensions/office-canvas/src/ui/", import.meta.url));

// --- a PNG, built rather than committed ------------------------------------

const CRC = (() => {
    const table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
        let c = n;
        for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
        table[n] = c;
    }
    return (buf) => {
        let c = -1;
        for (const byte of buf) c = table[(c ^ byte) & 0xff] ^ (c >>> 8);
        return (c ^ -1) >>> 0;
    };
})();

function pngChunk(type, data) {
    const head = Buffer.alloc(8);
    head.writeUInt32BE(data.length, 0);
    head.write(type, 4, "ascii");
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(CRC(Buffer.concat([head.subarray(4), data])), 0);
    return Buffer.concat([head, data, crc]);
}

/** A solid RGBA square, `size` on a side. Stands in for the extracted mark. */
function makePng(size) {
    const stride = size * 4 + 1;
    const raw = Buffer.alloc(stride * size);
    for (let y = 0; y < size; y++) {
        for (let x = 0; x < size; x++) {
            const o = y * stride + 1 + x * 4;
            raw[o] = 0x2b;
            raw[o + 1] = 0x57;
            raw[o + 2] = 0x9a;
            raw[o + 3] = 0xff;
        }
    }
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(size, 0);
    ihdr.writeUInt32BE(size, 4);
    ihdr[8] = 8; // bit depth
    ihdr[9] = 6; // RGBA
    return Buffer.concat([
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
        pngChunk("IHDR", ihdr),
        pngChunk("IDAT", deflateSync(raw)),
        pngChunk("IEND", Buffer.alloc(0)),
    ]);
}

// The extraction produces 32px and `.word-mark` sizes it down to 16 (#68), so
// the probe feeds the same shape the product actually receives.
const ICON = makePng(32);

// --- the viewer's own files, over the routes app.js expects -----------------

const TYPES = new Map([
    [".html", "text/html; charset=utf-8"],
    [".css", "text/css; charset=utf-8"],
    [".js", "text/javascript; charset=utf-8"],
    [".mjs", "text/javascript; charset=utf-8"],
    [".json", "application/json; charset=utf-8"],
]);

const STATE = {
    status: "ready",
    wordVersion: "16.0",
    doc: {
        key: "probe",
        name: "quarterly-report.docx",
        path: "C:\\Users\\probe\\Documents\\quarterly-report.docx",
        pageCount: 12,
        wordCount: 4210,
        title: "Quarterly Report",
        author: "A. Author",
    },
};

/** Serves `src/ui/` plus the handful of API routes the header touches. */
async function serveUi({ icon }) {
    const server = createServer(async (req, res) => {
        const { pathname } = new URL(req.url, "http://localhost");

        if (pathname === "/api/word-icon") {
            if (!icon) {
                res.writeHead(404, { "content-type": "application/json" });
                res.end(JSON.stringify({ error: { code: "not_found" } }));
                return;
            }
            res.writeHead(200, { "content-type": "image/png", "cache-control": "no-store" });
            res.end(icon);
            return;
        }
        if (pathname === "/api/state") {
            res.writeHead(200, { "content-type": "application/json" });
            res.end(JSON.stringify(STATE));
            return;
        }
        if (pathname === "/events") {
            // Held open and silent. The header is drawn from `/api/state`; this
            // only stops `EventSource` from reconnecting in a loop.
            res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache" });
            res.write(": open\n\n");
            return;
        }

        const rel = pathname === "/" ? "index.html" : pathname.replace(/^\//, "");
        const file = path.join(UI_DIR, rel);
        if (!file.startsWith(UI_DIR) || !existsSync(file)) {
            res.writeHead(404).end();
            return;
        }
        res.writeHead(200, { "content-type": TYPES.get(path.extname(file)) ?? "application/octet-stream" });
        res.end(await readFile(file));
    });

    await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
    const { port } = server.address();
    return {
        base: `http://127.0.0.1:${port}/`,
        // `closeAllConnections` first, always. `/events` is deliberately held
        // open, and `server.close()` alone waits for every live socket -- so the
        // probe hangs at teardown having already printed its table, which reads
        // exactly like a measurement that failed. Measured here before it was
        // fixed.
        close: () =>
            new Promise((resolve) => {
                server.closeAllConnections();
                server.close(resolve);
            }),
    };
}

// --- headless Chromium, over CDP -------------------------------------------

const BROWSERS = [
    path.join(process.env.ProgramFiles ?? "", "Google\\Chrome\\Application\\chrome.exe"),
    path.join(process.env["ProgramFiles(x86)"] ?? "", "Google\\Chrome\\Application\\chrome.exe"),
    path.join(process.env.LOCALAPPDATA ?? "", "Google\\Chrome\\Application\\chrome.exe"),
    path.join(process.env.ProgramFiles ?? "", "Microsoft\\Edge\\Application\\msedge.exe"),
    path.join(process.env["ProgramFiles(x86)"] ?? "", "Microsoft\\Edge\\Application\\msedge.exe"),
];

function findBrowser() {
    const found = BROWSERS.find((p) => p && existsSync(p));
    if (!found) throw new Error(`no Chromium found; looked in:\n  ${BROWSERS.join("\n  ")}`);
    return found;
}

/**
 * A CDP client over Node's global `WebSocket`.
 *
 * Flat sessions (`flatten: true`), so a page command is an ordinary message
 * carrying a `sessionId` rather than a string nested inside `Target.sendMessage`.
 */
class Cdp {
    #ws;
    #next = 0;
    #pending = new Map();

    static async connect(url) {
        const client = new Cdp();
        client.#ws = new WebSocket(url);
        client.#ws.addEventListener("message", (event) => {
            const message = JSON.parse(event.data);
            if (message.id === undefined) return;
            const entry = client.#pending.get(message.id);
            if (!entry) return;
            client.#pending.delete(message.id);
            if (message.error) entry.reject(new Error(`${message.error.message} (${message.error.code})`));
            else entry.resolve(message.result);
        });
        await new Promise((resolve, reject) => {
            client.#ws.addEventListener("open", resolve, { once: true });
            client.#ws.addEventListener("error", () => reject(new Error(`cannot reach ${url}`)), { once: true });
        });
        return client;
    }

    send(method, params = {}, sessionId) {
        const id = ++this.#next;
        this.#ws.send(JSON.stringify(sessionId ? { id, method, params, sessionId } : { id, method, params }));
        return new Promise((resolve, reject) => this.#pending.set(id, { resolve, reject }));
    }

    close() {
        this.#ws.close();
    }
}

async function launchBrowser() {
    const exe = findBrowser();
    const profile = await mkdtemp(path.join(tmpdir(), "viewer-header-"));

    // Discrete argv elements, no shell. A path interpolated into a command
    // string is the parser trap this repo has been bitten by twice; `%` and `&`
    // in a temp path are ordinary characters to `spawn` and are not to `cmd`.
    const child = spawn(
        exe,
        [
            "--headless=new",
            "--disable-gpu",
            "--hide-scrollbars",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            "--force-device-scale-factor=1",
            "--window-size=900,700",
            "--remote-allow-origins=*",
            "--remote-debugging-port=0",
            `--user-data-dir=${profile}`,
            "about:blank",
        ],
        { stdio: ["ignore", "ignore", "pipe"] },
    );

    const endpoint = await new Promise((resolve, reject) => {
        let buffer = "";
        const timer = setTimeout(() => reject(new Error(`browser printed no DevTools endpoint:\n${buffer}`)), 30_000);
        child.stderr.on("data", (chunk) => {
            buffer += chunk;
            const match = buffer.match(/ws:\/\/\S+/);
            if (match) {
                clearTimeout(timer);
                resolve(match[0]);
            }
        });
        child.on("exit", (code) => reject(new Error(`browser exited with ${code}:\n${buffer}`)));
    });

    const cdp = await Cdp.connect(endpoint);
    const { targetId } = await cdp.send("Target.createTarget", { url: "about:blank" });
    const { sessionId } = await cdp.send("Target.attachToTarget", { targetId, flatten: true });
    await cdp.send("Page.enable", {}, sessionId);
    await cdp.send("Runtime.enable", {}, sessionId);

    return {
        async evaluate(fn, args = {}) {
            const { result, exceptionDetails } = await cdp.send(
                "Runtime.evaluate",
                {
                    expression: `(${fn})(${JSON.stringify(args)})`,
                    returnByValue: true,
                    awaitPromise: true,
                },
                sessionId,
            );
            if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? "evaluation failed");
            return result.value;
        },
        async goto(url) {
            await cdp.send("Page.navigate", { url }, sessionId);
        },
        async close() {
            // Only ever the browser this probe started. Nothing here goes
            // looking for a Chromium it did not launch.
            try {
                await cdp.send("Browser.close");
            } catch {
                child.kill();
            }
            cdp.close();
            await new Promise((r) => child.on("exit", r));
            await rm(profile, { recursive: true, force: true }).catch(() => {});
        },
    };
}

// --- what is measured, in the page -----------------------------------------

/**
 * Runs in the browser. Reports the mark's box, the name's box, and the two
 * reference lines an eye actually compares the mark against.
 *
 * The name's baseline is *measured*, not derived from font metrics: a zero-height
 * inline-block aligned to the baseline has its top, bottom and the baseline at
 * the same y, so its rect reads the baseline straight off the layout. Cap height
 * comes from the canvas text metrics for the element's own computed font, so the
 * "optical centre" below is the middle of a capital letter rather than the middle
 * of a line box.
 */
const MEASURE = function measure() {
    const round = (n) => Math.round(n * 100) / 100;
    const mark = document.getElementById("docNameMark");
    const name = document.getElementById("docName");
    const main = document.querySelector(".bar-main");
    const actions = document.querySelector(".bar-actions");

    const probe = document.createElement("span");
    probe.style.cssText = "display:inline-block;width:0;height:0;vertical-align:baseline";
    name.appendChild(probe);
    const baseline = probe.getBoundingClientRect().top;
    probe.remove();

    const style = getComputedStyle(name);
    const ctx = document.createElement("canvas").getContext("2d");
    ctx.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const capHeight = ctx.measureText("H").actualBoundingBoxAscent;

    const markRect = mark.getBoundingClientRect();
    const nameRect = name.getBoundingClientRect();
    const opticalCentre = baseline - capHeight / 2;

    return {
        markHidden: mark.hidden,
        markTop: round(markRect.top),
        markHeight: round(markRect.height),
        markCentre: round(markRect.top + markRect.height / 2),
        nameTop: round(nameRect.top),
        nameHeight: round(nameRect.height),
        nameBoxCentre: round(nameRect.top + nameRect.height / 2),
        nameBaseline: round(baseline),
        capHeight: round(capHeight),
        opticalCentre: round(opticalCentre),
        // Positive means the mark sits *below* the middle of the filename's
        // capitals -- which is how the screenshot reads.
        offsetFromOptical: round(markRect.top + markRect.height / 2 - opticalCentre),
        offsetFromNameBox: round(markRect.top + markRect.height / 2 - (nameRect.top + nameRect.height / 2)),
        // The hypothesis under test lives here: which box does the flex line
        // take its cross size from?
        barMainHeight: round(main.getBoundingClientRect().height),
        barActionsHeight: round(actions.getBoundingClientRect().height),
        markAlignSelf: getComputedStyle(mark).alignSelf,
        barMainAlignSelf: getComputedStyle(main).alignSelf,
    };
};

/** The Open in Word button, as a browser would draw it. Used for defect 2. */
const MEASURE_BUTTON = function measureButton() {
    const round = (n) => Math.round(n * 100) / 100;
    const button = document.getElementById("openInWord");
    const rect = button.getBoundingClientRect();
    const visible = [...button.children]
        .filter((child) => child.getClientRects().length > 0)
        .map((child) => `${child.tagName.toLowerCase()}#${child.id || "-"}`);
    return {
        width: round(rect.width),
        height: round(rect.height),
        text: button.textContent.trim(),
        visibleChildren: visible.join(","),
        childCount: button.children.length,
    };
};

/**
 * Does `element.hidden = true` reach the *style* of an `<svg>`?
 *
 * `hidden` is an IDL attribute of `HTMLElement`. `SVGElement` does not inherit
 * from it, so the assignment may land as an ordinary JavaScript property that
 * reflects to no content attribute and matches no `[hidden]` selector. That is a
 * platform question, not a question about this repo's markup, so it is measured
 * on elements built here -- the probe keeps answering it after the viewer stops
 * containing a glyph to swap.
 */
const MEASURE_HIDDEN_REFLECTION = function measureHiddenReflection() {
    const host = document.createElement("div");
    host.style.cssText = "position:absolute;left:-9999px";
    document.body.appendChild(host);

    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("width", "16");
    svg.setAttribute("height", "16");
    const img = document.createElement("img");
    img.width = 16;
    img.height = 16;
    host.append(svg, img);

    const report = (el) => {
        el.hidden = true;
        const out = {
            element: el.tagName.toLowerCase(),
            interface: el.constructor.name,
            "hidden in prototype": Object.getPrototypeOf(el) === null ? "?" : String(hiddenIsAccessor(el)),
            "attribute set": String(el.hasAttribute("hidden")),
            display: getComputedStyle(el).display,
            "client rects": el.getClientRects().length,
        };
        el.hidden = false;
        return out;
    };
    function hiddenIsAccessor(el) {
        for (let p = Object.getPrototypeOf(el); p; p = Object.getPrototypeOf(p)) {
            const d = Object.getOwnPropertyDescriptor(p, "hidden");
            if (d) return true;
        }
        return false;
    }

    const rows = [report(svg), report(img)];
    host.remove();
    return rows;
};

const READY = function ready() {
    const mark = document.getElementById("docNameMark");
    const bar = document.getElementById("bar");
    return { barShown: !bar.hidden, markShown: !mark.hidden, complete: document.readyState === "complete" };
};

async function waitFor(browser, predicate, what, { timeoutMs = 15_000 } = {}) {
    const deadline = Date.now() + timeoutMs;
    let last;
    while (Date.now() < deadline) {
        last = await browser.evaluate(READY).catch(() => null);
        if (last && predicate(last)) return last;
        await new Promise((r) => setTimeout(r, 100));
    }
    throw new Error(`timed out waiting for ${what}: ${JSON.stringify(last)}`);
}

// --- variants ---------------------------------------------------------------

/**
 * Candidate rules, injected over the shipped stylesheet.
 *
 * `!important` only so a one-line override outranks the shipped rule; what each
 * candidate *means* is the single property named, not the flag.
 */
const VARIANTS = [
    { name: "as shipped (align-self: center)", css: "" },
    // The state #87 was reported against, kept as a variant rather than as a
    // number copied out of a run against 6dc3369. Once the fix lands, "as
    // shipped" *is* the fixed tree, and a before/after claim whose "before" half
    // can no longer be produced is an assertion. Undoing the one property the
    // fix adds reproduces it on demand: `.bar-main` stretches to the grid row
    // again, which is the whole mechanism.
    { name: "before #87 (bar-main stretched)", css: ".bar-main { align-self: stretch !important; }" },
    { name: "mark: align-self baseline", css: ".word-mark { align-self: baseline !important; }" },
    { name: "mark: align-self first baseline", css: ".word-mark { align-self: first baseline !important; }" },
    { name: "bar-main: align-self center (mark stays centred)", css: ".bar-main { align-self: center !important; }" },
    {
        name: "bar-main centred + mark baseline",
        css: ".bar-main { align-self: center !important; } .word-mark { align-self: baseline !important; }",
    },
];

/** Theme sizes the bar is expected to survive: the tokens app.css falls back on. */
const THEMES = [
    { name: "default (14/20)", css: "" },
    { name: "large (16/24)", css: ":root { --text-body-medium: 16px; --leading-body-medium: 24px; }" },
    { name: "meta absent", css: "", empty: true },
];

const APPLY = function apply({ css, empty }) {
    document.getElementById("probe-style")?.remove();
    if (css) {
        const style = document.createElement("style");
        style.id = "probe-style";
        style.textContent = css;
        document.head.appendChild(style);
    }
    const meta = document.getElementById("docMeta");
    if (empty) {
        meta.dataset.saved ??= meta.textContent;
        meta.textContent = "";
    } else if (meta.dataset.saved !== undefined) {
        meta.textContent = meta.dataset.saved;
    }
    return true;
};

function table(rows, columns) {
    const widths = columns.map((c) => Math.max(c.length, ...rows.map((r) => String(r[c] ?? "").length)));
    const line = (cells) => "| " + cells.map((cell, i) => String(cell ?? "").padEnd(widths[i])).join(" | ") + " |";
    return [line(columns), "|" + widths.map((w) => "-".repeat(w + 2)).join("|") + "|", ...rows.map((r) => line(columns.map((c) => r[c])))].join("\n");
}

// --- run --------------------------------------------------------------------

const server = await serveUi({ icon: ICON });
const browser = await launchBrowser();

try {
    await browser.goto(server.base);
    await waitFor(browser, (s) => s.barShown && s.markShown, "the bar and the mark to be drawn");

    console.log(`browser:  ${findBrowser()}`);
    console.log(`viewer:   ${server.base}\n`);

    const rows = [];
    for (const theme of THEMES) {
        for (const variant of VARIANTS) {
            await browser.evaluate(APPLY, { css: `${theme.css}\n${variant.css}`, empty: Boolean(theme.empty) });
            const m = await browser.evaluate(MEASURE);
            rows.push({
                theme: theme.name,
                variant: variant.name,
                "line h": m.barMainHeight,
                "actions h": m.barActionsHeight,
                "mark centre": m.markCentre,
                "optical centre": m.opticalCentre,
                "offset": m.offsetFromOptical,
            });
        }
    }
    await browser.evaluate(APPLY, { css: "", empty: false });

    console.log("## Mark against the filename's optical centre (px; + = mark sits low)\n");
    console.log(table(rows, ["theme", "variant", "line h", "actions h", "mark centre", "optical centre", "offset"]));

    const shipped = await browser.evaluate(MEASURE);
    console.log("\n## As shipped, in full\n");
    console.log(JSON.stringify(shipped, null, 2));

    console.log("\n## Does `element.hidden = true` hide an <svg>?\n");
    const reflection = await browser.evaluate(MEASURE_HIDDEN_REFLECTION);
    console.log(
        table(reflection, ["element", "interface", "hidden in prototype", "attribute set", "display", "client rects"]),
    );

    // Defect 2: the button must not depend on this machine having Word.
    const withWord = await browser.evaluate(MEASURE_BUTTON);
    await server.close();
    const without = await serveUi({ icon: null });
    await browser.goto(without.base);
    await waitFor(browser, (s) => s.barShown, "the bar to be drawn without an icon route");
    const withoutWord = await browser.evaluate(MEASURE_BUTTON);
    await without.close();

    console.log("\n## Open in Word, with and without an extractable mark\n");
    console.log(table([{ machine: "Word installed", ...withWord }, { machine: "no Word (404)", ...withoutWord }], [
        "machine",
        "width",
        "height",
        "text",
        "visibleChildren",
        "childCount",
    ]));
    const same = JSON.stringify(withWord) === JSON.stringify(withoutWord);
    console.log(`\nidentical: ${same ? "yes" : "NO -- the button depends on the machine"}`);
} finally {
    await browser.close();
    await server.close().catch(() => {});
}
