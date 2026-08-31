// What zooming actually does to the viewer, measured in a real browser (#106).
//
// Three questions this probe exists to answer, none of which a unit suite can:
//
//  1. **Is the left edge of a zoomed page reachable?** `.viewer` scrolls and
//     `.pages` centres its children. A centred flex item wider than its line
//     overflows at *both* ends, and the start-edge overflow is not scrollable --
//     so the left margin of a zoomed page can sit at a negative x that
//     `scrollLeft = 0` already is. This is a layout result. Reading the CSS can
//     produce a hypothesis about it and nothing more.
//
//  2. **Do the bitmaps repaint, or only the boxes?** `#renderPage` early-returns
//     on `page.rendered`. Miss the reset and `<canvas>.width` -- the bitmap's own
//     pixel width, not its CSS width -- stays at the scale of the first paint
//     while the box around it grows. A screenshot of a fresh load looks correct.
//
//  3. **Does the text layer survive?** pdf.js's own stylesheet sizes every span
//     with `round(down, var(--total-scale-factor) * Npx, ...)`. If that custom
//     property stops being set the declaration is invalid, the layer collapses,
//     and text selection and the change overlay go with it.
//
// ## Why the shipped files
//
// `src/ui/` as committed, and the real vendored pdf.js -- worker reassembled from
// its committed parts by the product's own `vendor-assets.mjs`. A stand-in would
// only restate the assumption under test; `test/unit/ui-harness.mjs` is explicit
// about computing no styles and laying nothing out.
//
// ## Why the PDF is built here
//
// Nothing document-shaped is committed to this repo, so the probe writes a
// minimal three-page PDF into memory: uncompressed, Helvetica, one line of text
// per page. pdf.js parses it with the same code path it uses on a Word export,
// and it gives the text layer something to lay out.
//
// ## Run
//
//   node spikes/viewer-zoom/probes/probe-zoom-overflow.mjs
//
// Needs Edge or Chrome (looked up below) and no Office at all.

import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { mkdtemp, readFile, rm } from "node:fs/promises";
import { existsSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { joinWorker } from "../../../.github/extensions/office-canvas/src/vendor-assets.mjs";

const UI_DIR = fileURLToPath(new URL("../../../.github/extensions/office-canvas/src/ui/", import.meta.url));
const VENDOR_DIR = path.join(UI_DIR, "vendor");

const PAGE_WIDTH = 612;
const PAGE_HEIGHT = 792;
const PAGE_COUNT = 3;

// --- a PDF, built rather than committed -------------------------------------

/**
 * A minimal uncompressed PDF, assembled object by object with a real xref.
 *
 * Byte offsets are taken as the buffer is built rather than computed afterwards:
 * an xref that disagrees with the file is the one way a hand-built PDF fails, and
 * it fails as "pdf.js could not open it", which reads like a probe that broke.
 */
function buildPdf() {
    const objects = [];
    const add = (body) => {
        objects.push(body);
        return `${objects.length} 0 R`;
    };

    const pagesRef = "2 0 R";
    add("<< /Type /Catalog /Pages 2 0 R >>");
    objects.push(null); // 2: the page tree, filled in once the kids exist.
    const fontRef = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

    const kids = [];
    for (let number = 1; number <= PAGE_COUNT; number++) {
        const text = `Page ${number} of a probe document. The quick brown fox jumps over the lazy dog.`;
        const stream = `BT /F1 18 Tf 72 ${PAGE_HEIGHT - 96} Td (${text}) Tj ET`;
        const contents = add(`<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`);
        kids.push(
            add(
                `<< /Type /Page /Parent ${pagesRef} /MediaBox [0 0 ${PAGE_WIDTH} ${PAGE_HEIGHT}] ` +
                    `/Resources << /Font << /F1 ${fontRef} >> >> /Contents ${contents} >>`,
            ),
        );
    }
    objects[1] = `<< /Type /Pages /Kids [${kids.join(" ")}] /Count ${kids.length} >>`;

    const chunks = [Buffer.from("%PDF-1.4\n")];
    let offset = chunks[0].length;
    const offsets = [];
    for (const [index, body] of objects.entries()) {
        const chunk = Buffer.from(`${index + 1} 0 obj\n${body}\nendobj\n`);
        offsets.push(offset);
        offset += chunk.length;
        chunks.push(chunk);
    }

    const xref = [`xref\n0 ${objects.length + 1}\n`, "0000000000 65535 f \n"];
    for (const at of offsets) xref.push(`${String(at).padStart(10, "0")} 00000 n \n`);
    chunks.push(
        Buffer.from(
            `${xref.join("")}trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${offset}\n%%EOF\n`,
        ),
    );
    return Buffer.concat(chunks);
}

// --- the server -------------------------------------------------------------

const TYPES = new Map([
    [".html", "text/html; charset=utf-8"],
    [".css", "text/css; charset=utf-8"],
    [".js", "text/javascript; charset=utf-8"],
    [".mjs", "text/javascript; charset=utf-8"],
]);

const STATE = {
    status: "ready",
    error: null,
    lastPage: 1,
    wordVersion: "16.0.18025",
    pdfUrl: "/pdf/probe.pdf",
    change: null,
    doc: {
        key: "probe",
        name: "probe.docx",
        path: "C:\\docs\\probe.docx",
        pageCount: PAGE_COUNT,
        wordCount: 42,
        title: "Zoom Probe",
        author: "A. Author",
    },
};

/** Serves `src/ui/`, the real vendored pdf.js, and the built PDF. */
async function serveUi({ pdf, worker }) {
    const server = createServer(async (req, res) => {
        const { pathname } = new URL(req.url, "http://localhost");

        if (pathname === "/pdf/probe.pdf") {
            res.writeHead(200, { "content-type": "application/pdf", "cache-control": "no-store" });
            return res.end(pdf);
        }
        if (pathname === "/vendor/pdf.worker.min.mjs") {
            // Reassembled by the product's own module, from the parts as
            // committed. Serving a part alone would hand pdf.js a fragment of a
            // program, which is why the real server does the same thing.
            res.writeHead(200, { "content-type": "text/javascript; charset=utf-8" });
            return res.end(worker);
        }
        if (pathname === "/api/state") {
            res.writeHead(200, { "content-type": "application/json" });
            return res.end(JSON.stringify(STATE));
        }
        if (pathname === "/events") {
            // Held open and silent, so `EventSource` does not reconnect in a loop.
            res.writeHead(200, { "content-type": "text/event-stream", "cache-control": "no-cache" });
            return res.write(": open\n\n");
        }
        if (pathname.startsWith("/api/")) {
            res.writeHead(200, { "content-type": "application/json" });
            return res.end("{}");
        }

        const root = pathname.startsWith("/vendor/") ? VENDOR_DIR : UI_DIR;
        const rel = pathname === "/" ? "index.html" : pathname.replace(/^\/(vendor\/)?/, "");
        const file = path.join(root, rel);
        if (!file.startsWith(root) || !existsSync(file)) {
            return res.writeHead(404).end();
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
        // probe would hang at teardown having already printed its table, which
        // reads exactly like a measurement that failed.
        close: () =>
            new Promise((resolve) => {
                server.closeAllConnections();
                server.close(resolve);
            }),
    };
}

// --- headless Chromium, over CDP --------------------------------------------

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

/** A CDP client over Node's global `WebSocket`, with flat sessions. */
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

async function launchBrowser({ width, height }) {
    const exe = findBrowser();
    const profile = await mkdtemp(path.join(tmpdir(), "viewer-zoom-"));

    // Discrete argv elements, no shell. A path interpolated into a command
    // string is the parser trap this repo has been bitten by twice; `%` and `&`
    // in a temp path are ordinary characters to `spawn` and are not to `cmd`.
    const child = spawn(
        exe,
        [
            "--headless=new",
            "--disable-gpu",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-extensions",
            "--force-device-scale-factor=1",
            `--window-size=${width},${height}`,
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
    // Not `Target.createTarget`'s width/height: those are a *window* position and
    // are refused for a target created in an existing window ("Target position
    // can only be set for new windows"). The emulation override is also what the
    // resize steps below use, so the panel is sized one way throughout.
    const setMetrics = (size) =>
        cdp.send(
            "Emulation.setDeviceMetricsOverride",
            { width: size.width, height: size.height, deviceScaleFactor: 1, mobile: false },
            sessionId,
        );
    await setMetrics({ width, height });

    return {
        async evaluate(fn, args = {}) {
            const { result, exceptionDetails } = await cdp.send(
                "Runtime.evaluate",
                { expression: `(${fn})(${JSON.stringify(args)})`, returnByValue: true, awaitPromise: true },
                sessionId,
            );
            if (exceptionDetails) throw new Error(exceptionDetails.exception?.description ?? "evaluation failed");
            return result.value;
        },
        /** Resizes the panel the way dragging its edge does. */
        async resize(next) {
            await setMetrics(next);
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

// --- what is measured, in the page ------------------------------------------

/** Waits until the viewer has laid pages out and painted the first one. */
const READY = function ready() {
    const deadline = Date.now() + 30_000;
    return new Promise((resolve, reject) => {
        const tick = () => {
            const canvas = document.querySelector(".pages .page .page-canvas");
            const layer = document.querySelector(".pages .page .textLayer");
            if (canvas && canvas.width > 0 && layer && layer.querySelectorAll("span").length > 0) {
                return resolve(true);
            }
            if (Date.now() > deadline) return reject(new Error("the viewer never painted a page"));
            setTimeout(tick, 100);
        };
        tick();
    });
};

/**
 * Runs in the browser. One reading of everything the three questions need.
 *
 * `leftReach` is the whole of question 1: the page's left edge in the viewer's
 * own coordinates once the viewer has been scrolled as far left as it goes. A
 * negative number is content that exists and cannot be reached.
 */
const MEASURE = function measure() {
    const round = (n) => Math.round(n * 100) / 100;
    const viewer = document.querySelector(".viewer");
    const pages = document.querySelector(".pages");
    const page = pages.children[0];
    const canvas = page.querySelector(".page-canvas");
    const layer = page.querySelector(".textLayer");

    viewer.scrollLeft = -99_999;
    const scrollLeft = viewer.scrollLeft;
    const viewerBox = viewer.getBoundingClientRect();
    const pageBox = page.getBoundingClientRect();
    // The *content* box's left edge. `getBoundingClientRect` returns the border
    // box, whose right edge sits outside the vertical scrollbar -- measuring a
    // right margin against it reports every centred page as 15px off-centre,
    // which is a scrollbar and not a layout fault. Measured once the naive way
    // and it made all six narrow rows look wrong by exactly the same 15px.
    const contentLeft = viewerBox.left + viewer.clientLeft;

    return {
        scale: Number(getComputedStyle(pages).getPropertyValue("--total-scale-factor")),
        viewerWidth: round(viewerBox.width),
        // What a fit actually measures: the border box less its scrollbar.
        viewerClientWidth: viewer.clientWidth,
        pageCssWidth: round(pageBox.width),
        // The bitmap's own pixel width, which is what a re-scale has to change.
        canvasBitmapWidth: canvas.width,
        scrollLeftAtStart: round(scrollLeft),
        scrollWidth: pages.scrollWidth,
        leftReach: round(pageBox.left - contentLeft),
        // The same margin off the border box, which is how this probe measured
        // before the correction above. Recorded so the figures quoted from the
        // earlier runs -- the -223.5 the CSS comment cites -- can be shown to be
        // untouched by the change of instrument, rather than argued to be.
        leftReachBorderBox: round(pageBox.left - viewerBox.left),
        // The other margin. Equal to `leftReach` is centred; a `flex-start` fix
        // for the overflow would cure `leftReach` and leave this one carrying
        // the whole difference.
        rightReach: round(contentLeft + viewer.clientWidth - pageBox.right),
        textSpans: layer.querySelectorAll("span").length,
        // Zero would mean the text layer collapsed, which is what a missing
        // `--total-scale-factor` does to it.
        textLayerWidth: round(layer.getBoundingClientRect().width),
    };
};

const press = (id) => `function () { document.getElementById(${JSON.stringify(id)}).click(); }`;
const settle = () => new Promise((r) => setTimeout(r, 1200));

/**
 * Puts `.pages` back to what it was before this change, so the trap can be
 * measured rather than argued about: `align-items: center` and no `min-width`.
 *
 * A flex line only grows to `max-content` because something asks it to. Without
 * that, a child wider than the line is centred *on the line* -- it hangs off
 * both ends, and the overflow at the start end is not in the scroll range.
 */
const NAIVE_CSS = function naive(args) {
    document.querySelector(".pages").style.minWidth = args.on ? "auto" : "";
};

// --- the run ----------------------------------------------------------------

const PANEL = { width: 900, height: 800 };
const NARROW = { width: 520, height: 800 };

async function main() {
    const worker = (await joinWorker()).buffer;
    const server = await serveUi({ pdf: buildPdf(), worker });
    const browser = await launchBrowser(PANEL);
    const rows = [];
    let control = null;
    const step = async (label) => rows.push([label, await browser.evaluate(MEASURE)]);

    try {
        await browser.goto(server.base);
        await browser.evaluate(READY);
        await step("open (fit-width)");

        await browser.evaluate(press("zoomIn"));
        await settle();
        await step("zoom in x1");

        await browser.evaluate(press("zoomIn"));
        await settle();
        await step("zoom in x2");

        await browser.evaluate(press("zoomOut"));
        await settle();
        await step("zoom out x1");

        await browser.evaluate(press("fitHeight"));
        await settle();
        await step("fit-height");

        await browser.evaluate(press("fitWidth"));
        await settle();
        await step("fit-width");

        // Below the container width, which is the case `min-width: max-content`
        // must *not* disturb. Two presses off a fit is 0.64 of it, so the page
        // is comfortably narrower than the viewer and there is a real margin on
        // both sides to compare.
        await browser.evaluate(press("zoomOut"));
        await settle();
        await browser.evaluate(press("zoomOut"));
        await settle();
        await step("zoom out below fit");

        await browser.evaluate(press("fitWidth"));
        await settle();
        await browser.resize(NARROW);
        await settle();
        await step(`fitted, resize to ${NARROW.width}`);

        await browser.evaluate(press("zoomIn"));
        await settle();
        const zoomedScale = rows.at(-1) && (await browser.evaluate(MEASURE)).scale;
        await browser.resize(PANEL);
        await settle();
        await step(`zoomed(${zoomedScale}), resize to ${PANEL.width}`);

        // The control. Same document, same scale, `.pages` put back to the
        // centring-only rule -- so the two numbers differ in one property.
        await browser.evaluate(press("fitWidth"));
        await settle();
        await browser.evaluate(press("zoomIn"));
        await browser.evaluate(press("zoomIn"));
        await settle();
        const fixed = await browser.evaluate(MEASURE);
        await browser.evaluate(NAIVE_CSS, { on: true });
        await settle();
        control = { fixed, naive: await browser.evaluate(MEASURE) };
        await browser.evaluate(NAIVE_CSS, { on: false });
    } finally {
        await browser.close();
        await server.close();
    }

    const columns = [
        ["step", (row) => row[0]],
        ["scale", (row) => row[1].scale.toFixed(3)],
        ["viewer", (row) => row[1].viewerClientWidth],
        ["page css", (row) => row[1].pageCssWidth],
        ["canvas px", (row) => row[1].canvasBitmapWidth],
        ["scrollW", (row) => row[1].scrollWidth],
        ["scrollLeft", (row) => row[1].scrollLeftAtStart],
        ["leftReach", (row) => row[1].leftReach],
        ["rightReach", (row) => row[1].rightReach],
        ["spans", (row) => row[1].textSpans],
        ["layer w", (row) => row[1].textLayerWidth],
    ];
    const table = [columns.map((c) => c[0]), ...rows.map((row) => columns.map((c) => String(c[1](row))))];
    const widths = table[0].map((_, i) => Math.max(...table.map((r) => r[i].length)));
    for (const [index, line] of table.entries()) {
        console.log(line.map((cell, i) => cell.padEnd(widths[i])).join("  "));
        if (index === 0) console.log(widths.map((w) => "-".repeat(w)).join("  "));
    }

    console.log(
        `\nthe same page at scale ${control.fixed.scale.toFixed(3)}, differing only in .pages:\n` +
            `  min-width: max-content  leftReach ${control.fixed.leftReach}, ` +
            `scrollWidth ${control.fixed.scrollWidth}, scrollLeft floor ${control.fixed.scrollLeftAtStart}\n` +
            `  centring alone         leftReach ${control.naive.leftReach}, ` +
            `scrollWidth ${control.naive.scrollWidth}, scrollLeft floor ${control.naive.scrollLeftAtStart}`,
    );

    // The questions, answered as pass/fail against the numbers above.
    const at = (label) => rows.find((row) => row[0].startsWith(label));
    const findings = [
        [
            "the left edge of the page is reachable at every scale",
            rows.every(([, m]) => m.leftReach >= -0.5),
            rows.map(([label, m]) => `${label}: ${m.leftReach}`).join(", "),
        ],
        [
            "centring alone puts the left edge out of reach, so the rule is load-bearing",
            control.naive.leftReach < -0.5 && control.fixed.leftReach >= -0.5,
            `${control.naive.leftReach} without it, ${control.fixed.leftReach} with it`,
        ],
        [
            // The other half of question 1, and the failure mode a `flex-start`
            // fix would introduce: it cures the overflow and silently loses
            // centring for every page narrower than the panel, which at a
            // default fit is most of them.
            "a page narrower than the viewer is still centred, not flushed left",
            (() => {
                const narrow = rows.filter(([, m]) => m.pageCssWidth < m.viewerClientWidth - 2);
                return (
                    narrow.length > 0 &&
                    narrow.every(([, m]) => Math.abs(m.leftReach - m.rightReach) <= 1 && m.leftReach > 1)
                );
            })(),
            rows
                .filter(([, m]) => m.pageCssWidth < m.viewerClientWidth - 2)
                .map(([label, m]) => `${label}: ${m.leftReach} left / ${m.rightReach} right`)
                .join(", ") || "no row had a page narrower than its viewer",
        ],
        [
            // The correction to the content box moved every *right* margin by
            // the scrollbar's 15px. It could not have moved a left margin --
            // `clientLeft` is the left border width, and the vertical scrollbar
            // is on the right in LTR -- but "could not have" is a re-reading,
            // so measure it: the two instruments must agree on every row.
            "the content-box correction leaves every leftReach figure unchanged",
            [...rows, ["centring alone", control.naive], ["with the rule", control.fixed]].every(
                ([, m]) => Math.abs(m.leftReach - m.leftReachBorderBox) <= 0.01,
            ),
            [...rows, ["centring alone", control.naive], ["with the rule", control.fixed]]
                .map(([label, m]) => `${label}: ${m.leftReach} vs ${m.leftReachBorderBox}`)
                .join(", "),
        ],
        [
            "the canvas bitmap is repainted at each new scale",
            new Set(rows.map((row) => `${row[1].scale}|${row[1].canvasBitmapWidth}`)).size ===
                new Set(rows.map((row) => row[1].scale)).size,
            rows.map(([label, m]) => `${label}: ${m.scale.toFixed(3)} -> ${m.canvasBitmapWidth}px`).join(", "),
        ],
        [
            "the text layer survives every re-scale",
            rows.every(([, m]) => m.textSpans > 0 && m.textLayerWidth > 0),
            rows.map(([label, m]) => `${label}: ${m.textSpans} spans, ${m.textLayerWidth}px`).join(", "),
        ],
        [
            "a standing fit follows a panel resize",
            (() => {
                const fitted = at("fitted, resize");
                return Boolean(fitted) && Math.abs(fitted[1].pageCssWidth - (fitted[1].viewerClientWidth - 32)) <= 1;
            })(),
            (() => {
                const fitted = at("fitted, resize");
                return fitted
                    ? `viewer ${fitted[1].viewerClientWidth} less a 32px gutter, page ${fitted[1].pageCssWidth}`
                    : "not measured";
            })(),
        ],
        [
            "a hand-picked zoom is not discarded by a panel resize",
            (() => {
                const after = at("zoomed(");
                const chosen = Number(after?.[0].match(/zoomed\(([\d.]+)\)/)?.[1]);
                return Boolean(after) && Math.abs(after[1].scale - chosen) < 0.001;
            })(),
            at("zoomed(")?.[0] ?? "not measured",
        ],
    ];

    console.log("");
    let failed = 0;
    for (const [name, ok, detail] of findings) {
        if (!ok) failed++;
        console.log(`${ok ? "PASS" : "FAIL"}  ${name}\n      ${detail}`);
    }
    process.exitCode = failed === 0 ? 0 : 1;
}

await main();
