// Per-instance loopback server: serves the viewer UI, the rendered PDF, and a
// small JSON API, and pushes state changes to the iframe over SSE.
//
// The iframe has no privileged bridge to the extension, so every control in the
// UI goes through these endpoints.

import { createServer } from "node:http";
import { createReadStream } from "node:fs";
import { readFile, readdir, stat } from "node:fs/promises";
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DocumentError, normalizeDocPath, SUPPORTED } from "./render-cache.mjs";
import { FileWatcher } from "./watcher.mjs";
import { addRecent, readRecents } from "./store.mjs";
import { loadWorker, VENDOR_DIR } from "./vendor-assets.mjs";

const UI_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "ui");

const MIME = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".mjs": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
};

/**
 * The vendored files this server will serve, and nothing else.
 *
 * An allowlist rather than a basename check. `src/ui/vendor/` also holds the
 * worker *parts* and the manifest, none of which any client has business
 * fetching -- a part served on its own is a fragment of a program, and handing
 * one out would make a truncated worker look like a working route.
 */
const VENDOR_FILES = new Set(["pdf.min.mjs", "pdf-text-layer.css"]);

/** Served by reassembling the committed parts; never present as one file. */
const VENDOR_WORKER_ROUTE = "/vendor/pdf.worker.min.mjs";

// The picker offers exactly what the cache will open. Its own copy of this set
// had drifted -- it omitted `.dotx`, so a template in the workspace was hidden
// from the picker while `open_document` on the same path succeeded. Exported so
// a unit test can assert the two stay equal.
export const DOC_EXTENSIONS = SUPPORTED;
const SCAN_SKIP = new Set(["node_modules", ".git", "bin", "obj", "dist", "build", ".venv", "__pycache__"]);

const json = (res, status, body) => {
    const payload = JSON.stringify(body);
    res.writeHead(status, {
        "Content-Type": "application/json; charset=utf-8",
        "Content-Length": Buffer.byteLength(payload),
        "Cache-Control": "no-store",
    });
    res.end(payload);
};

/**
 * Parses a single-range HTTP `Range` header against an entity of `size` bytes.
 *
 * Returns `null` when the whole entity should be sent (absent, malformed, or
 * multi-range — RFC 9110 says ignore what we cannot honour), the string
 * `"unsatisfiable"` for a 416, or an inclusive `{ start, end }`.
 *
 * `bytes=-N` is a *suffix* range meaning the final N bytes. Reading it as `0-N`
 * would return plausible-looking bytes from the wrong end of the file, which is
 * why it is handled separately.
 *
 * **This has no measured consumer.** An earlier version of this comment said
 * suffix ranges are "how PDF readers fetch the trailer to find the
 * cross-reference table". That is sound as a description of the format, and it
 * is wrong as a claim about this code: probed against the bundled pdf.js
 * (`spikes/pdfjs/probes/probe-range-requests.mjs`), the default configuration
 * issues one unranged GET and **zero** Range requests of any kind, computing the
 * tail from `Content-Length` instead. Forcing ranged mode produced 14 requests
 * and still no suffix range. The branch stays because it is correct and cheap,
 * and because the route is reachable by anything that speaks HTTP — but nothing
 * should be verified by assuming pdf.js exercises it.
 */
export function parseByteRange(headerValue, size) {
    const match = /^bytes=(\d*)-(\d*)$/.exec((headerValue ?? "").trim());
    if (!match) return null;
    const [, rawStart, rawEnd] = match;
    if (rawStart === "" && rawEnd === "") return null;

    if (rawStart === "") {
        const wanted = Number(rawEnd);
        if (!Number.isFinite(wanted) || wanted === 0 || size === 0) return "unsatisfiable";
        return { start: Math.max(0, size - wanted), end: size - 1 };
    }

    const start = Number(rawStart);
    if (!Number.isFinite(start) || start >= size) return "unsatisfiable";
    const end = rawEnd === "" ? size - 1 : Math.min(Number(rawEnd), size - 1);
    if (!Number.isFinite(end) || end < start) return "unsatisfiable";
    return { start, end };
}

async function readBody(req, limit = 64 * 1024) {
    const chunks = [];
    let size = 0;
    for await (const chunk of req) {
        size += chunk.length;
        if (size > limit) throw new DocumentError("payload_too_large", "Request body too large.");
        chunks.push(chunk);
    }
    if (!chunks.length) return {};
    try {
        return JSON.parse(Buffer.concat(chunks).toString("utf8"));
    } catch {
        throw new DocumentError("invalid_json", "Request body was not valid JSON.");
    }
}

/** Recursively lists Word documents under a root, breadth-first and bounded. */
async function scanForDocuments(root, { limit = 200, maxDepth = 6 } = {}) {
    if (!root) return [];
    const found = [];
    const queue = [{ dir: root, depth: 0 }];
    while (queue.length && found.length < limit) {
        const { dir, depth } = queue.shift();
        let entries;
        try {
            entries = await readdir(dir, { withFileTypes: true });
        } catch {
            continue;
        }
        for (const entry of entries) {
            if (found.length >= limit) break;
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                if (depth >= maxDepth) continue;
                if (entry.name.startsWith(".") || SCAN_SKIP.has(entry.name.toLowerCase())) continue;
                queue.push({ dir: full, depth: depth + 1 });
            } else if (DOC_EXTENSIONS.has(path.extname(entry.name).toLowerCase())) {
                if (entry.name.startsWith("~$")) continue; // Word lock file
                found.push({ path: full, name: entry.name, relative: path.relative(root, full) });
            }
        }
    }
    return found.sort((a, b) => a.relative.localeCompare(b.relative));
}

/**
 * Auto-refresh failures that will not clear on their own.
 *
 * Everything else is retried, and the asymmetry is deliberate. This set can be
 * stated from what the code knows: an ACL, an extension Word will not open, and
 * our own state errors are all conditions that are exactly as true in five
 * seconds as they are now — `document-reader.mjs` says so of `permission_denied`
 * in as many words, beside the `file_locked` it contrasts with. The transient
 * set cannot be stated that way, because the failing call reaches Word, the
 * filesystem and a PowerShell host, and a code we have not enumerated is far
 * likelier to be a passing condition than a permanent one. Retrying is bounded,
 * so treating an unknown code as transient costs a few seconds before it is
 * reported; treating a transient one as terminal cost the user their panel.
 */
const TERMINAL_REFRESH_CODES = new Set(["permission_denied", "unsupported_type", "invalid_path", "not_open"]);

/**
 * Waits between auto-refresh attempts. Four attempts across ~6s.
 *
 * Long enough to outlast the flush at the tail of a save, short enough that a
 * document which really is gone is reported while the user still associates the
 * message with what they did.
 */
const AUTO_REFRESH_DELAYS_MS = [500, 1500, 4000];

export class ViewerInstance {
    #server = null;
    #clients = new Set();
    #watcher = null;
    #refreshing = null;
    /** Resolves the live RenderCache. See the constructor. */
    #cacheSource = null;

    // Counts change records as they are recorded. `changed` tells you a render
    // happened; this tells you whether the record in `this.change` is still the
    // one a given render was started for. A render that began before a record
    // existed invoked `cache.refresh` before that edit landed, so its image
    // cannot be shown to contain it however recently it finished.
    #changeEpoch = 0;

    constructor({
        cache,
        instanceId,
        workspacePath,
        log = () => {},
        spawnFn = spawn,
        vendorDir = VENDOR_DIR,
        autoRefreshDelaysMs = AUTO_REFRESH_DELAYS_MS,
    }) {
        // A function, or a RenderCache directly. The function form is what the
        // extension passes, and it is not a convenience: a panel outlives Word
        // being shut down, and this instance used to keep the exact object it
        // was handed. Once that cache was disposed -- the last canvas closing,
        // or the idle timer -- the panel held a host that answers everything
        // with "The Word host has been shut down.", for the rest of the
        // session, with nothing in the code able to replace it. That was #61.
        // Resolving per use means the panel picks up whatever cache is live now.
        this.#cacheSource = typeof cache === "function" ? cache : () => cache;
        this.instanceId = instanceId;
        this.workspacePath = workspacePath ?? null;
        this.log = log;
        this.spawnFn = spawnFn;
        this.vendorDir = vendorDir;
        this.autoRefreshDelaysMs = autoRefreshDelaysMs;
        this.url = null;

        this.doc = null; // last known describe() payload
        this.status = "idle"; // idle | opening | ready | error
        this.error = null;
        this.lastPage = 1;

        // What the last operation changed, or null. Stamped with the render key
        // of the document state it describes; see #changeIfCurrent.
        this.change = null;

        this.#watcher = new FileWatcher({
            log,
            // Deliberately not caught here. A rejection is how the watcher is
            // told the change was not consumed, so it puts its fingerprint back
            // and the next write reports the change again. Swallowing it -- what
            // this did -- left the watcher believing the edit had been handled
            // while the panel sat in `error` with nothing left to re-fire on.
            onChange: () => this.#autoRefresh(),
        });
    }

    /** The live RenderCache, resolved per use rather than captured. */
    get cache() {
        return this.#cacheSource();
    }

    async start() {
        if (this.url) return this.url;
        this.#server = createServer((req, res) => {
            this.#handle(req, res).catch((err) => {
                this.log(`request failed: ${err.stack ?? err.message}`);
                if (!res.headersSent) json(res, 500, { error: { code: "internal", message: err.message } });
                else res.end();
            });
        });
        this.#server.on("clientError", (_err, socket) => socket.destroy());
        await new Promise((resolve, reject) => {
            this.#server.once("error", reject);
            this.#server.listen(0, "127.0.0.1", resolve);
        });
        this.url = `http://127.0.0.1:${this.#server.address().port}/`;
        return this.url;
    }

    // --- state ---------------------------------------------------------------

    get state() {
        return {
            status: this.status,
            error: this.error,
            doc: this.doc,
            lastPage: this.lastPage,
            wordVersion: this.cache.wordVersion,
            workspacePath: this.workspacePath,
            pdfUrl: this.doc ? `/pdf/${this.doc.key}.pdf` : null,
            change: this.#changeIfCurrent(),
        };
    }

    /**
     * The change record, but only while the document is still in the state it
     * describes.
     *
     * The render key is `hash(path|mtimeMs|size)`, so it moves on any write --
     * ours, the user's Word, or a generator script replacing the file. Keying
     * the record on it means an overlay cannot outlive the content it points at,
     * which is the address rule (ADR 0006) applied to the display. Anything else
     * would leave a highlight sitting over text that has since moved.
     *
     * This is not a new assumption: the same key already decides that a rendered
     * PDF may be cached immutably, so if it could go stale the viewer would
     * already be showing the wrong pages.
     */
    #changeIfCurrent() {
        if (!this.change || !this.doc) return null;
        return this.change.docKey === this.doc.key ? this.change : null;
    }

    /**
     * Ties whatever record is currently held to the render just produced.
     *
     * Stamping is one operation in one place, applied to whatever `this.change`
     * holds, rather than a second expression that re-decides what the record
     * should be. Two spellings of the same rule can drift, and here they would
     * drift silently: both produce a plausible-looking overlay state and neither
     * is wrong enough to throw.
     */
    #stampChange(docKey) {
        if (this.change) this.change = { ...this.change, docKey };
    }

    #broadcast(event, data = {}) {
        const frame = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
        for (const client of this.#clients) {
            try {
                client.write(frame);
            } catch {
                this.#clients.delete(client);
            }
        }
    }

    #pushState() {
        this.#broadcast("state", this.state);
    }

    // --- document operations -------------------------------------------------

    async openDocument(rawPath) {
        const docPath = normalizeDocPath(rawPath);
        this.status = "opening";
        this.error = null;
        this.#pushState();
        try {
            const info = await this.cache.open(docPath);
            // Render eagerly: the viewer is useless until the PDF exists, and
            // doing it here means one status transition instead of two.
            await this.cache.pdf(docPath);
            if (this.doc && this.doc.path !== info.path) this.lastPage = 1;
            this.doc = info;
            // A record describes one document; carrying it to another would be
            // meaningless. #changeIfCurrent would reject it anyway on the key,
            // but leaving it set would make that the only thing standing between
            // a stale overlay and the user.
            this.change = null;
            this.status = "ready";
            await this.#watcher.watch(info.path);
            await this.#watcher.acknowledge();
            await addRecent(info);
            this.#pushState();
            return info;
        } catch (err) {
            this.status = "error";
            this.error = { code: err.code ?? "open_failed", message: err.message };
            this.#pushState();
            throw err;
        }
    }

    /**
     * Re-renders if the file on disk changed. Returns the (possibly new) state.
     *
     * `change` describes what an operation just did, and is attached to the
     * render this refresh produces. It is stamped with that render's key here
     * rather than by the caller, because the caller cannot know the key until
     * the re-render has happened.
     */
    async refresh({ force = false, change } = {}) {
        if (!this.doc) throw new DocumentError("not_open", "No document is open in this canvas.");

        // Recorded before anything else can publish state. The render below
        // re-assigns it with a render key, so this bare assignment matters for
        // one reason only: it *invalidates the previous overlay immediately*.
        // Without it, an in-flight refresh that turns out to be a no-op would
        // broadcast state while `this.change` still held the last edit's record,
        // stamped with a key that has not moved -- so a stale highlight would go
        // out describing an edit that has since been superseded. Left unstamped
        // it is dropped by #changeIfCurrent, which is the safe direction:
        // stamping it with the *current* key here would instead publish the new
        // edit's text over the pre-edit render.
        if (change !== undefined) {
            this.change = change;
            this.#changeEpoch += 1;
        }

        if (this.#refreshing) {
            const joined = await this.#refreshing;
            // Only a caller with no opinion about the overlay may settle for a
            // joined render, and `changed` is what rules out a no-op: one that
            // returned early without advancing `this.doc` leaves the key at the
            // pre-edit render.
            //
            // A caller *with* a record may never settle for one, and the reason
            // is stronger than "usually unsafe" -- it is never provable. The
            // joined render invoked `cache.refresh` before this record existed,
            // so it read the file at some moment we cannot observe and may
            // predate this edit. `changed` says a render happened; it does not
            // say which. On two rapid edits -- A lands, its refresh starts, B
            // lands and joins it -- the join reports `changed: true` for an
            // image that is post-A and pre-B, and B's overlay lands on A's page.
            //
            // So this branch deliberately does not test the record against the
            // render. It was written that way first and the condition could
            // never be false: a render already in flight always began at an
            // earlier epoch than a record created by the call that joins it.
            // A guard with no reachable failure is worse than none, because it
            // reads as protection.
            if (joined.changed && change === undefined) return joined;
            // Fall through. A forced caller renders for real, which is what makes
            // its record safe to stamp -- the overlay is not lost, it is paid for
            // with the one render that can be proven to contain the edit. An
            // unforced caller returns the joined render with the record left
            // unstamped, and #changeIfCurrent drops it: losing an overlay is the
            // safe direction, drawing one over the wrong image is not.
            if (!force) return joined;
        }

        // Recorded before `cache.refresh` is invoked, which is the moment this
        // render's view of the file is fixed. If a newer record arrives while it
        // is in flight, `this.change` below is no longer the record this render
        // was started for, and stamping it would tie that newer edit to this
        // older image -- the same stale pairing the join above refuses, reached
        // from the opposite side.
        const startedEpoch = this.#changeEpoch;
        this.#refreshing = (async () => {
            const docPath = this.doc.path;
            const result = await this.cache.refresh(docPath);
            if (!result.changed && !force) return { changed: false, doc: this.doc };
            await this.cache.pdf(docPath);
            this.doc = result;
            this.status = "ready";
            this.error = null;
            // `#stampChange` stamps whatever `this.change` currently holds, which
            // is deliberate -- but a caller that arrived while this render was in
            // flight has already overwritten it with a newer record. Stamping then
            // would tie *their* edit to *this* render, which is the same stale
            // pairing the join above refuses, reached from the opposite side.
            if (change !== undefined && this.#changeEpoch === startedEpoch) this.#stampChange(result.key);
            await this.#watcher.acknowledge();
            this.#pushState();
            this.#broadcast("reloaded", { ...this.state, restorePage: this.lastPage });
            return { changed: true, doc: this.doc };
        })();

        // Cleared only if it is still ours. A forced caller that joined a no-op
        // falls through and installs its own promise here, and this owner's
        // `finally` may run afterwards -- clearing the field unconditionally would
        // null out that successor and let a third caller start a concurrent
        // refresh against the same document.
        const mine = this.#refreshing;
        try {
            return await mine;
        } finally {
            if (this.#refreshing === mine) this.#refreshing = null;
        }
    }

    /**
     * Re-renders because the file moved on disk. Retries a failure rather than
     * latching one, and rethrows if it gives up.
     *
     * The failure this exists for is a race, not a fault: the watcher fires
     * ~700 ms after the last write (300 ms debounce, then two 200 ms stable
     * rounds), which can land inside the tail of a save while the writer still
     * holds the file. One `refresh` at that instant throws, and the panel used
     * to stop there — `status: "error"`, no retry, no re-arm — over a document
     * that was readable again a moment later.
     *
     * **What actually failed in the reported case is not established.** The
     * panel was observed serving `file_locked` minutes after the save, but three
     * panels were watching the same document and contending on it, which is an
     * alternative cause that was not excluded. So this does not encode a belief
     * about which code Word's save produces: it retries anything not known to be
     * permanent (see `TERMINAL_REFRESH_CODES`) and reports whatever is left.
     *
     * Rethrowing is the third part, and it is not error propagation — nothing
     * above catches it. It is the signal to `FileWatcher` that this change was
     * not consumed, which is what lets a *later* save re-report it and recover a
     * panel this one gave up on.
     */
    async #autoRefresh() {
        if (!this.doc) return;
        this.log(`document changed on disk: ${path.basename(this.doc.path)}`);
        this.#broadcast("rendering", { path: this.doc.path });
        // Deliberately does not clear the change record. This fires for an
        // external writer *and* as an echo of our own save, and cannot tell them
        // apart -- clearing here would race the edit that set the record and
        // silently lose the overlay. The render key can tell them apart: it
        // moves only if the file actually differs, so #changeIfCurrent drops the
        // record exactly when an external write really did land.

        // A panel already in `error` has to render for real to leave it.
        // `refresh` clears the error only on the branch that re-renders, so an
        // unforced call that short-circuited on `changed: false` -- which is
        // exactly what a settle landing on an event that moved no bytes
        // produces -- would return successfully and leave the panel reading
        // `error` over a perfectly current image.
        const force = this.status === "error";

        for (let attempt = 0; ; attempt++) {
            try {
                await this.refresh({ force });
                return;
            } catch (err) {
                const code = err.code ?? "refresh_failed";
                const delay = TERMINAL_REFRESH_CODES.has(code) ? undefined : this.autoRefreshDelaysMs[attempt];
                if (delay === undefined) {
                    this.status = "error";
                    this.error = { code, message: err.message };
                    this.#pushState();
                    throw err;
                }
                this.log(`auto-refresh failed (${code}); retrying in ${delay}ms`);
                await new Promise((resolve) => setTimeout(resolve, delay));
            }
        }
    }

    setPage(page) {
        const n = Number(page);
        if (Number.isFinite(n) && n >= 1) this.lastPage = Math.floor(n);
        return this.lastPage;
    }

    goToPage(page) {
        const target = this.setPage(page);
        this.#broadcast("goto", { page: target });
        return target;
    }

    openInWord() {
        if (!this.doc) throw new DocumentError("not_open", "No document is open in this canvas.");
        // Launches the user's own visible Word on the original.
        //
        // This used to be justified with "safe because our hidden instance only
        // ever holds a temp copy, never the original file". That justification
        // died with edit_document, which opens the original by design (ADR 0005).
        // The invariant that replaces it is narrower but still holds: the hidden
        // instance holds the original only for the length of one operation, and
        // every edit is gated on a write-handle probe first. So the two can
        // collide, and when they do the edit is refused with `file_locked`
        // rather than either side corrupting the file. (`document_locked` is
        // the host's internal status for it; the code a caller sees is
        // `file_locked`.)
        //
        // The collision is asymmetric and only bites in this direction: a
        // document held by the user's Word still *reads* fine, because the read
        // works on a copy. Only the edit needs the original.
        //
        // Deliberately NOT `cmd.exe /c start`. Node quotes an argv element only
        // when it holds a space, tab or quote, and `cmd.exe` then applies its own
        // parse on top of that -- so a filename with no space and a `&`, `^` or a
        // `%VAR%` pair reaches cmd unquoted and is mangled. Measured in
        // spikes/isolation/probes/probe-open-in-word-quoting.mjs: 3 of 9 ordinary
        // filenames corrupted, including `%PATH%.docx` expanding the environment
        // variable into the path. `explorer.exe` opens the same default handler
        // with no shell parse in the way, and corrupted 0 of 9.
        //
        // explorer.exe exits 1 even on success, which is why nothing here reads
        // the exit code.
        const child = this.spawnFn("explorer.exe", [this.doc.path], {
            detached: true,
            stdio: "ignore",
            windowsHide: true,
        });
        child.unref();
        return { launched: true, path: this.doc.path };
    }

    async browse() {
        const [workspaceDocs, recents] = await Promise.all([
            scanForDocuments(this.workspacePath),
            readRecents(),
        ]);
        const existing = await Promise.all(
            recents.map(async (entry) => {
                try {
                    await stat(entry.path);
                    return entry;
                } catch {
                    return null;
                }
            }),
        );
        return { workspacePath: this.workspacePath, workspaceDocs, recents: existing.filter(Boolean) };
    }

    // --- http ----------------------------------------------------------------

    async #handle(req, res) {
        const url = new URL(req.url, "http://127.0.0.1");
        const route = url.pathname;

        if (route === "/events") return this.#serveEvents(req, res);
        if (route.startsWith("/pdf/")) return this.#servePdf(req, res);
        if (route.startsWith("/vendor/")) return this.#serveVendor(req, res, route);
        if (route.startsWith("/api/")) return this.#serveApi(req, res, route, url);
        return this.#serveStatic(res, route);
    }

    /**
     * Serves the vendored pdf.js, reassembling the worker from its parts.
     *
     * Cached by ETag rather than served immutable. The URL carries no version,
     * so a `max-age` far in the future would leave a browser holding last
     * version's worker against this version's `pdf.min.mjs` -- two halves of
     * pdf.js that were never tested together. Revalidation costs one loopback
     * round trip and a 304.
     */
    async #serveVendor(req, res, route) {
        if (route === VENDOR_WORKER_ROUTE) {
            let worker;
            try {
                // Awaited in full before a single header is written: a part that
                // fails to read must produce a failed request, not a 200 whose
                // body stops in the middle of an expression.
                worker = await loadWorker(this.vendorDir);
            } catch (err) {
                this.log(`could not serve the pdf.js worker: ${err.message}`);
                return json(res, 500, {
                    error: { code: err.code ?? "vendor_unavailable", message: err.message },
                });
            }
            if (req.headers["if-none-match"] === worker.etag) {
                res.writeHead(304, { ETag: worker.etag, "Cache-Control": "private, must-revalidate" });
                return res.end();
            }
            res.writeHead(200, {
                "Content-Type": MIME[".mjs"],
                "Content-Length": worker.buffer.byteLength,
                "Cache-Control": "private, must-revalidate",
                ETag: worker.etag,
            });
            if (req.method === "HEAD") return res.end();
            return res.end(worker.buffer);
        }

        const name = path.basename(route);
        if (!VENDOR_FILES.has(name)) {
            res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
            return res.end("Not found");
        }
        try {
            const body = await readFile(path.join(this.vendorDir, name));
            res.writeHead(200, {
                "Content-Type": MIME[path.extname(name).toLowerCase()] ?? "application/octet-stream",
                "Content-Length": body.length,
                "Cache-Control": "private, must-revalidate",
            });
            if (req.method === "HEAD") return res.end();
            return res.end(body);
        } catch (err) {
            this.log(`could not serve vendored ${name}: ${err.message}`);
            // Re-vendoring fixes exactly one cause, so only that cause is named.
            // A permissions or I/O failure reports its own code instead: telling
            // someone to re-run the vendoring step would send them after the
            // wrong thing, and for EACCES the step appears to succeed while the
            // file stays unservable. Discriminated on `err.code`, never on the
            // message, which is localized.
            const missing = err.code === "ENOENT";
            return json(res, 500, {
                error: {
                    code: missing ? "vendor_missing" : "vendor_unreadable",
                    message: missing
                        ? `The vendored pdf.js file '${name}' is missing. Run 'node tools/vendor-pdfjs.mjs'.`
                        : `The vendored pdf.js file '${name}' could not be read (${err.code ?? "unknown error"}).`,
                },
            });
        }
    }

    #serveEvents(req, res) {
        res.writeHead(200, {
            "Content-Type": "text/event-stream; charset=utf-8",
            "Cache-Control": "no-cache",
            Connection: "keep-alive",
            "X-Accel-Buffering": "no",
        });
        res.write(`event: state\ndata: ${JSON.stringify(this.state)}\n\n`);
        this.#clients.add(res);
        const keepAlive = setInterval(() => {
            try {
                res.write(": ping\n\n");
            } catch {
                /* dropped below */
            }
        }, 25_000);
        req.on("close", () => {
            clearInterval(keepAlive);
            this.#clients.delete(res);
        });
    }

    async #servePdf(req, res) {
        if (!this.doc) return json(res, 404, { error: { code: "not_open", message: "No document open." } });
        let file;
        try {
            ({ file } = await this.cache.pdf(this.doc.path));
        } catch (err) {
            return json(res, 500, { error: { code: err.code ?? "render_failed", message: err.message } });
        }

        const info = await stat(file);
        const headers = {
            "Content-Type": "application/pdf",
            // The render key is in the URL, so an unchanged document can be
            // cached hard; a new render is simply a new URL.
            "Cache-Control": "private, max-age=31536000, immutable",
            "Accept-Ranges": "bytes",
        };

        // Range support is kept because `Accept-Ranges` is part of what lets a
        // client choose *not* to use it: pdf.js reads `Content-Length` from the
        // unranged response and fetches the file once. It is not here because
        // pdf.js asks for ranges — measured, it never does.
        const range = parseByteRange(req.headers.range, info.size);
        if (range === "unsatisfiable") {
            res.writeHead(416, { "Content-Range": `bytes */${info.size}` });
            return res.end();
        }
        if (range) {
            const { start, end } = range;
            res.writeHead(206, {
                ...headers,
                "Content-Range": `bytes ${start}-${end}/${info.size}`,
                "Content-Length": end - start + 1,
            });
            if (req.method === "HEAD") return res.end();
            return createReadStream(file, { start, end }).pipe(res);
        }

        res.writeHead(200, { ...headers, "Content-Length": info.size });
        if (req.method === "HEAD") return res.end();
        return createReadStream(file).pipe(res);
    }

    async #serveApi(req, res, route, url) {
        try {
            switch (`${req.method} ${route}`) {
                case "GET /api/state":
                    return json(res, 200, this.state);

                case "GET /api/browse":
                    return json(res, 200, await this.browse());

                case "POST /api/open": {
                    const body = await readBody(req);
                    const doc = await this.openDocument(body.path);
                    return json(res, 200, { doc, state: this.state });
                }

                case "POST /api/refresh": {
                    const body = await readBody(req);
                    return json(res, 200, await this.refresh({ force: body.force === true }));
                }

                case "POST /api/page": {
                    const body = await readBody(req);
                    return json(res, 200, { lastPage: this.setPage(body.page) });
                }

                case "POST /api/open-in-word":
                    return json(res, 200, this.openInWord());

                case "GET /api/outline": {
                    if (!this.doc) throw new DocumentError("not_open", "No document is open.");
                    const limit = url.searchParams.get("limit");
                    return json(
                        res,
                        200,
                        await this.cache.outline(this.doc.path, { limit: limit ? Number(limit) : undefined }),
                    );
                }

                case "GET /api/search": {
                    if (!this.doc) throw new DocumentError("not_open", "No document is open.");
                    const query = url.searchParams.get("q") ?? "";
                    const limit = url.searchParams.get("limit");
                    return json(
                        res,
                        200,
                        await this.cache.search(this.doc.path, query, {
                            limit: limit ? Number(limit) : 100,
                            matchCase: url.searchParams.get("matchCase") === "1",
                            wholeWord: url.searchParams.get("wholeWord") === "1",
                        }),
                    );
                }

                default:
                    return json(res, 404, { error: { code: "not_found", message: `No route ${route}` } });
            }
        } catch (err) {
            const status = err instanceof DocumentError ? 400 : 500;
            return json(res, status, { error: { code: err.code ?? "error", message: err.message } });
        }
    }

    async #serveStatic(res, route) {
        const name = route === "/" ? "index.html" : path.basename(route);
        const file = path.join(UI_DIR, name);
        // Everything served here is a fixed asset next to this file; basename()
        // keeps a crafted path from escaping the ui directory.
        try {
            const body = await readFile(file);
            res.writeHead(200, {
                "Content-Type": MIME[path.extname(name).toLowerCase()] ?? "application/octet-stream",
                "Content-Length": body.length,
                "Cache-Control": "no-store",
            });
            res.end(body);
        } catch {
            res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
            res.end("Not found");
        }
    }

    async close() {
        this.#watcher.close();
        for (const client of this.#clients) {
            try {
                client.end();
            } catch {
                /* already gone */
            }
        }
        this.#clients.clear();
        if (this.#server) {
            const server = this.#server;
            this.#server = null;
            this.url = null;
            await new Promise((resolve) => server.close(resolve));
        }
    }
}
