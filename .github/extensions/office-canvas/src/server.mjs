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

const UI_DIR = path.join(path.dirname(fileURLToPath(import.meta.url)), "ui");

const MIME = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".svg": "image/svg+xml",
};

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
 * `bytes=-N` is a *suffix* range meaning the final N bytes, which is how PDF
 * readers fetch the trailer to find the cross-reference table. Reading it as
 * `0-N` returns plausible-looking bytes from the wrong end of the file.
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

export class ViewerInstance {
    #server = null;
    #clients = new Set();
    #watcher = null;
    #refreshing = null;

    constructor({ cache, instanceId, workspacePath, log = () => {}, spawnFn = spawn }) {
        this.cache = cache;
        this.instanceId = instanceId;
        this.workspacePath = workspacePath ?? null;
        this.log = log;
        this.spawnFn = spawnFn;
        this.url = null;

        this.doc = null; // last known describe() payload
        this.status = "idle"; // idle | opening | ready | error
        this.error = null;
        this.lastPage = 1;

        this.#watcher = new FileWatcher({
            log,
            onChange: () => {
                this.#autoRefresh().catch((err) => this.log(`auto-refresh failed: ${err.message}`));
            },
        });
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
        };
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

    /** Re-renders if the file on disk changed. Returns the (possibly new) state. */
    async refresh({ force = false } = {}) {
        if (!this.doc) throw new DocumentError("not_open", "No document is open in this canvas.");
        if (this.#refreshing) return this.#refreshing;

        this.#refreshing = (async () => {
            const docPath = this.doc.path;
            const result = await this.cache.refresh(docPath);
            if (!result.changed && !force) return { changed: false, doc: this.doc };
            await this.cache.pdf(docPath);
            this.doc = result;
            this.status = "ready";
            this.error = null;
            await this.#watcher.acknowledge();
            this.#pushState();
            this.#broadcast("reloaded", { ...this.state, restorePage: this.lastPage });
            return { changed: true, doc: this.doc };
        })();

        try {
            return await this.#refreshing;
        } finally {
            this.#refreshing = null;
        }
    }

    async #autoRefresh() {
        if (!this.doc) return;
        this.log(`document changed on disk: ${path.basename(this.doc.path)}`);
        this.#broadcast("rendering", { path: this.doc.path });
        try {
            await this.refresh();
        } catch (err) {
            this.status = "error";
            this.error = { code: err.code ?? "refresh_failed", message: err.message };
            this.#pushState();
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
        // collide, and when they do the edit is refused with document_locked
        // rather than either side corrupting the file.
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
        if (route.startsWith("/api/")) return this.#serveApi(req, res, route, url);
        return this.#serveStatic(res, route);
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

        // The embedded PDF viewer requests byte ranges; without this it either
        // refuses to load or downloads the whole file for every page jump.
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
