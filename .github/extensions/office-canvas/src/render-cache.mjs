// Document service and PDF cache.
//
// Turns a .docx path into a rendered PDF, and keeps the per-document state that
// the canvas needs (metadata, outline, search). One Word host serves every open
// document; documents are keyed by their absolute path, which is the durable
// domain ID -- never the canvas instance id.
//
// The cache key is path + mtime + size. That makes reopening an unchanged
// document instant and makes an edited document re-render exactly once.

import { createHash } from "node:crypto";
import { mkdir, readdir, rm, stat } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { WordHost } from "./word/word-host.mjs";
import { DocumentReader } from "./word/document-reader.mjs";

export class DocumentError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "DocumentError";
        this.code = code;
    }
}

const shortHash = (value) => createHash("sha256").update(value).digest("hex").slice(0, 16);

/** Absolute, normalized, case-folded on Windows -- the identity of a document. */
export function normalizeDocPath(input) {
    if (typeof input !== "string" || input.trim() === "") {
        throw new DocumentError("invalid_path", "A document path is required.");
    }
    const abs = path.resolve(input.trim().replace(/^"(.*)"$/, "$1"));
    return process.platform === "win32" ? abs.replace(/\//g, "\\") : abs;
}

const identityOf = (docPath) => (process.platform === "win32" ? docPath.toLowerCase() : docPath);

const SUPPORTED = new Set([".docx", ".docm", ".doc", ".dotx", ".rtf"]);

/** Normalizes a path and rejects anything Word will not open as a document. */
export function requireSupported(rawPath) {
    const docPath = normalizeDocPath(rawPath);
    const ext = path.extname(docPath).toLowerCase();
    if (!SUPPORTED.has(ext)) {
        throw new DocumentError(
            "unsupported_type",
            `${ext || "This file"} is not a Word document. Supported: ${[...SUPPORTED].join(", ")}.`,
        );
    }
    return docPath;
}

export class RenderCache {
    /** identity -> { docPath, docId, workDir, pdfDir, key, meta, pdfFile, pending } */
    #docs = new Map();
    #reader = null;

    constructor({ cacheRoot, log = () => {} }) {
        this.cacheRoot = cacheRoot;
        this.log = log;
        this.host = new WordHost({ log, pidDir: path.join(cacheRoot, "pids") });
    }

    get wordVersion() {
        return this.host.wordVersion;
    }

    #stateFor(docPath) {
        const identity = identityOf(docPath);
        let state = this.#docs.get(identity);
        if (state) return state;

        const docId = shortHash(identity);
        state = {
            docPath,
            docId,
            workDir: path.join(this.cacheRoot, "work", docId),
            pdfDir: path.join(this.cacheRoot, "pdf", docId),
            key: null,
            meta: null,
            pdfFile: null,
            pending: null,
        };
        this.#docs.set(identity, state);
        return state;
    }

    async #fingerprint(docPath) {
        let info;
        try {
            info = await stat(docPath);
        } catch {
            throw new DocumentError("file_not_found", `No such file: ${docPath}`);
        }
        if (!info.isFile()) throw new DocumentError("file_not_found", `Not a file: ${docPath}`);
        return {
            key: shortHash(`${identityOf(docPath)}|${info.mtimeMs}|${info.size}`),
            mtimeMs: info.mtimeMs,
            size: info.size,
        };
    }

    /**
     * Opens (or re-opens, if the file changed on disk) a document in Word and
     * returns its metadata. Safe to call repeatedly.
     */
    async open(rawPath) {
        const docPath = requireSupported(rawPath);

        // Validate the file *before* registering any state, so a bad path never
        // leaves a half-open document behind.
        const { key } = await this.#fingerprint(docPath);
        const state = this.#stateFor(docPath);
        if (state.key === key && state.meta) return this.#describe(state);

        // A changed file needs a fresh working copy, so drop the old one first.
        if (state.key && state.key !== key) {
            await this.host.closeDocument({ docId: state.docId }).catch(() => {});
            state.pdfFile = null;
        }

        try {
            await mkdir(state.workDir, { recursive: true });
            const meta = await this.host.openDocument({
                docId: state.docId,
                path: docPath,
                workDir: state.workDir,
            });
            state.key = key;
            state.meta = meta;
            state.pdfFile = null;
        } catch (err) {
            if (!state.meta) this.#docs.delete(identityOf(docPath));
            throw err;
        }
        return this.#describe(state);
    }

    /**
     * Reads the document's structure map and revision token.
     *
     * Independent of `open`: it takes no cache state, holds nothing open, and
     * works whether or not this document is displayed in a canvas — the tools
     * are the product, the canvas is a display surface. It shares the Word
     * instance, so it costs a cold start only when nothing else has needed one.
     */
    readStructure(rawPath, options = {}) {
        const docPath = requireSupported(rawPath);
        if (!this.#reader) {
            this.#reader = new DocumentReader({
                host: this.host,
                workRoot: this.cacheRoot,
                log: this.log,
            });
        }
        return this.#reader.read(docPath, options);
    }

    #describe(state) {
        return {
            path: state.docPath,
            docId: state.docId,
            key: state.key,
            name: state.meta?.name ?? path.basename(state.docPath),
            title: state.meta?.title ?? "",
            author: state.meta?.author ?? "",
            pageCount: state.meta?.pageCount ?? 0,
            wordCount: state.meta?.wordCount ?? 0,
            sizeBytes: state.meta?.sizeBytes ?? 0,
            modifiedIso: state.meta?.modifiedIso ?? null,
        };
    }

    #require(docPath) {
        const state = this.#docs.get(identityOf(normalizeDocPath(docPath)));
        if (!state || !state.meta) {
            throw new DocumentError("not_open", "That document is not open in this canvas.");
        }
        return state;
    }

    /**
     * Returns the rendered PDF for a document, exporting it if the cached copy
     * is missing or stale. Concurrent callers share one export.
     */
    async pdf(rawPath) {
        const info = await this.open(rawPath);
        const state = this.#require(info.path);

        const file = path.join(state.pdfDir, `${state.key}.pdf`);
        if (state.pdfFile === file && existsSync(file)) {
            return { file, key: state.key, pageCount: info.pageCount };
        }
        if (existsSync(file)) {
            state.pdfFile = file;
            return { file, key: state.key, pageCount: info.pageCount };
        }
        if (state.pending) return state.pending;

        state.pending = (async () => {
            await mkdir(state.pdfDir, { recursive: true });
            const started = Date.now();
            const result = await this.host.exportPdf({ docId: state.docId, out: file });
            state.pdfFile = file;
            this.log(`render: exported ${path.basename(state.docPath)} in ${Date.now() - started}ms`);
            await this.#pruneOldPdfs(state);
            return { file, key: state.key, pageCount: result?.pageCount ?? info.pageCount };
        })();

        try {
            return await state.pending;
        } finally {
            state.pending = null;
        }
    }

    /** Keeps only the current render; older versions are dead weight. */
    async #pruneOldPdfs(state) {
        try {
            const entries = await readdir(state.pdfDir);
            await Promise.all(
                entries
                    .filter((name) => name.endsWith(".pdf") && name !== `${state.key}.pdf`)
                    .map((name) => rm(path.join(state.pdfDir, name), { force: true }).catch(() => {})),
            );
        } catch {
            /* pruning is best-effort */
        }
    }

    /**
     * Re-checks the file on disk. Returns whether anything changed, so callers
     * can avoid pushing a pointless reload to the viewer.
     */
    async refresh(rawPath) {
        const docPath = normalizeDocPath(rawPath);
        const state = this.#docs.get(identityOf(docPath));
        const previousKey = state?.key ?? null;
        const info = await this.open(docPath);
        return { changed: info.key !== previousKey, ...info };
    }

    async outline(rawPath, { limit } = {}) {
        const state = this.#require(rawPath);
        return this.host.outline({ docId: state.docId, limit });
    }

    async search(rawPath, query, { limit, matchCase, wholeWord } = {}) {
        const state = this.#require(rawPath);
        if (typeof query !== "string" || query.trim() === "") {
            throw new DocumentError("invalid_query", "A search query is required.");
        }
        return this.host.search({ docId: state.docId, query, limit, matchCase, wholeWord });
    }

    async text(rawPath, { fromPage, toPage } = {}) {
        const state = this.#require(rawPath);
        return this.host.text({ docId: state.docId, fromPage, toPage });
    }

    async info(rawPath) {
        const state = this.#require(rawPath);
        const info = await this.host.info({ docId: state.docId });
        // The host only ever sees the temp working copy, so its file identity is
        // an internal detail -- report the document the user actually opened.
        return { ...info, name: path.basename(state.docPath), path: state.docPath };
    }

    async close(rawPath) {
        const docPath = normalizeDocPath(rawPath);
        const identity = identityOf(docPath);
        const state = this.#docs.get(identity);
        if (!state) return;
        this.#docs.delete(identity);
        await this.host.closeDocument({ docId: state.docId }).catch(() => {});
        await rm(state.workDir, { recursive: true, force: true }).catch(() => {});
    }

    /** Number of documents actually open in Word -- used to decide when to quit it. */
    get openCount() {
        let n = 0;
        for (const state of this.#docs.values()) if (state.meta) n += 1;
        return n;
    }

    async dispose() {
        const states = [...this.#docs.values()];
        this.#docs.clear();
        await this.host.dispose().catch(() => {});
        await Promise.all(
            states.map((s) => rm(s.workDir, { recursive: true, force: true }).catch(() => {})),
        );
    }

    /** Last-resort synchronous cleanup for process exit hooks. */
    reap() {
        return this.host.reapOwnedWord();
    }
}
