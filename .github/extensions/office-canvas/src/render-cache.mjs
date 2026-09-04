// Document service and PDF cache.
//
// Turns a .docx path into a rendered PDF, and keeps the per-document state that
// the canvas needs (metadata, outline, search). One Word host serves every open
// document; documents are keyed by their absolute path, which is the durable
// domain ID -- never the canvas instance id.
//
// The cache key is path + mtime + size. That makes reopening an unchanged
// document instant and makes an edited document re-render exactly once.

import { createHash, randomUUID } from "node:crypto";
import { mkdir, readFile, readdir, rm, stat } from "node:fs/promises";
import { existsSync } from "node:fs";
import path from "node:path";
import { WordHost } from "./word/word-host.mjs";
import { DocumentReader } from "./word/document-reader.mjs";
import { DocumentEditor } from "./word/document-editor.mjs";
import { DocumentAuthor } from "./word/document-author.mjs";
import { extractOutlineEntries, resolveOutlineEntries } from "./word/outline-map.mjs";

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

/**
 * The extensions Word will open as a document. **The single source of truth.**
 *
 * Every statement of this set elsewhere -- the tool parameter descriptions the
 * model reads, the picker's workspace scan -- is derived from it rather than
 * restated. Three independent copies had already drifted: two tool
 * descriptions and the picker's own set all omitted `.dotx`, so the agent was
 * told templates were unsupported and the picker hid them, while the code
 * opened them happily. A restatement is a copy that will drift; deriving is
 * what makes adding an extension here update the contract everywhere.
 */
export const SUPPORTED = new Set([".docx", ".docm", ".doc", ".dotx", ".rtf"]);

/** The supported extensions as prose, for tool descriptions and messages. */
export const supportedList = () => [...SUPPORTED].join(", ");

/** Normalizes a path and rejects anything Word will not open as a document. */
export function requireSupported(rawPath) {
    const docPath = normalizeDocPath(rawPath);
    const ext = path.extname(docPath).toLowerCase();
    if (!SUPPORTED.has(ext)) {
        throw new DocumentError(
            "unsupported_type",
            `${ext || "This file"} is not a Word document. Supported: ${supportedList()}.`,
        );
    }
    return docPath;
}

export class RenderCache {
    /** identity -> { docPath, docId, workDir, pdfDir, key, meta, pdfFile, pending, documentPending, closing } */
    #docs = new Map();
    #reader = null;
    #editor = null;
    #author = null;

    constructor({ cacheRoot, snapshotRoot = null, log = () => {} }) {
        this.cacheRoot = cacheRoot;
        // Defaults inside the cache root, but separable: snapshots are the only
        // copy of a document's previous state, so a deployment that wants them
        // somewhere more durable than a cache directory can say so.
        this.snapshotRoot = snapshotRoot ?? cacheRoot;
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
            documentPending: Promise.resolve(),
            closing: false,
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
        const state = this.#stateFor(docPath);
        if (state.closing) {
            await state.documentPending;
            return this.open(docPath);
        }

        return this.#enqueueDocumentOperation(state, async () => {
            const { key } = await this.#fingerprint(docPath);
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
        });
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
        return this.#readerFor().read(docPath, options);
    }

    #readerFor() {
        if (!this.#reader) {
            this.#reader = new DocumentReader({
                host: this.host,
                workRoot: this.cacheRoot,
                log: this.log,
            });
        }
        return this.#reader;
    }

    #editorFor() {
        if (!this.#editor) {
            this.#editor = new DocumentEditor({
                reader: this.#readerFor(),
                host: this.host,
                // Snapshots outlive the render cache on purpose: a PDF can be
                // rebuilt from the document, and the pre-edit bytes cannot be
                // rebuilt from anything. Clearing the cache must not throw away
                // the only copy of what the document said before an edit.
                snapshotRoot: this.snapshotRoot,
                log: this.log,
            });
        }
        return this.#editor;
    }

    #authorFor() {
        if (!this.#author) {
            this.#author = new DocumentAuthor({ reader: this.#readerFor(), host: this.host, log: this.log });
        }
        return this.#author;
    }

    /**
     * Authors a new document from a spec and returns it as it stands on disk.
     *
     * `normalizeDocPath` rather than `requireSupported`: the two ask different
     * questions. `SUPPORTED` is the set Word will *open*, and this path is not
     * opening anything — it is choosing what to write, which `DocumentAuthor`
     * narrows further because `SaveAs2` emits one format. Sharing the reader's
     * set here would accept `.rtf` and produce a file whose extension lies about
     * its contents.
     *
     * No `#invalidate`: the file did not exist a moment ago, so there is nothing
     * cached under that identity to drop.
     */
    async createDocument(rawPath, spec) {
        return this.#authorFor().create(normalizeDocPath(rawPath), spec);
    }

    /**
     * Applies one edit to the user's own document, in place (ADR 0005), and
     * returns the document as it stands afterwards.
     */
    async editDocument(rawPath, intent, options = {}) {
        const docPath = requireSupported(rawPath);
        const state = this.#docs.get(identityOf(docPath));
        if (!state) return this.#editorFor().edit(docPath, intent, options);
        if (state.closing) {
            throw new DocumentError("not_open", "That document is not open in this canvas.");
        }
        return this.#enqueueDocumentOperation(state, async () => {
            const result = await this.#editorFor().edit(docPath, intent, options);
            await this.#invalidateState(state);
            return result;
        });
    }

    /** Restores the newest snapshot of a document and discards it. */
    async revertDocument(rawPath, options = {}) {
        const docPath = requireSupported(rawPath);
        const state = this.#docs.get(identityOf(docPath));
        if (!state) return this.#editorFor().revert(docPath, options);
        if (state.closing) {
            throw new DocumentError("not_open", "That document is not open in this canvas.");
        }
        return this.#enqueueDocumentOperation(state, async () => {
            const result = await this.#editorFor().revert(docPath, options);
            await this.#invalidateState(state);
            return result;
        });
    }

    editHistory(rawPath) {
        return this.#editorFor().history(requireSupported(rawPath));
    }

    /**
     * Drops the cached render of a document we have just changed.
     *
     * The cache key already includes mtime and size, so a stale PDF would not
     * be served — but the Word-side document handle for an open canvas would
     * still be the pre-edit one, and the canvas would keep showing it.
     */
    async #invalidateState(state) {
        state.key = null;
        state.pdfFile = null;
        try {
            await this.host.closeDocument({ docId: state.docId });
        } catch {
            /* the document may already be closed; the next open re-creates it */
        }
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
        if (!state || !state.meta || state.closing) {
            throw new DocumentError("not_open", "That document is not open in this canvas.");
        }
        return state;
    }

    #enqueueDocumentOperation(state, operation) {
        const run = state.documentPending.then(operation);
        state.documentPending = run.then(
            () => undefined,
            () => undefined,
        );
        return run;
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
        return this.#enqueueDocumentOperation(state, async () => {
            const out = path.join(state.workDir, `outline-${randomUUID()}.xml`);
            try {
                await this.host.outlineMarkup({ docId: state.docId, out });
                const entries = extractOutlineEntries(await readFile(out, "utf8"), { limit: limit ?? 2000 });
                if (!entries.length) return { headings: [], count: 0 };
                const { positions } = await this.host.outlinePositions({
                    docId: state.docId,
                    wordIndices: entries.map((entry) => entry.wordIndex),
                });
                return resolveOutlineEntries(entries, positions);
            } finally {
                await rm(out, { force: true }).catch(() => {});
            }
        });
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
        state.closing = true;
        return this.#enqueueDocumentOperation(state, async () => {
            this.#docs.delete(identity);
            await this.host.closeDocument({ docId: state.docId }).catch(() => {});
            await rm(state.workDir, { recursive: true, force: true }).catch(() => {});
        });
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
