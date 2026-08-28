// read_document, end to end.
//
// One operation in the ADR 0005 sense: the document is opened, its markup taken
// in a single call, and it is closed again. Nothing stays open between reads.
//
// The document Word opens is a copy of the original, not the original itself.
// That is stronger than "close immediately": the original is never held at all,
// so a script can regenerate it and the canvas can auto-refresh even while a
// read is in flight. It also sidesteps mark-of-the-web -- Protected View
// refuses automation, and the only fix, `Unblock-File`, would modify the user's
// document to serve a read.
//
// The cost of reading a copy is that the copy and the file could diverge, so
// the revision token is taken on both sides of the read and a divergence is
// reported rather than papered over.

import { mkdir, readFile, rm, stat } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import path from "node:path";

import { fileRevisionToken } from "../revision-token.mjs";
import { buildStructureMap } from "./structure-map.mjs";

export class ReadError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "ReadError";
        this.code = code;
    }
}

const shortHash = (value) => createHash("sha256").update(value).digest("hex").slice(0, 16);

/** A cold Word start plus a full-package read. */
const FIRST_ATTEMPT_TIMEOUT_MS = 180_000;
/**
 * A retry runs against a *fresh* host, so it is not paying for whatever wedged
 * the first one. Retrying at the full startup budget turned one call into a
 * worst case of roughly twelve minutes of silence, which is indistinguishable
 * from a hang to whoever is waiting.
 */
const RETRY_TIMEOUT_MS = 45_000;
/** Hard ceiling on one `read`, retries and token re-reads included. */
const READ_BUDGET_MS = 240_000;

/** Errors that mean the host process died rather than the document being bad. */
const isHostFailure = (err) =>
    err?.code === "word_unavailable" || err?.code === "word_timeout" || err?.code === "no_such_document";

export class DocumentReader {
    #host;
    #workRoot;
    #log;
    #readToken;

    /**
     * `tokenOf` is injectable so the filesystem-error translation below can be
     * tested without a real exclusive lock, which is not reproducible on the
     * Linux runner that validates this repo.
     */
    constructor({ host, workRoot, log = () => {}, tokenOf = fileRevisionToken }) {
        this.#host = host;
        this.#workRoot = workRoot;
        this.#log = log;
        this.#readToken = tokenOf;
    }

    async #fetchMarkup(docPath, deadline) {
        // Unique per call, never derived from the document path alone. A
        // deterministic scratch id let two overlapping reads of one document
        // share a working directory, a `source.docx`, an output file *and* the
        // host-side docId -- so one call's cleanup could delete another's markup
        // mid-read, and one call's close could deregister the other's document.
        // Worse, it reopened the ADR 0005 hang: a scratch path re-derived while
        // a dying Word still held the previous `source.docx` asks a fresh Word
        // to open a file another Word has locked.
        const scratchId = `read-${shortHash(docPath.toLowerCase()).slice(0, 8)}-${randomUUID()}`;
        const workDir = path.join(this.#workRoot, "read", scratchId);
        const out = path.join(workDir, "structure.xml");
        await mkdir(workDir, { recursive: true });
        try {
            const remaining = () => deadline - Date.now();
            if (remaining() <= 0) throw new ReadError("word_timeout", "Ran out of time reading the document.");

            let result;
            try {
                result = await this.#host.structure({
                    docId: scratchId,
                    path: docPath,
                    workDir,
                    out,
                    timeoutMs: Math.min(FIRST_ATTEMPT_TIMEOUT_MS, remaining()),
                });
            } catch (err) {
                // A structure read registers no reopen args, so the host's own
                // replay cannot cover it. One retry: the next request starts a
                // fresh host transparently.
                if (!isHostFailure(err) || remaining() <= 0) throw err;
                this.#log(`read_document: retrying after ${err.code ?? "error"}`);
                result = await this.#host.structure({
                    docId: scratchId,
                    path: docPath,
                    workDir,
                    out,
                    timeoutMs: Math.min(RETRY_TIMEOUT_MS, remaining()),
                });
            }
            let xml;
            try {
                xml = await readFile(out, "utf8");
            } catch (err) {
                // The host reported success but produced no markup. Raw and
                // untyped this reads as a missing-file error about a path the
                // caller never named.
                throw new ReadError(
                    "document_unreadable",
                    `Word reported success but wrote no structure for ${path.basename(docPath)}: ${err?.message ?? err?.code}`,
                );
            }
            return { xml, meta: result };
        } finally {
            await rm(workDir, { recursive: true, force: true }).catch(() => {});
        }
    }

    /**
     * Reads the original's bytes for a token, translating filesystem errnos
     * into the typed vocabulary.
     *
     * This is the *first* thing in a read that touches the original, before
     * Word is involved at all, so it is where an exclusive lock surfaces. It
     * used to surface as a raw `EBUSY` from the stream -- not a `ReadError`,
     * not even `word_error` -- which escaped the whole stack and left a caller
     * unable to distinguish "locked" from any other failure.
     *
     * Worth knowing why a lock reaches here at all: Word itself opens a
     * document with `FileShare::Read`, so a document open in Word *can* still
     * be read and copied -- which is precisely why the copy-based read works
     * against a document the user has open. Only a stricter holder
     * (`FileShare::None`) blocks us, and that is a genuinely different
     * situation the caller needs told apart.
     */
    async #tokenOf(docPath) {
        try {
            return await this.#readToken(docPath);
        } catch (err) {
            const errno = err?.code;
            if (errno === "ENOENT") {
                throw new ReadError("file_not_found", `No such file: ${docPath}`);
            }
            if (errno === "EBUSY" || errno === "EACCES" || errno === "EPERM") {
                throw new ReadError(
                    "file_locked",
                    `${path.basename(docPath)} is locked by another process and cannot be read. ` +
                        `A document open in Word can still be read; this is a stricter lock than Word's own.`,
                );
            }
            throw new ReadError("document_unreadable", `Could not read ${path.basename(docPath)}: ${err?.message ?? errno}`);
        }
    }

    /**
     * Returns the structure map and the revision token for a document.
     *
     * The token is taken before the read and confirmed after it. A mismatch
     * means the file changed while we were reading, which would hand back a map
     * that addresses a document that no longer exists — so the read is retried
     * once and then refused.
     */
    async read(docPath, { limit = 0, offset = 0 } = {}) {
        const deadline = Date.now() + READ_BUDGET_MS;
        let info;
        try {
            info = await stat(docPath);
        } catch {
            throw new ReadError("file_not_found", `No such file: ${docPath}`);
        }
        if (!info.isFile()) throw new ReadError("file_not_found", `Not a file: ${docPath}`);

        for (let attempt = 0; attempt < 2; attempt++) {
            const before = await this.#tokenOf(docPath);
            const { xml, meta } = await this.#fetchMarkup(docPath, deadline);
            const after = await this.#tokenOf(docPath);

            if (before !== after) {
                this.#log(`read_document: ${path.basename(docPath)} changed during the read`);
                continue;
            }

            const map = buildStructureMap(xml, { limit, offset });
            return {
                path: docPath,
                name: meta.name ?? path.basename(docPath),
                revisionToken: before,
                // False means another process holds a write handle, so an edit
                // citing this token would be refused. Reported here so the
                // agent learns it before composing one.
                writable: meta.writable !== false,
                sizeBytes: Number(meta.sizeBytes ?? info.size),
                modifiedIso: meta.modifiedIso ?? info.mtime.toISOString(),
                title: meta.title || null,
                author: meta.author || null,
                markupBytes: Number(meta.bytes ?? 0),
                ...map,
            };
        }

        throw new ReadError(
            "document_changed_during_read",
            `${path.basename(docPath)} kept changing while it was being read. Nothing else should be writing to it; try again once it is settled.`,
        );
    }
}
