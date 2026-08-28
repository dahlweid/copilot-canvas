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
import { createHash } from "node:crypto";
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

/** Errors that mean the host process died rather than the document being bad. */
const isHostFailure = (err) =>
    err?.code === "word_unavailable" ||
    err?.code === "word_timeout" ||
    /No open document with id/i.test(err?.message ?? "");

export class DocumentReader {
    #host;
    #workRoot;
    #log;

    constructor({ host, workRoot, log = () => {} }) {
        this.#host = host;
        this.#workRoot = workRoot;
        this.#log = log;
    }

    async #fetchMarkup(docPath) {
        const scratchId = `read-${shortHash(docPath.toLowerCase())}`;
        const workDir = path.join(this.#workRoot, "read", scratchId);
        const out = path.join(workDir, "structure.xml");
        await mkdir(workDir, { recursive: true });
        try {
            let result;
            try {
                result = await this.#host.structure({ docId: scratchId, path: docPath, workDir, out });
            } catch (err) {
                // A structure read registers no reopen args, so the host's own
                // replay cannot cover it. One retry: the next request starts a
                // fresh host transparently.
                if (!isHostFailure(err)) throw err;
                this.#log(`read_document: retrying after ${err.code ?? "error"}`);
                result = await this.#host.structure({ docId: scratchId, path: docPath, workDir, out });
            }
            return { xml: await readFile(out, "utf8"), meta: result };
        } finally {
            await rm(workDir, { recursive: true, force: true }).catch(() => {});
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
        let info;
        try {
            info = await stat(docPath);
        } catch {
            throw new ReadError("file_not_found", `No such file: ${docPath}`);
        }
        if (!info.isFile()) throw new ReadError("file_not_found", `Not a file: ${docPath}`);

        for (let attempt = 0; attempt < 2; attempt++) {
            const before = await fileRevisionToken(docPath);
            const { xml, meta } = await this.#fetchMarkup(docPath);
            const after = await fileRevisionToken(docPath);

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
