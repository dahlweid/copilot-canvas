// A RenderCache that never goes near Word, and that reports being used after
// it was disposed the way the real one does.
//
// Stands in for `src/render-cache.mjs` when the extension itself is under test.
// Two things it has to model, because they are what #61 is made of:
//
//   * Disposal takes real time. `WordHost.dispose` sends `quit` under a 20 s
//     timeout and then waits up to 5 s more for the child to exit, so the real
//     window is up to ~25 s wide. `hold()` makes that window explicit and
//     deterministic instead of a sleep.
//   * The host is dead *before* that window closes -- `#disposed` is set inside
//     `WordHost.dispose`, not after `RenderCache.dispose` resolves -- and a
//     disposed host answers everything with the same sentence. So `disposed` is
//     set first here, and every method then throws that sentence.
//
// Everything else about the extension in these tests is the committed code.

export { DocumentError, normalizeDocPath, supportedList } from "../../src/render-cache.mjs";

import path from "node:path";

/** Every cache built, oldest first, so a test can say *which* one was used. */
export const built = [];

let gate = null;

/**
 * Makes disposals block until the returned function is called.
 * Models the seconds a real Word teardown takes.
 */
export function holdDisposals() {
    let release;
    gate = new Promise((resolve) => {
        release = resolve;
    });
    return () => {
        gate = null;
        release();
    };
}

export function resetCaches() {
    built.length = 0;
    gate = null;
}

let nextId = 1;

export class RenderCache {
    constructor(options = {}) {
        this.id = nextId++;
        this.options = options;
        this.disposed = false;
        /** Method names called on this cache, in order. */
        this.used = [];
        built.push(this);
    }

    #live(what) {
        if (this.disposed) {
            // Word-host.mjs:95, verbatim: this is the sentence the user saw.
            const err = new Error("The Word host has been shut down.");
            err.code = "word_unavailable";
            throw err;
        }
        this.used.push(what);
    }

    get wordVersion() {
        return this.disposed ? null : "16.0";
    }

    async open(docPath) {
        this.#live("open");
        return {
            path: docPath,
            name: path.basename(docPath),
            key: `key-${this.id}`,
            pageCount: 1,
            sizeBytes: 1,
            modifiedIso: new Date(0).toISOString(),
        };
    }

    async pdf(docPath) {
        this.#live("pdf");
        return { file: docPath, key: `key-${this.id}` };
    }

    async refresh(docPath) {
        this.#live("refresh");
        return { path: docPath, name: path.basename(docPath), key: `key-${this.id}`, pageCount: 1, changed: false };
    }

    async close() {
        this.#live("close");
    }

    async readStructure() {
        this.#live("readStructure");
        return { paragraphs: [], paragraphCount: 0, truncated: false };
    }

    async dispose() {
        this.disposed = true;
        if (gate) await gate;
    }

    reap() {
        return false;
    }
}
