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
        /** Document paths passed to path-taking methods, in order (#158). */
        this.docPaths = [];
        built.push(this);
    }

    #live(what) {
        if (this.disposed) {
            // `#ensureStarted` in word-host.mjs, verbatim: this is the sentence
            // the user saw.
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
        this.docPaths.push(docPath);
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

    async readStructure(docPath) {
        this.#live("readStructure");
        this.docPaths.push(docPath);
        return { paragraphs: [], paragraphCount: 0, truncated: false };
    }

    // create/edit/revert exist so the #158 negative controls are genuine
    // refusal->success flips rather than missing-method errors. With the
    // resolver fix in place a *refused* relative path never reaches these --
    // `resolveInputPath` declines before any cache is built -- but when the
    // resolver is broken to prove a control can go red (reverting to cwd
    // resolution, or dropping the field guard), each of these tools must be able
    // to *succeed* against a relative path resolved (wrongly) against cwd.
    // Without them the broken-resolver run would fail with an unrelated "is not
    // a function", which is red for the wrong reason: exactly the defect the
    // control exists to exclude. Each records its docPath the same way
    // `open`/`readStructure` do, and returns the `document.path` shape the
    // edit/revert handlers post-process (`changeRecordFrom` reads no `applied`
    // here, so it yields null and nothing is drawn).
    async createDocument(docPath, _spec) {
        this.#live("createDocument");
        this.docPaths.push(docPath);
        return {
            document: { path: docPath, name: path.basename(docPath), key: `key-${this.id}`, pageCount: 1 },
            structure: { paragraphs: [], paragraphCount: 0, truncated: false },
            revisionToken: `rev-${this.id}`,
        };
    }

    async editDocument(docPath, _intent, _options) {
        this.#live("editDocument");
        this.docPaths.push(docPath);
        return { document: { path: docPath, name: path.basename(docPath), key: `key-${this.id}`, pageCount: 1 } };
    }

    async revertDocument(docPath, _options) {
        this.#live("revertDocument");
        this.docPaths.push(docPath);
        return { document: { path: docPath, name: path.basename(docPath), key: `key-${this.id}`, pageCount: 1 } };
    }

    async dispose() {
        this.disposed = true;
        if (gate) await gate;
    }

    reap() {
        return false;
    }
}
