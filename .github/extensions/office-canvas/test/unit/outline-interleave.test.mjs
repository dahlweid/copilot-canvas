// A `refresh` must not land between the two halves of an outline handshake.
//
// The outline is no longer one host command: Word writes the document's markup
// to a file, Node derives the headings from it, and only then does Word resolve
// those paragraphs' positions. Nothing in `WordHost` serialises the two -- the
// cache must retain one document generation across this complete handshake.

import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { RenderCache } from "../../src/render-cache.mjs";
import { flatOpc, paragraph } from "./word-fixtures.mjs";

const GENERATIONS = [
    { text: "Chapter one", start: 100, page: 1 },
    { text: "Chapter two", start: 200, page: 7 },
];

const turns = async (count) => {
    for (let i = 0; i < count; i++) await new Promise((resolve) => setImmediate(resolve));
};

const deferred = () => {
    let resolve;
    const promise = new Promise((r) => {
        resolve = r;
    });
    return { promise, resolve };
};

test("a refresh cannot land between outline markup and outline positions", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "outline-interleave-"));
    const docPath = path.join(root, "interleaved.docx");
    await writeFile(docPath, "generation one");

    const calls = [];
    const markupEntered = deferred();
    const markupReleased = deferred();
    const closeSeen = deferred();
    let generation = 0;

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            calls.push("openDocument");
            await mkdir(workDir, { recursive: true });
            generation = calls.filter((call) => call === "openDocument").length - 1;
            return { docId, pageCount: 1, wordCount: 2, sizeBytes: 14, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async closeDocument() {
            calls.push("closeDocument");
            closeSeen.resolve();
            return { closed: true };
        },
        async outlineMarkup({ out }) {
            calls.push("outlineMarkup");
            await writeFile(out, flatOpc([paragraph(GENERATIONS[generation].text, { styleId: "berschrift1" })]));
            markupEntered.resolve();
            await markupReleased.promise;
        },
        async outlinePositions({ wordIndices }) {
            calls.push("outlinePositions");
            const { start, page } = GENERATIONS[generation];
            return { positions: wordIndices.map((wordIndex) => ({ wordIndex, start, page })) };
        },
    };

    try {
        await cache.open(docPath);
        assert.equal(generation, 0, "the fixture opens as generation one");

        const outline = cache.outline(docPath);
        await markupEntered.promise;

        await writeFile(docPath, "generation two, which is a different length");
        const refresh = cache.refresh(docPath);

        await Promise.race([closeSeen.promise, turns(20)]);
        markupReleased.resolve();

        const result = await outline;
        const refreshed = await refresh;

        assert.deepEqual(
            calls,
            ["openDocument", "outlineMarkup", "outlinePositions", "closeDocument", "openDocument"],
            "the reopen ran inside the outline handshake instead of queueing behind it",
        );
        assert.deepEqual(
            result,
            { headings: [{ level: 1, text: "Chapter one", start: 100, page: 1 }], count: 1 },
            "the outline joined markup and positions from two different generations",
        );

        assert.equal(refreshed.changed, true);
        assert.equal(generation, 1);
        assert.deepEqual(await cache.outline(docPath), {
            headings: [{ level: 1, text: "Chapter two", start: 200, page: 7 }],
            count: 1,
        });
    } finally {
        markupReleased.resolve();
        await rm(root, { recursive: true, force: true });
    }
});
