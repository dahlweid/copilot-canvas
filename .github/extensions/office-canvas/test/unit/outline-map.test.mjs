import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { extractOutlineEntries, OutlineError, resolveOutlineEntries } from "../../src/word/outline-map.mjs";
import { RenderCache } from "../../src/render-cache.mjs";
import { flatOpc, paragraph } from "./word-fixtures.mjs";

const markup = flatOpc([
    paragraph("First", { styleId: "berschrift1" }),
    `<w:tbl><w:tr><w:tc>${paragraph("Inside", { styleId: "berschrift2" })}</w:tc></w:tr></w:tbl>`,
    paragraph("After", { styleId: "Unterkapitel" }),
]);

test("outline entries keep XML document order and Word paragraph indices through a table", () => {
    assert.deepEqual(extractOutlineEntries(markup), [
        { level: 1, text: "First", wordIndex: 1 },
        { level: 2, text: "Inside", wordIndex: 2 },
        { level: 2, text: "After", wordIndex: 4 },
    ]);
    assert.deepEqual(extractOutlineEntries(markup, { limit: 2 }).map((entry) => entry.text), ["First", "Inside"]);
});

test("outline positions join headings by Word index rather than a translatable text value", () => {
    const entries = extractOutlineEntries(markup);
    const result = resolveOutlineEntries(entries, [
        { wordIndex: 4, start: 24, page: 2 },
        { wordIndex: 1, start: 0, page: 1 },
        { wordIndex: 2, start: 8, page: 1 },
    ]);

    assert.deepEqual(result, {
        headings: [
            { level: 1, text: "First", start: 0, page: 1 },
            { level: 2, text: "Inside", start: 8, page: 1 },
            { level: 2, text: "After", start: 24, page: 2 },
        ],
        count: 3,
    });
});

test("outline position mismatches fail rather than silently omitting a heading", () => {
    const entries = extractOutlineEntries(markup);
    assert.throws(
        () => resolveOutlineEntries(entries, [{ wordIndex: 1, start: 0, page: 1 }]),
        (error) => error instanceof OutlineError && error.code === "outline_mismatch",
    );
    assert.throws(
        () =>
            resolveOutlineEntries(entries, [
                { wordIndex: 1, start: 0, page: 1 },
                { wordIndex: 1, start: 4, page: 1 },
            ]),
        (error) => error instanceof OutlineError && error.code === "outline_mismatch",
    );
});

test("an empty outline needs no Word paragraph-position request", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "outline-map-"));
    const docPath = path.join(root, "empty.docx");
    await writeFile(docPath, "fixture");

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            await mkdir(workDir, { recursive: true });
            return { docId, pageCount: 1, wordCount: 0, sizeBytes: 7, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async outlineMarkup({ out }) {
            await writeFile(out, flatOpc([paragraph("Body")]));
        },
        async outlinePositions() {
            throw new Error("an empty outline must not ask Word for positions");
        },
    };

    try {
        await cache.open(docPath);
        assert.deepEqual(await cache.outline(docPath), { headings: [], count: 0 });
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("render cache takes XML once, resolves pages only for the limited headings, and removes its markup", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "outline-map-"));
    const docPath = path.join(root, "example.docx");
    const requested = [];
    await writeFile(docPath, "fixture");

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            await mkdir(workDir, { recursive: true });
            return { docId, pageCount: 2, wordCount: 3, sizeBytes: 7, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async outlineMarkup({ out }) {
            await writeFile(out, markup);
        },
        async outlinePositions({ wordIndices }) {
            requested.push(wordIndices);
            return {
                positions: wordIndices.map((wordIndex) => ({
                    wordIndex,
                    start: wordIndex * 10,
                    page: wordIndex === 4 ? 2 : 1,
                })),
            };
        },
    };

    try {
        await cache.open(docPath);
        const result = await cache.outline(docPath, { limit: 2 });

        assert.deepEqual(requested, [[1, 2]]);
        assert.deepEqual(result, {
            headings: [
                { level: 1, text: "First", start: 10, page: 1 },
                { level: 2, text: "Inside", start: 20, page: 1 },
            ],
            count: 2,
        });
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});
