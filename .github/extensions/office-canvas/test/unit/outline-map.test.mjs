import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
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

test("outline limit defaults only for non-finite route values and otherwise rejects invalid input", () => {
    assert.deepEqual(extractOutlineEntries(markup, { limit: Number.NaN }), extractOutlineEntries(markup));
    assert.deepEqual(extractOutlineEntries(markup, { limit: Infinity }), extractOutlineEntries(markup));
    assert.throws(
        () => extractOutlineEntries(markup, { limit: -1 }),
        (error) => error instanceof OutlineError && error.code === "invalid_request",
    );
    assert.throws(
        () => extractOutlineEntries(markup, { limit: 1.5 }),
        (error) => error instanceof OutlineError && error.code === "invalid_request",
    );
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
    let markupCalls = 0;
    let markupPath;
    await writeFile(docPath, "fixture");

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            await mkdir(workDir, { recursive: true });
            return { docId, pageCount: 2, wordCount: 3, sizeBytes: 7, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async outlineMarkup({ out }) {
            markupCalls++;
            markupPath = out;
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
        assert.equal(markupCalls, 1);
        assert.equal(existsSync(markupPath), false);
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

test("refresh cannot pair old outline markup with positions from a reopened document", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "outline-map-"));
    const docPath = path.join(root, "interleaved.docx");
    const markupFor = (heading) => flatOpc([paragraph(heading, { styleId: "berschrift1" })]);
    let signalMarkupWritten;
    let releaseMarkup;
    const markupWritten = new Promise((resolve) => {
        signalMarkupWritten = resolve;
    });
    const markupReleased = new Promise((resolve) => {
        releaseMarkup = resolve;
    });
    let generation = "old";
    await writeFile(docPath, "old");

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            await mkdir(workDir, { recursive: true });
            generation = generation === "old" ? "old" : "new";
            return { docId, pageCount: 1, wordCount: 1, sizeBytes: 7, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async closeDocument() {
            generation = "new";
            return { closed: true };
        },
        async outlineMarkup({ out }) {
            await writeFile(out, markupFor(generation));
            signalMarkupWritten();
            await markupReleased;
        },
        async outlinePositions({ wordIndices }) {
            return {
                positions: wordIndices.map((wordIndex) => ({
                    wordIndex,
                    start: generation === "old" ? 10 : 20,
                    page: 1,
                })),
            };
        },
    };

    try {
        await cache.open(docPath);
        const outline = cache.outline(docPath);
        await markupWritten;
        await writeFile(docPath, "new document");
        const refresh = cache.refresh(docPath);
        releaseMarkup();

        assert.deepEqual(await outline, {
            headings: [{ level: 1, text: "old", start: 10, page: 1 }],
            count: 1,
        });
        await refresh;
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("concurrent outlines serialize their complete markup-to-position handshakes", async () => {
    const root = await mkdtemp(path.join(os.tmpdir(), "outline-map-"));
    const docPath = path.join(root, "concurrent.docx");
    let signalFirstMarkup;
    let signalSecondMarkup;
    let releaseFirstMarkup;
    let releaseSecondMarkup;
    const firstMarkup = new Promise((resolve) => {
        signalFirstMarkup = resolve;
    });
    const secondMarkup = new Promise((resolve) => {
        signalSecondMarkup = resolve;
    });
    const firstRelease = new Promise((resolve) => {
        releaseFirstMarkup = resolve;
    });
    const secondRelease = new Promise((resolve) => {
        releaseSecondMarkup = resolve;
    });
    let calls = 0;
    let active = 0;
    let maximumActive = 0;
    await writeFile(docPath, "fixture");

    const cache = new RenderCache({ cacheRoot: path.join(root, "cache") });
    cache.host = {
        async openDocument({ docId, workDir }) {
            await mkdir(workDir, { recursive: true });
            return { docId, pageCount: 1, wordCount: 1, sizeBytes: 7, modifiedIso: "2026-01-01T00:00:00.000Z" };
        },
        async outlineMarkup({ out }) {
            calls++;
            active++;
            maximumActive = Math.max(maximumActive, active);
            try {
                if (calls === 1) {
                    signalFirstMarkup();
                    await firstRelease;
                } else if (calls === 2) {
                    signalSecondMarkup();
                    await secondRelease;
                }
                await writeFile(out, markup);
            } finally {
                active--;
            }
        },
        async outlinePositions({ wordIndices }) {
            return { positions: wordIndices.map((wordIndex) => ({ wordIndex, start: wordIndex, page: 1 })) };
        },
    };

    try {
        await cache.open(docPath);
        const first = cache.outline(docPath);
        await firstMarkup;
        const second = cache.outline(docPath);
        const third = cache.outline(docPath);
        releaseFirstMarkup();
        await secondMarkup;

        assert.equal(calls, 2, "a third outline began before the second completed its handshake");
        assert.equal(maximumActive, 1, "multiple outline markup requests overlapped");
        releaseSecondMarkup();
        await Promise.all([first, second, third]);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});
