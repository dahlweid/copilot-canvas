import { buildStructureMap } from "./structure-map.mjs";

export class OutlineError extends Error {
    constructor(message) {
        super(message);
        this.name = "OutlineError";
        this.code = "outline_mismatch";
    }
}

export function extractOutlineEntries(xml, { limit = 2000 } = {}) {
    if (!Number.isSafeInteger(limit) || limit < 0) {
        throw new OutlineError("Outline limit must be a non-negative safe integer.");
    }

    return buildStructureMap(xml).paragraphs
        .filter((paragraph) => paragraph.headingLevel !== null)
        .slice(0, limit)
        .map(({ headingLevel, text, wordIndex }) => ({
            level: headingLevel,
            text,
            wordIndex,
        }));
}

export function resolveOutlineEntries(entries, positions) {
    const byWordIndex = new Map();
    for (const position of positions) {
        if (!Number.isSafeInteger(position?.wordIndex) || byWordIndex.has(position.wordIndex)) {
            throw new OutlineError("Word returned duplicate or invalid outline positions.");
        }
        byWordIndex.set(position.wordIndex, position);
    }

    const headings = entries.map((entry) => {
        const position = byWordIndex.get(entry.wordIndex);
        if (!position) {
            throw new OutlineError("Word did not return a position for every outline heading.");
        }
        return {
            level: entry.level,
            text: entry.text,
            page: position.page,
            start: position.start,
        };
    });
    return { headings, count: headings.length };
}
