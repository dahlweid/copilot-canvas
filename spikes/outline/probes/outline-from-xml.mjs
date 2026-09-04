import { buildStructureMap } from "../../../.github/extensions/office-canvas/src/word/structure-map.mjs";

let xml = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
    xml += chunk;
});
process.stdin.on("end", () => {
    const { paragraphs } = buildStructureMap(xml);
    const headings = paragraphs
        .filter((paragraph) => paragraph.headingLevel !== null)
        .map(({ headingLevel, text, wordIndex }) => ({ headingLevel, text, wordIndex }));
    process.stdout.write(headings.map((heading) => JSON.stringify(heading)).join("\n"));
});
