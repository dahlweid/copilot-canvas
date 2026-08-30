// Mutation harness for #87. Applies one named mutation to the working tree,
// leaves it applied, and prints what it changed. Restore with `git stash` or by
// re-running with `--restore`, which reverses the same edit rather than asking
// git -- the index here is not a safe backup, because a `git checkout --` of a
// staged-then-edited file silently discards the later edit. That happened once
// during this change and made an unrelated test go red under a mutation that
// could not have caused it.
//
// Not shipped with the extension: this lives under spikes/.

import { readFileSync, writeFileSync } from "node:fs";

const UI = ".github/extensions/office-canvas/src/ui/";

const MUTATIONS = {
    // `.bar-main` stretches to the grid row again -- the state #87 was reported
    // against.
    stretch: {
        file: UI + "app.css",
        from: "\n    align-self: center;\n}\n\n.bar-actions {",
        to: "\n}\n\n.bar-actions {",
    },
    // The mark is nudged with a magic pixel instead of aligned.
    nudge: {
        file: UI + "app.css",
        from: ".word-mark {\n    width: 16px;",
        to: ".word-mark {\n    margin-top: 2px;\n    width: 16px;",
    },
    // The Open in Word button carries a second Word mark again.
    "second-mark": {
        file: UI + "index.html",
        from: '                    <svg\n                        class="icon"\n                        id="openInWordGlyph"',
        to:
            '                    <img class="word-mark" id="openInWordMark" alt="" width="16" height="16" hidden />\n' +
            '                    <svg\n                        class="icon"\n                        id="openInWordGlyph"',
    },
    // app.js wires the mark into the button again, with a fallback to hide.
    rewire: {
        file: UI + "app.js",
        from: "showWordMark({ img: el.docNameMark });",
        to: "showWordMark({ img: el.docNameMark });\nshowWordMark({ img: el.openInWordMark, fallback: el.openInWordGlyph });",
    },
    // word-mark.mjs writes to a second element again.
    fallback: {
        file: UI + "word-mark.mjs",
        from: "export function showWordMark({ img, src = \"/api/word-icon\" }) {",
        to:
            "export function showWordMark({ img, fallback, src = \"/api/word-icon\" }) {\n" +
            "    if (fallback) {\n" +
            "        fallback.hidden = false;\n" +
            "        fallback.addEventListener(\"load\", () => {}, { once: true });\n" +
            "        fallback.src = src;\n" +
            "    }",
    },
};

const [, , name, flag] = process.argv;
const m = MUTATIONS[name];
if (!m) {
    console.error(`usage: node mutate.mjs <${Object.keys(MUTATIONS).join("|")}> [--restore]`);
    process.exit(2);
}

const restore = flag === "--restore";
const [from, to] = restore ? [m.to, m.from] : [m.from, m.to];

const text = readFileSync(m.file, "utf8");

// The checkout here is CRLF and the anchors above are written with \n. Matching
// the literal string silently found nothing, which the guard below reports --
// but a mutation harness that can fail closed is exactly the thing that must
// not fail open, so the newlines are matched rather than assumed.
const eol = text.includes("\r\n") ? "\r\n" : "\n";
const pattern = new RegExp(from.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\n/g, "\\r?\\n"));

if (!pattern.test(text)) {
    // A mutation that quietly does nothing reports a surviving mutant, which is
    // the one result that must never be produced by accident.
    console.error(`FAIL — anchor not found in ${m.file}; nothing ${restore ? "restored" : "mutated"}`);
    process.exit(1);
}
writeFileSync(m.file, text.replace(pattern, to.split("\n").join(eol)));
console.log(`${restore ? "restored" : "applied"} ${name} in ${m.file}`);
