import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

// ADR 0004's accessibility section rests entirely on Word's export being
// tagged: /StructTreeRoot is what pdf.js reads back through getStructTree().
// That guarantee is `DocStructureTags = $true` -- except it is not written
// anywhere in that form. `ExportAsFixedFormat` is called with fourteen
// positional arguments and the parameter names live in a comment above the
// call, so the flag is identified by its *position* and by nothing else.
//
// Grepping for `DocStructureTags` in the host finds only that comment. An
// argument inserted or removed anywhere in the first twelve slots silently
// re-points the flag while leaving the comment reading correctly, and the only
// symptom is an untagged PDF -- which this repo would notice as an
// accessibility regression long after the commit that caused it.
//
// So the ADR must not restate the value. These tests derive it: the name list
// comes from the comment, the index from that list, and the value from the call.
// Office-free -- it is text, not COM -- so it runs on ubuntu-latest.

const HOST_SOURCE = readFileSync(
    fileURLToPath(new URL("../../src/word/word-host.ps1", import.meta.url)),
    "utf8",
);

const CALL = "$doc.ExportAsFixedFormat(";

/**
 * Pulls the export call apart into the parameter names the comment declares and
 * the arguments actually passed, both in source order.
 */
export function exportCallSignature(source) {
    const at = source.indexOf(CALL);
    if (at === -1) return null;

    const linesBefore = source.slice(0, at).split(/\r?\n/);
    const comment = [];
    // -2 skips the (partial) line the call itself starts on.
    for (let i = linesBefore.length - 2; i >= 0; i--) {
        const line = linesBefore[i].trim();
        if (!line.startsWith("#")) break;
        comment.unshift(line.replace(/^#\s?/, ""));
    }

    const names = comment
        .join(" ")
        .split(",")
        .map((n) => n.trim())
        .filter(Boolean);

    const args = [];
    let buffer = "";
    let depth = 1;
    for (let i = at + CALL.length; i < source.length; i++) {
        const ch = source[i];
        if (ch === "(") depth++;
        else if (ch === ")") {
            depth--;
            if (depth === 0) {
                args.push(buffer);
                break;
            }
        }
        if (depth === 1 && ch === ",") {
            args.push(buffer);
            buffer = "";
            continue;
        }
        buffer += ch;
    }

    return { names, args: args.map((a) => a.trim()) };
}

test("the PDF export call is found in the Word host", () => {
    const sig = exportCallSignature(HOST_SOURCE);
    assert.notEqual(sig, null, "ExportAsFixedFormat call not found in word-host.ps1");
    assert.ok(sig.args.length > 0, "export call parsed as having no arguments");
});

test("every positional argument of the PDF export is named by the comment above it", () => {
    const { names, args } = exportCallSignature(HOST_SOURCE);

    // This is the alignment the positional call depends on. If it ever drifts,
    // the index derived below stops meaning what its name says -- and that is
    // exactly the failure the comment cannot reveal on its own.
    assert.equal(
        args.length,
        names.length,
        `the comment names ${names.length} parameters but the call passes ${args.length} arguments`,
    );
});

test("the PDF export asks Word for a tagged document", () => {
    const { names, args } = exportCallSignature(HOST_SOURCE);

    const index = names.indexOf("DocStructureTags");
    assert.notEqual(index, -1, "DocStructureTags is not named in the export call's parameter comment");

    assert.equal(
        args[index],
        "$true",
        `DocStructureTags is argument ${index + 1} and is passed as ${args[index]}, so the export is not tagged`,
    );
});

test("the tagging flag is identified by position, not by name", () => {
    // Guards the premise of the whole file. If the call is ever converted to
    // named arguments, an argument carries an `=` and the position-derived
    // tests above stop meaning anything -- this is the signal to delete them
    // rather than let them keep deriving an index nobody uses.
    const { args } = exportCallSignature(HOST_SOURCE);
    for (const arg of args) {
        assert.ok(
            !arg.includes("="),
            `argument \`${arg}\` is passed by name -- the positional derivation in this file is obsolete`,
        );
    }
});
