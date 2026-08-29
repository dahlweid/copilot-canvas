// The stdio protocol is UTF-8 in both directions, and the host has to say so
// twice.
//
// It said it once. `[Console]::OutputEncoding` was set and
// `[Console]::InputEncoding` was not, so Windows PowerShell decoded the UTF-8
// JSON Node writes to its stdin as the OEM codepage. Every non-ASCII character
// an agent sent was silently mojibake'd on the way in (issue #40): `Grüße`
// reached the disk as `Gr├╝├ƒe`, and a canvas search for a non-ASCII term
// returned `count: 0` -- no error, indistinguishable from a genuine miss, which
// is the symptom a user meets first.
//
// The third symptom is the one worth stating precisely, because the obvious
// account of it is wrong and self-defeating. A non-ASCII paragraph becomes
// permanently *uneditable*, and not because the stored text and `expectedText`
// both crossed the same boundary -- one transform applied to both sides would
// make them match. It is because the corruption **compounds**: the read
// direction is faithful, so the agent gets the on-disk mojibake back exactly,
// and sending that in re-decodes the previous result's UTF-8 bytes as OEM a
// second time (measured: `Gr├╝nchen` in, `GrÔö£ÔòØnchen` out). Each crossing
// expands it, so `expectedText` can never equal what is stored, for any number
// of retries.
//
// It survived two merged layers not because anything hid it -- a faithful read
// makes it visible end to end on first look -- but because every fixture in the
// repo was ASCII, including the line in `make-fixture.ps1` labelled "Umlaut
// check:" that contained no umlaut. The integration suites now send real ones,
// but they need Word and cannot run in CI. This file is the half that can: it
// is a source assertion, in the shape of the `MAX_READ_LIMIT` schema guard, and
// it would have caught the whole defect on `ubuntu-latest` without Word
// installed.
//
// Measured by spikes/isolation/probes/probe-console-input-encoding.mjs, which
// drives the same spawn shape three ways: control corrupts, both UTF-8 forms
// are intact, and the setter does not throw on a redirected pipe.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const hostScript = path.join(here, "..", "..", "src", "word", "word-host.ps1");

const rawSource = await readFile(hostScript, "utf8");

/**
 * The script with its comments blanked out, which is what the encoding
 * assertions below run against.
 *
 * A grep over the raw text accepts a **commented-out** setter, and that is not
 * hypothetical: the first form of this file passed cleanly while the fix sat
 * behind a `#`, so it could not have failed for the reason it claims. Deleting
 * the line is only one of the ways it goes away; disabling it is the likelier
 * one during a debugging session, and the likeliest to be committed by accident.
 *
 * Comments are blanked rather than dropped so every offset is preserved -- the
 * ordering test compares positions, and removing text would shift them.
 * Line-comments only: PowerShell's `<# #>` block form does not appear here, and
 * a matcher for a construct the file does not use is untested code in a test.
 */
const source = rawSource.replace(/#[^\n]*/g, (m) => " ".repeat(m.length));

/** Matches an assignment of any UTF-8 encoding to one Console encoding property. */
const assignsUtf8 = (property) =>
    new RegExp(
        String.raw`\[Console\]::${property}\s*=\s*(New-Object\s+(System\.)?Text\.UTF8Encoding|\[(System\.)?Text\.Encoding\]::UTF8)`,
    );

test("the host decodes its stdin as UTF-8", () => {
    // The one line whose absence was issue #40. Written as its own test so a
    // regression names the direction that broke rather than "encoding".
    assert.match(
        source,
        assignsUtf8("InputEncoding"),
        "word-host.ps1 must set [Console]::InputEncoding to UTF-8, or Windows PowerShell decodes the JSON Node sends as the OEM codepage",
    );
});

test("the host encodes its stdout as UTF-8", () => {
    assert.match(
        source,
        assignsUtf8("OutputEncoding"),
        "word-host.ps1 must set [Console]::OutputEncoding to UTF-8",
    );
});

test("both encodings are set before anything reads or writes the console", () => {
    // Not ceremony, and not a style rule about preambles. .NET builds the
    // reader behind `[Console]::In` on first use and caches it; assigning
    // `InputEncoding` afterwards replaces the property while that reader keeps
    // decoding with the old codepage. So a correct pair of assignments moved
    // below the read loop -- the kind of thing a refactor does -- restores the
    // defect while leaving both greps above satisfied.
    const input = source.search(assignsUtf8("InputEncoding"));
    const output = source.search(assignsUtf8("OutputEncoding"));
    const firstUse = source.search(/\[Console\]::(In|Out)\b/);

    assert.ok(firstUse > 0, "expected word-host.ps1 to use [Console]::In or [Console]::Out somewhere");
    // `search` reports -1 for absent, and -1 is less than every offset -- so
    // without these two lines this test passes on a file that sets neither
    // encoding, which is precisely the defect it exists to catch.
    assert.ok(input >= 0, "[Console]::InputEncoding is never assigned a UTF-8 encoding");
    assert.ok(output >= 0, "[Console]::OutputEncoding is never assigned a UTF-8 encoding");
    assert.ok(
        input < firstUse,
        "[Console]::InputEncoding is assigned after the console reader is first used, which leaves that reader on the old codepage",
    );
    assert.ok(
        output < firstUse,
        "[Console]::OutputEncoding is assigned after the console writer is first used",
    );
});

test("the probe backing the encoding claim is cited where the encodings are set", () => {
    // This repo's rule is that a claim about platform behaviour is backed by a
    // probe that was actually run. The citation is what makes the next reader
    // able to re-run it rather than re-derive it, and `tools/check-citations.mjs`
    // is what stops the path going stale.
    //
    // Against `rawSource`, not `source`: the citation *is* a comment, and the
    // comment-blanking the other tests rely on would erase the thing being
    // asserted. The offset is still taken from `source`, so a `#` in a comment
    // cannot move the boundary.
    const preamble = rawSource.slice(0, source.search(/\[Console\]::(In|Out)\b/));
    assert.match(preamble, /probe-console-input-encoding\.mjs/, "the encoding preamble must cite the probe that measured it");
});
