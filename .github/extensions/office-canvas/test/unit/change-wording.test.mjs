// What the change banner says, and when it offers to take the reader there.
//
// The interesting case is the one review round 1 found: a record whose page is
// null. The banner interpolated it and printed "page null", and the "Show me"
// button stayed visible beside it while its handler declined to move -- an
// interface asserting a capability the code does not have, which is the same
// class as an error message naming a cause the code never distinguished.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";

import {
    describeChange,
    describeChangeBanner,
    DESCRIBED_OPS,
    GENERIC_PHRASE,
} from "../../src/ui/change-wording.mjs";

const record = (over = {}) => ({ op: "replace_text", page: 3, locatable: true, text: "hi", ...over });

test("a located change names its page", () => {
    const { text, jumpable } = describeChangeBanner(record());
    assert.equal(text, "Text replaced — page 3");
    assert.equal(jumpable, true);
});

test("an unlocatable change says the page is marked but not highlighted", () => {
    const { text, jumpable } = describeChangeBanner(
        record({ op: "delete_paragraph", locatable: false, text: null }),
    );
    assert.equal(text, "Paragraph deleted — page 3, marked but not highlighted");
    // The page is known, so the offer to go there is one we can keep.
    assert.equal(jumpable, true);
});

test("a record with no page never renders the word null", () => {
    for (const page of [null, undefined, 0, -1, 2.5, NaN, "3"]) {
        const { text } = describeChangeBanner(record({ page }));
        assert.doesNotMatch(
            text,
            /null|undefined|NaN/,
            `page ${JSON.stringify(page)} leaked a placeholder into the banner`,
        );
    }
});

test("a record with no page says so, and is not offered as a destination", () => {
    const { text, jumpable } = describeChangeBanner(record({ page: null }));
    assert.equal(text, "Text replaced — page not reported");
    assert.equal(jumpable, false);
});

test("the unreported-page wording does not promise a marker", () => {
    // With no page there is no candidate, so nothing is marked anywhere. Saying
    // "marked but not highlighted" here would send the reader hunting for a
    // badge that was never drawn.
    const { text } = describeChangeBanner(record({ page: null, locatable: false, text: null }));
    assert.doesNotMatch(text, /marked/);
});

test("the unreported-page wording does not name a cause", () => {
    // All the record preserves is that no usable number arrived. Whether Word
    // declined to report one or the editor never asked is not distinguished, so
    // neither may be named.
    const { text } = describeChangeBanner(record({ page: null }));
    assert.doesNotMatch(text, /Word|could not|failed|unable/i);
});

test("a page is only jumpable when the banner actually names it", () => {
    // The two halves must not be able to disagree: an offer to jump is exactly
    // an assertion that a page number is known.
    for (const page of [null, undefined, 0, -1, NaN, "3", 1, 7]) {
        const { text, jumpable } = describeChangeBanner(record({ page }));
        assert.equal(
            jumpable,
            /page \d+/.test(text),
            `page ${JSON.stringify(page)}: jumpable=${jumpable} disagrees with "${text}"`,
        );
    }
});

test("every described operation reaches the banner", () => {
    // Derived from the table rather than restated, so a new operation cannot be
    // added without this noticing.
    for (const op of DESCRIBED_OPS) {
        const { text } = describeChangeBanner(record({ op }));
        assert.doesNotMatch(text, new RegExp(`^${GENERIC_PHRASE} `), `${op} fell through to generic`);
        assert.match(text, new RegExp(`^${describeChange({ op })} `));
    }
});

test("an unknown operation still produces a usable line", () => {
    const { text } = describeChangeBanner(record({ op: "some_future_op" }));
    assert.equal(text, `${GENERIC_PHRASE} — page 3`);
});
