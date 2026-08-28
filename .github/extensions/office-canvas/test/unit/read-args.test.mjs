// Paging-argument validation for `read_document`.
//
// The defect these exist to catch: the handler defaulted its limit with
// `args?.limit ?? DEFAULT_READ_LIMIT`, and `??` substitutes only for
// null/undefined. `buildStructureMap` treats `limit = 0` as "return
// everything", so `limit: 0` reached that convention and returned the whole
// document -- no error, and the response looked coherent (`truncated: false`),
// which is the failure mode the 300-paragraph default exists to prevent.
//
// The caller is a language model, so 0 is an ordinary input: a remaining-count
// that reached zero, or a guess that 0 means "no preference".

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";

import { DEFAULT_READ_LIMIT, MAX_READ_LIMIT, ReadArgsError, normalizeReadArgs } from "../../src/word/read-args.mjs";
import { buildStructureMap } from "../../src/word/structure-map.mjs";

const here = path.dirname(fileURLToPath(import.meta.url));

const rejects = (args, field) => {
    let err;
    try {
        normalizeReadArgs(args);
    } catch (caught) {
        err = caught;
    }
    assert.ok(err instanceof ReadArgsError, `expected ${JSON.stringify(args)} to be rejected`);
    assert.equal(err.code, "invalid_request", "must carry the typed code the tool layer reports");
    assert.match(err.message, new RegExp(`\`${field}\``), "the message must name the offending field");
    return err;
};

test("an absent limit takes the default", () => {
    assert.deepEqual(normalizeReadArgs({}), { limit: DEFAULT_READ_LIMIT, offset: 0 });
    assert.deepEqual(normalizeReadArgs(undefined), { limit: DEFAULT_READ_LIMIT, offset: 0 });
    // Explicit null is "no preference" too -- it is what a model emits for an
    // optional field it has decided not to use.
    assert.deepEqual(normalizeReadArgs({ limit: null, offset: null }), { limit: DEFAULT_READ_LIMIT, offset: 0 });
});

test("a present limit is honoured, including offset 0", () => {
    assert.deepEqual(normalizeReadArgs({ limit: 5, offset: 0 }), { limit: 5, offset: 0 });
    assert.deepEqual(normalizeReadArgs({ limit: 1 }), { limit: 1, offset: 0 });
    assert.deepEqual(normalizeReadArgs({ limit: MAX_READ_LIMIT }), { limit: MAX_READ_LIMIT, offset: 0 });
});

test("a limit above the declared maximum is rejected", () => {
    // This test previously asserted the opposite -- that 10_000_000 was accepted
    // -- on the argument that slice clamps it anyway, so refusing would refuse a
    // request we can satisfy. That missed that the tool schema *declares*
    // `maximum`, which makes the bound a promise to the caller rather than an
    // internal preference. Unenforced, the contract answered differently
    // depending on whether the host pre-validates: either an upstream rejection
    // in some other shape, or a declaration that is false.
    const err = rejects({ limit: MAX_READ_LIMIT + 1 }, "limit");
    assert.match(err.message, new RegExp(`${MAX_READ_LIMIT} or less`));
    rejects({ limit: 10_000_000 }, "limit");
});

test("the schema declares the same maximum the validator enforces", async () => {
    // The bound is stated twice by necessity -- once as a JSON-schema number the
    // model reads, once as a comparison the runtime makes. Deriving both from
    // MAX_READ_LIMIT is what keeps them one number; this is the test that says
    // so, in the same spirit as the supported-extension list.
    //
    // Scoped to read_document's own schema on purpose: the canvas actions
    // declare numeric maximums of their own, and a file-wide assertion would
    // fail on unrelated code -- claiming a rule over a surface this module does
    // not own.
    const source = await readFile(path.join(here, "..", "..", "extension.mjs"), "utf8");
    const start = source.indexOf('name: "read_document"');
    assert.ok(start > 0, "expected to find the read_document tool");
    const schema = source.slice(start, source.indexOf("\n};", start));

    assert.match(schema, /maximum: MAX_READ_LIMIT/, "the schema must declare the enforced constant");
    assert.doesNotMatch(schema, /maximum: \d+/, "a literal maximum in the schema can drift from the validator");
});

test("limit: 0 is rejected rather than silently meaning everything", () => {
    const err = rejects({ limit: 0 }, "limit");
    assert.match(err.message, /1 or greater/);
    // The hint matters: the caller is a model, and this is what lets it correct
    // itself instead of retrying the same call.
    assert.match(err.message, new RegExp(`default of ${DEFAULT_READ_LIMIT}`));
});

test("limit: 0 would otherwise return every paragraph", () => {
    // Guards the reasoning, not just the branch. If buildStructureMap ever stops
    // treating 0 as "everything", this test fails and the rejection above can be
    // reconsidered -- rather than surviving as a rule whose cause has gone.
    const xml = `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
        <w:p><w:r><w:t>one</w:t></w:r></w:p>
        <w:p><w:r><w:t>two</w:t></w:r></w:p>
        <w:p><w:r><w:t>three</w:t></w:r></w:p>
    </w:body></w:document>`;
    assert.equal(buildStructureMap(xml, { limit: 0 }).returned, 3);
    assert.equal(buildStructureMap(xml, { limit: 2 }).returned, 2);
});

test("a negative or fractional limit is rejected", () => {
    rejects({ limit: -1 }, "limit");
    rejects({ limit: 1.5 }, "limit");
});

test("a non-numeric limit is rejected rather than coerced", () => {
    // The declared parameter type is `integer`. Coercing "300" would make the
    // published contract a lie in the same way the hardcoded `.dotx` list did --
    // a second, looser contract living in the code.
    for (const value of ["300", true, {}, [], NaN, Infinity]) {
        rejects({ limit: value }, "limit");
    }
});

test("offset is validated on the same terms, and 0 is legitimate", () => {
    assert.equal(normalizeReadArgs({ offset: 0 }).offset, 0);
    assert.equal(normalizeReadArgs({ offset: 12 }).offset, 12);
    rejects({ offset: -1 }, "offset");
    rejects({ offset: 2.5 }, "offset");
    rejects({ offset: "0" }, "offset");
});

test("the read_document handler defaults through normalizeReadArgs, not ??", async () => {
    // extension.mjs cannot be imported: it calls joinSession() at module scope.
    // Reading the source still pins the property that matters -- that the
    // handler routes arguments through the validator rather than reintroducing
    // a `??` default, which is the exact line the reviewer found.
    const source = await readFile(path.join(here, "..", "..", "extension.mjs"), "utf8");

    assert.match(source, /normalizeReadArgs\(args\)/, "the handler must validate its paging arguments");
    assert.doesNotMatch(
        source,
        /args\?\.limit\s*\?\?/,
        "`??` does not substitute for 0, which is why limit: 0 returned the whole document",
    );
    assert.doesNotMatch(source, /const DEFAULT_READ_LIMIT/, "the default must be imported, not restated");
});
