// The XML reader. Office-free.
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";

import { attr, childNamed, childrenNamed, findFirst, localNameOf, parseXml, textOf, XmlError } from "../../src/word/ooxml.mjs";

const first = (root, local) => findFirst(root, local);

test("elements, attributes and text round-trip", () => {
    const root = parseXml(`<w:p w:rsidR="00A1"><w:r><w:t>hello</w:t></w:r></w:p>`);
    const p = first(root, "p");
    assert.equal(p.name, "w:p");
    assert.equal(p.local, "p");
    assert.equal(attr(p, "rsidR"), "00A1");
    assert.equal(textOf(p), "hello");
});

test("a self-closing tag has no children and does not swallow its siblings", () => {
    const root = parseXml(`<w:r><w:tab/><w:t>after</w:t></w:r>`);
    const run = first(root, "r");
    assert.equal(run.children.length, 2);
    assert.equal(childNamed(run, "tab").children.length, 0);
    assert.equal(textOf(run), "after");
});

test("a '>' inside a quoted attribute value does not end the tag", () => {
    // indexOf(">") would truncate here and reparent everything after it.
    const root = parseXml(`<w:p w:note="a > b"><w:t>text</w:t></w:p>`);
    assert.equal(attr(first(root, "p"), "note"), "a > b");
    assert.equal(textOf(first(root, "p")), "text");
});

test("single-quoted attribute values are read", () => {
    assert.equal(attr(first(parseXml(`<w:t w:val='x y'/>`), "t"), "val"), "x y");
});

test("named, decimal and hex entities are decoded in text and attributes", () => {
    const root = parseXml(`<w:t w:x="&lt;&amp;&gt;">a &amp; b &#233; &#xDC; &quot;q&quot; &apos;</w:t>`);
    const t = first(root, "t");
    assert.equal(textOf(t), `a & b é Ü "q" '`);
    assert.equal(attr(t, "x"), "<&>");
});

test("an unknown entity is left alone rather than dropped", () => {
    assert.equal(textOf(first(parseXml(`<w:t>a &frobnicate; b</w:t>`), "t")), "a &frobnicate; b");
});

test("comments, processing instructions and a doctype are skipped", () => {
    const root = parseXml(
        `<?xml version="1.0"?><?mso-application progid="Word.Document"?><!DOCTYPE x><!-- <w:t>ignored</w:t> --><w:body><w:t>real</w:t></w:body>`,
    );
    assert.equal(textOf(first(root, "body")), "real");
    assert.equal(first(root, "t").children.length, 1);
});

test("CDATA content is taken literally", () => {
    assert.equal(textOf(first(parseXml(`<w:t><![CDATA[a & <b>]]></w:t>`), "t")), "a & <b>");
});

test("whitespace inside a preserved run is kept", () => {
    const root = parseXml(`<w:t xml:space="preserve">  spaced  </w:t>`);
    assert.equal(textOf(first(root, "t")), "  spaced  ");
    assert.equal(attr(first(root, "t"), "space"), "preserve");
});

test("lookups ignore the namespace prefix", () => {
    // Word writes `w:`, but the prefix is a serialization detail and binding to
    // it would fail on markup that is still perfectly valid.
    const root = parseXml(`<x:p xmlns:x="..." x:rsidR="7"><x:t>text</x:t></x:p>`);
    assert.equal(first(root, "p").name, "x:p");
    assert.equal(attr(first(root, "p"), "rsidR"), "7");
    assert.equal(textOf(first(root, "t")), "text");
    assert.equal(localNameOf("x:p"), "p");
    assert.equal(localNameOf("p"), "p");
});

test("childNamed and childrenNamed look one level down, findFirst looks all the way", () => {
    const root = parseXml(`<w:body><w:p><w:t>a</w:t></w:p><w:p><w:t>b</w:t></w:p></w:body>`);
    const body = first(root, "body");
    assert.equal(childrenNamed(body, "p").length, 2);
    assert.equal(childNamed(body, "t"), null, "t is a grandchild, not a child");
    assert.equal(textOf(findFirst(body, "t")), "a");
});

test("a closing tag with no match does not reparent the rest of the document", () => {
    const root = parseXml(`<w:body><w:p><w:t>a</w:t></w:foreign></w:p><w:p><w:t>b</w:t></w:p></w:body>`);
    assert.equal(childrenNamed(first(root, "body"), "p").length, 2);
});

test("empty or non-string input is a typed error", () => {
    for (const bad of ["", "   ", null, undefined, 42, {}]) {
        assert.throws(
            () => parseXml(bad),
            (err) => err instanceof XmlError && err.code === "invalid_xml",
            `expected XmlError for ${JSON.stringify(bad)}`,
        );
    }
});

test("an unterminated tag is reported, not silently truncated", () => {
    assert.throws(() => parseXml(`<w:body><w:p `), (err) => err instanceof XmlError);
});

test("attr and textOf are null-safe", () => {
    assert.equal(attr(null, "val"), null);
    assert.equal(textOf(null), "");
    assert.deepEqual(childrenNamed(null, "p"), []);
    assert.equal(findFirst(null, "p"), null);
});

test("a large document parses in linear time", () => {
    // Guards against a quadratic scan: the real input is a megabyte of markup
    // on a single line, and O(n^2) there would be minutes rather than millis.
    const body = `<w:p><w:r><w:t>paragraph text here</w:t></w:r></w:p>`.repeat(20_000);
    const started = Date.now();
    const root = parseXml(`<w:body>${body}</w:body>`);
    assert.equal(childrenNamed(first(root, "body"), "p").length, 20_000);
    assert.ok(Date.now() - started < 5_000, `parsing took ${Date.now() - started}ms`);
});
