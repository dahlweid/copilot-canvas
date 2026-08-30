// The vendored pdf.js: what is committed, and that it reassembles.
//
// Office-free, so this runs on ubuntu-latest. Everything here derives from
// `pdfjs.manifest.json` rather than restating a byte count -- a hardcoded list
// has drifted from its source three times in this repo, and a test that carries
// its own copy of the answer stops testing the thing it names.
//
// `tools/vendor-pdfjs.mjs` is driven through `tool-bridge.mjs` rather than
// imported: this file is inside the extension folder and an import reaching up
// into `tools/` is exactly the C3 escape the validator rejects.

import { test, after } from "node:test";
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdir, mkdtemp, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { joinWorker, readManifest, VendorAssetError, VENDOR_DIR } from "../../src/vendor-assets.mjs";
import { callTool, readToolValue, cleanupBridge } from "./tool-bridge.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");
const VENDOR_TOOL = path.join(REPO, "tools", "vendor-pdfjs.mjs");
const VALIDATOR = path.join(REPO, "tools", "validate-extensions.mjs");
const APP_CSS = path.join(REPO, ".github", "extensions", "office-canvas", "src", "ui", "app.css");

const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");

after(cleanupBridge);

async function withTempVendor(files) {
    const dir = await mkdtemp(path.join(tmpdir(), "vendor-test-"));
    for (const [name, contents] of Object.entries(files)) {
        await writeFile(path.join(dir, name), contents);
    }
    return dir;
}

// --- what is committed ------------------------------------------------------

test("the committed manifest describes the files that are actually on disk", async () => {
    const manifest = await readManifest();
    for (const entry of manifest.files) {
        const body = await readFile(path.join(VENDOR_DIR, entry.name));
        assert.equal(body.byteLength, entry.bytes, `${entry.name} byte count`);
        assert.equal(sha256(body), entry.sha256, `${entry.name} digest`);
    }
    for (const part of manifest.worker.parts) {
        const body = await readFile(path.join(VENDOR_DIR, part.name));
        assert.equal(body.byteLength, part.bytes, `${part.name} byte count`);
    }
});

test("nothing is committed under src/ui/vendor that the manifest does not name", async () => {
    // An orphan left by a version bump that produced fewer parts is invisible:
    // no route serves it, no manifest lists it, and it counts against the total
    // install budget forever.
    const manifest = await readManifest();
    const named = new Set([
        "pdfjs.manifest.json",
        manifest.license.name,
        ...manifest.files.map((f) => f.name),
        ...manifest.worker.parts.map((p) => p.name),
    ]);
    const onDisk = await readdir(VENDOR_DIR);
    assert.deepEqual(
        onDisk.filter((name) => !named.has(name)),
        [],
    );
});

test("the parts are not named .mjs, which the validator would try to parse", async () => {
    const manifest = await readManifest();
    for (const part of manifest.worker.parts) {
        assert.ok(!part.name.endsWith(".mjs"), `${part.name} would be handed to 'node --check'`);
    }
});

test("every committed part decodes as UTF-8 on its own", async () => {
    // The install envelope decodes each file separately with {fatal: true}, so a
    // boundary inside a multi-byte sequence fails the gate rather than the
    // viewer. This is the same check, close to the thing it constrains.
    const manifest = await readManifest();
    const decoder = new TextDecoder("utf-8", { fatal: true });
    for (const part of manifest.worker.parts) {
        const body = await readFile(path.join(VENDOR_DIR, part.name));
        assert.doesNotThrow(() => decoder.decode(body), `${part.name} must decode standalone`);
    }
});

// --- reassembly -------------------------------------------------------------

test("the worker reassembles byte-exactly from the committed parts", async () => {
    const manifest = await readManifest();
    const { buffer, etag } = await joinWorker();
    assert.equal(buffer.byteLength, manifest.worker.bytes);
    assert.equal(sha256(buffer), manifest.worker.sha256);
    assert.equal(etag, `"${manifest.worker.sha256}"`);
});

test("a part missing from disk fails rather than yielding a short worker", async () => {
    const manifest = await readManifest();
    const first = manifest.worker.parts[0].name;
    const dir = await withTempVendor({
        "pdfjs.manifest.json": JSON.stringify(manifest),
        [first]: await readFile(path.join(VENDOR_DIR, first)),
    });
    try {
        await assert.rejects(joinWorker(dir), (err) => {
            assert.ok(err instanceof VendorAssetError);
            assert.equal(err.code, "vendor_incomplete");
            return true;
        });
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

test("parts read out of order are caught, even though the total size is right", async () => {
    // The discriminating case for the digest check. Two parts are exactly the
    // same size, so swapping them leaves every byte count correct and produces a
    // worker of the right length and complete nonsense. A test that merely
    // truncated a part would pass against a size-only check and prove nothing.
    const manifest = await readManifest();
    const [first, second] = manifest.worker.parts;
    assert.equal(first.bytes, second.bytes, "this test is only meaningful if two parts are the same size");

    const files = { "pdfjs.manifest.json": JSON.stringify(manifest) };
    for (const part of manifest.worker.parts) {
        files[part.name] = await readFile(path.join(VENDOR_DIR, part.name));
    }
    [files[first.name], files[second.name]] = [files[second.name], files[first.name]];

    const dir = await withTempVendor(files);
    try {
        await assert.rejects(joinWorker(dir), (err) => {
            assert.equal(err.code, "vendor_corrupt");
            return true;
        });
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

test("concatenation follows manifest order, not the order filenames sort in", async () => {
    // Three parts today, so `part0 part1 part2` sorts correctly by accident and
    // nothing here could ever fail. Eleven parts is the discriminating case:
    // lexicographically `part10` lands between `part1` and `part2`.
    const parts = Array.from({ length: 11 }, (_, i) => Buffer.from(`[${i}]`, "utf8"));
    const whole = Buffer.concat(parts);
    const names = parts.map((_, i) => `w.part${i}`);
    assert.notDeepEqual([...names].sort(), names, "fixture must actually distinguish the two orderings");

    const files = {};
    parts.forEach((part, i) => {
        files[names[i]] = part;
    });
    files["pdfjs.manifest.json"] = JSON.stringify({
        files: [],
        worker: {
            name: "w",
            bytes: whole.byteLength,
            sha256: sha256(whole),
            parts: names.map((name, i) => ({ name, bytes: parts[i].byteLength })),
        },
    });

    const dir = await withTempVendor(files);
    try {
        const { buffer } = await joinWorker(dir);
        assert.equal(buffer.toString("utf8"), "[0][1][2][3][4][5][6][7][8][9][10]");
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

test("a part name in the manifest is read from the vendor directory, not through it", async () => {
    // Measures the mechanism rather than the outcome. Asserting only that a
    // traversing name *throws* proves nothing: the digest check throws too, so
    // that test passed with the basenaming removed entirely.
    //
    // Instead the same basename exists in both places with different contents,
    // and the manifest describes the inner one. Reading through the traversal
    // yields the outer bytes and fails the digest; basenaming yields the inner
    // bytes and succeeds. Only one of those is a passing join.
    const parent = await mkdtemp(path.join(tmpdir(), "vendor-escape-"));
    try {
        const dir = path.join(parent, "vendor");
        await mkdir(dir);
        const inside = Buffer.from("INSIDE");
        await writeFile(path.join(parent, "decoy.part"), "OUTSIDE");
        await writeFile(path.join(dir, "decoy.part"), inside);
        await writeFile(
            path.join(dir, "pdfjs.manifest.json"),
            JSON.stringify({
                files: [],
                worker: {
                    name: "w",
                    bytes: inside.byteLength,
                    sha256: sha256(inside),
                    parts: [{ name: "../decoy.part", bytes: inside.byteLength }],
                },
            }),
        );

        const { buffer } = await joinWorker(dir);
        assert.equal(buffer.toString("utf8"), "INSIDE");
    } finally {
        await rm(parent, { recursive: true, force: true });
    }
});

test("a traversing part name that resolves nowhere fails as an incomplete vendoring", async () => {
    const dir = await withTempVendor({
        "pdfjs.manifest.json": JSON.stringify({
            files: [],
            worker: { name: "w", bytes: 1, sha256: "x", parts: [{ name: "../../../etc/hosts", bytes: 1 }] },
        }),
    });
    try {
        await assert.rejects(joinWorker(dir), (err) => err instanceof VendorAssetError && err.code === "vendor_incomplete");
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

test("a manifest with no parts is refused instead of serving nothing", async () => {
    const dir = await withTempVendor({
        "pdfjs.manifest.json": JSON.stringify({ worker: { parts: [] }, files: [] }),
    });
    try {
        await assert.rejects(joinWorker(dir), (err) => err.code === "vendor_invalid");
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

test("an absent manifest names the script that produces it", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "vendor-empty-"));
    try {
        await assert.rejects(readManifest(dir), (err) => {
            assert.equal(err.code, "vendor_missing");
            assert.match(err.message, /tools\/vendor-pdfjs\.mjs/);
            return true;
        });
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

// --- the splitter -----------------------------------------------------------

test("splitUtf8Safe never cuts inside a multi-byte sequence", async () => {
    // Every boundary is forced into a multi-byte character: the input is made
    // entirely of 4-byte astral characters, so a naive fixed-width split lands
    // mid-sequence at three offsets out of four. A test over ASCII would agree
    // with a broken splitter and measure nothing.
    const buffer = Buffer.from("\u{1F600}".repeat(500), "utf8");
    const decoder = new TextDecoder("utf-8", { fatal: true });
    for (const target of [7, 13, 64, 999]) {
        const parts = await callTool(VENDOR_TOOL, "splitUtf8Safe", [buffer, target]);
        assert.ok(parts.length > 1, `target ${target} must actually split`);
        for (const part of parts) {
            assert.ok(part.byteLength <= target, `target ${target} produced an oversized part`);
            assert.doesNotThrow(() => decoder.decode(part), `target ${target} produced an undecodable part`);
        }
        assert.deepEqual(Buffer.concat(parts), buffer, `target ${target} did not round-trip`);
    }
});

test("splitUtf8Safe prefers a newline boundary when one is near", async () => {
    // Both branches exist, and only the walk-back one fires on the real worker,
    // which is a single enormous line. Without this the newline branch is dead
    // code that no test would notice being deleted.
    const parts = await callTool(VENDOR_TOOL, "splitUtf8Safe", [Buffer.from("aaaa\nbbbbbbbbbb", "utf8"), 8]);
    assert.equal(parts[0].toString("utf8"), "aaaa\n");
});

test("splitUtf8Safe reproduces the committed part sizes from the committed worker", async () => {
    const manifest = await readManifest();
    const { buffer } = await joinWorker();
    const target = await readToolValue(VENDOR_TOOL, "PART_TARGET_BYTES");
    const parts = await callTool(VENDOR_TOOL, "splitUtf8Safe", [buffer, target]);
    assert.deepEqual(
        parts.map((p) => p.byteLength),
        manifest.worker.parts.map((p) => p.bytes),
    );
});

// --- the text-layer stylesheet ---------------------------------------------

test("the rule extractor takes a whole nested block and stops at its end", async () => {
    const css = ".a{x:1}\n.textLayer{y:2; span{z:3} &.on{w:4}}\n.b{q:5}";
    const rule = await callTool(VENDOR_TOOL, "extractRule", [css, ".textLayer"]);
    assert.equal(rule, ".textLayer{y:2; span{z:3} &.on{w:4}}");
});

test("the extracted text-layer CSS uses every property the bundle sets on it", async () => {
    // The bundle sets these inline, per span; the stylesheet is what turns them
    // into position and size. Ship one without the other and the text layer is
    // present, invisible and unselectable -- while the PDF itself still renders
    // perfectly, so nothing else in the viewer would notice.
    const contract = await readToolValue(VENDOR_TOOL, "TEXT_LAYER_CONTRACT");
    assert.ok(contract.length > 0);
    const css = await readFile(path.join(VENDOR_DIR, "pdf-text-layer.css"), "utf8");
    const bundle = await readFile(path.join(VENDOR_DIR, "pdf.min.mjs"), "utf8");
    for (const property of contract) {
        assert.ok(bundle.includes(property), `pdf.min.mjs should set ${property}`);
        assert.ok(css.includes(property), `pdf-text-layer.css should use ${property}`);
    }
});

test("the page CSS defines the rounding steps pdf.js divides the layer by", async () => {
    // setLayerDimensions writes `round(down, var(--total-scale-factor) * Npx,
    // var(--scale-round-x))`. An undefined --scale-round-x makes the whole
    // declaration invalid, the layer keeps its intrinsic size, and every text
    // span lands at the wrong percentage of it.
    const bundle = await readFile(path.join(VENDOR_DIR, "pdf.min.mjs"), "utf8");
    const appCss = await readFile(APP_CSS, "utf8");
    for (const property of ["--scale-round-x", "--scale-round-y"]) {
        assert.ok(bundle.includes(property), `pdf.min.mjs should read ${property}`);
        assert.match(appCss, new RegExp(`${property}\\s*:`), `app.css should define ${property}`);
    }
});

// --- attribution (Apache-2.0 §4) --------------------------------------------

test("the vendored tree ships the Apache-2.0 licence its notices point at", async () => {
    // §4(a): recipients of the Work get a copy of the licence. The extension is
    // delivered by copying one folder, so "recipient" means whoever has that
    // folder -- which is why the licence lives inside it and not only at the
    // repo root. Digest-checked like everything else here, because a truncated
    // or half-written licence file would still satisfy a mere existence check.
    const manifest = await readManifest();
    assert.equal(manifest.license.spdx, "Apache-2.0");
    const body = await readFile(path.join(VENDOR_DIR, manifest.license.name));
    assert.equal(body.byteLength, manifest.license.bytes, "licence byte count");
    assert.equal(sha256(body), manifest.license.sha256, "licence digest");

    const markers = await readToolValue(VENDOR_TOOL, "APACHE_LICENSE_MARKERS");
    const text = body.toString("utf8");
    for (const marker of markers) {
        assert.ok(text.includes(marker), `the vendored licence should contain '${marker}'`);
    }
});

test("the extracted stylesheet retains Mozilla's copyright and licence notice", async () => {
    // §4(c): retain the copyright and attribution notices from the Source form.
    // Upstream web/pdf_viewer.css opens with one, and the first version of the
    // vendoring script cut the .textLayer rule out from underneath it -- so the
    // committed file carried no attribution at all. Asserting on the committed
    // artifact is what makes that visible; the generator test below is what
    // stops the next version bump undoing it.
    const css = await readFile(path.join(VENDOR_DIR, "pdf-text-layer.css"), "utf8");
    assert.match(css, /Copyright \d{4} Mozilla Foundation/);
    assert.match(css, /Licensed under the Apache License, Version 2\.0/);

    // The pointer has to resolve, or it is decoration. Read from the CSS rather
    // than restated, so renaming the licence file fails here instead of leaving
    // a header naming a file nobody ships.
    const named = css.match(/licence text is in (\S+),/)?.[1];
    assert.ok(named, "the header should name the licence file");
    assert.ok(
        (await readdir(VENDOR_DIR)).includes(named),
        `pdf-text-layer.css points at ${named}, which is not in the vendor directory`,
    );
});

test("the generator emits the upstream notice ahead of the rule it extracts", async () => {
    // The load-bearing half. `tools/vendor-pdfjs.mjs` rewrites this whole
    // directory on a version bump, so a stylesheet fixed by hand is undone the
    // next time pdf.js moves. This drives the generator's own emitter, with no
    // network, and fails if the notice stops coming out the other side.
    const notice = "/* Copyright 2031 Example Foundation\n * Licensed under the Apache License, Version 2.0\n */";
    const source = `${notice}\n.messageBar{a:1}\n.textLayer{b:2}\n`;

    const css = await callTool(VENDOR_TOOL, "buildTextLayerCss", [source, "9.9.9"]);
    assert.ok(css.startsWith(notice), "the upstream notice must lead the file");
    assert.ok(css.includes(".textLayer{b:2}"), "the rule must still be extracted");
    assert.ok(!css.includes(".messageBar"), "only the .textLayer rule is taken");
    assert.ok(css.includes("pdfjs-dist@9.9.9"), "§4(b): the file must state what was changed, and from where");
});

test("the generator refuses to emit a stylesheet whose source carries no notice", async () => {
    // Without this the emitter degrades quietly: an upstream file with no
    // leading comment would produce exactly the unattributed stylesheet this
    // whole section exists to prevent, and nothing would say so.
    for (const source of [".textLayer{b:2}\n", "/* just a note, no attribution */\n.textLayer{b:2}\n"]) {
        await assert.rejects(
            callTool(VENDOR_TOOL, "buildTextLayerCss", [source, "9.9.9"]),
            (err) => {
                // On the error, not merely on rejection: the bridge rejects for
                // any non-zero exit, so a typo in the tool name would otherwise
                // pass this test while proving nothing.
                assert.match(`${err.stderr ?? ""}`, /notice|copyright/i);
                return true;
            },
            JSON.stringify(source),
        );
    }
});

test("no other tracked vendored or generated artifact is missing its notice", async () => {
    // Verified against the tree rather than by eye. Every file under the vendor
    // directory must either be one of ours (the manifest), or carry an Apache
    // notice, or be a continuation part of a file that does. The worker split is
    // the reason for that last clause: parts 1 and 2 are the middle of one file
    // whose notice is in part 0, and demanding a notice in each would be asking
    // for something that would corrupt them.
    const manifest = await readManifest();
    const ours = new Set(["pdfjs.manifest.json", manifest.license.name]);
    const continuation = new Set(manifest.worker.parts.slice(1).map((p) => p.name));

    const missing = [];
    for (const name of await readdir(VENDOR_DIR)) {
        if (ours.has(name) || continuation.has(name)) continue;
        const text = await readFile(path.join(VENDOR_DIR, name), "utf8");
        if (!/Apache License, Version 2\.0/.test(text) || !/Mozilla Foundation/.test(text)) missing.push(name);
    }
    assert.deepEqual(missing, [], "every vendored file must carry its Apache-2.0 attribution");
});

// --- the gate itself --------------------------------------------------------

test("a mid-sequence split is rejected by the real install-envelope check", async () => {
    // Fault injection against the gate CI runs. The committed parts are shaped
    // to avoid exactly this, so without an injected fault nothing shows the gate
    // can still see it.
    const root = await mkdtemp(path.join(tmpdir(), "envelope-"));
    const extDir = path.join(root, ".github", "extensions", "broken");
    await mkdir(extDir, { recursive: true });
    await writeFile(path.join(extDir, "extension.mjs"), "export default {};\n");
    await writeFile(
        path.join(extDir, "copilot-extension.json"),
        JSON.stringify({ name: "broken", version: 1, productVersion: "1.0.0" }),
    );
    // The lead byte of a two-byte character with its continuation byte in the
    // sibling part -- what a naive split of a UTF-8 file produces.
    await writeFile(path.join(extDir, "half.part0"), Buffer.from([0xc3]));
    await writeFile(path.join(extDir, "half.part1"), Buffer.from([0xa9]));

    // Staged so the validator scans this tree and not the real repo.
    const stagedTools = path.join(root, "tools");
    await mkdir(stagedTools, { recursive: true });
    await writeFile(path.join(stagedTools, "validate-extensions.mjs"), await readFile(VALIDATOR));

    try {
        await assert.rejects(
            execFileAsync(process.execPath, [path.join(stagedTools, "validate-extensions.mjs")]),
            (err) => {
                assert.match(`${err.stdout ?? ""}${err.stderr ?? ""}`, /not valid UTF-8/);
                return true;
            },
        );
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});
