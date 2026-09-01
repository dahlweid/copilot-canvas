// Fault-injection check for tools/validate-extensions.mjs.
//
// A validator that has never failed is indistinguishable from one that cannot
// fail. Each case below copies the repo to a temp dir, breaks exactly one
// invariant, and asserts the validator rejects it — and that it accepts the
// untouched copy, so the failures are not just "everything is broken".
//
// Office-free. Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { cp, mkdtemp, mkdir, rm, stat, writeFile, readFile, appendFile, readdir, copyFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { REPO } from "./tracked-files.mjs";

const execFileAsync = promisify(execFile);
const EXT = path.join(".github", "extensions", "office-canvas");

/** Copies just enough of the repo for the validator to run. */
async function stageRepo() {
    const root = await mkdtemp(path.join(tmpdir(), "office-canvas-validate-"));
    await cp(path.join(REPO, "tools"), path.join(root, "tools"), { recursive: true });
    await cp(path.join(REPO, EXT), path.join(root, EXT), { recursive: true });
    return root;
}

async function runValidator(root) {
    try {
        const { stdout } = await execFileAsync(process.execPath, [path.join(root, "tools", "validate-extensions.mjs")]);
        return { ok: true, output: stdout };
    } catch (err) {
        return { ok: false, output: `${err.stdout ?? ""}${err.stderr ?? ""}` };
    }
}

/** Breaks one invariant and asserts the validator notices. */
async function rejects(mutate, expected, label = "") {
    const root = await stageRepo();
    const suffix = label ? ` (${label})` : "";
    try {
        await mutate(path.join(root, EXT), root);
        const result = await runValidator(root);
        assert.equal(result.ok, false, `validator accepted a repo that ${expected}${suffix}`);
        assert.match(result.output, expected, `wrong reason${suffix}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
}

test("an untouched copy passes", async () => {
    const root = await stageRepo();
    try {
        const result = await runValidator(root);
        assert.equal(result.ok, true, `validator rejected a clean repo:\n${result.output}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a console.log in the extension process is rejected", async () => {
    // The documented failure mode: stdout carries JSON-RPC.
    await rejects(
        (ext) => appendFile(path.join(ext, "src", "server.mjs"), '\nconsole.log("oops");\n'),
        /corrupts JSON-RPC/,
    );
});

test("a console.log under src/ui is allowed", async () => {
    // The iframe has its own console; a repo-wide ban would be wrong.
    const root = await stageRepo();
    try {
        await appendFile(path.join(root, EXT, "src", "ui", "app.js"), '\nconsole.log("fine here");\n');
        const result = await runValidator(root);
        assert.equal(result.ok, true, `validator wrongly rejected src/ui console use:\n${result.output}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a package.json in the extension is rejected", async () => {
    await rejects((ext) => writeFile(path.join(ext, "package.json"), "{}"), /C2/);
});

test("an import escaping the extension folder is rejected", async () => {
    // Every shape that can reach a module, because the first version of this
    // check handled only `import ... from` and let the bare form through.
    // The specifier is interpolated rather than written literally, so these
    // stay data — this file must not itself trip the check it is exercising.
    const OUT = "../../../../tools/validate-extensions.mjs";
    const escapes = {
        "bare side-effect import": `import "${OUT}";`,
        "default import": `import v from "${OUT}";`,
        "named import": `import { x } from "${OUT}";`,
        "re-export": `export * from "${OUT}";`,
        "dynamic import": `const m = await import("${OUT}");`,
    };
    for (const [label, statement] of Object.entries(escapes)) {
        await rejects(
            (ext) => appendFile(path.join(ext, "src", "store.mjs"), `\n${statement}\n`),
            /outside the extension/,
            label,
        );
    }
});

test("a relative import inside the extension is allowed", async () => {
    const root = await stageRepo();
    try {
        const inside = "./watcher.mjs";
        await appendFile(path.join(root, EXT, "src", "store.mjs"), `\nimport "${inside}";\n`);
        const result = await runValidator(root);
        assert.equal(result.ok, true, `validator wrongly rejected an internal import:\n${result.output}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a manifest name that disagrees with its folder is rejected", async () => {
    await rejects(
        (ext) =>
            writeFile(
                path.join(ext, "copilot-extension.json"),
                JSON.stringify({ name: "wrong", version: 1, productVersion: "1.0.0" }),
            ),
        /does not match its folder/,
    );
});

test("a missing manifest is rejected", async () => {
    await rejects((ext) => rm(path.join(ext, "copilot-extension.json")), /copilot-extension\.json missing/);
});

// The bug this pins was real: the manifest carried `"version": "1.0.0"`, the
// in-place loader accepted it, and `install_extension` refused the whole
// extension with `invalid type: string "1.0.0", expected u32`. So the extension
// ran perfectly in the dev checkout and could not be installed by anyone.
test("a semver string in `version` is rejected — the installer parses it as u32", async () => {
    await rejects(
        (ext) =>
            writeFile(
                path.join(ext, "copilot-extension.json"),
                JSON.stringify({ name: "office-canvas", version: "1.0.0", productVersion: "1.0.0" }),
            ),
        /manifest version must be the number 1/,
    );
});

test("`version` must be 1 exactly, not merely a number", async () => {
    await rejects(
        (ext) =>
            writeFile(
                path.join(ext, "copilot-extension.json"),
                JSON.stringify({ name: "office-canvas", version: 2, productVersion: "1.0.0" }),
            ),
        /manifest version must be the number 1/,
    );
});

test("a non-semver productVersion is rejected", async () => {
    await rejects(
        (ext) =>
            writeFile(
                path.join(ext, "copilot-extension.json"),
                JSON.stringify({ name: "office-canvas", version: 1, productVersion: "1" }),
            ),
        /not semver/,
    );
});

test("a missing productVersion is rejected — release tagging has no other source", async () => {
    await rejects(
        (ext) =>
            writeFile(
                path.join(ext, "copilot-extension.json"),
                JSON.stringify({ name: "office-canvas", version: 1 }),
            ),
        /not semver/,
    );
});

test("a binary file in the extension is rejected", async () => {
    // C4: gists refuse binaries. This is what keeps spikes and screenshots out.
    await rejects(
        (ext) => writeFile(path.join(ext, "src", "blob.bin"), Buffer.from([0xff, 0xfe, 0x00, 0x80])),
        /not valid UTF-8/,
    );
});

test("a file over the per-file size limit is rejected", async () => {
    await rejects(
        (ext) => writeFile(path.join(ext, "src", "huge.mjs"), "// filler\n".repeat(120_000)),
        /C4 limit 1000000/,
    );
});

// The two tests below pin the per-file cap to the decimal 1,000,000 that
// install_extension actually enforces, measured against the running app. They
// exist to stop someone "tidying" the constant back to 1024*1024: that is the
// bug they replaced, and it left a 48,576-byte window in which a file passed
// validation and then failed at install. Asserting the constant alone would
// not catch it — only straddling the boundary does.
test("a file of exactly the per-file limit is accepted", async () => {
    const root = await stageRepo();
    try {
        await writeFile(path.join(root, EXT, "src", "edge.mjs"), "/".repeat(999_999) + "\n");
        assert.equal((await stat(path.join(root, EXT, "src", "edge.mjs"))).size, 1_000_000, "fixture must be exact");
        const result = await runValidator(root);
        assert.equal(result.ok, true, `1,000,000 bytes must pass:\n${result.output}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a file one byte over the per-file limit is rejected", async () => {
    await rejects(async (ext) => {
        const file = path.join(ext, "src", "edge.mjs");
        await writeFile(file, "/".repeat(1_000_000) + "\n");
        assert.equal((await stat(file)).size, 1_000_001, "fixture must be exact");
    }, /C4 limit 1000000/);
});

test("the total limit is the decimal 5,000,000, not 5 MiB", async () => {
    // 5 MiB is 5,242,880. A tree between the two sizes passes a binary-constant
    // validator and is then refused by the installer.
    const root = await stageRepo();
    const filler = "/".repeat(999_998) + "\n";
    try {
        for (let i = 0; i < 5; i++) await writeFile(path.join(root, EXT, "src", `bulk${i}.mjs`), filler);
        const result = await runValidator(root);
        assert.equal(result.ok, false, "a tree over 5,000,000 bytes must be rejected");
        assert.match(result.output, /exceeds the C4 limit of 5000000/);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a syntax error is rejected", async () => {
    await rejects((ext) => appendFile(path.join(ext, "src", "store.mjs"), "\nfunction ( {\n"), /does not parse/);
});

// --- what the run reports, as opposed to what it rejects ---------------------
//
// The routine success line used to carry a byte total beside the file count. It
// was accumulated from disk, so a CRLF checkout and an LF checkout disagreed
// about the same commit -- ~23 KB apart, measured across two clones -- while the
// count was identical in both. It was nevertheless quoted between machines as if
// it identified a tree. It is gone; bytes now appear only against the envelope,
// where they are the thing being decided (#82).
//
// Nothing else can see this. The envelope band is far above where this repo
// sits, so the warning never fires on the real tree, and a validator whose
// reporting has never been observed is one whose reporting cannot be trusted --
// the same argument the head of this file makes about its checks.

/** Total bytes of the files under a directory, counted the way the validator counts. */
async function dirBytes(dir) {
    let total = 0;
    for (const entry of await readdir(dir, { recursive: true, withFileTypes: true })) {
        if (entry.isFile()) total += (await stat(path.join(entry.parentPath ?? entry.path, entry.name))).size;
    }
    return total;
}

/**
 * Stages a minimal extension beside the real validator, rather than a copy of
 * the real extension.
 *
 * The boundary cases below need an exact total, and the real extension is both
 * large and free to grow. Padding a copy of it up to 4,000,000 bytes would start
 * failing the day it grows past that on its own -- which is growth, not a defect
 * -- and would quietly make the warning threshold a hard cap enforced by CI. The
 * validator under test is the real one either way; only the tree it is pointed
 * at is synthetic.
 */
async function stageMinimal() {
    const root = await mkdtemp(path.join(tmpdir(), "office-canvas-sizing-"));
    const ext = path.join(root, ".github", "extensions", "sizing");
    await mkdir(ext, { recursive: true });
    await mkdir(path.join(root, "tools"), { recursive: true });
    await copyFile(
        path.join(REPO, "tools", "validate-extensions.mjs"),
        path.join(root, "tools", "validate-extensions.mjs"),
    );
    await writeFile(path.join(ext, "extension.mjs"), "export default {};\n");
    await writeFile(
        path.join(ext, "copilot-extension.json"),
        JSON.stringify({ name: "sizing", version: 1, productVersion: "1.0.0" }),
    );
    return { root, ext };
}

/** Pads a staged extension to an exact total, in files the per-file cap allows. */
async function padTo(ext, target) {
    let current = await dirBytes(ext);
    assert.ok(current < target, `fixture is already ${current} bytes, at or over the ${target} this case needs`);
    // "x" is one byte in UTF-8, so `repeat(chunk)` is exactly `chunk` bytes --
    // which is what lets the totals below be stated exactly rather than
    // approximately. Nothing rewrites these: they are written straight into an
    // untracked temp tree, so no checkout or filter is in play.
    for (let i = 0; current < target; i++) {
        const chunk = Math.min(900_000, target - current);
        await writeFile(path.join(ext, `pad${i}.txt`), "x".repeat(chunk));
        current += chunk;
    }
    assert.equal(await dirBytes(ext), target, "fixture must be exact");
}

test("the success line reports the file count and no byte total", async () => {
    const root = await stageRepo();
    try {
        const result = await runValidator(root);
        assert.equal(result.ok, true, `validator rejected a clean repo:\n${result.output}`);
        assert.match(result.output, /^office-canvas: \d+ files$/m, "expected a bare file count");
        assert.doesNotMatch(result.output, /KB/, "the checkout-dependent KB figure is gone for good");
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

// The pair below straddles the threshold, because asserting the constant would
// not catch the boundary being moved by one. Exactly at 5,000,000 - 1,000,000
// the headroom is exactly one whole legal file, and a file of that size lands
// the total *on* the envelope, which passes -- so no single addition can breach
// it from there and a warning would name a danger that does not exist.
test("a total of exactly the envelope less one file is silent — no single legal file can breach it", async () => {
    const { root, ext } = await stageMinimal();
    try {
        await padTo(ext, 4_000_000);
        const result = await runValidator(root);
        assert.equal(result.ok, true, `4,000,000 bytes must pass:\n${result.output}`);
        assert.doesNotMatch(result.output, /on disk/, "no single legal file can breach the envelope from here");
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("one byte more warns, and states the headroom and why it matters", async () => {
    const { root, ext } = await stageMinimal();
    try {
        await padTo(ext, 4_000_001);
        const result = await runValidator(root);
        assert.equal(result.ok, true, `4,000,001 bytes is under the envelope and must pass:\n${result.output}`);
        // The whole line, not a fragment of it. Every clause is a claim the
        // message makes -- 999,999 of headroom, that this is less than the
        // per-file limit, and that one further file could therefore exceed the
        // total -- and all three are why the warning is true at this size. A
        // partial match would let a message that reversed the reasoning pass.
        const expected =
            "sizing: 4000001 of 5000000 bytes on disk, 999999 left — " +
            "less than the 1000000-byte per-file limit, so one more file could exceed the total";
        assert.ok(result.output.includes(expected), `wanted exactly:\n${expected}\ngot:\n${result.output}`);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("bytes are still reported when the envelope is exceeded", async () => {
    const { root, ext } = await stageMinimal();
    try {
        await padTo(ext, 5_000_001);
        const result = await runValidator(root);
        assert.equal(result.ok, false, "5,000,001 bytes must be rejected");
        // One byte over, so this also pins the envelope boundary itself, which
        // the sibling 5 MiB test above does not: its fixture lands well past
        // both constants and discriminates only through the limit named in the
        // message.
        assert.match(result.output, /total 5000001 bytes exceeds the C4 limit of 5000000/);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("a missing extension.mjs is rejected", async () => {
    await rejects((ext) => rm(path.join(ext, "extension.mjs")), /missing extension\.mjs/);
});

test("an empty extensions directory fails rather than passing vacuously", async () => {
    const root = await stageRepo();
    try {
        await rm(path.join(root, EXT), { recursive: true, force: true });
        const result = await runValidator(root);
        assert.equal(result.ok, false, "an empty extensions dir must not report success");
        assert.match(result.output, /vacuously/);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});

test("the validator checks the extension actually shipped", async () => {
    // Guards against the staging helper silently copying nothing.
    const root = await stageRepo();
    try {
        const manifest = JSON.parse(await readFile(path.join(root, EXT, "copilot-extension.json"), "utf8"));
        assert.equal(manifest.name, "office-canvas");
    } finally {
        await rm(root, { recursive: true, force: true });
    }
});
