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
import { cp, mkdtemp, rm, stat, writeFile, readFile, appendFile } from "node:fs/promises";
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
