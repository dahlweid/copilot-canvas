// Tests for tools/package-extension.mjs.
//
// The packager is driven as a subprocess rather than imported, for the same
// reason validator.test.mjs does it: this file lives inside the extension
// folder, and an `import` reaching up into tools/ would be exactly the C3
// escape the validator is there to reject. Spawning keeps the test honest
// about the constraint it is testing under.
//
// Everything here is file layout, so it is Office-free and runs on
// ubuntu-latest. Run:
//   node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import test from "node:test";
import assert from "node:assert/strict";
import { cp, mkdir, mkdtemp, readdir, readFile, rm, stat, writeFile, appendFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");
const EXT = path.join(".github", "extensions", "office-canvas");
const NAME = "office-canvas";

// Artefact filenames carry the *product* version, so this file tracks whatever
// the manifest says rather than a literal. A release bump is paperwork; it must
// not redden the packager's own tests, and a test that has to be edited to
// permit a version change is a test that stops being run truthfully.
const VERSION = JSON.parse(
    await readFile(path.join(REPO, EXT, "copilot-extension.json"), "utf8"),
).productVersion;

// If that read ever goes wrong, VERSION becomes undefined and every filename
// below turns into "office-canvas-undefined...", which fails loudly rather than
// vacuously — but only if something checks. This is the check.
if (!/^\d+\.\d+\.\d+$/.test(VERSION ?? "")) {
    throw new Error(`copilot-extension.json productVersion is not semver: ${VERSION}`);
}

async function stageRepo() {
    const root = await mkdtemp(path.join(tmpdir(), "office-canvas-package-"));
    await cp(path.join(REPO, "tools"), path.join(root, "tools"), { recursive: true });
    await cp(path.join(REPO, EXT), path.join(root, EXT), { recursive: true });
    return root;
}

async function runPackager(root, args = []) {
    const script = path.join(root, "tools", "package-extension.mjs");
    try {
        const { stdout } = await execFileAsync(process.execPath, [script, ...args], { cwd: root });
        return { ok: true, output: stdout };
    } catch (err) {
        return { ok: false, output: `${err.stdout ?? ""}${err.stderr ?? ""}` };
    }
}

const readManifest = async (root, version = VERSION) =>
    JSON.parse(await readFile(path.join(root, "dist", `${NAME}-${version}.package.json`), "utf8"));

/** Every file in a packaged folder, as sorted POSIX-relative paths. */
async function tree(dir, base = dir) {
    const out = [];
    for (const entry of await readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...(await tree(full, base)));
        else out.push(path.relative(base, full).split(path.sep).join("/"));
    }
    return out.sort();
}

async function withStagedRepo(fn) {
    const root = await stageRepo();
    try {
        return await fn(root);
    } finally {
        await rm(root, { recursive: true, force: true });
    }
}

test("packages the extension into a folder, a gist body and a manifest", async () => {
    await withStagedRepo(async (root) => {
        const result = await runPackager(root);
        assert.equal(result.ok, true, `packaging failed:\n${result.output}`);

        const manifest = await readManifest(root);
        assert.equal(manifest.name, NAME);
        assert.equal(manifest.version, VERSION);

        const files = await tree(path.join(root, "dist", NAME));
        // C1: the two files the runtime requires must survive packaging.
        assert.ok(files.includes("extension.mjs"), "extension.mjs must be in the artefact");
        assert.ok(files.includes("copilot-extension.json"), "the manifest must be in the artefact");
        assert.deepEqual(
            files,
            manifest.files.map((f) => f.path).sort(),
            "the artefact on disk must match the manifest exactly",
        );

        await stat(path.join(root, "dist", `${NAME}-${VERSION}.gist.json`));
    });
});

// Measured against the running app: `install_extension` parses the manifest's
// `version` as a Rust u32 and refuses the whole extension otherwise, so an
// artefact whose manifest carries a semver string there cannot be installed by
// either path. The product version is a separate key.
test("the shipped manifest is the shape install_extension will accept", async () => {
    await withStagedRepo(async (root) => {
        assert.equal((await runPackager(root)).ok, true);

        const shipped = JSON.parse(
            await readFile(path.join(root, "dist", NAME, "copilot-extension.json"), "utf8"),
        );
        assert.equal(shipped.name, NAME, "the installer matches the folder on `name`");
        assert.equal(shipped.version, 1, "`version` is the manifest format version, parsed as u32");
        assert.equal(typeof shipped.version, "number", "a semver string here makes the extension uninstallable");
        assert.match(shipped.productVersion, /^\d+\.\d+\.\d+$/, "the product version has its own key");

        // And the report names the release by the product version, not by the
        // manifest format version — otherwise every release is "1".
        const report = await readManifest(root);
        assert.equal(report.version, shipped.productVersion);
        assert.equal(report.manifestFormatVersion, 1);
    });
});

test("the packaged folder is byte-identical to the source it came from", async () => {
    // C3: install_extension copies the folder and runs it as-is, so packaging
    // must not normalise line endings or rewrite anything on the way through.
    await withStagedRepo(async (root) => {
        assert.equal((await runPackager(root)).ok, true);
        for (const rel of await tree(path.join(root, "dist", NAME))) {
            const packaged = await readFile(path.join(root, "dist", NAME, rel));
            const source = await readFile(path.join(root, EXT, rel));
            assert.ok(packaged.equals(source), `${rel} was altered during packaging`);
        }
    });
});

test("test/ is excluded, and the reason is reported rather than assumed", async () => {
    await withStagedRepo(async (root) => {
        const result = await runPackager(root);
        const manifest = await readManifest(root);
        const files = await tree(path.join(root, "dist", NAME));

        assert.equal(
            files.some((f) => f.startsWith("test/")),
            false,
            "no development test material may reach the artefact",
        );
        const excludedTests = manifest.excluded.filter((e) => e.rule === "test");
        assert.ok(excludedTests.length > 0, "the test files must be recorded as excluded");
        assert.ok(
            excludedTests.some((e) => e.path.startsWith("test/integration/")),
            "the Word-driving suites in particular must be excluded",
        );
        // The justification is the point of the list; an unexplained exclusion
        // is indistinguishable from an accident.
        assert.match(result.output, /Excluded, and why:/);
        assert.match(result.output, /test — \d+ file\(s\)/);
    });
});

test("development-only material is excluded, each under a named rule", async () => {
    // One staged repo, one packaging run: every rule that has ever mattered is
    // exercised at once, and each must be attributed to the rule a reader
    // would name.
    const cases = {
        "node_modules/left-pad/index.js": "node_modules",
        "spikes/screenshot-notes.md": "spikes",
        "artifacts/render-cache.json": "artifacts",
        "dist/stale.mjs": "output",
        "build/stale.mjs": "output",
        ".vscode/settings.json": "editor",
        ".gitignore": "vcs",
        ".DS_Store": "os-junk",
        "src/debug.log": "scratch",
        "src/server.mjs.orig": "scratch",
        ".env.production": "secrets",
    };

    await withStagedRepo(async (root) => {
        for (const rel of Object.keys(cases)) {
            const full = path.join(root, EXT, rel);
            await mkdir(path.dirname(full), { recursive: true });
            await writeFile(full, "// development-only\n");
        }

        assert.equal((await runPackager(root)).ok, true);
        const manifest = await readManifest(root);
        const files = await tree(path.join(root, "dist", NAME));
        const byPath = new Map(manifest.excluded.map((e) => [e.path, e.rule]));

        for (const [rel, rule] of Object.entries(cases)) {
            assert.equal(files.includes(rel), false, `${rel} must not be in the artefact`);
            assert.equal(byPath.get(rel), rule, `${rel} should be excluded by the '${rule}' rule`);
        }
    });
});

test("source files are not excluded by any rule", async () => {
    // The mirror of the test above: an over-eager rule that swallowed src/
    // would still produce a passing artefact right up until someone installed it.
    await withStagedRepo(async (root) => {
        assert.equal((await runPackager(root)).ok, true);
        const files = await tree(path.join(root, "dist", NAME));
        for (const rel of ["src/server.mjs", "src/render-cache.mjs", "src/ui/app.js", "src/word/word-host.ps1"]) {
            assert.ok(files.includes(rel), `${rel} must survive packaging`);
        }
    });
});

test("packaging the same source twice produces an identical artefact", async () => {
    await withStagedRepo(async (root) => {
        const first = path.join(root, "out-a");
        const second = path.join(root, "out-b");
        assert.equal((await runPackager(root, ["--out", first])).ok, true);
        assert.equal((await runPackager(root, ["--out", second])).ok, true);

        const manifestA = JSON.parse(await readFile(path.join(first, `${NAME}-${VERSION}.package.json`), "utf8"));
        const manifestB = JSON.parse(await readFile(path.join(second, `${NAME}-${VERSION}.package.json`), "utf8"));
        assert.equal(manifestA.digest, manifestB.digest, "the digest must not depend on when packaging ran");
        assert.deepEqual(manifestA, manifestB, "the whole manifest must be reproducible");

        // A digest that matches while the bytes differ would be worse than no
        // digest, so the files are compared directly too.
        const filesA = await tree(path.join(first, NAME));
        assert.deepEqual(filesA, await tree(path.join(second, NAME)));
        for (const rel of filesA) {
            const a = await readFile(path.join(first, NAME, rel));
            const b = await readFile(path.join(second, NAME, rel));
            assert.ok(a.equals(b), `${rel} differs between two runs`);
        }
        const gistA = await readFile(path.join(first, `${NAME}-${VERSION}.gist.json`), "utf8");
        const gistB = await readFile(path.join(second, `${NAME}-${VERSION}.gist.json`), "utf8");
        assert.equal(gistA, gistB, "the gist body must be reproducible, key order included");
    });
});

test("output timestamps are fixed, so a downstream zip is reproducible", async () => {
    await withStagedRepo(async (root) => {
        assert.equal((await runPackager(root)).ok, true);
        const manifest = await readManifest(root);
        for (const rel of await tree(path.join(root, "dist", NAME))) {
            const { mtime } = await stat(path.join(root, "dist", NAME, rel));
            assert.equal(
                Math.floor(mtime.getTime() / 1000),
                manifest.sourceDateEpoch,
                `${rel} carries a checkout timestamp, which would make a zip non-reproducible`,
            );
        }
    });
});

test("SOURCE_DATE_EPOCH overrides the stamp", async () => {
    await withStagedRepo(async (root) => {
        const script = path.join(root, "tools", "package-extension.mjs");
        await execFileAsync(process.execPath, [script], { cwd: root, env: { ...process.env, SOURCE_DATE_EPOCH: "946684800" } });
        const manifest = await readManifest(root);
        assert.equal(manifest.sourceDateEpoch, 946684800);
        const { mtime } = await stat(path.join(root, "dist", NAME, "extension.mjs"));
        assert.equal(Math.floor(mtime.getTime() / 1000), 946684800);
    });
});

test("refuses to package anything the validator rejects", async () => {
    // The gate that matters: packaging must never publish what CI would fail.
    await withStagedRepo(async (root) => {
        await appendFile(path.join(root, EXT, "src", "server.mjs"), '\nconsole.log("oops");\n');
        const result = await runPackager(root);
        assert.equal(result.ok, false, "a console.log in the extension process must block packaging");
        assert.match(result.output, /corrupts JSON-RPC/);
        assert.match(result.output, /validation failed for the source folder/);
    });
});

test("an exclusion cannot launder a validation failure", async () => {
    // A package.json under test/ is excluded from the artefact, so validating
    // only the artefact would let it through — while `node
    // tools/validate-extensions.mjs` in CI still fails on the source. Packaging
    // that disagrees with validation is the failure mode this prevents.
    await withStagedRepo(async (root) => {
        await writeFile(path.join(root, EXT, "test", "package.json"), "{}");
        const result = await runPackager(root);
        assert.equal(result.ok, false, "an excluded-but-invalid file must still block packaging");
        assert.match(result.output, /C2/);
    });
});

test("validates the artefact as well as the source", async () => {
    // The opposite mistake: an exclusion that removes something the extension
    // needs. Widening a rule to swallow extension.mjs must be caught, and only
    // the artefact pass can catch it.
    await withStagedRepo(async (root) => {
        const tool = path.join(root, "tools", "package-extension.mjs");
        const text = await readFile(tool, "utf8");
        await writeFile(
            tool,
            text.replace(
                'match: (rel) => segments(rel).includes("test"),',
                'match: (rel) => segments(rel).includes("test") || rel === "extension.mjs",',
            ),
        );
        const result = await runPackager(root);
        assert.equal(result.ok, false, "an exclusion that guts the artefact must fail");
        assert.match(result.output, /validation failed for the packaged artefact/);
        assert.match(result.output, /missing extension\.mjs/);
    });
});

test("the gist body is flat, backslash-separated and complete", async () => {
    // Verified against a real shared gist: nested paths are encoded with a
    // backslash. They cannot be written as files on Windows, where the
    // backslash *is* the separator, so the bundle is the API request body.
    await withStagedRepo(async (root) => {
        assert.equal((await runPackager(root)).ok, true);
        const payload = JSON.parse(await readFile(path.join(root, "dist", `${NAME}-${VERSION}.gist.json`), "utf8"));
        const manifest = await readManifest(root);

        const names = Object.keys(payload.files);
        assert.deepEqual(
            names.slice().sort(),
            manifest.files.map((f) => f.path.split("/").join("\\")).sort(),
            "the gist body must carry exactly the artefact's files, backslash-encoded",
        );
        assert.equal(
            names.some((n) => n.includes("/")),
            false,
            "a forward slash in a gist name means the bundle is not flat",
        );
        assert.ok(names.includes("src\\ui\\app.js"), "nested paths must be encoded, not dropped");
        assert.equal(payload.public, false, "share_extension creates a private gist; the body must match");

        const content = payload.files["extension.mjs"].content;
        assert.equal(content, await readFile(path.join(root, EXT, "extension.mjs"), "utf8"));
    });
});

test("--expect-version gates the tag against the manifest", async () => {
    await withStagedRepo(async (root) => {
        const wrong = await runPackager(root, ["--expect-version", "v9.9.9"]);
        assert.equal(wrong.ok, false, "a tag that disagrees with the manifest must not publish");
        assert.match(wrong.output, /version mismatch/);

        // Both spellings, because git tags carry the v and manifests do not.
        assert.equal((await runPackager(root, ["--expect-version", `v${VERSION}`])).ok, true);
        assert.equal((await runPackager(root, ["--expect-version", VERSION])).ok, true);
    });
});

test("reports size against the limits install_extension actually enforces", async () => {
    await withStagedRepo(async (root) => {
        const result = await runPackager(root);
        const { budget, totalBytes, files } = await readManifest(root);

        // Decimal, and measured — not 1 MiB / 5 MiB. This is the assertion that
        // stops the budget silently drifting back to the wrong constants.
        assert.equal(budget.maxFileBytes, 1_000_000);
        assert.equal(budget.maxTotalBytes, 5_000_000);
        assert.equal(budget.headroomBytes, 5_000_000 - totalBytes);

        const largest = files.reduce((a, b) => (b.bytes > a.bytes ? b : a));
        assert.equal(budget.largestFile.path, largest.path);
        assert.equal(budget.largestFile.bytes, largest.bytes);
        assert.equal(
            totalBytes,
            files.reduce((sum, f) => sum + f.bytes, 0),
            "the reported total must be the sum of what shipped",
        );

        // Printed on every run so the headroom a later layer spends stays visible.
        assert.match(result.output, /C4 budget/);
        assert.match(result.output, /headroom/);
    });
});

test("refuses to package an extension that does not exist", async () => {
    await withStagedRepo(async (root) => {
        const result = await runPackager(root, ["no-such-extension"]);
        assert.equal(result.ok, false);
        assert.match(result.output, /no extension 'no-such-extension'/);
    });
});

test("refuses an output directory that contains the repository", async () => {
    // `--out .` used to mean "delete the repository and then write to it".
    await withStagedRepo(async (root) => {
        const result = await runPackager(root, ["--out", root]);
        assert.equal(result.ok, false, "writing output over the repo must be refused");
        assert.match(result.output, /contains the repository/);
        await stat(path.join(root, EXT, "extension.mjs"));
    });
});

test("refuses an unknown option rather than ignoring it", async () => {
    await withStagedRepo(async (root) => {
        const result = await runPackager(root, ["--dry-run"]);
        assert.equal(result.ok, false);
        assert.match(result.output, /unknown option/);
    });
});
