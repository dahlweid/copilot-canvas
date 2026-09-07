// The anchor gate's own liveness: proof that `tools/check-mutation-anchors.mjs`
// can go red, and for the reason it claims.
//
// This matters more here than for most tools. The gate exists because a mutant
// whose anchor no longer matches is *counted* while measuring nothing -- and a
// gate that cannot fail is the same defect one level up. So the end-to-end test
// below is a pair: two fixture harnesses identical but for one anchor string,
// one resolving and one not, asserting exit 0 and exit 1 respectively. Asserting
// only the red would leave "the gate always exits 1" indistinguishable from "the
// gate detected the stale anchor".
//
// Office-free. It does need PowerShell, because asking the harness for its own
// array is the whole design -- a hosted `ubuntu-latest` runner has `pwsh`
// preinstalled. A machine without it fails loudly rather than skipping: the gate
// cannot run there at all, and a skip would hide that behind a green suite.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");
const GATE = path.join(REPO, "tools", "check-mutation-anchors.mjs");

// Loaded by URL rather than by a relative specifier, the way
// `check-citation-lines.test.mjs` reaches its own gate. C3 forbids the
// extension importing anything outside itself, and it is right to: the folder
// must run exactly as committed, and `tools/` is not shipped with it. A test is
// the one thing here that legitimately reaches up, so it does so at runtime and
// stays invisible to the packaged tree.
const { auditMutants } = await import(`file://${GATE.split(path.sep).join("/")}`);
/** An in-memory stand-in for the reader the gate hands `auditMutants`. */
const reader = (files) => async (p) => (p in files ? files[p] : null);

const mutant = (over) => ({
    harness: "mutate-fixture.ps1",
    name: "a defect worth injecting",
    file: "subject.mjs",
    path: "/x/subject.mjs",
    from: "const alive = 1;",
    to: "const alive = 0;",
    ...over,
});

test("an anchor occurring exactly once is the only clean case", async () => {
    const problems = await auditMutants(
        [mutant()],
        reader({ "/x/subject.mjs": "before\nconst alive = 1;\nafter\n" }),
    );
    assert.deepEqual(problems, []);
});

test("an anchor that occurs nowhere is reported dead rather than passed over", async () => {
    const problems = await auditMutants([mutant()], reader({ "/x/subject.mjs": "the line was reworded\n" }));

    assert.equal(problems.length, 1);
    assert.equal(problems[0].kind, "DEAD");
    // On the anchor itself, not on a position: the anchor is an exact string,
    // so quoting it is both the precise reference and the one that cannot rot.
    assert.match(problems[0].detail, /const alive = 1;/);
});

test("an anchor that occurs twice is reported ambiguous, because the harness would replace both", async () => {
    const problems = await auditMutants(
        [mutant()],
        reader({ "/x/subject.mjs": "const alive = 1;\nconst alive = 1;\n" }),
    );

    assert.equal(problems.length, 1);
    assert.equal(problems[0].kind, "AMBIGUOUS");
    assert.match(problems[0].detail, /matches 2x/);
});

test("occurrences are counted the way String.Replace consumes them, not by overlap", async () => {
    // "aa" appears twice in "aaaa" non-overlapping and three times overlapping.
    // .NET advances past each match, so an overlap-counting gate would invent an
    // AMBIGUOUS the harness does not have.
    const problems = await auditMutants(
        [mutant({ from: "aa", to: "bb" })],
        reader({ "/x/subject.mjs": "aaaa" }),
    );
    assert.equal(problems[0].detail.match(/matches (\d+)x/)[1], "2");
});

test("a mutant whose from equals its to is reported, since it rewrites the file unchanged", async () => {
    const problems = await auditMutants(
        [mutant({ to: "const alive = 1;" })],
        reader({ "/x/subject.mjs": "const alive = 1;\n" }),
    );

    assert.equal(problems.length, 1);
    assert.equal(problems[0].kind, "NO-OP");
});

test("a file the gate cannot read is reported, not silently treated as empty", async () => {
    // An empty read would come back as DEAD, which sends a reader to re-anchor a
    // mutant whose anchor is fine while the path is what moved.
    const problems = await auditMutants([mutant()], reader({}));

    assert.equal(problems.length, 1);
    assert.equal(problems[0].kind, "UNREADABLE");
});

/**
 * A harness directory holding one subject file and one harness that lists a
 * single mutant against it. `from` is the caller's, so the two arms differ in
 * exactly one string and nothing else.
 *
 * The fixture emits the same JSON shape the real harnesses do, deliberately
 * hand-written rather than copied from one of them: what is under test here is
 * the gate's reaction to the shape, and a fixture that imported the real
 * harness could not carry a stale anchor without editing the tree.
 */
async function harnessDir(from) {
    const dir = await mkdtemp(path.join(tmpdir(), "anchor-gate-"));
    await writeFile(path.join(dir, "subject.mjs"), "before\nconst alive = 1;\nafter\n");
    await writeFile(
        path.join(dir, "mutate-fixture.ps1"),
        [
            "param([switch]$ListMutants)",
            "$mutants = @(",
            `    @{ name = 'the fixture mutant'; file = 'subject.mjs'; from = '${from}'; to = 'const alive = 0;' }`,
            ")",
            "if ($ListMutants) {",
            "    [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)",
            "    $listed = @($mutants | ForEach-Object {",
            "        [pscustomobject]@{",
            "            harness = 'mutate-fixture.ps1'",
            "            name    = $_.name",
            "            file    = $_.file",
            "            path    = (Join-Path $PSScriptRoot $_.file)",
            "            from    = $_.from",
            "            to      = $_.to",
            "        }",
            "    })",
            "    ConvertTo-Json -InputObject $listed -Depth 4",
            "    exit 0",
            "}",
            "",
        ].join("\n"),
    );
    return dir;
}

async function runGate(dir) {
    try {
        const { stdout, stderr } = await execFileAsync(process.execPath, [GATE, "--harness-dir", dir]);
        return { code: 0, out: stdout + stderr };
    } catch (err) {
        return { code: err.code, out: (err.stdout ?? "") + (err.stderr ?? "") };
    }
}

test("the gate exits 0 on a resolving anchor and 1 on a stale one, driven end to end", async () => {
    const live = await harnessDir("const alive = 1;");
    const stale = await harnessDir("const gone = 1;");
    try {
        const green = await runGate(live);
        const red = await runGate(stale);

        // The green arm first, because it is what makes the red arm evidence:
        // without it, a gate broken in any way at all would produce the same
        // exit 1 and this test would still pass.
        assert.equal(green.code, 0, `the resolving anchor did not pass:\n${green.out}`);
        assert.match(green.out, /Every anchor resolves exactly once/);

        assert.equal(red.code, 1, `a stale anchor did not fail the gate:\n${red.out}`);
        assert.match(red.out, /DEAD/);
        assert.match(red.out, /the fixture mutant/);
        // Named by its anchor, so the failure says which string to re-pin.
        assert.match(red.out, /const gone = 1;/);
    } finally {
        await rm(live, { recursive: true, force: true });
        await rm(stale, { recursive: true, force: true });
    }
});

test("a harness directory with no harnesses fails rather than passing vacuously", async () => {
    const empty = await mkdtemp(path.join(tmpdir(), "anchor-gate-empty-"));
    try {
        const { code, out } = await runGate(empty);
        assert.equal(code, 1, `an empty harness directory was accepted:\n${out}`);
        assert.match(out, /no mutation harnesses found/);
    } finally {
        await rm(empty, { recursive: true, force: true });
    }
});
