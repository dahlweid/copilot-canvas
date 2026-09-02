// The two failure shapes `gitAvailable()` has to tell apart, measured rather
// than asserted -- plus the `false` branch and its consequence, driven through
// the subject in a real non-repository layout.
//
// The platform claim is backed by `spikes/git-guard/probes/probe-execfile-error-shapes.mjs`,
// which drives both arms and reports `err.code`, its `typeof` and whether
// `stderr` is populated. Measured there:
//
//   ARM 1  spawn failure (no such binary):     code = "ENOENT", typeof string, stderr 0 bytes
//   ARM 2  git ran and exited non-zero:        code = 128,      typeof number, stderr 69 bytes
//
// Both arms are re-asserted below, so the table cannot go stale the way a
// comment quoting a probe's old output can. The probe is not redundant with
// them and this file is not a substitute for it: `tools/check-citations.mjs`
// tracks `probe-*.{ps1,mjs}` paths and nothing else, so a claim discharged onto
// a neighbouring test is an unchecked citation that cannot be mutation-tested
// (CONTEXT.md). The division is that the probe is the citable, portable arm --
// it can be pointed at another platform by hand -- while these run on every
// commit against the platform CI is on.
//
// Office-free. Also Windows-free and network-free: this is `node:child_process`
// behaviour and runs on `ubuntu-latest` unchanged.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { cp, mkdir, mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { REPO, gitAvailable } from "./tracked-files.mjs";

const execFileAsync = promisify(execFile);

const UNIT = path.dirname(fileURLToPath(import.meta.url));
const EXTENSION = path.resolve(UNIT, "..", "..");

/**
 * A copy of the extension laid out so `tracked-files.mjs` resolves `REPO` to a
 * directory that is not a repository -- the installed-extension mode the guards
 * exist for, built rather than described.
 *
 * `REPO` is five levels up from `test/unit/`, so the copy is buried four deep
 * and the fifth level up is the returned `repo`. `GIT_CEILING_DIRECTORIES` then
 * stops git's upward walk below the real checkout: without it this would answer
 * differently depending on whether the machine's temp directory happens to sit
 * inside a repository, which would make the test a property of the machine.
 *
 * Only the files `console-encoding.test.mjs` actually reaches are copied, and
 * they are the **real** files rather than fixtures written to look like them --
 * a fixture would be a second record of the subject, free to disagree with it.
 */
async function nonRepoInstall() {
    const root = await mkdtemp(path.join(tmpdir(), "install-mode-"));
    const ext = path.join(root, "repo", "a", "b", "extension");
    await mkdir(path.join(ext, "test", "unit"), { recursive: true });
    await mkdir(path.join(ext, "src", "word"), { recursive: true });
    for (const name of ["tracked-files.mjs", "ps-encoding-rule.mjs", "console-encoding.test.mjs"]) {
        await cp(path.join(UNIT, name), path.join(ext, "test", "unit", name));
    }
    await cp(path.join(EXTENSION, "src", "word", "word-host.ps1"), path.join(ext, "src", "word", "word-host.ps1"));

    // `NODE_TEST_CONTEXT` is set by the runner we are *inside*, and a nested
    // `node --test` that sees it switches to the machine-readable protocol its
    // parent expects and writes no human summary at all. Inheriting it left the
    // child's stdout empty -- which the tally parse below correctly refused to
    // read as green.
    const env = { ...process.env, GIT_CEILING_DIRECTORIES: root };
    delete env.NODE_TEST_CONTEXT;
    return { root, ext, repo: path.join(root, "repo"), env };
}

async function failureOf(promise) {
    try {
        await promise;
        return null;
    } catch (err) {
        return err;
    }
}

test("a spawn failure arrives with a string code, not a numeric one", async () => {
    // A name no PATH entry can hold: `execFile` does not go through a shell, so
    // this is a literal filename lookup and nothing expands it.
    const err = await failureOf(execFileAsync("git-that-does-not-exist-8f3a1c", ["--version"]));

    assert.ok(err, "spawning a nonexistent binary resolved instead of rejecting");
    assert.equal(typeof err.code, "string", `expected a string code, got ${typeof err.code} (${String(err.code)})`);
    assert.equal(err.code, "ENOENT");
});

test("a non-zero exit arrives with a numeric code and populated stderr", async () => {
    const dir = await mkdtemp(path.join(tmpdir(), "git-available-"));
    try {
        // The temp directory is not a repository, and `GIT_CEILING_DIRECTORIES`
        // stops git's upward walk at its parent so it cannot find one above --
        // otherwise this arm would pass or fail depending on where the
        // machine's temp directory happens to sit.
        const env = { ...process.env, GIT_CEILING_DIRECTORIES: path.dirname(dir) };
        const err = await failureOf(execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: dir, env }));

        if (err && typeof err.code === "string") {
            // This *is* arm 1, on a machine with no git. Reporting it as a
            // failure of arm 2 would name the wrong cause.
            return assert.fail(`git is not runnable here (${err.code}), so this arm cannot be measured`);
        }
        assert.ok(err, `git rev-parse --git-dir succeeded in ${dir}, so it is inside a repository after all`);
        assert.equal(typeof err.code, "number", `expected a numeric code, got ${typeof err.code} (${String(err.code)})`);
        assert.equal(err.code, 128);
        // Asserted because the split rests on the two shapes being genuinely
        // different objects, not just on one field.
        assert.ok(err.stderr.length > 0, "a non-zero exit carried no stderr");
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
});

// Placed after the two raw-shape measurements above and before the install-mode
// pair below, because it mutates `PATH` for its duration and settles
// `gitAvailable()`'s module-level memo. Top-level tests in `node:test` run one
// at a time, so nothing above observes the mutation -- and nothing below does
// either: the two install-mode tests use a *separately imported* copy of the
// module and a child process, neither of which shares this memo.
test("gitAvailable throws when git cannot be run, and the throw does not latch", async () => {
    // What a fresh evaluation should answer once `PATH` is back. Taken here,
    // from git itself, rather than assumed -- in an installed extension folder
    // the honest answer is `false`, and hardcoding `true` would make this test
    // pass for the wrong reason in the one mode the guard exists for.
    const inRepo = (await failureOf(execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: REPO }))) === null;

    const saved = process.env.PATH;
    let thrown;
    try {
        process.env.PATH = "";
        thrown = await failureOf(gitAvailable());
    } finally {
        process.env.PATH = saved;
    }

    assert.ok(thrown, "gitAvailable() answered instead of throwing while git was unreachable; a broken git would skip all 17 guarded tests");
    // On our own message, which is a constant in this repo -- not on a platform
    // message, which would be localized. It is here to distinguish the
    // deliberate throw from some unrelated error escaping the same call.
    assert.match(thrown.message, /no exit status/);
    assert.equal(thrown.cause?.code, "ENOENT");

    // The memoization interaction #148 asks about. Only success is cached, and
    // the throw happens before any assignment, so this must re-ask and get the
    // real answer rather than a latched failure.
    assert.equal(await gitAvailable(), inRepo);
});

test("gitAvailable answers false where git runs but there is no repository", async () => {
    // The `available = false` branch, driven through the **subject**. The two
    // arms above measure `child_process`'s error shapes on raw `execFileAsync`,
    // which establishes that the split is possible but never runs the function
    // that makes it: mutate `available = false` to `true` in `tracked-files.mjs`
    // and those three tests stay green. This one and the next go red.
    const install = await nonRepoInstall();
    try {
        // The copied module, not the imported one: `REPO` is baked in from
        // `import.meta.url` at module scope, so the only way to ask this
        // question of the real function is to load it from a location where
        // that arithmetic lands somewhere else.
        const copied = path.join(install.ext, "test", "unit", "tracked-files.mjs");
        const { REPO: copiedRepo, gitAvailable: copiedGitAvailable } = await import(`file://${copied.split(path.sep).join("/")}`);

        // Asserted, not assumed. If the layout above ever stops putting `REPO`
        // where this test thinks it does, the rest of this measures nothing.
        assert.equal(copiedRepo, install.repo, "the copied module did not resolve REPO to the non-repository root");

        const saved = process.env.GIT_CEILING_DIRECTORIES;
        let answer;
        try {
            // `gitAvailable` spawns with the ambient environment, so the ceiling
            // has to be set here rather than passed.
            process.env.GIT_CEILING_DIRECTORIES = install.root;
            // Confirms the premise before trusting the answer: git must be
            // *runnable* here, or `false` would be arriving from the spawn-failure
            // arm and this test would pass for the opposite reason.
            const err = await failureOf(execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: install.repo }));
            assert.equal(typeof err?.code, "number", `git did not answer with an exit status in ${install.repo}`);

            answer = await copiedGitAvailable();
        } finally {
            if (saved === undefined) delete process.env.GIT_CEILING_DIRECTORIES;
            else process.env.GIT_CEILING_DIRECTORIES = saved;
        }

        // `strictEqual` against `false`, not `assert.ok(!answer)`: a thrown
        // error or an `undefined` would satisfy a falsiness check while meaning
        // something entirely different.
        assert.strictEqual(answer, false, "a directory that is not a repository must answer false, not throw and not true");
    } finally {
        await rm(install.root, { recursive: true, force: true });
    }
});

test("the guarded tests skip, and the unguarded one still runs, in a non-repository install", async () => {
    // The consequence of the test above, observed where it actually matters.
    // That one asserts a return value; this one runs a **real** guarded test
    // file in the install-mode layout and reads the reporter's tally. They are
    // not two records of one quantity: a correct `false` that somehow failed to
    // produce skips would pass the first and fail this.
    //
    // It also pins #147's constraint, which nothing else in the suite can see:
    // the `word-host.ps1` citation assertion is extension-relative and
    // unguarded, and is the only assertion this file makes without a checkout.
    // Widening it to cover `live-word.ps1` would take that pass count to zero,
    // and until now that could only be caught by a human re-running the
    // measurement by hand.
    const install = await nonRepoInstall();
    try {
        const target = path.join(install.ext, "test", "unit", "console-encoding.test.mjs");
        let stdout;
        let failure = null;
        try {
            ({ stdout } = await execFileAsync(process.execPath, ["--test", target], { cwd: install.ext, env: install.env }));
        } catch (err) {
            failure = err;
            stdout = err.stdout ?? "";
        }

        assert.equal(failure, null, `the suite failed in install mode:\n${stdout}`);

        // The reporter's prefix is environment-dependent -- a TTY writes `ℹ`
        // and a pipe writes `#` -- so a filter hard-coded to either matches
        // nothing in the other. `.` matches whichever arrived.
        //
        // The strictness is not style. A parse failure is asserted as a failure
        // because the tally being *absent* is a real shape, not a hypothetical
        // one: a nested `node --test` that inherits `NODE_TEST_CONTEXT`
        // switches to the machine-readable protocol its parent expects and
        // writes no human summary -- measured at 184 bytes with zero parseable
        // tally lines -- and this test hit exactly that before `nonRepoInstall`
        // began deleting the variable. A lenient parser would then have asserted
        // `pass 1 / skipped 7` against empty stdout and passed forever.
        // Verified against eleven inputs including that one: it refuses empty
        // stdout, a TAP body with no summary, the machine-readable protocol,
        // and near-miss lines such as `ok 1 - pass 1` and `# pass 1 of 8`,
        // while still reading both prefixes and CRLF.
        const tally = (key) => {
            const m = stdout.match(new RegExp(String.raw`^.\s*${key}\s+(\d+)\s*$`, "m"));
            assert.ok(m, `could not parse '${key}' from the reporter output:\n${stdout}`);
            return Number(m[1]);
        };

        assert.equal(tally("fail"), 0);
        // Equalities, not `> 0`. The whole failure mode here is a count moving
        // by an amount nobody checks: a guard added or removed upstream should
        // land as a red here and a deliberate decision, not slide through.
        assert.equal(tally("pass"), 1, "install mode must keep exactly the one unguarded, extension-relative assertion (#147)");
        assert.equal(tally("skipped"), 7, "the seven REPO-rooted tests must skip without a checkout, not fail and not run");
    } finally {
        await rm(install.root, { recursive: true, force: true });
    }
});
