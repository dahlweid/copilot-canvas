// The two failure shapes `gitAvailable()` has to tell apart, measured rather
// than asserted -- and the memoization behaviour of the throw it now raises.
//
// #148 asked for a committed probe under `spikes/` for the `ENOENT`-vs-numeric
// claim. This file is that measurement written as a test instead, and the
// reason is not convenience. `spikes/` exists for claims CI *cannot* check --
// the ones needing an installed, licensed Office. This claim needs neither
// Office nor Windows: it is `node:child_process` behaviour, it is deterministic,
// and it runs on `ubuntu-latest` unchanged. A probe would be re-run by hand when
// someone remembered; this is re-run on every commit, which is strictly the
// stronger guarantee. The measurement is recorded here rather than lost:
//
//   ARM 1  spawn failure (no such binary):     code = "ENOENT", typeof string
//   ARM 2  git ran and exited non-zero:        code = 128,      typeof number, stderr populated
//
// Both arms are asserted below, so the table cannot go stale the way a comment
// quoting a probe's old output can.
//
// Office-free.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

import { REPO, gitAvailable } from "./tracked-files.mjs";

const execFileAsync = promisify(execFile);

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

// Last, and deliberately: it mutates `PATH` for the duration, and it settles
// `gitAvailable()`'s module-level memo. Top-level tests in `node:test` run one
// at a time, so nothing above can observe either.
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

    assert.ok(thrown, "gitAvailable() answered instead of throwing while git was unreachable; a broken git would skip all 16 guarded tests");
    assert.match(thrown.message, /no exit status/);
    assert.equal(thrown.cause?.code, "ENOENT");

    // The memoization interaction #148 asks about. Only success is cached, and
    // the throw happens before any assignment, so this must re-ask and get the
    // real answer rather than a latched failure.
    assert.equal(await gitAvailable(), inRepo);
});
