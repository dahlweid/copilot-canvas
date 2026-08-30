// Enumerating the repository's tracked files, for the guards that assert a
// property over "every script in the tree" rather than over a named list.
//
// Extracted because two tests grew their own copy and the copies had already
// diverged: one normalised path separators inside the helper, the other did it
// at each call site and missed one of them. That is the same failure mode as the
// regex in `ps-encoding-rule.mjs` -- two records of one quantity, free to
// disagree -- caught here before it produced a wrong answer rather than after.
//
// Separate from `ps-encoding-rule.mjs` on purpose: that file is the definition
// of the encoding rule, and enumeration is not part of that rule. Folding this
// into it would mean the next guard needing a file list has to import a module
// named for something it does not use.
//
// Not a `.test.mjs`, so the CI glob does not try to run it -- same convention as
// `tool-bridge.mjs`.
//
// Office-free.

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import path from "node:path";
import { fileURLToPath } from "node:url";

const execFileAsync = promisify(execFile);

/** The repository root, five levels up from `test/unit/`. */
export const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..", "..", "..", "..", "..");

/**
 * Tracked files matching a git pathspec, as repo-relative paths with forward
 * slashes.
 *
 * Asked of git rather than the filesystem for two reasons. A worktree can carry
 * untracked scratch scripts that are nobody's business, and this repo's probes
 * are split across sibling worktrees, so the filesystem answers a different
 * question than "what is in this commit".
 *
 * `-z` because a path may contain a character git would otherwise quote, and a
 * quoted path silently stops matching the constants the callers compare against.
 *
 * Separators are normalised here, once. Both callers compare against constants
 * written with forward slashes, and doing it at the call sites is what let the
 * two copies of this function disagree.
 *
 * @param {string} pathspec e.g. `"*.ps1"`
 * @returns {Promise<string[]>}
 */
export async function trackedFiles(pathspec) {
    const { stdout } = await execFileAsync("git", ["ls-files", "-z", pathspec], {
        cwd: REPO,
        maxBuffer: 8 * 1024 * 1024,
    });
    return stdout
        .split("\0")
        .filter(Boolean)
        .map((f) => f.replace(/\\/g, "/"));
}

/**
 * Whether there is a git checkout to ask at all.
 *
 * Memoized per process: every test in both guards calls it, and the answer
 * cannot change during a run.
 *
 * Callers skip rather than fail when this is false -- an installed extension
 * folder is a plain directory, and these are properties of the repository, not
 * of the shipped artifact.
 */
let available = null;
export async function gitAvailable() {
    if (available !== null) return available;
    try {
        await execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: REPO });
        available = true;
    } catch {
        available = false;
    }
    return available;
}
