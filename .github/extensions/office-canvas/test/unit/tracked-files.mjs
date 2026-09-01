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
 * Only a *successful* answer is memoized. Caching the failure would be the
 * cheaper thing to do and is what the two copies of this helper used to do, but
 * it latches: measured by clearing `PATH`, calling this, restoring `PATH` and
 * calling again, the second call still answered `false` while git was working.
 * Every caller turns a `false` into `t.skip()`, so a single transient failure
 * would silently retire the guards for the rest of the process -- the exact
 * failure mode these callers exist to prevent. The failure path is left
 * uncached so it can correct itself; it costs one fast-failing spawn per call
 * in the only case it happens, an installed extension folder that is not a
 * repository.
 *
 * Callers skip rather than fail when this is false -- an installed extension
 * folder is a plain directory, and these are properties of the repository, not
 * of the shipped artifact.
 *
 * That mode is real, and #104 turned on establishing it, so the evidence is
 * recorded here rather than left as an assertion. There are two install paths
 * and they differ:
 *
 * - **Packaged.** `tools/package-extension.mjs` excludes any path with a `test`
 *   segment (`id: "test"`), so the artefact carries no tests. This suite cannot
 *   run there at all, with or without git.
 * - **Repo-folder.** `install_extension` "copies a folder, not a repository"
 *   (`tools/package-extension.mjs:103`) and its own skip list is `dist` /
 *   `build` / `out` (`:97-98`) -- **`test` is not in it**. The packager excludes
 *   `test` precisely because this path will not. So the installed directory
 *   contains `test/unit/*.test.mjs` and has no `.git`, which is this mode
 *   exactly.
 *
 * Observed on the machine this was written on, against real repo-folder installs
 * under `~/.copilot/installed-plugins/`: one carries `cli/test/unit/` test files,
 * another a `tests/` directory of them, and neither contains a `.git` anywhere
 * -- while
 * `.gitignore` and `.github/` are both present, so dotted entries are being
 * listed and the absence is a fact rather than an artefact of how it was looked
 * for. Not a direct install of *this* folder, which would be stronger; two
 * independent lines of evidence agreeing is what this rests on.
 *
 * **What this justifies, and what it does not.** All of the above is a reason to
 * skip when there is no *repository*. It is not a reason to skip when there is
 * no working *git*, and the bare `catch` below cannot tell those apart: a
 * `PATH` with no git fails the spawn, a plain directory fails the command, and
 * both land here as `false` and become `t.skip()` at all 16 call sites. So a
 * runner whose git installation is broken reports a green suite that executed
 * nothing -- the vacuous pass this file's callers exist to prevent.
 *
 * Node does distinguish them, on the error *shape* rather than on a localized
 * message: a spawn failure carries `code === "ENOENT"`, a non-zero exit carries
 * a numeric `code`. Splitting on that would let the first throw and the second
 * skip, with no call site changed. Not done here, because it is a behaviour
 * change under 16 dependents and the error shape is a platform claim this repo
 * requires a committed probe for. **See #148**, which carries both.
 */
let available = null;
export async function gitAvailable() {
    if (available === true) return available;
    try {
        await execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: REPO });
        available = true;
    } catch {
        available = false;
    }
    return available;
}
