// `Application.Quit` must be called with no argument.
//
// Under Windows PowerShell 5.1 -- the runtime every `.ps1` in this repo is
// spawned under -- the argument form does not bind. `Quit` declares its three
// parameters as `VARIANT*`, so PowerShell's binder refuses the value and throws
// before the COM call happens:
//
//   Argument: "1" muss System.Management.Automation.PSReference sein.
//
// Every site that had it also wrapped it in `try { ... } catch { }`, so the
// throw was swallowed and the probe printed a tidy cleanup section while the
// `WINWORD.EXE` it had created stayed up. Process exit does not reap it either:
// measured in `spikes/isolation/probes/probe-quit0-leak.ps1`, which runs the two
// arms in a child that then exits, and reports the `Quit(0)` arm's Word still
// alive 30 s later while the `Quit()` control is gone. Re-run at 6dc3369: pid
// 38820 leaked, pid 10888 reaped.
//
// There is no literal/variable split -- `Quit($var)` throws exactly as `Quit(0)`
// does (PLAN.md 20). So this cannot be checked by eye for "looks like a
// constant"; the argument form is wrong in every spelling but `[ref]`.
//
// Why a source assertion rather than a leak assertion in the integration
// suites: the sites are probes, which CI cannot run and which nobody runs all
// of. A leak assertion only fires on the probe someone happened to execute,
// which is precisely the reason this survived -- it leaks on probe runs, not on
// product runs, so the population that would have noticed never ran the code.
//
// Office-free.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { blankComments } from "./ps-encoding-rule.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");

// The one file allowed to use the argument form, because measuring the throw is
// what it is for. Exempting it by name is not a hole: the count below pins how
// many times it may do so, so a second site added here still reddens.
const MEASURES_THE_THROW = "spikes/isolation/probes/probe-quit0-leak.ps1";
const MEASURED_OCCURRENCES = 1;

/**
 * Tracked `.ps1` files, asked of git rather than the filesystem.
 *
 * Two reasons. A worktree can carry untracked scratch scripts that are nobody's
 * business, and this repo's probes are split across sibling worktrees, so the
 * filesystem answers a different question than "what is in this commit".
 */
async function trackedPowerShellFiles() {
    const { stdout } = await execFileAsync("git", ["ls-files", "-z", "*.ps1"], {
        cwd: REPO,
        maxBuffer: 8 * 1024 * 1024,
    });
    return stdout.split("\0").filter(Boolean);
}

// Skip rather than fail where there is no git to ask -- an installed extension
// folder is a plain directory, and this property is about the repository.
let available = null;
async function gitAvailable() {
    if (available !== null) return available;
    try {
        await execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: REPO });
        available = true;
    } catch {
        available = false;
    }
    return available;
}

/**
 * The script with its line comments blanked, so a call named in prose is not
 * mistaken for a call. Two of the three surviving mentions of `Quit(0)` in this
 * tree are comments explaining the defect, and a grep over raw text would report
 * the explanation as the bug.
 *
 * Imported from `ps-encoding-rule.mjs`, which owns the one copy: blanking is
 * offset-preserving there, which is what lets the failure message below carry a
 * line number a reader can go to.
 *
 * Here-strings are deliberately *not* stripped: several probes build their
 * worker as a single-quoted here-string, so the code that actually calls `Quit`
 * lives inside one. Skipping them would exempt exactly the sites that leaked.
 */

/** Every `.Quit(` call in `source` whose argument list is not empty and not `[ref]`. */
function argumentFormCalls(source) {
    const found = [];
    const re = /\.Quit\(([^)]*)\)/g;
    let m;
    while ((m = re.exec(source)) !== null) {
        const arg = m[1].trim();
        if (arg === "") continue;
        // `[ref]$x` binds correctly and is the documented escape hatch for a
        // genuinely non-default wdSaveChanges. Nothing in the tree needs it
        // today; allowing it keeps this guard from banning the correct remedy.
        if (/^\[ref\]/i.test(arg)) continue;
        found.push({ arg, line: source.slice(0, m.index).split("\n").length });
    }
    return found;
}

test("no .ps1 calls Application.Quit with a bare argument", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    const files = await trackedPowerShellFiles();
    // A file list that came back empty would make every assertion below pass
    // without examining anything -- the same shape as a glob that matches
    // nothing and exits 0. This repo has measured that failure once already.
    assert.ok(files.length > 0, "git ls-files found no .ps1 files; the enumeration is broken, not the tree");

    const offenders = [];
    for (const file of files) {
        if (file.replace(/\\/g, "/") === MEASURES_THE_THROW) continue;
        const source = blankComments(await readFile(path.join(REPO, file), "utf8"));
        for (const call of argumentFormCalls(source)) {
            offenders.push(`${file}:${call.line}  .Quit(${call.arg})`);
        }
    }

    assert.deepEqual(
        offenders,
        [],
        "Application.Quit(<arg>) does not bind under Windows PowerShell 5.1: it throws, and the swallowing catch " +
            "around it turns that into a leaked WINWORD.EXE. Use the no-argument Quit(), which takes the same " +
            "default. See PLAN.md 20 and spikes/isolation/probes/probe-quit0-leak.ps1.\n  " +
            offenders.join("\n  "),
    );
});

test("the probe that measures the throw still contains exactly the calls it needs", async (t) => {
    if (!(await gitAvailable())) return t.skip("no git checkout to enumerate");

    // Without this, the exemption above is unbounded: any number of new
    // argument-form calls could be added to that file and nothing would notice.
    // Pinning the count is what keeps the one allowed exception able to go red.
    const files = (await trackedPowerShellFiles()).map((f) => f.replace(/\\/g, "/"));
    assert.ok(
        files.includes(MEASURES_THE_THROW),
        `${MEASURES_THE_THROW} is not in the tree, so the exemption above now covers nothing and would hide a real site`,
    );

    const source = blankComments(await readFile(path.join(REPO, MEASURES_THE_THROW), "utf8"));
    const calls = argumentFormCalls(source);
    assert.equal(
        calls.length,
        MEASURED_OCCURRENCES,
        `${MEASURES_THE_THROW} is exempt because one arm must call Quit(0) to measure that it throws. ` +
            `Expected exactly ${MEASURED_OCCURRENCES} such call, found ${calls.length}: ` +
            calls.map((c) => `line ${c.line} .Quit(${c.arg})`).join("; "),
    );
});
