// Probe: how does Node report a child process that could not be spawned, versus
// one that ran and exited non-zero?
//
// `gitAvailable()` in `.github/extensions/office-canvas/test/unit/tracked-files.mjs`
// has to tell those two apart, because they mean opposite things and want
// opposite responses. A `PATH` with no git means the environment is broken and
// the guards must fail loudly; a plain directory that is not a repository is the
// installed-extension mode the guards exist to skip in. Before #148 a bare
// `catch` collapsed both into "no repository here", so a runner with a broken
// git reported a green suite that executed nothing at 17 call sites.
//
// The split has to be made on the error **shape**, never on the message: this
// repo is developed on a German-language Windows, where both the platform's
// errors and git's own stderr are localized, and discriminating on message text
// is a defect class this repo has been bitten by before.
//
// Run: node spikes/git-guard/probes/probe-execfile-error-shapes.mjs
//
// No Office, no Windows, no network, no browser: pure `node:child_process`.
// This one is genuinely re-runnable anywhere, and the guard it backs is also
// asserted on every commit by `test/unit/git-available.test.mjs`. The probe is
// still the citable artefact -- `tools/check-citations.mjs` tracks
// `probe-*.{ps1,mjs}` paths and nothing else, so a claim discharged onto a
// neighbouring test is an unchecked citation (CONTEXT.md). It is also the arm
// that can be pointed at a *different* platform: the unit test measures the
// machine CI happens to run on, and this can be run by hand on another.
//
// Exit codes:
//   0  both arms measured and they are distinguishable
//   1  both arms measured and they are NOT distinguishable -- the guard's
//      premise is false on this platform and #148's fix does not hold here
//   2  an arm could not be measured at all, so nothing was learned. Deliberately
//      distinct from 1: "the claim is false" and "I failed to look" are the two
//      things this repo most insists on not collapsing.

import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";

const execFileAsync = promisify(execFile);

// A name no `PATH` entry can hold. `execFile` does not go through a shell, so
// this is a literal filename lookup and nothing expands it.
const NO_SUCH_BINARY = "git-that-does-not-exist-8f3a1c";

async function failureOf(promise) {
    try {
        await promise;
        return null;
    } catch (err) {
        return err;
    }
}

function describe(err) {
    return {
        code: String(err.code),
        typeofCode: typeof err.code,
        stderrBytes: typeof err.stderr === "string" ? err.stderr.length : null,
        // Reported, never matched on. It is here so a reader can see *why*
        // matching on it would be a mistake -- on this machine it arrives in
        // German -- rather than being told so.
        message: err.message.split("\n")[0],
    };
}

async function armSpawnFailure() {
    const err = await failureOf(execFileAsync(NO_SUCH_BINARY, ["--version"]));
    if (!err) return { ok: false, why: `spawning ${NO_SUCH_BINARY} succeeded; the name is not unused after all` };
    return { ok: true, observed: describe(err) };
}

async function armNonZeroExit() {
    const dir = await mkdtemp(path.join(tmpdir(), "git-guard-probe-"));
    try {
        // Without a ceiling, git walks upward and this arm passes or fails
        // depending on whether the machine's temp directory happens to sit
        // inside a repository -- which would make the probe's answer a property
        // of the machine's layout rather than of git.
        const env = { ...process.env, GIT_CEILING_DIRECTORIES: path.dirname(dir) };
        const err = await failureOf(execFileAsync("git", ["rev-parse", "--git-dir"], { cwd: dir, env }));
        if (!err) return { ok: false, why: `git rev-parse --git-dir succeeded in ${dir}; the ceiling did not hold` };
        if (typeof err.code === "string") {
            // This is arm 1 wearing arm 2's clothes. Reporting it as a
            // measurement of a non-zero exit would name a cause that was never
            // observed.
            return { ok: false, why: `git is not runnable here (${err.code}), so a non-zero exit cannot be measured` };
        }
        return { ok: true, observed: describe(err) };
    } finally {
        await rm(dir, { recursive: true, force: true });
    }
}

const arms = [
    ["ARM1 spawn failure (no such binary)", await armSpawnFailure()],
    ["ARM2 non-repo directory (git present, exits non-zero)", await armNonZeroExit()],
];

let unmeasured = false;
for (const [label, result] of arms) {
    console.log(`\n${label}`);
    if (!result.ok) {
        unmeasured = true;
        console.log(`  NOT MEASURED: ${result.why}`);
        continue;
    }
    for (const [k, v] of Object.entries(result.observed)) console.log(`  ${k.padEnd(12)} ${v}`);
}

console.log("");
if (unmeasured) {
    console.log("VERDICT: at least one arm was not measured. Nothing is established either way.");
    process.exit(2);
}

const [spawnFailure, nonZeroExit] = arms.map(([, r]) => r.observed);
const distinguishable = spawnFailure.typeofCode !== nonZeroExit.typeofCode;
console.log(
    distinguishable
        ? `VERDICT: distinguishable on typeof err.code -- ${spawnFailure.typeofCode} vs ${nonZeroExit.typeofCode}. ` +
              `gitAvailable() may treat a numeric code as "git answered: not a repository" and anything else as ` +
              `"git never reported an exit status".`
        : `VERDICT: NOT distinguishable -- both arms report typeof err.code === ${spawnFailure.typeofCode}. ` +
              `The premise #148's fix rests on does not hold on this platform.`,
);
process.exit(distinguishable ? 0 : 1);
