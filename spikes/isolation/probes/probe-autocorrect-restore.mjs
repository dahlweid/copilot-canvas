// Does create_document put the user's autocorrect settings back?
//
// Run: node spikes/isolation/probes/probe-autocorrect-restore.mjs
// Requires Word. Perturbs the user's Word settings and puts them back; the
// found state is captured first and restored in a `finally`.
//
// WHY THIS EXISTS, AND WHY THE SMOKE TEST CANNOT REPLACE IT
//
// `probe-autocorrect.ps1` arm C concluded these settings are per-process. That
// is RETRACTED: it read instance B while A was still alive, and the value is
// not flushed until the writer exits, so a concurrent read cannot tell
// isolation from persistence-with-lag. They persist for the user. So
// create_document now captures, disables, and restores around the authoring
// call, and `suppressed`/`restored` are read-backs rather than "the assignment
// did not throw".
//
// The read-backs are mutation-checked in create-smoke by assigning a wrong
// value, and both go red naming the setting. But there is a gap that mutation
// cannot close, and it is the reason for this file:
//
//   On this machine the found state is already all-`False`, which is also the
//   value suppression writes. So prior == target, and a write that is SKIPPED
//   ENTIRELY leaves the correct value behind and both read-backs stay green.
//
// That vacuity is benign for suppression -- if the setting is already off, not
// writing it is not a defect. It is NOT benign for the restore, because the
// case the restore exists for is a user whose autocorrect is ON: there prior
// != target, a skipped restore leaves our `False` on their machine, and that is
// precisely the defect the whole redesign was for. On an all-`False` machine
// the smoke can never enter that case.
//
// So this probe manufactures it. It sets all five ON first -- standing in for
// the user we are protecting -- and only then runs the tool. Now prior !=
// target in both directions, both read-backs are falsifiable, and the final
// check reads a FRESH instance so it observes what was actually persisted
// rather than what the authoring instance held in memory.
//
// It is a probe rather than a smoke check because it deliberately writes the
// user's real settings, and a suite that dies mid-run would leave them changed.
// The smoke stays non-perturbing beyond what the tool itself does.

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const here = path.dirname(fileURLToPath(import.meta.url));
const extension = path.resolve(here, "../../../.github/extensions/office-canvas");
// pathToFileURL, not the bare path: on Windows a dynamic import of `C:\...`
// fails with ERR_UNSUPPORTED_ESM_URL_SCHEME because the drive letter parses as
// a URL scheme.
const { RenderCache } = await import(pathToFileURL(path.join(extension, "src/render-cache.mjs")).href);

// The five, listed once. Every pass below derives from this, for the same
// reason $script:AC_SETTINGS exists in word-host.ps1: a setting added to one
// pass and not another reads as handled while being unchecked.
const SETTINGS = [
    ["AutoCorrect", "ReplaceText"],
    ["AutoCorrect", "CorrectSentenceCaps"],
    ["AutoCorrect", "CorrectInitialCaps"],
    ["Options", "AutoFormatAsYouTypeReplaceQuotes"],
    ["Options", "AutoFormatAsYouTypeReplaceSymbols"],
];

// No path is interpolated into this script text -- it is static, and the only
// varying part is a boolean rendered by us. A path here would have to travel as
// a discrete argv element instead; see the note in create-smoke.mjs.
const script = (body) => [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-Command",
    `$ErrorActionPreference = 'Stop'
     $w = New-Object -ComObject Word.Application
     try { ${body} } finally {
        $w.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w)
     }`,
];

/**
 * Reads the five from a FRESH Word instance, which is the whole point: a value
 * is not flushed until the writer exits, so reading the instance that wrote it
 * -- or any instance running alongside it -- tells you nothing about what the
 * next Word will see.
 */
async function readFresh() {
    const body = SETTINGS.map(([c, n]) => `"${n}=$($w.${c}.${n})"`).join("\n");
    const { stdout } = await execFileAsync("powershell.exe", script(body));
    const values = {};
    for (const line of stdout.trim().split(/\r?\n/)) {
        const [name, value] = line.split("=");
        if (name) values[name.trim()] = value.trim() === "True";
    }
    return values;
}

async function writeAll(values) {
    const body = SETTINGS.map(([c, n]) => `$w.${c}.${n} = $${values[n] ? "true" : "false"}`).join("\n");
    await execFileAsync("powershell.exe", script(body));
}

const names = SETTINGS.map(([, n]) => n);
const all = (values, expected) => names.every((n) => values[n] === expected);
const show = (values) => names.map((n) => `${n}=${values[n]}`).join(" ");

const workRoot = await mkdtemp(path.join(tmpdir(), "word-ac-probe-"));
let found = null;
let failures = 0;

const step = (label, ok, detail) => {
    if (!ok) failures += 1;
    process.stdout.write(`${ok ? "ok  " : "FAIL"}  ${label}${detail ? ` -- ${detail}` : ""}\n`);
};

try {
    found = await readFresh();
    process.stdout.write(`found state: ${show(found)}\n`);

    // Stand in for a user who has autocorrect ON. Without this the probe runs
    // on prior == target and proves nothing the smoke does not already cover.
    await writeAll(Object.fromEntries(names.map((n) => [n, true])));
    const armed = await readFresh();
    step("all five read back ON from a fresh instance before the tool runs", all(armed, true), show(armed));
    assert.ok(all(armed, true), "could not arm the probe; the rest would be vacuous");

    const cache = new RenderCache({ cacheRoot: path.join(workRoot, "artifacts"), log: () => {} });
    let created;
    try {
        created = await cache.createDocument(path.join(workRoot, "probe.docx"), {
            blocks: [{ kind: "paragraph", text: 'It said "quote" -- and it stayed.' }],
        });
    } finally {
        await cache.dispose?.().catch(() => {});
    }

    const ac = created.autoCorrect ?? {};
    step("the tool reports it suppressed autocorrect", ac.suppressed === true, `${ac.reason ?? ""} ${(ac.settings ?? []).join(",")}`);
    step("the tool reports it restored the settings", ac.restored === true, `${ac.restoreReason ?? ""} ${(ac.restoreSettings ?? []).join(",")}`);

    // The captured prior is what the restore writes back, so if it did not
    // observe the ON state the restore cannot be correct even when it claims
    // success -- it would faithfully put back the wrong values.
    step(
        "it captured the ON values, not the values it was about to write",
        ac.prior && names.every((n) => ac.prior[n] === true),
        JSON.stringify(ac.prior ?? null),
    );

    // The load-bearing check. A fresh instance sees what was persisted, so this
    // is the user's next Word, not ours.
    const after = await readFresh();
    step("a fresh instance still reads all five ON after the tool ran", all(after, true), show(after));
} finally {
    if (found) {
        await writeAll(found).catch(() => {});
        const back = await readFresh().catch(() => null);
        const ok = back && names.every((n) => back[n] === found[n]);
        step("the probe put the machine back the way it found it", Boolean(ok), back ? show(back) : "unreadable");
    }
    await rm(workRoot, { recursive: true, force: true }).catch(() => {});
}

process.stdout.write(failures === 0 ? "\nall checks passed\n" : `\n${failures} check(s) failed\n`);
process.exitCode = failures === 0 ? 0 : 1;
