// Does create_document leave the user's autocorrect settings alone?
//
// Run: node spikes/isolation/probes/probe-autocorrect-untouched.mjs
// Requires Word. Perturbs the user's Word settings and puts them back; the
// found state is captured first and restored in a `finally`.
//
// WHAT THIS ASKED BEFORE, AND WHY THE QUESTION CHANGED
//
// This file used to ask whether create_document RESTORED the five settings it
// suppressed. It no longer suppresses them, so the question is now the simpler
// and stronger one: does it touch them at all?
//
// The history is worth keeping because the file only exists because of it.
// `probe-autocorrect.ps1` arm C concluded these settings are per-process. That
// is RETRACTED: it read instance B while A was still alive, and a concurrent
// reader sees the pre-write value, so it cannot tell isolation from
// persistence-with-lag. They persist for the user. The response at the time was
// capture-and-restore around the authoring call, and this probe was written to
// prove the restore, because the smoke test could not:
//
//   On this machine the found state was all-`False`, which was also the value
//   suppression wrote. So prior == target, and a write that was SKIPPED
//   ENTIRELY left the correct value behind and every read-back stayed green.
//   The case the restore existed for -- a user whose autocorrect is ON -- could
//   not occur in the suite at all.
//
// That manufacture is exactly what is still needed, so it is kept verbatim: set
// all five ON first, standing in for the user we are protecting, and only then
// run the tool. What changed is the assertion at the end. A restore that puts
// back what it captured and a tool that never wrote anything are
// indistinguishable from the outside -- which is the point. The user's next
// Word must read ON either way, and now nothing has to go right for that to
// happen.
//
// The final read is from a FRESH instance, which is the whole point: an
// instance running alongside the writer reads the pre-write value, so it
// observes what was persisted rather than what the authoring instance held in
// memory.
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

// The five, listed once. Every pass below derives from this: a setting added to
// one pass and not another reads as handled while being unchecked.
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
 * Reads the five from a FRESH Word instance, which is the whole point: an
 * instance running alongside the writer reads the pre-write value, so reading
 * the instance that wrote it -- or any instance running alongside it -- tells
 * you nothing about what the next Word will see. Only a reader started after
 * the writer exited answers that question.
 *
 * Why the wording is careful: "the value is not flushed until the writer exits"
 * is a mechanism, and nothing measured here separates it from the reader having
 * cached at startup. The probe depends on the observation, not the mechanism.
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

    // Nothing about autocorrect crosses the tool boundary any more, and that is
    // asserted rather than assumed: a reappearing report field is the visible
    // symptom of a reappearing suppression.
    step(
        "the tool reports nothing about autocorrect",
        created.autoCorrect === undefined,
        JSON.stringify(created.autoCorrect ?? null),
    );

    // The load-bearing check. A fresh instance sees what was persisted, so this
    // is the user's next Word, not ours. It reads ON because the tool never
    // wrote anything -- not because a restore put it back.
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
