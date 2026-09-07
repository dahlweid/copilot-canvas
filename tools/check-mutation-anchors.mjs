#!/usr/bin/env node
// Static anchor resolution for the mutation harnesses under
// `.github/extensions/office-canvas/test/unit/`.
//
// ## What it checks, and the narrower claim a green run makes
//
// Each harness defines a `$mutants` array whose entries carry a `from` string
// and the `file` it is expected to occur in. The harness replaces `from` with
// `to` and requires the suite to go red. If `from` does not occur, nothing is
// replaced and the mutant measures nothing while still being counted; if it
// occurs more than once, PowerShell's `String.Replace` changes every site, so a
// kill cannot be attributed to the defect the mutant is named for. Two of the
// four harnesses test only for *presence*, so the second case is unchecked
// there today.
//
// So this asserts one thing per mutant: `from` occurs in `file` **exactly
// once**. That is anchor resolution and nothing else. It does **not** establish
// that a mutant is killed -- that needs the mutation applied and the suite run
// once per mutant, on Windows with PowerShell 5.1, which is issue #182's
// remaining question and stays open. The two failures are distinct: this one is
// the silent half, and it is the half that is pure text.
//
// It would have caught all three anchors that sat dead on `main` -- two in
// `mutate-webview.ps1` against `pdf-view.mjs`, repaired by #181, and
// `mutate-create.ps1`'s `creatable extensions spelled out`, retired by a
// rewording in #165 and dead from `fc50f98` until it was re-anchored.
//
// ## Why it asks the harnesses instead of reading them
//
// It runs each harness with `-ListMutants`, which prints that harness's own
// array as JSON and exits before any suite runs. The harness stays the single
// source of truth for both the array and the root each `file` resolves against.
//
// Parsing the `.ps1` text instead was tried and produced three confidently
// wrong answers, which is why the switch exists:
//
//   * a text regex found 29 mutants in `mutate-pdfjs.ps1`, which has 62 -- a
//     third of the corpus dropped under a total that looked authoritative;
//   * not every `from` is a literal. `mutate-create.ps1`'s BLOCK_HELP anchor is
//     an expression, `'...' + $dash + '...' + [char]0x60 + ','`, and capturing
//     only its first quoted fragment reported a bogus AMBIGUOUS;
//   * the harnesses use two path conventions. `mutate-pdfjs.ps1`'s entries are
//     repo-root-relative and carry `at = 'repo'`; the other three are
//     extension-relative. One root for all of them gave 33 bogus missing files.
//
// All three are the same mistake -- a second parser of data PowerShell already
// parses -- which is why the remedy is to remove the parser rather than harden
// it.
//
// ## Office-free
//
// No Word, no PowerPoint, no mutation, no suite run. It needs PowerShell, which
// on a hosted `ubuntu-latest` runner means the preinstalled `pwsh`. Reading the
// arrays under `pwsh` rather than Windows PowerShell 5.1 is sound here for a
// reason worth stating: only `mutate-create.ps1` is BOM-less, where 5.1 and
// `pwsh` could disagree, and its only non-ASCII byte sequence sits in a comment
// -- every anchor in it is built from code points (`$dash`, `[char]0x60`)
// precisely so that the decoding cannot change what it says.
//
// Run: node tools/check-mutation-anchors.mjs
//      node tools/check-mutation-anchors.mjs --harness-dir <dir>

import { spawnSync } from "node:child_process";
import { readdir, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const REPO = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DEFAULT_HARNESS_DIR = path.join(REPO, ".github", "extensions", "office-canvas", "test", "unit");

/**
 * `pwsh` first everywhere, because it is what a hosted Linux runner has and
 * what a modern Windows box has as well. Windows PowerShell is the fallback on
 * Windows only -- it is the interpreter that actually runs the harnesses there,
 * so falling back to it keeps the gate runnable on a machine that never
 * installed 7.
 */
const SHELLS = process.platform === "win32" ? ["pwsh", "powershell"] : ["pwsh"];

function resolveShell() {
    const tried = [];
    for (const exe of SHELLS) {
        // A constant argument, not an interpolated one: nothing user-supplied
        // is ever assembled into a PowerShell command line here or below.
        const probe = spawnSync(exe, ["-NoProfile", "-NonInteractive", "-Command", "exit 0"], { encoding: "utf8" });
        if (!probe.error) return exe;
        tried.push(`${exe} (${probe.error.code ?? probe.error.message})`);
    }
    throw new Error(`no PowerShell available; tried ${tried.join(", ")}`);
}

/**
 * Ask one harness for its own mutants.
 *
 * Every value is passed as a discrete argv element -- no shell, no `cmd.exe`,
 * no `-Command` carrying an interpolated path. A harness living under a
 * directory with an apostrophe or an `&` in its name is then just an argument.
 */
function listMutants(shell, harnessPath) {
    const run = spawnSync(shell, ["-NoProfile", "-NonInteractive", "-File", harnessPath, "-ListMutants"], {
        encoding: "utf8",
        // The four harnesses together emit well under a megabyte, but the
        // default 1 MiB truncates silently rather than failing, and truncated
        // JSON would arrive as a parse error naming the wrong cause.
        maxBuffer: 64 * 1024 * 1024,
    });

    if (run.error) throw new Error(`could not run ${path.basename(harnessPath)}: ${run.error.message}`);
    if (run.status !== 0) {
        throw new Error(
            `${path.basename(harnessPath)} -ListMutants exited ${run.status}\n${run.stderr || run.stdout}`.trim(),
        );
    }

    let parsed;
    try {
        parsed = JSON.parse(run.stdout);
    } catch (err) {
        throw new Error(`${path.basename(harnessPath)} -ListMutants did not print JSON: ${err.message}`);
    }

    // Windows PowerShell's `ConvertTo-Json` unwraps a one-element array, so a
    // harness that ever held a single mutant would arrive as a bare object.
    return Array.isArray(parsed) ? parsed : [parsed];
}

/**
 * Non-overlapping occurrences, which is what both the harnesses and the check
 * they would perform count: .NET's `String.Replace` and `Regex.Matches` each
 * advance past a match rather than by one character.
 */
function countOccurrences(haystack, needle) {
    let count = 0;
    let at = haystack.indexOf(needle);
    while (at !== -1) {
        count += 1;
        at = haystack.indexOf(needle, at + needle.length);
    }
    return count;
}

/** One line, bounded, so a multi-line anchor cannot flood the report. */
function quoteAnchor(from) {
    const [first = ""] = String(from).split("\n");
    const trimmed = first.replace(/\r$/, "");
    const shown = trimmed.length > 96 ? `${trimmed.slice(0, 96)}...` : trimmed;
    return String(from).includes("\n") ? `${shown} [+more lines]` : shown;
}

/**
 * The audit itself, over already-listed mutants. Separated from the I/O above
 * so it can be driven from a test with anchors chosen to fail.
 *
 * `read` takes an absolute path and returns the file's text, or null if it
 * cannot be read.
 */
export async function auditMutants(mutants, read) {
    const problems = [];
    const sources = new Map();

    for (const m of mutants) {
        const where = `${m.harness} / ${m.name}`;

        if (!sources.has(m.path)) sources.set(m.path, await read(m.path));
        const source = sources.get(m.path);

        if (source === null || source === undefined) {
            problems.push({ kind: "UNREADABLE", where, detail: `cannot read ${m.file}` });
            continue;
        }
        // The harnesses treat this as a stale anchor too, and they are right to:
        // the replacement writes the file back unchanged, so the suite is run
        // against unmutated source and reports a kill nobody made.
        if (m.from === m.to) {
            problems.push({ kind: "NO-OP", where, detail: `from and to are identical: ${quoteAnchor(m.from)}` });
            continue;
        }

        const occurrences = countOccurrences(source, m.from);
        if (occurrences === 0) {
            problems.push({ kind: "DEAD", where, detail: `not found in ${m.file}: ${quoteAnchor(m.from)}` });
        } else if (occurrences > 1) {
            problems.push({
                kind: "AMBIGUOUS",
                where,
                detail: `matches ${occurrences}x in ${m.file}, so the mutation would change every site: ${quoteAnchor(m.from)}`,
            });
        }
    }

    return problems;
}

async function main(argv) {
    let harnessDir = DEFAULT_HARNESS_DIR;
    for (let i = 0; i < argv.length; i += 1) {
        if (argv[i] === "--harness-dir") {
            harnessDir = path.resolve(argv[i + 1] ?? "");
            i += 1;
        } else {
            throw new Error(`unrecognised argument: ${argv[i]}`);
        }
    }

    const names = (await readdir(harnessDir)).filter((n) => n.startsWith("mutate-") && n.endsWith(".ps1")).sort();

    // A gate that finds nothing and exits 0 is the failure this whole corpus
    // exists to prevent, so an empty harness set is an error rather than a
    // vacuous pass.
    if (names.length === 0) {
        console.error(`no mutation harnesses found in ${harnessDir}; the path moved and nothing was checked`);
        return 1;
    }

    const shell = resolveShell();
    const all = [];
    const perHarness = [];
    for (const name of names) {
        const listed = listMutants(shell, path.join(harnessDir, name));
        if (listed.length === 0) {
            console.error(`${name} listed no mutants; its array moved or -ListMutants no longer reaches it`);
            return 1;
        }
        perHarness.push([name, listed.length]);
        all.push(...listed);
    }

    const problems = await auditMutants(all, async (p) => {
        try {
            return await readFile(p, "utf8");
        } catch {
            return null;
        }
    });

    console.log(`mutation anchors: ${all.length} across ${names.length} harnesses, via ${shell}`);
    const width = Math.max(...perHarness.map(([n]) => n.length));
    for (const [name, count] of perHarness) console.log(`  ${name.padEnd(width)}  ${String(count).padStart(3)}`);

    if (problems.length === 0) {
        // Stated as narrowly as it is true. Every anchor resolving is not every
        // mutant being killed, and a reader who takes the second from this line
        // is the reason the first is worth printing carefully.
        console.log(`\nEvery anchor resolves exactly once. This says the mutations would be applied, not that any is killed.`);
        return 0;
    }

    console.error(`\n${problems.length} anchor(s) would measure nothing:`);
    for (const p of problems) console.error(`  ${p.kind}  ${p.where}\n      ${p.detail}`);
    console.error(
        `\nDEAD and UNREADABLE mean the mutation is never applied; AMBIGUOUS means it is applied more widely than its name claims; NO-OP means the file is rewritten unchanged. All four are counted as mutants and none is evidence.`,
    );
    return 1;
}

// Only when run as a script: the audit is imported by its test.
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
    main(process.argv.slice(2)).then(
        (code) => process.exit(code),
        (err) => {
            console.error(err.message);
            process.exit(1);
        },
    );
}
