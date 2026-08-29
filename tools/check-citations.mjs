#!/usr/bin/env node
// Every probe cited by name in a tracked file must resolve to exactly one probe
// that exists in *this* repository.
//
// This exists because of a specific, repeatable mistake rather than a general
// tidiness worry. Work here is split across sibling git worktrees — one per
// stack layer, all under the same parent directory — and the coordinator runs
// other sessions' probes to verify their claims rather than accepting them on
// report. A probe run that way is real: it executes, it prints a table, the
// table goes into an ADR. But it lives on *another branch*, so the citation
// that quotes it dangles for everyone reading `main`.
//
// That is exactly what happened: five files landed citing a probe that was
// never on `main` and still resolved on this machine, under a sibling
// worktree's `spikes\isolation\probes\`.
//
// So the check must not ask the filesystem. `Test-Path` (or `fs.existsSync`
// against an absolute path built from a sibling worktree) reproduces the bug
// instead of catching it. It asks **git**, for files tracked in this
// repository, which cannot see a sibling worktree's branch content.
//
// Scope is deliberately narrow: filenames matching probe-*.ps1 / probe-*.mjs.
// Those are the repo's evidence trail — they are what ADRs cite when they say
// a claim was measured — and a dangling one silently converts a measurement
// into an assertion.
//
// ## Resolving a citation, and why "it exists somewhere" is not enough
//
// The first version of this checker keyed on **basename alone**. That is too
// weak, and measurably so: this repo has three colliding basenames —
// `probe-export.ps1`, `probe-hide.ps1` and `probe-bulk-read.ps1` each exist
// under both `spikes/isolation/probes/` and `spikes/powerpoint/probes/`. Under
// basename matching, a citation naming a *full path* that does not exist still
// passed, because the same basename existed in the other directory. The guard
// would have reported OK on precisely the dangling reference it was written to
// catch.
//
// So a citation resolves by the suffix the author actually wrote, matched on
// directory boundaries:
//
//   - 0 candidates                 -> dangling; the probe does not ship here.
//   - exactly 1 candidate          -> resolved.
//   - more than 1 candidate        -> ambiguous, then disambiguated by
//                                     proximity: the candidate sharing the
//                                     longest directory prefix with the citing
//                                     file wins, provided it wins outright.
//
// Proximity is not a convenience. A bare `probe-export.ps1` inside
// `spikes/powerpoint/` means the PowerPoint one to every human who reads it,
// and writing a repo-root path to name a sibling file would make those
// documents worse. But it deliberately does **not** rescue a citation from
// somewhere else in the tree: an ADR or `CONTEXT.md` naming a colliding
// basename has no nearest candidate, stays ambiguous, and must cite the path.
// That is the case where a reader genuinely cannot tell which probe is the
// evidence.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { join } from "node:path";

// The whole written citation, including any directory segments in front of it,
// so the check can tell `spikes/powerpoint/probes/probe-export.ps1` from a bare
// `probe-export.ps1`. Both separators, because prose here quotes Windows paths.
const CITATION = /(?:[A-Za-z0-9._-]+[/\\])*probe-[A-Za-z0-9._-]*\.(?:ps1|mjs)\b/g;

// Binary and vendored paths have no citations and would only add noise.
const SKIP = /^(?:\.git\/|.*\.(?:png|jpg|jpeg|gif|pdf|docx|pptx|xlsx|zip|ico)$)/i;

// This file names probes in the commentary above, and describes the very
// collisions it resolves. Exempting the checker by path keeps the exemption
// from spreading: no other file gets to opt out.
const EXEMPT = "tools/check-citations.mjs";

function git(...args) {
  return execFileSync("git", args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
}

// Everything is anchored to the repo root rather than the current directory.
// `git ls-files` scopes itself to the directory it runs in *and* reports paths
// relative to it, so running from a subdirectory silently changes both the set
// of files searched and the set of probes considered to exist. Measured, both
// halves go wrong at once and in opposite directions: from `docs/` the probe
// list is empty so every correct citation is reported dangling, and from
// `spikes/isolation/` a genuinely dangling citation at the repo root is not
// seen at all. The second is the dangerous one, because a guard that reports
// OK is indistinguishable from a healthy repo. CI does not set a
// `working-directory` today, so this was latent -- but it is one workflow edit
// away from silently passing forever.
const ROOT = git("rev-parse", "--show-toplevel").trim();

const tracked = git("-C", ROOT, "ls-files", "-z").split("\0").filter(Boolean);

// Every probe that actually ships, by full repo-relative path. Keeping the
// paths rather than the basenames is the fix: the set has to be able to say
// *which* probe, not merely that something by that name exists.
const probePaths = tracked.filter((p) => /(?:^|\/)probe-[^/]*\.(?:ps1|mjs)$/.test(p));

/** Directory of a repo-relative path, "" for a file at the root. */
function dirOf(p) {
  const i = p.lastIndexOf("/");
  return i === -1 ? "" : p.slice(0, i);
}

/** How many leading directory segments two paths share. */
function sharedDepth(a, b) {
  const x = a === "" ? [] : a.split("/");
  const y = b === "" ? [] : b.split("/");
  let n = 0;
  while (n < x.length && n < y.length && x[n] === y[n]) n++;
  return n;
}

/**
 * Probes whose path ends with the written citation on a directory boundary.
 *
 * The boundary test is what stops `probe-export.ps1` from matching a
 * hypothetical `probe-export.ps1` suffix of `my-probe-export.ps1`, and stops a
 * partial `probes/probe-x.ps1` from matching `other-probes/probe-x.ps1`.
 */
function candidatesFor(written) {
  const want = written.replace(/\\/g, "/").replace(/^\.\//, "");
  return probePaths.filter((p) => p === want || p.endsWith("/" + want));
}

const dangling = [];
const ambiguous = [];
const unreadable = [];

for (const path of tracked) {
  if (path === EXEMPT || SKIP.test(path)) continue;

  let text;
  try {
    text = readFileSync(join(ROOT, path), "utf8");
  } catch (err) {
    // A tracked path can legitimately be unreadable: a submodule lists as a
    // directory (EISDIR), and a sparse or partial checkout lists files that are
    // not materialized (ENOENT). Those are skipped -- but they are *reported*,
    // never swallowed. A guard that prints OK while having quietly declined to
    // look at part of the repo is the exact failure this tool exists to stop,
    // and it would look identical to a clean run.
    unreadable.push({ path, why: err.code ?? String(err) });
    continue;
  }
  if (text.includes("\0")) continue;

  const fromDir = dirOf(path);
  const lines = text.split(/\r?\n/);

  for (let i = 0; i < lines.length; i++) {
    for (const written of lines[i].match(CITATION) ?? []) {
      const found = candidatesFor(written);
      const where = { path, line: i + 1, written, text: lines[i].trim() };

      if (found.length === 0) {
        dangling.push(where);
        continue;
      }
      if (found.length === 1) continue;

      // More than one probe answers to this citation. Let the nearest one win,
      // but only if it wins outright — a tie means the reader cannot tell
      // either, which is the whole harm being guarded against.
      let best = -1;
      let winners = [];
      for (const cand of found) {
        const depth = sharedDepth(fromDir, dirOf(cand));
        if (depth > best) {
          best = depth;
          winners = [cand];
        } else if (depth === best) {
          winners.push(cand);
        }
      }
      if (winners.length !== 1) ambiguous.push({ ...where, found });
    }
  }
}

if (unreadable.length > 0) {
  console.error(`WARN — ${unreadable.length} tracked file(s) could not be read and were not checked:`);
  for (const u of unreadable) console.error(`  ${u.path}  (${u.why})`);
  console.error("");
}

if (dangling.length > 0 || ambiguous.length > 0) {
  if (dangling.length > 0) {
    console.error(`FAIL — ${dangling.length} citation(s) name a probe not tracked in this repo:\n`);
    for (const d of dangling) {
      console.error(`  ${d.path}:${d.line}  ->  ${d.written}`);
      console.error(`    ${d.text}\n`);
    }
    console.error(
      "If the probe exists in a sibling worktree, it is on another branch and is\n" +
        "not evidence anyone reading this branch can follow. Either bring it into\n" +
        "this branch or cite a probe that ships with the claim.\n",
    );
  }
  if (ambiguous.length > 0) {
    console.error(`FAIL — ${ambiguous.length} citation(s) match more than one probe:\n`);
    for (const a of ambiguous) {
      console.error(`  ${a.path}:${a.line}  ->  ${a.written}`);
      console.error(`    ${a.text}`);
      console.error(`    could be: ${a.found.join(", ")}\n`);
    }
    console.error(
      "A citation that names two probes names neither. Write enough of the path\n" +
        "to pick one, e.g. spikes/powerpoint/probes/probe-export.ps1.\n",
    );
  }
  process.exit(1);
}

console.log(`OK — every probe citation resolves (${probePaths.length} probe(s) tracked).`);
