#!/usr/bin/env node
// Positional citations (`file.ext:NN` and `file.ext:NN-NN`) in tracked text,
// checked for the failures a machine can decide **without knowing what the
// citation meant** — and only those.
//
// ## Why this is a sibling of check-citations.mjs, not an extension of it
//
// That checker answers one question: does a `probe-*.ps1` / `probe-*.mjs`
// filename resolve to exactly one probe tracked on *this* branch? It is a
// dangling-reference guard, scoped on purpose to the repo's evidence trail, and
// it has no notion of a line number. This tool answers a different question
// about a broader syntax — does a `:NN` coordinate name a file that exists here,
// unambiguously, with that line in range? — and carries a materially weaker
// guarantee (see COVERAGE below). Folding the weak guarantee into the strong,
// well-scoped one would blur what each promises. They stay separate.
//
// ## COVERAGE — read this before trusting a green run
//
// This is the honest, load-bearing part. Issue #160 audited 8 rotted/unresolvable
// positional citations. This check reliably catches exactly the ones whose
// failure is **decidable from the filesystem alone**:
//
//   * a cited file that does not exist as tracked, INCLUDING a case-only
//     mismatch like `Word-host.mjs` for `word-host.mjs` (which also breaks a
//     case-sensitive checkout outright); and
//   * a bare basename that matches more than one tracked file (ambiguous), with
//     the same proximity tie-break the sibling uses; and
//   * a line (or the high end of a range) past the end of the file.
//
// Of #160's eight, that is **2**: `Word-host.mjs:95` (case) and
// `make-fixture.ps1:92` (ambiguous). Both were also line-wrong, but the case and
// ambiguity failures are what this catches them on.
//
// It does **not** catch the other six, and cannot. Every one of them lands on a
// real line of a real file — `word-pids.mjs:59-61` on an unrelated
// `execFileAsync` call, `structure-map.mjs:266` on plausible code, a `#` or a
// bare `}`. A resolver sees a syntactically valid line at a valid offset and has
// nothing to compare it against, because the thing needed to check it — what the
// citation was asserting — was never written down. **A coordinate that still
// resolves to real code while no longer meaning anything is invisible to any
// gate.** Do not read a green run as "every citation is correct"; read it as
// "no citation names a missing file, an ambiguous file, or an out-of-range line".
//
// A line-existence check that called itself a citation checker would claim 8 and
// catch 2. This one catches 2 and says 2.
//
// ## Scope
//
// Filenames ending .mjs/.ps1/.js/.ts/.md immediately followed by :NN or :NN-NN.
// Fenced code blocks are skipped, because ```...``` in Markdown routinely holds
// captured tool output (e.g. a grep result printing `file:line`) that is a
// recording, not an authored citation into this tree. This checker's own two
// files are exempt — its source (which quotes example coordinates above) and its
// unit test (which must carry rot-shaped citations as fixtures to assert on) —
// and nothing else is.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

// A filename with a code/doc extension, then :NN or :NN-NN. Directory segments in
// front are kept (both separators) so a full path can be told from a basename.
const CITATION =
  /(?:[A-Za-z0-9._-]+[/\\])*[A-Za-z0-9._-]+\.(?:mjs|ps1|js|ts|md):\d+(?:-\d+)?/g;

const SKIP = /\.(?:png|jpg|jpeg|gif|pdf|docx|pptx|xlsx|zip|ico)$/i;

// The checker's own two files are exempt. Its source quotes example coordinates
// in the header; its unit test must carry rot-shaped citations as fixtures and
// assert on them by literal value — a checker whose test could not name the rot
// it catches could not test itself. This is that one file-pair, not a general
// opt-out: no third path is exempt, and a citation-shaped string anywhere else,
// including in another test, is still checked.
const EXEMPT = new Set([
  "tools/check-citation-lines.mjs",
  ".github/extensions/office-canvas/test/unit/check-citation-lines.test.mjs",
]);

function dirOf(p) {
  const i = p.lastIndexOf("/");
  return i === -1 ? "" : p.slice(0, i);
}

function sharedDepth(a, b) {
  const x = a === "" ? [] : a.split("/");
  const y = b === "" ? [] : b.split("/");
  let n = 0;
  while (n < x.length && n < y.length && x[n] === y[n]) n++;
  return n;
}

// Tracked files whose path ends with the written path on a directory boundary.
// Case-sensitive by design: a case-only match is a real failure (`Word-host.mjs`
// is not `word-host.mjs`), and it breaks a case-sensitive checkout, so it must
// surface as a miss rather than be silently rescued.
function candidatesFor(written, trackedNorm) {
  const want = written.replace(/\\/g, "/").replace(/^\.\//, "");
  return trackedNorm.filter((p) => p === want || p.endsWith("/" + want));
}

// One winner by proximity, or null if it does not win outright. Mirrors the
// sibling: the nearest candidate to the citing file wins, a tie stays ambiguous.
function resolveOne(candidates, fromDir) {
  if (candidates.length === 1) return candidates[0];
  let best = -1;
  let winners = [];
  for (const cand of candidates) {
    const depth = sharedDepth(fromDir, dirOf(cand));
    if (depth > best) {
      best = depth;
      winners = [cand];
    } else if (depth === best) {
      winners.push(cand);
    }
  }
  return winners.length === 1 ? winners[0] : null;
}

// Number of lines in a file's text. A trailing newline does NOT add a line:
// `"a\nb\n".split(/\r?\n/)` is `["a","b",""]`, whose length (3) over-counts the
// two-line file by one — and that one is exactly the EOF+1 offset an out-of-range
// check exists to catch, so the naive length fails the check open at its own
// boundary. Dropping a single trailing empty segment corrects it. An empty file
// is deliberately 0 lines: `"".split(/\r?\n/)` is `[""]`, and a citation to line 1
// of an empty file has no line 1, so it must read as out of range, not in range.
function countLines(text) {
  const parts = text.split(/\r?\n/);
  if (parts.length > 0 && parts[parts.length - 1] === "") return parts.length - 1;
  return parts.length;
}

// Split a citation "path:NN" or "path:NN-NN" into { file, lo, hi }. The regex
// guarantees a trailing :digits(-digits)?, so the match is total.
function parseCitation(written) {
  const m = written.match(/^(.*):(\d+)(?:-(\d+))?$/);
  const file = m[1];
  const lo = Number(m[2]);
  const hi = m[3] === undefined ? lo : Number(m[3]);
  return { file, lo, hi };
}

/**
 * The whole decision, as a pure function of a file list and a reader, so it can
 * be driven by a unit test without a git checkout. `tracked` is the list of
 * repo-relative paths; `readFile(path)` returns the file's text or throws an
 * error carrying `.code` (as `fs.readFileSync` does). Nothing here touches git,
 * the filesystem, the process exit code, or stdout — the CLI below owns those.
 *
 * Returns { missing, ambiguous, outOfRange, unreadable }, each an array of the
 * offending sites. An empty missing/ambiguous/outOfRange is a pass.
 */
export function analyzeCitations(tracked, readFile) {
  const trackedNorm = tracked.map((p) => p.replace(/\\/g, "/"));
  const missing = []; // file does not resolve (nonexistent, or case-only mismatch)
  const ambiguous = []; // basename matches >1 tracked file, no proximity winner
  const outOfRange = []; // line, or high end of range, past end of file
  const unreadable = [];

  for (const path of tracked) {
    const norm = path.replace(/\\/g, "/");
    if (EXEMPT.has(norm) || SKIP.test(norm)) continue;

    let text;
    try {
      text = readFile(path);
    } catch (err) {
      // Reported, never swallowed — a guard that quietly declines to look at part
      // of the repo is indistinguishable from a clean run, which is the failure
      // this class of tool exists to stop.
      unreadable.push({ path: norm, why: err.code ?? String(err) });
      continue;
    }
    if (text.includes("\0")) continue;

    const lines = text.split(/\r?\n/);
    const fromDir = dirOf(norm);
    let inFence = false;

    for (let i = 0; i < lines.length; i++) {
      // Skip fenced code blocks: they carry captured output, not authored
      // citations. A fence marker is ``` (or ~~~) at the start of the line,
      // possibly indented, possibly with an info string.
      if (/^\s*(?:```|~~~)/.test(lines[i])) {
        inFence = !inFence;
        continue;
      }
      if (inFence) continue;

      for (const written of lines[i].match(CITATION) ?? []) {
        const { file, hi } = parseCitation(written);
        const where = { path: norm, line: i + 1, written, text: lines[i].trim() };

        const candidates = candidatesFor(file, trackedNorm);
        if (candidates.length === 0) {
          missing.push(where);
          continue;
        }
        const resolved = resolveOne(candidates, fromDir);
        if (resolved === null) {
          ambiguous.push({ ...where, found: candidates });
          continue;
        }

        let target;
        try {
          target = readFile(resolved);
        } catch (err) {
          unreadable.push({ path: resolved, why: err.code ?? String(err) });
          continue;
        }
        const lineCount = countLines(target);
        if (hi > lineCount) {
          outOfRange.push({ ...where, resolved, lineCount });
        }
      }
    }
  }

  return { missing, ambiguous, outOfRange, unreadable };
}

/** Prints the report and returns the intended exit code (0 pass, 1 fail). */
export function report({ missing, ambiguous, outOfRange, unreadable }) {
  if (unreadable.length > 0) {
    console.error(`WARN — ${unreadable.length} file(s) could not be read and were not checked:`);
    for (const u of unreadable) console.error(`  ${u.path}  (${u.why})`);
    console.error("");
  }

  let failed = false;

  if (missing.length > 0) {
    failed = true;
    console.error(`FAIL — ${missing.length} positional citation(s) name a file not tracked here (nonexistent, or case-only mismatch):\n`);
    for (const m of missing) {
      console.error(`  ${m.path}:${m.line}  ->  ${m.written}`);
      console.error(`    ${m.text}\n`);
    }
    console.error(
      "A case-only mismatch (Word-host.mjs vs word-host.mjs) also breaks a\n" +
        "case-sensitive checkout. Fix the path to the tracked file.\n",
    );
  }

  if (ambiguous.length > 0) {
    failed = true;
    console.error(`FAIL — ${ambiguous.length} positional citation(s) match more than one tracked file:\n`);
    for (const a of ambiguous) {
      console.error(`  ${a.path}:${a.line}  ->  ${a.written}`);
      console.error(`    ${a.text}`);
      console.error(`    could be: ${a.found.join(", ")}\n`);
    }
    console.error("Write enough of the path to pick one file.\n");
  }

  if (outOfRange.length > 0) {
    failed = true;
    console.error(`FAIL — ${outOfRange.length} positional citation(s) point past the end of the cited file:\n`);
    for (const o of outOfRange) {
      console.error(`  ${o.path}:${o.line}  ->  ${o.written}  (${o.resolved} has ${o.lineCount} line(s))`);
      console.error(`    ${o.text}\n`);
    }
  }

  if (failed) return 1;

  console.log(
    "OK — every positional citation names a tracked file, unambiguously, in range.\n" +
      "(This does not verify a coordinate still means what the prose claims — see\n" +
      "the COVERAGE note in this file. A citation landing on real but unrelated code\n" +
      "passes here and cannot be caught mechanically.)",
  );
  return 0;
}

// CLI: ask git for the tracked file list, read from disk, print, exit. Runs only
// when executed directly, never on import — the unit test imports the pure core
// above and never spawns git.
function main() {
  const git = (...args) =>
    execFileSync("git", args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });

  // Anchored to the repo root for the same reason the sibling is: ls-files scopes
  // and re-bases its output on the directory it runs in, so a subdirectory run
  // silently changes both the files searched and the files considered to exist.
  const root = git("rev-parse", "--show-toplevel").trim();
  const tracked = git("-C", root, "ls-files", "-z").split("\0").filter(Boolean);

  const result = analyzeCitations(tracked, (p) => readFileSync(join(root, p), "utf8"));
  process.exit(report(result));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main();
}
