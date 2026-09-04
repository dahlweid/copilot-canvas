#!/usr/bin/env node
// Positional coordinates (`file.ext:NN`, `file.ext:NN-NN`, and a bare `(:NN)` or
// `` `:NN` `` into the file's own body) are **rejected** in tracked text.
//
// ## What changed, and why the old job was the wrong one
//
// This tool used to *validate* coordinates: it resolved the named file, checked
// the line was in range, and passed anything that resolved. ADR 0009 retired
// that job. Validation makes a forbidden construct comfortable, and it cannot
// buy what it appears to: a coordinate that resolved correctly when written is
// relocated by the next insertion **above** it, in a file its author never
// touched. That is the #141 mechanism precisely — the coordinate was correct at
// the instant it was written and rotted seventeen minutes later — and no
// resolver can see it, because the thing needed to check a coordinate (what it
// was asserting) was never written down.
//
// So the guarantee is inverted. This gate no longer asks whether a coordinate
// resolves; it asks whether one is present, which is decidable and stays
// decidable. A ban on a *syntax* is enforceable where a check on whether a
// coordinate still *means* what it claimed is not.
//
// ## What a green run means now
//
// No tracked *line this gate read* carries a positional coordinate, except where
// one is pinned to a commit. It reads every tracked text file outside the exempt
// list, minus fenced blocks and vendored code. That is a real property of the
// tree, unlike the old message, which promised only that a coordinate named a
// file that existed at a line that existed.
//
// The qualifier is load-bearing and was learned the hard way: this said "no
// tracked file carries a positional coordinate" while two sat unread inside a
// fenced transcript. A gate that overstates its coverage is worse than a
// narrower honest one, because the overstatement is what stops the next person
// looking. The success message now prints what it skipped, with counts.
//
// It is still not a claim that every *reference* is correct. ADR 0009 says why:
// dropping the number does not make a reference right, because a renamed
// function or a moved file breaks a name too. The narrow claim is the measured
// one — the rot in this tree is overwhelmingly relocation by an edit elsewhere,
// and a name is immune to exactly that.
//
// ## Sibling: check-citations.mjs
//
// That checker answers a different question — does a `probe-*.ps1` / `probe-*.mjs`
// filename resolve to exactly one probe tracked on *this* branch? It is a
// dangling-reference guard over the repo's evidence trail, it has no notion of a
// line, and a bare filename with no coordinate is exactly what ADR 0009 wants
// written. The two gates are complementary and stay separate.
//
// ## Scope and the two allowances
//
// Matched: a filename ending .mjs/.ps1/.js/.ts/.md/.yml/.yaml immediately
// followed by :NN or :NN-NN, and a bare :NN / :NN-NN introduced by `(`, a
// backtick, or the words "at" / "line" / "lines" — the self-reference form, which ADR 0009 records as
// simultaneously the most fragile (it rots on any insertion above it in its own
// file, the most frequent edit a file receives) and the least checkable (a
// matcher needs a filename to recognise a citation at all).
//
// Fenced code blocks are skipped, and this is an **exception to ADR 0009, named
// there** — a verbatim record of a probe run is evidence, and evidence is not
// edited to satisfy a linter. This repo's standard is that a claim about
// platform behaviour is backed by a probe that was actually run; rewriting a
// coordinate inside that probe's printed output destroys exactly the thing the
// standard rests on. The worked case is the transcript in
// `spikes/word-icon/FINDINGS.md`, which prints `"icon" appears in
// [generated\rpc.d.ts:18, generated\session-events.d.ts:13]` — two coordinates
// into **Copilot CLI 1.0.80's generated typings**, an external artifact, not
// into this tree, which has no `generated/` at all.
//
// Note the limit, because the skip is wider than that justification: it drops
// every fenced line, including one holding an unpinned coordinate into this
// repo. So the skip protects the recording; it is emphatically **not** a claim
// that a coordinate inside a fence still resolves. The success message names
// this, and says how many lines went unread, rather than reporting a clean tree.
//
// Vendored third-party code under `src/ui/vendor/` is skipped for a related
// reason: it is not authored here, and nothing in it is a reference into this
// tree.
//
// The size of the fence skip, measured at `9c81f9b` rather than assumed: turning
// fence tracking off and changing nothing else takes the count from 43 to 45,
// and both extra hits are that one FINDINGS.md pair. Worth stating, because a
// scanner that lacked fence tracking *and* matched far more widely read 88 here,
// and it would be easy to credit the fence with a gap that was almost all
// matcher width.
//
// Three allowances, all from ADR 0009:
//
//   1. **A commit pin.** A coordinate followed *immediately* by "as of `<sha>`"
//      cannot rot — the SHA fixes the tree the number indexes. Pinned
//      coordinates are counted and reported on a green run, so the allowance is
//      visible rather than silent. This is the allowance an author who hits the
//      gate should reach for, and the failure message says so.
//   2. **A verbatim transcript**, inside a fenced block, per the paragraph
//      above. Reported as a line count on a green run, never as a clean bill.
//   3. **An exempt list**, for the files that carry rot-shaped *fixtures*: this
//      checker, its unit test, and ADR 0009 itself, which analyses rotted
//      coordinates and must quote them. Naming a string is not pointing at a
//      line. This is that short list and nothing else. It is deliberately not
//      offered as a remedy in the failure message: an exemption an author can
//      grant themselves at an inconvenient moment is how a gate gets worked
//      around rather than obeyed. Growing it is an ADR-0009 question.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

// A filename with a code/doc extension, then :NN or :NN-NN. Directory segments in
// front are kept (both separators) so the whole written form is reported back.
const QUALIFIED =
  /(?:[A-Za-z0-9._-]+[/\\])*[A-Za-z0-9._-]+\.(?:mjs|ps1|js|ts|md|yml|yaml):\d+(?:-\d+)?/g;

// A bare coordinate into the file's own body. The introducer is required: a
// colon followed by digits is far too common on its own to ban outright — a JSON
// example (`{"id":1}`), a clock time, a ratio — and a matcher that flagged those
// would be turned off rather than obeyed. `(`, a backtick, and the words
// "at"/"line"/"lines" are the forms this tree actually used, measured across
// every tracked file.
const BARE = /(?:[(`]|\b(?:at|lines?)\s+):\d+(?:-\d+)?(?![\w.-])/g;

// "as of `4abf952`" **immediately** after the coordinate — only whitespace,
// backticks or closing punctuation may sit between. Anchored deliberately: a
// rule that accepted a pin anywhere later on the line would let one pinned
// coordinate cover every unpinned one written before it on that line, which is
// how a pinned-and-unpinned pair on one line silently passes. Seven hex digits
// is git's own abbreviation floor; a full 40 is accepted too.
const PIN = /^[\s`,;)\]]*as of\s+`?[0-9a-f]{7,40}`?/i;

const SKIP = /\.(?:png|jpg|jpeg|gif|pdf|docx|pptx|xlsx|zip|ico)$/i;

// Third-party code, vendored verbatim. Not authored here, and a coordinate-shaped
// string inside minified output is not a reference into this tree.
const VENDORED = /(?:^|\/)src\/ui\/vendor\//;

// The files that carry rot-shaped coordinates as their subject matter rather than
// as references being followed. ADR 0009's own second exception, and no wider:
// a coordinate-shaped string anywhere else, including in another test, is still
// rejected.
const EXEMPT = new Set([
  "tools/check-citation-lines.mjs",
  ".github/extensions/office-canvas/test/unit/check-citation-lines.test.mjs",
  "docs/adr/0009-issue-and-pr-text-quotes-the-claim.md",
]);

/**
 * The whole decision, as a pure function of a file list and a reader, so it can
 * be driven by a unit test without a git checkout. `tracked` is the list of
 * repo-relative paths; `readFile(path)` returns the file's text or throws an
 * error carrying `.code` (as `fs.readFileSync` does). Nothing here touches git,
 * the filesystem, the process exit code, or stdout — the CLI below owns those.
 *
 * Returns { coordinates, pinned, unreadable, skipped }. `coordinates` is the
 * rejection list and an empty one is a pass; `pinned` records what the commit-pin
 * allowance let through; `skipped` counts what was never looked at, so the
 * success message can state its own coverage instead of overstating it. That
 * last one is not bookkeeping: a green run previously claimed "no positional
 * coordinate in any tracked file" while two sat unread inside a fenced
 * transcript, which is precisely the overstatement that stops the next person
 * looking.
 */
export function findCoordinates(tracked, readFile) {
  const coordinates = [];
  const pinned = [];
  const unreadable = [];
  const skipped = { exemptFiles: 0, vendoredFiles: 0, binaryFiles: 0, fencedLines: 0, fencedFiles: 0 };

  for (const path of tracked) {
    const norm = path.replace(/\\/g, "/");
    if (EXEMPT.has(norm)) {
      skipped.exemptFiles++;
      continue;
    }
    if (VENDORED.test(norm)) {
      skipped.vendoredFiles++;
      continue;
    }
    if (SKIP.test(norm)) {
      skipped.binaryFiles++;
      continue;
    }

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
    if (text.includes("\0")) {
      skipped.binaryFiles++;
      continue;
    }

    const lines = text.split(/\r?\n/);
    let inFence = false;
    let fencedHere = 0;

    for (let i = 0; i < lines.length; i++) {
      // Skip fenced code blocks: they carry captured output, not authored
      // references. A fence marker is ``` (or ~~~) at the start of the line,
      // possibly indented, possibly with an info string.
      if (/^\s*(?:```|~~~)/.test(lines[i])) {
        inFence = !inFence;
        continue;
      }
      if (inFence) {
        fencedHere++;
        continue;
      }

      const line = lines[i];

      // Blank out the qualified matches before hunting bare ones, so the `:12`
      // inside `a.mjs:12` is not also reported as a bare self-reference. The
      // blanks are the same length, so every offset stays the line's own.
      let rest = line;
      const hits = [];
      for (const m of line.matchAll(QUALIFIED)) {
        hits.push({ written: m[0], at: m.index, kind: "qualified" });
        rest =
          rest.slice(0, m.index) + " ".repeat(m[0].length) + rest.slice(m.index + m[0].length);
      }
      for (const m of rest.matchAll(BARE)) {
        // The introducer is part of the match; report the coordinate itself.
        const off = m[0].indexOf(":");
        hits.push({ written: m[0].slice(off), at: m.index + off, kind: "bare" });
      }

      for (const { written, at, kind } of hits) {
        const where = { path: norm, line: i + 1, written, kind, text: line.trim() };
        // The pin must follow this coordinate immediately, on the same line —
        // the form the ADR states: "`structure-map.mjs:266` as of `4abf952`".
        // Offsets come from the match, so a line carrying several coordinates
        // governs each one separately rather than letting one pin cover them all.
        if (PIN.test(line.slice(at + written.length))) pinned.push(where);
        else coordinates.push(where);
      }
    }
    skipped.fencedLines += fencedHere;
    if (fencedHere > 0) skipped.fencedFiles++;
  }

  return { coordinates, pinned, unreadable, skipped };
}

/** Prints the report and returns the intended exit code (0 pass, 1 fail). */
export function report({ coordinates, pinned, unreadable, skipped }) {
  if (unreadable.length > 0) {
    console.error(`WARN — ${unreadable.length} file(s) could not be read and were not checked:`);
    for (const u of unreadable) console.error(`  ${u.path}  (${u.why})`);
    console.error("");
  }

  if (coordinates.length > 0) {
    console.error(`FAIL — ${coordinates.length} positional coordinate(s) in tracked files (ADR 0009):\n`);
    for (const c of coordinates) {
      console.error(`  ${c.path}:${c.line}  ->  ${c.written}  (${c.kind})`);
      console.error(`    ${c.text}\n`);
    }
    console.error(
      "A committed file references code by name, not by position: write\n" +
        "`Set-ParagraphText (word-host.ps1)`, or say what the thing is when it has\n" +
        "no name — `the census helper in _common.ps1`. A coordinate hides the claim\n" +
        "it stands for, and an insertion above it relocates that claim without\n" +
        "touching it.\n\n" +
        "ADR 0009 allows three exceptions. One is yours to apply:\n\n" +
        "  * A COMMIT PIN. If the number is genuinely needed as evidence, pin it:\n" +
        "    `structure-map.mjs:266` as of `4abf952`. The SHA fixes the tree the\n" +
        "    number indexes, so it cannot rot, and this gate accepts it.\n\n" +
        "The other two are not, and neither is a way round this message:\n\n" +
        "  * A VERBATIM TRANSCRIPT, inside a fenced block. A recorded probe run is\n" +
        "    evidence; evidence is not edited to satisfy a linter. This gate skips\n" +
        "    fenced blocks wholesale, so it will not have sent you here for one —\n" +
        "    and it does not check that a coordinate inside one still resolves.\n" +
        "  * THE EXEMPT LIST, for files whose subject matter IS a rotted coordinate\n" +
        "    (this gate, its unit test, ADR 0009). If you believe a file belongs on\n" +
        "    it, argue that in ADR 0009 rather than editing the list — an exemption\n" +
        "    granted in passing is how this gate stops meaning anything.\n\n" +
        "All three, and why the syntax is banned at all, are in\n" +
        "docs/adr/0009-issue-and-pr-text-quotes-the-claim.md.\n",
    );
    return 1;
  }

  // State the coverage, not a claim wider than it. The earlier wording here —
  // "no positional coordinate in any tracked file" — was false while two sat
  // unread inside a fenced transcript in spikes/word-icon/FINDINGS.md. A gate
  // that overstates what it verified is worse than a narrower honest one,
  // because the overstatement is what stops the next person looking.
  const s = skipped ?? {};
  const pinNote =
    pinned.length === 0
      ? "no commit-pinned coordinate was allowed through"
      : `${pinned.length} commit-pinned coordinate(s) were allowed through, which cannot rot`;
  console.log(
    `OK — no positional coordinate in any tracked line this gate read; ${pinNote}.\n\n` +
      "Not read, and therefore not certified:\n" +
      `  ${s.fencedLines ?? 0} line(s) inside fenced blocks, across ${s.fencedFiles ?? 0} file(s) — verbatim\n` +
      "    transcripts, which are evidence and are not edited to satisfy a gate.\n" +
      `  ${s.exemptFiles ?? 0} exempt file(s), whose subject matter is a rotted coordinate.\n` +
      `  ${s.vendoredFiles ?? 0} vendored and ${s.binaryFiles ?? 0} binary file(s), not authored here.\n\n` +
      "Nor does a pass certify that every reference is correct: a renamed function\n" +
      "or a moved file breaks a name too. It certifies that no reference this gate\n" +
      "read is a bare position, which is the rot this tree actually measured.",
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

  const result = findCoordinates(tracked, (p) => readFileSync(join(root, p), "utf8"));
  process.exit(report(result));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main();
}
