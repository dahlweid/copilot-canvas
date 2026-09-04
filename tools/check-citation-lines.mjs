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
// that a coordinate inside a fence still resolves.
//
// That hole is therefore *measured* rather than merely disclaimed. The matcher
// is run over the skipped lines too, and a green run reports how many of them
// are coordinate-shaped and which files they are in — without failing on them.
// The number is a size, not a defect count: on this tree it is 2, both in the
// FINDINGS.md transcript above, and driving it to zero would mean editing that
// transcript. Read it as "what this gate cannot see", and never as a to-do
// list. Its absence is what let an earlier wording here claim a clean tree
// while those two sat unread — a review had to go and find them by hand.
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
// Three allowances from ADR 0009, plus one limit of scope. Only the first is
// something an author may reach for, and the failure message says so:
//
//   1. **A commit pin.** A coordinate followed *immediately* by "as of `<sha>`"
//      cannot rot — the SHA fixes the tree the number indexes. The SHA is
//      **resolved against git**, not merely matched: `as of `deadbee`` has the
//      shape of a pin and names no commit, and a pin resolving to nothing fixes
//      nothing. A checkout without the history cannot answer, and that state is
//      reported rather than accepted. Pinned coordinates are counted on a green
//      run, so the allowance is visible rather than silent.
//   2. **A verbatim transcript**, inside a fenced block, per the paragraph
//      above. Reported on a green run as a line count *and* as the number of
//      those lines that are coordinate-shaped, never as a clean bill.
//   3. **The subject under discussion** — the exempt list, for files that carry
//      rot-shaped coordinates as their *subject matter*: this checker, its unit
//      test, and ADR 0009, which analyses rotted coordinates and must quote
//      them. Naming a string is not pointing at a line. It is deliberately not
//      offered as a remedy in the failure message: an exemption an author can
//      grant themselves at an inconvenient moment is how a gate gets worked
//      around rather than obeyed. Growing it is an ADR-0009 question.
//   4. **Vendored and binary files**, skipped as out of scope rather than
//      excused — see VENDORED and SKIP below. Listed here because the gate does
//      it: documenting three allowances while applying four is the same
//      overstatement this header was corrected for once already.

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

// A filename with a text-file extension, then :NN or :NN-NN. Directory segments in
// front are kept (both separators) so the whole written form is reported back.
//
// The set is wide on purpose. It was once `mjs|ps1|js|ts|md|yml|yaml`, which is
// the set of things this tree *mostly* references — but a coordinate into
// `data.json` rots exactly as one into `data.mjs` does, and no principle
// separates them. Measured on the tracked list, that narrower set was blind to
// `json`, `css` and `html`, all of which are present here. What stays out is
// anything that is not a text file, since a line number into a PNG is not a
// citation, and any extension is not accepted outright because "word.number"
// matches version strings and prose.
const QUALIFIED =
  /(?:[A-Za-z0-9._-]+[/\\])*[A-Za-z0-9._-]+\.(?:mjs|cjs|mts|cts|js|jsx|ts|tsx|ps1|psm1|psd1|md|markdown|yml|yaml|json|jsonc|txt|xml|html|htm|css|scss|sh|bash|bat|cmd|toml|ini|cfg|conf|csv|sql|py|rs|go|c|h|cpp|hpp|cs|java|rb):\d+(?:-\d+)?/g;

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
//
// The SHA is captured, not merely matched, because the shape of a SHA is not
// evidence that a commit exists: `as of `deadbee`` satisfies this pattern and
// names nothing. A pin that resolves to no commit protects nothing, so it is
// resolved against git and rejected when it does not — see verifyPins.
const PIN = /^[\s`,;)\]]*as of\s+`?([0-9a-f]{7,40})`?/i;

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
/**
 * Every coordinate-shaped hit on one line, qualified first then bare.
 *
 * Split out because the gate needs to run it over lines it will *not* act on —
 * the ones inside fenced blocks and inside exempt files — so that a skip can be
 * reported as a measured hole rather than a silent one.
 */
function coordinatesIn(line) {
  // Blank out the qualified matches before hunting bare ones, so the `:12`
  // inside `a.mjs:12` is not also reported as a bare self-reference. The
  // blanks are the same length, so every offset stays the line's own.
  let rest = line;
  const hits = [];
  for (const m of line.matchAll(QUALIFIED)) {
    hits.push({ written: m[0], at: m.index, kind: "qualified" });
    rest = rest.slice(0, m.index) + " ".repeat(m[0].length) + rest.slice(m.index + m[0].length);
  }
  for (const m of rest.matchAll(BARE)) {
    // The introducer is part of the match; report the coordinate itself.
    const off = m[0].indexOf(":");
    hits.push({ written: m[0].slice(off), at: m.index + off, kind: "bare" });
  }
  return hits;
}

export function findCoordinates(tracked, readFile) {
  const coordinates = [];
  const pinned = [];
  const unreadable = [];
  const skipped = {
    exemptFiles: 0,
    vendoredFiles: 0,
    binaryFiles: 0,
    fencedLines: 0,
    fencedFiles: 0,
    // The measured size of the two skips that hide *authored* text. Not a defect
    // count — see report() — but the number whose absence let a false success
    // message stand.
    fencedCoordinateLines: 0,
    fencedCoordinateFiles: [],
    exemptCoordinateLines: 0,
  };

  for (const path of tracked) {
    const norm = path.replace(/\\/g, "/");
    if (EXEMPT.has(norm)) {
      skipped.exemptFiles++;
      // Counted, not acted on. An exemption nobody can size is an exemption
      // nobody rechecks.
      try {
        for (const line of readFile(path).split(/\r?\n/)) {
          if (coordinatesIn(line).length > 0) skipped.exemptCoordinateLines++;
        }
      } catch {
        // An exempt file that cannot be read costs only this census figure; it
        // was never going to be judged. unreadable is for files the gate had a
        // verdict to give.
      }
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
    let fencedCoordsHere = 0;

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
        if (coordinatesIn(lines[i]).length > 0) fencedCoordsHere++;
        continue;
      }

      const line = lines[i];

      for (const { written, at, kind } of coordinatesIn(line)) {
        const where = { path: norm, line: i + 1, written, kind, text: line.trim() };
        // The pin must follow this coordinate immediately, on the same line —
        // the form the ADR states: "`structure-map.mjs:266` as of `4abf952`".
        // Offsets come from the match, so a line carrying several coordinates
        // governs each one separately rather than letting one pin cover them all.
        const pin = PIN.exec(line.slice(at + written.length));
        if (pin) pinned.push({ ...where, pin: pin[1] });
        else coordinates.push(where);
      }
    }
    skipped.fencedLines += fencedHere;
    if (fencedHere > 0) skipped.fencedFiles++;
    if (fencedCoordsHere > 0) {
      skipped.fencedCoordinateLines += fencedCoordsHere;
      skipped.fencedCoordinateFiles.push(norm);
    }
  }

  return { coordinates, pinned, unreadable, skipped };
}

/**
 * Sorts pinned coordinates into resolved, fabricated and unverifiable.
 *
 * The pin is the one exception an author may apply to themselves, and its whole
 * force is that a SHA fixes the tree the number indexes. A SHA that names no
 * commit fixes nothing, so `as of \`deadbee\`` would otherwise be a way of
 * writing "trust me" that passes a gate. Verified because the shape of a SHA is
 * cheap to forge and cheap to mistype.
 *
 * `resolveCommit(sha)` returns true when the object exists, false when git says
 * it does not, and null when it could not be asked — which is not a pass. A
 * shallow clone holds only the head commit, so every historical pin is
 * unresolvable there; that state is REPORTED rather than treated as either
 * verdict, because a check that passes because it did nothing is the failure
 * this whole document is about. `validate.yml` therefore checks out with
 * `fetch-depth: 0`, which costs a 3.21 MiB pack over 114 commits here — measured
 * with `git count-objects -vH`, not assumed — so in this repo's own CI the
 * unverifiable branch should never be taken.
 */
export function verifyPins(pinned, resolveCommit) {
  const fabricated = [];
  let resolved = 0;
  let unverifiable = 0;
  for (const p of pinned) {
    const answer = p.pin ? resolveCommit(p.pin) : false;
    if (answer === null) unverifiable++;
    else if (answer) resolved++;
    else fabricated.push(p);
  }
  return { resolved, fabricated, unverifiable };
}

/** Prints the report and returns the intended exit code (0 pass, 1 fail). */
export function report({ coordinates, pinned, unreadable, skipped, pins }) {
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
        "ADR 0009 states three exceptions to the rule, and this gate allows a\n" +
        "fourth thing that is a limit of scope rather than a licence. Exactly one\n" +
        "of them is yours to apply:\n\n" +
        "  * A COMMIT PIN. If the number is genuinely needed as evidence, pin it:\n" +
        "    `structure-map.mjs:266` as of `4abf952`. The SHA fixes the tree the\n" +
        "    number indexes, so it cannot rot, and this gate accepts it — after\n" +
        "    resolving the SHA, so a fabricated hash is rejected, not waved past.\n\n" +
        "The rest are not remedies, and none is a way round this message:\n\n" +
        "  * A VERBATIM TRANSCRIPT, inside a fenced block. A recorded probe run is\n" +
        "    evidence; evidence is not edited to satisfy a linter. This gate skips\n" +
        "    fenced blocks wholesale, so it will not have sent you here for one —\n" +
        "    and it does not check that a coordinate inside one still resolves.\n" +
        "  * THE SUBJECT UNDER DISCUSSION — a coordinate quoted because the rotted\n" +
        "    coordinate *is* what the text is about. In the tree that is the exempt\n" +
        "    list: this gate, its unit test, ADR 0009. If you believe a file belongs\n" +
        "    on it, argue that in ADR 0009 rather than editing the list — an\n" +
        "    exemption granted in passing is how a gate stops meaning anything.\n" +
        "  * VENDORED AND BINARY FILES, skipped as out of scope rather than\n" +
        "    excused: `src/ui/vendor/` is third-party code not authored here, and a\n" +
        "    coordinate-shaped string inside minified output is not a reference into\n" +
        "    this tree.\n\n" +
        "All of them, and why the syntax is banned at all, are in\n" +
        "docs/adr/0009-issue-and-pr-text-quotes-the-claim.md.\n",
    );
    return 1;
  }

  const p = pins ?? {};
  const fabricated = p.fabricated ?? [];
  if (fabricated.length > 0) {
    console.error(
      `FAIL — ${fabricated.length} coordinate(s) claim a commit pin that names no commit:\n`,
    );
    for (const c of fabricated) {
      console.error(`  ${c.path}:${c.line}  ->  ${c.written}  as of \`${c.pin}\``);
      console.error(`    ${c.text}\n`);
    }
    console.error(
      "A pin's whole force is that the SHA fixes the tree the number indexes. A\n" +
        "SHA that resolves to nothing fixes nothing, so this is an unpinned\n" +
        "coordinate wearing the one exception an author may apply to themselves.\n" +
        "Either pin it to a commit that exists, or — better — reference the code by\n" +
        "name and drop the number.\n\n" +
        "If the SHA is real but absent here, this checkout lacks the history: see\n" +
        "`fetch-depth` in .github/workflows/validate.yml.\n",
    );
    return 1;
  }

  // Reported, never treated as a pass. A shallow clone can answer no pin, and a
  // check that stayed quiet about that would be passing because it did nothing.
  //
  // State the coverage, not a claim wider than it. The earlier wording here —
  // "no positional coordinate in any tracked file" — was false while two sat
  // unread inside a fenced transcript in spikes/word-icon/FINDINGS.md. A gate
  // that overstates what it verified is worse than a narrower honest one,
  // because the overstatement is what stops the next person looking.
  const s = skipped ?? {};
  const pinNote =
    pinned.length === 0
      ? "no commit-pinned coordinate was allowed through"
      : p.unverifiable > 0
        ? `${pinned.length} commit-pinned coordinate(s) were allowed through, of which ` +
          `${p.unverifiable} could NOT be resolved against git — this checkout lacks the ` +
          `history, so those pins are unverified rather than accepted`
        : `${pinned.length} commit-pinned coordinate(s) were allowed through, each resolved ` +
          `to a real commit, so they cannot rot`;
  // Sized, not just named. The hole this reports is the one that let the old
  // wording stand: a coordinate-shaped line the gate skipped was indistinguishable
  // from no coordinate at all. It is deliberately NOT called a defect count —
  // see the label below.
  const fencedCoordFiles = s.fencedCoordinateFiles ?? [];
  const fencedCoords =
    (s.fencedCoordinateLines ?? 0) === 0
      ? "    None of them is coordinate-shaped.\n"
      : `    ${s.fencedCoordinateLines} of those lines ${s.fencedCoordinateLines === 1 ? "is" : "are"} coordinate-shaped, in ` +
        `${fencedCoordFiles.join(", ")}.\n` +
        "    That is the size of what this gate cannot see, NOT a count of work\n" +
        "    left undone: a coordinate inside a recorded run is protected on\n" +
        "    purpose, and driving this number to zero would mean editing\n" +
        "    transcripts, which is the harm the exception exists to prevent.\n";
  console.log(
    `OK — no positional coordinate in any tracked line this gate read; ${pinNote}.\n\n` +
      "Not read, and therefore not certified:\n" +
      `  ${s.fencedLines ?? 0} line(s) inside fenced blocks, across ${s.fencedFiles ?? 0} file(s) — verbatim\n` +
      "    transcripts, which are evidence and are not edited to satisfy a gate.\n" +
      fencedCoords +
      `  ${s.exemptFiles ?? 0} exempt file(s), whose subject matter is a rotted coordinate;\n` +
      `    ${s.exemptCoordinateLines ?? 0} coordinate-shaped line(s) in them, which is what being a\n` +
      "    fixture for this rule looks like.\n" +
      `  ${s.vendoredFiles ?? 0} vendored and ${s.binaryFiles ?? 0} binary file(s), not scanned at all.\n\n` +
      "Two gaps are in the matcher itself, and no count can size them: a bare\n" +
      "`:NN` with no `(`, backtick or \"at\"/\"line\" in front is not recognised as a\n" +
      "citation, because unintroduced `:\\d+` is a clock, a port or a JSON value far\n" +
      "more often than a reference, and a gate that flagged those would be turned\n" +
      "off rather than obeyed. A coordinate into a file whose extension is not in\n" +
      "the whitelist is invisible for the same reason. Both are trades, not\n" +
      "oversights, and both are stated here so nobody reads this pass as wider.\n\n" +
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

  // A shallow clone holds only the head commit, so it can answer no historical
  // pin. Asked once, up front, so the "unverifiable" report names the cause
  // rather than inferring one from a failed lookup — the same discipline the
  // rest of this tree applies to error codes.
  let shallow = false;
  try {
    shallow = git("-C", root, "rev-parse", "--is-shallow-repository").trim() === "true";
  } catch {
    shallow = true;
  }
  const resolveCommit = (sha) => {
    if (shallow) return null;
    try {
      git("-C", root, "rev-parse", "--verify", "--quiet", `${sha}^{commit}`);
      return true;
    } catch {
      // rev-parse --verify exits non-zero for an object that is not here. With a
      // full clone that means the SHA names nothing, which is a finding; the
      // shallow case was already separated out above, so this branch cannot be
      // "we could not look".
      return false;
    }
  };

  process.exit(report({ ...result, pins: verifyPins(result.pinned, resolveCommit) }));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main();
}
