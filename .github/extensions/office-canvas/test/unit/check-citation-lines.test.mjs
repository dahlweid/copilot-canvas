// Unit tests for the positional-citation checker (`tools/check-citation-lines.mjs`).
//
// The checker's whole value is a coverage claim: it catches the citation rots
// that are decidable from the filesystem alone (missing/case-mismatched file,
// ambiguous basename, out-of-range line) and provably does NOT catch a
// coordinate that still lands on real but unrelated code. A test that could not
// distinguish those two is worthless here, so every assertion below is built to
// fail if the checker's behaviour drifts:
//
//   * the CATCHES tests feed a rot and assert it is reported. Delete the
//     corresponding branch in the checker and these go red.
//   * the NEGATIVE CONTROL feeds the exact shape of the six #160 rots the
//     checker cannot see — a citation onto real, unrelated code — and asserts it
//     passes. If someone "strengthens" the checker into pretending it catches
//     this class (e.g. by flagging any comment-only or brace-only target), this
//     test goes red and stops the overstatement. It is the assertion that keeps
//     the coverage claim honest.
//
// The checker is driven through its pure core `analyzeCitations(tracked,
// readFile)`, with an in-memory file map. No git, no disk, no shared helper —
// the fixtures are local to this file so nothing another worktree owns is
// touched. Office-free; runs on ubuntu-latest.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";

const UNIT = path.dirname(fileURLToPath(import.meta.url));
const CHECKER = path.resolve(UNIT, "..", "..", "..", "..", "..", "tools", "check-citation-lines.mjs");
const { analyzeCitations } = await import(`file://${CHECKER.split(path.sep).join("/")}`);

/**
 * Build a reader over an in-memory { path: text } map. A path not in the map
 * throws an ENOENT-shaped error, exactly as `fs.readFileSync` does, so the
 * checker's unreadable-vs-missing split is exercised for real.
 */
function readerOver(files) {
  return (p) => {
    const norm = p.replace(/\\/g, "/");
    if (!(norm in files)) {
      const err = new Error(`ENOENT: ${norm}`);
      err.code = "ENOENT";
      throw err;
    }
    return files[norm];
  };
}

function analyze(files) {
  return analyzeCitations(Object.keys(files), readerOver(files));
}

test("CATCHES: a citation whose file exists only under a different case is missing", () => {
  // The `Word-host.mjs:95` shape from #160 site 7. The tree has word-host.mjs;
  // the citation writes Word-host.mjs. Case-sensitive resolution must miss it.
  const files = {
    "src/word/word-host.mjs": "line1\nline2\nline3\n",
    "test/notes.md": "see `Word-host.mjs:2`, verbatim.\n",
  };
  const { missing, ambiguous, outOfRange } = analyze(files);
  assert.equal(ambiguous.length, 0);
  assert.equal(outOfRange.length, 0);
  assert.equal(missing.length, 1, "a case-only mismatch was not reported missing");
  assert.equal(missing[0].written, "Word-host.mjs:2");

  // Control that this is really about case, not about the file being absent: the
  // correctly-cased citation to the same line must NOT be reported.
  const ok = analyze({
    "src/word/word-host.mjs": "line1\nline2\nline3\n",
    "test/notes.md": "see `word-host.mjs:2`.\n",
  });
  assert.equal(ok.missing.length, 0, "the correctly-cased citation was wrongly reported missing");
});

test("CATCHES: a bare basename matching two tracked files is ambiguous", () => {
  // The `make-fixture.ps1:92` shape from #160 site 8. Two files share the
  // basename and neither is nearer the citing doc, so no proximity winner.
  const files = {
    "a/make-fixture.ps1": "x\ny\n",
    "b/make-fixture.ps1": "x\ny\n",
    "CONTEXT.md": "the guard at `make-fixture.ps1:1`.\n",
  };
  const { ambiguous, missing } = analyze(files);
  assert.equal(missing.length, 0);
  assert.equal(ambiguous.length, 1, "an ambiguous basename was not reported");
  assert.equal(ambiguous[0].found.length, 2);

  // Control: writing enough of the path to pick one file resolves it.
  const ok = analyze({
    "a/make-fixture.ps1": "x\ny\n",
    "b/make-fixture.ps1": "x\ny\n",
    "CONTEXT.md": "the guard at `a/make-fixture.ps1:1`.\n",
  });
  assert.equal(ok.ambiguous.length, 0, "a path-qualified citation was still reported ambiguous");
});

test("CATCHES: a proximity tie-break resolves a colliding basename cited from a nearby probe", () => {
  // The disambiguation the checker inherits from its sibling: a bare basename
  // cited from within one of the colliding directories resolves to the nearer
  // file, and does NOT count as ambiguous. (Fixtures avoid the `probe-` prefix
  // so the sibling probe checker does not also scan this file's strings.)
  const files = {
    "spikes/pp/fx-x.ps1": "a\nb\n",
    "spikes/iso/fx-x.ps1": "a\nb\n",
    "spikes/pp/other.ps1": "cite `fx-x.ps1:1`\n",
  };
  const { ambiguous } = analyze(files);
  assert.equal(ambiguous.length, 0, "the nearer colliding probe should win outright");
});

test("CATCHES: a line past the end of the cited file is out of range", () => {
  const files = {
    "src/a.mjs": "one\ntwo\nthree\n", // split on \n -> 4 entries (trailing "")
    "doc.md": "at `src/a.mjs:99`.\n",
  };
  const { outOfRange, missing, ambiguous } = analyze(files);
  assert.equal(missing.length, 0);
  assert.equal(ambiguous.length, 0);
  assert.equal(outOfRange.length, 1, "an out-of-range line was not reported");
  assert.equal(outOfRange[0].written, "src/a.mjs:99");

  // The HIGH end of a range is what is checked: a range straddling the end must
  // still fail. If the checker ever tested lo instead of hi this goes red.
  const straddle = analyze({
    "src/a.mjs": "one\ntwo\nthree\n",
    "doc.md": "at `src/a.mjs:3-99`.\n",
  });
  assert.equal(straddle.outOfRange.length, 1, "a range straddling EOF was not reported");
});

test("NEGATIVE CONTROL: a citation onto real but unrelated code passes — the class the checker cannot see", () => {
  // This is the exact shape of the six #160 rots the checker is blind to, and
  // the assertion that keeps its coverage claim honest. `word-pids.mjs:1` here
  // is a live, in-range coordinate on a real file — it just lands on the wrong
  // line for what the prose claims ("pure set difference"). The checker has no
  // way to know the intended referent, so it MUST pass. If it ever reports this,
  // it is pretending to a coverage it does not have.
  const files = {
    "test/word-pids.mjs": "import x from 'y';\n// unrelated line\nexport const z = 1;\n",
    "CONTEXT.md": "`newWordPids` (`word-pids.mjs:1`) is a pure set difference\n",
  };
  const { missing, ambiguous, outOfRange } = analyze(files);
  assert.equal(missing.length, 0, "a live in-range coordinate must not be reported missing");
  assert.equal(ambiguous.length, 0);
  assert.equal(outOfRange.length, 0, "a live in-range coordinate must not be reported out of range");

  // Stated as an equality so the point is explicit: nothing is flagged, even
  // though a human knows the citation is wrong. That is the coverage boundary.
  assert.deepEqual(
    { missing: missing.length, ambiguous: ambiguous.length, outOfRange: outOfRange.length },
    { missing: 0, ambiguous: 0, outOfRange: 0 },
  );
});

test("a citation onto a real, correct line passes", () => {
  const files = {
    "test/word-pids.mjs": "a\nb\nexport async function newWordPids() {}\n",
    "CONTEXT.md": "`newWordPids` (`word-pids.mjs:3`) is a pure set difference\n",
  };
  const { missing, ambiguous, outOfRange } = analyze(files);
  assert.equal(missing.length + ambiguous.length + outOfRange.length, 0);
});

test("both citations on a shared line are seen — the >1-per-line blind spot", () => {
  // The parser blind spot found reconciling the #160 audit: two positional
  // citations on one line. A one-per-line scan sees the first and drops the
  // second. The property currently holds only because `line.match(/…/g)`
  // enumerates all matches; a later refactor to `exec`, a capture-group change,
  // or a `for…of` with an early `break` could silently restore the one-per-line
  // scan, and the checker would still exit 0 while going blind to a whole class.
  //
  // So this asserts BOTH citations are emitted, not merely that the second one
  // is caught when it happens to be missing. Both are made to fail in DIFFERENT
  // buckets — first missing, second out-of-range — so a scan that dropped either
  // would change a count this test pins exactly.
  const files = {
    "spikes/real.mjs": "a\nb\n", // 2 real lines; :99 is out of range
    "spikes/fx-b.ps1": "line\ncite `gone.mjs:5` and `real.mjs:99` together\n",
  };
  const { missing, outOfRange } = analyze(files);
  assert.equal(missing.length, 1, "the first citation on the line was not emitted");
  assert.equal(missing[0].written, "gone.mjs:5");
  assert.equal(outOfRange.length, 1, "the second citation on the line was dropped (one-per-line regression)");
  assert.equal(outOfRange[0].written, "real.mjs:99");

  // The count itself, pinned: exactly two citations were scanned off this line.
  // A scanner that finds nothing because it scanned nothing also exits 0 and
  // still reports a count — this equality is what refuses that.
  assert.equal(
    missing.length + outOfRange.length,
    2,
    "exactly two citations share this line; a count other than 2 means the line was under- or over-scanned",
  );
});

test("every authored positional citation on a line is scanned — count guard against a silent-empty scan", () => {
  // A second, direct guard on the same property, phrased as a total. Three
  // citations across two lines, every one pointing at a missing file, so the
  // emitted count equals the authored count. If a refactor makes the scanner
  // find fewer (or nothing), this equality fails rather than passing green on an
  // empty scan.
  const files = {
    "doc.md": "first `a.mjs:1` and `b.mjs:2` on one line\nthen `c.mjs:3` alone\n",
  };
  const { missing } = analyze(files);
  assert.equal(missing.length, 3, "expected all three authored citations to be scanned and reported missing");
  assert.deepEqual(
    missing.map((m) => m.written).sort(),
    ["a.mjs:1", "b.mjs:2", "c.mjs:3"],
    "a citation was dropped from the scan",
  );
});

test("citations inside a fenced code block are skipped — captured output, not authored claims", () => {
  // The FINDINGS.md shape: a `file:line` inside ```...``` is transcript. It must
  // not be resolved, even when the named file does not exist.
  const files = {
    "spikes/FINDINGS.md": "prose\n```\n\"icon\" appears in [generated/rpc.d.ts:18]\n```\nmore prose\n",
  };
  const { missing, ambiguous, outOfRange } = analyze(files);
  assert.equal(
    missing.length + ambiguous.length + outOfRange.length,
    0,
    "a citation inside a fenced code block was checked; it should be skipped",
  );

  // Control: the SAME missing citation OUTSIDE a fence must be caught, or the
  // test above would pass simply because the checker never catches anything.
  const outside = analyze({
    "spikes/FINDINGS.md": "prose naming generated/rpc.d.ts:18 directly\n",
  });
  assert.equal(outside.missing.length, 1, "the fence-skip test is vacuous: the citation is not caught even outside a fence");
});

test("an unreadable file is reported, never silently skipped", () => {
  // A tracked path the reader cannot open must surface as unreadable, so a guard
  // that quietly declined to look at part of the repo is distinguishable from a
  // clean run. Modelled by a reader that throws for one path.
  const reader = (p) => {
    if (p === "broken.md") {
      const err = new Error("EISDIR");
      err.code = "EISDIR";
      throw err;
    }
    return "ok\n";
  };
  const { unreadable } = analyzeCitations(["broken.md", "fine.md"], reader);
  assert.equal(unreadable.length, 1);
  assert.equal(unreadable[0].path, "broken.md");
  assert.equal(unreadable[0].why, "EISDIR");
});

test("the checker exempts its own two files (source and this test)", () => {
  // The checker names example coordinates (Word-host.mjs:95) in its own header,
  // and this test file carries rot-shaped citations as fixtures throughout; the
  // checker must flag neither, or the gate would fail on its own machinery.
  // Modelled by putting a would-be-missing citation at each exempt path and
  // asserting silence — and, as a control, the SAME citation at a third path is
  // caught, so the exemption is proven narrow rather than global.
  const exemptSource = analyze({
    "tools/check-citation-lines.mjs": "// example `Word-host.mjs:95` in a comment\n",
  });
  assert.equal(exemptSource.missing.length, 0, "the checker flagged a citation in its own source");

  const exemptTest = analyze({
    ".github/extensions/office-canvas/test/unit/check-citation-lines.test.mjs":
      'assert.equal(missing[0].written, "Word-host.mjs:95");\n',
  });
  assert.equal(exemptTest.missing.length, 0, "the checker flagged a citation in its own test file");

  // Control: the identical citation in any OTHER file is still caught. If this
  // did not fail, the exemption would be silently global and the two asserts
  // above would be vacuous.
  const elsewhere = analyze({
    "docs/notes.md": "see `Word-host.mjs:95`\n",
  });
  assert.equal(elsewhere.missing.length, 1, "the exemption leaked: a third file was not checked");
});
