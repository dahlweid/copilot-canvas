// Unit tests for the positional-coordinate gate (`tools/check-citation-lines.mjs`).
//
// The tool was inverted: it used to *validate* coordinates (resolve the file,
// check the line was in range) and now *rejects* them outright, per ADR 0009.
// That changes what a test here has to prove. Validation was checkable and
// worthless — a coordinate correct when written is relocated by the next
// insertion above it, which no resolver can see. Rejection is a claim about
// syntax, so the risk moves to the two edges of the matcher:
//
//   * too narrow, and the gate is silent on the form actually used. #168 was
//     exactly this: four bare `(:1531)` self-references the old matcher could
//     not see, sitting under a safety claim. The CATCHES tests pin each form.
//   * too wide, and it flags a clock time or a JSON value, gets turned off, and
//     protects nothing. The FALSE-POSITIVE CONTROL pins that boundary.
//
// The inversion itself has one load-bearing regression guard: a coordinate onto
// a real, in-range line — the input the old tool deliberately passed — must now
// be rejected. If someone restores resolution semantics, that test goes red.
//
// Every allowance is tested with a control at a non-allowed path, so an
// exemption that leaked and became global would fail rather than pass silently.
//
// The gate is driven through its pure core `findCoordinates(tracked, readFile)`,
// with an in-memory file map. No git, no disk, no shared helper — the fixtures
// are local to this file so nothing another worktree owns is touched. This file
// is on the gate's exempt list precisely because its fixtures are rot-shaped.
// Office-free; runs on ubuntu-latest.
//
// Run: node --test ".github/extensions/office-canvas/test/unit/*.test.mjs"

import { test } from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { fileURLToPath } from "node:url";

const UNIT = path.dirname(fileURLToPath(import.meta.url));
const CHECKER = path.resolve(UNIT, "..", "..", "..", "..", "..", "tools", "check-citation-lines.mjs");
const { findCoordinates, report, verifyPins } = await import(`file://${CHECKER.split(path.sep).join("/")}`);

/**
 * Build a reader over an in-memory { path: text } map. A path not in the map
 * throws an ENOENT-shaped error, exactly as `fs.readFileSync` does, so the
 * gate's unreadable branch is exercised for real.
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
  return findCoordinates(Object.keys(files), readerOver(files));
}

/** The written forms rejected, sorted, for comparing against an expected set. */
function written(result) {
  return result.coordinates.map((c) => c.written).sort();
}

/**
 * Runs `fn` with console.log/error captured, so a test can assert on what the
 * gate actually prints. The message is part of the contract here: this tool
 * shipped a *correct* verdict under a message that claimed more coverage than
 * it had, and the message is what a reader acts on.
 */
function capture(fn) {
  const out = [];
  const realLog = console.log;
  const realError = console.error;
  console.log = (s) => out.push(String(s));
  console.error = (s) => out.push(String(s));
  try {
    return { code: fn(), text: out.join("\n") };
  } finally {
    console.log = realLog;
    console.error = realError;
  }
}

test("REGRESSION GUARD: a coordinate onto a real, in-range line is rejected — the input the old tool passed", () => {
  // This is the whole inversion in one assertion. `word-pids.mjs:3` names a file
  // that exists, at a line that exists, holding the very function the prose
  // claims — the old validate-semantics tool passed it, correctly by its own
  // rules. Under reject semantics it must fail, because resolving today says
  // nothing about resolving after the next insertion above line 3.
  //
  // If anyone reinstates resolution ("only flag coordinates that are broken"),
  // this goes red. It is the test that keeps the inversion from being undone.
  const files = {
    "test/word-pids.mjs": "a\nb\nexport async function newWordPids() {}\n",
    "CONTEXT.md": "`newWordPids` (`word-pids.mjs:3`) is a pure set difference\n",
  };
  const result = analyze(files);
  assert.equal(result.coordinates.length, 1, "a resolvable coordinate was not rejected");
  assert.equal(result.coordinates[0].written, "word-pids.mjs:3");
  assert.equal(result.coordinates[0].kind, "qualified");
  assert.equal(result.coordinates[0].path, "CONTEXT.md");
  assert.equal(result.coordinates[0].line, 1, "the reporting line number is the citing line");

  // Control: the same sentence written the way ADR 0009 asks for passes. Without
  // this, the assertion above could be satisfied by a gate that rejects the file
  // for some unrelated reason, or by one that rejects everything.
  const named = analyze({
    "test/word-pids.mjs": "a\nb\nexport async function newWordPids() {}\n",
    "CONTEXT.md": "`newWordPids` (`word-pids.mjs`) is a pure set difference\n",
  });
  assert.equal(named.coordinates.length, 0, "a name-only reference must pass");
});

test("CATCHES: a bare self-reference — the #168 form the old matcher could not see", () => {
  // The four coordinates in ADR 0009's own evidence are bare: `(:1531)`, with no
  // filename in front. The old matcher required a filename before the colon and
  // returned no matches at all, so the file was reported clean. A bare
  // self-reference is simultaneously the most fragile form (it rots on any
  // insertion above it in its own file) and the least checkable, which is
  // exactly why it has to be pinned here.
  const paren = analyze({
    "src/word/host.ps1": "# every write goes through Set-ParagraphText (:1531)\n",
  });
  assert.equal(paren.coordinates.length, 1, "a parenthesised bare self-reference was not rejected");
  assert.equal(paren.coordinates[0].written, ":1531");
  assert.equal(paren.coordinates[0].kind, "bare");

  const backtick = analyze({ "notes.md": "the guard at `:486` handles this\n" });
  assert.equal(backtick.coordinates.length, 1, "a backticked bare self-reference was not rejected");
  assert.equal(backtick.coordinates[0].written, ":486");

  const range = analyze({ "spikes/x.ps1": "# the job above (:98-100) proves it\n" });
  assert.equal(range.coordinates.length, 1, "a bare range was not rejected");
  assert.equal(range.coordinates[0].written, ":98-100");

  for (const prose of ["see at :126 for the census", "the line :129 above", "lines :12-14 below"]) {
    const r = analyze({ "doc.md": `${prose}\n` });
    assert.equal(r.coordinates.length, 1, `a prose-introduced bare coordinate was missed: ${prose}`);
  }
});

test("FALSE-POSITIVE CONTROL: a colon and digits that is not a citation is left alone", () => {
  // The other edge. A matcher that flagged every `:\d+` would hit a clock time,
  // a JSON value, a port, a ratio — and would be turned off rather than obeyed,
  // protecting nothing. The introducer requirement is what buys the difference,
  // so these must pass while the bare forms above are caught.
  const files = {
    "doc.md":
      "the run started at 09:41 and took 1:30\n" +
      'a payload like {"id":1, "level":2} is fine\n' +
      "listening on localhost:8080\n" +
      "a ratio of 3:1 either way\n",
  };
  const result = analyze(files);
  assert.deepEqual(
    written(result),
    [],
    "a non-citation colon-digit shape was rejected; the gate would be turned off rather than obeyed",
  );

  // Control that the file is genuinely being read, not skipped for some other
  // reason: adding one real coordinate to it must produce exactly one finding.
  const withOne = analyze({ "doc.md": `${files["doc.md"]}and the guard at \`:486\`\n` });
  assert.equal(withOne.coordinates.length, 1, "the false-positive test is vacuous: this file is not being scanned");
});

test("ALLOWANCE: a coordinate pinned to a commit passes, is counted, and the unpinned twin fails", () => {
  // ADR 0009's first exception. A SHA fixes the tree the number indexes, so a
  // pinned coordinate cannot rot. It is reported in `pinned` rather than silently
  // dropped, so a green run can state how much the allowance let through.
  const pinned = analyze({
    "CONTEXT.md": "it appears once, at `structure-map.mjs:266` as of `4abf952`.\n",
  });
  assert.equal(pinned.coordinates.length, 0, "a commit-pinned coordinate was rejected");
  assert.equal(pinned.pinned.length, 1, "a commit-pinned coordinate was not counted as an allowance");
  assert.equal(pinned.pinned[0].written, "structure-map.mjs:266");

  // Control: the identical coordinate without the pin fails. If this passed, the
  // allowance would be matching something other than the pin.
  const bare = analyze({ "CONTEXT.md": "it appears once, at `structure-map.mjs:266`.\n" });
  assert.equal(bare.coordinates.length, 1, "the pin allowance leaked: an unpinned coordinate passed");

  // The pin must follow the coordinate it governs, immediately. A line carrying
  // a pinned and two unpinned coordinates must report both unpinned ones — if
  // the check were "does this line mention a SHA anywhere after this", the
  // coordinate written *before* the pin would be covered by it too, and this
  // deepEqual would be missing an entry.
  const mixed = analyze({
    "CONTEXT.md": "`b.mjs:2` and `a.mjs:1` as of `4abf952`, then `c.mjs:3` after.\n",
  });
  assert.deepEqual(
    written(mixed),
    ["b.mjs:2", "c.mjs:3"],
    "the pin must govern only the coordinate it immediately follows",
  );
  assert.deepEqual(mixed.pinned.map((c) => c.written), ["a.mjs:1"]);
});

test("ALLOWANCE: the three fixture-carrying files are exempt, and the exemption is narrow", () => {
  // ADR 0009's second exception: a coordinate that is the subject under
  // discussion rather than a reference being followed. Three files qualify — the
  // gate itself, this test, and the ADR — and no others.
  for (const p of [
    "tools/check-citation-lines.mjs",
    ".github/extensions/office-canvas/test/unit/check-citation-lines.test.mjs",
    "docs/adr/0009-issue-and-pr-text-quotes-the-claim.md",
  ]) {
    const r = analyze({ [p]: "example `Word-host.mjs:95` and a bare `:1531`\n" });
    assert.equal(r.coordinates.length, 0, `the gate flagged a coordinate in its own exempt file ${p}`);
  }

  // Control: the identical text in any OTHER file is still rejected, both forms.
  // Without this the three asserts above would be satisfied by a global pass.
  const elsewhere = analyze({ "docs/notes.md": "example `Word-host.mjs:95` and a bare `:1531`\n" });
  assert.equal(elsewhere.coordinates.length, 2, "the exemption leaked: a fourth file was not checked");

  // And narrow in the other direction: a *near-miss* path is not exempt. An
  // exemption matched by basename or by prefix would silently cover the whole
  // ADR directory or every test whose name starts the same way.
  const nearMiss = analyze({
    "docs/adr/0010-something-else.md": "see `a.mjs:1`\n",
    "tools/check-citations.mjs": "see `b.mjs:2`\n",
  });
  assert.deepEqual(
    written(nearMiss),
    ["a.mjs:1", "b.mjs:2"],
    "the exemption is matching by prefix or basename, not by exact path",
  );
});

test("a coordinate inside a fenced code block is skipped — and is COUNTED, not silently dropped", () => {
  // ADR 0009's transcript exception: a `file:line` inside ```...``` is the
  // captured output of a probe run. Rewriting it would falsify evidence.
  //
  // The counting half is the part that was missing and that a review caught. The
  // gate reported "no positional coordinate in any tracked file" while two
  // coordinates sat unread inside a fenced transcript in
  // spikes/word-icon/FINDINGS.md, which made the success message false as
  // written. Skipping is right; claiming to have read them is not.
  const fenced = analyze({
    "spikes/FINDINGS.md": "prose\n```\n\"icon\" appears in [generated/rpc.d.ts:18]\n```\nmore prose\n",
  });
  assert.equal(fenced.coordinates.length, 0, "a coordinate inside a fenced block was rejected");
  assert.equal(fenced.skipped.fencedLines, 1, "the skipped fenced line was not counted");
  assert.equal(fenced.skipped.fencedFiles, 1, "the file containing a skipped fence was not counted");

  // Control: the same coordinate outside a fence must be caught, or the test
  // above passes simply because nothing is ever caught in this file.
  const outside = analyze({ "spikes/FINDINGS.md": "prose naming generated/rpc.d.ts:18 directly\n" });
  assert.equal(outside.coordinates.length, 1, "the fence-skip test is vacuous: the coordinate is not caught outside one either");
  assert.equal(outside.skipped.fencedLines, 0, "a file with no fence reported skipped fenced lines");

  // A fence that opens and closes must not swallow the rest of the file. If the
  // toggle were a one-way latch, everything after the first fence would go
  // unscanned — a silent hole exactly where long documents keep their examples.
  const reopened = analyze({
    "spikes/FINDINGS.md": "```\na.mjs:1\n```\nprose citing b.mjs:2 outside\n```\nc.mjs:3\n```\n",
  });
  assert.deepEqual(written(reopened), ["b.mjs:2"], "fence tracking is not toggling; part of the file went unscanned");
  assert.equal(reopened.skipped.fencedLines, 2, "both fenced lines should be counted as unread");
});

test("the skipped lines are MATCHED, not merely counted — the census that would have caught round 1", () => {
  // Counting skipped lines says how much went unread; it does not say whether
  // any of it was a coordinate. The gap between those two is exactly where
  // round 1's High lived: two coordinates sat inside a fenced transcript in
  // spikes/word-icon/FINDINGS.md, and a reviewer had to find them by hand
  // because the gate reported a silence rather than a number.
  //
  // So the matcher runs over the skipped lines too, and the result is reported
  // without failing on it. This test asserts the number is real by making the
  // fenced content differ from the unfenced content.
  const result = analyze({
    "spikes/FINDINGS.md":
      "prose with no coordinate\n" +
      "```\n" +
      '"icon" appears in [generated/rpc.d.ts:18]\n' +
      "a plain output line with no coordinate at all\n" +
      "and generated/session-events.d.ts:13 here\n" +
      "```\n",
    "spikes/OTHER.md": "```\nno coordinates in this fence at all\n```\n",
  });

  assert.equal(result.coordinates.length, 0, "fenced coordinates must not fail the gate");
  assert.equal(result.skipped.fencedLines, 4, "the unread volume is wrong");
  assert.equal(
    result.skipped.fencedCoordinateLines,
    2,
    "the census is not matching the skipped lines — it is only counting them, which is the state that let round 1's High hide",
  );
  assert.deepEqual(
    result.skipped.fencedCoordinateFiles,
    ["spikes/FINDINGS.md"],
    "the census must name the file holding unread coordinates, and must not name a file whose fence is clean",
  );

  // The number must be able to be zero, or "2" above proves nothing: a census
  // hard-wired to the fenced-line count would pass the assertion above too.
  const clean = analyze({ "spikes/OTHER.md": "```\njust output\nmore output\n```\n" });
  assert.equal(clean.skipped.fencedLines, 2, "control fixture is not fenced as expected");
  assert.equal(clean.skipped.fencedCoordinateLines, 0, "the census counts fenced lines, not coordinate-shaped ones");
  assert.deepEqual(clean.skipped.fencedCoordinateFiles, []);

  // Exempt files are sized the same way, for the same reason: an exemption
  // nobody can measure is an exemption nobody rechecks.
  const exempt = analyze({ "tools/check-citation-lines.mjs": "// `a.mjs:1`\n// prose\n// `b.mjs:2-3`\n" });
  assert.equal(exempt.coordinates.length, 0, "an exempt file must still be exempt");
  assert.equal(exempt.skipped.exemptCoordinateLines, 2, "exempt files are skipped without being sized");
});

test("the success message reports the unread coordinates as a size, never as work to do", () => {
  // The label matters as much as the number. Called "unmigrated coordinates" it
  // reads as a defect count, and the next person drives it to zero by editing a
  // recorded probe run — the precise harm the transcript exception exists to
  // prevent. It must read as the gate's blind spot.
  const result = analyze({
    "spikes/FINDINGS.md": "prose\n```\ncaptured generated/rpc.d.ts:18 here\n```\n",
  });
  const out = capture(() => report(result));

  assert.match(out.text, /1 of those lines is coordinate-shaped, in spikes\/FINDINGS\.md/, "the count and its location are not reported");
  assert.match(out.text, /NOT a count of work\s*\n?\s*left undone/, "the number is not labelled as a size rather than a to-do list");
  assert.match(out.text, /editing\s*\n?\s*transcripts/, "the message does not say why driving the number to zero is the harm");
  assert.equal(out.code, 0, "reporting unread coordinates must not fail the gate");
});

test("the success message states what it read, and never claims a clean tree", () => {
  // The regression guard for the overstatement above. A gate that reports more
  // coverage than it has is worse than a narrower honest one, because the
  // overstatement is what stops the next person looking. So the passing output
  // must (a) qualify its claim to what it read and (b) name the skipped volume.
  const result = analyze({
    "spikes/FINDINGS.md": "prose\n```\ncaptured generated/rpc.d.ts:18 here\n```\n",
    "tools/check-citation-lines.mjs": "// `a.mjs:1`\n",
  });
  assert.equal(result.coordinates.length, 0, "fixture should pass; otherwise this tests the failure path");

  const lines = [];
  const realLog = console.log;
  console.log = (s) => lines.push(s);
  let code;
  try {
    code = report(result);
  } finally {
    console.log = realLog;
  }
  const out = lines.join("\n");

  assert.equal(code, 0);
  assert.match(out, /this gate read/, "the success message does not qualify its claim to what it read");
  assert.doesNotMatch(
    out,
    /no positional coordinate in any tracked file/,
    "the success message makes the unqualified claim that shipped false",
  );
  assert.match(out, /1 line\(s\) inside fenced blocks/, "the skipped fenced volume is not reported");
  assert.match(out, /1 exempt file\(s\)/, "the skipped exempt files are not reported");
});

test("vendored third-party code is skipped, and only under the vendor path", () => {
  // Minified vendor bundles carry coordinate-shaped strings by the dozen and are
  // not authored here, so nothing in them is a reference into this tree.
  const vendored = analyze({
    ".github/extensions/office-canvas/src/ui/vendor/pdf.min.mjs": "n(e,t):9,r=a.mjs:12\n",
  });
  assert.equal(vendored.coordinates.length, 0, "a vendored file was scanned");

  // Control: a file with "vendor" in its name that is not under the vendor path
  // is still scanned, so the skip cannot be widened by an accidental substring.
  const notVendor = analyze({
    ".github/extensions/office-canvas/src/ui/vendor-notes.md": "see `a.mjs:12`\n",
  });
  assert.equal(notVendor.coordinates.length, 1, "the vendor skip is matching a substring, not the path");
});

test("every coordinate on a line is reported — count guard against a silent-empty scan", () => {
  // A scanner that finds nothing also exits 0. These equalities are what refuse
  // that: the emitted count equals the authored count, across one line and
  // several, and a refactor that dropped the second match on a line (an `exec`
  // loop with an early break, say) changes a number pinned here.
  const oneLine = analyze({ "spikes/fx-b.ps1": "cite `gone.mjs:5` and `real.mjs:99` together\n" });
  assert.deepEqual(written(oneLine), ["gone.mjs:5", "real.mjs:99"], "a coordinate sharing a line was dropped");

  const several = analyze({ "doc.md": "first `a.mjs:1` and `b.mjs:2` on one line\nthen `c.mjs:3` alone\n" });
  assert.deepEqual(written(several), ["a.mjs:1", "b.mjs:2", "c.mjs:3"], "a coordinate was dropped from the scan");
  assert.deepEqual(
    several.coordinates.map((c) => c.line),
    [1, 1, 2],
    "the reported line numbers do not match where the coordinates were authored",
  );
});

test("a qualified coordinate is counted once, not also as a bare self-reference", () => {
  // `a.mjs:12` contains `:12`, and the bare matcher would find it if the
  // qualified matches were not blanked out first — turning every cross-file
  // citation into two findings and making the count meaningless. The backtick in
  // front of the filename is a bare-form introducer, so this is not hypothetical.
  const r = analyze({ "doc.md": "see `a.mjs:12` here\n" });
  assert.equal(r.coordinates.length, 1, "a qualified coordinate was double-counted as bare");
  assert.equal(r.coordinates[0].kind, "qualified");
  assert.equal(r.coordinates[0].written, "a.mjs:12");
});

test("the extension whitelist is wide, because a coordinate into data.json rots like one into data.mjs", () => {
  // It was once mjs|ps1|js|ts|md|yml|yaml — the set this tree *mostly* uses.
  // Measured against the tracked list, that left `json`, `css` and `html`
  // invisible although all three are tracked here, and review could find no
  // principle separating them from the rest. Widened.
  const covered = analyze({
    "doc.md":
      "`a.mjs:1` `b.ps1:2` `c.js:3` `d.ts:4` `e.md:5` `validate.yml:6` `f.yaml:7`\n" +
      "`g.json:8` `h.css:9` `i.html:10` `j.txt:11` `k.xml:12` `l.cjs:13` `m.mts:14`\n",
  });
  assert.equal(covered.coordinates.length, 14, "an extension in the documented whitelist is not matched");

  // Still not "any extension": that would match version strings and prose, and a
  // line number into a PNG is not a citation.
  const uncovered = analyze({ "doc.md": "release 1.0.80:2 and image.png:4 and thing.wibble:9\n" });
  assert.equal(uncovered.coordinates.length, 0, "the whitelist has widened to things that are not text files");
});

test("a pin is resolved against git, so a SHA-shaped string that names no commit is rejected", () => {
  // The pin is the one exception an author may apply to themselves, and its
  // whole force is that the SHA fixes the tree the number indexes. Matching the
  // *shape* of a SHA verifies nothing: `deadbee` is seven hex digits and names
  // nothing in this repo. Without resolution, "as of `deadbee`" is a way of
  // writing "trust me" that passes a gate.
  const found = analyze({
    "doc.md": "`a.mjs:1` as of `4abf952`\n`b.mjs:2` as of `deadbee`\n",
  });
  assert.equal(found.coordinates.length, 0, "both are pin-shaped, so neither should fail matching");
  assert.equal(found.pinned.length, 2);
  assert.deepEqual(found.pinned.map((p) => p.pin), ["4abf952", "deadbee"], "the SHA is not captured, so it cannot be resolved");

  const real = new Set(["4abf952"]);
  const pins = verifyPins(found.pinned, (sha) => real.has(sha));
  assert.equal(pins.resolved, 1);
  assert.equal(pins.fabricated.length, 1, "a pin naming no commit was accepted");
  assert.equal(pins.fabricated[0].written, "b.mjs:2");
  assert.equal(pins.unverifiable, 0);

  // And it must actually fail the build, not merely be tallied.
  const out = capture(() => report({ ...found, pins }));
  assert.equal(out.code, 1, "a fabricated pin did not fail the gate");
  assert.match(out.text, /names no commit/);
});

test("a checkout that cannot answer is REPORTED, never counted as a pass", () => {
  // The shallow-clone case, and the reason validate.yml sets fetch-depth: 0.
  // A depth-1 clone holds only the head commit, so every historical pin is
  // unresolvable. Treating "could not look" as "looked and it was fine" is a
  // check that passes because it did nothing — the exact failure this gate
  // exists to argue against.
  const found = analyze({ "doc.md": "`a.mjs:1` as of `4abf952`\n" });
  const pins = verifyPins(found.pinned, () => null);

  assert.equal(pins.resolved, 0, "an unanswerable lookup was scored as a resolution");
  assert.equal(pins.fabricated.length, 0, "an unanswerable lookup was scored as a forgery");
  assert.equal(pins.unverifiable, 1);

  const out = capture(() => report({ ...found, pins }));
  assert.equal(out.code, 0, "an unverifiable pin must not fail the build; it must be disclosed");
  assert.match(out.text, /could NOT be resolved against git/, "the gate stayed silent about what it could not check");

  // Control: when the pin does resolve, the message must say so rather than
  // printing the same hedge either way.
  const ok = capture(() => report({ ...found, pins: verifyPins(found.pinned, () => true) }));
  assert.match(ok.text, /each resolved\s*\n?\s*to a real commit/);
  assert.doesNotMatch(ok.text, /could NOT be resolved/);
});

test("the failure message enumerates every way past the gate, and offers only the pin", () => {
  // Round 1 and round 2 found the same root twice: the tool's stated exception
  // set was narrower than the set it applied — first the fence, then vendored
  // files. A message that lists three ways past a gate with four is how the
  // fourth stays unexamined.
  const out = capture(() => report(analyze({ "doc.md": "see `a.mjs:12` there\n" })));

  assert.equal(out.code, 1);
  assert.match(out.text, /COMMIT PIN/, "the author's actual remedy is not named");
  assert.match(out.text, /VERBATIM TRANSCRIPT/);
  assert.match(out.text, /SUBJECT UNDER DISCUSSION/);
  assert.match(out.text, /VENDORED AND BINARY/, "the fourth allowance is applied but not disclosed");
  assert.match(out.text, /not remedies/, "the message does not distinguish the one exception an author may apply");
});

test("an unreadable file is reported, never silently skipped", () => {
  // A guard that quietly declines to look at part of the repo is
  // indistinguishable from a clean run — the failure this class of tool exists
  // to stop. Modelled by a reader that throws for one path.
  const reader = (p) => {
    if (p === "broken.md") {
      const err = new Error("EISDIR");
      err.code = "EISDIR";
      throw err;
    }
    return "ok\n";
  };
  const { unreadable, coordinates } = findCoordinates(["broken.md", "fine.md"], reader);
  assert.equal(unreadable.length, 1);
  assert.equal(unreadable[0].path, "broken.md");
  assert.equal(unreadable[0].why, "EISDIR");
  assert.equal(coordinates.length, 0);
});
