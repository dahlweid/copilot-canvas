#!/usr/bin/env node
// What does `node --test` put at the start of its summary lines when its stdout
// is a pipe, and does a gate keyed on `^# ` survive it? Issue #117.
//
// This is the load-bearing instrument the issue's own resolution calls for: a
// committed, re-runnable probe that runs the same invocation the gates run,
// captures the redirected output, and proves — with a negative control — that
// the `^# ` form matches nothing *while the suite is red*, which is the silent
// failure the whole issue is about.
//
// It needs no Office and touches nothing in the repo: it writes a two-test
// passing suite and a one-failure broken suite into an OS temp directory, runs
// `node --test` against each with stdout captured (never a TTY), and reads the
// summary back. Office-free, so it can be re-run anywhere the repo is checked
// out. It is a spike probe rather than a unit test because it shells out to a
// second `node --test` and asserts on that child's *console output shape*, which
// is a property of the runner and the invocation, not of this codebase.
//
// ## What the issue claimed, and what this settles
//
// The issue stated the reporter is chosen by TTY detection, so `^# pass`
// "succeeds when run interactively and silently matches nothing when piped".
// The coordinator refuted the mechanism by measurement on node v24.18.0: the
// piped arm emits the **spec** reporter's `ℹ` (U+2139), not tap's `#`. So the
// trap is not intermittent — on this version `^# ` matches nothing *however the
// command is invoked*, which is worse than an invocation-dependent trap.
//
// This probe cannot hand a child a real console without a pty dependency the
// extension folder is forbidden to carry, so it does not claim to measure the
// TTY arm. It measures the arm the gates actually run in — redirected — and
// records the glyph it finds. If a future node changes the default, re-running
// this turns the recorded glyph red against the assertion below rather than
// leaving a document asserting something about a version nobody runs.
//
// Usage:  node spikes/test-reporter-gate/probes/probe-reporter-glyph.mjs

import { execFileSync } from "node:child_process";
import { mkdtempSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const SPEC_GLYPH = "\u2139"; // ℹ — node's spec reporter
const TAP_GLYPH = "#"; //       node's tap reporter

// The robust matcher the gates must use: accept BOTH glyphs, key on the label,
// and never depend on which reporter node picked. Anchored, label-bearing, and
// it captures the count so a caller can assert the count is a number.
const passLine = /^(?:#|\u2139)\s+pass\s+(\d+)\b/;
const failLine = /^(?:#|\u2139)\s+fail\s+(\d+)\b/;

// The fragile form the issue is about — tap's `#` only. This is the form that
// must be shown matching nothing, so nobody writes it again.
const fragilePass = /^# pass\b/;

const PASS_SRC = `import test from "node:test";
test("a", () => {});
test("b", () => {});
`;

const BROKEN_SRC = `import test from "node:test";
import assert from "node:assert";
test("a", () => {});
test("b-broken", () => { assert.equal(1, 2); });
`;

/**
 * Run `node --test <file>` with stdout+stderr CAPTURED (a pipe, never a TTY),
 * as a gate or CI step does. Returns the combined output and the child's exit
 * code. `execFileSync` throws on a non-zero exit, so the broken arm's status is
 * read from the thrown error rather than lost.
 */
function runTest(file, extraArgs = []) {
  const args = ["--test", ...extraArgs, file];
  try {
    const out = execFileSync(process.execPath, args, { encoding: "utf8" });
    return { out, code: 0 };
  } catch (err) {
    // Node's runner exits 1 on a failing suite; that is expected on the broken
    // arm. Anything without captured stdout is a real failure to launch.
    if (typeof err.stdout === "string") return { out: err.stdout, code: err.status ?? 1 };
    throw err;
  }
}

function summaryLine(out, re) {
  for (const raw of out.split(/\r?\n/)) {
    const m = re.exec(raw);
    if (m) return { line: raw, count: Number(m[1]) };
  }
  return null;
}

const dir = mkdtempSync(join(tmpdir(), "reporter-gate-"));
const passFile = join(dir, "pass.test.mjs");
const brokenFile = join(dir, "broken.test.mjs");
writeFileSync(passFile, PASS_SRC);
writeFileSync(brokenFile, BROKEN_SRC);

const failures = [];
const record = [];

try {
  // ---- Arm 1: default reporter, redirected, PASSING ----------------------
  const good = runTest(passFile);
  const goodPass = summaryLine(good.out, passLine);
  const goodFragile = good.out.split(/\r?\n/).some((l) => fragilePass.test(l));
  const goodGlyph = goodPass ? goodPass.line.charCodeAt(0) : null;

  record.push(`node: ${process.version}  platform: ${process.platform}`);
  record.push("");
  record.push("Arm 1 — default reporter, stdout captured (non-TTY), passing suite:");
  record.push(`  exit code                : ${good.code}`);
  record.push(`  summary line             : ${goodPass ? JSON.stringify(goodPass.line) : "(none)"}`);
  record.push(
    `  first char              : U+${goodGlyph == null ? "????" : goodGlyph.toString(16).toUpperCase().padStart(4, "0")}` +
      (goodGlyph === SPEC_GLYPH.charCodeAt(0)
        ? " (ℹ — spec reporter)"
        : goodGlyph === TAP_GLYPH.charCodeAt(0)
          ? " (# — tap reporter)"
          : ""),
  );
  record.push(`  both-glyph matcher       : pass ${goodPass ? goodPass.count : "MISS"}`);
  record.push(`  fragile '^# pass' matches: ${goodFragile}`);

  if (!goodPass) failures.push("both-glyph matcher did not match the passing summary");
  if (good.code !== 0) failures.push(`passing suite exited ${good.code}, expected 0`);

  // ---- Arm 2: explicit tap, PASSING --------------------------------------
  const tap = runTest(passFile, ["--test-reporter=tap"]);
  const tapPass = summaryLine(tap.out, passLine);
  const tapFragile = tap.out.split(/\r?\n/).some((l) => fragilePass.test(l));
  record.push("");
  record.push("Arm 2 — pinned --test-reporter=tap, captured, passing suite:");
  record.push(`  summary line             : ${tapPass ? JSON.stringify(tapPass.line) : "(none)"}`);
  record.push(`  fragile '^# pass' matches: ${tapFragile}`);
  if (!tapFragile) failures.push("pinned tap reporter did not produce a '# pass' line");

  // ---- Arm 3: NEGATIVE CONTROL — default reporter, redirected, BROKEN -----
  // The point of the whole issue. With the suite RED, the fragile '^# pass'
  // form still matches nothing — so a gate written that way reports no verdict
  // on a broken suite. The both-glyph matcher, by contrast, reads a non-zero
  // fail count and can drive the gate red.
  const bad = runTest(brokenFile);
  const badFragile = bad.out.split(/\r?\n/).some((l) => fragilePass.test(l));
  const badFail = summaryLine(bad.out, failLine);
  record.push("");
  record.push("Arm 3 — NEGATIVE CONTROL: default reporter, captured, BROKEN suite:");
  record.push(`  exit code                : ${bad.code}`);
  record.push(`  fragile '^# pass' matches: ${badFragile}  <- must be false: it is silent on a red suite`);
  record.push(`  both-glyph fail line     : ${badFail ? JSON.stringify(badFail.line) : "(none)"}`);
  record.push(`  both-glyph fail count    : ${badFail ? badFail.count : "MISS"}`);

  if (badFragile) failures.push("fragile '^# pass' unexpectedly matched on the broken suite — control invalid");
  if (!badFail) failures.push("both-glyph fail matcher did not find a fail line on the broken suite");
  else if (badFail.count < 1) failures.push(`broken suite reported fail ${badFail.count}, expected >= 1`);
  if (bad.code === 0) failures.push("broken suite exited 0 — the runner did not signal failure");
} finally {
  rmSync(dir, { recursive: true, force: true });
}

console.log(record.join("\n"));
console.log("");

if (failures.length > 0) {
  console.error(`FAIL — ${failures.length} expectation(s) not met:`);
  for (const f of failures) console.error(`  - ${f}`);
  process.exit(1);
}

console.log(
  "OK — on this node the redirected default reporter emits the spec glyph, the\n" +
    "fragile '^# pass' form matches nothing on BOTH a passing and a broken suite,\n" +
    "and the both-glyph label matcher reads pass on green and fail on red.",
);
