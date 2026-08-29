#!/usr/bin/env node
// Every probe cited by name in a tracked file must exist in *this* repository.
//
// This exists because of a specific, repeatable mistake rather than a general
// tidiness worry. Work here is split across sibling git worktrees — one per
// stack layer, all under the same parent directory — and the coordinator runs
// other sessions' probes to verify their claims rather than accepting them on
// report. A probe run that way is real: it executes, it prints a table, the
// table goes into an ADR. But it lives on *another branch*, so the citation
// that quotes it dangles for everyone reading `main`.
//
// That is exactly what happened: five files landed citing
// `probe-share-vs-access.ps1`, which was never on `main` and still resolved on
// this machine, at
// C:\...\copilot-canvas\<other-worktree>\spikes\isolation\probes\.
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

import { execFileSync } from "node:child_process";
import { readFileSync } from "node:fs";

const CITATION = /\bprobe-[A-Za-z0-9._-]*\.(?:ps1|mjs)\b/g;

// Binary and vendored paths have no citations and would only add noise.
const SKIP = /^(?:\.git\/|.*\.(?:png|jpg|jpeg|gif|pdf|docx|pptx|xlsx|zip|ico)$)/i;

function git(...args) {
  return execFileSync("git", args, { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
}

const tracked = git("ls-files", "-z").split("\0").filter(Boolean);

// The set of probes that actually ship, keyed by basename. Citations in prose
// are usually bare filenames rather than full paths, so basename is the only
// key both sides reliably share.
const existing = new Set(
  tracked
    .filter((p) => /(?:^|\/)probe-[^/]*\.(?:ps1|mjs)$/.test(p))
    .map((p) => p.slice(p.lastIndexOf("/") + 1)),
);

const dangling = [];

for (const path of tracked) {
  if (SKIP.test(path)) continue;

  let text;
  try {
    text = readFileSync(path, "utf8");
  } catch {
    continue; // unreadable or genuinely binary; nothing to cite
  }
  if (text.includes("\0")) continue;

  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    // This file names the very probe it was written about, in the comment
    // above. Exempting the checker by path keeps the exemption from
    // spreading: no other file gets to opt out.
    if (path === "tools/check-citations.mjs") continue;

    for (const name of lines[i].match(CITATION) ?? []) {
      if (!existing.has(name)) {
        dangling.push({ path, line: i + 1, name, text: lines[i].trim() });
      }
    }
  }
}

if (dangling.length > 0) {
  console.error(`FAIL — ${dangling.length} citation(s) name a probe not tracked in this repo:\n`);
  for (const d of dangling) {
    console.error(`  ${d.path}:${d.line}  ->  ${d.name}`);
    console.error(`    ${d.text}\n`);
  }
  console.error(
    "If the probe exists in a sibling worktree, it is on another branch and is\n" +
      "not evidence anyone reading this branch can follow. Either bring it into\n" +
      "this branch or cite a probe that ships with the claim.",
  );
  process.exit(1);
}

console.log(`OK — every probe citation resolves (${existing.size} probe(s) tracked).`);
