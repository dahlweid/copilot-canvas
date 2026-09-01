# Test-reporter gate — findings

Written for issue #117. One question, three arms, and a mechanism that
contradicts the issue's own hypothesis — measured, not reasoned about.

## The question

Every gate report in this repo quotes the unit-suite summary. A verification
step written as

```
node --test … | Select-String '^# pass'
```

was observed to match **nothing while the suite passed and exit 0** — the
worst available outcome, because a run that matched nothing and a run that
passed are the same observable. The issue explained this as TTY detection: tap
when interactive, and — it claimed — tap's `# pass` present when interactive but
absent when piped.

The real question the issue's own §"Provenance" flagged as unmeasured: **what
does `node --test` actually put at the start of its summary lines when its
stdout is a pipe, and does the reporter selection depend on the TTY at all?**

## The probe

`probes/probe-reporter-glyph.mjs`. Office-free and self-contained: it writes a
two-test passing suite and a one-failure broken suite into an OS temp directory,
runs `node --test` against each with stdout **captured** — a pipe, never a TTY,
which is the arm every gate and CI step runs in — and reads the summary back.
Re-runnable anywhere the repo is checked out.

## What it measures — node v24.18.0, Windows

| arm | invocation | summary line | `^# pass` | both-glyph matcher | exit |
| --- | --- | --- | --- | --- | --- |
| 1 | `node --test` (default), passing | `ℹ pass 2` | **0 matches** | `pass 2` | 0 |
| 2 | `node --test --test-reporter=tap`, passing | `# pass 2` | matches | `pass 2` | 0 |
| 3 | `node --test` (default), **broken** | `ℹ fail 1` | **0 matches** | `fail 1` | 1 |

The first character of arm 1's summary line is **U+2139 (`ℹ`)** — node's
**spec** reporter — not tap's `#`. Verified by code point, not by eye.

## What this establishes

**The mechanism in the issue is wrong; the symptom is real.** On this version
the redirected arm selects the spec reporter, so `^# pass` matches nothing
*however the command is invoked*. That is not the intermittent, TTY-dependent
trap the issue described — it is **unconditional**. A session that followed the
issue's step 1 literally (run twice, TTY vs redirected, "the difference is the
finding") would very likely see no difference between the two arms and could
wrongly conclude "no trap". The load-bearing instrument is therefore **arm 3,
the negative control**, not a two-arm comparison.

**Arm 3 is the whole point.** With the suite *red*, the fragile `^# pass` form
still matches nothing — so a gate written that way prints no verdict on a broken
suite, exactly as it prints none on a passing one. The both-glyph matcher, keyed
on the label and accepting either glyph, reads `fail 1` off the same output and
can drive the gate red. A control that only showed the correct form working
would not have demonstrated that the incorrect form fails *silently*.

**Pinning the reporter removes the variable.** Arm 2 shows `--test-reporter=tap`
restores `# pass` regardless of how stdout is attached, which is why issue #117
step 4 (pin the reporter) is the right call rather than an optional extra: the
default has evidently moved from tap to spec already, so a gate that depends on
the default depends on a moving target. This is the same reasoning the repo
applies to parsers — prefer removing the parser to escaping it.

## The rule this yields

Recorded in `CONTEXT.md` beside the other gate traps. A summary matcher must:

- **accept both glyphs and key on the label** — `^(?:#|ℹ)\s+pass\s+\d` — never
  `^# ` alone, which by default never matches on this version;
- **assert the matched line count is non-zero** before trusting a summary — a
  run that matched nothing and a run that passed are the same observable;
- **capture `$LASTEXITCODE` before any pipe** — `Select-String` overwrites it,
  so a check read after the pipe reads the matcher's status, not the runner's;
- and where a gate is scripted, **pin `--test-reporter=tap`** so `^# ` is
  correct by construction and the output shape stops depending on the default.

## What this does not settle

**The TTY arm is unmeasured.** Handing a child process a real console needs a
pty, and the extension folder may carry no `node_modules` (constraint C2), so
there is no clean way to measure it here. What can be said: the *redirected* arm
— the only one the gates run in — selects spec on v24.18.0, and that is the arm
that matters. Whether a TTY selects something different is not required to
establish the trap, because the trap does not depend on it.

**The version is part of the claim.** Everything above is node v24.18.0 on
Windows. The default reporter selection has changed at least once already
(the issue's mechanism describes an older default), so re-run the probe on a new
node before trusting the glyph. The probe asserts its own expectations, so a
changed default turns it red rather than letting a stale figure travel.

---

*Reported for #117. The coordinator ran arms 1–2 first in a throwaway directory
and refuted the TTY mechanism; this probe re-derives those figures and adds the
negative control (arm 3) as a committed, re-runnable instrument.*
