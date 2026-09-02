# git-guard — findings

One question, from #148: **can `gitAvailable()` tell "git is not runnable" from
"this directory is not a repository"?** Those mean opposite things — the first
is a broken environment and must fail loudly, the second is the installed
extension folder and must skip — and a bare `catch` collapsed them into one
`false` at every guarded call site, so a runner with a broken git reported a
green suite that had executed nothing.

Probe: `probes/probe-execfile-error-shapes.mjs`. No Office, no Windows, no
network. Run it with `node`.

## Measured

Node 22, Windows 11, `git version 2.x` on `PATH`:

| arm | `err.code` | `typeof` | `err.stderr` |
| --- | --- | --- | --- |
| spawn failure (binary not on `PATH`) | `ENOENT` | `string` | 0 bytes |
| git ran, directory is not a repository | `128` | `number` | 69 bytes |

**Distinguishable on `typeof err.code`, and only on that.** The messages are not
usable: this machine's Windows and git both report in German, and message
matching is the defect class this repo has been bitten by before.

## What the split may and may not conclude

A **numeric** code is the only evidence that git ran and gave an answer, so it
is the only thing allowed to produce a skip. Everything else — `ENOENT`,
`EACCES`, a signal kill, an absent `code` — means the child never reported an
exit status, which says nothing about whether a repository exists. The guard is
fail-closed in that direction on purpose.

Note what the probe does **not** establish: it does not show that a missing git
is *reachable* in CI. It is not — `ubuntu-latest` with `actions/checkout` has
git by construction. The conflation was real, its consequence was severe, and
the path to it was hypothetical; #148's own comment records that correction.
This measures the mechanism, not the likelihood.

## Why the probe exists beside a unit test that asserts the same thing

`test/unit/git-available.test.mjs` re-asserts both arms on every commit, which
is the stronger *guarantee*. The probe is the stronger *citation*:
`tools/check-citations.mjs` tracks `probe-*.{ps1,mjs}` paths and nothing else,
so prose pointing at a neighbouring test is unchecked and cannot be
mutation-tested (`CONTEXT.md`). It is also the portable arm — it can be pointed
at another OS or Node version by hand, where the unit test only ever measures
whatever machine the suite is running on.
