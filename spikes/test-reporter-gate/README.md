# Test-reporter gate

Why a unit-suite gate keyed on `^# ` can pass a broken suite silently, measured
end to end. Issue #117.

`FINDINGS.md` is the record. The short version: `node --test` with its stdout
redirected does **not** emit tap's `# pass` — on node v24.18.0 it emits the
**spec** reporter's `ℹ pass` (U+2139). So a gate matching `^# pass` matches
nothing whether the suite is green or red, and combined with reading exit status
after a pipe it yields *exit 0 with no verdict*.

## Probe

### `probes/probe-reporter-glyph.mjs`

Self-contained and Office-free. It writes a two-test passing suite and a
one-failure broken suite into an OS temp directory, runs `node --test` against
each with stdout **captured** (a pipe, never a TTY — the arm the gates run in),
and asserts three things:

- the redirected default reporter's summary line starts with `ℹ`, and the
  both-glyph label matcher reads `pass 2`;
- pinning `--test-reporter=tap` restores the `# pass` line;
- **the negative control** — on the broken suite the fragile `^# pass` form
  matches nothing while the both-glyph matcher reads `fail 1`. A control that
  only showed the correct form working would not prove the incorrect form fails
  silently, and silent failure is the entire point.

```
node spikes/test-reporter-gate/probes/probe-reporter-glyph.mjs
```

It is a spike probe rather than a unit test because it shells out to a second
`node --test` and asserts on that child's console-output *shape*, which is a
property of the runner and the invocation, not of this codebase. If a future
node changes the default reporter, re-running this turns the recorded glyph red
against its own assertions rather than leaving a stale claim standing.

It does **not** measure the TTY arm: handing a child a real console needs a pty
dependency the extension folder is forbidden to carry. See `FINDINGS.md` §"What
this does not settle".
