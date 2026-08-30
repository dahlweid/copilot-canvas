# An orphaned panel is not recovered from inside the extension

Reloading the extension kills the host process, and every `ViewerInstance` HTTP
server dies with it. The app then rehydrates by calling `open` on the fresh
host, which mints a new server on a new ephemeral port. Nothing re-navigates the
panel's webview, so it goes on requesting the dead port forever. This is #77.

The question #77 left open, and this record settles, is whether the panel
**should** recover on its own. **It should not — not from here.** The only
mechanism available inside this repo is a deterministic port per `instanceId`,
and it is rejected below on three grounds. The correct fix is a host-app change
and is named at the end.

This is a decision about what *not* to build, which is why it is written down.
Without it the next reader finds a small, obvious-looking patch and no record of
why it was passed over.

## The gap any fix has to cross

`server.mjs` binds `this.#server.listen(0, "127.0.0.1", …)` and sets `this.url`
from whatever port the OS hands back. The app stores that URL in the panel
record at open time. So after a restart there are two URLs in play: the fresh one
the extension knows, and the stale one the panel is pointing at. Closing the gap
means either **moving the panel to the new port** — which the extension cannot
do, as it has no way to push a navigation to its own panel — or **putting the new
server back on the old port**.

Only the second is reachable from this repo, so it is the one that had to be
argued.

## Why a deterministic port is rejected

### It squats on a range the operating system is actively allocating from

Measured on this machine with `netsh int ipv4 show dynamicport tcp`: the dynamic
range is **49152–65535**, 16384 ports, and every viewer port ever recorded on #77
falls inside it. A fixed port per `instanceId` would have to be chosen from that
same range, which the OS hands out to every process that asks for an ephemeral
port. Losing the race is not an exotic failure; it is the expected behaviour of a
range designed to be allocated from.

It is worse than a race. `netsh int ipv4 show excludedportrange tcp` reports that
Windows has **already carved reserved blocks out of that range** on this machine
— 49677–49776, 50000–50059, 50229–50328 and more. A deterministic port that
lands in one of those is not merely contended, it is **permanently unbindable**,
and which blocks exist varies by machine and by boot.

### Losing the race turns a dead panel into a lying one

This is the argument that decides it. The panel's URL is baked in, and the whole
point of the change would be that the panel keeps requesting it after a restart.
So if our server does not win the port, the panel's webview loads **whatever else
answers there** — inside a trusted surface that still carries the user's document
name in its tab, in an app where the panel is understood to be showing their
file.

Today's failure mode is a panel that has stopped and says so. The proposed fix's
failure mode is a panel that appears to work and shows content from a server
nobody in this system opened. Trading an honest stop for a silent substitution is
the wrong direction, and it is a security argument rather than a tidiness one.

### The fallback is worse than either branch

The obvious mitigation — try the deterministic port, fall back to `listen(0)` if
it is taken — produces a recovery that works sometimes, fails silently the rest
of the time, and gives the user no way to tell which happened. A mechanism that
is unreliable *and* silent about its own reliability is harder to reason about
than no mechanism.

## The timing constraint, which no port scheme escapes

Even granting a won port, recovery is bounded by something else entirely.

`connection-status.mjs` gives up at `DEADLINE_MS = 10_000`: it closes the
`EventSource` and sets a `finished` flag that nothing clears. That is deliberate
(#66, #72) and it is backed by a measured defect — a panel that claimed
"Reconnecting…" for roughly 90 minutes against a port nothing was listening on.

So a deterministic port recovers a panel **only if the replacement server binds
inside that 10-second window**. Past it the source is closed and terminal, and no
port scheme reaches it. Whether a host restart plus rehydrate typically completes
inside 10 s is **not measured here**, and the decision does not rest on which
side it usually falls — "usually" is precisely the problem. The losing branch is
permanent and indistinguishable from the winning one.

Widening the window means walking back #66's terminal give-up and restoring an
indefinite retry. That is the defect #66 exists to fix. **Recovery and #72's
honesty pull against each other**, and #72 is the one carrying a measurement.

## What stays, and what should change instead

**#72's message stays exactly as it is.**

> Connection lost — this panel has stopped updating and cannot resume.

It is true, it names no cause it cannot know, and it promises no remedy. Because
no recovery is being added, it does not become a lie, and it is not touched. A
message promising nothing beside a mechanism that does something is the shape
this decision avoids.

**The fix belongs in the host app.** The extension already returns the fresh URL
from `open` on rehydrate — `extension.mjs` returns `url: instance.url`, and
re-opening the same `instanceId` correctly reuses the live instance. The
extension side is idempotent and is not at fault. What is missing is that
**nothing re-navigates the panel's webview to the URL `open` just returned.**
That is host-app behaviour, and it is the whole fix: no port scheme, no retry
window, no squatting, and the panel moves to a server this system actually
started.

Two smaller host-app gaps found alongside it, recorded because they shaped this
decision:

- There is **no counterpart to `open_canvas`** — an agent cannot close a panel.
  So an investigation into panel lifetime cannot clean up after itself.
- A panel record can outlive its server, so **the panel list is not evidence
  about liveness at the moment you read it, in either direction.** Measured on
  #77: a list showing three panels, where the live servers were the three it
  reported as *gone*. Liveness is answered by
  `Get-NetTCPConnection -LocalPort <p> -State Listen`, not by the list.

  But the list is **reconciled eventually, and by the host app itself.** Two
  measurements bracket one such reconciliation. Roughly ten minutes before a
  durable session archive, six panels were listed and only one of their ports had
  anything bound. Immediately after that archive the list held exactly one entry
  — the live one — and a second pass over all seven ports agreed with the first
  on every port. The five dead entries were purged and the live one was kept.

  This is a single co-occurrence, it was not deliberately triggered, and a redraw
  that would have happened anyway is not excluded. Taken only for what it shows:
  **purging dead panel records is already within the host app's power and it
  discriminates correctly.** An app that can already tell a dead panel record
  from a live one, and act on it, is the right place to also re-point a live
  panel at its new server. That is an argument for where the fix belongs, not
  evidence that recovery exists.

Servers, by contrast, do not outlive their session: the previous
investigation's `orphan-probe-a` server was measured dead here without anyone
closing it. That is the one-host-per-session fact seen from the other end, and it
is why no in-repo disposal path can be the hook for a fix — see below.

## No disposal hook can carry a fix here

Worth stating because it is the natural place to reach for. Orphaning is **not** a
`close()` that fails to clean up — it is the host process ceasing to exist,
taking every `ViewerInstance` with it. No `close()` runs, no `#closed` flag is
set, no backoff is woken. #86 made disposal cancel in-flight work, which is right
for the path it covers and **cannot fire on this one at all**.

So anything hung off disposal will simply be skipped whenever a reload is the
cause, which is the common case. The same fact explains why a stranded server is
never leaked for long: servers die with their session's host, measured here on
the previous investigation's `orphan-probe-a`. Panels outlive their servers;
servers do not outlive their host.

## What this repo does instead

Nothing about the orphaning, deliberately. What #77 *did* make possible is a
reliable trigger — open a panel, reload the extension, the panel is dead, with no
Word and no race — and that trigger is the first dependable way to exercise #72's
terminal message. The verification #72 and #75 both had to declare open is closed
at the app layer in `test/unit/connection-lost.test.mjs`, which drives the
committed `app.js` against the committed markup and asserts the panel's own
status element. Its limit is stated there: it is Node, not Chromium.

## Measurements

| | instrument |
| --- | --- |
| Windows dynamic TCP range: **49152–65535** (16384 ports) | `netsh int ipv4 show dynamicport tcp` |
| Reserved blocks already excluded *inside* that range (49677–49776, 50000–50059, 50229–50328, …) | `netsh int ipv4 show excludedportrange tcp` |
| Every viewer port recorded on #77 falls inside the dynamic range | the ports named on #77 |
| Of ten viewer ports recorded across #77, **nine have nothing bound**; only 51822 answers | `Get-NetTCPConnection -LocalPort <p> -State Listen` |
| Exactly **one** listening socket exists across every `copilot.exe` on the machine: `127.0.0.1:51822` | `Get-NetTCPConnection -State Listen` joined to `copilot.exe` pids |
| 51822 is a real viewer: `200`, `status: ready`, `demo.docx`, 13 pages | `Invoke-WebRequest http://127.0.0.1:51822/api/state` |
| Six panels listed while five of their ports were dead; after a session archive, one listed and it was the live one. Both passes agree on all seven ports | the two `Listen` passes bracketing the archive |

The port readings describe the world, not the panel list. They are joined to it
only in the direction that holds — nothing bound means that panel is dead —
never the reverse, since a bound port does not prove which record owns it. The
one live server's host process was launched well before this worktree reached
`6dc3369`, so it is serving an earlier tree: a running process is pinned to a SHA
exactly as a figure is, and on this project the only way to reconcile the two is
the reload that orphans the panel.
