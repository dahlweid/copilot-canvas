# Tool errors

What a tool failure carries from this extension to the agent, measured end to
end. Issue #45.

`FINDINGS.md` is the record. Read it before changing anything at the tool
boundary — the obvious shape (throw a rich error) is measured to deliver
nothing.

## Probes

### `probes/error-channel-probe/`

The rig. A Copilot extension with five tools and three canvas actions, one per
candidate channel, each carrying a distinctive marker.

It has to run **live**, as a real extension inside a real session, because the
discard happens in the host and nothing beneath it can observe the far end.

```
# Windows
$dir = "$env:USERPROFILE\.copilot\session-state\<session-id>\extensions\error-channel-probe"
New-Item -ItemType Directory -Force -Path $dir
Copy-Item spikes\tool-errors\probes\error-channel-probe\extension.mjs $dir
```

Then reload extensions and call each tool and action, recording verbatim what
comes back. Delete the directory and reload again afterwards — the probe tools
are noise in every later session otherwise.

No Word: every arm returns or throws immediately.

### `probes/probe-bridge-discard.mjs`

The citation check. Asserts that the two discard sites quoted in `FINDINGS.md`
are still present in the installed SDK and the newest installed CLI runtime, so a
CLI update that moves or fixes the discard turns the citation red rather than
leaving the document asserting something about a version nobody runs.

```
node spikes/tool-errors/probes/probe-bridge-discard.mjs
```

Machine-local by nature — it reads the installed CLI — so it is a spike probe,
not a test. CI has no Copilot CLI to read.
