# Downloaded documents are edited through Protected View, not by unblocking them

A document that arrived from the internet carries a mark-of-the-web — an
alternate data stream named `Zone.Identifier` containing `ZoneId=3`. Word opens
such a file in Protected View, which refuses automation. Since `edit_document`
must touch the original, this blocks the whole feature for any document the user
downloaded, which is most of them.

We open those documents through
`Application.ProtectedViewWindows.Open(path, password, false, false)` and then
call `.Edit()` on the resulting window, which returns an ordinary writable
`Document`. This is the automation equivalent of the user clicking **Enable
Editing** on the yellow bar, and it is what we do.

The one-line alternative is `Unblock-File`, which deletes the `Zone.Identifier`
stream. We rejected it. It silently removes a security marker from a file the
user owns, it is not restorable, and it changes how *every other* program on the
machine subsequently treats that file — for a change the user never asked for
and is never told about. A tool that edits a paragraph should not also be
downgrading the trust classification of the user's documents.

## Mark-of-the-web does not refuse. It hangs.

This is the part that matters beyond the choice above, and it was measured:
`Documents.Open` on a file carrying `ZoneId=3` **does not return an error. It
never returns at all.** The call hangs indefinitely and the process must be
killed externally. A clean copy of the same document opened in 430 ms as a
control.

Two consequences follow, and both are load-bearing:

- This is a **second, independent reason every `Documents.Open` must be
  timeout-bounded**, alongside the file-lock hang of ADR 0005. The two causes
  are unrelated but the failure shape is identical, so the mitigation is shared.
  Removing the timeout because "the lock is now checked up front" would
  reintroduce the hang through this door.
- **Never probe for Protected View by attempting to open the document.** The
  probe would wedge exactly as the operation does. Detection must be a read of
  the `Zone.Identifier` alternate data stream, which is a file-system call
  costing well under a millisecond, and it is done *before* Word is asked for
  anything.

## What it costs

`.Edit()` writes a **Trust Record** for that file under
`HKCU\Software\Microsoft\Office\16.0\Word\Security\Trusted Documents\TrustRecords`.
This is a real cost and it is stated plainly rather than buried: after an edit,
that document is trusted for that user, and a later plain `Documents.Open` on it
takes 85 ms instead of hanging.

It is the same record a human produces by clicking Enable Editing. It is
per-file, not a blanket setting; it is scoped to the current user; it is visible
and clearable in the Word Trust Center; and the file's own `Zone.Identifier`
survives untouched, verified after the edit. Compared with `Unblock-File` —
global in effect, invisible, irreversible — this is the smaller and more honest
footprint.

## Measurements

| | |
| --- | --- |
| `Documents.Open` on a marked file | never returns (killed externally) |
| Same document, mark removed (control) | 430 ms |
| Protected View open + `.Edit()`, warm | 1076–1548 ms |
| Edit + save through that window | 375–395 ms |
| Full marked-document edit, end to end | 2737 ms |
| Plain open after a trust record exists | 85 ms |
| `Zone.Identifier` after the edit | unchanged |

The Protected View route costs roughly a second more than a direct open. That is
paid only by documents that carry the mark, only on the first edit of a session,
and it buys the file's security marker staying where the user put it.
