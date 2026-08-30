// The one definition of "this script pins UTF-8 on that console direction".
//
// Extracted because two tests need it and a second copy would be a second
// record of the same quantity, free to disagree with the first. It already did:
// the regex grown against `word-host.ps1` recognised only that file's spelling
// (`New-Object System.Text.UTF8Encoding`), so the breadth check reported
// `live-word.ps1` as unpinned on the very run where it had just been pinned --
// with `[System.Text.UTF8Encoding]::new($false)`, which is the same object by a
// different constructor. A guard that rejects a correct spelling does not get
// obeyed, it gets worked around.
//
// Office-free; no imports beyond nothing at all.

/**
 * Matches an assignment of any UTF-8 encoding to one `[Console]` encoding
 * property.
 *
 * All three spellings PowerShell offers are accepted, because all three produce
 * a UTF-8 encoding and the rule is about the encoding, not the syntax:
 *
 *   [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
 *   [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
 *   [Console]::InputEncoding = [System.Text.Encoding]::UTF8
 *
 * The `System.` prefix is optional in PowerShell and therefore optional here.
 *
 * Note what is deliberately *not* accepted: `[Text.Encoding]::Default`. It names
 * a different encoding on each runtime -- the ANSI codepage under Windows
 * PowerShell 5.1, UTF-8 under PowerShell 7 on .NET Core -- so a script using it
 * behaves differently depending on which host launched it, with no failure
 * either way. `probe-console-input-encoding.mjs` carries an arm that measures
 * exactly that split rather than asserting it.
 *
 * @param {"InputEncoding" | "OutputEncoding"} property
 */
export const assignsUtf8 = (property) =>
    new RegExp(
        String.raw`\[Console\]::${property}\s*=\s*(` +
            String.raw`New-Object\s+(System\.)?Text\.UTF8Encoding` +
            String.raw`|\[(System\.)?Text\.UTF8Encoding\]::new` +
            String.raw`|\[(System\.)?Text\.Encoding\]::UTF8` +
            String.raw`)`,
    );

/**
 * A PowerShell script with its line comments blanked out and every offset
 * preserved.
 *
 * Blanked because a grep over raw text accepts a **commented-out** setter, and
 * that is not hypothetical: the first form of `word-host-encoding.test.mjs`
 * passed cleanly while the fix it asserted sat behind a `#`. Deleting the line
 * is only one of the ways the fix goes away; disabling it during a debugging
 * session is the likelier one, and the likeliest to be committed by accident.
 *
 * Preserved rather than removed because the ordering check compares positions,
 * and shortening the text would move them.
 *
 * Line comments only: PowerShell's `<# #>` block form does not appear in any
 * `.ps1` in this repo, and a matcher for a construct the tree does not use is
 * untested code sitting in a test.
 *
 * @param {string} source
 */
export const blankComments = (source) => source.replace(/#[^\n]*/g, (m) => " ".repeat(m.length));

/** Matches a read of the console's stdin. */
export const READS_STDIN = /\[Console\]::In\b/;

/**
 * Matches the first touch of the console in any direction.
 *
 * `Error` belongs here with `In` and `Out`: .NET builds and caches the writer
 * behind `[Console]::Error` on first use exactly as it does the other two, so an
 * `OutputEncoding` assignment sitting below a stderr write is too late for that
 * writer while looking correct. Diagnostics are the strand most likely to be
 * written early -- a "cannot find X, giving up" branch near the top of a script
 * is the normal shape -- and, on this machine, the strand most likely to carry
 * non-ASCII, because PowerShell reports its own errors in German.
 */
export const USES_CONSOLE = /\[Console\]::(In|Out|Error)\b/;

/** Matches a write to the console in either outbound direction. */
export const WRITES_CONSOLE = /\[Console\]::(Out|Error)\b/;
