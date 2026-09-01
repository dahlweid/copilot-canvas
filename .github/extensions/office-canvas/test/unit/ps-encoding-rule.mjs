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
 * Matches one of three enumerated textual spellings of a UTF-8 assignment to a
 * `[Console]` encoding property.
 *
 * Scope, stated exactly, because this file is the one whose whole job is to be
 * the rule and a rule that overstates itself is the worst place for it:
 *
 * - It recognises **these three spellings and no others**. They are the three
 *   PowerShell offers and the three the tree uses; a fourth way of naming a
 *   UTF-8 encoding would be reported as unpinned. That failure is loud and
 *   safe — it reddens on correct code, which gets noticed and fixed — whereas
 *   the reverse would not be.
 * - It is a **textual** match over the source. It cannot tell code from a
 *   string, so an assignment written inside a quoted string or a here-string
 *   would count as pinning the encoding. No here-string in the tree contains
 *   one today (checked), so this is a known limit rather than an observed
 *   defect. Note the sibling guard in `quit-argument.test.mjs` deliberately
 *   does *not* strip here-strings, because there the leaking `Quit` genuinely
 *   lives inside one — the two guards want opposite things from the same
 *   construct, which is why neither hands this off to a shared stripper.
 * - Comments are handled, and only because they bit: see `blankComments`.
 *
 * What it does not do is verify that the assignment executes, or that it
 * executes first. Ordering is a separate assertion in
 * `console-encoding.test.mjs`, built on `USES_CONSOLE`, and it is separate
 * precisely because this regex cannot see it.
 *
 *   [Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
 *   [Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
 *   [Console]::InputEncoding = [System.Text.Encoding]::UTF8
 *
 * All three are accepted because all three produce a UTF-8 encoding and the
 * rule is about the encoding rather than the syntax. The `System.` prefix is
 * optional in PowerShell and therefore optional here.
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
 * that is not hypothetical: the first form of the `word-host.ps1` encoding
 * assertion passed cleanly while the fix it asserted sat behind a `#`. Deleting
 * the line is only one of the ways the fix goes away; disabling it during a
 * debugging session is the likelier one, and the likeliest to be committed by
 * accident.
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
