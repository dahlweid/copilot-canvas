// Does `spawn("cmd.exe", ["/c", "start", "", path])` survive an ordinary
// document filename?
//
// Node quotes an argv element only when it contains a space, a tab or a quote.
// `cmd.exe` then applies its *own* parse on top, in which `&`, `|`, `^`, `<`
// and `>` are separators. So a filename with no space and an `&` reaches cmd
// unquoted and is split. This probe observes the split instead of arguing
// about it: it substitutes `echo` for `start`, so the arguments cmd actually
// received are visible in stdout.
//
// Run: node spikes/isolation/probes/probe-open-in-word-quoting.mjs

import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

const NAMES = [
    "plain.docx",
    "with space.docx",
    "R&D.docx",
    "R&D report.docx",
    "100%budget.docx",
    "%PATH%.docx",
    "a(1).docx",
    "it's mine.docx",
    "caret^v.docx",
];

// `echo` stands in for `start`: same cmd parse, observable result.
const viaCmd = (p) => execFileAsync("cmd.exe", ["/c", "echo", p], { windowsHide: true });

// `echo` prints the quotes Node added; `start` consumes them. Strip one
// balanced pair so a quoted-and-therefore-safe path is not counted as a
// corruption. Without this the probe overstates the defect by 3 of 9.
const unquote = (s) => (s.startsWith('"') && s.endsWith('"') ? s.slice(1, -1) : s);

// The argv-only alternative: no shell parse at all. `where` is a real
// CommandLineToArgvW consumer, so it reports back exactly one argument.
const viaArgv = (p) =>
    execFileAsync(process.execPath, ["-e", "process.stdout.write(process.argv[1])", p], {
        windowsHide: true,
    });

const rows = [];

for (const name of NAMES) {
    const full = `C:\\Docs\\${name}`;

    let cmdSaw = null;
    let cmdErr = null;
    try {
        const { stdout } = await viaCmd(full);
        cmdSaw = unquote(stdout.trim());
    } catch (err) {
        cmdErr = err.code ?? err.message;
    }

    let argvSaw = null;
    try {
        const { stdout } = await viaArgv(full);
        argvSaw = stdout;
    } catch (err) {
        argvSaw = `THREW ${err.code ?? err.message}`;
    }

    rows.push({
        name,
        cmd: cmdErr ? `ERROR ${cmdErr}` : cmdSaw,
        cmdIntact: cmdSaw === full,
        argvIntact: argvSaw === full,
    });
}

const pad = (s, n) => String(s).padEnd(n);
console.log(pad("filename", 20), pad("cmd.exe /c saw", 34), pad("cmd ok", 8), "argv ok");
console.log("-".repeat(80));
for (const r of rows) {
    console.log(pad(r.name, 20), pad(r.cmd, 34), pad(r.cmdIntact ? "yes" : "NO", 8), r.argvIntact ? "yes" : "NO");
}

const broken = rows.filter((r) => !r.cmdIntact);
console.log("");
console.log(`${broken.length} of ${rows.length} filenames are corrupted by the cmd.exe parse.`);
if (broken.length) console.log("corrupted:", broken.map((r) => r.name).join(", "));
console.log(`${rows.filter((r) => !r.argvIntact).length} of ${rows.length} are corrupted on the argv-only path.`);
