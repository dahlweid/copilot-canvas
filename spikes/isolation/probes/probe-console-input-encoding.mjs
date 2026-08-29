// Drives probe-console-input-encoding.ps1 exactly as word-host.mjs drives the
// host: spawn powershell.exe -File, write one line of UTF-8 JSON to stdin.
// Three arms so the probe can discriminate; a probe whose arms all agree has
// measured nothing.
//
// Needs no Word and no extension code, so it stays runnable whatever else is in
// flight: the whole question is what Windows PowerShell does to bytes on their
// way in, and that is answerable with a script that only reports what it read.
//
// Named for the property it measures rather than for the issue, because the
// create_document branch (#26) carries a probe of its own that reaches the same
// defect through that handler, under the name this file would otherwise have
// taken. Two probes of one basename would make every basename-only citation
// ambiguous to tools/check-citations.mjs.
//
// Run: node spikes/isolation/probes/probe-console-input-encoding.mjs
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const script = join(here, "probe-console-input-encoding.ps1");

// Built from codepoints so this source file's own encoding is not a variable.
const TEXT = [0x47, 0x72, 0xfc, 0xdf, 0x65, 0x20, 0xc4, 0xd6, 0xdc]
    .map((c) => String.fromCodePoint(c))
    .join("");
const EXPECTED = [...TEXT]
    .map((ch) => "U+" + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0"))
    .join(" ");

function run(arm) {
    return new Promise((resolve) => {
        const child = spawn(
            "powershell.exe",
            ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script],
            { windowsHide: true, env: { ...process.env, PROBE_ARM: arm } },
        );
        let out = "";
        child.stdout.setEncoding("utf8");
        child.stdout.on("data", (d) => (out += d));
        let err = "";
        child.stderr.setEncoding("utf8");
        child.stderr.on("data", (d) => (err += d));
        child.on("close", () => resolve({ out: out.trim(), err: err.trim() }));
        child.stdin.write(JSON.stringify({ text: TEXT }) + "\n", "utf8");
        child.stdin.end();
    });
}

console.log("expected: " + EXPECTED);
console.log("");
for (const arm of ["control", "utf8-nobom", "utf8-static"]) {
    const { out, err } = await run(arm);
    let parsed;
    try {
        parsed = JSON.parse(out);
    } catch {
        console.log(`${arm.padEnd(12)} UNPARSEABLE: ${out || err}`);
        continue;
    }
    const verdict = parsed.codepoints === EXPECTED ? "INTACT" : "CORRUPTED";
    console.log(`${arm.padEnd(12)} ${verdict.padEnd(10)} applied=${parsed.applied}`);
    console.log(`${"".padEnd(12)} got: ${parsed.codepoints}`);
    if (parsed.setError) console.log(`${"".padEnd(12)} setter threw: ${parsed.setError}`);
    console.log("");
}
