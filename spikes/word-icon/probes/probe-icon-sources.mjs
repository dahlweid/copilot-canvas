// Where a Word mark can come from, and where it can be put (#68).
//
// Three questions, three arms, each of which can fail. Run:
//
//   node spikes/word-icon/probes/probe-icon-sources.mjs
//
// Windows and an installed Word for arm 2; network for arm 1; the installed
// Copilot CLI's bundled SDK for arm 3. Each arm reports `skipped` rather than
// passing when its prerequisite is absent, because an arm that cannot run and
// says "ok" is the failure mode this repo pays for most.

import { spawn } from "node:child_process";
import { readFile, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import os from "node:os";

const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..");
const EXTRACTOR = path.join(REPO, ".github", "extensions", "office-canvas", "src", "word", "word-icon.ps1");

const results = [];
const record = (name, outcome, detail) => {
    results.push({ name, outcome, detail });
    console.log(`${outcome.padEnd(7)} ${name}\n        ${detail}`);
};

// --- 1. Is there an open-licensed Word mark to use instead? -------------------
//
// The premise behind extracting at runtime is that there is no vendorable mark.
// If simple-icons carried one, the whole apparatus below would be unnecessary,
// so this arm is the one that could make the change wrong.

async function probeSimpleIcons() {
    const slugs = ["microsoft", "microsoftword", "microsoft-word", "microsoftoffice", "libreoffice"];
    const seen = {};
    for (const slug of slugs) {
        try {
            const res = await fetch(`https://cdn.simpleicons.org/${slug}`, { redirect: "follow" });
            seen[slug] = res.status;
        } catch (err) {
            record("simple-icons has no Microsoft mark", "skipped", `no network: ${err.message}`);
            return;
        }
    }

    // `libreoffice` is the control. Without it, four 404s would be equally well
    // explained by a CDN that is down or a URL shape that changed, and the
    // conclusion "the marks are gone" would rest on an instrument nobody checked.
    const detail = Object.entries(seen)
        .map(([slug, status]) => `${slug}=${status}`)
        .join(" ");
    const controlWorks = seen.libreoffice === 200;
    const microsoftAbsent = slugs.slice(0, 4).every((slug) => seen[slug] === 404);

    if (!controlWorks) return record("simple-icons has no Microsoft mark", "FAIL", `control failed: ${detail}`);
    record(
        "simple-icons has no Microsoft mark",
        microsoftAbsent ? "ok" : "FAIL",
        microsoftAbsent ? detail : `a Microsoft mark exists after all: ${detail}`,
    );
}

// --- 2. Does the user's own Word yield one? -----------------------------------

function runExtractor() {
    return new Promise((resolve) => {
        const child = spawn(
            "powershell.exe",
            ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", EXTRACTOR],
            { stdio: ["ignore", "pipe", "pipe"], windowsHide: true },
        );
        let out = "";
        let err = "";
        child.stdout.on("data", (c) => (out += c));
        child.stderr.on("data", (c) => (err += c));
        child.on("error", (e) => resolve({ code: -1, out: "", err: e.message }));
        child.on("close", (code) => resolve({ code, out: out.trim(), err: err.trim() }));
    });
}

/** Extraction must not launch Word. Counted rather than taken from the docs. */
async function wordProcessCount() {
    if (os.platform() !== "win32") return 0;
    return await new Promise((resolve) => {
        const child = spawn("tasklist.exe", ["/FI", "IMAGENAME eq WINWORD.EXE", "/NH"], { windowsHide: true });
        let out = "";
        child.stdout.on("data", (c) => (out += c));
        child.on("error", () => resolve(-1));
        child.on("close", () => resolve((out.match(/WINWORD\.EXE/gi) ?? []).length));
    });
}

async function probeExtraction() {
    if (os.platform() !== "win32") {
        return record("the installed Word yields a PNG", "skipped", `not Windows (${os.platform()})`);
    }

    const before = await wordProcessCount();
    const started = Date.now();
    const { code, out, err } = await runExtractor();
    const took = Date.now() - started;

    if (code === 2) return record("the installed Word yields a PNG", "skipped", `no Word on this machine: ${err}`);
    if (code !== 0) return record("the installed Word yields a PNG", "FAIL", `exit ${code}: ${err}`);

    const png = Buffer.from(out, "base64");
    const signature = png.subarray(0, 8).toString("hex");
    const after = await wordProcessCount();

    // The size is read out of the IHDR chunk rather than assumed: `width x
    // height` is the claim the CSS is written against.
    const width = png.readUInt32BE(16);
    const height = png.readUInt32BE(20);

    const ok = signature === "89504e470d0a1a0a" && png.length > 8 && after === before;
    record(
        "the installed Word yields a PNG",
        ok ? "ok" : "FAIL",
        `${width}x${height}, ${png.length} bytes, ${out.length} base64 chars, ${took}ms; ` +
            `WINWORD processes ${before} -> ${after}`,
    );
}

// --- 3. Can the canvas *tab* carry it? ----------------------------------------
//
// The question the issue leaves open, and the one worth being exact about,
// because the answer decides whether a third placement ships at all. The claim
// under test is that `canvas.declaration.icon` is supported and takes a file
// path relative to the entrypoint -- which would fight "never a resource in the
// sources" head-on.

async function findSdk() {
    const root = path.join(os.homedir(), "AppData", "Local", "copilot", "pkg");
    const candidates = [];
    const walk = async (dir, depth) => {
        if (depth > 3) return;
        let entries;
        try {
            entries = await readdir(dir, { withFileTypes: true });
        } catch {
            return;
        }
        for (const entry of entries) {
            if (!entry.isDirectory()) continue;
            const full = path.join(dir, entry.name);
            if (entry.name === "copilot-sdk") candidates.push(full);
            else await walk(full, depth + 1);
        }
    };
    await walk(root, 0);
    return candidates.sort().at(-1) ?? null;
}

async function probeTabIcon() {
    const sdk = await findSdk();
    if (!sdk) return record("the canvas tab can carry an icon", "skipped", "no bundled SDK found under ~/AppData");

    const files = [];
    const walk = async (dir) => {
        for (const entry of await readdir(dir, { withFileTypes: true })) {
            const full = path.join(dir, entry.name);
            if (entry.isDirectory()) await walk(full);
            else if (/\.(js|mjs|ts)$/.test(entry.name)) files.push(full);
        }
    };
    await walk(sdk);

    // Two independent instruments. A grep for the word alone would be satisfied
    // by a comment; the declaration's shape is what decides whether a value
    // could reach the wire at all.
    const mentions = [];
    let declarationLiteral = null;
    for (const file of files) {
        const text = await readFile(file, "utf8");
        const hits = (text.match(/\bicon\b/gi) ?? []).length;
        if (hits) mentions.push(`${path.relative(sdk, file)}:${hits}`);
        declarationLiteral ??= text.match(/this\.declaration\s*=\s*\{[\s\S]{0,400}?\n\s*\}/)?.[0] ?? null;
    }

    const spreads = declarationLiteral ? /\.\.\./.test(declarationLiteral) : null;
    const fields = declarationLiteral ? [...declarationLiteral.matchAll(/(\w+)\s*:/g)].map((m) => m[1]) : [];

    // "ok" means the *question was answered*, whichever way it went. A probe
    // that passed only when the feature existed could not report an absence,
    // which is the likely outcome and the one that has to be reportable.
    record(
        "the canvas tab can carry an icon",
        declarationLiteral ? "ok" : "FAIL",
        declarationLiteral
            ? `${path.basename(path.dirname(sdk))}: ${files.length} source files; ` +
                  `"icon" appears in [${mentions.join(", ") || "nothing"}]; ` +
                  `declaration built from [${fields.join(", ")}], spread: ${spreads}`
            : `could not find the declaration literal in ${files.length} files -- instrument broken, not an answer`,
    );
}

await probeSimpleIcons();
await probeExtraction();
await probeTabIcon();

const failed = results.filter((r) => r.outcome === "FAIL");
const skipped = results.filter((r) => r.outcome === "skipped");
console.log(`\n${results.length} arm(s), ${failed.length} failed, ${skipped.length} skipped`);
process.exit(failed.length ? 1 : 0);
