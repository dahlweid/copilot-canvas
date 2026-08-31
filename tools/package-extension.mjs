#!/usr/bin/env node
// Produces the installable artefact for one extension folder.
//
// C2 + C3 mean this repository has no conventional build: the runtime loads the
// committed source directly, so packaging is copy-and-verify, not bundle. What
// is left for a packager to do is therefore narrow and worth stating:
//
//   1. decide what is *not* part of the artefact, and say why;
//   2. emit the two shapes the two distribution channels accept —
//      a folder (C3) and a gist payload (C4);
//   3. prove the artefact is still valid by running the validator on it,
//      rather than by asserting the same invariants a second time;
//   4. report size against the C4 budget, so the headroom that the vendored
//      pdf.js worker spends stays visible.
//
// One measured correction shapes all of this: the C4 limits are enforced by
// `install_extension` on **every** source it accepts, repo folder as well as
// gist, and they are decimal — 1,000,000 bytes per file and 5,000,000 in
// total. The restructure plan this repository was built to had held that
// repo-folder install carried no such limit; measurement against the running
// app said otherwise.
//
// Run: node tools/package-extension.mjs [name] [--out dist] [--expect-version v1.0.0]

import { createHash } from "node:crypto";
import { cp, mkdir, readdir, readFile, rm, stat, utimes, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const EXTENSIONS_REL = path.join(".github", "extensions");

// Measured against the running app (v1.0.80), not taken from documentation,
// and decimal rather than binary. Both limits apply to every source
// `install_extension` accepts — gist *and* repo folder — so this is the
// install envelope, not merely the gist one.
const MAX_FILE_BYTES = 1_000_000; // C4: install per-file limit
const MAX_TOTAL_BYTES = 5_000_000; // C4: install total limit

// Fixed mtime on every output file, so a downstream `zip` is byte-reproducible
// rather than carrying whatever the checkout happened to stamp. 1980-01-01Z is
// the zip epoch floor; anything earlier cannot be represented in a DOS date.
const DEFAULT_EPOCH = 315532800;

/**
 * What is deliberately not in the artefact.
 *
 * Every rule carries its reason, and the reasons are printed at package time,
 * because an exclusion list nobody can see is an exclusion list nobody checks.
 * `expected` marks the rules that must match something in this repository today
 * — the rest are guards against material that would be a mistake to ship if it
 * ever appeared, and a unit test holds that distinction.
 */
export const EXCLUSIONS = [
    {
        id: "test",
        expected: true,
        why:
            "Development-only. Nothing under src/ or extension.mjs imports it (the validator's C3 " +
            "import graph would say so otherwise), test/integration/ drives Word through COM and " +
            "cannot run on an installed copy anyway, and make-fixture.ps1 generates a fixture " +
            "no user has a use for.",
        match: (rel) => segments(rel).includes("test"),
    },
    {
        id: "spikes",
        why:
            "Throwaway exploration, already moved out of the extension folder. It carried " +
            "300 KB of JPEG screenshots that C4 refuses outright, so this rule exists to stop a " +
            "returning spike from silently re-entering the artefact.",
        match: (rel) => segments(rel).includes("spikes"),
    },
    {
        id: "artifacts",
        why:
            "The user's own render cache and recents.json, written at runtime by src/store.mjs " +
            "into $COPILOT_HOME/extensions/<name>/artifacts/. Measured on this machine: it holds " +
            "exported PDFs of documents the user opened. Publishing those would be a privacy " +
            "failure, not merely a size one — and the PDFs are binary, which C4 refuses.",
        match: (rel) => segments(rel).includes("artifacts"),
    },
    {
        id: "node_modules",
        why: "C2 forbids it outright; excluded so a stray local install cannot reach a published artefact.",
        match: (rel) => segments(rel).includes("node_modules"),
    },
    {
        id: "output",
        why:
            "This packager's own output, so pointing it at a folder it has already written to does not " +
            "nest. `dist` and `build` also match install_extension's own skip list, so excluding them " +
            "keeps the artefact identical to what a repo-folder install would have produced.",
        match: (rel) => segments(rel).some((s) => s === "dist" || s === "build" || s === "out"),
    },
    {
        id: "vcs",
        why: "Version-control metadata. install_extension copies a folder, not a repository.",
        match: (rel) => segments(rel).includes(".git") || /^\.git(ignore|attributes|modules)$/.test(path.basename(rel)),
    },
    {
        id: "editor",
        why: "Editor and IDE state, specific to one machine and one person.",
        match: (rel) =>
            segments(rel).some((s) => s === ".vscode" || s === ".idea") || /\.code-workspace$/.test(rel),
    },
    {
        id: "os-junk",
        why:
            "Filesystem droppings. .DS_Store and Thumbs.db are binary, so leaving them in would " +
            "fail C4 at share time — long after packaging said it was fine.",
        match: (rel) => [".DS_Store", "Thumbs.db", "desktop.ini"].includes(path.basename(rel)),
    },
    {
        id: "scratch",
        why: "Logs, temporaries and merge leftovers. A .orig or .rej in an artefact means a botched merge shipped.",
        match: (rel) => /(\.(log|tmp|bak|swp|orig|rej)|~)$/.test(rel),
    },
    {
        id: "secrets",
        why: "A .env in a published artefact is a credential leak. There is no version of this that is correct.",
        match: (rel) => /^\.env(\..+)?$/.test(path.basename(rel)),
    },
];

const segments = (rel) => rel.split("/");
const toPosix = (rel) => rel.split(path.sep).join("/");
const sha256 = (buffer) => createHash("sha256").update(buffer).digest("hex");
const kb = (bytes) => `${(bytes / 1024).toFixed(0)} KB`;

class PackagingError extends Error {}

async function listFiles(dir, base = dir) {
    const out = [];
    for (const entry of await readdir(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) out.push(...(await listFiles(full, base)));
        else if (entry.isFile()) out.push(toPosix(path.relative(base, full)));
    }
    return out;
}

/** First matching rule wins, so the reported reason is the one a reader would give. */
export function classify(relPath) {
    return EXCLUSIONS.find((rule) => rule.match(relPath)) ?? null;
}

/**
 * The gist channel is flat and encodes a path by replacing the separator with a
 * backslash — verified by sharing a real extension and reading the filenames
 * back, not assumed.
 *
 * That mapping is only unambiguous while no source filename contains a
 * backslash, which is unrepresentable on Windows but legal on Linux, so it is
 * checked rather than trusted.
 */
export function gistName(relPath) {
    if (relPath.includes("\\")) {
        throw new PackagingError(`'${relPath}' contains a backslash, which the flat gist naming scheme cannot encode`);
    }
    return relPath.split("/").join("\\");
}

/**
 * Builds the gist bundle as an API request body rather than a directory.
 *
 * A directory is not an option: the encoded names contain backslashes, which
 * *are* the path separator on Windows, so `src\ui\app.js` cannot exist as a
 * file there at all. The gist channel's real interface is the POST body
 * anyway, so this is both the portable shape and the directly usable one:
 *
 *   gh api -X POST /gists --input dist/office-canvas-1.0.0.gist.json
 *
 * Key order follows the sorted file list, so the payload is deterministic.
 */
export function buildGistPayload({ name, version, description, files }) {
    const body = { description: description ?? `${name} ${version}`, public: false, files: {} };
    for (const { path: rel, content } of files) {
        const flat = gistName(rel);
        if (flat in body.files) throw new PackagingError(`two files collide on the gist name '${flat}'`);
        body.files[flat] = { content };
    }
    return body;
}

/** Runs the real validator against a tree staged to look like a repository. */
async function validateTree(label, sourceDir, extensionName, stagingRoot) {
    await mkdir(path.join(stagingRoot, EXTENSIONS_REL), { recursive: true });
    await cp(path.join(ROOT, "tools"), path.join(stagingRoot, "tools"), { recursive: true });
    await cp(sourceDir, path.join(stagingRoot, EXTENSIONS_REL, extensionName), { recursive: true });
    try {
        await execFileAsync(process.execPath, [path.join(stagingRoot, "tools", "validate-extensions.mjs")]);
    } catch (err) {
        const output = `${err.stdout ?? ""}${err.stderr ?? ""}`.trim();
        throw new PackagingError(`validation failed for ${label}:\n${output}`);
    }
}

/**
 * Packages one extension folder.
 *
 * The validator runs twice, and the second run is not redundant. Validating the
 * source refuses to package anything the repository would already reject —
 * including material that the exclusions would have hidden, such as a
 * package.json under test/. Validating the artefact then catches the opposite
 * mistake: an exclusion that removed something the extension needs.
 */
export async function packageExtension({
    sourceDir,
    outDir,
    name = path.basename(sourceDir),
    epoch = Number(process.env.SOURCE_DATE_EPOCH ?? DEFAULT_EPOCH),
    expectVersion = null,
    gist = true,
    log = () => {},
} = {}) {
    const manifestPath = path.join(sourceDir, "copilot-extension.json");
    let manifest;
    try {
        manifest = JSON.parse(await readFile(manifestPath, "utf8"));
    } catch (err) {
        throw new PackagingError(`cannot read ${path.relative(ROOT, manifestPath)}: ${err.message}`);
    }

    // `manifest.version` is the manifest *format* version (u32) that
    // install_extension validates; the product version this artefact is named
    // for lives in `productVersion`. See tools/validate-extensions.mjs.
    const productVersion = manifest.productVersion;
    if (!/^\d+\.\d+\.\d+$/.test(productVersion ?? "")) {
        throw new PackagingError(
            `copilot-extension.json productVersion '${productVersion}' is not semver`,
        );
    }

    if (expectVersion !== null) {
        const wanted = String(expectVersion).replace(/^v/, "");
        if (productVersion !== wanted) {
            throw new PackagingError(
                `version mismatch: asked for '${wanted}' but copilot-extension.json productVersion is '${productVersion}'`,
            );
        }
    }

    const staging = path.join(outDir, ".staging");
    // Only this extension's own outputs are cleared. An earlier draft removed
    // `outDir` wholesale, which turns a mistyped `--out .` into deleting the
    // repository.
    if (outDir === ROOT || ROOT.startsWith(outDir + path.sep)) {
        throw new PackagingError(`refusing to write output to '${outDir}': it contains the repository`);
    }
    await rm(path.join(outDir, name), { recursive: true, force: true });
    await rm(path.join(outDir, `${name}-${productVersion}.package.json`), { force: true });
    await rm(path.join(outDir, `${name}-${productVersion}.gist.json`), { force: true });
    await rm(staging, { recursive: true, force: true });
    await mkdir(staging, { recursive: true });

    log(`validating source ${name}...`);
    await validateTree(`the source folder '${name}'`, sourceDir, name, path.join(staging, "source"));

    // Sorted so the artefact, its manifest and its digest do not depend on
    // readdir order, which differs between filesystems.
    const all = (await listFiles(sourceDir)).sort();
    const included = [];
    const excluded = [];
    for (const rel of all) {
        const rule = classify(rel);
        const bytes = (await stat(path.join(sourceDir, rel))).size;
        if (rule) excluded.push({ path: rel, rule: rule.id, bytes });
        else included.push(rel);
    }

    if (included.length === 0) {
        throw new PackagingError(`every file in '${name}' was excluded — refusing to publish an empty artefact`);
    }

    const folderDir = path.join(outDir, name);
    const entries = [];
    const gistFiles = [];
    let totalBytes = 0;

    for (const rel of included) {
        const buffer = await readFile(path.join(sourceDir, rel));
        // Copied byte-for-byte: C3 requires the folder to run exactly as
        // committed, so any normalisation here would be a behaviour change
        // disguised as tidying.
        const destination = path.join(folderDir, rel);
        await mkdir(path.dirname(destination), { recursive: true });
        await writeFile(destination, buffer);
        await utimes(destination, epoch, epoch);

        if (gist) gistFiles.push({ path: rel, content: decodeUtf8(rel, buffer) });

        entries.push({ path: rel, bytes: buffer.byteLength, sha256: sha256(buffer) });
        totalBytes += buffer.byteLength;
    }

    log(`validating artefact ${name}...`);
    await validateTree(`the packaged artefact '${name}'`, folderDir, name, path.join(staging, "artefact"));

    await stampDirectories(folderDir, epoch);

    let gistPath = null;
    if (gist) {
        const payload = buildGistPayload({
            name,
            version: productVersion,
            description: manifest.description,
            files: gistFiles,
        });
        gistPath = path.join(outDir, `${name}-${productVersion}.gist.json`);
        await writeFile(gistPath, `${JSON.stringify(payload, null, 2)}\n`, "utf8");
        await utimes(gistPath, epoch, epoch);
    }

    const largest = entries.reduce((a, b) => (b.bytes > a.bytes ? b : a));
    const digest = `sha256:${sha256(entries.map((e) => `${e.sha256}  ${e.path}\n`).join(""))}`;

    const report = {
        name,
        version: productVersion,
        manifestFormatVersion: manifest.version,
        digest,
        totalBytes,
        sourceDateEpoch: epoch,
        files: entries,
        excluded,
        budget: {
            maxFileBytes: MAX_FILE_BYTES,
            maxTotalBytes: MAX_TOTAL_BYTES,
            largestFile: { path: largest.path, bytes: largest.bytes },
            headroomBytes: MAX_TOTAL_BYTES - totalBytes,
        },
        outputs: {
            folder: toPosix(path.relative(outDir, folderDir)),
            gist: gistPath ? toPosix(path.relative(outDir, gistPath)) : null,
        },
    };

    const reportPath = path.join(outDir, `${name}-${productVersion}.package.json`);
    await writeFile(reportPath, `${JSON.stringify(report, null, 2)}\n`, "utf8");
    await utimes(reportPath, epoch, epoch);
    await rm(staging, { recursive: true, force: true });

    return { report, reportPath, folderDir, gistPath };
}

/** C4 refuses binaries, so a file that will not decode cannot go in the payload. */
function decodeUtf8(rel, buffer) {
    try {
        return new TextDecoder("utf-8", { fatal: true }).decode(buffer);
    } catch {
        throw new PackagingError(`'${rel}' is not valid UTF-8 and cannot be shared as a gist (C4)`);
    }
}

/** Directory mtimes matter to `zip`; stamp deepest-first so parents stay put. */
async function stampDirectories(dir, epoch) {
    for (const entry of await readdir(dir, { withFileTypes: true })) {
        if (entry.isDirectory()) await stampDirectories(path.join(dir, entry.name), epoch);
    }
    await utimes(dir, epoch, epoch);
}

function parseArgs(argv) {
    const options = { name: null, out: path.join(ROOT, "dist"), expectVersion: null, gist: true };
    for (let i = 0; i < argv.length; i++) {
        const arg = argv[i];
        if (arg === "--out") options.out = path.resolve(argv[++i]);
        else if (arg === "--expect-version") options.expectVersion = argv[++i];
        else if (arg === "--no-gist") options.gist = false;
        else if (arg.startsWith("--")) throw new PackagingError(`unknown option '${arg}'`);
        else options.name = arg;
    }
    return options;
}

async function main() {
    const options = parseArgs(process.argv.slice(2));
    const extensionsDir = path.join(ROOT, EXTENSIONS_REL);
    const available = (await readdir(extensionsDir, { withFileTypes: true }))
        .filter((e) => e.isDirectory())
        .map((e) => e.name);

    // C3 packages exactly one folder. Guessing which one, when there is a
    // choice, would publish the wrong extension under the right name.
    let name = options.name;
    if (!name) {
        if (available.length !== 1) {
            throw new PackagingError(
                `name required: ${available.length} extensions available (${available.join(", ") || "none"})`,
            );
        }
        name = available[0];
    }
    if (!available.includes(name)) {
        throw new PackagingError(`no extension '${name}' (found: ${available.join(", ") || "none"})`);
    }

    const { report, reportPath, folderDir, gistPath } = await packageExtension({
        sourceDir: path.join(extensionsDir, name),
        outDir: options.out,
        name,
        expectVersion: options.expectVersion,
        gist: options.gist,
        log: (m) => console.log(m),
    });

    const byRule = new Map();
    for (const item of report.excluded) {
        const entry = byRule.get(item.rule) ?? { count: 0, bytes: 0 };
        byRule.set(item.rule, { count: entry.count + 1, bytes: entry.bytes + item.bytes });
    }

    console.log(`\n${report.name} ${report.version} — ${report.files.length} files, ${kb(report.totalBytes)}`);
    console.log(`  digest      ${report.digest}`);
    console.log(`  folder      ${path.relative(ROOT, folderDir)}   (C3: what install_extension copies)`);
    if (gistPath) console.log(`  gist body   ${path.relative(ROOT, gistPath)}   (C4: POST to /gists as-is)`);
    console.log(`  manifest    ${path.relative(ROOT, reportPath)}`);

    if (byRule.size > 0) {
        console.log("\nExcluded, and why:");
        for (const [id, { count, bytes }] of byRule) {
            const rule = EXCLUSIONS.find((r) => r.id === id);
            console.log(`  ${id} — ${count} file(s), ${kb(bytes)}`);
            console.log(`      ${rule.why.replace(/\s+/g, " ")}`);
        }
    }

    // The vendored pdf.js worker spends this headroom. Printing it every time is
    // the point: it should shrink visibly, not be discovered empty. Both
    // limits bind every install path, not just gist sharing.
    const { budget } = report;
    console.log("\nC4 budget (enforced by install_extension on gist and repo-folder installs alike):");
    console.log(
        `  largest file  ${budget.largestFile.bytes} of ${budget.maxFileBytes} bytes  (${budget.largestFile.path})`,
    );
    console.log(
        `  total         ${report.totalBytes} of ${budget.maxTotalBytes} bytes  ` +
            `— ${kb(budget.headroomBytes)} headroom`,
    );
}

if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith("package-extension.mjs")) {
    try {
        await main();
    } catch (err) {
        console.error(err instanceof PackagingError ? `\npackaging failed: ${err.message}` : err);
        process.exit(1);
    }
}
