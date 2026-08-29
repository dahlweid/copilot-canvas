// Git must hand back the vendored bytes it was given.
//
// `core.autocrlf=true` is the default on a Windows install, and under it git
// rewrites every LF to CRLF on checkout. The vendored worker is minified but not
// newline-free: measured on this tree, `pdf.worker.min.mjs.part0` carries 28 LFs
// and `pdf.min.mjs` 28, so a fresh Windows clone would receive 600,028 and
// 454,697 bytes against a manifest that records 600,000 and 454,669. The
// reassembled worker then fails its sha256 in `joinWorker` and the canvas
// renders nothing.
//
// This suite exists because nothing else can see that. CI is Linux, where the
// conversion does not happen, so every other test in this repo stays green while
// the artifact is broken for the people most likely to be running Word. The
// guard is a `-text` attribute in the repo root `.gitattributes`; what follows
// asks git itself whether that guard is in force, rather than reading the file
// and trusting our own parse of it.
//
// Asking git rather than reading the file is not fastidiousness. Attributes are
// resolved from the index when the working-tree `.gitattributes` is missing, so
// the guard applies during a checkout that cannot see the file -- a fact a parse
// of the working tree would get wrong in both directions.
//
// Office-free.

import { test } from "node:test";
import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { readdir, stat } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { VENDOR_DIR } from "../../src/vendor-assets.mjs";

const execFileAsync = promisify(execFile);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const REPO = path.resolve(HERE, "..", "..", "..", "..", "..");

async function git(...args) {
    const { stdout } = await execFileAsync("git", args, { cwd: REPO, maxBuffer: 8 * 1024 * 1024 });
    return stdout;
}

// Skip rather than fail where there is no git to ask -- an installed extension
// folder is a plain directory, and this property is about the repository.
let available = null;
async function gitAvailable() {
    if (available !== null) return available;
    try {
        await git("rev-parse", "--git-dir");
        available = true;
    } catch {
        available = false;
    }
    return available;
}

// The list is read off disk, not restated here: a vendored file added later is
// covered the moment it lands, which is the case a hardcoded list would miss.
async function vendoredFiles() {
    const names = await readdir(VENDOR_DIR);
    const files = [];
    for (const name of names) {
        const full = path.join(VENDOR_DIR, name);
        if ((await stat(full)).isFile()) files.push(path.relative(REPO, full).split(path.sep).join("/"));
    }
    return files.sort();
}

test("every vendored file has line-ending conversion disabled", async (t) => {
    if (!(await gitAvailable())) return t.skip("not a git checkout");

    const files = await vendoredFiles();
    assert.ok(files.length > 0, "expected vendored files to enumerate");

    // `check-attr` answers per path with `unset` only when something actively
    // sets `-text`; an unmentioned path answers `unspecified`, which is the
    // state that lets the conversion happen.
    const stdout = await git("check-attr", "text", "--", ...files);
    const answers = new Map(
        stdout
            .split("\n")
            .filter(Boolean)
            .map((line) => {
                const marker = ": text: ";
                const at = line.lastIndexOf(marker);
                return [line.slice(0, at), line.slice(at + marker.length).trim()];
            }),
    );

    for (const file of files) {
        assert.equal(answers.get(file), "unset", `${file} would be line-ending converted on checkout`);
    }
});

test("checked-out bytes match the bytes git stores", async (t) => {
    if (!(await gitAvailable())) return t.skip("not a git checkout");

    // The attribute is the mechanism; this is the outcome. Comparing the size of
    // the blob git holds against the size on disk catches any other route to the
    // same corruption -- a smudge filter, a future `working-tree-encoding` -- and
    // fails on a machine where the attribute is right but the file was written
    // before it applied.
    const files = await vendoredFiles();
    const tracked = new Set(
        (await git("ls-files", "--", path.relative(REPO, VENDOR_DIR).split(path.sep).join("/")))
            .split("\n")
            .filter(Boolean),
    );

    let compared = 0;
    for (const file of files) {
        if (!tracked.has(file)) continue;
        const stored = Number((await git("cat-file", "-s", `:${file}`)).trim());
        const onDisk = (await stat(path.join(REPO, file))).size;
        assert.equal(onDisk, stored, `${file} differs from the bytes git stores`);
        compared += 1;
    }
    assert.ok(compared > 0, "expected at least one tracked vendored file to compare");
});
