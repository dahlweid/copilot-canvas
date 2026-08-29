// Calls an export of a repo-level tool from a test that lives inside the
// extension folder.
//
// The extension must run exactly as committed from one copied folder (C3), and
// `tools/` is outside it -- `validate-extensions.mjs` rejects any import that
// reaches up there, including from `test/`. `packaging.test.mjs` and
// `validator.test.mjs` therefore drive their tools as subprocesses, and this is
// the same move for a tool whose *exports* are what needs testing rather than
// its command line.
//
// ## Arguments and results go through files, not argv or stdout
//
// The thing under test is a splitter for a 1.26 MB buffer. Passing that on the
// command line fails with `ENAMETOOLONG` -- measured here, not guessed -- and a
// result of that size on stdout is at the mercy of pipe buffering. Files have
// neither limit.
//
// ## Buffers are base64 in both directions
//
// `Buffer` has its own `toJSON`, so a buffer put through `JSON.stringify`
// without a replacer arrives as `{type: "Buffer", data: [...]}`: an object with
// no `length`, which a byte splitter reads as an empty input and answers with an
// empty array. That produced a plausible-looking `[]` in the first version of
// this helper, so both directions are encoded explicitly.

import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { pathToFileURL } from "node:url";

const execFileAsync = promisify(execFile);

function b64Replace(key, value) {
    const raw = this[key];
    if (Buffer.isBuffer(raw) || raw instanceof Uint8Array) {
        return { __b64: Buffer.from(raw).toString("base64") };
    }
    return value;
}

const b64Revive = (_key, value) =>
    value && typeof value === "object" && typeof value.__b64 === "string"
        ? Buffer.from(value.__b64, "base64")
        : value;

const DRIVER = `
import { readFile, writeFile } from "node:fs/promises";
const [, , toolUrl, kind, name, argsFile, outFile] = process.argv;
const revive = (_k, v) =>
    v && typeof v === "object" && typeof v.__b64 === "string" ? Buffer.from(v.__b64, "base64") : v;
function replace(k, v) {
    const raw = this[k];
    if (Buffer.isBuffer(raw) || raw instanceof Uint8Array) return { __b64: Buffer.from(raw).toString("base64") };
    return v;
}
const mod = await import(toolUrl);
if (!(name in mod)) {
    process.stderr.write("no export named " + name + " in " + toolUrl);
    process.exit(2);
}
const args = JSON.parse(await readFile(argsFile, "utf8"), revive);
const result = kind === "value" ? mod[name] : await mod[name](...args);
await writeFile(outFile, JSON.stringify({ result }, replace));
`;

let workDir = null;

async function workspace() {
    if (!workDir) {
        workDir = await mkdtemp(path.join(tmpdir(), "tool-bridge-"));
        await writeFile(path.join(workDir, "driver.mjs"), DRIVER);
    }
    return workDir;
}

/** Drops the driver and its scratch files. Harmless if never set up. */
export async function cleanupBridge() {
    if (workDir) await rm(workDir, { recursive: true, force: true });
    workDir = null;
}

let callId = 0;

async function run(toolPath, kind, name, args) {
    const dir = await workspace();
    const id = ++callId;
    const argsFile = path.join(dir, `args-${id}.json`);
    const outFile = path.join(dir, `out-${id}.json`);
    await writeFile(argsFile, JSON.stringify(args ?? [], b64Replace));

    // Every value is a discrete argv element and `execFile` does not go through
    // a shell, so no path is ever handed to `cmd` or PowerShell to reparse. An
    // interpolated command string here would break on an apostrophe or an `&` in
    // the repo path -- a bug this repo has already found twice.
    await execFileAsync(process.execPath, [
        path.join(dir, "driver.mjs"),
        pathToFileURL(toolPath).href,
        kind,
        name,
        argsFile,
        outFile,
    ]);
    return JSON.parse(await readFile(outFile, "utf8"), b64Revive).result;
}

/** Calls `export function name(...args)` in the tool at `toolPath`. */
export const callTool = (toolPath, name, args = []) => run(toolPath, "call", name, args);

/** Reads `export const name` from the tool at `toolPath`. */
export const readToolValue = (toolPath, name) => run(toolPath, "value", name);
