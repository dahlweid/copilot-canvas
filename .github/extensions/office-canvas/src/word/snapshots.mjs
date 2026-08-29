// On-disk snapshots: the mechanism behind "revert the last edit".
//
// Word's own undo stack is not usable here. ADR 0005 has the host open, edit,
// save and *close* the document within a single operation, and Word discards
// the in-process undo history when the document closes. By the time the agent
// could ask to undo, there is nothing left to undo. So the snapshot is a copy
// of the file's bytes, taken immediately before the operation, and revert is a
// copy back.
//
// One snapshot per operation, not one per session, so a mistake three edits
// ago is still recoverable. Revert *pops* the newest: it restores those bytes
// and then removes that snapshot, so repeated reverts walk backwards through
// history instead of ping-ponging between the last two states.
//
// The snapshot directory is derived from the document path alone and is
// therefore stable across calls, but each snapshot filename carries a UTC
// timestamp, the operation and the revision token it replaces, so two
// concurrent operations write different files rather than fighting over one.

import { createHash } from "node:crypto";
import { copyFile, mkdir, readdir, readFile, rename, rm, stat, writeFile } from "node:fs/promises";
import path from "node:path";

/** How many snapshots to keep per document before the oldest are discarded. */
export const MAX_SNAPSHOTS_PER_DOCUMENT = 20;

const SNAPSHOT_EXTENSION = ".snapshot";
const MANIFEST_EXTENSION = ".json";

/**
 * Monotonic within the process, so two snapshots taken in the same millisecond
 * still have a defined order.
 *
 * The timestamp alone is not enough. An edit takes a few hundred milliseconds
 * but the snapshot is taken at the very start of it, so two operations queued
 * together are snapshotted within the same millisecond routinely rather than
 * exceptionally. Without this the names tie-break on the operation name and
 * then a random nonce, which is an ordering unrelated to time -- so
 * `revertToLatest`, which takes the lexically greatest name, could restore the
 * *older* of two snapshots and report that it had undone the newer.
 */
let sequence = 0;
const nextSequence = () => ++sequence;

/** Test seam: makes ordering assertions independent of how many ran before. */
export function __resetSequenceForTests(value = 0) {
    sequence = value;
}

export class SnapshotError extends Error {
    constructor(code, message) {
        super(message);
        this.name = "SnapshotError";
        this.code = code;
    }
}

const shortHash = (value) => createHash("sha256").update(value).digest("hex").slice(0, 16);

/**
 * Snapshots live beside each other per document. Keyed on the lowercased path
 * because Windows paths are case-insensitive, so `C:\A.docx` and `c:\a.docx`
 * are the same document and must share a history.
 */
export function snapshotDirFor(root, docPath) {
    return path.join(root, "snapshots", shortHash(docPath.toLowerCase()));
}

/** `2026-08-28T18:30:12.123Z` -> `20260828T183012123Z`, which sorts lexically. */
export function compactStamp(date = new Date()) {
    return date.toISOString().replace(/[-:]/g, "").replace(".", "");
}

/**
 * Names are parseable, and sort oldest-first as plain strings, so listing a
 * history never needs to stat every file.
 *
 * The sequence sits immediately after the timestamp and is zero-padded to a
 * fixed width, so lexical order and numeric order agree and the tie-break at
 * equal timestamps is time of creation rather than operation name.
 */
export function snapshotName({ takenAt = new Date(), op, token, nonce = "", seq = nextSequence() }) {
    const stamp = typeof takenAt === "string" ? takenAt : compactStamp(takenAt);
    const tokenPart = String(token ?? "none")
        .replace(/^sha256:/, "")
        .slice(0, 16)
        .replace(/[^0-9a-z]/gi, "") || "none";
    const opPart = String(op ?? "edit").replace(/[^a-z_]/gi, "") || "edit";
    const noncePart = nonce ? `-${String(nonce).replace(/[^0-9a-z]/gi, "")}` : "";
    const seqPart = String(Math.max(0, Number(seq) || 0)).padStart(9, "0");
    return `${stamp}-${seqPart}-${opPart}-${tokenPart}${noncePart}${SNAPSHOT_EXTENSION}`;
}

export function parseSnapshotName(name) {
    if (!name.endsWith(SNAPSHOT_EXTENSION)) return null;
    const stem = name.slice(0, -SNAPSHOT_EXTENSION.length);
    // The sequence is optional so a snapshot written before it existed still
    // parses, lists and restores rather than becoming invisible.
    const match = /^(\d{8}T\d{9}Z)(?:-(\d{9}))?-([a-z_]+)-([0-9a-z]+)(?:-([0-9a-z]+))?$/i.exec(stem);
    if (!match) return null;
    return {
        name,
        stamp: match[1],
        seq: match[2] === undefined ? null : Number(match[2]),
        op: match[3],
        token: match[4],
        nonce: match[5] ?? null,
    };
}

/**
 * Pure: given every snapshot name in a directory, decide which to keep.
 * Split out from the filesystem so retention is unit-testable without Word,
 * without disk, and without a clock.
 */
export function pruneSnapshots(names, keep = MAX_SNAPSHOTS_PER_DOCUMENT) {
    const parsed = names.map(parseSnapshotName).filter(Boolean);
    parsed.sort((a, b) => (a.name < b.name ? 1 : a.name > b.name ? -1 : 0)); // newest first
    return { keep: parsed.slice(0, keep), discard: parsed.slice(keep) };
}

const manifestPathFor = (snapshotPath) => `${snapshotPath}${MANIFEST_EXTENSION}`;

/**
 * Copies the document as it stands, before an operation touches it.
 *
 * Deliberately ordered: bytes first, manifest second. A crash between the two
 * leaves a restorable snapshot with no metadata, which is recoverable. The
 * reverse order would leave a manifest promising bytes that do not exist.
 */
export async function takeSnapshot({ root, docPath, op, token, description = null, nonce = "", takenAt = new Date() }) {
    const dir = snapshotDirFor(root, docPath);
    await mkdir(dir, { recursive: true });

    const name = snapshotName({ takenAt, op, token, nonce });
    const file = path.join(dir, name);

    await copyFile(docPath, file);

    const info = await stat(file);
    const manifest = {
        name,
        documentPath: docPath,
        documentName: path.basename(docPath),
        op,
        description,
        revisionToken: token ?? null,
        takenAt: takenAt.toISOString(),
        bytes: info.size,
    };
    await writeFile(manifestPathFor(file), JSON.stringify(manifest, null, 2), "utf8");

    await applyRetention(dir);
    return { ...manifest, file };
}

async function applyRetention(dir) {
    const names = await readdir(dir).catch(() => []);
    const { discard } = pruneSnapshots(names);
    for (const entry of discard) {
        const file = path.join(dir, entry.name);
        await rm(file, { force: true }).catch(() => {});
        await rm(manifestPathFor(file), { force: true }).catch(() => {});
    }
}

/** Newest first. Missing or unreadable manifests degrade to name-derived data. */
export async function listSnapshots({ root, docPath }) {
    const dir = snapshotDirFor(root, docPath);
    const names = await readdir(dir).catch(() => []);
    const { keep } = pruneSnapshots(names, Number.MAX_SAFE_INTEGER);

    const out = [];
    for (const entry of keep) {
        const file = path.join(dir, entry.name);
        let manifest = null;
        try {
            manifest = JSON.parse(await readFile(manifestPathFor(file), "utf8"));
        } catch {
            manifest = { name: entry.name, op: entry.op, revisionToken: null, takenAt: null };
        }
        out.push({ ...manifest, name: entry.name, file });
    }
    return out;
}

export async function latestSnapshot({ root, docPath }) {
    const [newest] = await listSnapshots({ root, docPath });
    return newest ?? null;
}

/**
 * Restores the newest snapshot over the original and discards it, so a second
 * revert reaches the state before that.
 *
 * Written to a sibling temp and renamed over the original, rather than copied
 * straight onto it. `copyFile` truncates the destination and then fills it, so
 * a failure part-way through leaves the user's document truncated; `rename`
 * within a volume is atomic, so the document is either the old bytes or the
 * restored ones and never a prefix of either. The temp is a sibling so that the
 * rename stays on one volume -- across volumes it degrades to copy-and-delete
 * and the atomicity is lost.
 *
 * Alternate data streams -- notably `Zone.Identifier`, the mark of the web --
 * survive this, but not for the reason an earlier version of this comment gave.
 * They do not survive because the destination is written in place; measured,
 * `copyFile` from an unmarked source over a marked destination *drops* the
 * destination's mark. They survive because the stream travels with the
 * *source*: `takeSnapshot` copies the original, so the snapshot carries the
 * mark, and the temp copied from the snapshot carries it too. Verified both
 * ways round, including that `rename` preserves the source's streams. Losing it
 * would silently strip a security marker from the user's file, which is the
 * outcome ADR 0007 refuses elsewhere.
 */
export async function revertToLatest({ root, docPath }) {
    const newest = await latestSnapshot({ root, docPath });
    if (!newest) {
        throw new SnapshotError(
            "no_snapshot",
            `No snapshot exists for ${path.basename(docPath)}. Snapshots are taken when edit_document changes a document; there is nothing to revert to.`,
        );
    }

    const staging = `${docPath}.revert-${process.pid.toString(36)}${(nextSequence() % 1000).toString(36)}.tmp`;
    // The two steps fail for unrelated reasons and want different remediation,
    // and only the throw site knows which one ran. Measured on this machine:
    // reading a held snapshot gives EBUSY (copyfile), while replacing the
    // document gives EPERM (rename) for *every* cause -- a holder, a read-only
    // attribute, a denying ACL alike. So the errno alone cannot say which side
    // failed, let alone why, and `step` is recorded rather than reconstructed.
    try {
        await copyFile(newest.file, staging);
    } catch (err) {
        await rm(staging, { force: true }).catch(() => {});
        err.step = "copy";
        throw err;
    }
    try {
        await rename(staging, docPath);
    } catch (err) {
        await rm(staging, { force: true }).catch(() => {});
        err.step = "replace";
        throw err;
    }

    await rm(newest.file, { force: true }).catch(() => {});
    await rm(manifestPathFor(newest.file), { force: true }).catch(() => {});

    const remaining = await listSnapshots({ root, docPath });
    return { restored: newest, remaining: remaining.length };
}
