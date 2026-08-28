// Which errno does Node actually report for each way a file can be unreadable
// on Windows, and which exception type does PowerShell's Copy-Item raise? The
// typed-error mapping asserts a *cause* -- "locked by another process" -- so the
// mapping has to be measured, not guessed.
//
// Both halves are here because the read path crosses both runtimes:
// `document-reader.mjs` branches on the errno, `word-host.ps1` branches on the
// exception type. A message-based discriminator is not an option: Word and
// Windows are German on this machine.
//
// The question this answers is whether `file_locked` and `permission_denied`
// are genuinely distinguishable, or whether two codes would claim a precision
// the platform does not give us. Result: distinguishable on the read path
// (EBUSY vs EPERM), and *not* distinguishable through a write handle, which is
// why `Test-FileWritable`'s `writable` field stays deliberately ambiguous.
//
// This is a probe, not a test: it mutates ACLs and spawns holder processes, so
// it is run by hand rather than in the suite. It is .mjs rather than .ps1
// because half of what it measures is Node's own errno translation.
//
// Run: node <this file>

import { writeFileSync, createReadStream, rmSync, mkdtempSync, chmodSync } from "node:fs";
import { execFileSync, spawn } from "node:child_process";
import { tmpdir } from "node:os";
import path from "node:path";

const dir = mkdtempSync(path.join(tmpdir(), "errno-probe-"));
const results = [];

const tokenRead = (p) =>
    new Promise((resolve) => {
        const stream = createReadStream(p);
        stream.on("error", (err) => resolve({ code: err.code, message: err.message }));
        stream.on("data", () => {});
        stream.on("end", () => resolve({ code: null, message: "read SUCCEEDED" }));
    });

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function withHolder(label, share) {
    const target = path.join(dir, `${label}.txt`);
    writeFileSync(target, "hello");
    // Path and share travel in the environment rather than interpolated into
    // the command string: a Windows profile may contain an apostrophe, which
    // would close the PowerShell string and quietly make the holder open
    // nothing -- turning a measurement into a fiction.
    const holder = spawn(
        "powershell.exe",
        [
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            "$fs=[IO.File]::Open($env:PROBE_PATH,'Open','ReadWrite',$env:PROBE_SHARE); Start-Sleep -Seconds 10; $fs.Close()",
        ],
        { env: { ...process.env, PROBE_PATH: target, PROBE_SHARE: share } },
    );
    await sleep(2000);
    const r = await tokenRead(target);
    holder.kill();
    await sleep(400);
    return r;
}

// 1. exclusive lock -- stricter than Word
results.push(["FileShare::None (exclusive)", await withHolder("exclusive", "None")]);

// 2. what Word itself takes
results.push(["FileShare::Read (Word's own lock)", await withHolder("wordlike", "Read")]);

// 3. read-only attribute
{
    const target = path.join(dir, "readonly.txt");
    writeFileSync(target, "hello");
    try {
        execFileSync("attrib", ["+R", target], { stdio: "ignore" });
        results.push(["read-only attribute", await tokenRead(target)]);
        execFileSync("attrib", ["-R", target], { stdio: "ignore" });
    } catch (err) {
        results.push(["read-only attribute", { code: "PROBE-FAILED", message: String(err.message).slice(0, 40) }]);
    }
}

// 4. ACL denying read. icacls rather than Get-Acl: the
// Microsoft.PowerShell.Security module does not load on this machine.
{
    const target = path.join(dir, "denied.txt");
    writeFileSync(target, "hello");
    const user = process.env.USERNAME ?? "";
    try {
        execFileSync("icacls", [target, "/deny", `${user}:(R)`], { stdio: "ignore" });
        results.push(["ACL deny read (current user)", await tokenRead(target)]);
    } catch (err) {
        results.push(["ACL deny read (current user)", { code: "PROBE-FAILED", message: String(err.message).slice(0, 40) }]);
    }
    try {
        execFileSync("icacls", [target, "/remove:d", user], { stdio: "ignore" });
    } catch {
        /* best effort */
    }
}

// 5. chmod 000, for comparison with the Linux runner
{
    const target = path.join(dir, "chmod.txt");
    writeFileSync(target, "hello");
    try {
        chmodSync(target, 0o000);
        results.push(["chmod 000", await tokenRead(target)]);
        chmodSync(target, 0o644);
    } catch (err) {
        results.push(["chmod 000", { code: "PROBE-FAILED", message: String(err.message).slice(0, 40) }]);
    }
}

console.log("\n  Node read handle -- what document-reader.mjs branches on\n");
console.log("  scenario                             errno          message");
console.log("  " + "-".repeat(84));
for (const [name, r] of results) {
    console.log(`  ${name.padEnd(36)} ${String(r.code ?? "(none)").padEnd(14)} ${r.message.slice(0, 58)}`);
}
console.log();

// 6. The PowerShell side of the same question. `word-host.ps1` copies the
// original with Copy-Item and branches on the exception *type*, so measure the
// type names rather than the messages.
{
    const script = `
$dir = Join-Path $env:TEMP ("errno-ps-" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$rows = @()
foreach ($case in @(@{n='FileShare::None (exclusive)';s='None'}, @{n="FileShare::Read (Word's own lock)";s='Read'})) {
    $src = Join-Path $dir ($case.s + '.txt')
    Set-Content -Path $src -Value 'hello'
    $fs = [IO.File]::Open($src, 'Open', 'ReadWrite', $case.s)
    try { Copy-Item $src (Join-Path $dir ($case.s + '.copy')) -ErrorAction Stop; $rows += ,@($case.n, 'copy SUCCEEDED') }
    catch { $rows += ,@($case.n, $_.Exception.GetType().FullName) }
    finally { $fs.Close() }
}
$src = Join-Path $dir 'denied.txt'
Set-Content -Path $src -Value 'hello'
icacls $src /deny "$($env:USERNAME):(R)" | Out-Null
try { Copy-Item $src (Join-Path $dir 'denied.copy') -ErrorAction Stop; $rows += ,@('ACL deny read', 'copy SUCCEEDED') }
catch { $rows += ,@('ACL deny read', $_.Exception.GetType().FullName) }
icacls $src /remove:d $env:USERNAME | Out-Null
$src = Join-Path $dir 'readonly.txt'
Set-Content -Path $src -Value 'hello'
attrib +R $src
try { Copy-Item $src (Join-Path $dir 'ro.copy') -ErrorAction Stop; $rows += ,@('read-only attribute', 'copy SUCCEEDED') }
catch { $rows += ,@('read-only attribute', $_.Exception.GetType().FullName) }
attrib -R $src
foreach ($r in $rows) { Write-Output ($r[0] + '|' + $r[1]) }
Remove-Item $dir -Recurse -Force -ErrorAction SilentlyContinue
`;
    console.log("  PowerShell Copy-Item -- what word-host.ps1 branches on\n");
    console.log("  scenario                             exception type");
    console.log("  " + "-".repeat(84));
    try {
        const out = execFileSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script], {
            encoding: "utf8",
        });
        for (const line of out.split(/\r?\n/).filter(Boolean)) {
            const [name, type] = line.split("|");
            console.log(`  ${name.padEnd(36)} ${type}`);
        }
    } catch (err) {
        console.log(`  PROBE FAILED: ${err.message}`);
    }
    console.log();
}

try {
    rmSync(dir, { recursive: true, force: true });
} catch {
    /* the denied file may resist removal */
}
