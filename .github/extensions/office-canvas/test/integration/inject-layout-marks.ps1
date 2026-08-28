# Injects paragraphs carrying Word's layout marks into an existing .docx.
#
# These marks are the ones where the two text representations can disagree: the
# structure map reads them out of the markup, while the host reads them over COM
# as control characters. A mismatch makes the affected paragraph permanently
# uneditable, and reports it as "the file changed", which is false -- so this
# fixture exists to keep that class of bug caught.
#
# It patches a *real* Word-authored package rather than building one. A
# hand-built minimal .docx is accepted by the format but opens with zero
# paragraphs, because Word wants parts the minimal package omits. ZipArchive in
# Update mode leaves every other part byte for byte intact.
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'

Copy-Item -LiteralPath $Source -Destination $Out -Force
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

$zip = [IO.Compression.ZipFile]::Open($Out, [IO.Compression.ZipArchiveMode]::Update)
try {
    $entry = $zip.GetEntry('word/document.xml')
    if ($null -eq $entry) { throw 'no word/document.xml in the package' }
    $reader = New-Object IO.StreamReader($entry.Open())
    $xml = $reader.ReadToEnd()
    $reader.Close()

    # One paragraph per mark, each tagged so the test can find it by text.
    #   MARKbr    w:br            map "\n"   COM \u000B
    #   MARKnbh   w:noBreakHyphen map "-"    COM \u001E
    #   MARKtab   w:tab           map "\t"   COM \u0009
    #   MARKsoft  w:softHyphen    map ""     COM \u001F
    $inject = @'
<w:p><w:r><w:t>MARKbr abc</w:t><w:br/><w:t>def</w:t></w:r></w:p><w:p><w:r><w:t>MARKnbh e</w:t><w:noBreakHyphen/><w:t>mail</w:t></w:r></w:p><w:p><w:r><w:t>MARKtab a</w:t><w:tab/><w:t>b</w:t></w:r></w:p><w:p><w:r><w:t>MARKsoft soft</w:t><w:softHyphen/><w:t>hyphen</w:t></w:r></w:p>
'@

    $idx = $xml.IndexOf('<w:body>')
    if ($idx -lt 0) { throw 'no <w:body> in document.xml' }
    $at = $idx + '<w:body>'.Length
    $patched = $xml.Substring(0, $at) + $inject + $xml.Substring($at)

    $stream = $entry.Open()
    $stream.SetLength(0)
    # Set-Content -Encoding fails to bind on a German-locale PowerShell here;
    # UTF8Encoding($false) also guarantees no BOM, which the part must not have.
    $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
    $writer.Write($patched)
    $writer.Flush()
    $writer.Close()
} finally {
    $zip.Dispose()
}
