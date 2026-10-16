# Save a COPY of a document in the current Word format (out of Compatibility Mode: Office add-ins are not
# available for documents in Compatibility Mode, so the MY_SCRIPTS tab does not appear for them).
#   pwsh -File tests/tools/convert-docx.ps1 -Path <in.docx> -Out <out.docx>
param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("convert-" + [guid]::NewGuid() + [IO.Path]::GetExtension($Path))
Copy-Item -LiteralPath $Path -Destination $tmp
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($tmp)
  "compatibility mode before: $($doc.CompatibilityMode)"
  $doc.Convert()
  $doc.SaveAs2($Out, 16)   # wdFormatDocumentDefault
  "compatibility mode after: $($doc.CompatibilityMode) -> $Out"
  $doc.Close(0)
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
