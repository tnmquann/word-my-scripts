# Render a Word-readable file (.docx or Flat OPC .xml) to PDF through Word COM (read-only, private copy).
#   pwsh -File tests/tools/render-pdf.ps1 -Path <in> -Out <out.pdf>
param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("render-" + [guid]::NewGuid() + [IO.Path]::GetExtension($Path))
Copy-Item -LiteralPath $Path -Destination $tmp
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($tmp, $false, $true)
  $doc.ExportAsFixedFormat($Out, 17)   # wdExportFormatPDF
  "pages: $($doc.ComputeStatistics(2))  -> $Out"
  $doc.Close(0)
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
