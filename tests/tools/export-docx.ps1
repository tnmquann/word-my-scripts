# Export a .docx READ-ONLY through Word COM into JSON for the real-document test harnesses.
# Run with PowerShell 7 (pwsh): Vietnamese text in this file needs UTF-8.
#   pwsh -File tests/tools/export-docx.ps1 -Path <docx> -Out <json>
# Output: paragraphs (text, list string, inTable, start, end), tables (start, end, values), and the
# whole body WordOpenXML (flat OPC) so harnesses can build per-range packages like Office JS getOoxml().
param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Out)
$ErrorActionPreference = 'Stop'
# Work on a private copy so the user's file (possibly open in Word) is never touched.
$tmp = Join-Path ([IO.Path]::GetTempPath()) ("export-" + [guid]::NewGuid() + [IO.Path]::GetExtension($Path))
Copy-Item -LiteralPath $Path -Destination $tmp
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  $doc = $word.Documents.Open($tmp, $false, $true)   # ConfirmConversions=false, ReadOnly=true
  $paras = New-Object System.Collections.Generic.List[object]
  $i = 0
  foreach ($p in $doc.Paragraphs) {
    $r = $p.Range
    # COM counts table end-of-row markers as paragraphs; Office JS (and OOXML w:p) does not.
    $c = $r.Duplicate; $c.Collapse(1)       # wdCollapseStart: Information(31) is only True on the collapsed start
    if ($c.Information(31)) { continue }    # wdAtEndOfRowMarker
    $paras.Add([pscustomobject]@{
      i = $i; text = ($r.Text -replace "[\r\a]$", ""); list = $r.ListFormat.ListString
      inTable = [bool]$r.Information(12); start = $r.Start; end = $r.End
    })
    $i++
  }
  $tables = foreach ($t in $doc.Tables) {
    $rows = @{}
    foreach ($c in $t.Range.Cells) {
      if (-not $rows.ContainsKey($c.RowIndex)) { $rows[$c.RowIndex] = @{} }
      $rows[$c.RowIndex][$c.ColumnIndex] = ($c.Range.Text -replace "[\r\a]+$", "")
    }
    $values = foreach ($ri in ($rows.Keys | Sort-Object)) { , @(foreach ($ci in ($rows[$ri].Keys | Sort-Object)) { $rows[$ri][$ci] }) }
    [pscustomobject]@{ start = $t.Range.Start; end = $t.Range.End; values = @($values) }
  }
  $bodyXml = $doc.Content.WordOpenXML
  $doc.Close(0)
  [pscustomobject]@{ paragraphs = $paras; tables = @($tables); bodyXml = $bodyXml } |
    ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $Out -Encoding UTF8
  "paragraphs: $($paras.Count)  tables: $(@($tables).Count)  bodyXml: $([math]::Round($bodyXml.Length/1KB)) KB"
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}
