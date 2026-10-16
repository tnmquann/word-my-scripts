# Run a macro of an old VBA template on a COPY of a document and save the result (.docx + .pdf).
# Used to capture the reference output of the original tools (e.g. BigKey bkAutoTab) for comparison.
# Run with PowerShell 7:
#   pwsh -File tests/tools/run-vba-macro.ps1 -Template <x.dotm> -Macro bkAutoTab -Path <doc.docx> -OutDir <dir>
# Word started via COM does not load Startup templates, so -Template is loaded explicitly.
# A MsgBox at the end of the macro is dismissed by pressing Enter on the Word window.
param(
  [Parameter(Mandatory)][string]$Template,
  [Parameter(Mandatory)][string]$Macro,
  [Parameter(Mandatory)][string]$Path,
  [Parameter(Mandatory)][string]$OutDir,
  [string]$SelectAll,          # non-empty: select the whole document first (for "selection" macros)
  [int]$TimeoutSec = 90
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$name = [IO.Path]::GetFileNameWithoutExtension($Path)
$copy = Join-Path $OutDir ("$name.input" + [IO.Path]::GetExtension($Path))
Copy-Item -LiteralPath $Path -Destination $copy -Force

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue).Id
$word = New-Object -ComObject Word.Application
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id }).Id
$word.Visible = $true
try {
  [void]$word.AddIns.Add($Template, $true)
  $doc = $word.Documents.Open($copy)
  if ($SelectAll) { $doc.Content.Select() } else { $word.Selection.HomeKey(6) | Out-Null }
  $job = Start-ThreadJob { param($w, $m) try { $w.Run($m); 'macro returned' } catch { 'macro error: ' + $_.Exception.Message } } -ArgumentList $word, $Macro
  $shell = New-Object -ComObject WScript.Shell
  $t0 = Get-Date
  while ($job.State -eq 'Running' -and ((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec) {
    Start-Sleep -Seconds 2
    # close a message box left by the macro (MsgBox / MsgBoxVN)
    if (((Get-Date) - $t0).TotalSeconds -gt 6 -and $shell.AppActivate('BigKey')) { $shell.SendKeys('~') }
  }
  "job: $($job.State) $(Receive-Job $job -ErrorAction SilentlyContinue)"
  $out = Join-Path $OutDir "$name.$Macro"
  $doc.SaveAs2("$out.docx", 16)
  $doc.SaveAs2("$out.pdf", 17)
  "saved $out.docx / .pdf"
} finally {
  try { $word.Quit(0) } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Start-Sleep -Seconds 2
  $mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}
