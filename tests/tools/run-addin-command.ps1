# Run a MY_SCRIPTS ribbon command for real (installed add-in) on a COPY of a document, via UI Automation.
# Word started through COM never loads web add-ins, so this starts Word normally and clicks the ribbon.
# Run with PowerShell 7 while no other Word window is open (the script refuses otherwise):
#   pwsh -File tests/tools/run-addin-command.ps1 -Path <doc.docx> -OutDir <dir> -Menu "Chuẩn hoá BigKey" -Item "Chuẩn hoá BigKey (cả đề)"
# Output: <OutDir>/<name>.addin.docx (the processed copy) and the text of the result popup on stdout.
param(
  [Parameter(Mandatory)][string]$Path,
  [Parameter(Mandatory)][string]$OutDir,
  [Parameter(Mandatory)][string]$Menu,
  [Parameter(Mandatory)][string]$Item,
  [string]$Tab = 'MY_SCRIPTS',
  [string]$Answer,             # button to click if the command asks a question (askUser popup), e.g. "Đánh số"
  [int]$WaitSec = 90
)
$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close every Word window first (this script drives the Word UI).' }
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$PC = [System.Windows.Automation.PropertyCondition]

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$name = [IO.Path]::GetFileNameWithoutExtension($Path)
$copy = Join-Path (Resolve-Path $OutDir) "$name.addin.docx"
Copy-Item -LiteralPath $Path -Destination $copy -Force

function Find-ByName($root, [string]$n, [int]$timeout = 30) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $timeout) {
    $e = $root.FindFirst($TS::Descendants, (New-Object $PC($AE::NameProperty, $n)))
    if ($e) { return $e }
    Start-Sleep -Milliseconds 500
  }
  throw "UI element not found: '$n'"
}
function Invoke-El($e) {
  foreach ($p in [System.Windows.Automation.InvokePattern]::Pattern, [System.Windows.Automation.SelectionItemPattern]::Pattern, [System.Windows.Automation.ExpandCollapsePattern]::Pattern) {
    $o = $null
    if ($e.TryGetCurrentPattern($p, [ref]$o)) {
      if ($o -is [System.Windows.Automation.InvokePattern]) { $o.Invoke() }
      elseif ($o -is [System.Windows.Automation.SelectionItemPattern]) { $o.Select() }
      else { $o.Expand() }
      return
    }
  }
  throw "Element '$($e.Current.Name)' cannot be invoked"
}

$proc = Start-Process -FilePath 'WINWORD.EXE' -ArgumentList '/q', "`"$copy`"" -PassThru
try {
  $t0 = Get-Date
  $win = $null
  while (-not $win -and ((Get-Date) - $t0).TotalSeconds -lt 60) {
    Start-Sleep -Seconds 1
    $win = $AE::RootElement.FindFirst($TS::Children, (New-Object $PC($AE::ClassNameProperty, 'OpusApp')))
  }
  if (-not $win) { throw 'Word window not found' }
  Start-Sleep -Seconds 4   # let the add-in register its ribbon
  Invoke-El (Find-ByName $win $Tab)
  Start-Sleep -Seconds 1
  Invoke-El (Find-ByName $win $Menu)
  Start-Sleep -Seconds 1
  Invoke-El (Find-ByName $AE::RootElement $Item 10)
  "clicked: $Tab > $Menu > $Item"

  # The result popup is an Office dialog (WebView). Wait for its close button "Đóng", then read its text.
  # A question popup (askUser) shows $Answer instead: click it once and keep waiting for the result.
  $dlg = $null; $t0 = Get-Date; $answered = $false
  while (-not $dlg -and ((Get-Date) - $t0).TotalSeconds -lt $WaitSec) {
    Start-Sleep -Seconds 2
    if ($Answer -and -not $answered) {
      $q = $AE::RootElement.FindFirst($TS::Descendants, (New-Object $PC($AE::NameProperty, $Answer)))
      if ($q) { Invoke-El $q; $answered = $true; "answered: $Answer"; continue }
    }
    $dlg = $AE::RootElement.FindFirst($TS::Descendants, (New-Object $PC($AE::NameProperty, 'Đóng')))
  }
  if ($dlg) {
    $walker = [System.Windows.Automation.TreeWalker]::RawViewWalker
    $box = $walker.GetParent($walker.GetParent($dlg))
    $texts = $box.FindAll($TS::Descendants, (New-Object $PC($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))
    "popup: " + (($texts | ForEach-Object { $_.Current.Name } | Where-Object { $_ }) -join ' | ')
    Invoke-El $dlg
  } else { 'popup: (not found within timeout)' }
  Start-Sleep -Seconds 2
  # Save the copy in place (Ctrl+S), then close Word
  $shell = New-Object -ComObject WScript.Shell
  [void]$shell.AppActivate($proc.Id)
  Start-Sleep -Milliseconds 500
  $shell.SendKeys('^s')
  Start-Sleep -Seconds 4
  "saved: $copy"
} finally {
  Start-Sleep -Seconds 1
  if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
}
