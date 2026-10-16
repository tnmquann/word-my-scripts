# Smoke test for installer/MY_SCRIPTS_Desktop.dotm (companion MathType tools) through Word COM.
# Needs Windows + Word + MathType. Run with PowerShell 7:  pwsh -File tests/tools/companion-smoke.ps1
# Word started via COM does not load Startup templates, so both templates are loaded explicitly.
# Macros that end with a MsgBox (LoiEquation, empty-selection prompts) are not run: MsgBox blocks a hidden Word.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$dotm = Join-Path $root 'installer\MY_SCRIPTS_Desktop.dotm'
$mt = Get-ChildItem 'C:\Program Files\Microsoft Office\root\Office16\STARTUP', "$env:APPDATA\Microsoft\Word\STARTUP", 'C:\Program Files (x86)\MathType\Office Support\64' -Filter 'MathType Commands 2016.dotm' -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $mt) { 'SKIP companion-smoke: MathType Commands 2016.dotm not found'; exit 0 }
$fail = 0
function Check($ok, $name, $info = '') { if ($ok) { "PASS $name $info" } else { "FAIL $name $info"; $script:fail++ } }

[xml](Get-Content (Join-Path $root 'companion\customUI14.xml') -Raw -Encoding UTF8) | Out-Null
'PASS customUI14.xml is well-formed'

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue).Id
$word = New-Object -ComObject Word.Application
# this test's own Word (killed in finally if MathType's dialog blocks Quit; never the user's Word)
$mine = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id }).Id
$word.Visible = $true
try {
  [void]$word.AddIns.Add($mt.FullName, $true)
  [void]$word.AddIns.Add($dotm, $true)
  $doc = $word.Documents.Add()

  # MathType hoá: undelimited selection with a trailing space -> wrapped in $...$, inner spaces kept
  $word.Selection.TypeText('Cho ')
  $s = $word.Selection.Start
  $word.Selection.TypeText('\alpha x^2 + 1 ')
  $e = $word.Selection.End
  $word.Selection.TypeText('xong.')
  $doc.Range($s, $e).Select()
  $word.Run('MathTypeHoa')
  $shapes = @($doc.InlineShapes)
  Check ($shapes.Count -eq 1 -and $shapes[0].OLEFormat.ProgID -like 'Equation.DSMT*') 'MathType hoa: one MathType object' "($($shapes.Count))"
  Check ($doc.Content.Text -match '^Cho .* xong\.') 'MathType hoa: trailing space kept outside the formula' ("text=" + $doc.Content.Text.Trim())

  # Lệch Pic-Math: shape lowered by -h/2 + 3.5 (h >= 14.4) or + 3; surrounding text untouched
  $doc.Range(0, 3).Font.Bold = $true
  $doc.Content.Select()
  $word.Run('LechPicMath')
  $h = $shapes[0].Height
  $want = if ($h -ge 14.4) { [math]::Round(-$h / 2 + 3.5) } else { [math]::Round(-$h / 2 + 3) }
  Check ($shapes[0].Range.Font.Position -eq $want) 'Lech Pic-Math: position from own height' "(h=$h pos=$($shapes[0].Range.Font.Position) want=$want)"
  Check ($doc.Range(0, 3).Font.Bold -eq -1 -and $doc.Range(0, 3).Font.Position -eq 0) 'Lech Pic-Math: text formatting kept'

  # Lệch MathType (last: it leaves MathType's Format Equations dialog open, Word is killed afterwards)
  $doc.Content.Select()
  [void](Start-ThreadJob { param($w) try { $w.Run('LechMathType') } catch {} } -ArgumentList $word)
  Start-Sleep -Seconds 4
  Check ($shapes[0].Range.Font.Position -eq 0) 'Lech MathType: position reset' "(pos=$($shapes[0].Range.Font.Position))"
} finally {
  try { $word.Quit(0) } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  Start-Sleep -Seconds 2   # Quit returns but the open MathType dialog keeps Word alive
  $mine | ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }
}
if ($fail) { exit 1 }
