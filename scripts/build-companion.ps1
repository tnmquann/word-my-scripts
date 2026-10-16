# Build the companion VBA template installer/MY_SCRIPTS_Desktop.dotm from companion/*.bas + companion/customUI14.xml.
# Needs Windows + Word (COM) + "Trust access to the VBA project object model" (HKCU ...\Word\Security AccessVBOM=1).
# Run with PowerShell 7:  pwsh -File scripts/build-companion.ps1   (or: npm run companion)
# The .dotm is committed: CI (windows-latest) has no Word, so build.mjs / release.mjs only copy it.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$out = Join-Path $root 'installer\MY_SCRIPTS_Desktop.dotm'

# VBA source is UTF-8; the VBE stores code in the ANSI code page, so every string literal with
# non-ASCII text becomes "..." & ChrW(n) & "...". Comments are made ASCII (diacritics dropped).
function Convert-Literal([string]$lit) {
  $parts = New-Object System.Collections.Generic.List[string]
  $buf = ''
  foreach ($ch in $lit.ToCharArray()) {
    if ([int]$ch -lt 128) { $buf += $ch; continue }
    if ($buf) { $parts.Add('"' + $buf + '"'); $buf = '' }
    $parts.Add("ChrW($([int]$ch))")
  }
  if ($buf -or -not $parts.Count) { $parts.Add('"' + $buf + '"') }
  return ($parts -join ' & ')
}
function Remove-Diacritics([string]$s) {
  $n = $s.Replace([char]0x0111, 'd').Replace([char]0x0110, 'D').Normalize([Text.NormalizationForm]::FormD)
  return (-join ($n.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne 'NonSpacingMark' })) -replace '[^\x00-\x7f]', '?'
}
function Convert-VbaLine([string]$line) {
  $sb = New-Object System.Text.StringBuilder
  $i = 0
  while ($i -lt $line.Length) {
    $c = $line[$i]
    if ($c -eq "'") { [void]$sb.Append((Remove-Diacritics $line.Substring($i))); break }   # comment to end of line
    if ($c -eq '"') {
      $j = $i + 1; $lit = ''
      while ($j -lt $line.Length) {
        if ($line[$j] -eq '"') { if ($j + 1 -lt $line.Length -and $line[$j + 1] -eq '"') { $lit += '""'; $j += 2; continue } else { break } }
        $lit += $line[$j]; $j++
      }
      if ($lit -match '[^\x00-\x7f]') { [void]$sb.Append((Convert-Literal $lit)) } else { [void]$sb.Append('"' + $lit + '"') }
      $i = $j + 1; continue
    }
    [void]$sb.Append($c); $i++
  }
  return $sb.ToString()
}

$modules = Get-ChildItem (Join-Path $root 'companion') -Filter *.bas
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  # Check imageMso names used by the ribbon (an unknown name shows a blank button). GetImageMso cannot cross
  # the COM process boundary (IPictureDisp), so check the control label instead: unknown ids throw.
  $ui = Get-Content (Join-Path $root 'companion\customUI14.xml') -Raw -Encoding UTF8
  foreach ($m in [regex]::Matches($ui, 'imageMso="([^"]+)"')) {
    try { [void]$word.CommandBars.GetLabelMso($m.Groups[1].Value) } catch { throw "Unknown imageMso '$($m.Groups[1].Value)' in customUI14.xml" }
  }
  $doc = $word.Documents.Add()
  foreach ($f in $modules) {
    $lines = (Get-Content $f.FullName -Encoding UTF8) | Where-Object { $_ -notmatch '^Attribute VB_' }
    $code = ($lines | ForEach-Object { Convert-VbaLine $_ }) -join "`r`n"
    if ($code -match '[^\x00-\x7f]') { throw "$($f.Name): non-ASCII left after conversion" }
    $comp = $doc.VBProject.VBComponents.Add(1)   # vbext_ct_StdModule
    $comp.Name = $f.BaseName
    $comp.CodeModule.AddFromString($code)
  }
  $tmp = Join-Path ([IO.Path]::GetTempPath()) ("companion-" + [guid]::NewGuid() + '.dotm')
  $doc.SaveAs2($tmp, 15)   # wdFormatXMLTemplateMacroEnabled
  $doc.Close(0)
} finally {
  $word.Quit(0)
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}

# Inject the ribbon: customUI/customUI14.xml + relationship in _rels/.rels
Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::Open($tmp, 'Update')
try {
  $e = $zip.CreateEntry('customUI/customUI14.xml')
  $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false)))
  $w.Write($ui); $w.Close()
  $rels = $zip.GetEntry('_rels/.rels')
  $r = New-Object IO.StreamReader($rels.Open()); $xml = $r.ReadToEnd(); $r.Close()
  $xml = $xml.Replace('</Relationships>', '<Relationship Id="msCustomUI14" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/></Relationships>')
  $rels.Delete()
  $w = New-Object IO.StreamWriter($zip.CreateEntry('_rels/.rels').Open(), (New-Object Text.UTF8Encoding($false)))
  $w.Write($xml); $w.Close()
} finally { $zip.Dispose() }

Move-Item -LiteralPath $tmp -Destination $out -Force
"Built $out ($([math]::Round((Get-Item $out).Length / 1KB)) KB): $($modules.Name -join ', ') + customUI14.xml"
