<#
.SYNOPSIS
  Installs (sideloads) the MY_SCRIPTS Word add-in for the current Windows user.

.DESCRIPTION
  Registers the add-in manifest under
    HKCU\Software\Microsoft\Office\16.0\WEF\Developer
  which Word reads at startup. No admin rights required. The web part of the add-in
  is served from GitHub Pages, so nothing else needs to be installed.

  Manifest source, in order of precedence:
    1. -ManifestPath <file>
    2. manifest.xml next to this script (release package)
    3. -ManifestUrl  (defaults to the published GitHub Pages manifest)

.EXAMPLE
  .\install.ps1
.EXAMPLE
  irm https://tnmquann.github.io/word-my-scripts/install.ps1 | iex
#>
[CmdletBinding()]
param(
  [string]$ManifestPath,
  [string]$ManifestUrl = '{{BASE_URL}}manifest.xml',
  [switch]$ClearCache,
  [switch]$NoUpdater
)

$ErrorActionPreference = 'Stop'
$InstallRoot = Join-Path $env:LOCALAPPDATA 'MY_SCRIPTS\manifests'
$DevKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

function Write-Step($msg) { Write-Host "==> $msg" -ForegroundColor Cyan }

# ---- 1. Resolve manifest --------------------------------------------------
$here = if ($PSScriptRoot) { $PSScriptRoot } else { $null }
if (-not $ManifestPath -and $here -and (Test-Path (Join-Path $here 'manifest.xml'))) {
  $ManifestPath = Join-Path $here 'manifest.xml'
}

$tmp = $null
if (-not $ManifestPath) {
  if ($ManifestUrl -like '*{{*') { throw 'No manifest found. Pass -ManifestPath or -ManifestUrl.' }
  Write-Step "Downloading manifest: $ManifestUrl"
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $tmp = Join-Path $env:TEMP ("my_scripts_manifest_{0}.xml" -f [guid]::NewGuid())
  Invoke-WebRequest -Uri $ManifestUrl -OutFile $tmp -UseBasicParsing
  $ManifestPath = $tmp
}
$ManifestPath = (Resolve-Path $ManifestPath).Path

[xml]$xml = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8
$id = $xml.OfficeApp.Id
$name = $xml.OfficeApp.DisplayName.DefaultValue
$version = $xml.OfficeApp.Version
if (-not $id) { throw "Invalid manifest (no <Id>): $ManifestPath" }
Write-Step "Add-in: $name  v$version  ($id)"

# ---- 2. Copy manifest to a stable per-user location ----------------------
New-Item -ItemType Directory -Force -Path $InstallRoot | Out-Null
$target = Join-Path $InstallRoot "$id.xml"
Copy-Item -LiteralPath $ManifestPath -Destination $target -Force
if ($tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
Write-Step "Manifest copied to $target"

# ---- 3. Register with Word -----------------------------------------------
if (-not (Test-Path $DevKey)) { New-Item -Path $DevKey -Force | Out-Null }
New-ItemProperty -Path $DevKey -Name $id -Value $target -PropertyType String -Force | Out-Null
Write-Step "Registered in $DevKey"

# ---- 4. Auto-update: updater.ps1 + logon check + myscripts: protocol -------
if (-not $NoUpdater) {
  $appRoot = Join-Path $env:LOCALAPPDATA 'MY_SCRIPTS'
  $updater = Join-Path $appRoot 'updater.ps1'
  $localUpdater = if ($here) { Join-Path $here 'updater.ps1' } else { $null }
  try {
    if ($localUpdater -and (Test-Path $localUpdater) -and -not ((Get-Content -LiteralPath $localUpdater -Raw) -like ("*'{{" + "BASE_URL}}'*"))) {
      Copy-Item -LiteralPath $localUpdater -Destination $updater -Force
    } else {
      $base = $ManifestUrl -replace 'manifest\.xml$', ''
      Invoke-WebRequest -Uri ($base + 'updater.ps1') -OutFile $updater -UseBasicParsing
    }
    $ps = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $cmd = "`"$ps`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$updater`""
    # Check at Windows logon
    New-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'MY_SCRIPTS Updater' `
      -Value "$cmd -AtLogon" -PropertyType String -Force | Out-Null
    # myscripts:update protocol (used by the add-in update button)
    $proto = 'HKCU:\Software\Classes\myscripts'
    New-Item -Path "$proto\shell\open\command" -Force | Out-Null
    Set-ItemProperty -Path $proto -Name '(default)' -Value 'URL:MY_SCRIPTS updater'
    New-ItemProperty -Path $proto -Name 'URL Protocol' -Value '' -PropertyType String -Force | Out-Null
    Set-ItemProperty -Path "$proto\shell\open\command" -Name '(default)' -Value "$cmd `"%1`""
    Write-Step "Auto-update enabled (logon check + myscripts:update)"
  } catch {
    Write-Warning "Could not set up the auto-updater: $($_.Exception.Message)"
  }
}

# ---- 5. Optional: clear Office add-in web cache ---------------------------
$wordRunning = [bool](Get-Process WINWORD -ErrorAction SilentlyContinue)
if ($ClearCache) {
  if ($wordRunning) {
    Write-Warning 'Word is running; close it and re-run with -ClearCache to clear the add-in cache.'
  } else {
    $cache = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
    if (Test-Path $cache) {
      Get-ChildItem -LiteralPath $cache -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
      Write-Step "Cleared Office add-in cache: $cache"
    }
  }
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
if ($wordRunning) {
  Write-Host 'Word is currently running: close ALL Word windows and open Word again to load the MY_SCRIPTS tab.' -ForegroundColor Yellow
} else {
  Write-Host 'Open Word: the MY_SCRIPTS tab will appear on the ribbon.'
}
