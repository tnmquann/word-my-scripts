<#
.SYNOPSIS
  Removes the MY_SCRIPTS Word add-in registration for the current user.

.DESCRIPTION
  Add-in ID is taken from (in order): -Id, manifest.xml next to this script,
  addin.config.json in the repo (use -Dev for the dev build, otherwise prod),
  or the prod ID baked in at build time (hosted copy: irm <baseUrl>uninstall.ps1 | iex).
  -All removes every manifest registered from %LOCALAPPDATA%\MY_SCRIPTS.
#>
[CmdletBinding()]
param(
  [string]$Id,
  [switch]$Dev,
  [switch]$All,
  [string]$DefaultId = '{{ID}}'
)

$ErrorActionPreference = 'Stop'
$InstallRoot = Join-Path $env:LOCALAPPDATA 'MY_SCRIPTS\manifests'
$DevKey = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'

$ids = @()
if ($All) {
  if (Test-Path $InstallRoot) { $ids = Get-ChildItem $InstallRoot -Filter *.xml | ForEach-Object { $_.BaseName } }
} elseif ($Id) {
  $ids = @($Id)
} else {
  $local = if ($PSScriptRoot) { Join-Path $PSScriptRoot 'manifest.xml' } else { $null }
  $config = if ($PSScriptRoot) { Join-Path $PSScriptRoot '..\addin.config.json' } else { $null }
  if (-not $Dev -and $local -and (Test-Path $local)) {
    [xml]$xml = Get-Content -LiteralPath $local -Raw -Encoding UTF8
    $ids = @($xml.OfficeApp.Id)
  } elseif ($config -and (Test-Path $config)) {
    $cfg = Get-Content -LiteralPath $config -Raw -Encoding UTF8 | ConvertFrom-Json
    $ids = @($(if ($Dev) { $cfg.dev.id } else { $cfg.prod.id }))
  }
}
if (-not $ids -and $DefaultId -notlike '*{{*') { $ids = @($DefaultId) }
if (-not $ids) { throw 'Could not determine the add-in ID. Pass -Id <guid> or -All.' }

foreach ($i in $ids) {
  if (Get-ItemProperty -Path $DevKey -Name $i -ErrorAction SilentlyContinue) {
    Remove-ItemProperty -Path $DevKey -Name $i
    Write-Host "==> Unregistered $i" -ForegroundColor Cyan
  } else {
    Write-Host "==> $i was not registered" -ForegroundColor DarkGray
  }
  $file = Join-Path $InstallRoot "$i.xml"
  if (Test-Path $file) { Remove-Item -LiteralPath $file -Force }
}
# Auto-updater (only when no MY_SCRIPTS manifest remains)
if (-not (Test-Path $InstallRoot) -or -not (Get-ChildItem $InstallRoot -Filter *.xml -ErrorAction SilentlyContinue)) {
  Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'MY_SCRIPTS Updater' -ErrorAction SilentlyContinue
  Remove-Item -Path 'HKCU:\Software\Classes\myscripts' -Recurse -Force -ErrorAction SilentlyContinue
  Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA 'MY_SCRIPTS') -Recurse -Force -ErrorAction SilentlyContinue
  Write-Host '==> Removed auto-updater' -ForegroundColor Cyan
}
Write-Host 'Done. Restart Word to remove the MY_SCRIPTS tab.' -ForegroundColor Green
