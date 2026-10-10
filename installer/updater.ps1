<#
.SYNOPSIS
  MY_SCRIPTS updater: compares the installed manifest version with the published one and,
  if the user accepts, reinstalls the latest manifest.

.DESCRIPTION
  Installed by install.ps1 to %LOCALAPPDATA%\MY_SCRIPTS\updater.ps1 and triggered by:
    - Windows logon (HKCU Run key)          -> -AtLogon : silent unless an update exists
    - the add-in ("Cập nhật" button)        -> myscripts:update protocol
    - manual run                            -> powershell -File updater.ps1
  Published copy is saved as UTF-8 with BOM so Windows PowerShell 5.1 shows Vietnamese correctly.
#>
[CmdletBinding()]
param(
  [switch]$AtLogon,
  [Parameter(ValueFromRemainingArguments = $true)] [string[]]$Rest  # protocol URL ("myscripts:update")
)

$ErrorActionPreference = 'Stop'
$BaseUrl = '{{BASE_URL}}'
$AddinId = '{{ID}}'
$Title = 'MY_SCRIPTS – Cập nhật'

Add-Type -AssemblyName System.Windows.Forms
function Show-Box($text, $buttons = 'OK', $icon = 'Information') {
  [System.Windows.Forms.MessageBox]::Show($text, $Title, $buttons, $icon)
}

# Chỉ chạy 1 phiên cùng lúc
$mutex = New-Object System.Threading.Mutex($false, 'Local\MY_SCRIPTS_Updater')
if (-not $mutex.WaitOne(0)) { exit 0 }

try {
  if ($BaseUrl -like '*{{*') { throw 'updater.ps1 chưa được build (thiếu BASE_URL).' }
  if ($AtLogon) { Start-Sleep -Seconds 20 }  # chờ mạng sẵn sàng sau khi đăng nhập

  # ---- Phiên bản đang cài ----
  $manifest = Join-Path $env:LOCALAPPDATA "MY_SCRIPTS\manifests\$AddinId.xml"
  $installed = [version]'0.0.0.0'
  if (Test-Path $manifest) {
    [xml]$xml = Get-Content -LiteralPath $manifest -Raw -Encoding UTF8
    $installed = [version]$xml.OfficeApp.Version
  }

  # ---- Phiên bản mới nhất ----
  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  try {
    $info = Invoke-RestMethod -Uri ("{0}version.json?t={1}" -f $BaseUrl, [DateTime]::UtcNow.Ticks) -UseBasicParsing -TimeoutSec 15
  } catch {
    if (-not $AtLogon) { Show-Box "Không kiểm tra được bản cập nhật (mất mạng?).`n`n$($_.Exception.Message)" 'OK' 'Warning' | Out-Null }
    exit 1
  }
  $latest = [version]("{0}.0" -f $info.version)

  if ($latest -le $installed) {
    if (-not $AtLogon) { Show-Box "Bạn đang dùng bản mới nhất (v$($info.version))." | Out-Null }
    exit 0
  }

  # ---- Hỏi người dùng ----
  $cur = if ($installed.Major -eq 0) { 'chưa cài' } else { "v$($installed.Major).$($installed.Minor).$($installed.Build)" }
  $answer = Show-Box ("Có bản MY_SCRIPTS mới: v{0} (đang dùng: {1}).`n`nCài đặt ngay? Word cần được đóng để hoàn tất." -f $info.version, $cur) 'YesNo' 'Question'
  if ($answer -ne 'Yes') { exit 0 }

  # ---- Đợi đóng Word ----
  while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    $r = Show-Box "Hãy LƯU và ĐÓNG tất cả cửa sổ Word, sau đó bấm OK để tiếp tục cài đặt." 'OKCancel' 'Warning'
    if ($r -ne 'OK') { exit 0 }
  }

  # ---- Cài bản mới bằng install.ps1 đã publish ----
  $tmp = Join-Path $env:TEMP ("my_scripts_install_{0}.ps1" -f [guid]::NewGuid())
  Invoke-WebRequest -Uri ("{0}install.ps1?t={1}" -f $BaseUrl, [DateTime]::UtcNow.Ticks) -OutFile $tmp -UseBasicParsing
  try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $tmp -ClearCache | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "install.ps1 exit code $LASTEXITCODE" }
  } finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
  }
  Show-Box "Đã cập nhật MY_SCRIPTS lên v$($info.version).`nMở lại Word để sử dụng.`n`nNếu không thấy tab MY_SCRIPTS: Home > Add-ins > Advanced > DEVELOPER ADD-INS > MY_SCRIPTS > Add." | Out-Null
} catch {
  Show-Box "Cập nhật thất bại:`n`n$($_.Exception.Message)" 'OK' 'Error' | Out-Null
  exit 1
} finally {
  $mutex.ReleaseMutex()
}
