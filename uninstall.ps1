#Requires -Version 5.1
<#
.SYNOPSIS
    Uninstalls W365 CLI.

.DESCRIPTION
    Runs the installer's silent uninstaller (registered in Apps & Features), which also removes
    W365 CLI from your user PATH.

.EXAMPLE
    irm https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/uninstall.ps1 | iex
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "Uninstalling W365 CLI..." -ForegroundColor Cyan

# --- Stop running app -------------------------------------------------------
Get-Process -Name W365Cli -ErrorAction SilentlyContinue | ForEach-Object {
    Stop-Process -Id $_.Id -Force
}
Start-Sleep -Milliseconds 300

# --- Run Inno Setup uninstaller (Apps & Features entry) ---------------------
$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7BA17D73-6A35-4892-82CA-CA4AB563BAAD}_is1'
$uninstaller = (Get-ItemProperty -Path $uninstallKey -Name 'QuietUninstallString' -ErrorAction SilentlyContinue).QuietUninstallString
if (-not $uninstaller) {
    $uninstaller = (Get-ItemProperty -Path $uninstallKey -Name 'UninstallString' -ErrorAction SilentlyContinue).UninstallString
    if ($uninstaller) { $uninstaller += ' /VERYSILENT' }
}

if ($uninstaller) {
    Write-Host "Running registered uninstaller..." -ForegroundColor Yellow
    # The registered string already includes the full path + flags. Use cmd to honor it verbatim.
    Start-Process -FilePath 'cmd.exe' -ArgumentList '/c', $uninstaller -Wait
} else {
    Write-Host "No W365 CLI uninstaller registered - is it installed?" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "W365 CLI uninstalled." -ForegroundColor Green
Write-Host ""
