#Requires -Version 5.1
<#
.SYNOPSIS
    Installs W365 CLI.

.DESCRIPTION
    Downloads the latest W365 CLI installer from GitHub and runs it. The installer puts W365 CLI
    in "%LocalAppData%\Programs\W365CLI", registers a clean uninstaller in Apps & Features, and
    (by default) adds that folder to your user PATH so you can just type "w365cli" from any new
    terminal. No admin rights or UAC prompt are required — this is a per-user install.

.PARAMETER NoPath
    Install without adding W365 CLI to your PATH. You'll need to run it by full path or add it
    to PATH yourself later.

.EXAMPLE
    irm https://raw.githubusercontent.com/bwya77/W365-CLI-Native/main/install.ps1 | iex

.EXAMPLE
    .\install.ps1 -NoPath
#>
[CmdletBinding()]
param(
    [switch]$NoPath
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$repo = 'bwya77/W365-CLI-Native'

# --- Architecture detection ------------------------------------------------
$archEnv = $env:PROCESSOR_ARCHITECTURE
if ($archEnv -eq 'ARM64') {
    $arch = 'arm64'
} elseif ([Environment]::Is64BitOperatingSystem) {
    $arch = 'x64'
} else {
    throw "W365 CLI requires 64-bit Windows. Detected: $archEnv"
}

Write-Host ""
Write-Host "  W365 CLI" -ForegroundColor Cyan
Write-Host "  --------"
Write-Host "  Architecture : $arch"
Write-Host ""

# --- Resolve latest release asset -----------------------------------------
Write-Host "Looking up latest release..."
try {
    $release = Invoke-RestMethod `
        -Uri "https://api.github.com/repos/$repo/releases/latest" `
        -Headers @{ 'User-Agent' = 'W365CLI-Installer'; 'Accept' = 'application/vnd.github+json' } `
        -ErrorAction Stop
} catch {
    throw "Couldn't reach GitHub: $($_.Exception.Message)"
}

$asset = $release.assets | Where-Object {
    $_.name -like "W365CLISetup-*-win-$arch.exe"
} | Select-Object -First 1

if (-not $asset) {
    $available = ($release.assets | ForEach-Object { $_.name }) -join ', '
    throw "Release $($release.tag_name) doesn't contain a W365 CLI installer for $arch. Available: $available"
}

# --- Download -------------------------------------------------------------
$tempPath = Join-Path $env:TEMP $asset.name
Write-Host "Downloading $($release.tag_name) ($([math]::Round($asset.size / 1MB, 1)) MB)..."
Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tempPath -UseBasicParsing

if (-not (Test-Path $tempPath) -or (Get-Item $tempPath).Length -lt 500KB) {
    throw "Download failed or file is too small."
}

# --- Run installer (no UAC prompt — per-user install) ----------------------
# The installer's "Add to PATH" task is ticked by default; use /TASKS=! to opt out.
$installerArgs = @('/VERYSILENT', '/NORESTART')
if ($NoPath) {
    $installerArgs += '/TASKS=!addtopath'
}

Write-Host "Installing..."
$proc = Start-Process -FilePath $tempPath -ArgumentList $installerArgs -Wait -PassThru
if ($proc.ExitCode -ne 0) {
    throw "Installer exited with code $($proc.ExitCode)."
}

Remove-Item $tempPath -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "W365 CLI installed." -ForegroundColor Green
if (-not $NoPath) {
    Write-Host "Open a new terminal and type 'w365cli' to get started." -ForegroundColor Green
} else {
    $exePath = Join-Path $env:LOCALAPPDATA 'Programs\W365CLI\W365Cli.exe'
    Write-Host "Run it with: $exePath" -ForegroundColor Green
}
Write-Host ""
