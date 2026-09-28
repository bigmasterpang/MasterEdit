param(
    [string]$FilePath = "",
    [string]$Version = "",
    [string]$ReleaseNotes = "",
    [string]$NotesFilePath = "",
    [ValidateSet("portable", "installer")]
    [string]$Mode = "portable",
    [switch]$Build
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot

# 1. Build the release (optional)
if ($Build) {
    Write-Host "[*] Building frontend + Tauri release bundle..." -ForegroundColor Cyan
    Push-Location $projectRoot
    try {
        & pnpm tauri build
        if ($LASTEXITCODE -ne 0) { throw "pnpm tauri build failed" }
    } finally {
        Pop-Location
    }
}

# 2. Read version from tauri.conf.json when not provided
if (-not $Version) {
    $confPath = Join-Path $projectRoot "src-tauri\tauri.conf.json"
    $conf = Get-Content -LiteralPath $confPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $Version = $conf.version
}
if (-not $Version) { throw "Version could not be determined" }

$releaseDir = Join-Path $projectRoot "src-tauri\target\release"
$bundleDir = Join-Path $releaseDir "bundle\nsis"

# 3. Locate the artifact
if (-not $FilePath) {
    if ($Mode -eq "portable") {
        # Portable single-file build: app frontend is embedded in the exe
        $built = Join-Path $releaseDir "MasterEdit.exe"
        if (-not (Test-Path -LiteralPath $built)) {
            $built = Join-Path $releaseDir "MasterMD.exe"
        }
        if (-not (Test-Path -LiteralPath $built)) {
            $built = Join-Path $releaseDir "mastermd.exe"
        }
        if (-not (Test-Path -LiteralPath $built)) {
            throw "Portable binary not found: $built (run with -Build first)"
        }
        $target = Join-Path $releaseDir "MasterEdit_${Version}_x64.exe"
        Copy-Item -LiteralPath $built -Destination $target -Force
        $FilePath = $target
    } else {
        $target = Join-Path $bundleDir "MasterEdit_${Version}_x64-setup.exe"
        if (-not (Test-Path -LiteralPath $target)) {
            $target = Join-Path $bundleDir "MasterMD_${Version}_x64-setup.exe"
        }
        if (-not (Test-Path -LiteralPath $target)) {
            $target = Join-Path $bundleDir "mastermd_${Version}_x64-setup.exe"
        }
        if (-not (Test-Path -LiteralPath $target)) {
            throw "Installer not found in $bundleDir (run with -Build first)"
        }
        $FilePath = $target
    }
}
Write-Host "Artifact: $(Split-Path $FilePath -Leaf) ($([math]::Round((Get-Item -LiteralPath $FilePath).Length / 1MB, 2)) MB)" -ForegroundColor Cyan

# 4. Release notes: file > git log > default
if (-not $ReleaseNotes -and $NotesFilePath -and (Test-Path -LiteralPath $NotesFilePath)) {
    $ReleaseNotes = (Get-Content -LiteralPath $NotesFilePath -Raw -Encoding UTF8).Trim()
}
if (-not $ReleaseNotes) {
    try {
        $ReleaseNotes = (git -C $projectRoot log -1 --pretty=%B).Trim()
    } catch { }
    if (-not $ReleaseNotes) { $ReleaseNotes = "MasterEdit $Version released" }
}

# 5. Publish to the overseas VM portal (single upload).
#    The domestic Aliyun node mirrors the VM automatically via cron
#    (/etc/cron.d/portal-sync -> /opt/app-portal/sync-from-vm.sh), so no second upload.
#    -SkipNotify: WeChat cards are sent by the DeepSeek Harness plugin, not by this script.
& (Join-Path $projectRoot "..\tools\publish-release.ps1") `
    -App "masteredit" `
    -Platform "windows" `
    -FilePath $FilePath `
    -Version $Version `
    -ReleaseNotes $ReleaseNotes `
    -Variant $Mode `
    -SkipNotify
