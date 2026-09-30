# release-installer.ps1 - Stage a newly built SXUPaper installer for the
# promo/download site (sxupaper.b100.top). Encodes the release checklist so
# no file can be forgotten:
#   1. versioned exe  (SXUPaper-Beta-<v>-x64-Setup.exe   -> promo page button)
#   2. latest exe     (SXUPaper-Beta-latest-x64-Setup.exe -> client check-update)
#   3. sha256.txt     (hash of latest, auto-shown on the promo page)
#   4. version.txt    (version only, auto-shown on the promo page)
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File release-installer.ps1 -InstallerPath <built-exe>
#   (-Version is optional; derived from the filename, else from package.json)
#
# After running: page needs NO change (it fetches version.txt + sha256 at runtime).
# Manual follow-ups it will remind you about are printed at the end.

param(
    [Parameter(Mandatory = $true)][string]$InstallerPath,
    [string]$Version = ""
)

$ErrorActionPreference = 'Stop'

$repoRoot   = Split-Path -Parent $PSScriptRoot          # deploy/ -> repo root
$downloadDir = Join-Path $repoRoot 'deploy\download'

# --- resolve installer -------------------------------------------------------
$item = Get-Item -LiteralPath $InstallerPath
if (-not $item.Exists) { throw "Installer not found: $InstallerPath" }

# --- resolve version ---------------------------------------------------------
if ($Version -eq "" -and $item.Name -match '^SXUPaper-Beta-([0-9.]+)-x64-Setup\.exe$') {
    $Version = $Matches[1]
}
if ($Version -eq "") {
    $pkg = Get-Content (Join-Path $repoRoot 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    $Version = $pkg.version
    Write-Host "Version derived from package.json: $Version"
}
if ($Version -notmatch '^[0-9]+(\.[0-9]+){1,3}$') {
    throw "Cannot determine a valid version (got: '$Version'). Pass -Version explicitly."
}

# --- stage the 4 files -------------------------------------------------------
$verName    = "SXUPaper-Beta-$Version-x64-Setup.exe"
$latestName = "SXUPaper-Beta-latest-x64-Setup.exe"

New-Item -ItemType Directory -Force -Path $downloadDir | Out-Null
function Copy-IfDifferent([string]$src, [string]$dst) {
    $srcFull = (Get-Item -LiteralPath $src).FullName
    $dstFull = [IO.Path]::GetFullPath($dst)
    if ($srcFull -ieq $dstFull) { Write-Host "skip copy (same file): $(Split-Path -Leaf $dst)"; return }
    Copy-Item -LiteralPath $src -Destination $dst -Force
}
Copy-IfDifferent $item.FullName (Join-Path $downloadDir $verName)
Copy-IfDifferent $item.FullName (Join-Path $downloadDir $latestName)

$hash = (Get-FileHash (Join-Path $downloadDir $latestName) -Algorithm SHA256).Hash.ToLower()
[IO.File]::WriteAllText((Join-Path $downloadDir "$latestName.sha256.txt"), "$hash  $latestName`n")
[IO.File]::WriteAllText((Join-Path $downloadDir 'version.txt'), $Version)

# --- verify ------------------------------------------------------------------
foreach ($f in @($verName, $latestName)) {
    $p = Join-Path $downloadDir $f
    if (-not (Test-Path -LiteralPath $p)) { throw "Staging failed, missing: $f" }
    $h = (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash.ToLower()
    if ($h -ne $hash) { throw "SHA-256 mismatch after copy: $f" }
}

Write-Host ""
Write-Host "OK  version=$Version  staged into deploy/download/"
Write-Host ("  {0}  {1:N0} bytes" -f $verName,    (Get-Item (Join-Path $downloadDir $verName)).Length)
Write-Host ("  {0}  {1:N0} bytes" -f $latestName, (Get-Item (Join-Path $downloadDir $latestName)).Length)
Write-Host ("  {0}.sha256.txt" -f $latestName)
Write-Host ("  version.txt = {0}" -f $Version)
Write-Host ""
Write-Host "Promo page needs NO change (it fetches version.txt + sha256 at runtime)."
Write-Host ""
Write-Host "Manual follow-ups:"
Write-Host ("  1. Client check-update: server env SXUPAPER_LATEST_VERSION={0}" -f $Version)
Write-Host ("     and SXUPAPER_DOWNLOAD_URL=https://sxupaperdownload.b100.top/{0}" -f $verName)
Write-Host "     (recreate the controller container after changing its env)"
Write-Host "  2. If a second production host exists, sync these 4 files to its download dir too."
Write-Host "  3. Verify: https://sxupaper.b100.top shows the new version and the button"
Write-Host "     downloads the versioned filename."
