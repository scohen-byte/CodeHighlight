# Checks whether this machine can run the PyCodeHighlight add-in.
# Safe to send to a colleague - it only reads, it changes nothing.
#
#   powershell.exe -NoProfile -ExecutionPolicy Bypass -File check-env.ps1

$ErrorActionPreference = 'SilentlyContinue'
Write-Output "=== PowerPoint add-in readiness ==="

# --- version -------------------------------------------------------------
$ppt = New-Object -ComObject PowerPoint.Application
if (-not $ppt) {
    Write-Output "FAIL  PowerPoint not installed, or not the desktop version."
    exit 1
}
$ver = $ppt.Version
$exe = Get-Item (Join-Path $ppt.Path 'POWERPNT.EXE') -EA SilentlyContinue
$ppt.Quit()

$build = if ($exe) { $exe.VersionInfo.ProductVersion } else { 'unknown' }
$bits = if ([Environment]::Is64BitOperatingSystem) { '64-bit OS' } else { '32-bit OS' }
Write-Output "Version   : $ver  (build $build)"
Write-Output "Path      : $(if ($exe) { $exe.FullName } else { 'not found' })"
Write-Output "OS        : $bits"

# Click-to-Run tells us 365 vs perpetual, which $ver cannot.
$c2r = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -EA SilentlyContinue
if ($c2r) {
    Write-Output "Product   : $($c2r.ProductReleaseIds)"
    Write-Output "Channel   : $($c2r.CDNBaseUrl -replace '.*/','')"
}

$name = switch -Regex ($ver) {
    '^16' { 'Office 2016 / 2019 / 2021 / 2024 / Microsoft 365' }
    '^15' { 'Office 2013' }
    '^14' { 'Office 2010' }
    default { "Office $ver" }
}
Write-Output "Family    : $name"

if ([double]$ver -ge 14) {
    Write-Output "RESULT    : OK - meets the 2010+ floor for .ppam + custom ribbon."
} else {
    Write-Output "RESULT    : TOO OLD - needs PowerPoint 2010 or newer."
}

# --- macro trust ---------------------------------------------------------
Write-Output ""
Write-Output "=== Macro security ==="
$paths = @(
    "HKCU:\Software\Policies\Microsoft\Office\$ver\PowerPoint\Security",  # policy wins
    "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security"
)
$warn = $null; $src = $null
foreach ($p in $paths) {
    $v = (Get-ItemProperty -Path $p -Name VBAWarnings -EA SilentlyContinue).VBAWarnings
    if ($null -ne $v -and $null -eq $warn) { $warn = $v; $src = $p }
}
$meaning = @{
    1 = 'Enable all macros (no prompt)'
    2 = 'Disable with notification  <- Windows default, fine'
    3 = 'Disable except digitally signed  <- sign the .ppam'
    4 = 'Disable all without notification  <- BLOCKED, needs a Trusted Location'
}
if ($null -eq $warn) {
    Write-Output "VBAWarnings not set - Office default applies (disable with notification). Fine."
} else {
    Write-Output "VBAWarnings = $warn : $($meaning[[int]$warn])"
    if ($src -like '*Policies*') { Write-Output "  (set by IT policy - you cannot change this yourself)" }
}

# --- VBA project access --------------------------------------------------
# Only the BUILD needs this, never the install: tools/build-addin.ps1 imports
# the modules over COM, and that goes through the VBA project object model.
# Off is the Office default and is reported as such, not as a failure - someone
# who was handed a .ppam is fine without it.
Write-Output ""
Write-Output "=== VBA project access (build machines only) ==="
$vbom = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security" `
                          -Name AccessVBOM -EA SilentlyContinue).AccessVBOM
if ($vbom -eq 1) {
    Write-Output "AccessVBOM = 1 : trusted - tools/build-addin.sh can build the add-in."
} else {
    Write-Output "AccessVBOM = $(if ($null -eq $vbom) { 'not set' } else { $vbom }) : blocked - installing works, BUILDING does not."
    Write-Output "  To build: File > Options > Trust Center > Trust Center Settings >"
    Write-Output "            Macro Settings > Trust access to the VBA project object model"
}

# --- trusted locations ---------------------------------------------------
Write-Output ""
Write-Output "=== Trusted Locations (PowerPoint) ==="
$tl = "HKCU:\Software\Microsoft\Office\$ver\PowerPoint\Security\Trusted Locations"
$found = Get-ChildItem $tl -EA SilentlyContinue
if ($found) {
    foreach ($k in $found) {
        $p = (Get-ItemProperty $k.PSPath).Path
        if ($p) { Write-Output "  $p" }
    }
} else {
    Write-Output "  none configured (normal - we can add one at install time)"
}

Write-Output ""
Write-Output "Done. Send this whole output back."
