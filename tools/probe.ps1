# Runs one modSelfTest macro with a single string argument, for one-off
# experiments. Same shape as the other drivers: fixed module list, few calls.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,
    [Parameter(Mandatory = $true)][string]$Proc,
    [string]$Arg = "",
    [string]$Scratch = ''
)
$ErrorActionPreference = 'Stop'

# A UNIQUE scratch file per run. A shared name gets left locked by any run
# that died badly, and then every later run fails on a file it cannot delete -
# which looks like a fault in whatever is being tested.
# The default is EMPTY on purpose. This used to default to a fixed filename,
# which meant the guard below never fired: -not $Scratch was never true, every
# run reused one file, and any run that died badly left it locked so that every
# later run failed on a file it could not delete. The comment claimed a unique
# name per run and the parameter quietly took it away.
if (-not $Scratch) { $Scratch = (Join-Path $env:USERPROFILE "ppt-lab\probe-run-") + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptm' }
New-Item -ItemType Directory -Path (Split-Path $Scratch -Parent) -Force | Out-Null
$MODULES = @('modTheme','modSpec','modLangRegistry','modLangPython',
             'modLexer','modBlock','modRender','modGutter','modGuides',
             'modOptions','modSwatch','modOutput','modNote','modArrow','modRibbon','modSelfTest')
if (Test-Path $Scratch) { Remove-Item $Scratch -Force }
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = -1
$pres = $ppt.Presentations.Add($true)
try {
    $pres.SaveAs($Scratch, 25)
    foreach ($m in $MODULES) { $null = $pres.VBProject.VBComponents.Import((Join-Path $SrcDir "$m.bas")) }
    $all = New-Object object[] 2
    $all[0] = (Split-Path $Scratch -Leaf) + "!modSelfTest.$Proc"
    $all[1] = $Arg
    Write-Output $ppt.GetType().InvokeMember('Run',
        [System.Reflection.BindingFlags]::InvokeMethod, $null, $ppt, $all)
} finally {
    $pres.Saved = $true; $pres.Close(); $ppt.Quit()
    Start-Sleep -Milliseconds 500
    Remove-Item $Scratch -Force -ErrorAction SilentlyContinue
}
