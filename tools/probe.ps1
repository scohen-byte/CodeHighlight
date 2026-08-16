# Runs one modSelfTest macro with a single string argument, for one-off
# experiments. Same shape as the other drivers: fixed module list, few calls.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,
    [Parameter(Mandatory = $true)][string]$Proc,
    [string]$Arg = "",
    [string]$Scratch = 'C:\Users\User\ppt-lab\probe-run.pptm'
)
$ErrorActionPreference = 'Stop'
$MODULES = @('modTheme','modSpec','modLangRegistry','modLangPython',
             'modLexer','modBlock','modRender','modGutter','modGuides','modRibbon','modSelfTest')
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
} finally { $pres.Saved = $true; $pres.Close(); $ppt.Quit() }
