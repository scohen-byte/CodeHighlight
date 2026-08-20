# Runs one modSelfTest function against a real PowerPoint, and brings back a PNG.
#
#   tools/run-feature.sh NoteTest [sample-name]
#
# The slice and lexer harnesses each drive one fixed entry point. Every feature
# since - emphasis, step through, hide, notes - needed the same thing with a
# different function name, and was tested through a throwaway script that then
# went missing. This is that script, kept.
#
# Same rules as tools/slice-test.ps1: keep the macro-call count low, and use a
# UNIQUE scratch filename. A shared one gets left locked by any run that died
# badly, and every later run then fails on a file it cannot delete - which looks
# like a fault in whatever is being tested.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,
    [Parameter(Mandatory = $true)][string]$Proc,      # a Public Function in modSelfTest
    [Parameter(Mandatory = $true)][string]$Sample,
    [Parameter(Mandatory = $true)][string]$Png
)

$ErrorActionPreference = 'Stop'

# Dependency order. modNote needs modBlock and modGutter; modRibbon needs all of
# them and is included because these tests drive the Do* commands.
$MODULES = @('modTheme', 'modSpec', 'modLangRegistry', 'modLangPython',
             'modLexer', 'modBlock', 'modRender', 'modGutter', 'modGuides',
             'modOptions', 'modSwatch', 'modOutput', 'modNote', 'modArrow', 'modRibbon', 'modSelfTest')

$Scratch = (Join-Path $env:USERPROFILE "ppt-lab\feature-") +
           [guid]::NewGuid().ToString('N').Substring(0, 8) + '.pptm'
New-Item -ItemType Directory -Path (Split-Path $Scratch -Parent) -Force | Out-Null
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

New-Item -ItemType Directory -Path (Split-Path $Png -Parent) -Force | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
# msoTrue is -1, and Visible is an MsoTriState, not a Boolean. Assigning $true
# raises "Invalid cast from System.Boolean". Export needs a visible window.
$ppt.Visible = -1
$pres = $ppt.Presentations.Add($true)
try {
    $pres.SaveAs($Scratch, $ppSaveAsOpenXMLPresentationMacroEnabled)

    if ($null -eq $pres.VBProject) {
        throw ("VBA project access is blocked. Turn on: File > Options >`n" +
               "  Trust Center > Trust Center Settings > Macro Settings >`n" +
               "  Trust access to the VBA project object model")
    }

    foreach ($m in $MODULES) {
        $null = $pres.VBProject.VBComponents.Import((Join-Path $SrcDir "$m.bas"))
    }
    Write-Output ("imported: " + ($MODULES -join ', '))

    $all = New-Object object[] 3
    $all[0] = (Split-Path $Scratch -Leaf) + "!modSelfTest.$Proc"
    $all[1] = $Sample
    $all[2] = $Png
    Write-Output "== $Proc =="
    Write-Output $ppt.GetType().InvokeMember(
        'Run', [Reflection.BindingFlags]::InvokeMethod, $null, $ppt, $all)
}
finally {
    if ($null -ne $pres) { $pres.Saved = $true; $pres.Close() }
    $ppt.Quit()
    Start-Sleep -Milliseconds 500
    Remove-Item $Scratch -Force -ErrorAction SilentlyContinue
}
