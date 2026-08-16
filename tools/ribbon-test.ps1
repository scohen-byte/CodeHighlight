# Drives the ribbon commands end to end, then closes and reopens the deck to
# check that shape tags survived the round trip.
#
# Same rules as tools/lexer-test.ps1: fixed module list, few macro calls.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,
    [Parameter(Mandatory = $true)][string]$Sample,
    [Parameter(Mandatory = $true)][string]$Png,
    [string]$Lang = 'python',
    [string]$Scratch = 'C:\Users\User\ppt-lab\ribbon.pptm'
)

$ErrorActionPreference = 'Stop'

# A UNIQUE scratch file per run. A shared name gets left locked by any run
# that died badly, and then every later run fails on a file it cannot delete -
# which looks like a fault in whatever is being tested.
if (-not $Scratch) { $Scratch = "C:\Users\User\ppt-lab\ribbon-" + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptm' }
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

# modRibbon IS imported here, unlike in the lexer harness - it is the thing
# under test. It needs IRibbonUI, which the Office library provides by default.
$MODULES = @('modTheme', 'modSpec', 'modLangRegistry', 'modLangPython',
             'modLexer', 'modBlock', 'modRender', 'modGutter', 'modGuides',
             'modOptions', 'modSwatch', 'modNote', 'modArrow', 'modRibbon', 'modSelfTest')

if (Test-Path $Scratch) { Remove-Item $Scratch -Force }
New-Item -ItemType Directory -Path (Split-Path $Png -Parent) -Force | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$type = $ppt.GetType()
$flags = [System.Reflection.BindingFlags]::InvokeMethod
$name = Split-Path $Scratch -Leaf

function Run-Macro {
    param([string]$Proc, [object[]]$MacroArgs = @())
    $all = New-Object object[] ($MacroArgs.Count + 1)
    $all[0] = "$script:name!modSelfTest.$Proc"
    for ($i = 0; $i -lt $MacroArgs.Count; $i++) { $all[$i + 1] = $MacroArgs[$i] }
    return $script:type.InvokeMember('Run', $script:flags, $null, $script:ppt, $all)
}

$ppt.Visible = -1     # msoTrue. Selection and Export both need a real window.
$pres = $ppt.Presentations.Add($true)
try {
    $pres.SaveAs($Scratch, $ppSaveAsOpenXMLPresentationMacroEnabled)
    if ($null -eq $pres.VBProject) { throw "VBA project access is blocked." }

    foreach ($m in $MODULES) {
        $null = $pres.VBProject.VBComponents.Import((Join-Path $SrcDir "$m.bas"))
    }
    Write-Output ("imported: " + ($MODULES -join ', '))

    Write-Output "== commands =="
    Write-Output (Run-Macro 'RibbonSliceTest' @($Sample, $Png))

    $pres.Save()
    $tagsBefore = Run-Macro 'ReportTags'
    $pres.Close()

    # Reopen from disk. Shape tags are expected to persist in the file, but the
    # design leans on it, so it is worth ten minutes of proof.
    $pres = $ppt.Presentations.Open($Scratch)
    $tagsAfter = Run-Macro 'ReportTags'

    Write-Output "== tags before save =="
    Write-Output $tagsBefore.TrimEnd()
    Write-Output "== tags after reopen =="
    Write-Output $tagsAfter.TrimEnd()
    if ($tagsBefore.Trim() -eq $tagsAfter.Trim()) {
        Write-Output "tags_survive_reopen=1"
    } else {
        Write-Output "tags_survive_reopen=0"
    }
}
finally {
    if ($null -ne $pres) { $pres.Saved = $true; $pres.Close() }
    $ppt.Quit()
}
