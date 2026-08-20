# Thin slice: build one code block on one slide, highlight it, export a PNG.
#
# Called by tools/run-slice.sh. Same rules as tools/lexer-test.ps1 - keep this
# script simple and the macro-call count low. See the header there.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,     # folder of .bas files
    [Parameter(Mandatory = $true)][string]$Sample,     # the .py file to render
    [Parameter(Mandatory = $true)][string]$Png,        # where to write the PNG
    [string]$Lang = 'python',
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
if (-not $Scratch) { $Scratch = (Join-Path $env:USERPROFILE "ppt-lab\slice-") + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptm' }
New-Item -ItemType Directory -Path (Split-Path $Scratch -Parent) -Force | Out-Null
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

# Dependency order: modLexer needs TokenClass from modTheme and LangDef from
# modLangRegistry, modBlock needs modSpec and modTheme, modRender needs both.
# modRibbon is excluded - it needs IRibbonUI and plays no part here.
$MODULES = @('modTheme', 'modSpec', 'modLangRegistry', 'modLangPython',
             'modLexer', 'modBlock', 'modRender', 'modGutter', 'modGuides',
             'modOptions', 'modSwatch', 'modOutput', 'modNote', 'modArrow',
             'modRibbon', 'modSelfTest')

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

# The slice needs a visible window: ActivePresentation and Slide.Export both
# refuse to work properly on a headless presentation.
# msoTrue is -1, and Visible is an MsoTriState, not a Boolean. Assigning $true
# raises "Invalid cast from System.Boolean".
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

    $null = Run-Macro 'SetLanguage' @($Lang)
    Write-Output (Run-Macro 'BuildSlice' @($Sample, $Png))

    # Kept so the result can be opened by hand and poked at.
    $pres.Save()
}
finally {
    $pres.Saved = $true
    $pres.Close()
    $ppt.Quit()
    Start-Sleep -Milliseconds 500
    Remove-Item $Scratch -Force -ErrorAction SilentlyContinue
}
