# Runs the VBA scanner over the sample corpus and writes one mask file per
# sample. Called by tools/run-lexer-tests.sh, which does the diffing.
#
# Builds a throwaway macro-enabled deck, imports the modules into it, and calls
# modSelfTest over COM. Nothing here touches the installed .ppam, so a broken
# lexer cannot break the add-in you have loaded.
#
# Requires "Trust access to the VBA project object model" (Trust Center > Macro
# Settings). Without it VBProject access is blocked - see the null check below.
#
# HOW TO CALL A MACRO FROM HERE. Three traps, each reporting as something else:
#
#   1. $ppt.Run($macro, $a, $b) does not bind - PowerShell cannot match the COM
#      signature. Go through InvokeMember with the macro name and arguments in
#      ONE FLAT array. Nesting the arguments as a sub-array fails too.
#   2. Never name a PowerShell parameter $Args. It collides with the automatic
#      $args variable, so the value silently is not what you passed.
#   3. Keep this script simple. An earlier version that built the import list
#      from a hashtable and made several warm-up calls first got "Sub or
#      function not defined" for procedures that were demonstrably present, and
#      that ran fine from a simpler script importing the identical files in the
#      identical order. Never root-caused. This version deliberately follows the
#      shape that works: fixed module list, one pass, no warm-up calls.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,      # folder of .bas files
    [Parameter(Mandatory = $true)][string]$SampleDir,   # folder of .py files
    [Parameter(Mandatory = $true)][string]$OutDir,      # where masks are written
    [string]$Lang = 'python',
    [string]$Scratch = 'C:\Users\User\ppt-lab\lexertest.pptm'
)

$ErrorActionPreference = 'Stop'

# A UNIQUE scratch file per run. A shared name gets left locked by any run
# that died badly, and then every later run fails on a file it cannot delete -
# which looks like a fault in whatever is being tested.
if (-not $Scratch) { $Scratch = "C:\Users\User\ppt-lab\lexertest-" + [guid]::NewGuid().ToString('N').Substring(0,8) + '.pptm' }
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

# Dependency order. modLexer refers to the TokenClass enum in modTheme and to
# LangDef in modLangRegistry, so those come first. modRibbon needs IRibbonUI and
# has nothing to do with the lexer, so it is not in the list at all.
$MODULES = @('modTheme', 'modLangRegistry', 'modLangPython', 'modLexer', 'modSelfTest')

if (Test-Path $Scratch) { Remove-Item $Scratch -Force }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

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

$pres = $ppt.Presentations.Add($true)
try {
    # Saved first so the macro can be addressed as <file>!<Module>.<Sub>.
    $pres.SaveAs($Scratch, $ppSaveAsOpenXMLPresentationMacroEnabled)

    # Blocked VBA project access does not raise - PowerPoint hands back a null
    # VBProject and an empty VBE.VBProjects collection, so checking that the VBE
    # object merely exists proves nothing.
    if ($null -eq $pres.VBProject) {
        throw ("VBA project access is blocked, so the modules cannot be imported.`n" +
               "  Turn on: File > Options > Trust Center > Trust Center Settings >`n" +
               "           Macro Settings > Trust access to the VBA project object model")
    }

    foreach ($m in $MODULES) {
        $null = $pres.VBProject.VBComponents.Import((Join-Path $SrcDir "$m.bas"))
    }
    Write-Output ("imported: " + ($MODULES -join ', '))

    # Language is set once rather than passed on every call.
    $null = Run-Macro 'SetLanguage' @($Lang)

    # The whole corpus in ONE macro call. See modSelfTest.RunCorpus for why:
    # repeated Application.Run calls against a freshly imported project proved
    # unreliable in a way never root-caused, and crossing the COM boundary once
    # is faster regardless.
    $report = Run-Macro 'RunCorpus' @($SampleDir, $OutDir)

    foreach ($line in ($report -split "`n")) {
        $line = $line.Trim()
        if ($line.Length -eq 0) { continue }
        $f = $line -split ' ', 3
        if ($f[0] -eq 'TOTAL') {
            Write-Output ("{0,-16} {1} ms for {2} chars" -f 'whole corpus', $f[1], $f[2])
        } elseif ($f[1] -eq 'ERROR') {
            Write-Output ("{0,-16} FAILED :: {1}" -f $f[0], $f[2])
        } else {
            Write-Output ("{0,-16} mask {1,6} chars  {2,5} spans" -f $f[0], $f[1], $f[2])
        }
    }

}
finally {
    $pres.Saved = $true
    $pres.Close()
    $ppt.Quit()
    Start-Sleep -Milliseconds 500
    Remove-Item $Scratch -Force -ErrorAction SilentlyContinue
}
