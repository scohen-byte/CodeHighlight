# Runs the VBA scanner over the sample corpus and writes one mask file per
# sample. Called by tools/run-lexer-tests.sh, which does the diffing.
#
# Builds a throwaway macro-enabled deck, imports the modules into it, and calls
# modSelfTest over COM. Nothing here touches the installed .ppam, so a broken
# lexer cannot break the add-in you have loaded.
#
# Requires "Trust access to the VBA project object model" (Trust Center > Macro
# Settings). Without it VBProject access raises, and the only route is importing
# the modules by hand in the VBA editor.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,      # folder of .bas files
    [Parameter(Mandatory = $true)][string]$SampleDir,   # folder of .py files
    [Parameter(Mandatory = $true)][string]$OutDir,      # where masks are written
    [string]$Lang = 'python',
    [string]$Scratch = 'C:\Users\User\ppt-lab\lexertest.pptm'
)

$ErrorActionPreference = 'Stop'
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

if (Test-Path $Scratch) { Remove-Item $Scratch -Force }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add($true)
try {
    # Saving first so the macro can be addressed as <file>!<Module>.<Sub>.
    # Running a macro in an unsaved presentation depends on its window title,
    # which is exactly the kind of thing that breaks silently.
    $pres.SaveAs($Scratch, $ppSaveAsOpenXMLPresentationMacroEnabled)
    $name = Split-Path $Scratch -Leaf

    # Blocked VBA project access does not raise - PowerPoint hands back a null
    # VBProject and an empty VBE.VBProjects collection, so checking that the VBE
    # object merely exists proves nothing. Check for the null and say what to do.
    if ($null -eq $pres.VBProject) {
        throw ("VBA project access is blocked, so the modules cannot be imported.`n" +
               "  Turn on: File > Options > Trust Center > Trust Center Settings >`n" +
               "           Macro Settings > Trust access to the VBA project object model`n" +
               "  (Per-user setting. No policy key is present, so it is yours to change.)")
    }

    foreach ($bas in Get-ChildItem -Path $SrcDir -Filter *.bas | Sort-Object Name) {
        # modRibbon needs IRibbonUI and is not part of the lexer, so skip it.
        if ($bas.BaseName -eq 'modRibbon') { continue }
        $null = $pres.VBProject.VBComponents.Import($bas.FullName)
        Write-Output ("imported " + $bas.BaseName)
    }

    foreach ($py in Get-ChildItem -Path $SampleDir -Filter *.py | Sort-Object Name) {
        $out = Join-Path $OutDir ($py.BaseName + '.mask')
        $len = $ppt.Run("$name!modSelfTest.MaskFileToFile", $py.FullName, $out, $Lang)
        $spans = $ppt.Run("$name!modSelfTest.SpanCountOf", $py.FullName, $Lang)
        Write-Output ("{0,-16} mask {1,6} chars  {2,5} spans" -f $py.Name, $len, $spans)
    }

    $ms = $ppt.Run("$name!modSelfTest.TimeTokenize",
                   (Join-Path $SampleDir 'long.py'), $Lang, 20)
    Write-Output ("long.py x20 tokenize: {0} ms" -f $ms)
}
finally {
    $pres.Saved = $true
    $pres.Close()
    $ppt.Quit()
}
