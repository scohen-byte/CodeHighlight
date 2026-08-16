# Builds CodeHighlight.ppam from the .bas sources.
#
# Called by tools/build-addin.sh, which then injects the ribbon. Produces a
# fresh add-in every time rather than mutating the installed one, so a bad build
# cannot corrupt a working install.
param(
    [Parameter(Mandatory = $true)][string]$SrcDir,
    [Parameter(Mandatory = $true)][string]$Out     # full path to the .pptm
)

$ErrorActionPreference = 'Stop'

# PowerPoint REFUSES to SaveAs an add-in format over COM - both ppSaveAsAddIn (8)
# and ppSaveAsOpenXMLAddin (30) come back E_FAIL, while the macro-enabled
# presentation format saves fine. So this saves a .pptm and the shell script
# converts it: a .ppam differs only in the content type of the main part.
$ppSaveAsOpenXMLPresentationMacroEnabled = 25

# Dependency order, and SHIPPING modules only. modSelfTest is test scaffolding -
# it reads and writes files by path and has no business in a distributed add-in.
$MODULES = @('modTheme', 'modSpec', 'modLangRegistry', 'modLangPython',
             'modLexer', 'modBlock', 'modRender', 'modRibbon')

if (Test-Path $Out) { Remove-Item $Out -Force }
New-Item -ItemType Directory -Path (Split-Path $Out -Parent) -Force | Out-Null

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add($true)
try {
    if ($null -eq $pres.VBProject) {
        throw ("VBA project access is blocked, so the add-in cannot be built.`n" +
               "  Turn on: File > Options > Trust Center > Trust Center Settings >`n" +
               "           Macro Settings > Trust access to the VBA project object model")
    }

    foreach ($m in $MODULES) {
        $path = Join-Path $SrcDir "$m.bas"
        if (-not (Test-Path $path)) { throw "missing module: $path" }
        try { $null = $pres.VBProject.VBComponents.Import($path) }
        catch { throw ("importing $m failed: " + $_.Exception.Message) }
    }

    # An add-in has no slides, but PowerPoint will not save a presentation with
    # none of them. The slide never shows: loading a .ppam does not display it.
    if ($pres.Slides.Count -eq 0) { $null = $pres.Slides.Add(1, 12) }

    try { $pres.SaveAs($Out, $ppSaveAsOpenXMLPresentationMacroEnabled) }
    catch { throw ("SaveAs failed: " + $_.Exception.Message) }
    Write-Output ("built " + (Split-Path $Out -Leaf) + " with " + $MODULES.Count + " modules")
}
finally {
    $pres.Saved = $true
    $pres.Close()
    $ppt.Quit()
}
