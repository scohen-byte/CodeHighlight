# Check whether imageMso ids exist in PowerPoint, before putting them in the
# ribbon XML.
#
#   powershell.exe -NoProfile -ExecutionPolicy Bypass `
#       -File "$(wslpath -w tools/check-imagemso.ps1)" -Ids Repaint,RefreshAll
#
# Worth having because a bad imageMso fails silently: the control renders with a
# blank image and the ribbon loads normally, so the mistake survives review and
# only shows up when someone looks at the tab. Several plausible-sounding ids
# ("Repaint", "ListNumbering", "ShowLineNumbers") do not exist in PowerPoint.
#
# How the test works: GetImageMso raises E_INVALIDARG - "Value does not fall
# within the expected range" - for an unknown id. For a KNOWN id it still raises
# here, with E_UNEXPECTED, because PowerShell cannot marshal the IPictureDisp
# that comes back. So the distinction is in the error message, not in whether
# the call succeeds. Do not "fix" this by treating any exception as failure.
param(
    [Parameter(Mandatory = $true)][string[]]$Ids
)

$ErrorActionPreference = 'Stop'

# Invoked with -File, PowerShell hands "a,b,c" over as a single string rather
# than splitting it, so do the splitting here. Space-separated args still work.
$Ids = $Ids | ForEach-Object { $_ -split ',' } | Where-Object { $_ -ne '' }

$ppt = New-Object -ComObject PowerPoint.Application
# GetImageMso needs a presentation open, otherwise every id reports as invalid.
$pres = $ppt.Presentations.Add($true)
try {
    foreach ($id in $Ids) {
        try {
            $null = $ppt.CommandBars.GetImageMso($id, 16, 16)
            Write-Output ("VALID    " + $id)
        }
        catch {
            if ($_.Exception.Message -match 'expected range') {
                Write-Output ("INVALID  " + $id)
            }
            else {
                Write-Output ("VALID    " + $id)
            }
        }
    }
}
finally {
    $pres.Close()
    $ppt.Quit()
}
