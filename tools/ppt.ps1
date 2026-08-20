# PowerPoint COM helpers, invoked from WSL via powershell.exe.
#
#   powershell.exe -NoProfile -File "$(wslpath -w tools/ppt.ps1)" -Action export `
#       -Path $env:USERPROFILE\ppt-lab\lab.pptx -Out $env:USERPROFILE\ppt-lab\png
#
param(
    [Parameter(Mandatory = $true)][ValidateSet('export', 'runmacro')] [string]$Action,
    [Parameter(Mandatory = $true)][string]$Path,
    [string]$Out,
    [string]$Macro,
    [int]$Width = 1600,
    [int]$Height = 900
)

$ErrorActionPreference = 'Stop'

$ppt = New-Object -ComObject PowerPoint.Application
try {
    # WithWindow must be true - PowerPoint refuses many operations headless.
    $pres = $ppt.Presentations.Open($Path, $true, $false, $true)
    try {
        switch ($Action) {
            'export' {
                if (Test-Path $Out) { Remove-Item $Out -Recurse -Force }
                New-Item -ItemType Directory -Path $Out -Force | Out-Null
                $pres.Export($Out, 'PNG', $Width, $Height)
                Write-Output "exported $($pres.Slides.Count) slides to $Out"
            }
            'runmacro' {
                $ppt.Run($Macro) | Out-Null
                Write-Output "ran $Macro"
            }
        }
    }
    finally { $pres.Close() }
}
finally { $ppt.Quit() }
