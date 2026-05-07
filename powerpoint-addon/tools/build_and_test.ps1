#!/usr/bin/env pwsh
<#
Builds a PowerPoint add-in from exported VBA modules.
Requirements:
- Windows with PowerPoint installed
- PowerPoint Trust Center: "Trust access to the VBA project object model" enabled
#>

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$libDir = Join-Path $scriptDir 'lib'

. (Join-Path $libDir 'powerpoint_build.ps1')
. (Join-Path $libDir 'ribbon_openxml.ps1')

$projectRoot = Split-Path -Parent $scriptDir
$vbaDir = Join-Path $projectRoot 'src\vba'
$dist = Join-Path $projectRoot 'dist'
if (!(Test-Path $dist)) {
    New-Item -Path $dist -ItemType Directory | Out-Null
}

Write-Host "Using project root: $projectRoot"

$ppt = $null
$tempPptmPath = $null
try {
    $ppt = New-PowerPointApplication

    $tempPptmPath = Join-Path $dist ('_build-temp-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.pptm')

    $sourcePresentation = $ppt.Presentations.Add()
    try {
        # Force a macro-enabled file on disk so VBProject is available.
        $sourcePresentation.SaveAs($tempPptmPath) | Out-Null
        $sourcePresentation.Close()

        $sourcePresentation = $ppt.Presentations.Open($tempPptmPath, $false, $false, $false)
        Import-VbaModules -Presentation $sourcePresentation -VbaDirectory $vbaDir
        $addInPath = Save-PresentationAsAddIn -Presentation $sourcePresentation -DistDirectory $dist
    }
    finally {
        $sourcePresentation.Close()
    }

    Write-Host 'Embedding custom ribbon UI into add-in package...'
    Add-CustomRibbonToAddIn -PpamPath $addInPath

    Write-Host 'Loading the built add-in into PowerPoint...'
    Load-AddInIntoPowerPoint -PowerPoint $ppt -AddInPath $addInPath

    Write-Host "Build complete. Add-in output: $addInPath"
}
finally {
    if ($tempPptmPath -and (Test-Path $tempPptmPath)) {
        Remove-Item -Path $tempPptmPath -Force -ErrorAction SilentlyContinue
    }
    if ($ppt) {
        $ppt.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
    }
}
