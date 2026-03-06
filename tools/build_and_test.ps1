#!/usr/bin/env pwsh
<#
Builds an Excel add-in from exported VBA modules and runs tests.
Requirements:
- Windows with Excel installed
- Excel Trust Center: "Trust access to the VBA project object model" enabled
#>

$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$libDir = Join-Path $scriptDir 'lib'

. (Join-Path $libDir 'ribbon_openxml.ps1')
. (Join-Path $libDir 'excel_build.ps1')
. (Join-Path $libDir 'addin_tests.ps1')

$repoRoot = Split-Path -Parent $scriptDir
$vbaDir = Join-Path $repoRoot 'src\vba'
$dist = Join-Path $repoRoot 'dist'
if (!(Test-Path $dist)) {
    New-Item -Path $dist -ItemType Directory | Out-Null
}

Write-Host "Using repo root: $repoRoot"

$excel = $null
try {
    $excel = New-ExcelApplication

    $sourceWorkbook = $excel.Workbooks.Add()
    try {
        Import-VbaModules -Workbook $sourceWorkbook -VbaDirectory $vbaDir
        $addInPath = Save-WorkbookAsAddIn -Workbook $sourceWorkbook -DistDirectory $dist
    }
    finally {
        $sourceWorkbook.Close($false)
    }

    Write-Host 'Embedding custom ribbon UI into add-in package...'
    Add-CustomRibbonToAddIn -XlamPath $addInPath

    $resultsText = Invoke-AddInTests -Excel $excel -AddInPath $addInPath
    $resultsFile = Join-Path $dist 'test-results.txt'
    Set-Content -Path $resultsFile -Value $resultsText -Encoding UTF8

    Write-Host "Build and tests complete. Results at $resultsFile"
}
finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}
