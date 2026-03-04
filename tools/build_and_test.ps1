#!/usr/bin/env pwsh
<#
Builds an Excel add-in from exported VBA modules and runs tests.
Requirements:
- Windows with Excel installed
- Excel Trust Center: "Trust access to the VBA project object model" enabled
#>

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$repoRoot = Split-Path -Parent $scriptDir
$vbaDir = Join-Path $repoRoot 'src\vba'
$dist = Join-Path $repoRoot 'dist'
if (!(Test-Path $dist)) { New-Item -Path $dist -ItemType Directory | Out-Null }

Write-Host "Using repo root: $repoRoot"
Write-Host "Importing VBA modules from: $vbaDir"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 1

# Create new workbook and import modules
$wb = $excel.Workbooks.Add()
foreach ($f in Get-ChildItem -Path $vbaDir -Filter *.bas -ErrorAction SilentlyContinue) {
    Write-Host "Importing: $($f.FullName)"
    $wb.VBProject.VBComponents.Import($f.FullName)
}

$addInPath = Join-Path $dist 'MyAddin.xlam'
Write-Host "Saving add-in to: $addInPath"
# xlOpenXMLAddIn = 55 (macro-enabled add-in)
$xlOpenXMLAddIn = 55
$wb.SaveAs($addInPath, $xlOpenXMLAddIn) | Out-Null
$wb.Close($false)

# Open the add-in and run tests
$ai = $excel.Workbooks.Open($addInPath)
try {
    Write-Host "Running tests (RunAllTests)..."
    $excel.Run("RunAllTests")
}
catch {
    Write-Host "Error running RunAllTests: $($_.Exception.Message)"
}

# Read test results sheet
try {
    $ws = $ai.Worksheets.Item("TestResults")
    $used = $ws.UsedRange
    $lines = @()
    for ($r = 1; $r -le $used.Rows.Count; $r++) {
        $lines += $ws.Cells.Item($r, 1).Text
    }
    $resultsText = $lines -join "`r`n"
}
catch {
    $resultsText = "No TestResults sheet found or reading failed: $($_.Exception.Message)"
}

$resultsFile = Join-Path $dist 'test-results.txt'
Set-Content -Path $resultsFile -Value $resultsText -Encoding UTF8

# Cleanup
$ai.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Build and tests complete. Results at $resultsFile"
