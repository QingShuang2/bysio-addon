function New-ExcelApplication {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 1
    return $excel
}

function Import-VbaModules {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,
        [Parameter(Mandatory = $true)]
        [string]$VbaDirectory
    )

    Write-Host "Importing VBA modules from: $VbaDirectory"
    foreach ($moduleFile in Get-ChildItem -Path $VbaDirectory -Filter *.bas -ErrorAction SilentlyContinue) {
        Write-Host "Importing: $($moduleFile.FullName)"
        $Workbook.VBProject.VBComponents.Import($moduleFile.FullName)
    }
}

function Save-WorkbookAsAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Workbook,
        [Parameter(Mandatory = $true)]
        [string]$DistDirectory
    )

    $Workbook.IsAddin = $true
    $xlOpenXMLAddIn = 55
    $addInPath = Join-Path $DistDirectory 'MyAddin.xlam'
    Write-Host "Saving add-in to: $addInPath"

    try {
        $Workbook.SaveAs($addInPath, $xlOpenXMLAddIn) | Out-Null
        return $addInPath
    }
    catch {
        $fallbackPath = Join-Path $DistDirectory ("MyAddin-{0}.xlam" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        Write-Host "Primary add-in file is locked. Saving to fallback path: $fallbackPath"
        $Workbook.SaveAs($fallbackPath, $xlOpenXMLAddIn) | Out-Null
        return $fallbackPath
    }
}
