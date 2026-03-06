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
    # Determine build name from tools\.env (parent of this lib folder)
    $defaultName = 'MyAddin'
    try {
        $envFile = Join-Path (Split-Path -Parent $PSScriptRoot) '.env'
        if (Test-Path $envFile) {
            $envLines = Get-Content -Path $envFile -ErrorAction SilentlyContinue
            foreach ($line in $envLines) {
                if ($line -match '^\s*build_name\s*=\s*(.+)') {
                    $val = $matches[1].Trim()
                    # Strip surrounding single or double quotes if present
                    $val = $val.Trim([char[]]@("'", '"'))
                    if ($val -ne '') { $defaultName = $val; break }
                }
            }
        }
    }
    catch {
        # ignore parse errors and use default
    }

    $addInFileName = $defaultName + '.xlam'
    $addInPath = Join-Path $DistDirectory $addInFileName
    Write-Host "Saving add-in to: $addInPath"

    try {
        $Workbook.SaveAs($addInPath, $xlOpenXMLAddIn) | Out-Null
        return $addInPath
    }
    catch {
        $fallbackPath = Join-Path $DistDirectory ($defaultName + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.xlam')
        Write-Host "Primary add-in file is locked. Saving to fallback path: $fallbackPath"
        $Workbook.SaveAs($fallbackPath, $xlOpenXMLAddIn) | Out-Null
        return $fallbackPath
    }
}
