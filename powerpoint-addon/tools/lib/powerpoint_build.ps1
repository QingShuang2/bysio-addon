function New-PowerPointApplication {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $true
    return $ppt
}

function Import-VbaModules {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Presentation,
        [Parameter(Mandatory = $true)]
        [string]$VbaDirectory
    )

    if (-not $Presentation) {
        throw 'Presentation object is null. Unable to import VBA modules.'
    }

    $vbProject = $null
    try {
        $vbProject = $Presentation.VBProject
    }
    catch {
        throw 'Unable to access Presentation.VBProject. Enable "Trust access to the VBA project object model" in PowerPoint Trust Center.'
    }

    if (-not $vbProject) {
        throw 'Presentation.VBProject is unavailable. Enable "Trust access to the VBA project object model" in PowerPoint Trust Center.'
    }

    Write-Host "Importing VBA modules from: $VbaDirectory"
    foreach ($moduleFile in Get-ChildItem -Path $VbaDirectory -Filter *.bas -ErrorAction SilentlyContinue) {
        Write-Host "Importing: $($moduleFile.FullName)"
        $vbProject.VBComponents.Import($moduleFile.FullName)
    }
}

function Save-PresentationAsAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Presentation,
        [Parameter(Mandatory = $true)]
        [string]$DistDirectory
    )

    $buildName = 'MyPowerPointAddIn'
    try {
        $envFile = Join-Path (Split-Path -Parent $PSScriptRoot) '.env'
        if (Test-Path $envFile) {
            $envLines = Get-Content -Path $envFile -ErrorAction SilentlyContinue
            foreach ($line in $envLines) {
                if ($line -match '^\s*build_name\s*=\s*(.+)') {
                    $val = $matches[1].Trim()
                    $val = $val.Trim([char[]]@("'", '"'))
                    if ($val -ne '') {
                        $buildName = $val
                        break
                    }
                }
            }
        }
    }
    catch {
        # Ignore parse errors and use default.
    }

    $addInPath = Join-Path $DistDirectory ($buildName + '.ppam')
    Write-Host "Saving add-in to: $addInPath"

    try {
        $Presentation.SaveAs($addInPath) | Out-Null
        return $addInPath
    }
    catch {
        $fallbackPath = Join-Path $DistDirectory ($buildName + '-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '.ppam')
        Write-Host "Primary add-in file is locked. Saving to fallback path: $fallbackPath"
        $Presentation.SaveAs($fallbackPath) | Out-Null
        return $fallbackPath
    }
}

function Load-AddInIntoPowerPoint {
    param(
        [Parameter(Mandatory = $true)]
        [object]$PowerPoint,
        [Parameter(Mandatory = $true)]
        [string]$AddInPath
    )

    $resolvedPath = [System.IO.Path]::GetFullPath($AddInPath)

    # Unload duplicate entries to avoid stale builds remaining active.
    foreach ($item in @($PowerPoint.AddIns)) {
        try {
            $itemPath = [System.IO.Path]::GetFullPath([string]$item.FullName)
            if ($itemPath -ieq $resolvedPath) {
                $item.Loaded = $false
            }
        }
        catch {
            # Ignore entries that cannot be inspected.
        }
    }

    $addIn = $PowerPoint.AddIns.Add($resolvedPath)
    $addIn.Loaded = $true
}
