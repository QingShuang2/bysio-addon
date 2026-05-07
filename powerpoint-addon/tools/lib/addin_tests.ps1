function Invoke-AddInTests {
    param(
        [Parameter(Mandatory = $true)]
        [object]$PowerPoint,
        [Parameter(Mandatory = $true)]
        [string]$AddInPath
    )

    $addInPresentation = $PowerPoint.Presentations.Open($AddInPath, $false, $false, $false)
    try {
        try {
            Write-Host 'Running tests (RunAllTests)...'
            $PowerPoint.Run('RunAllTests') | Out-Null
        }
        catch {
            Write-Host "Error running RunAllTests: $($_.Exception.Message)"
        }

        try {
            return [string]($PowerPoint.Run('GetTestResults'))
        }
        catch {
            return "No test results available: $($_.Exception.Message)"
        }
    }
    finally {
        $addInPresentation.Close()
    }
}
