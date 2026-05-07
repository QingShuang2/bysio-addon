function Invoke-AddInTests {
    param(
        [Parameter(Mandatory = $true)]
        [object]$PowerPoint,
        [Parameter(Mandatory = $true)]
        [string]$AddInPath
    )

    $addIn = $null
    try {
        # .ppam files must be loaded through AddIns.Add, not Presentations.Open.
        $addIn = $PowerPoint.AddIns.Add($AddInPath)
        $addIn.Loaded = $true

        $fileName = [System.IO.Path]::GetFileName($AddInPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($AddInPath)

        $runCandidates = @(
            'RunAllTests',
            "'$fileName'!RunAllTests",
            "'$baseName'!RunAllTests"
        )

        $resultsCandidates = @(
            'GetTestResults',
            "'$fileName'!GetTestResults",
            "'$baseName'!GetTestResults"
        )

        $runSucceeded = $false
        foreach ($macro in $runCandidates) {
            try {
                Write-Host "Running tests ($macro)..."
                $PowerPoint.Run($macro) | Out-Null
                $runSucceeded = $true
                break
            }
            catch {
                Write-Host "Unable to run ${macro}: $($_.Exception.Message)"
            }
        }

        if (-not $runSucceeded) {
            return 'Unable to run tests: no RunAllTests macro entrypoint could be invoked.'
        }

        foreach ($macro in $resultsCandidates) {
            try {
                return [string]($PowerPoint.Run($macro))
            }
            catch {
                Write-Host "Unable to fetch results via ${macro}: $($_.Exception.Message)"
            }
        }

        return 'No test results available: GetTestResults macro could not be invoked.'
    }
    finally {
        if ($addIn) {
            try {
                $addIn.Loaded = $false
            }
            catch {
                # Best effort cleanup.
            }
        }
    }
}
