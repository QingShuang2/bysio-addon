function Read-TestResults {
    param(
        [Parameter(Mandatory = $true)]
        [object]$AddInWorkbook
    )

    try {
        $ws = $AddInWorkbook.Worksheets.Item('TestResults')
        $used = $ws.UsedRange
        $lines = @()
        for ($r = 1; $r -le $used.Rows.Count; $r++) {
            $lines += $ws.Cells.Item($r, 1).Text
        }
        return $lines -join "`r`n"
    }
    catch {
        return "No TestResults sheet found or reading failed: $($_.Exception.Message)"
    }
}

function Invoke-AddInTests {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Excel,
        [Parameter(Mandatory = $true)]
        [string]$AddInPath
    )

    $addInWorkbook = $Excel.Workbooks.Open($AddInPath)
    try {
        try {
            Write-Host 'Running tests (RunAllTests)...'
            $Excel.Run('RunAllTests')
        }
        catch {
            Write-Host "Error running RunAllTests: $($_.Exception.Message)"
        }

        return Read-TestResults -AddInWorkbook $addInWorkbook
    }
    finally {
        $addInWorkbook.Close($false)
    }
}
