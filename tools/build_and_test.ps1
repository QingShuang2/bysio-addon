#!/usr/bin/env pwsh
<#
Builds an Excel add-in from exported VBA modules and runs tests.
Requirements:
- Windows with Excel installed
- Excel Trust Center: "Trust access to the VBA project object model" enabled
#>

$ErrorActionPreference = 'Stop'

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory = $true)]
        [string]$EntryPath
    )

    $entry = $Zip.GetEntry($EntryPath)
    if (-not $entry) {
        return $null
    }

    $reader = [System.IO.StreamReader]::new($entry.Open())
    try {
        return $reader.ReadToEnd()
    }
    finally {
        $reader.Dispose()
    }
}

function Set-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip,
        [Parameter(Mandatory = $true)]
        [string]$EntryPath,
        [Parameter(Mandatory = $true)]
        [string]$Content
    )

    $existingEntry = $Zip.GetEntry($EntryPath)
    if ($existingEntry) {
        $existingEntry.Delete()
    }

    $newEntry = $Zip.CreateEntry($EntryPath)
    $writer = [System.IO.StreamWriter]::new($newEntry.Open(), [System.Text.UTF8Encoding]::new($false))
    try {
        $writer.Write($Content)
    }
    finally {
        $writer.Dispose()
    }
}

function Add-CustomRibbonToAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlamPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $customUiPartPath = 'customUI/customUI14.xml'
    $workbookRelsPath = 'xl/_rels/workbook.xml.rels'
    $contentTypesPath = '[Content_Types].xml'
    $ribbonRelationshipType = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'

    $customUiXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabBysioTools" label="Bysio Tools">
        <group id="grpMathLib" label="Math Library">
          <button id="btnAddNumbers"
                  label="Add Numbers"
                  imageMso="FunctionWizard"
                  size="large"
                  onAction="RibbonAddNumbers_OnAction"
                  screentip="Add two numbers"
                  supertip="Prompts for two values and shows the AddNumbers result." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@

    $fileStream = [System.IO.File]::Open($XlamPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
        try {
            Set-ZipEntryText -Zip $zip -EntryPath $customUiPartPath -Content $customUiXml

            $relsText = Get-ZipEntryText -Zip $zip -EntryPath $workbookRelsPath
            if (-not $relsText) {
                throw "Missing workbook relationships entry: $workbookRelsPath"
            }

            [xml]$relsXml = $relsText
            $relsNs = [System.Xml.XmlNamespaceManager]::new($relsXml.NameTable)
            $relsNs.AddNamespace('r', $relsXml.DocumentElement.NamespaceURI)

            $existingRibbonRels = $relsXml.SelectNodes("//r:Relationship[@Type='$ribbonRelationshipType']", $relsNs)
            foreach ($rel in @($existingRibbonRels)) {
                [void]$rel.ParentNode.RemoveChild($rel)
            }

            $existingIds = @($relsXml.SelectNodes('//r:Relationship', $relsNs) | ForEach-Object { $_.Id })
            $newRibbonRelId = 'rIdRibbon'
            $suffix = 1
            while ($existingIds -contains $newRibbonRelId) {
                $suffix++
                $newRibbonRelId = "rIdRibbon$suffix"
            }

            $newRel = $relsXml.CreateElement('Relationship', $relsXml.DocumentElement.NamespaceURI)
            [void]$newRel.SetAttribute('Id', $newRibbonRelId)
            [void]$newRel.SetAttribute('Type', $ribbonRelationshipType)
            [void]$newRel.SetAttribute('Target', '../customUI/customUI14.xml')
            [void]$relsXml.DocumentElement.AppendChild($newRel)

            Set-ZipEntryText -Zip $zip -EntryPath $workbookRelsPath -Content $relsXml.OuterXml

            $contentTypesText = Get-ZipEntryText -Zip $zip -EntryPath $contentTypesPath
            if (-not $contentTypesText) {
                throw "Missing content types entry: $contentTypesPath"
            }

            [xml]$typesXml = $contentTypesText
            $typesNs = [System.Xml.XmlNamespaceManager]::new($typesXml.NameTable)
            $typesNs.AddNamespace('ct', $typesXml.DocumentElement.NamespaceURI)

            $existingOverride = $typesXml.SelectSingleNode("//ct:Override[@PartName='/customUI/customUI14.xml']", $typesNs)
            if (-not $existingOverride) {
                $override = $typesXml.CreateElement('Override', $typesXml.DocumentElement.NamespaceURI)
                [void]$override.SetAttribute('PartName', '/customUI/customUI14.xml')
                [void]$override.SetAttribute('ContentType', 'application/vnd.ms-office.customUI+xml')
                [void]$typesXml.DocumentElement.AppendChild($override)
            }

            Set-ZipEntryText -Zip $zip -EntryPath $contentTypesPath -Content $typesXml.OuterXml
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

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

Write-Host 'Embedding custom ribbon UI into add-in package...'
Add-CustomRibbonToAddIn -XlamPath $addInPath

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
