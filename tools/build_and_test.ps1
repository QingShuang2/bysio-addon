#!/usr/bin/env pwsh
<#
Builds an Excel add-in from exported VBA modules and runs tests.
Requirements:
- Windows with Excel installed
- Excel Trust Center: "Trust access to the VBA project object model" enabled
#>

$ErrorActionPreference = 'Stop'

$RibbonSettings = @{
    CustomUiPartPath         = 'customUI/customUI.xml'
    LegacyCustomUiPartPath   = 'customUI/customUI14.xml'
    PackageRelsPath          = '_rels/.rels'
    WorkbookRelsPath         = 'xl/_rels/workbook.xml.rels'
    ContentTypesPath         = '[Content_Types].xml'
    RelationshipType         = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    CustomUiContentType      = 'application/vnd.ms-office.customUI+xml'
    PackageRelationshipPath  = 'customUI/customUI.xml'
    WorkbookRelationshipPath = '../customUI/customUI.xml'
}

function Get-CustomUiXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
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
}

# OpenXML helpers
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

function Set-RibbonRelationship {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$RelationshipsXml,
        [Parameter(Mandatory = $true)]
        [string]$RelationshipType,
        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    $relsNs = [System.Xml.XmlNamespaceManager]::new($RelationshipsXml.NameTable)
    $relsNs.AddNamespace('r', $RelationshipsXml.DocumentElement.NamespaceURI)

    $existingRibbonRels = $RelationshipsXml.SelectNodes("//r:Relationship[@Type='$RelationshipType']", $relsNs)
    foreach ($rel in @($existingRibbonRels)) {
        [void]$rel.ParentNode.RemoveChild($rel)
    }

    $existingIds = @($RelationshipsXml.SelectNodes('//r:Relationship', $relsNs) | ForEach-Object { $_.Id })
    $newRibbonRelId = 'rIdRibbon'
    $suffix = 1
    while ($existingIds -contains $newRibbonRelId) {
        $suffix++
        $newRibbonRelId = "rIdRibbon$suffix"
    }

    $newRel = $RelationshipsXml.CreateElement('Relationship', $RelationshipsXml.DocumentElement.NamespaceURI)
    [void]$newRel.SetAttribute('Id', $newRibbonRelId)
    [void]$newRel.SetAttribute('Type', $RelationshipType)
    [void]$newRel.SetAttribute('Target', $Target)
    [void]$RelationshipsXml.DocumentElement.AppendChild($newRel)
}

function Update-RibbonRelationships {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip
    )

    $packageRelsText = Get-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.PackageRelsPath
    if (-not $packageRelsText) {
        throw "Missing package relationships entry: $($RibbonSettings.PackageRelsPath)"
    }

    [xml]$packageRelsXml = $packageRelsText
    Set-RibbonRelationship -RelationshipsXml $packageRelsXml -RelationshipType $RibbonSettings.RelationshipType -Target $RibbonSettings.PackageRelationshipPath
    Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.PackageRelsPath -Content $packageRelsXml.OuterXml

    $workbookRelsText = Get-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.WorkbookRelsPath
    if ($workbookRelsText) {
        [xml]$workbookRelsXml = $workbookRelsText
        Set-RibbonRelationship -RelationshipsXml $workbookRelsXml -RelationshipType $RibbonSettings.RelationshipType -Target $RibbonSettings.WorkbookRelationshipPath
        Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.WorkbookRelsPath -Content $workbookRelsXml.OuterXml
    }
}

function Update-ContentTypes {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip
    )

    $contentTypesText = Get-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.ContentTypesPath
    if (-not $contentTypesText) {
        throw "Missing content types entry: $($RibbonSettings.ContentTypesPath)"
    }

    [xml]$typesXml = $contentTypesText
    $typesNs = [System.Xml.XmlNamespaceManager]::new($typesXml.NameTable)
    $typesNs.AddNamespace('ct', $typesXml.DocumentElement.NamespaceURI)

    $legacyOverride = $typesXml.SelectSingleNode("//ct:Override[@PartName='/$($RibbonSettings.LegacyCustomUiPartPath)']", $typesNs)
    if ($legacyOverride) {
        [void]$legacyOverride.ParentNode.RemoveChild($legacyOverride)
    }

    $currentOverride = $typesXml.SelectSingleNode("//ct:Override[@PartName='/$($RibbonSettings.CustomUiPartPath)']", $typesNs)
    if (-not $currentOverride) {
        $override = $typesXml.CreateElement('Override', $typesXml.DocumentElement.NamespaceURI)
        [void]$override.SetAttribute('PartName', "/$($RibbonSettings.CustomUiPartPath)")
        [void]$override.SetAttribute('ContentType', $RibbonSettings.CustomUiContentType)
        [void]$typesXml.DocumentElement.AppendChild($override)
    }

    Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.ContentTypesPath -Content $typesXml.OuterXml
}

function Add-CustomRibbonToAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$XlamPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $fileStream = [System.IO.File]::Open($XlamPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
        try {
            $legacyPart = $zip.GetEntry($RibbonSettings.LegacyCustomUiPartPath)
            if ($legacyPart) {
                $legacyPart.Delete()
            }

            Set-ZipEntryText -Zip $zip -EntryPath $RibbonSettings.CustomUiPartPath -Content (Get-CustomUiXml)
            Update-RibbonRelationships -Zip $zip
            Update-ContentTypes -Zip $zip
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

# Excel automation helpers
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

    $workbook.IsAddin = $true
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

# Entrypoint
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
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
