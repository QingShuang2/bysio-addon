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
                                <group id="grpFormatting" label="Formatting">
                                        <button id="btnApplyFont"
                                                                        label="Apply Font"
                                                                        imageMso="FontDialog"
                                                                        size="large"
                                                                        onAction="RibbonApplyFont_OnAction"
                                                                        screentip="Apply font to all sheets"
                                                                        supertip="Applies the configured font to all worksheets (skips names starting/ending with _)." />
                                        <button id="btnFormatNumbers"
                                                                        label="Format Numbers"
                                                                        imageMso="NumberFormat"
                                                                        size="large"
                                                                        onAction="RibbonFormatNumbers_OnAction"
                                                                        screentip="Format selected numeric cells"
                                                                        supertip="Converts numeric-looking values to Number format; zeros get gray background, positives get red font." />
                                        <button id="btnZoom100"
                                                                        label="Zoom 100%"
                                                                        imageMso="ZoomTo100Percent"
                                                                        size="large"
                                                                        onAction="RibbonZoom100_OnAction"
                                                                        screentip="Set view zoom to 100% for all sheets"
                                                                        supertip="Sets the view zoom to 100% for all worksheets, skipping sheets whose name starts or ends with an underscore." />
                                        <button id="btnResizePicture"
                                                                        label="Resize to 70%"
                                                                        imageMso="FormatPicture"
                                                                        size="large"
                                                                        onAction="RibbonResizePicture_OnAction"
                                                                        screentip="Resize selected pictures to 70%"
                                                                        supertip="Resize selected picture(s) to 70% of their original size, preserving aspect ratio." />
                                </group>
                        </tab>
        </tabs>
    </ribbon>
</customUI>
'@
}

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
