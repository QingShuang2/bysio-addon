$RibbonSettings = @{
    CustomUiPartPath           = 'customUI/customUI.xml'
    CustomUi14PartPath         = 'customUI/customUI14.xml'
    PackageRelsPath            = '_rels/.rels'
    ContentTypesPath           = '[Content_Types].xml'
    CustomUiRelationshipType   = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    CustomUi14RelationshipType = 'http://schemas.microsoft.com/office/2007/relationships/ui/extensibility'
}

function Get-CustomUiXml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui" onLoad="RibbonOnLoad">
    <ribbon>
        <tabs>
            <tab id="tabBysioPpt" label="Bysio">
                <group id="grpResize" label="Resize Image">
                    <button id="btnResizeImage"
                            label="Resize Image"
                            size="large"
                            onAction="RibbonResizeImage_OnAction"
                            screentip="Resize selected image to specified dimensions" />
                    <editBox id="txtScale"
                             label="Scale %"
                             getText="RibbonScale_GetText"
                             onChange="RibbonScale_OnChange"
                             screentip="Resize based on 5.85in x 9.1in slide size (default 64)" />
                    <editBox id="txtHorizontal"
                             label="Horizontal (in)"
                             getText="RibbonHorizontal_GetText"
                             onChange="RibbonHorizontal_OnChange"
                             screentip="Left position in inches from slide left edge" />
                    <editBox id="txtVertical"
                             label="Vertical (in)"
                             getText="RibbonVertical_GetText"
                             onChange="RibbonVertical_OnChange"
                             screentip="Top position in inches from slide top edge" />
                </group>
            </tab>
        </tabs>
    </ribbon>
</customUI>
'@
}

function Get-CustomUi14Xml {
    return @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonOnLoad">
    <ribbon>
        <tabs>
            <tab id="tabBysioPpt" label="Bysio">
                <group id="grpResize" label="Resize Image">
                    <button id="btnResizeImage"
                            label="Resize Image"
                            size="large"
                            onAction="RibbonResizeImage_OnAction"
                            screentip="Resize selected image to specified dimensions" />
                    <editBox id="txtScale"
                             label="Scale %"
                             getText="RibbonScale_GetText"
                             onChange="RibbonScale_OnChange"
                             screentip="Resize based on 5.85in x 9.1in slide size (default 64)" />
                    <editBox id="txtHorizontal"
                             label="Horizontal (in)"
                             getText="RibbonHorizontal_GetText"
                             onChange="RibbonHorizontal_OnChange"
                             screentip="Left position in inches from slide left edge" />
                    <editBox id="txtVertical"
                             label="Vertical (in)"
                             getText="RibbonVertical_GetText"
                             onChange="RibbonVertical_OnChange"
                             screentip="Top position in inches from slide top edge" />
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

function Set-PackageRelationship {
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

    $existing = $RelationshipsXml.SelectNodes("//r:Relationship[@Type='$RelationshipType']", $relsNs)
    foreach ($rel in @($existing)) {
        [void]$rel.ParentNode.RemoveChild($rel)
    }

    $existingIds = @($RelationshipsXml.SelectNodes('//r:Relationship', $relsNs) | ForEach-Object { $_.Id })
    $newId = 'rIdRibbon'
    $suffix = 1
    while ($existingIds -contains $newId) {
        $suffix++
        $newId = "rIdRibbon$suffix"
    }

    $newRel = $RelationshipsXml.CreateElement('Relationship', $RelationshipsXml.DocumentElement.NamespaceURI)
    [void]$newRel.SetAttribute('Id', $newId)
    [void]$newRel.SetAttribute('Type', $RelationshipType)
    [void]$newRel.SetAttribute('Target', $Target)
    [void]$RelationshipsXml.DocumentElement.AppendChild($newRel)
}

function Remove-ContentTypeOverride {
    param(
        [Parameter(Mandatory = $true)]
        [xml]$TypesXml,
        [Parameter(Mandatory = $true)]
        [string]$PartPath
    )

    $typesNs = [System.Xml.XmlNamespaceManager]::new($TypesXml.NameTable)
    $typesNs.AddNamespace('ct', $TypesXml.DocumentElement.NamespaceURI)

    $existing = $TypesXml.SelectSingleNode("//ct:Override[@PartName='/$PartPath']", $typesNs)
    if ($existing) {
        [void]$existing.ParentNode.RemoveChild($existing)
    }
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
    Set-PackageRelationship -RelationshipsXml $packageRelsXml -RelationshipType $RibbonSettings.CustomUiRelationshipType -Target $RibbonSettings.CustomUiPartPath
    Set-PackageRelationship -RelationshipsXml $packageRelsXml -RelationshipType $RibbonSettings.CustomUi14RelationshipType -Target $RibbonSettings.CustomUi14PartPath
    Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.PackageRelsPath -Content $packageRelsXml.OuterXml
}

function Update-ContentTypes {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.Compression.ZipArchive]$Zip
    )

    $typesText = Get-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.ContentTypesPath
    if (-not $typesText) {
        throw "Missing content types entry: $($RibbonSettings.ContentTypesPath)"
    }

    [xml]$typesXml = $typesText
    Remove-ContentTypeOverride -TypesXml $typesXml -PartPath $RibbonSettings.CustomUiPartPath
    Remove-ContentTypeOverride -TypesXml $typesXml -PartPath $RibbonSettings.CustomUi14PartPath
    Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.ContentTypesPath -Content $typesXml.OuterXml
}

function Add-CustomRibbonToAddIn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PpamPath
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $fileStream = [System.IO.File]::Open($PpamPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zip = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
        try {
            Set-ZipEntryText -Zip $zip -EntryPath $RibbonSettings.CustomUiPartPath -Content (Get-CustomUiXml)
            Set-ZipEntryText -Zip $zip -EntryPath $RibbonSettings.CustomUi14PartPath -Content (Get-CustomUi14Xml)
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
