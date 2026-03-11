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
                <tab id="tabBysioTools" label="Bysio">
                <group id="grpFontPicker" label="Font">
                    <dropDown id="ddlRibbonFont"
                              label="Font"
                              getSelectedItemIndex="RibbonFont_GetSelectedItemIndex"
                              onAction="RibbonFont_OnAction"
                              screentip="Select font for testing">
                        <item id="font_msgothic" label="ＭＳ ゴシック" />
                        <item id="font_meiryo" label="Meiryo UI" />
                    </dropDown>
                    <editBox id="txtRibbonSize"
                             label="Size"
                             getText="RibbonSize_GetText"
                             onChange="RibbonSize_OnChange"
                             screentip="Font size (default 11)" />
                    <checkBox id="chkRibbonAllSheets"
                              label="All Sheets?"
                              getPressed="RibbonAllSheets_GetPressed"
                              onAction="RibbonAllSheets_OnAction"
                              screentip="Apply to all sheets when checked" />
                </group>

                <group id="grpApplyFont" label="Apply">
                    <button id="btnApplyFont"
                        label="Apply Font And Size"
                        imageMso="FontDialog"
                        size="large"
                        onAction="RibbonApplyFont_OnAction"
                        screentip="Apply selected font and size to the active sheet" />
                </group>

                <group id="grpZoom" label="Zoom">
                    <button id="btnZoom100"
                        label="Zoom 100%"
                        imageMso="Zoom100"
                        size="large"
                        onAction="RibbonZoom100_OnAction"
                        screentip="Set zoom to 100%" />
                    <checkBox id="chkZoomAllSheets"
                              label="All Sheets?"
                              getPressed="RibbonZoomAllSheets_GetPressed"
                              onAction="RibbonZoomAllSheets_OnAction"
                              screentip="Apply zoom changes to all sheets when checked" />
                    <box id="boxZoomBtns" boxStyle="horizontal">
                        <button id="btnZoomUp"
                            label="Zoom +"
                            imageMso="ZoomIn"
                            onAction="RibbonZoomUp_OnAction"
                            screentip="Increase zoom by 10%" />
                        <button id="btnZoomDown"
                            label="Zoom -"
                            imageMso="ZoomOut"
                            onAction="RibbonZoomDown_OnAction"
                            screentip="Decrease zoom by 10%" />
                    </box>
                </group>

                <group id="grpResize" label="Picture">
                    <button id="btnResizePicture"
                        label="Resize 100%"
                        imageMso="PictureCrop"
                        size="large"
                        onAction="RibbonResizePicture_OnAction"
                        screentip="Resize pictures to configured percent" />
                    <checkBox id="chkResizeAllSheets"
                              label="All Sheets?"
                              getPressed="RibbonResizeAllSheets_GetPressed"
                              onAction="RibbonResizeAllSheets_OnAction"
                              screentip="Apply resize changes to all sheets when checked" />
                    <box id="boxResizeBtns" boxStyle="horizontal">
                        <button id="btnResizeUp"
                            label="Resize +"
                            imageMso="ZoomIn"
                            onAction="RibbonResizeUp_OnAction"
                            screentip="Increase picture size by 5%" />
                        <button id="btnResizeDown"
                            label="Resize -"
                            imageMso="ZoomOut"
                            onAction="RibbonResizeDown_OnAction"
                            screentip="Decrease picture size by 5%" />
                    </box>
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
                <tab id="tabBysioTools" label="Bysio">
                <group id="grpFont" label="Font">
                    <button id="btnApplyFont"
                        label="Apply Font And Size"
                        imageMso="FontDialog"
                        size="large"
                        onAction="RibbonApplyFont_OnAction"
                        screentip="Apply selected font and size to every worksheet" />
                    <checkBox id="chkRibbonAllSheets"
                              label="All Sheets?"
                              getPressed="RibbonAllSheets_GetPressed"
                              onAction="RibbonAllSheets_OnAction"
                              screentip="Apply to all sheets when checked" />
                    <dropDown id="ddlRibbonFont"
                              label="Font"
                              getSelectedItemIndex="RibbonFont_GetSelectedItemIndex"
                              onAction="RibbonFont_OnAction"
                              screentip="Select font for testing">
                        <item id="font_msgothic" label="ＭＳ ゴシック" />
                        <item id="font_meiryo" label="Meiryo UI" />
                    </dropDown>
                    <editBox id="txtRibbonSize"
                             label="Size"
                             getText="RibbonSize_GetText"
                             onChange="RibbonSize_OnChange"
                             screentip="Font size (default 11)" />
                </group>
                <group id="grpZoom" label="Zoom">
                    <button id="btnZoom100"
                        label="Zoom 100%"
                        imageMso="Zoom100"
                        size="large"
                        onAction="RibbonZoom100_OnAction"
                        screentip="Set zoom to 100%" />

                    <box id="boxZoomBtns" boxStyle="vertical">
                        <checkBox id="chkZoomAllSheets"
                            label="All Sheets?"
                            getPressed="RibbonZoomAllSheets_GetPressed"
                            onAction="RibbonZoomAllSheets_OnAction"
                            screentip="Apply zoom changes to all sheets when checked" />
                        <button id="btnZoomUp"
                            label="Zoom +"
                            imageMso="ZoomIn"
                            onAction="RibbonZoomUp_OnAction"
                            screentip="Increase zoom by 10%" />
                        <button id="btnZoomDown"
                            label="Zoom -"
                            imageMso="ZoomOut"
                            onAction="RibbonZoomDown_OnAction"
                            screentip="Decrease zoom by 10%" />
                    </box>
                </group>
                <group id="grpResize" label="Picture">
                    <button id="btnResizePicture"
                        label="Resize 100%"
                        imageMso="PictureCrop"
                        size="large"
                        onAction="RibbonResizePicture_OnAction"
                        screentip="Resize pictures to configured percent" />
                    <box id="boxResizeBtns" boxStyle="vertical">
                        <checkBox id="chkResizeAllSheets"
                            label="All Sheets?"
                            getPressed="RibbonResizeAllSheets_GetPressed"
                            onAction="RibbonResizeAllSheets_OnAction"
                            screentip="Apply resize changes to all sheets when checked" />
                        <button id="btnResizeUp"
                            label="Resize +"
                            imageMso="ZoomIn"
                            onAction="RibbonResizeUp_OnAction"
                            screentip="Increase picture size by 5%" />
                        <button id="btnResizeDown"
                            label="Resize -"
                            imageMso="ZoomOut"
                            onAction="RibbonResizeDown_OnAction"
                            screentip="Decrease picture size by 5%" />
                    </box>
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
    Set-PackageRelationship -RelationshipsXml $packageRelsXml -RelationshipType $RibbonSettings.CustomUiRelationshipType -Target $RibbonSettings.CustomUiPartPath
    Set-PackageRelationship -RelationshipsXml $packageRelsXml -RelationshipType $RibbonSettings.CustomUi14RelationshipType -Target $RibbonSettings.CustomUi14PartPath
    Set-ZipEntryText -Zip $Zip -EntryPath $RibbonSettings.PackageRelsPath -Content $packageRelsXml.OuterXml
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

    $currentOverride = $typesXml.SelectSingleNode("//ct:Override[@PartName='/$($RibbonSettings.CustomUiPartPath)']", $typesNs)
    if ($currentOverride) {
        [void]$currentOverride.ParentNode.RemoveChild($currentOverride)
    }

    $currentOverride14 = $typesXml.SelectSingleNode("//ct:Override[@PartName='/$($RibbonSettings.CustomUi14PartPath)']", $typesNs)
    if ($currentOverride14) {
        [void]$currentOverride14.ParentNode.RemoveChild($currentOverride14)
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
