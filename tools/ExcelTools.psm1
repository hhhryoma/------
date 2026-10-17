#     tools 以下のスクリプトで共通に使う部品。Excel ファイル（Open XML）を Excel なしで読み書きする。
#
#     使い方（スクリプトの先頭で）:
#         Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
#
#     このファイルは UTF-8（BOM 付き）で保存すること（Windows PowerShell 5.1 が日本語を正しく読むため）。

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml

$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 外部実体の読み込みを禁止して XML を読む（XXE 対策）
$script:XmlSettings = New-Object System.Xml.XmlReaderSettings
$script:XmlSettings.DtdProcessing  = [System.Xml.DtdProcessing]::Prohibit
$script:XmlSettings.XmlResolver    = $null
$script:XmlSettings.IgnoreComments = $true
$script:XmlSettings.CloseInput     = $true

# ---------------------------------------------------------------- 共通

function ConvertTo-ColumnName {
    # 1 始まりの列番号を A, B, ..., Z, AA, ... に変換する
    param([int] $Index)

    $name = ''
    while ($Index -gt 0) {
        $m = ($Index - 1) % 26
        $name = [string][char](65 + $m) + $name
        $Index = [int][math]::Floor(($Index - 1) / 26)
    }
    return $name
}

function Get-ColumnIndex {
    # "AB12" のようなセル番地から列番号（1 始まり）を取り出す
    param([string] $Ref)

    $col = 0
    foreach ($ch in $Ref.ToUpperInvariant().ToCharArray()) {
        if ($ch -lt 'A' -or $ch -gt 'Z') { break }
        $col = $col * 26 + ([int]$ch - 64)
    }
    return $col
}

function ConvertTo-XmlText {
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    $Text = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', '')
    return [System.Security.SecurityElement]::Escape($Text)
}

# ---------------------------------------------------------------- 読み取り（xlsx / xlsm の ZIP パッケージ）

function Open-ZipPackage {
    param([string] $FilePath)

    # Excel で開いている最中のファイルも読めるよう、共有モードで開く
    $share  = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $stream = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
    }
    catch {
        $stream.Dispose()
        throw
    }

    # パーツ名の大文字小文字が規定と違うファイルがあるため、小文字で引けるようにしておく
    $entries = @{}
    foreach ($e in $zip.Entries) { $entries[$e.FullName.ToLowerInvariant()] = $e }

    return [pscustomobject]@{ Zip = $zip; Entries = $entries }
}

function Get-PackageEntry {
    param($Package, [string] $PartName)
    return $Package.Entries[$PartName.TrimStart('/').ToLowerInvariant()]
}

function New-PartXmlReader {
    param($Package, [string] $PartName)

    $entry = Get-PackageEntry -Package $Package -PartName $PartName
    if ($null -eq $entry) { return $null }
    return [System.Xml.XmlReader]::Create($entry.Open(), $script:XmlSettings)
}

function Read-PartXml {
    param($Package, [string] $PartName)

    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return $null }
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.Load($reader)
        return $doc
    }
    finally {
        $reader.Dispose()
    }
}

function Resolve-PartName {
    # リレーションの Target（相対パス）をパッケージ内の絶対パーツ名にする
    param([string] $SourcePart, [string] $TargetPath)

    if ($TargetPath.StartsWith('/')) { return $TargetPath.TrimStart('/') }

    $dir = ''
    $slash = $SourcePart.LastIndexOf('/')
    if ($slash -ge 0) { $dir = $SourcePart.Substring(0, $slash) }

    $parts = New-Object System.Collections.Generic.List[string]
    if ($dir) { foreach ($seg in $dir.Split('/')) { $parts.Add($seg) } }
    foreach ($seg in $TargetPath.Split('/')) {
        if ($seg -eq '..') { if ($parts.Count -gt 0) { $parts.RemoveAt($parts.Count - 1) } }
        elseif ($seg -ne '.' -and $seg -ne '') { $parts.Add($seg) }
    }
    return ($parts -join '/')
}

function Get-PartRelationships {
    #     パーツのリレーション（_rels/*.rels）を読み、Id / 種類 / 解決済みパーツ名を返す
    param($Package, [string] $PartName)

    $slash = $PartName.LastIndexOf('/')
    if ($slash -ge 0) {
        $relsName = $PartName.Substring(0, $slash) + '/_rels/' + $PartName.Substring($slash + 1) + '.rels'
    }
    else {
        $relsName = '_rels/' + $PartName + '.rels'
    }

    $doc = Read-PartXml -Package $Package -PartName $relsName
    $list = New-Object System.Collections.Generic.List[object]
    if ($null -eq $doc) { return ,$list }

    foreach ($rel in $doc.SelectNodes("//*[local-name()='Relationship']")) {
        if ($rel.GetAttribute('TargetMode') -eq 'External') { continue }
        $type = $rel.GetAttribute('Type')
        $list.Add([pscustomobject]@{
            Id     = $rel.GetAttribute('Id')
            Kind   = $type.Substring($type.LastIndexOf('/') + 1)
            Target = Resolve-PartName -SourcePart $PartName -TargetPath $rel.GetAttribute('Target')
        })
    }
    return ,$list
}

function Get-RelationshipIdAttribute {
    # r:id 属性を名前空間（Transitional / Strict）によらず取り出す
    param([System.Xml.XmlElement] $Element)

    foreach ($a in $Element.Attributes) {
        if ($a.LocalName -eq 'id' -and $a.NamespaceURI -like '*relationships') { return $a.Value }
    }
    return $null
}

# ---------------------------------------------------------------- 書き出し（xlsx）


# 部分ごとに書式を付けた文字列のフォント（削除は赤の取り消し線、追加は赤の太字・下線）
$runFont = '<sz val="10"/>{0}<rFont val="游ゴシック"/><family val="3"/><charset val="128"/>'
$script:RichFonts = @{
    ''         = $runFont -f ''
    'Deleted'  = '<strike/>' + ($runFont -f '<color rgb="FFC00000"/>')
    'Inserted' = '<b/><u/>' + ($runFont -f '<color rgb="FFC00000"/>')
}

function ConvertTo-InlineStringXml {
    #     セルの文字列を <is> 要素にする（空なら ''）。
    #     @{ Runs = @(@{ Text; Style }) } を渡すと、部分ごとに書式を付ける（Style: Deleted / Inserted / なし）
    param($Value)

    if ($Value -is [System.Collections.IDictionary] -and $Value.Contains('Runs')) {
        # 書式が同じ隣どうしの部分はつなげる
        $parts = New-Object System.Collections.Generic.List[object]
        foreach ($run in $Value.Runs) {
            $t = [string]$run.Text
            if ($t.Length -eq 0) { continue }
            $st = [string]$run.Style
            if (-not $script:RichFonts.ContainsKey($st)) { $st = '' }
            if ($parts.Count -gt 0 -and $parts[$parts.Count - 1][0] -eq $st) { $parts[$parts.Count - 1][1] += $t }
            else { $parts.Add(@($st, $t)) }
        }
        if ($parts.Count -eq 0) { return '' }
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<is>')
        foreach ($part in $parts) {
            [void]$sb.Append('<r><rPr>').Append($script:RichFonts[$part[0]]).Append('</rPr><t xml:space="preserve">').Append((ConvertTo-XmlText $part[1])).Append('</t></r>')
        }
        return $sb.Append('</is>').ToString()
    }
    $v = [string]$Value
    if ($v.Length -eq 0) { return '' }
    if ($v.Length -gt 32767) { $v = $v.Substring(0, 32767) }
    return ('<is><t xml:space="preserve">{0}</t></is>' -f (ConvertTo-XmlText $v))
}

function Export-Workbook {
    #     複数シートの表を xlsx に書き出す（Excel 不要）。各シートは見出し行の固定とオートフィルター付き。
    #     Sheets: @(@{ Name; Columns = @(@{ Header; Width }); Rows = @(@{ Values = @(...); Link = @{ Column; Path; Location; Tooltip } }) })
    #       Values の数値（int / long / double）は数値のセルとして書く
    #       Values に @{ Runs = @(@{ Text; Style }) } を渡すと、部分ごとに書式を付けた文字列にする（Style: Deleted / Inserted / なし）
    #       Link は Values の Column 番目（0 始まり）のセルに付ける外部リンク。Path のファイルの Location（'シート'!セル）を開く
    param($Sheets, [string] $FilePath)

    $font = '<sz val="10"/><name val="游ゴシック"/><family val="3"/><charset val="128"/>'
    $sheetEntries = New-Object System.Collections.Generic.List[object]
    $definedNames = New-Object System.Text.StringBuilder

    $index = 0
    foreach ($s in $Sheets) {
        $index++
        $colCount = $s.Columns.Count
        $lastCol  = ConvertTo-ColumnName $colCount
        $rowCount = $s.Rows.Count + 1

        $xml   = New-Object System.Text.StringBuilder
        $rels  = New-Object System.Text.StringBuilder
        $links = New-Object System.Text.StringBuilder
        [void]$xml.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        [void]$xml.Append(('<dimension ref="A1:{0}{1}"/>' -f $lastCol, $rowCount))
        $tab = ''
        if ($index -eq 1) { $tab = ' tabSelected="1"' }
        [void]$xml.Append(('<sheetViews><sheetView{0} workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' -f $tab))
        [void]$xml.Append('<sheetFormatPr defaultRowHeight="18"/><cols>')
        for ($i = 0; $i -lt $colCount; $i++) {
            [void]$xml.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $s.Columns[$i].Width))
        }
        [void]$xml.Append('</cols><sheetData><row r="1">')
        for ($i = 0; $i -lt $colCount; $i++) {
            [void]$xml.Append(('<c r="{0}1" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f (ConvertTo-ColumnName ($i + 1)), (ConvertTo-XmlText $s.Columns[$i].Header)))
        }
        [void]$xml.Append('</row>')
        [void]$rels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

        $r = 1; $linkCount = 0
        foreach ($row in $s.Rows) {
            $r++
            [void]$xml.Append(('<row r="{0}">' -f $r))
            for ($i = 0; $i -lt $colCount; $i++) {
                $raw = $row.Values[$i]
                if ($raw -is [int] -or $raw -is [long] -or $raw -is [double]) {
                    # 数値は数値のセルとして書く（文字列にすると Excel で「数値が文字列として保存」の警告が出る）
                    [void]$xml.Append(('<c r="{0}{1}" s="2"><v>{2}</v></c>' -f (ConvertTo-ColumnName ($i + 1)), $r, ([double]$raw).ToString('R', $invariant)))
                    continue
                }
                $content = ConvertTo-InlineStringXml $raw
                if (-not $content) { continue }
                $style = 2
                if ($row.Link -and $row.Link.Column -eq $i) { $style = 3 }
                [void]$xml.Append(('<c r="{0}{1}" s="{2}" t="inlineStr">{3}</c>' -f (ConvertTo-ColumnName ($i + 1)), $r, $style, $content))
            }
            [void]$xml.Append('</row>')

            if ($row.Link) {
                if ($linkCount -ge 65530) {
                    # Excel の 1 シートあたりのハイパーリンクの上限
                    if ($linkCount -eq 65530) { Write-Warning ("シート「{0}」のリンクが上限（65,530 個）に達したため、{1} 行目以降にはリンクを付けていません。" -f $s.Name, $r) }
                    $linkCount++
                    continue
                }
                $linkCount++
                # Excel が他ブックへのリンクを保存するときと同じ file:/// 形式にする
                $target = $row.Link.Path
                if ($target.StartsWith('/')) { $target = 'file://' + $target } else { $target = 'file:///' + $target }
                $tip = ''
                if ($row.Link.Tooltip) { $tip = ' tooltip="{0}"' -f (ConvertTo-XmlText $row.Link.Tooltip) }
                # Location が空ならファイルを開くだけのリンクにする
                if ($row.Link.Location) { $tip = (' location="{0}"' -f (ConvertTo-XmlText $row.Link.Location)) + $tip }
                [void]$rels.Append(('<Relationship Id="rIdL{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{1}" TargetMode="External"/>' -f $linkCount, (ConvertTo-XmlText $target)))
                [void]$links.Append(('<hyperlink ref="{0}{1}" r:id="rIdL{2}"{3}/>' -f (ConvertTo-ColumnName ($row.Link.Column + 1)), $r, $linkCount, $tip))
            }
        }
        [void]$xml.Append('</sheetData>')
        [void]$xml.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $lastCol, $rowCount))
        if ($linkCount -gt 0) { [void]$xml.Append('<hyperlinks>').Append($links.ToString()).Append('</hyperlinks>') }
        [void]$xml.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape"/></worksheet>')
        [void]$rels.Append('</Relationships>')

        $sheetEntries.Add([pscustomobject]@{ Name = $s.Name; Xml = $xml.ToString(); Rels = $rels.ToString(); HasLinks = ($linkCount -gt 0) })
        [void]$definedNames.Append(('<definedName name="_xlnm._FilterDatabase" localSheetId="{0}" hidden="1">''{1}''!$A$1:${2}${3}</definedName>' -f ($index - 1), (ConvertTo-XmlText $s.Name.Replace("'", "''")), $lastCol, $rowCount))
    }

    $ct = New-Object System.Text.StringBuilder
    [void]$ct.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
    $wbSheets = New-Object System.Text.StringBuilder
    $wbRels = New-Object System.Text.StringBuilder
    [void]$wbRels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
    for ($i = 1; $i -le $sheetEntries.Count; $i++) {
        [void]$ct.Append(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $i))
        [void]$wbSheets.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (ConvertTo-XmlText $sheetEntries[$i - 1].Name), $i))
        [void]$wbRels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $i))
    }
    [void]$ct.Append('</Types>')
    [void]$wbRels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' -f ($sheetEntries.Count + 1)))

    $parts = [ordered]@{
        '[Content_Types].xml' = $ct.ToString()
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>{0}</sheets><definedNames>{1}</definedNames></workbook>' -f $wbSheets.ToString(), $definedNames.ToString())
        'xl/_rels/workbook.xml.rels' = $wbRels.ToString()
        # スタイル 0:標準 1:見出し 2:折り返し・上詰め 3:リンク
        'xl/styles.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font>{0}</font><font><b/>{0}</font><font><u/><sz val="10"/><color rgb="FF0563C1"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>' -f $font)
    }
    for ($i = 1; $i -le $sheetEntries.Count; $i++) {
        $parts["xl/worksheets/sheet$i.xml"] = $sheetEntries[$i - 1].Xml
        if ($sheetEntries[$i - 1].HasLinks) { $parts["xl/worksheets/_rels/sheet$i.xml.rels"] = $sheetEntries[$i - 1].Rels }
    }

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $stream = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($name in $parts.Keys) {
                $writer = New-Object System.IO.StreamWriter($zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal).Open(), $utf8)
                try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $stream.Dispose() }
}

Export-ModuleMember -Function *
