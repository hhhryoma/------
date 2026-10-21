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

function Read-TextFile {
    #     テキストファイル（CSV など）を読む。BOM 付き・BOM なしの UTF-8 と、Excel が日本語の Windows で保存する Shift_JIS に対応する
    param([string] $Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3) }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) { return [System.Text.Encoding]::Unicode.GetString($bytes, 2, $bytes.Length - 2) }
    try { return (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
    catch { return [System.Text.Encoding]::GetEncoding(932).GetString($bytes) }
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

function Read-PartXmlForEdit {
    #     書き換えるパーツを読む。空白やコメントも含めてそのまま保つ（書き戻したときに、変えていない文字列が変わらないように）
    param($Package, [string] $PartName)

    $entry = Get-PackageEntry -Package $Package -PartName $PartName
    if ($null -eq $entry) { return $null }
    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $settings.CloseInput = $true
    $reader = [System.Xml.XmlReader]::Create($entry.Open(), $settings)
    try {
        $doc = New-Object System.Xml.XmlDocument
        $doc.XmlResolver = $null
        $doc.PreserveWhitespace = $true
        $doc.Load($reader)
        return $doc
    }
    finally { $reader.Dispose() }
}

function ConvertTo-XmlString {
    #     XmlDocument を、UTF-8 の宣言付きの文字列にする（字下げや改行は足さない）
    param([System.Xml.XmlDocument] $Document)

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Encoding = New-Object System.Text.UTF8Encoding($false)
    $settings.Indent = $false
    $settings.NewLineHandling = [System.Xml.NewLineHandling]::None
    $ms = New-Object System.IO.MemoryStream
    $writer = [System.Xml.XmlWriter]::Create($ms, $settings)
    try { $Document.Save($writer) } finally { $writer.Dispose() }
    return [System.Text.Encoding]::UTF8.GetString($ms.ToArray())
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

# ---------------------------------------------------------------- 書き換え（xlsx / xlsm の一部のパーツだけ）


function Save-ZipPackage {
    #     ZIP パッケージの一部のパーツだけを書き換えて保存する。変更しないパーツは同じ名前・日時・中身で書き直す
    #     （ZipArchive は圧縮済みのままコピーできないため展開して書き直す）。
    #     Changes: パーツ名（小文字）→ 新しい内容。NewParts: 足すパーツ名 → 内容（末尾に足す）。
    #     OutPath を省略すると元のファイルに上書きする。その際、NTFS の代替データストリーム（インターネットから取得した
    #     ファイルの印 Zone.Identifier など）やアクセス権・作成日時を失わないよう、元のファイルを開いて中身だけを書き換える
    param([string] $FilePath, [hashtable] $Changes, $NewParts, [string] $OutPath)

    $target = $OutPath
    if (-not $target) { $target = $FilePath + '.~tmp' + [guid]::NewGuid().ToString('N').Substring(0, 8) }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $src = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    try {
        $inZip = New-Object System.IO.Compression.ZipArchive($src, [System.IO.Compression.ZipArchiveMode]::Read)
        $mode = [System.IO.FileMode]::Create
        if (-not $OutPath) { $mode = [System.IO.FileMode]::CreateNew }
        $dst = New-Object System.IO.FileStream($target, $mode, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
        try {
            $outZip = New-Object System.IO.Compression.ZipArchive($dst, [System.IO.Compression.ZipArchiveMode]::Create)
            try {
                foreach ($e in $inZip.Entries) {
                    $o = $outZip.CreateEntry($e.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
                    $o.LastWriteTime = $e.LastWriteTime
                    $os = $o.Open()
                    try {
                        $key = $e.FullName.ToLowerInvariant()
                        if ($Changes -and $Changes.ContainsKey($key)) {
                            $bytes = $utf8.GetBytes($Changes[$key])
                            $os.Write($bytes, 0, $bytes.Length)
                        }
                        else {
                            $is = $e.Open()
                            try { $is.CopyTo($os) } finally { $is.Dispose() }
                        }
                    }
                    finally { $os.Dispose() }
                }
                if ($NewParts) {
                    foreach ($name in $NewParts.Keys) {
                        $writer = New-Object System.IO.StreamWriter($outZip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal).Open(), $utf8)
                        try { $writer.Write($NewParts[$name]) } finally { $writer.Dispose() }
                    }
                }
            }
            finally { $outZip.Dispose() }
        }
        finally { $dst.Dispose() }
    }
    finally { $src.Dispose() }
    if ($OutPath) { return }

    # 元のファイルに上書きする
    $written = $false
    try {
        $in = [System.IO.File]::OpenRead($target)
        try {
            $out = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try {
                $out.SetLength(0)
                $in.CopyTo($out)
                $written = $true
            }
            finally { $out.Dispose() }
        }
        finally { $in.Dispose() }
    }
    catch {
        if (-not $written) {
            # 途中で失敗した場合は、書き換え後の内容で置き換えて壊れた状態を残さない
            try { [System.IO.File]::Copy($target, $FilePath, $true); $written = $true } catch { }
        }
        if (-not $written) { throw ("書き換えに失敗しました。書き換え後の内容が {0} に残っています: {1}" -f $target, $_.Exception.Message) }
    }
    Remove-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
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

# ---------------------------------------------------------------- 設計書（Excel 方眼紙）を読む

$script:Nfkc = [System.Text.NormalizationForm]::FormKC

function ConvertTo-Label {
    #     見出しや項目名を比べるための形（全角・半角をそろえ、空白を除く）
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    return ($Text.Normalize($script:Nfkc) -replace '\s', '')
}

function Read-CsvFile {
    #     設定の CSV を読み、必要な列がそろっているかを確かめる
    param([string] $FilePath, [string[]] $Required, [string] $What)
    if (-not (Test-Path -LiteralPath $FilePath)) { throw "$What が見つかりません: $FilePath" }
    $rows = @(Read-TextFile $FilePath | ConvertFrom-Csv)
    if ($rows.Count -gt 0) {
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        foreach ($r in $Required) { if ($names -notcontains $r) { throw "$What に「$r」の列がありません: $FilePath" } }
    }
    return , $rows
}

function ConvertTo-DesignBooks {
    #     Search-Excel.ps1 の結果（-Pattern '\S' -Target Value）を、ファイル → シート → 行番号 → 列番号 → 文字 の形にまとめる
    #     戻り値: パス → @{ File; Path; Sheets = シート名 → シート }（ファイル・シートは読んだ順）
    #     シート: @{ Cells = 行番号 → @{ 列番号 → 文字 }; Rows = 行番号（昇順）; Cols = 行番号 → 列番号（昇順）; Labels = 行番号 → @{ 列番号 → ConvertTo-Label 済みの文字 } }
    param($Hits)
    $books = [ordered]@{}
    foreach ($h in $Hits) {
        $m = [regex]::Match([string]$h.Cell, '^([A-Z]+)(\d+)$')
        if (-not $m.Success) { continue }
        if (-not $books.Contains($h.Path)) { $books[$h.Path] = [pscustomobject]@{ File = $h.File; Path = $h.Path; Sheets = [ordered]@{} } }
        $sheets = $books[$h.Path].Sheets
        if (-not $sheets.Contains($h.Sheet)) { $sheets[$h.Sheet] = @{} }
        $row = [int]$m.Groups[2].Value
        if (-not $sheets[$h.Sheet].ContainsKey($row)) { $sheets[$h.Sheet][$row] = @{} }
        $sheets[$h.Sheet][$row][(Get-ColumnIndex $m.Groups[1].Value)] = [string]$h.Text
    }
    # 見出しを探すたびに比べる形へ変換しないよう、先に変換しておく
    foreach ($book in $books.Values) {
        foreach ($name in @($book.Sheets.Keys)) {
            $cells = $book.Sheets[$name]
            $rows = [int[]]@($cells.Keys)
            [Array]::Sort($rows)
            $cols = @{}; $labels = @{}
            foreach ($r in $rows) {
                $c = [int[]]@($cells[$r].Keys)
                [Array]::Sort($c)
                $cols[$r] = $c
                $l = @{}
                foreach ($col in $c) { $l[$col] = ($cells[$r][$col].Normalize($script:Nfkc) -replace '\s', '') }
                $labels[$r] = $l
            }
            $book.Sheets[$name] = [pscustomobject]@{ Cells = $cells; Rows = $rows; Cols = $cols; Labels = $labels }
        }
    }
    return $books
}

function Find-LabelCell {
    #     「画面名」などの見出し（ConvertTo-Label 済み）のセルの右にある値と、その行・列。見出しが無ければ $null
    #     Grid は ConvertTo-DesignBooks のシート
    param($Grid, [string] $Label)
    if (-not $Label) { return $null }
    foreach ($row in $Grid.Rows) {
        $cols = $Grid.Cols[$row]
        $labels = $Grid.Labels[$row]
        for ($i = 0; $i -lt $cols.Count - 1; $i++) {
            if ($labels[$cols[$i]] -eq $Label) {
                return [pscustomobject]@{ Value = $Grid.Cells[$row][$cols[$i + 1]].Trim(); Row = $row; Col = $cols[$i + 1] }
            }
        }
    }
    return $null
}

function Find-LabelValue {
    #     「画面名」などの見出しのセルの右にある値（見出しが無ければ $null）
    param($Grid, [string] $Label)
    $c = Find-LabelCell $Grid $Label
    if ($c) { return $c.Value }
    return $null
}

function Find-DesignTables {
    #     見出し（ConvertTo-Label 済み）がすべてそろっている行を探し、その下で KeyHeader の列に値のある行を返す。
    #     「■」で始まる行（次の節）までを 1 つの表とみなす。KeyHeader が無い表では最初の見出しの列を使う
    #     Grid は ConvertTo-DesignBooks のシート
    #     戻り値: @{ HeaderRow; ColumnOf = 見出し → 列番号; Rows = @(@{ Row; Cells = 列番号 → 文字 }) } の一覧
    param($Grid, [string[]] $Headers, [string] $KeyHeader)
    $tables = New-Object System.Collections.Generic.List[object]
    if (@($Headers).Count -eq 0) { return , $tables }
    $rows = $Grid.Rows
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $labels = $Grid.Labels[$rows[$i]]
        $colOf = @{}
        foreach ($c in $Grid.Cols[$rows[$i]]) { $label = $labels[$c]; if (-not $colOf.ContainsKey($label)) { $colOf[$label] = $c } }
        $ok = $true
        foreach ($hdr in $Headers) { if (-not $colOf.ContainsKey($hdr)) { $ok = $false; break } }
        if (-not $ok) { continue }
        $keyCol = $colOf[$Headers[0]]
        if ($KeyHeader -and $colOf.ContainsKey($KeyHeader)) { $keyCol = $colOf[$KeyHeader] }
        $dataRows = New-Object System.Collections.Generic.List[object]
        for ($j = $i + 1; $j -lt $rows.Count; $j++) {
            $r = $Grid.Cells[$rows[$j]]
            $section = $false
            foreach ($v in $r.Values) { if ($v.TrimStart().StartsWith('■')) { $section = $true; break } }
            if ($section) { break }
            if ($r.ContainsKey($keyCol) -and $r[$keyCol].Trim()) { $dataRows.Add([pscustomobject]@{ Row = $rows[$j]; Cells = $r }) }
        }
        $tables.Add([pscustomobject]@{ HeaderRow = $rows[$i]; ColumnOf = $colOf; Rows = $dataRows })
        $i = $j - 1
    }
    return , $tables
}

# ---------------------------------------------------------------- シートの XML のセルを書き換える


function Get-WorkbookSheetParts {
    #     ブックのパーツ名と、シート名 → シートのパーツ名（シートの順）
    param($Package)
    $wbPart = @((Get-PartRelationships -Package $Package -PartName '') | Where-Object { $_.Kind -eq 'officeDocument' })[0].Target
    $wb = Read-PartXml -Package $Package -PartName $wbPart
    $rels = Get-PartRelationships -Package $Package -PartName $wbPart
    $sheets = [ordered]@{}
    foreach ($s in $wb.SelectNodes("//*[local-name()='sheet']")) {
        $rid = Get-RelationshipIdAttribute $s
        $t = @($rels | Where-Object { $_.Id -eq $rid })[0]
        if ($t) { $sheets[$s.GetAttribute('name')] = $t.Target }
    }
    return [pscustomobject]@{ WorkbookPart = $wbPart; Sheets = $sheets }
}

function Get-XmlChildElements {
    param($Node, [string] $LocalName)
    return @($Node.ChildNodes | Where-Object { $_.NodeType -eq [System.Xml.XmlNodeType]::Element -and $_.LocalName -eq $LocalName })
}

function Get-SheetRowElement {
    #     行の要素（無ければ $null。-Create なら行番号の順の位置に作り、TemplateRow の属性（高さなど）を写す）
    param([System.Xml.XmlDocument] $Doc, [int] $Row, [switch] $Create, $TemplateRow)
    $sheetData = @(Get-XmlChildElements $Doc.DocumentElement 'sheetData')[0]
    $n = 0; $next = $null
    foreach ($el in (Get-XmlChildElements $sheetData 'row')) {
        $r = $el.GetAttribute('r')
        if ($r) { $n = [int]$r } else { $n++; $el.SetAttribute('r', [string]$n) }
        if ($n -eq $Row) { return $el }
        if ($n -gt $Row -and -not $next) { $next = $el }
    }
    if (-not $Create) { return $null }
    $new = $Doc.CreateElement($sheetData.Prefix, 'row', $sheetData.NamespaceURI)
    if ($TemplateRow) { foreach ($a in @($TemplateRow.Attributes)) { if ($a.Name -ne 'r') { $new.SetAttribute($a.Name, $a.Value) } } }
    $new.SetAttribute('r', [string]$Row)
    if ($next) { [void]$sheetData.InsertBefore($new, $next) } else { [void]$sheetData.AppendChild($new) }
    return $new
}

function Get-SheetCellElement {
    #     セルの要素（無ければ $null。-Create なら列の順の位置に作る）
    param([System.Xml.XmlDocument] $Doc, $RowElement, [int] $Row, [int] $Col, [switch] $Create, [string] $Style)
    $n = 0; $next = $null; $found = $null
    foreach ($c in (Get-XmlChildElements $RowElement 'c')) {
        $ref = $c.GetAttribute('r')
        if ($ref) { $n = Get-ColumnIndex $ref } else { $n++; $c.SetAttribute('r', ('{0}{1}' -f (ConvertTo-ColumnName $n), $Row)) }
        if ($n -eq $Col) { $found = $c }
        elseif ($n -gt $Col -and -not $next) { $next = $c }
    }
    if ($found -or -not $Create) { return $found }
    $new = $Doc.CreateElement($RowElement.Prefix, 'c', $RowElement.NamespaceURI)
    $new.SetAttribute('r', ('{0}{1}' -f (ConvertTo-ColumnName $Col), $Row))
    if ($Style) { $new.SetAttribute('s', $Style) }
    if ($next) { [void]$RowElement.InsertBefore($new, $next) } else { [void]$RowElement.AppendChild($new) }
    return $new
}

function Test-NumericCellElement {
    #     数値のセル（日付のシリアル値など）か
    param($Cell)
    if (-not $Cell) { return $false }
    $t = $Cell.GetAttribute('t')
    return (-not $t -or $t -eq 'n') -and @(Get-XmlChildElements $Cell 'v').Count -gt 0
}

function Set-SheetCellValue {
    #     セルの値を書き換える（文字はセル内の文字列、数は数値、$null は空。書式 s はそのまま。数式は消える）
    param([System.Xml.XmlDocument] $Doc, $Cell, $Value)
    foreach ($child in @($Cell.ChildNodes)) { [void]$Cell.RemoveChild($child) }
    if ($null -eq $Value) { $Cell.RemoveAttribute('t'); return }
    if ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal]) {
        $Cell.RemoveAttribute('t')
        $v = $Doc.CreateElement($Cell.Prefix, 'v', $Cell.NamespaceURI)
        $v.InnerText = ([double]$Value).ToString('R', $invariant)
        [void]$Cell.AppendChild($v)
        return
    }
    $Cell.SetAttribute('t', 'inlineStr')
    $is = $Doc.CreateElement($Cell.Prefix, 'is', $Cell.NamespaceURI)
    $t = $Doc.CreateElement($Cell.Prefix, 't', $Cell.NamespaceURI)
    [void]$t.SetAttribute('space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
    $t.InnerText = [string]$Value
    [void]$is.AppendChild($t)
    [void]$Cell.AppendChild($is)
}

function Copy-SheetRowMerges {
    #     もとの行の結合（1 行の中だけのもの）を新しい行にも付ける
    param([System.Xml.XmlDocument] $Doc, [int] $FromRow, [int] $ToRow)
    $merges = @(Get-XmlChildElements $Doc.DocumentElement 'mergeCells')[0]
    if (-not $merges) { return }
    $ranges = New-Object System.Collections.Generic.List[object]
    foreach ($mc in (Get-XmlChildElements $merges 'mergeCell')) {
        $m = [regex]::Match($mc.GetAttribute('ref'), '^([A-Z]+)(\d+):([A-Z]+)(\d+)$')
        if ($m.Success) { $ranges.Add([pscustomobject]@{ C1 = (Get-ColumnIndex $m.Groups[1].Value); R1 = [int]$m.Groups[2].Value; C2 = (Get-ColumnIndex $m.Groups[3].Value); R2 = [int]$m.Groups[4].Value }) }
    }
    $added = 0
    foreach ($g in @($ranges | Where-Object { $_.R1 -eq $FromRow -and $_.R2 -eq $FromRow })) {
        $overlap = @($ranges | Where-Object { $_.R1 -le $ToRow -and $_.R2 -ge $ToRow -and $_.C1 -le $g.C2 -and $_.C2 -ge $g.C1 }).Count -gt 0
        if ($overlap) { continue }
        $mc = $Doc.CreateElement($merges.Prefix, 'mergeCell', $merges.NamespaceURI)
        $mc.SetAttribute('ref', ('{0}{2}:{1}{2}' -f (ConvertTo-ColumnName $g.C1), (ConvertTo-ColumnName $g.C2), $ToRow))
        [void]$merges.AppendChild($mc)
        $added++
    }
    if ($added -gt 0 -and $merges.HasAttribute('count')) { $merges.SetAttribute('count', [string]((Get-XmlChildElements $merges 'mergeCell').Count)) }
}

function Update-SheetDimension {
    #     シートの使用範囲（dimension）を、Row 行目まで広げる
    param([System.Xml.XmlDocument] $Doc, [int] $Row)
    $dim = @(Get-XmlChildElements $Doc.DocumentElement 'dimension')[0]
    if (-not $dim) { return }
    $m = [regex]::Match($dim.GetAttribute('ref'), '^([A-Z]+\d+):([A-Z]+)(\d+)$')
    if ($m.Success -and [int]$m.Groups[3].Value -lt $Row) { $dim.SetAttribute('ref', ('{0}:{1}{2}' -f $m.Groups[1].Value, $m.Groups[2].Value, $Row)) }
}

function Get-VersionlessName {
    #     ファイル名（相対パス）から版数（_v1.0・Ver1.0・第1.0版 など）を除いた形（版数だけが違うファイルどうしを対応付けるため）
    param([string] $Name)
    return [regex]::Replace($Name.Normalize($script:Nfkc).ToLowerInvariant(), '(?:(?<![a-z])v(?:er)?\.?\s*|第\s*)\d+(?:\.\d+)*(?:版)?', '{v}')
}

function ConvertTo-ExcelDate {
    #     日付を Excel のシリアル値にする
    param([datetime] $Date)
    return [double]($Date.Date - [datetime]'1899-12-30').TotalDays
}

function Set-FullCalcOnLoad {
    #     workbook.xml の文字列に、開いたときに再計算させる指定（calcPr fullCalcOnLoad="1"）を付ける
    param([string] $WorkbookXml)
    $px = [regex]::Match($WorkbookXml, '<(\w+:)?workbook\b').Groups[1].Value
    $calc = [regex]::Match($WorkbookXml, ('<{0}calcPr\b[^>]*?/?>' -f $px))
    if ($calc.Success) {
        $tag = $calc.Value
        if ($tag -match '\sfullCalcOnLoad="[^"]*"') { $tag = $tag -replace '\sfullCalcOnLoad="[^"]*"', ' fullCalcOnLoad="1"' }
        elseif ($tag.EndsWith('/>')) { $tag = $tag.Substring(0, $tag.Length - 2).TrimEnd() + ' fullCalcOnLoad="1"/>' }
        else { $tag = $tag.Substring(0, $tag.Length - 1) + ' fullCalcOnLoad="1">' }
        return $WorkbookXml.Substring(0, $calc.Index) + $tag + $WorkbookXml.Substring($calc.Index + $calc.Length)
    }
    # calcPr は definedNames の後、oleSize などの前に置く
    $after = [regex]::Match($WorkbookXml, ('<{0}(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b' -f $px))
    $at = $WorkbookXml.LastIndexOf(('</{0}workbook>' -f $px))
    if ($after.Success) { $at = $after.Index }
    return $WorkbookXml.Substring(0, $at) + ('<{0}calcPr fullCalcOnLoad="1"/>' -f $px) + $WorkbookXml.Substring($at)
}

# ---------------------------------------------------------------- 数式・ヘッダー／フッター


# 数式の中のセル参照（A1・$A$1・A:A・1:1）。文字列（"..."）・シート名（'...'）・テーブルの列名（[...]）は飛ばす。
# 関数名（LOG10( など）や、シート名（Sheet1! など）の一部には一致させない
$script:FormulaRefRx = New-Object System.Text.RegularExpressions.Regex(
    '"(?:[^"]|"")*"|''(?:[^'']|'''')*''|\[[^\]]*\]|(?<![A-Za-z0-9_.$])(?:(\$?)([A-Z]{1,3})(\$?)(\d+)|(\$?)([A-Z]{1,3}):(\$?)([A-Z]{1,3})|(\$?)(\d+):(\$?)(\d+))(?![A-Za-z0-9_.(!])')

function Move-FormulaReference {
    #     数式の相対参照を RowOffset 行・ColumnOffset 列ずらす（$ の付いた行・列はずらさない）。
    #     共有数式（Excel が数式をコピーしたときの保存形式。先頭のセルにだけ数式があり、写しのセルには数式が無い）の、
    #     写しのセルの数式を求めるのに使う。シートの外に出た参照は #REF! にする
    param([string] $Formula, [int] $RowOffset, [int] $ColumnOffset)

    if ($RowOffset -eq 0 -and $ColumnOffset -eq 0) { return $Formula }
    return $script:FormulaRefRx.Replace($Formula, {
        param($m)
        $g = $m.Groups
        # 列（C）・行（R）の並び。A1 は列と行を区切らずにつなげ、A:B・1:2 は : でつなげる
        $parts = New-Object System.Collections.Generic.List[string]
        if ($g[2].Success) { $parts.Add('C' + $g[1].Value + $g[2].Value); $parts.Add('R' + $g[3].Value + $g[4].Value); $sep = '' }
        elseif ($g[6].Success) { $parts.Add('C' + $g[5].Value + $g[6].Value); $parts.Add('C' + $g[7].Value + $g[8].Value); $sep = ':' }
        elseif ($g[10].Success) { $parts.Add('R' + $g[9].Value + $g[10].Value); $parts.Add('R' + $g[11].Value + $g[12].Value); $sep = ':' }
        else { return $m.Value }
        $out = New-Object System.Text.StringBuilder
        for ($i = 0; $i -lt $parts.Count; $i++) {
            $p = $parts[$i]
            if ($i -gt 0) { [void]$out.Append($sep) }
            if ($p.Substring(1, 1) -eq '$') { [void]$out.Append($p.Substring(1)); continue }
            if ($p[0] -eq 'C') {
                $c = (Get-ColumnIndex $p.Substring(1)) + $ColumnOffset
                if ($c -lt 1 -or $c -gt 16384) { return '#REF!' }
                [void]$out.Append((ConvertTo-ColumnName $c))
            }
            else {
                $r = [long]$p.Substring(1) + $RowOffset
                if ($r -lt 1 -or $r -gt 1048576) { return '#REF!' }
                [void]$out.Append($r)
            }
        }
        return $out.ToString()
    })
}

function Split-HeaderFooterText {
    #     ヘッダー・フッターの文字（&L 左 &C 中央 &R 右。&P などはページ番号などの項目、&"フォント" &12 &B などは書式）を区切る。
    #     返り値: 区切りの並び。Section（L / C / R）、Kind（Text: 文字 / Field: ページ番号などの項目 / Format: 書式 / Section: 位置の指定）、
    #     Raw（元の書き方）、InnerText（Text は文字（&& は & に戻す）、Field は Excel の画面と同じ書き方（&[ページ番号] など）、それ以外は空）。
    #     Text の InnerText を書き換えてから Join-HeaderFooterText に渡すと、書式や項目を保ったまま文字を変えられる
    param([string] $Text)

    $fields = @{ P = '&[ページ番号]'; N = '&[総ページ数]'; D = '&[日付]'; T = '&[時刻]'; Z = '&[パス]'; F = '&[ファイル名]'; A = '&[シート名]'; G = '&[図]' }
    $tokens = New-Object System.Collections.Generic.List[object]
    # 位置の指定が無い文字は中央に出る
    $section = 'C'
    foreach ($m in [regex]::Matches([string]$Text, '&&|&[LCR]|&[PNDTZFAG]|&"[^"]*"?|&K(?:[0-9A-Fa-f]{6}|\d\d[+-]\d{3})|&\d+|&.?|[^&]+', 'Singleline')) {
        $raw = $m.Value
        if ($raw -eq '&&' -or -not $raw.StartsWith('&')) {
            $t = $raw
            if ($raw -eq '&&') { $t = '&' }
            # 続いている文字は 1 つにまとめる
            if ($tokens.Count -gt 0 -and $tokens[$tokens.Count - 1].Kind -eq 'Text') {
                $last = $tokens[$tokens.Count - 1]
                $last.Raw += $raw; $last.InnerText += $t
                continue
            }
            $tokens.Add([pscustomobject]@{ Section = $section; Kind = 'Text'; Raw = $raw; InnerText = $t })
        }
        elseif ($raw -cmatch '^&[LCR]$') {
            $section = $raw.Substring(1)
            $tokens.Add([pscustomobject]@{ Section = $section; Kind = 'Section'; Raw = $raw; InnerText = '' })
        }
        elseif ($raw -cmatch '^&[PNDTZFAG]$') {
            $tokens.Add([pscustomobject]@{ Section = $section; Kind = 'Field'; Raw = $raw; InnerText = $fields[$raw.Substring(1)] })
        }
        else {
            $tokens.Add([pscustomobject]@{ Section = $section; Kind = 'Format'; Raw = $raw; InnerText = '' })
        }
    }
    return , $tokens
}

function Join-HeaderFooterText {
    #     Split-HeaderFooterText の区切りをつなげて、ヘッダー・フッターの文字に戻す（文字の & は && にする）
    param($Tokens)
    $sb = New-Object System.Text.StringBuilder
    foreach ($t in $Tokens) {
        if ($t.Kind -eq 'Text') { [void]$sb.Append($t.InnerText.Replace('&', '&&')) }
        else { [void]$sb.Append($t.Raw) }
    }
    return $sb.ToString()
}

function Get-HeaderFooterSections {
    #     ヘッダー・フッターの文字を、位置（左・中央・右）ごとの表示の文字にする。返り値: @(@{ Section = '左'; Text }) （空の位置は除く）
    param($Tokens)
    $names = [ordered]@{ L = '左'; C = '中央'; R = '右' }
    $result = New-Object System.Collections.Generic.List[object]
    foreach ($key in $names.Keys) {
        $text = -join @($Tokens | Where-Object { $_.Section -eq $key -and ($_.Kind -eq 'Text' -or $_.Kind -eq 'Field') } | ForEach-Object { $_.InnerText })
        if ($text) { $result.Add([pscustomobject]@{ Key = $key; Section = $names[$key]; Text = $text }) }
    }
    return , $result
}

# ヘッダー・フッターの要素名 → 表示の名前
$script:HeaderFooterNames = [ordered]@{
    oddHeader = 'ヘッダー'; oddFooter = 'フッター'; evenHeader = '偶数ページのヘッダー'; evenFooter = '偶数ページのフッター'
    firstHeader = '先頭ページのヘッダー'; firstFooter = '先頭ページのフッター'
}

function Get-HeaderFooterName {
    #     ヘッダー・フッターの要素名（oddHeader など）の表示の名前。ヘッダー・フッターの要素でなければ $null
    param([string] $LocalName)
    if ($script:HeaderFooterNames.Contains($LocalName)) { return $script:HeaderFooterNames[$LocalName] }
    return $null
}

function Get-ValidationTexts {
    #     入力規則（dataValidation 要素。拡張部分の x14:dataValidation も）の、検索・置換する文字。
    #     返り値: @(@{ Name; Text; Node; Attribute; IsList }) Name は「リスト」「条件」「入力時メッセージ」など。
    #     リスト（"未着手,完了" のように直接書いたもの）の Text は値の並び。参照・条件の Text は =数式。
    #     Node・Attribute は書き換える場所（Attribute が空なら Node の文字）
    param([System.Xml.XmlElement] $Element)

    $result = New-Object System.Collections.Generic.List[object]
    $type = $Element.GetAttribute('type')
    foreach ($fn in @('formula1', 'formula2')) {
        $f = @(Get-XmlChildElements $Element $fn)[0]
        if (-not $f) { continue }
        # 拡張部分（x14）は <x14:formula1><xm:f>…</xm:f></x14:formula1>
        $inner = @(Get-XmlChildElements $f 'f')[0]
        if ($inner) { $f = $inner }
        $text = $f.InnerText
        if ($type -eq 'list' -and $text -match '^"([\s\S]*)"$') {
            $result.Add([pscustomobject]@{ Name = 'リスト'; Text = $Matches[1].Replace('""', '"'); Node = $f; Attribute = ''; IsList = $true })
        }
        elseif ($text) {
            $name = '条件'
            if ($type -eq 'list') { $name = 'リスト' }
            $result.Add([pscustomobject]@{ Name = $name; Text = '=' + $text; Node = $f; Attribute = ''; IsList = $false })
        }
    }
    foreach ($a in @(@('promptTitle', '入力時メッセージのタイトル'), @('prompt', '入力時メッセージ'), @('errorTitle', 'エラーメッセージのタイトル'), @('error', 'エラーメッセージ'))) {
        $v = $Element.GetAttribute($a[0])
        if ($v) { $result.Add([pscustomobject]@{ Name = $a[1]; Text = $v; Node = $Element; Attribute = $a[0]; IsList = $false }) }
    }
    return , $result
}

function Get-ValidationRange {
    #     入力規則の範囲（sqref 属性。x14 では xm:sqref 要素）
    param([System.Xml.XmlElement] $Element)
    $sqref = $Element.GetAttribute('sqref')
    if (-not $sqref) {
        $node = @(Get-XmlChildElements $Element 'sqref')[0]
        if ($node) { $sqref = $node.InnerText }
    }
    return $sqref.Trim()
}

# ---------------------------------------------------------------- Markdown にする（Export-ExcelToMarkdown・Invoke-DesignReview）


function ConvertTo-MarkdownText {
    #     セルの文字を Markdown の 1 行に収める（改行は <br>、表の中の | はエスケープ）
    param([string] $Text, [switch] $InTable)
    # Search-Excel はセル内の改行を ↵ で表す
    $t = ([string]$Text).Trim() -replace "\r\n|\r|\n|↵", '<br>'
    if ($InTable) { $t = $t.Replace('|', '\|') }
    return $t
}

function ConvertTo-SheetMarkdown {
    #     1 シートの値（行番号 → 列の順のセル）を Markdown の行にして Lines に足す。
    #       見出しの行（3 つ以上の値）の下に、見出しの列だけを使う行が続くところは表に、それ以外の行は箇条書きに、
    #       「■」で始まる行は小見出しにする。
    #       -CellAddress: 表に行番号、箇条書きにセル番地を付ける。-ColumnLetter: 表の見出しに列を付ける（C:項目名）
    param($Rows, [System.Collections.Generic.List[string]] $Lines, [switch] $CellAddress, [switch] $ColumnLetter)
    $rowNums = [int[]]@($Rows.Keys)
    [Array]::Sort($rowNums)
    $last = ''
    $i = 0
    while ($i -lt $rowNums.Count) {
        $cells = $Rows[$rowNums[$i]]
        if ($cells[0].Text.TrimStart().StartsWith('■')) {
            if ($last) { $Lines.Add('') }
            $Lines.Add('### ' + (ConvertTo-MarkdownText (@($cells | ForEach-Object { $_.Text.Trim() }) -join ' ')))
            $Lines.Add('')
            $last = ''
            $i++
            continue
        }
        $headerCols = @{}
        foreach ($c in $cells) { $headerCols[$c.Col] = $true }
        $j = $i + 1
        if ($cells.Count -ge 3) {
            while ($j -lt $rowNums.Count -and $rowNums[$j] -eq $rowNums[$j - 1] + 1) {
                $next = $Rows[$rowNums[$j]]
                if ($next[0].Text.TrimStart().StartsWith('■')) { break }
                $inside = $true
                foreach ($c in $next) { if (-not $headerCols.ContainsKey($c.Col)) { $inside = $false; break } }
                if (-not $inside) { break }
                $j++
            }
        }
        if ($j - $i -ge 2) {
            if ($last) { $Lines.Add('') }
            $cols = @($cells | ForEach-Object { $_.Col })
            $head = @($cells | ForEach-Object {
                    $h = ConvertTo-MarkdownText $_.Text -InTable
                    if ($ColumnLetter) { $h = '{0}:{1}' -f (ConvertTo-ColumnName $_.Col), $h }
                    $h
                })
            $sep = @($cols | ForEach-Object { '---' })
            if ($CellAddress) { $head = @('行') + $head; $sep = @('---') + $sep }
            $Lines.Add('| ' + ($head -join ' | ') + ' |')
            $Lines.Add('| ' + ($sep -join ' | ') + ' |')
            for ($k = $i + 1; $k -lt $j; $k++) {
                $byCol = @{}
                foreach ($c in $Rows[$rowNums[$k]]) { $byCol[$c.Col] = $c.Text }
                $vals = @($cols | ForEach-Object { ConvertTo-MarkdownText $byCol[$_] -InTable })
                if ($CellAddress) { $vals = @([string]$rowNums[$k]) + $vals }
                $Lines.Add('| ' + ($vals -join ' | ') + ' |')
            }
            $Lines.Add('')
            $last = ''
            $i = $j
            continue
        }
        if ($last -ne 'list' -and $Lines.Count -gt 0 -and $Lines[$Lines.Count - 1] -ne '') { $Lines.Add('') }
        if ($CellAddress) { $parts = @($cells | ForEach-Object { '{0}{1}: {2}' -f (ConvertTo-ColumnName $_.Col), $rowNums[$i], (ConvertTo-MarkdownText $_.Text) }) }
        else { $parts = @($cells | ForEach-Object { ConvertTo-MarkdownText $_.Text }) }
        $Lines.Add('- ' + ($parts -join ' | '))
        $last = 'list'
        $i++
    }
}

function ConvertTo-SheetMarkdownLines {
    #     Search-Excel の結果（1 ブック分。Target は Value・Comment・Shape・SheetName）を、
    #     シート名 → そのシートの Markdown の行（「## シート名」から）の順序付き辞書にする。
    #     図形・コメントはシートごとに後ろへまとめる
    param($Hits, [switch] $CellAddress, [switch] $ColumnLetter)
    # シート → 行番号 → セル（列の順）
    $sheets = [ordered]@{}
    foreach ($h in $Hits) {
        if (-not $sheets.Contains($h.Sheet)) { $sheets[$h.Sheet] = [pscustomobject]@{ Rows = @{}; Shapes = (New-Object System.Collections.Generic.List[object]); Comments = (New-Object System.Collections.Generic.List[object]) } }
        $s = $sheets[$h.Sheet]
        switch ($h.Kind) {
            '値' {
                $m = [regex]::Match([string]$h.Cell, '^([A-Z]+)(\d+)$')
                if (-not $m.Success) { break }
                $row = [int]$m.Groups[2].Value
                if (-not $s.Rows.ContainsKey($row)) { $s.Rows[$row] = New-Object System.Collections.Generic.List[object] }
                $s.Rows[$row].Add([pscustomobject]@{ Col = (Get-ColumnIndex $m.Groups[1].Value); Text = [string]$h.Text })
            }
            '図形' { $s.Shapes.Add($h) }
            'コメント' { $s.Comments.Add($h) }
        }
    }
    $result = [ordered]@{}
    foreach ($name in $sheets.Keys) {
        $s = $sheets[$name]
        $lines = New-Object System.Collections.Generic.List[string]
        $result[$name] = $lines
        $lines.Add('## ' + $name)
        $lines.Add('')
        if ($s.Rows.Count -eq 0 -and $s.Shapes.Count -eq 0 -and $s.Comments.Count -eq 0) { $lines.Add('（空）'); $lines.Add(''); continue }
        foreach ($row in @($s.Rows.Keys)) {
            $sorted = @($s.Rows[$row] | Sort-Object Col)
            $s.Rows[$row] = $sorted
        }
        ConvertTo-SheetMarkdown $s.Rows $lines -CellAddress:$CellAddress -ColumnLetter:$ColumnLetter
        if ($lines[$lines.Count - 1] -ne '') { $lines.Add('') }
        if ($s.Shapes.Count -gt 0) {
            $lines.Add('### 図形')
            $lines.Add('')
            foreach ($h in $s.Shapes) {
                $label = $h.Name
                if ($CellAddress) { $label = '{0}（{1} 付近）' -f $h.Name, $h.Cell }
                $lines.Add(('- {0}: {1}' -f $label, (ConvertTo-MarkdownText $h.Text)))
            }
            $lines.Add('')
        }
        if ($s.Comments.Count -gt 0) {
            $lines.Add('### コメント')
            $lines.Add('')
            foreach ($h in $s.Comments) { $lines.Add(('- {0}: {1}' -f $h.Cell, (ConvertTo-MarkdownText $h.Text))) }
            $lines.Add('')
        }
    }
    return $result
}

function ConvertTo-BookMarkdown {
    #     Search-Excel の結果（1 ブック分）を Markdown の文字列にする（Title があれば「# Title」の見出しを付ける）
    param([string] $Title, $Hits, [switch] $CellAddress)
    $lines = New-Object System.Collections.Generic.List[string]
    if ($Title) { $lines.Add('# ' + $Title); $lines.Add('') }
    $sheets = ConvertTo-SheetMarkdownLines $Hits -CellAddress:$CellAddress
    foreach ($name in $sheets.Keys) { $lines.AddRange($sheets[$name]) }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return (($lines -join "`n") + "`n")
# ---------------------------------------------------------------- 品質の分析（Measure-ReviewQuality など）


function Read-ZoneTable {
    #     ゾーン分析の表（CSV）を読む。RowAxis・ColumnAxis は 2 つの軸の列名（例: 工数密度・指摘密度）で、
    #     値は「少ない」「基準内」「多い」。戻り値: '行の値|列の値' → @{ Zone; View; Action; NeedsAction }
    param([string] $FilePath, [string] $RowAxis, [string] $ColumnAxis)
    $table = @{}
    foreach ($r in (Read-CsvFile $FilePath @($RowAxis, $ColumnAxis, 'ゾーン') 'ゾーン分析の表')) {
        $get = { param($name) if ($r.PSObject.Properties[$name]) { ([string]$r.$name).Trim() } else { '' } }
        $key = '{0}|{1}' -f (& $get $RowAxis), (& $get $ColumnAxis)
        $table[$key] = [pscustomobject]@{ Zone = (& $get 'ゾーン'); View = (& $get '見方'); Action = (& $get '対応の例'); NeedsAction = ((& $get '要対応') -match '^(○|〇|要|Y|yes|1)$') }
    }
    return $table
}

Export-ModuleMember -Function *
