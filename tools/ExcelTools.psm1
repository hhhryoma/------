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

function ConvertTo-ChartXml {
    #     Export-Workbook のシートの表から作るグラフの XML（Excel 不要。DrawingML のグラフ）
    #     Chart: @{ Type = 'Line' | 'Scatter' | 'Column'; Title; Category（0 始まりの列。Line・Column の項目、Scatter の X）;
    #               Series = @(@{ Column（0 始まりの列）; Name（無ければ見出し）; Type = 'Line'（Column のグラフに重ねる折れ線）; Secondary（右の軸）; Rows = @(行の位置) }) ;
    #               XTitle; YTitle; SecondaryTitle; SecondaryPercent（右の軸を % にする）; Labels（Scatter で系列の名前を点に表示）; Width; Height（セルの数） }
    #     データは 2 行目から LastRow 行目まで（Series の Rows を渡すと、その行だけ。1 行目が 2 行目にあたる）
    param($Chart, [string] $SheetName, [int] $LastRow)
    $sheetRef = "'" + $SheetName.Replace("'", "''") + "'"
    $range = {
        param([int] $Col, $Rows)
        $c = ConvertTo-ColumnName ($Col + 1)
        # @(0) は if では偽になるため、数で確かめる
        if ($null -ne $Rows -and @($Rows).Count -gt 0) { return '{0}!${1}${2}' -f $sheetRef, $c, ([int]@($Rows)[0] + 2) }
        return '{0}!${1}$2:${1}${2}' -f $sheetRef, $c, $LastRow
    }
    $text = {
        param([string] $Value, [int] $Size, [bool] $Bold)
        $b = ''
        if ($Bold) { $b = ' b="1"' }
        return ('<c:tx><c:rich><a:bodyPr/><a:p><a:pPr><a:defRPr sz="{0}"{1}/></a:pPr><a:r><a:rPr lang="ja-JP" sz="{0}"{1}/><a:t>{2}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/>' -f $Size, $b, (ConvertTo-XmlText $Value))
    }
    $seriesName = {
        param($Ser)
        if ($Ser.Name) { return '<c:tx><c:v>{0}</c:v></c:tx>' -f (ConvertTo-XmlText $Ser.Name) }
        return '<c:tx><c:strRef><c:f>{0}!${1}$1</c:f></c:strRef></c:tx>' -f $sheetRef, (ConvertTo-ColumnName ($Ser.Column + 1))
    }
    $valueAxis = {
        param([int] $Id, [int] $Cross, [string] $Pos, [string] $Title, [bool] $Grid, [bool] $Percent, [bool] $AtMax)
        $x = '<c:valAx><c:axId val="{0}"/><c:scaling><c:orientation val="minMax"/>' -f $Id
        if ($Percent) { $x += '<c:max val="1"/><c:min val="0"/>' }
        $x += '</c:scaling><c:delete val="0"/><c:axPos val="{0}"/>' -f $Pos
        if ($Grid) { $x += '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' }
        if ($Title) { $x += '<c:title>' + (& $text $Title 900 $false) + '</c:title>' }
        if ($Percent) { $x += '<c:numFmt formatCode="0%" sourceLinked="0"/>' } else { $x += '<c:numFmt formatCode="General" sourceLinked="1"/>' }
        $x += '<c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{0}"/>' -f $Cross
        if ($AtMax) { $x += '<c:crosses val="max"/>' } else { $x += '<c:crosses val="autoZero"/>' }
        $x += '<c:crossBetween val="between"/></c:valAx>'
        return $x
    }
    $categoryAxis = {
        param([int] $Id, [int] $Cross, [string] $Title, [bool] $Hidden)
        $x = '<c:catAx><c:axId val="{0}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="{1}"/><c:axPos val="b"/>' -f $Id, $(if ($Hidden) { 1 } else { 0 })
        if ($Title) { $x += '<c:title>' + (& $text $Title 900 $false) + '</c:title>' }
        $x += '<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="{0}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' -f $Cross
        return $x
    }

    # 系列の色（テーマの色に頼ると、テーマの無いブックでは色が付かないアプリがあるため、色を書く）
    $palette = @('4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E')
    $plot = New-Object System.Text.StringBuilder
    $series = @($Chart.Series)
    $index = 0
    if ($Chart.Type -eq 'Scatter') {
        [void]$plot.Append('<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>')
        foreach ($ser in $series) {
            [void]$plot.Append(('<c:ser><c:idx val="{0}"/><c:order val="{0}"/>' -f $index) + (& $seriesName $ser))
            # 点だけ（線を引かない）
            [void]$plot.Append(('<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="7"/><c:spPr><a:solidFill><a:srgbClr val="{0}"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:marker>' -f $palette[0]))
            if ($Chart.Labels) { [void]$plot.Append('<c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="800"/></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr><c:dLblPos val="r"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="1"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>') }
            [void]$plot.Append(('<c:xVal><c:numRef><c:f>{0}</c:f></c:numRef></c:xVal><c:yVal><c:numRef><c:f>{1}</c:f></c:numRef></c:yVal><c:smooth val="0"/></c:ser>' -f (& $range $Chart.Category $ser.Rows), (& $range $ser.Column $ser.Rows)))
            $index++
        }
        [void]$plot.Append('<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>')
        [void]$plot.Append((& $valueAxis 5001 5002 'b' $Chart.XTitle $false $false $false))
        [void]$plot.Append((& $valueAxis 5002 5001 'l' $Chart.YTitle $true $false $false))
    }
    else {
        $main = @($series | Where-Object { -not $_.Secondary -and -not ($Chart.Type -eq 'Column' -and $_.Type -eq 'Line') })
        $overlay = @($series | Where-Object { $_.Secondary -or ($Chart.Type -eq 'Column' -and $_.Type -eq 'Line') })
        $writeSeries = {
            param($List, [bool] $AsLine)
            foreach ($ser in $List) {
                [void]$plot.Append(('<c:ser><c:idx val="{0}"/><c:order val="{0}"/>' -f $script:chartSeriesIndex) + (& $seriesName $ser))
                $color = $palette[$script:chartSeriesIndex % $palette.Count]
                if ($AsLine) { [void]$plot.Append(('<c:spPr><a:ln w="22225" cap="rnd"><a:solidFill><a:srgbClr val="{0}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker>' -f $color)) }
                else { [void]$plot.Append(('<c:spPr><a:solidFill><a:srgbClr val="{0}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>' -f $color)) }
                [void]$plot.Append(('<c:cat><c:strRef><c:f>{0}</c:f></c:strRef></c:cat><c:val><c:numRef><c:f>{1}</c:f></c:numRef></c:val>' -f (& $range $Chart.Category $null), (& $range $ser.Column $null)))
                if ($AsLine) { [void]$plot.Append('<c:smooth val="0"/>') }
                [void]$plot.Append('</c:ser>')
                $script:chartSeriesIndex++
            }
        }
        $script:chartSeriesIndex = 0
        if ($Chart.Type -eq 'Column') {
            [void]$plot.Append('<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>')
            & $writeSeries $main $false
            [void]$plot.Append('<c:gapWidth val="60"/><c:axId val="5001"/><c:axId val="5002"/></c:barChart>')
        }
        else {
            [void]$plot.Append('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>')
            & $writeSeries $main $true
            [void]$plot.Append('<c:marker val="1"/><c:axId val="5001"/><c:axId val="5002"/></c:lineChart>')
        }
        if ($overlay.Count -gt 0) {
            [void]$plot.Append('<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>')
            & $writeSeries $overlay $true
            $ids = '<c:axId val="5001"/><c:axId val="5002"/>'
            if (@($overlay | Where-Object Secondary).Count -gt 0) { $ids = '<c:axId val="5003"/><c:axId val="5004"/>' }
            [void]$plot.Append('<c:marker val="1"/>' + $ids + '</c:lineChart>')
        }
        [void]$plot.Append((& $categoryAxis 5001 5002 $Chart.XTitle $false))
        [void]$plot.Append((& $valueAxis 5002 5001 'l' $Chart.YTitle $true $false $false))
        if (@($overlay | Where-Object Secondary).Count -gt 0) {
            [void]$plot.Append((& $categoryAxis 5003 5004 '' $true))
            [void]$plot.Append((& $valueAxis 5004 5003 'r' $Chart.SecondaryTitle $false ([bool]$Chart.SecondaryPercent) $true))
        }
    }
    $legend = '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>'
    if ($Chart.Type -eq 'Scatter' -and $Chart.Labels) { $legend = '' }
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart><c:title>{0}</c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>{1}</c:plotArea>{2}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="游ゴシック"/><a:ea typeface="游ゴシック"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr></c:chartSpace>' -f (& $text $Chart.Title 1200 $true), $plot.ToString(), $legend)
}

function Export-Workbook {
    #     複数シートの表を xlsx に書き出す（Excel 不要）。各シートは見出し行の固定とオートフィルター付き。
    #     Sheets: @(@{ Name; Columns = @(@{ Header; Width }); Rows = @(@{ Values = @(...); Link = @{ Column; Path; Location; Tooltip } }) })
    #       Values の数値（int / long / double）は数値のセルとして書く
    #       Values に @{ Runs = @(@{ Text; Style }) } を渡すと、部分ごとに書式を付けた文字列にする（Style: Deleted / Inserted / なし）
    #       Link は Values の Column 番目（0 始まり）のセルに付ける外部リンク。Path のファイルの Location（'シート'!セル）を開く
    #       Charts: そのシートの表から作るグラフ（ConvertTo-ChartXml）。表の右に縦に並べて置く
    param($Sheets, [string] $FilePath)

    $font = '<sz val="10"/><name val="游ゴシック"/><family val="3"/><charset val="128"/>'
    $sheetEntries = New-Object System.Collections.Generic.List[object]
    $definedNames = New-Object System.Text.StringBuilder
    $chartParts = [ordered]@{}   # パーツの名前 → XML（グラフと、その置き場所の描画）
    $chartCount = 0; $drawingCount = 0

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
        [void]$xml.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape"/>')
        $charts = @($s.Charts | Where-Object { $_ })
        if ($charts.Count -gt 0 -and $s.Rows.Count -gt 0) {
            # グラフは描画（drawing）に置き、描画からグラフを参照する
            $drawingCount++
            $anchors = New-Object System.Text.StringBuilder
            $drawingRels = New-Object System.Text.StringBuilder
            [void]$drawingRels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
            $top = 1
            for ($ci = 0; $ci -lt $charts.Count; $ci++) {
                $chartCount++
                $chart = $charts[$ci]
                $chartParts["xl/charts/chart$chartCount.xml"] = ConvertTo-ChartXml $chart $s.Name ($s.Rows.Count + 1)
                [void]$drawingRels.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart{1}.xml"/>' -f ($ci + 1), $chartCount))
                $width = 9; $height = 18
                if ($chart.Width) { $width = $chart.Width }
                if ($chart.Height) { $height = $chart.Height }
                $left = $colCount + 1
                [void]$anchors.Append(('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' -f $left, $top, ($left + $width), ($top + $height)))
                [void]$anchors.Append(('<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="{0}" name="{1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId{2}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>' -f ($ci + 2), (ConvertTo-XmlText $chart.Title), ($ci + 1)))
                $top += $height + 1
            }
            [void]$drawingRels.Append('</Relationships>')
            $chartParts["xl/drawings/drawing$drawingCount.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + $anchors.ToString() + '</xdr:wsDr>'
            $chartParts["xl/drawings/_rels/drawing$drawingCount.xml.rels"] = $drawingRels.ToString()
            [void]$rels.Append(('<Relationship Id="rIdD1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{0}.xml"/>' -f $drawingCount))
            [void]$xml.Append('<drawing r:id="rIdD1"/>')
            $hasDrawing = $true
        }
        else { $hasDrawing = $false }
        [void]$xml.Append('</worksheet>')
        [void]$rels.Append('</Relationships>')

        $sheetEntries.Add([pscustomobject]@{ Name = $s.Name; Xml = $xml.ToString(); Rels = $rels.ToString(); HasLinks = ($linkCount -gt 0 -or $hasDrawing) })
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
    foreach ($name in $chartParts.Keys) {
        if ($name -match '^xl/charts/.+\.xml$') { [void]$ct.Append(('<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>' -f $name)) }
        elseif ($name -match '^xl/drawings/[^/]+\.xml$') { [void]$ct.Append(('<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' -f $name)) }
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
    foreach ($name in $chartParts.Keys) { $parts[$name] = $chartParts[$name] }

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

function Read-ScreenDesign {
    #     画面設計書の画面項目定義・イベント定義と、コード定義書・メッセージ一覧を読む（Books は ConvertTo-DesignBooks の結果）。
    #     戻り値: @{ Items（画面項目。Input は入力のある項目）; Events（イベント）; Codes（コード区分名 → 表示名の一覧）;
    #             Messages（メッセージID（大文字）→ @{ Text; File; Sheet; Row }） }
    param($Books)
    $nfkc = [System.Text.NormalizationForm]::FormKC
    $items = New-Object System.Collections.Generic.List[object]
    $events = New-Object System.Collections.Generic.List[object]
    $codes = @{}
    $messages = @{}
    foreach ($book in $Books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # コード定義書: コード区分名ごとの表示名
            foreach ($t in (Find-DesignTables $grid @('コード区分名', '表示名') 'コード区分名')) {
                foreach ($dr in $t.Rows) {
                    $k = ConvertTo-Label ([string]$dr.Cells[$t.ColumnOf['コード区分名']])
                    $name = ([string]$dr.Cells[$t.ColumnOf['表示名']]).Trim()
                    if (-not $k -or -not $name) { continue }
                    if (-not $codes.ContainsKey($k)) { $codes[$k] = New-Object System.Collections.Generic.List[string] }
                    if (-not $codes[$k].Contains($name)) { $codes[$k].Add($name) }
                }
            }
            # メッセージ一覧: メッセージID → 文言
            foreach ($t in (Find-DesignTables $grid @('メッセージID', 'メッセージ') 'メッセージID')) {
                foreach ($dr in $t.Rows) {
                    $id = ([string]$dr.Cells[$t.ColumnOf['メッセージID']]).Normalize($nfkc).Trim().ToUpperInvariant()
                    $text = ([string]$dr.Cells[$t.ColumnOf['メッセージ']]).Normalize($nfkc).Trim()
                    # 画面設計書のメッセージ欄（ID だけを書いた列）は文言ではない
                    if ($id -notmatch '^[A-Z]+\d+$' -or -not $text -or $text -match '^[A-Za-z]+\d+$' -or $messages.ContainsKey($id)) { continue }
                    $messages[$id] = [pscustomobject]@{ Text = $text; File = $book.File; Sheet = $sheetName; Row = $dr.Row }
                }
            }
            $screen = Find-LabelValue $grid '画面ID'
            if (-not $screen) { continue }
            $screen = $screen.Trim()
            $screenName = Find-LabelValue $grid '画面名'
            if ($screenName) { $screenName = $screenName.Trim() }
            # 画面項目定義
            foreach ($t in (Find-DesignTables $grid @('項目名', '桁数') '項目名')) {
                $cell = { param($dr, $name) if ($t.ColumnOf.ContainsKey($name)) { ([string]$dr.Cells[$t.ColumnOf[$name]]).Normalize($nfkc).Trim() } else { '' } }
                foreach ($dr in $t.Rows) {
                    $io = & $cell $dr '入出力'
                    $digits = & $cell $dr '桁数'
                    $items.Add([pscustomobject]@{
                            Screen = $screen; ScreenName = $screenName; Item = ([string]$dr.Cells[$t.ColumnOf['項目名']]).Trim(); ItemId = (& $cell $dr '項目ID')
                            Kind = (& $cell $dr '区分'); Io = $io; Input = (-not $t.ColumnOf.ContainsKey('入出力') -or $io -match 'I')
                            Type = (& $cell $dr '型'); Digits = $(if ($digits -match '^\d+$') { [int]$digits } else { $null })
                            Required = ((& $cell $dr '必須') -match '^(○|〇|◯|必須|要|Y|YES|1)$'); Initial = (& $cell $dr '初期値')
                            Checks = (& $cell $dr '入力チェック'); Note = (& $cell $dr '備考')
                            File = $book.File; Sheet = $sheetName; Row = $dr.Row
                        })
                }
            }
            # イベント定義
            foreach ($t in (Find-DesignTables $grid @('イベントID', 'イベント名') 'イベントID')) {
                $cell = { param($dr, $name) if ($t.ColumnOf.ContainsKey($name)) { ([string]$dr.Cells[$t.ColumnOf[$name]]).Trim() } else { '' } }
                foreach ($dr in $t.Rows) {
                    $events.Add([pscustomobject]@{
                            Screen = $screen; ScreenName = $screenName; Id = (& $cell $dr 'イベントID'); Name = (& $cell $dr 'イベント名')
                            Trigger = (& $cell $dr 'トリガー'); Summary = (& $cell $dr '処理概要'); Messages = (& $cell $dr 'メッセージ'); Next = (& $cell $dr '遷移先')
                            File = $book.File; Sheet = $sheetName; Row = $dr.Row
                        })
                }
            }
        }
    }
    return [pscustomobject]@{ Items = $items; Events = $events; Codes = $codes; Messages = $messages }
}

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

# ---------------------------------------------------------------- 表を見出しの候補で読む（テスト仕様書・障害管理表など）


function ConvertTo-Hours {
    #     工数の文字（1.5・1.5h・1時間30分・1:30・90分）を時間にする（読めなければ $null）
    param([string] $Text)
    $t = ([string]$Text).Normalize($script:Nfkc).Trim() -replace '[,\s]', ''
    if (-not $t) { return $null }
    $invariant = [System.Globalization.CultureInfo]::InvariantCulture
    $m = [regex]::Match($t, '^(\d+):(\d{1,2})$')
    if ($m.Success) { return [double]$m.Groups[1].Value + [double]$m.Groups[2].Value / 60 }
    $m = [regex]::Match($t, '^(?:(\d+(?:\.\d+)?)時間)?(?:(\d+(?:\.\d+)?)分)?$')
    if ($m.Success -and ($m.Groups[1].Success -or $m.Groups[2].Success)) {
        $h = 0.0
        if ($m.Groups[1].Success) { $h += [double]::Parse($m.Groups[1].Value, $invariant) }
        if ($m.Groups[2].Success) { $h += [double]::Parse($m.Groups[2].Value, $invariant) / 60 }
        return $h
    }
    $t = $t -replace '(?i)(h|hr|hrs|時間)$', ''
    $n = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$n)) { return $n }
    return $null
}

function Read-EffortLog {
    #     工数実績の表（1 行が日付・担当・WBS 番号かタスク・工数）を読む（Measure-Effort.ps1・New-WeeklyReport.ps1）。
    #       Headers: 見出しの候補の CSV。Tasks: Measure-Progress.ps1 のタスク（WBS 番号、無ければタスク名で対応付ける）
    #       工数は時間で読む（見出しに「人日」があれば HoursPerDay を掛ける）。AsOf より後の日付の行は数えない
    #       日付・担当が空の行は同じ表の上の行のもの。担当の列が無い表は、シート名（表が複数のブック）かファイル名を担当にする
    #     戻り値: @{ Entries = @(@{ Date; Owner; WbsNo; TaskName; Hours; Task（Tasks の位置。対応しなければ -1）; Place }); Unreadable; Future }
    param($Inputs, [string] $Headers, $Tasks, [datetime] $AsOf, [double] $HoursPerDay = 8, $Skip, [string] $Exclude, [switch] $NoRecurse, [switch] $NoExcel, $ReadLog)
    $byWbs = @{}
    $byName = @{}
    $list = @($Tasks)
    for ($i = 0; $i -lt $list.Count; $i++) {
        if ($list[$i].WbsNo) { $byWbs[$list[$i].WbsNo] = $i }
        $k = ConvertTo-NameKey $list[$i].Task
        if ($byName.ContainsKey($k)) { $byName[$k] = -1 } else { $byName[$k] = $i }   # 同じ名前のタスクが複数あれば名前では対応付けない
    }
    $fields = @('日付', '担当', 'WBS番号', 'タスク', '工数')
    $records = Read-HeaderRecords -Inputs $Inputs -Candidates (Read-HeaderCandidates $Headers $fields) -Fields $fields -Need @('工数', @('WBS番号', 'タスク')) `
        -Key @('工数') -What '工数実績' -Skip $Skip -Exclude $Exclude -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $ReadLog
    $entries = New-Object System.Collections.Generic.List[object]
    $unreadable = 0; $future = 0
    $last = @{}   # 表ごとの前の行の日付・担当（結合したセルや、同じ日の 2 行目以降の空欄を埋める）
    foreach ($rec in $records) {
        $v = $rec.Values
        $hours = ConvertTo-Hours $v['工数']
        if ($null -eq $hours) { $unreadable++; continue }
        if ($rec.Labels['工数'] -match '人日') { $hours = $hours * $HoursPerDay }
        $table = '{0}|{1}' -f $rec.Path, $rec.Sheet
        if (-not $last.ContainsKey($table)) { $last[$table] = @{ Date = $null; Owner = '' } }
        $date = ConvertTo-DateValue $v['日付'] $AsOf
        if ($date) { $last[$table].Date = $date } elseif (-not ([string]$v['日付']).Trim()) { $date = $last[$table].Date }
        if ($date -and $date -gt $AsOf) { $future++; continue }
        $owner = ([string]$v['担当']).Trim()
        if ($owner) { $last[$table].Owner = $owner }
        elseif ($rec.Labels['担当']) { $owner = $last[$table].Owner }
        elseif ($rec.SheetCount -gt 1) { $owner = $rec.Sheet }
        else { $owner = [System.IO.Path]::GetFileNameWithoutExtension($rec.File) }
        $wbsNo = ((([string]$v['WBS番号']).Normalize($script:Nfkc).Trim() -replace '[\-_]', '.').TrimEnd('.'))
        $taskName = ([string]$v['タスク']).Trim()
        $task = -1
        if ($wbsNo -and $byWbs.ContainsKey($wbsNo)) { $task = $byWbs[$wbsNo] }
        elseif ($taskName -and $byName.ContainsKey((ConvertTo-NameKey $taskName))) { $task = $byName[(ConvertTo-NameKey $taskName)] }
        $cell = '{0} 行目' -f $rec.Row
        if ($rec.Sheet) { $cell = '{0}{1}' -f (ConvertTo-ColumnName $rec.Column), $rec.Row }
        $entries.Add([pscustomobject]@{
                Date = $date; Owner = $owner; WbsNo = $wbsNo; TaskName = $taskName; Hours = $hours; Task = $task
                Place = [pscustomobject]@{ File = $rec.File; Path = $rec.Path; Sheet = $rec.Sheet; Cell = $cell }
            })
    }
    return [pscustomobject]@{ Entries = $entries; Unreadable = $unreadable; Future = $future }
}

function Get-EffortJudgement {
    #     工数の判定: 予定超過（未完了で実績が予定を超えた）・超過の見込み（実績 ÷ 進捗率 が予定 ×（1 + Tolerance）を超える）・
    #     完了・超過（完了して実績が予定 ×（1 + Tolerance）を超えた）。予定が無ければ空
    param($Planned, [double] $Actual, [double] $Progress, [bool] $Done, [double] $Tolerance)
    if ($null -eq $Planned -or $Planned -le 0) { return '' }
    $limit = $Planned * (1 + $Tolerance)
    if ($Done) { if ($Actual -gt $limit) { return '完了・超過' } else { return '' } }
    if ($Actual -gt $Planned) { return '予定超過' }
    if ($Progress -gt 0 -and $Actual / $Progress -gt $limit) { return '超過の見込み' }
    return ''
}

function Read-HeaderCandidates {
    #     見出しの候補の CSV（項目・見出し。見出しは | 区切り）を、項目 → ConvertTo-Label 済みの候補の配列にする。
    #     Fields のうち CSV に無い項目は空の配列
    param([string] $FilePath, [string[]] $Fields)
    $candidates = @{}
    foreach ($r in (Read-CsvFile $FilePath @('項目', '見出し') '見出しの候補')) {
        $field = ([string]$r.'項目').Trim()
        if ($field) { $candidates[$field] = @(([string]$r.'見出し') -split '\|' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ }) }
    }
    foreach ($f in $Fields) { if (-not $candidates.ContainsKey($f)) { $candidates[$f] = @() } }
    return $candidates
}

function Get-InputFiles {
    #     ファイル・フォルダを、Excel（Search-Excel に渡すパス）と CSV に分ける
    param([string[]] $Paths, [string] $What, [switch] $NoRecurse)
    $excel = New-Object System.Collections.Generic.List[string]
    $csv = New-Object System.Collections.Generic.List[object]
    foreach ($p in $Paths) {
        $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
        if ($items.Count -eq 0) { Write-Warning "${What}が見つかりません: $p"; continue }
        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                if (@(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.xls[xmb]?$' -and -not $_.Name.StartsWith('~$') }).Count -gt 0) { $excel.Add($item.FullName) }
                foreach ($c in @(Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.csv' -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { $csv.Add($c) }
            }
            elseif ($item.Extension -eq '.csv') { $csv.Add($item) }
            else { $excel.Add([System.Management.Automation.WildcardPattern]::Escape($item.FullName)) }
        }
    }
    return [pscustomobject]@{ Excel = $excel; Csv = $csv }
}

function Get-FilePathSet {
    #     ファイル・フォルダの中のファイルのフルパス（大文字小文字を区別しない集合）
    param([string[]] $Paths, [switch] $NoRecurse)
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $Paths) {
        foreach ($item in @(Get-Item -Path $p -ErrorAction SilentlyContinue)) {
            if ($item.PSIsContainer) { foreach ($f in @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { [void]$set.Add($f.FullName) } }
            else { [void]$set.Add($item.FullName) }
        }
    }
    return , $set
}

function Read-HeaderRecords {
    #     見出しの候補で表を探し、1 行ずつ読む（Excel は Search-Excel.ps1 で、CSV はそのまま）。
    #       Candidates: 項目 → 見出しの候補（Read-HeaderCandidates）。Fields: 読む項目
    #       Need: 見出しの行の条件。文字は必ず要る項目、配列はそのうちどれか 1 つが要る項目
    #       Key: 行があるかを決める項目（前から順に、見出しにあるもの）
    #       Skip: 読まないファイルのフルパスの集合。Exclude: 読まないファイル（結果の xlsx など）
    #       ReadLog: 渡すと、ファイル・シートごとの読み取り結果（Add-ReadLog）を足す
    #     戻り値: @{ Values = 項目 → 文字; Labels = 項目 → 見出し; Columns = 項目 → 列番号（CSV は 0）; File; Path; Sheet; Row;
    #             Column（Key の列）; SheetCount（そのブックで表のあるシートの数） }
    param($Inputs, [hashtable] $Candidates, [string[]] $Fields, [object[]] $Need, [string[]] $Key, [string] $What, $Skip, [string] $Exclude, [switch] $NoRecurse, [switch] $NoExcel, $ReadLog)
    $records = New-Object System.Collections.Generic.List[object]
    $test = {
        param($m)
        foreach ($g in $Need) { if (@(@($g) | Where-Object { $m[$_] }).Count -eq 0) { return $false } }
        return $true
    }
    $findLabel = {
        param($labels, $names)
        foreach ($n in $names) { if ($labels -contains $n) { return $n } }
        return $null
    }
    if ($Inputs.Excel.Count -gt 0) {
        $searchArgs = @{ Pattern = '\S'; Path = $Inputs.Excel.ToArray(); Target = 'Value' }
        if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
        if ($NoExcel) { $searchArgs['NoExcel'] = $true }
        $hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object {
                -not ($Exclude -and [string]::Equals($_.Path, $Exclude, [System.StringComparison]::OrdinalIgnoreCase)) -and -not ($Skip -and $Skip.Contains($_.Path))
            })
        $books = ConvertTo-DesignBooks $hits
        foreach ($book in $books.Values) {
            # 先に表のあるシートを数える（シートごとに機能を分けたブックでは、シート名を機能にするため）
            $found = New-Object System.Collections.Generic.List[object]
            foreach ($sheetName in $book.Sheets.Keys) {
                $grid = $book.Sheets[$sheetName]
                foreach ($row in $grid.Rows) {
                    $labels = @($grid.Labels[$row].Values)
                    $m = @{}
                    foreach ($f in $Fields) { $l = & $findLabel $labels $Candidates[$f]; if ($l) { $m[$f] = $l } }
                    if (& $test $m) { $found.Add([pscustomobject]@{ Sheet = $sheetName; Grid = $grid; Map = $m }); break }
                }
            }
            if ($found.Count -eq 0) { Add-ReadLog $ReadLog $What $book.File '' $null 0 $null $Fields '表が見つかりません（必要な見出しの行がありません）' }
            foreach ($entry in $found) {
                $map = $entry.Map
                $keyField = @($Key | Where-Object { $map[$_] })[0]
                foreach ($t in (Find-DesignTables $entry.Grid @($map.Values | Sort-Object -Unique) $map[$keyField])) {
                    Add-ReadLog $ReadLog $What $book.File $entry.Sheet $t.HeaderRow $t.Rows.Count $map $Fields ''
                    $cols = @{}
                    foreach ($f in $map.Keys) { $cols[$f] = $t.ColumnOf[$map[$f]] }
                    foreach ($dr in $t.Rows) {
                        $v = @{}
                        foreach ($f in $map.Keys) { $v[$f] = [string]$dr.Cells[$cols[$f]] }
                        $records.Add([pscustomobject]@{ Values = $v; Labels = $map; Columns = $cols; File = $book.File; Path = $book.Path; Sheet = $entry.Sheet; Row = $dr.Row; Column = $cols[$keyField]; SheetCount = $found.Count })
                    }
                }
            }
        }
    }
    foreach ($c in $Inputs.Csv) {
        if ($Skip -and $Skip.Contains($c.FullName)) { continue }
        $rows = Read-CsvFile $c.FullName @() $What
        if ($rows.Count -eq 0) { continue }
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $map = @{}
        foreach ($f in $Fields) { foreach ($n in $names) { if (-not $map[$f] -and $Candidates[$f] -contains (ConvertTo-Label $n)) { $map[$f] = $n } } }
        if (-not (& $test $map)) {
            Write-Warning ("{0}: {1}の列が見つかりません" -f $c.Name, $What)
            Add-ReadLog $ReadLog $What $c.Name '' $null 0 $null $Fields '表が見つかりません（必要な見出しの列がありません）'
            continue
        }
        $keyField = @($Key | Where-Object { $map[$_] })[0]
        $before = $records.Count
        $line = 1
        foreach ($r in $rows) {
            $line++
            $v = @{}
            foreach ($f in $map.Keys) { $v[$f] = [string]$r.($map[$f]) }
            if (-not ([string]$v[$keyField]).Trim()) { continue }
            $records.Add([pscustomobject]@{ Values = $v; Labels = $map; Columns = @{}; File = $c.Name; Path = $c.FullName; Sheet = ''; Row = $line; Column = 0; SheetCount = 1 })
        }
        Add-ReadLog $ReadLog $What $c.Name '' 1 ($records.Count - $before) $map $Fields ''
    }
    return , $records
}

function Add-ReadLog {
    #     読み取り結果（どのファイル・シートの表を、どの見出しで何行読んだか）を ReadLog に足す。ReadLog が $null なら何もしない
    #       Map: 項目 → 読んだ見出し（表が見つからなければ $null）。Fields: 読もうとした項目
    param($ReadLog, [string] $What, [string] $File, [string] $Sheet, $HeaderRow, [int] $Rows, $Map, [string[]] $Fields, [string] $Note)
    if ($null -eq $ReadLog) { return }
    $found = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($f in $Fields) {
        if ($Map -and $Map[$f]) {
            if ($Map[$f] -eq (ConvertTo-Label $f)) { $found.Add($f) } else { $found.Add(('{0}（{1}）' -f $f, $Map[$f])) }
        }
        elseif ($Map) { $missing.Add($f) }
    }
    $ReadLog.Add([pscustomobject]@{ What = $What; File = $File; Sheet = $Sheet; HeaderRow = $HeaderRow; Rows = $Rows; Found = $found.ToArray(); Missing = $missing.ToArray(); Note = $Note })
}

function New-ReadLogSheet {
    #     読み取り結果のシート（Export-Workbook の 1 シートの定義）
    param($ReadLog)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($l in $ReadLog) {
        $header = ''
        if ($null -ne $l.HeaderRow) { $header = [double]$l.HeaderRow }
        $rows.Add(@{ Values = @($l.What, $l.File, $l.Sheet, $header, [double]$l.Rows, ($l.Found -join '、'), ($l.Missing -join '、'), $l.Note) })
    }
    return @{ Name = '読み取り結果'; Columns = @(@{ Header = '種類'; Width = 12 }, @{ Header = 'ファイル'; Width = 30 }, @{ Header = 'シート'; Width = 16 }, @{ Header = '見出しの行'; Width = 9 }, @{ Header = '読んだ行数'; Width = 9 }, @{ Header = '読んだ項目（見出し）'; Width = 50 }, @{ Header = '見出しが無い項目'; Width = 30 }, @{ Header = '備考'; Width = 36 }); Rows = $rows }
}

function Write-ReadLogSummary {
    #     読み取り結果の要約を表示する（表が見つからないファイルと、見出しが無い項目）
    param($ReadLog)
    foreach ($g in @($ReadLog | Group-Object What)) {
        $tables = @($g.Group | Where-Object { -not $_.Note })
        $files = @($tables | ForEach-Object { $_.File } | Select-Object -Unique)
        $notFound = @($g.Group | Where-Object { $_.Note } | ForEach-Object { $_.File } | Select-Object -Unique)
        $line = '読み取り（{0}）: {1} ファイル・{2} 行' -f $g.Name, $files.Count, (($tables | Measure-Object -Property Rows -Sum).Sum + 0)
        $missingCounts = [ordered]@{}
        foreach ($t in $tables) { foreach ($m in $t.Missing) { if (-not $missingCounts.Contains($m)) { $missingCounts[$m] = 0 }; $missingCounts[$m]++ } }
        if ($missingCounts.Count -gt 0) { $line += '。見出しが無い項目: ' + ((@($missingCounts.Keys | ForEach-Object { '{0}（{1}/{2} 表）' -f $_, $missingCounts[$_], $tables.Count })) -join '・') }
        Write-Host $line
        if ($notFound.Count -gt 0) { Write-Host ('  表が見つからないファイル: {0}' -f ($notFound -join '・')) }
    }
}

function Read-PhaseList {
    #     工程の CSV（工程・別名・摘出すべき工程）を読む。戻り値: @{ Name; Aliases（ConvertTo-Label 済み・大文字）; Expected; Index } の一覧（CSV の順）
    param([string] $FilePath)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in (Read-CsvFile $FilePath @('工程', '別名', '摘出すべき工程') '工程の CSV')) {
        $name = ([string]$r.'工程').Trim()
        if (-not $name) { continue }
        $aliases = @(@($name) + @(([string]$r.'別名') -split '\|') | ForEach-Object { (ConvertTo-Label $_).ToUpperInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
        $list.Add([pscustomobject]@{ Name = $name; Aliases = $aliases; Expected = ([string]$r.'摘出すべき工程').Trim(); Index = $list.Count })
    }
    if ($list.Count -eq 0) { throw "工程の CSV に工程がありません: $FilePath" }
    return , $list
}

function Resolve-PhaseName {
    #     工程の文字（「詳細設計レビュー」「UT」など）を、Read-PhaseList の工程にする。読めなければ $null
    param($PhaseList, [string] $Text)
    $t = (ConvertTo-Label $Text).ToUpperInvariant() -replace '(レビュー|工程|フェーズ|段階)$', ''
    if (-not $t) { return $null }
    foreach ($p in $PhaseList) { if ($p.Aliases -contains $t) { return $p } }
    # 英字だけの短い別名（UT・IT など）は完全一致だけ。日本語の別名は含んでいればよい（長い別名を優先）
    $best = $null; $bestLength = 0
    foreach ($p in $PhaseList) {
        foreach ($a in $p.Aliases) {
            if ($a -match '^[A-Z0-9]+$') { continue }
            if ($t.Contains($a) -and $a.Length -gt $bestLength) { $best = $p; $bestLength = $a.Length }
        }
    }
    return $best
}

function Read-SettingsFile {
    #     設定の CSV（項目・値）を、パラメーター → 値 にする。Path がフォルダなら、その中の FileName を読む。
    #       Keys: 項目 → @(パラメーター, 複数か, パスか)。値の相対パスは設定ファイルのフォルダからのパスとし、複数の値は | で区切る
    #       What: 設定の名前（メッセージ用）。Template: 写して使うひな形（見つからないときのメッセージ用）
    #     戻り値: @{ Settings = パラメーター → 値; BaseDir = 設定ファイルのフォルダ }
    param([string] $Path, [string] $FileName, $Keys, [string] $What, [string] $Template)
    $file = $Path
    if (Test-Path -LiteralPath $Path -PathType Container) { $file = Join-Path $Path $FileName }
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "${What}ファイルが見つかりません: $file（$Template を写して書いてください）" }
    $file = (Resolve-Path -LiteralPath $file).ProviderPath
    $base = Split-Path -Path $file -Parent
    $settings = [ordered]@{}
    foreach ($r in (Read-CsvFile $file @('項目', '値') $What)) {
        $key = ([string]$r.'項目').Trim()
        $value = ([string]$r.'値').Trim()
        if (-not $key -or -not $value) { continue }
        if (-not $Keys.Contains($key)) { throw ("{0}に知らない項目があります: {1}（使える項目: {2}）: {3}" -f $What, $key, (@($Keys.Keys) -join '・'), $file) }
        $def = $Keys[$key]
        $values = @($value)
        if ($def[1]) { $values = @($value -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
        if ($def[2]) {
            # Windows の区切り（\）で書いた相対パスを、ほかの OS でも読めるようにする
            if ([System.IO.Path]::DirectorySeparatorChar -eq '/') { $values = @($values | ForEach-Object { $_.Replace('\', '/') }) }
            $values = @($values | ForEach-Object { if ([System.IO.Path]::IsPathRooted($_)) { $_ } else { [System.IO.Path]::GetFullPath((Join-Path $base $_)) } })
        }
        if ($def[1]) { $settings[$def[0]] = [string[]]$values } else { $settings[$def[0]] = [string]$values[0] }
    }
    return [pscustomobject]@{ Settings = $settings; BaseDir = $base }
}

function Read-QualityReportConfig {
    #     品質評価の設定の CSV（項目・値）を、New-QualityReport.ps1 のパラメーター → 値 にする。
    #     Path がフォルダなら、その中の「品質評価の設定.csv」を読む。値の相対パスは設定ファイルのフォルダからのパスとし、
    #     | で区切って複数書ける。出力先が無ければ設定ファイルのフォルダの「品質報告」、プロジェクトが無ければそのフォルダの名前
    param([string] $Path)
    # 項目 → パラメーター・複数か・パスか
    $keys = [ordered]@{
        '出力先' = @('OutDir', $false, $true); 'タイトル' = @('Title', $false, $false); 'プロジェクト' = @('Project', $false, $false)
        '設計書' = @('DesignPath', $true, $true); '指摘管理表' = @('Findings', $true, $true); 'AIレビュー' = @('AiReview', $true, $true); 'レビュー記録' = @('ReviewLog', $true, $true)
        'ページ数' = @('PageCsv', $false, $true); 'テスト仕様書' = @('TestPath', $true, $true); '障害管理表' = @('BugList', $true, $true)
        '要件一覧' = @('Requirement', $true, $true); '規模' = @('SizeCsv', $false, $true); '機能の対応表' = @('FunctionMap', $false, $true)
        'レビューの基準値' = @('ReviewLimits', $false, $true); 'テストの基準値' = @('TestLimits', $false, $true); '判定基準' = @('Criteria', $false, $true)
        '実績' = @('HistoryCsv', $false, $true)
    }
    $read = Read-SettingsFile $Path '品質評価の設定.csv' $keys '品質評価の設定' 'tools\quality-report\品質評価の設定_ひな形.csv'
    $settings = $read.Settings
    if (-not $settings.Contains('OutDir')) { $settings['OutDir'] = Join-Path $read.BaseDir '品質報告' }
    if (-not $settings.Contains('Project')) { $settings['Project'] = Split-Path -Path $read.BaseDir -Leaf }
    return $settings
}

function Read-WeeklyReportConfig {
    #     週次報告の設定の CSV（項目・値）を、New-WeeklyReport.ps1 のパラメーター → 値 にする。
    #     Path がフォルダなら、その中の「週次報告の設定.csv」を読む。出力先が無ければ設定ファイルのフォルダの「週次報告」
    param([string] $Path)
    $keys = [ordered]@{
        '出力先' = @('OutDir', $false, $true); 'タイトル' = @('Title', $false, $false); 'WBS' = @('Wbs', $true, $true); '休日' = @('Holidays', $false, $true)
        '指摘管理表' = @('Findings', $true, $true); '障害管理表' = @('BugList', $true, $true); '課題管理表' = @('IssueList', $true, $true); '工数実績' = @('EffortLog', $true, $true)
        '工数の単位' = @('Unit', $false, $false); '1日の時間' = @('HoursPerDay', $false, $false); '実績' = @('HistoryCsv', $false, $true)
    }
    $read = Read-SettingsFile $Path '週次報告の設定.csv' $keys '週次報告の設定' 'tools\weekly-report\週次報告の設定_ひな形.csv'
    $settings = $read.Settings
    if (-not $settings.Contains('OutDir')) { $settings['OutDir'] = Join-Path $read.BaseDir '週次報告' }
    return $settings
}

function Read-FunctionMap {
    #     機能の対応表の CSV（機能・別名・設計書。別名と設計書は | で区切って複数書ける）を、名前（ConvertTo-NameKey）→ 機能の名前 にする
    #     テスト仕様書・障害管理表・規模・設計書で機能の書き方が違うとき（SCR001 / 受注一覧 / DD02_画面設計書_SCR001_受注一覧 など）にそろえる
    param([string] $FilePath)
    $map = @{}
    if (-not $FilePath) { return $map }
    foreach ($r in (Read-CsvFile $FilePath @('機能') '機能の対応表')) {
        $name = ([string]$r.'機能').Trim()
        if (-not $name) { continue }
        $names = @($name)
        foreach ($col in @('別名', '設計書')) { if ($r.PSObject.Properties[$col]) { $names += @(([string]$r.$col) -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) } }
        foreach ($n in $names) {
            $k = ConvertTo-NameKey $n
            if ($map.ContainsKey($k) -and $map[$k] -ne $name) { Write-Warning ("機能の対応表で「{0}」が 2 つの機能（{1}・{2}）にあります。先の行を使います: {3}" -f $n, $map[$k], $name, $FilePath); continue }
            $map[$k] = $name
        }
    }
    return $map
}

function Resolve-FunctionName {
    #     名前を機能の対応表の機能の名前にする（対応表に無ければ $null）
    param([hashtable] $Map, [string] $Name)
    if (-not $Map -or $Map.Count -eq 0 -or -not ([string]$Name).Trim()) { return $null }
    $k = ConvertTo-NameKey $Name
    if ($Map.ContainsKey($k)) { return $Map[$k] }
    return $null
}

function ConvertTo-NameKey {
    #     機能・設計書の名前を比べる形（フォルダと拡張子を除き、全角・半角と空白をそろえる）
    param([string] $Name)
    $n = ([string]$Name).Trim()
    $n = $n.Split(@('\', '/'))[-1]
    if ($n -match '\.(xlsx|xlsm|xls|xlsb|csv|pdf|docx?)$') { $n = [System.IO.Path]::GetFileNameWithoutExtension($n) }
    return (ConvertTo-Label $n).ToUpperInvariant()
}

function ConvertTo-DateValue {
    #     日付の文字（2026/10/1・2026-10-01・2026年10月1日・10/1）を日付にする（無ければ $null）。年の無い日付は AsOf の年
    param([string] $Text, [datetime] $AsOf = (Get-Date))
    $t = ([string]$Text).Normalize($script:Nfkc).Trim()
    if (-not $t) { return $null }
    $m = [regex]::Match($t, '(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})')
    try {
        if ($m.Success) { return New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value) }
        $m = [regex]::Match($t, '^(\d{1,2})\s*[/月]\s*(\d{1,2})日?$')
        if ($m.Success) { return New-Object DateTime $AsOf.Year, ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value) }
    }
    catch { }
    return $null
}

function Get-TestResultKind {
    #     テストの結果を OK・NG・保留・対象外・未実施にする（読めない値は未実施）
    param([string] $Text)
    $t = ([string]$Text).Normalize($script:Nfkc).Trim().ToUpperInvariant()
    if (-not $t) { return '未実施' }
    if ($t -match '^(OK|○|〇|合格|PASS|良)') { return 'OK' }
    if ($t -match '^(NG|×|✕|不合格|FAIL|不可)') { return 'NG' }
    if ($t -match '^(対象外|-|―|N/?A|スキップ|SKIP|削除|欠番)') { return '対象外' }
    if ($t -match '保留|中断|ブロック|BLOCK') { return '保留' }
    return '未実施'
}

function ConvertTo-BugIds {
    #     バグ番号のセル（「B-001、B-002」「#12」など）を、比べる形の番号の並びにする
    param([string] $Text)
    return @(([string]$Text).Normalize($script:Nfkc) -split '[、,;/\s↵]+' | ForEach-Object { $_.Trim().TrimStart('#').ToUpperInvariant() } | Where-Object { $_ -and $_ -notmatch '^[-―]$' })
}

function ConvertTo-TestCases {
    #     Read-HeaderRecords で読んだテストケースの行に、機能を付ける。
    #     機能は、機能の列（グループの最初の行だけに書くことがあるので、空なら上の行の値）、
    #     無ければシート名（表のあるシートが複数のブック）かファイル名
    param($Records, [datetime] $AsOf = (Get-Date))
    $cases = New-Object System.Collections.Generic.List[object]
    $lastFunction = @{}
    foreach ($rec in $Records) {
        $v = $rec.Values
        $function = ([string]$v['機能']).Trim()
        $tableKey = $rec.Path + '|' + $rec.Sheet
        if ($rec.Labels['機能']) {
            if ($function) { $lastFunction[$tableKey] = $function }
            elseif ($lastFunction.ContainsKey($tableKey)) { $function = $lastFunction[$tableKey] }
        }
        if (-not $function) {
            if ($rec.SheetCount -gt 1) { $function = $rec.Sheet }
            else { $function = [System.IO.Path]::GetFileNameWithoutExtension($rec.File) }
        }
        $id = ([string]$v['テストID']).Trim()
        if (-not $id) { $id = '{0} 行目' -f $rec.Row }
        $cases.Add([pscustomobject]@{
            Function = $function; Key = (ConvertTo-NameKey $function); Id = $id; Kind = (Get-TestResultKind $v['結果'])
            Planned = (ConvertTo-DateValue $v['予定日'] $AsOf); Done = (ConvertTo-DateValue $v['実施日'] $AsOf); BugIds = @(ConvertTo-BugIds $v['バグ番号'])
            Process = ([string]$v['工程']).Trim(); HasBugColumn = [bool]$rec.Labels['バグ番号']; HasDoneColumn = [bool]$rec.Labels['実施日']
            Values = $v; Place = $rec
        })
    }
    return , $cases
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

function Get-NameVersion {
    #     ファイル名・フォルダ名の版数（_v1.2・Ver1.2・第1.2版）を比べられる形（System.Version）にする。版数が無ければ $null
    param([string] $Name)
    $m = [regex]::Match($Name.Normalize($script:Nfkc).ToLowerInvariant(), '(?:(?<![a-z])v(?:er)?\.?\s*|第\s*)(\d+(?:\.\d+){0,3})')
    if (-not $m.Success) { return $null }
    $parts = @($m.Groups[1].Value.Split('.') | ForEach-Object { [int]$_ })
    while ($parts.Count -lt 4) { $parts += 0 }
    return New-Object System.Version($parts[0], $parts[1], $parts[2], $parts[3])
}

function Get-NameVersionText {
    #     ファイル名・フォルダ名の版数の部分（v1.2・Ver1.2・第1.2版）。版数が無ければ名前（拡張子を除く）
    param([string] $Name)
    $m = [regex]::Match($Name, '(?:(?<![A-Za-z])[vVｖＶ](?:[eE][rR])?\.?\s*|第\s*)[\d０-９]+(?:[.．][\d０-９]+){0,3}(?:版)?')
    if ($m.Success) { return $m.Value.Trim() }
    return [System.IO.Path]::GetFileNameWithoutExtension($Name)
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
}

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

# ---------------------------------------------------------------- Word（.docx / .docm）の本文（Search-Excel・Test-ExcelDeliverable）


function New-WordNamespaceManager {
    #     Word の XML を XPath で読むための名前空間（w・mc）
    param($Doc)
    $nsm = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $nsm.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main')
    $nsm.AddNamespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
    return , $nsm   # 列挙できる型なので、そのまま返すと名前空間の一覧に展開されてしまう
}

function Get-WordParts {
    #     本文・ヘッダー・フッター・脚注・文末脚注の XML（Name・Part・Doc）。本文が無ければ例外
    param($Package)
    $doc = Read-PartXml -Package $Package -PartName 'word/document.xml'
    if ($null -eq $doc) { throw 'Word の文書ではありません（word/document.xml が見つかりません）' }
    $parts = New-Object System.Collections.Generic.List[object]
    $parts.Add([pscustomobject]@{ Name = '本文'; Part = 'word/document.xml'; Doc = $doc })
    foreach ($e in @($Package.Entries.Keys | Sort-Object)) {
        $name = $null
        if ($e -match '^word/header\d*\.xml$') { $name = 'ヘッダー' }
        elseif ($e -match '^word/footer\d*\.xml$') { $name = 'フッター' }
        elseif ($e -eq 'word/footnotes.xml') { $name = '脚注' }
        elseif ($e -eq 'word/endnotes.xml') { $name = '文末脚注' }
        if (-not $name) { continue }
        $d = Read-PartXml -Package $Package -PartName $e
        if ($d) { $parts.Add([pscustomobject]@{ Name = $name; Part = $e; Doc = $d }) }
    }
    return , $parts
}

function Get-WordParagraphText {
    #     段落（w:p）の文字。テキストボックス（w:txbxContent）の中と、互換用の写し（mc:Fallback）・段落の書式（タブ位置）は含めない。
    #     削除した変更履歴の文字（w:delText）は含めない
    param($Paragraph, $Nsm)
    $depth = $Paragraph.SelectNodes('ancestor::w:txbxContent', $Nsm).Count
    $fallback = $Paragraph.SelectNodes('ancestor::mc:Fallback', $Nsm).Count
    $xpath = './/*[(self::w:t or self::w:tab or self::w:br or self::w:cr) and not(ancestor::w:pPr) and count(ancestor::w:txbxContent) = {0} and count(ancestor::mc:Fallback) = {1}]' -f $depth, $fallback
    $sb = New-Object System.Text.StringBuilder
    foreach ($n in $Paragraph.SelectNodes($xpath, $Nsm)) {
        if ($n.LocalName -eq 't') { [void]$sb.Append($n.InnerText) }
        elseif ($n.LocalName -eq 'tab') { [void]$sb.Append("`t") }
        else { [void]$sb.Append("`n") }
    }
    return $sb.ToString().Trim()
}

function Get-WordBlocks {
    #     本文・ヘッダーなどの段落と表の行を、文書の順に場所付きで返す（内容コントロール w:sdt の中も。脚注は区切り線などを除く）。
    #     Cell は「段落 n」（表の外の段落を上から数えた番号。空の段落も数える）と「表 n の m 行目」。どのツールでも同じ場所を指す。
    #     Text は段落の文字、表の行はセルの文字を「 | 」でつないだもの（1 行）。Cells はセルごとの文字（セルの中の段落は改行でつなぐ）、Node は w:p・w:tr
    param($Doc, $Nsm)
    $wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $root = $Doc.DocumentElement
    if ($root.LocalName -eq 'document') { $containers = @($root.SelectNodes('w:body', $Nsm)) }
    elseif ($root.LocalName -eq 'footnotes' -or $root.LocalName -eq 'endnotes') { $containers = @($root.SelectNodes("*[not(@w:type) or @w:type='normal']", $Nsm)) }
    else { $containers = @($root) }
    $stack = New-Object System.Collections.Generic.Stack[object]
    for ($q = $containers.Count - 1; $q -ge 0; $q--) {
        $children = @($containers[$q].ChildNodes)
        for ($i = $children.Count - 1; $i -ge 0; $i--) { $stack.Push($children[$i]) }
    }
    $paragraphNo = 0; $tableNo = 0
    while ($stack.Count -gt 0) {
        $node = $stack.Pop()
        if ($node.NamespaceURI -ne $wordNs) { continue }
        if ($node.LocalName -eq 'sdt') {
            $content = $node.SelectSingleNode('w:sdtContent', $Nsm)
            if ($content) { $children = @($content.ChildNodes); for ($i = $children.Count - 1; $i -ge 0; $i--) { $stack.Push($children[$i]) } }
        }
        elseif ($node.LocalName -eq 'p') {
            $paragraphNo++
            $text = Get-WordParagraphText $node $Nsm
            [pscustomobject]@{ Kind = '段落'; Cell = ('段落 {0}' -f $paragraphNo); Text = $text; Cells = @($text); Node = $node }
        }
        elseif ($node.LocalName -eq 'tbl') {
            $tableNo++; $rowNo = 0
            foreach ($tr in $node.SelectNodes('w:tr', $Nsm)) {
                $rowNo++
                $cells = @(foreach ($tc in $tr.SelectNodes('w:tc | w:sdt/w:sdtContent/w:tc', $Nsm)) {
                        # セルの中のテキストボックスは Get-WordTextBoxes で別に読む
                        (@(foreach ($p in $tc.SelectNodes('.//w:p[not(ancestor::w:txbxContent) and not(ancestor::mc:Fallback)]', $Nsm)) { Get-WordParagraphText $p $Nsm }) | Where-Object { $_ }) -join "`n"
                    })
                $text = (@($cells | ForEach-Object { $_ -replace '\s*\n\s*', ' ' })) -join ' | '
                [pscustomobject]@{ Kind = '表'; Cell = ('表 {0} の {1} 行目' -f $tableNo, $rowNo); Text = $text; Cells = $cells; Node = $tr }
            }
        }
    }
}

function Get-WordTextBoxText {
    #     テキストボックス（w:txbxContent）の中の段落の文字を改行でつなぐ（中にある別のテキストボックスの文字は含めない）
    param($Box, $Nsm)
    $depth = $Box.SelectNodes('ancestor-or-self::w:txbxContent', $Nsm).Count
    $fallback = $Box.SelectNodes('ancestor::mc:Fallback', $Nsm).Count
    $xpath = './/w:p[count(ancestor::w:txbxContent) = {0} and count(ancestor::mc:Fallback) = {1}]' -f $depth, $fallback
    return (@(foreach ($p in $Box.SelectNodes($xpath, $Nsm)) { Get-WordParagraphText $p $Nsm }) | Where-Object { $_ }) -join "`n"
}

function Get-WordTextBoxes {
    #     テキストボックスの文字（互換用の写し mc:Fallback は除く）。Name は「テキストボックス n」
    param($Doc, $Nsm)
    $n = 0
    foreach ($box in $Doc.SelectNodes('//w:txbxContent[not(ancestor::mc:Fallback)]', $Nsm)) {
        $n++
        [pscustomobject]@{ Name = ('テキストボックス {0}' -f $n); Text = (Get-WordTextBoxText $box $Nsm) }
    }
}

# ---------------------------------------------------------------- ショートカット


function Add-ShellLinkType {
    #     ショートカット（.lnk）を読み書きする型（ExcelTools.ShellLink）を読み込む（Windows だけ。Install-SendTo.ps1・Install-ExcelTools.ps1 が使う）
    # ショートカット（.lnk）の読み書き。WScript.Shell はファイル名を ANSI のコードページで扱うため、
    # 日本語の名前は日本語版以外の Windows で保存できない。Unicode の IShellLinkW を直接使う
    if (-not ('ExcelTools.ShellLink' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ExcelTools {
    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellLinkW {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    class CShellLink { }

    public class ShellLinkInfo {
        public string TargetPath;
        public string Arguments;
        public string WorkingDirectory;
        public string Description;
    }

    public static class ShellLink {
        public static void Save(string path, string targetPath, string arguments, string workingDirectory, string description) {
            IShellLinkW link = (IShellLinkW)new CShellLink();
            try {
                link.SetPath(targetPath);
                link.SetArguments(arguments);
                link.SetWorkingDirectory(workingDirectory);
                link.SetDescription(description);
                ((IPersistFile)link).Save(path, true);
            }
            finally { Marshal.FinalReleaseComObject(link); }
        }

        public static ShellLinkInfo Read(string path) {
            IShellLinkW link = (IShellLinkW)new CShellLink();
            try {
                ((IPersistFile)link).Load(path, 0);
                ShellLinkInfo info = new ShellLinkInfo();
                StringBuilder sb = new StringBuilder(32768);
                link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
                info.TargetPath = sb.ToString();
                sb.Length = 0;
                link.GetArguments(sb, sb.Capacity);
                info.Arguments = sb.ToString();
                sb.Length = 0;
                link.GetWorkingDirectory(sb, sb.Capacity);
                info.WorkingDirectory = sb.ToString();
                sb.Length = 0;
                link.GetDescription(sb, sb.Capacity);
                info.Description = sb.ToString();
                return info;
            }
            finally { Marshal.FinalReleaseComObject(link); }
        }
    }
}
'@
    }
}

Export-ModuleMember -Function *
