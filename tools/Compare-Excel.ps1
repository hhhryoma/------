[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Old,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $New,

    [string[]] $KeyColumn,

    [switch] $NoRowAlign,

    [switch] $NoColumnAlign,

    [switch] $CompareFormat,

    [switch] $CheckHistory,

    [string] $HistorySheet = '改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録',

    [string] $OutHistory,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx,

    [string] $OutMarkedDir
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$maxAlignCells = 1000000   # 行の対応付けを厳密な LCS で行う「旧の行数 × 新の行数」の上限（超えたら patience diff）
$pairWindow = 50           # 一致しなかった行どうしを「変更された行」として対応付けるとき、先を探す行数

if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}
if ($OutMarkedDir) { $OutMarkedDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutMarkedDir) }
if ($OutHistory) { $OutHistory = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutHistory) }

Update-TypeData -TypeName 'CompareExcel.Diff' -DefaultDisplayPropertySet File, Location, Change, Detail -Force -WhatIf:$false

# ---------------------------------------------------------------- ブックの読み取り


function Read-Book {
    #     Search-Excel.ps1 でブックの中身（値・数式・コメント・図形・シート名）を読み、シートごとに整理する。
    #     Rows のキーは行番号の文字列（[ordered] に整数で引くと位置指定になるため）
    param([string] $FilePath)

    $searchArgs = @{
        Pattern = '[\s\S]'
        Path    = [System.Management.Automation.WildcardPattern]::Escape($FilePath)
        Target  = @('Value', 'Formula', 'Comment', 'Shape', 'SheetName')
    }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $w = $null
    $hits = @(& $searchScript @searchArgs -WarningVariable w -WarningAction SilentlyContinue 6>$null)
    if ($w) { throw (@($w | ForEach-Object { $_.Message }) -join ' / ') }

    $book = [ordered]@{}
    foreach ($h in $hits) {
        if (-not $book.Contains($h.Sheet)) {
            $book[$h.Sheet] = [pscustomobject]@{ Rows = @{}; Comments = @{}; Shapes = @{}; Formats = @{} }
        }
        $s = $book[$h.Sheet]
        switch ($h.Kind) {
            { $_ -eq '値' -or $_ -eq '数式' } {
                $m = [regex]::Match($h.Cell, '^([A-Z]+)(\d+)$')
                $row = $m.Groups[2].Value
                if (-not $s.Rows.ContainsKey($row)) { $s.Rows[$row] = @{ V = @{}; F = @{} } }
                if ($h.Kind -eq '値') { $s.Rows[$row].V[$m.Groups[1].Value] = $h.Text }
                else { $s.Rows[$row].F[$m.Groups[1].Value] = $h.Text }
            }
            'コメント' { $s.Comments[$h.Cell] = $h.Text }
            '図形' {
                $key = $h.Name
                if (-not $key) { $key = $h.Cell }
                $s.Shapes[$key] = [pscustomobject]@{ Text = $h.Text; Cell = $h.Cell }
            }
        }
    }

    if ($CompareFormat) {
        $formats = $null
        try { $formats = Read-BookFormat $FilePath }
        catch { Write-Warning ("{0}: 書式を読めませんでした: {1}" -f [System.IO.Path]::GetFileName($FilePath), $_.Exception.Message) }
        if ($null -eq $formats) { Write-Warning ("{0}: 書式の比較は .xlsx / .xlsm だけが対象です" -f [System.IO.Path]::GetFileName($FilePath)) }
        else { foreach ($name in $book.Keys) { if ($formats.ContainsKey($name)) { $book[$name].Formats = $formats[$name] } } }
    }
    return $book
}

# ---------------------------------------------------------------- 書式の読み取り（-CompareFormat）


$script:FormatSep = [string][char]0x1F
$script:FormatOrder = @('太字', '斜体', '下線', '取り消し線', '文字色', '塗りつぶし', '部分書式')
$script:ThemeColorNames = @('背景1', 'テキスト1', '背景2', 'テキスト2', 'アクセント1', 'アクセント2', 'アクセント3', 'アクセント4', 'アクセント5', 'アクセント6', 'ハイパーリンク', '表示済みのハイパーリンク')
# 既定のカラーパレット（indexed の 0～63。64・65 はシステムの前景色・背景色）
$script:IndexedColors = @(
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333')
# Excel の「標準の色」と黒・白は名前で表す
$script:ColorNames = @{
    'C00000' = '濃い赤'; 'FF0000' = '赤'; 'FFC000' = 'オレンジ'; 'FFFF00' = '黄'; '92D050' = '薄い緑'
    '00B050' = '緑'; '00B0F0' = '薄い青'; '0070C0' = '青'; '002060' = '濃い青'; '7030A0' = '紫'
    '000000' = '黒'; 'FFFFFF' = '白'
}

function Get-ColorText {
    #     色の指定（rgb / theme / indexed / auto）を「赤」「#1F4E79」「アクセント1 +40%」のような文字にする。自動は ''
    param([string] $Rgb, [string] $Theme, [string] $Indexed, [string] $Tint, [string] $Auto)

    if ($Auto -eq '1' -or $Auto -eq 'true') { return '' }
    $n = 0
    $hex = $null
    if ($Rgb -and $Rgb.Length -ge 6) { $hex = $Rgb.Substring($Rgb.Length - 6).ToUpperInvariant() }
    elseif ($Indexed -and [int]::TryParse($Indexed, [ref]$n)) {
        if ($n -ge $script:IndexedColors.Count) { return '' }
        $hex = $script:IndexedColors[$n]
    }
    elseif ($Theme -and [int]::TryParse($Theme, [ref]$n)) {
        $name = "テーマの色$n"
        if ($n -lt $script:ThemeColorNames.Count) { $name = $script:ThemeColorNames[$n] }
        $tint = 0.0
        if ($Tint -and [double]::TryParse($Tint, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$tint) -and [math]::Abs($tint) -ge 0.005) {
            # 正は明るく、負は暗くした色
            $sign = ''
            if ($tint -gt 0) { $sign = '+' }
            $name += ' {0}{1}%' -f $sign, [math]::Round($tint * 100)
        }
        return $name
    }
    if (-not $hex) { return '' }
    if ($script:ColorNames.ContainsKey($hex)) { return $script:ColorNames[$hex] }
    return "#$hex"
}

function Set-FontAttribute {
    #     フォントの子要素（b / i / u / strike / color）1 つを、書式の辞書に反映する
    param([hashtable] $Format, [string] $Name, [string] $Val, [string] $Rgb, [string] $Theme, [string] $Indexed, [string] $Tint, [string] $Auto)

    $on = -not ($Val -eq '0' -or $Val -eq 'false')
    switch ($Name) {
        'b' { if ($on) { $Format['太字'] = 'あり' } }
        'i' { if ($on) { $Format['斜体'] = 'あり' } }
        'strike' { if ($on) { $Format['取り消し線'] = 'あり' } }
        'u' {
            if ($Val -eq 'double' -or $Val -eq 'doubleAccounting') { $Format['下線'] = '二重' }
            elseif ($Val -ne 'none') { $Format['下線'] = 'あり' }
        }
        'color' {
            $c = Get-ColorText -Rgb $Rgb -Theme $Theme -Indexed $Indexed -Tint $Tint -Auto $Auto
            # 自動・黒・テキスト1 は見た目が同じなので、指定なしと同じに扱う
            if ($c -and $c -ne '黒' -and $c -ne 'テキスト1') { $Format['文字色'] = $c }
        }
    }
}

function ConvertTo-FormatKey {
    #     書式の辞書を、比較に使う 1 つの文字列にする（標準の書式は ''）
    param([hashtable] $Format)
    $parts = @()
    foreach ($k in $script:FormatOrder) { if ($Format.ContainsKey($k)) { $parts += ('{0}={1}' -f $k, $Format[$k]) } }
    return ($parts -join $script:FormatSep)
}

function ConvertFrom-FormatKey {
    param([string] $Key)
    $f = @{}
    if (-not $Key) { return $f }
    foreach ($part in $Key.Split($script:FormatSep)) {
        $i = $part.IndexOf('=')
        $f[$part.Substring(0, $i)] = $part.Substring($i + 1)
    }
    return $f
}

function Format-FontText {
    #     文字の書式を「太字・赤」の形にする（部分書式の表示用）
    param([hashtable] $Format)
    $parts = @()
    foreach ($k in @('太字', '斜体', '下線', '取り消し線')) {
        if (-not $Format.ContainsKey($k)) { continue }
        if ($k -eq '下線' -and $Format[$k] -eq '二重') { $parts += '二重下線' } else { $parts += $k }
    }
    if ($Format.ContainsKey('文字色')) { $parts += $Format['文字色'] }
    return ($parts -join '・')
}

function Format-FormatText {
    #     セルの書式を「太字、文字色 赤、塗りつぶし 黄」の形にする（標準なら「標準」）
    param([hashtable] $Format)
    $parts = @()
    foreach ($k in $script:FormatOrder) {
        if (-not $Format.ContainsKey($k)) { continue }
        switch ($k) {
            '下線' { if ($Format[$k] -eq '二重') { $parts += '二重下線' } else { $parts += '下線' } }
            { @('太字', '斜体', '取り消し線') -contains $_ } { $parts += $k }
            default { $parts += ('{0} {1}' -f $k, $Format[$k]) }
        }
    }
    if ($parts.Count -eq 0) { return '標準' }
    return ($parts -join '、')
}

function Get-FormatValueText {
    #     書式の 1 項目の値を表示用にする（指定なしは「なし」、文字色だけ「自動」）
    param([string] $Name, [string] $Value)
    if ($Value) { return $Value }
    if ($Name -eq '文字色') { return '自動' }
    return 'なし'
}

function Read-StyleFormats {
    #     styles.xml の cellXfs を読み、スタイル番号ごとの書式のキー（ConvertTo-FormatKey）を返す
    param($Package, [string] $PartName)

    $result = New-Object System.Collections.Generic.List[string]
    if (-not $PartName) { return , $result }
    $doc = Read-PartXml -Package $Package -PartName $PartName
    if ($null -eq $doc) { return , $result }

    $fonts = New-Object System.Collections.Generic.List[object]
    foreach ($font in $doc.SelectNodes("//*[local-name()='fonts']/*[local-name()='font']")) {
        $f = @{}
        foreach ($ch in $font.ChildNodes) {
            if ($ch.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            Set-FontAttribute $f $ch.LocalName $ch.GetAttribute('val') $ch.GetAttribute('rgb') $ch.GetAttribute('theme') $ch.GetAttribute('indexed') $ch.GetAttribute('tint') $ch.GetAttribute('auto')
        }
        $fonts.Add($f)
    }
    $fills = New-Object System.Collections.Generic.List[string]
    foreach ($fill in $doc.SelectNodes("//*[local-name()='fills']/*[local-name()='fill']")) {
        $text = ''
        $pf = $fill.SelectSingleNode("*[local-name()='patternFill']")
        if ($pf) {
            $type = $pf.GetAttribute('patternType')
            if ($type -and $type -ne 'none' -and $type -ne 'gray125') {
                $fg = $pf.SelectSingleNode("*[local-name()='fgColor']")
                if ($fg) { $text = Get-ColorText $fg.GetAttribute('rgb') $fg.GetAttribute('theme') $fg.GetAttribute('indexed') $fg.GetAttribute('tint') $fg.GetAttribute('auto') }
                if (-not $text) { $text = '自動' }
                if ($type -ne 'solid') { $text += "（模様 $type）" }
            }
        }
        elseif ($fill.SelectSingleNode("*[local-name()='gradientFill']")) { $text = 'グラデーション' }
        $fills.Add($text)
    }
    foreach ($xf in $doc.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
        $fontId = 0; $fillId = 0
        [void][int]::TryParse($xf.GetAttribute('fontId'), [ref]$fontId)
        [void][int]::TryParse($xf.GetAttribute('fillId'), [ref]$fillId)
        $f = @{}
        if ($fontId -lt $fonts.Count) { foreach ($k in $fonts[$fontId].Keys) { $f[$k] = $fonts[$fontId][$k] } }
        if ($fillId -lt $fills.Count -and $fills[$fillId]) { $f['塗りつぶし'] = $fills[$fillId] }
        $result.Add((ConvertTo-FormatKey $f))
    }
    return , $result
}

function Read-RunFormats {
    #     XmlReader が <si> / <is> の開始位置にあるとき、文字列の一部だけに付いた書式を「「正式名称」赤 / …」の形で返す（無ければ ''）。
    #     書式が同じ隣どうしの部分はつなげる。終了時、リーダーは終了タグの上にある
    param([System.Xml.XmlReader] $Reader)

    if ($Reader.IsEmptyElement) { return '' }
    $depth = $Reader.Depth
    $runs = New-Object System.Collections.Generic.List[object]   # @(書式の表示, 文字列)
    $fmt = $null; $text = $null
    $inRun = $false; $inPr = $false; $inPhonetic = $false
    $advance = $true
    while ($true) {
        if ($advance -and -not $Reader.Read()) { break }
        $advance = $true
        if ($Reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            $name = $Reader.LocalName
            if ($name -eq 'r') {
                if ($Reader.IsEmptyElement) { continue }
                $inRun = $true; $fmt = @{}; $text = New-Object System.Text.StringBuilder
            }
            elseif ($name -eq 'rPr') { $inPr = -not $Reader.IsEmptyElement }
            elseif ($name -eq 'rPh') { $inPhonetic = -not $Reader.IsEmptyElement }
            elseif ($name -eq 't') {
                $t = $Reader.ReadElementContentAsString()
                $advance = $false
                if ($inRun -and -not $inPhonetic) { [void]$text.Append($t) }
            }
            elseif ($inPr) {
                Set-FontAttribute $fmt $name $Reader.GetAttribute('val') $Reader.GetAttribute('rgb') $Reader.GetAttribute('theme') $Reader.GetAttribute('indexed') $Reader.GetAttribute('tint') $Reader.GetAttribute('auto')
            }
        }
        elseif ($Reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
            if ($Reader.Depth -eq $depth) { break }
            $name = $Reader.LocalName
            if ($name -eq 'rPr') { $inPr = $false }
            elseif ($name -eq 'rPh') { $inPhonetic = $false }
            elseif ($name -eq 'r' -and $inRun) {
                $inRun = $false
                $desc = Format-FontText $fmt
                if ($runs.Count -gt 0 -and $runs[$runs.Count - 1][0] -ceq $desc) { $runs[$runs.Count - 1][1] += $text.ToString() }
                else { $runs.Add(@($desc, $text.ToString())) }
            }
        }
    }
    $parts = @()
    foreach ($r in $runs) { if ($r[0]) { $parts += ('「{0}」{1}' -f (Get-Short $r[1] 20), $r[0]) } }
    return ($parts -join ' / ')
}

function Read-SharedRunFormats {
    #     共有文字列テーブルの文字列ごとに、部分書式（Read-RunFormats）を返す
    param($Package, [string] $PartName)

    $result = New-Object System.Collections.Generic.List[string]
    if (-not $PartName) { return , $result }
    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return , $result }
    try {
        while ($reader.Read()) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element -and $reader.LocalName -eq 'si') { $result.Add((Read-RunFormats $reader)) }
        }
    }
    finally { $reader.Dispose() }
    return , $result
}

function Read-SheetFormats {
    #     ワークシートのセルの書式を読む。返り値: 行番号の文字列 → (列 → 書式のキー)。標準の書式のセルは含めない
    param($Package, [string] $PartName, $Styles, $SharedRuns)

    $result = @{}
    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return $result }
    $sep = $script:FormatSep
    $idx = 0
    $save = {
        $key = ''
        if ($style -lt $Styles.Count) { $key = $Styles[$style] }
        if ($runs) {
            if ($key) { $key += $sep }
            $key += '部分書式=' + $runs
        }
        if ($key) {
            $m = [regex]::Match($ref, '^([A-Z]+)(\d+)$')
            if ($m.Success) {
                $r = $m.Groups[2].Value
                if (-not $result.ContainsKey($r)) { $result[$r] = @{} }
                $result[$r][$m.Groups[1].Value] = $key
            }
        }
    }
    try {
        $row = 0; $col = 0
        $lastRef = $null
        $advance = $true
        $inCell = $false
        $ref = $null; $type = $null; $style = 0; $runs = ''
        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true

            $nodeType = $reader.NodeType
            if ($nodeType -eq [System.Xml.XmlNodeType]::Element) {
                $name = $reader.LocalName
                if ($name -eq 'c') {
                    $ref = $reader.GetAttribute('r')
                    if ($ref) { $lastRef = $ref }
                    else {
                        # r 属性を省略するツールで作られたファイル向けに位置を数える
                        if ($lastRef) { $col = Get-ColumnIndex $lastRef; $lastRef = $null }
                        $col++
                        $ref = (ConvertTo-ColumnName $col) + $row
                    }
                    $type = $reader.GetAttribute('t')
                    $s = $reader.GetAttribute('s')
                    $style = 0
                    if ($s) { [void][int]::TryParse($s, [ref]$style) }
                    $runs = ''
                    if ($reader.IsEmptyElement) { . $save; continue }
                    $inCell = $true
                }
                elseif ($name -eq 'row') {
                    $r = $reader.GetAttribute('r')
                    if ($r) { $row = [int]$r } else { $row++ }
                    $col = 0; $lastRef = $null
                }
                elseif ($inCell -and $name -eq 'v' -and $type -eq 's') {
                    $v = $reader.ReadElementContentAsString()
                    $advance = $false
                    if ([int]::TryParse($v, [ref]$idx) -and $idx -ge 0 -and $idx -lt $SharedRuns.Count) { $runs = $SharedRuns[$idx] }
                }
                elseif ($inCell -and $name -eq 'is') { $runs = Read-RunFormats $reader }
            }
            elseif ($nodeType -eq [System.Xml.XmlNodeType]::EndElement -and $inCell -and $reader.LocalName -eq 'c') {
                $inCell = $false
                . $save
            }
        }
    }
    finally { $reader.Dispose() }
    return $result
}

function Read-BookFormat {
    #     xlsx / xlsm の各シートのセルの書式を読む。返り値: シート名 → Read-SheetFormats の結果。
    #     .xls / .xlsb など OpenXML でないブックは $null
    param([string] $FilePath)

    if (@('.xlsx', '.xlsm') -notcontains [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()) { return $null }
    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) {
            if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break }
        }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        if ($null -eq $wbDoc) { return $null }
        $relById = @{}
        $sstPart = $null; $stylesPart = $null
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName $wbPart)) {
            $relById[$rel.Id] = $rel
            if ($rel.Kind -eq 'sharedStrings') { $sstPart = $rel.Target }
            elseif ($rel.Kind -eq 'styles') { $stylesPart = $rel.Target }
        }
        $styles = Read-StyleFormats -Package $pkg -PartName $stylesPart
        $sharedRuns = Read-SharedRunFormats -Package $pkg -PartName $sstPart

        $result = @{}
        foreach ($sheet in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $sheet)]
            if ($null -eq $rel -or ($rel.Kind -ne 'worksheet' -and $rel.Kind -ne 'macrosheet')) { continue }
            $result[$sheet.GetAttribute('name')] = Read-SheetFormats -Package $pkg -PartName $rel.Target -Styles $styles -SharedRuns $sharedRuns
        }
        return $result
    }
    finally { $pkg.Zip.Dispose() }
}

function Get-RowSignature {
    # 行の内容を 1 つの文字列にする（同じ内容の行は同じ文字列になる）
    param($Row)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($c in ($Row.V.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts.Add("$c`t" + $Row.V[$c]) }
    foreach ($c in ($Row.F.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts.Add("$c=`t" + $Row.F[$c]) }
    return ($parts -join "`n")
}

function Get-RowSimilarity {
    #     2 つの行がどれだけ似ているか（0〜1）。同じ列で同じ内容なら 1 点、同じ列に何か入っているだけなら 0.5 点
    param($A, $B)
    $cols = @{}
    foreach ($k in @($A.V.Keys) + @($A.F.Keys) + @($B.V.Keys) + @($B.F.Keys)) { $cols[$k] = $true }
    if ($cols.Count -eq 0) { return 1.0 }
    $score = 0.0
    foreach ($c in $cols.Keys) {
        $ia = $A.V.ContainsKey($c) -or $A.F.ContainsKey($c)
        $ib = $B.V.ContainsKey($c) -or $B.F.ContainsKey($c)
        if (-not ($ia -and $ib)) { continue }
        if ($A.V[$c] -ceq $B.V[$c] -and $A.F[$c] -ceq $B.F[$c]) { $score += 1.0 } else { $score += 0.5 }
    }
    return $score / $cols.Count
}

# ---------------------------------------------------------------- 行の対応付け


function Get-RowMapping {
    #     旧と新の行を対応付ける。返り値は Pairs（旧行, 新行）、Removed（旧だけの行）、Added（新だけの行）。
    #     1) 先頭と末尾の同じ行を対応付け  2) 残りを LCS（同じ内容の行の最長共通部分列）で対応付け
    #     3) 間に残った行は、似ている行どうしを順に「変更された行」として対応付ける
    param([string[]] $OldRows, [string[]] $NewRows, $OldBook, $NewBook)

    $pairs = New-Object System.Collections.Generic.List[object]
    $removed = New-Object System.Collections.Generic.List[string]
    $added = New-Object System.Collections.Generic.List[string]

    if ($NoRowAlign) {
        # 行番号どうしで比べる
        $all = @{}
        foreach ($r in $OldRows) { $all[$r] = $true }
        foreach ($r in $NewRows) { $all[$r] = $true }
        foreach ($r in $all.Keys) {
            $inOld = $OldBook.Rows.ContainsKey($r); $inNew = $NewBook.Rows.ContainsKey($r)
            if ($inOld -and $inNew) { $pairs.Add(@($r, $r)) } elseif ($inOld) { $removed.Add($r) } else { $added.Add($r) }
        }
        return [pscustomobject]@{ Pairs = $pairs; Removed = $removed; Added = $added }
    }

    $a = @($OldRows | ForEach-Object { Get-RowSignature $OldBook.Rows[$_] })
    $b = @($NewRows | ForEach-Object { Get-RowSignature $NewBook.Rows[$_] })
    $n = $a.Count; $m = $b.Count

    $head = 0
    while ($head -lt $n -and $head -lt $m -and $a[$head] -ceq $b[$head]) { $pairs.Add(@($OldRows[$head], $NewRows[$head])); $head++ }
    $tail = 0
    while ($tail -lt ($n - $head) -and $tail -lt ($m - $head) -and $a[$n - 1 - $tail] -ceq $b[$m - 1 - $tail]) { $tail++ }

    $oi = $head; $oEnd = $n - $tail
    $ni = $head; $nEnd = $m - $tail
    $matches = New-Object System.Collections.Generic.List[object]
    $on = $oEnd - $oi; $nn = $nEnd - $ni
    if ($on -gt 0 -and $nn -gt 0) {
        if ($on * $nn -le $maxAlignCells) {
            # LCS の表を作り、後ろからたどって一致する行を拾う
            $dp = New-Object 'int[,]' ($on + 1), ($nn + 1)
            for ($i = $on - 1; $i -ge 0; $i--) {
                for ($j = $nn - 1; $j -ge 0; $j--) {
                    if ($a[$oi + $i] -ceq $b[$ni + $j]) { $dp[$i, $j] = $dp[($i + 1), ($j + 1)] + 1 }
                    elseif ($dp[($i + 1), $j] -ge $dp[$i, ($j + 1)]) { $dp[$i, $j] = $dp[($i + 1), $j] }
                    else { $dp[$i, $j] = $dp[$i, ($j + 1)] }
                }
            }
            $i = 0; $j = 0
            while ($i -lt $on -and $j -lt $nn) {
                if ($a[$oi + $i] -ceq $b[$ni + $j]) { $matches.Add(@(($oi + $i), ($ni + $j))); $i++; $j++ }
                elseif ($dp[($i + 1), $j] -ge $dp[$i, ($j + 1)]) { $i++ }
                else { $j++ }
            }
        }
        else {
            # 行数が多いときは patience diff: 旧と新のどちらでも 1 回しか出てこない行を目印にし、
            # 順序が保たれる最長の組（最長増加部分列）を一致とする
            $countA = @{}; $countB = @{}; $posB = @{}
            for ($i = $oi; $i -lt $oEnd; $i++) { $countA[$a[$i]] = 1 + [int]$countA[$a[$i]] }
            for ($j = $ni; $j -lt $nEnd; $j++) { $countB[$b[$j]] = 1 + [int]$countB[$b[$j]]; $posB[$b[$j]] = $j }
            $cand = New-Object System.Collections.Generic.List[object]
            for ($i = $oi; $i -lt $oEnd; $i++) {
                if ($countA[$a[$i]] -eq 1 -and $countB[$a[$i]] -eq 1) { $cand.Add(@($i, $posB[$a[$i]])) }
            }
            # 最長増加部分列（新側の位置が増えていく最長の並び）
            $tails = New-Object System.Collections.Generic.List[int]
            $prev = New-Object int[] $cand.Count
            for ($k = 0; $k -lt $cand.Count; $k++) {
                $lo = 0; $hi = $tails.Count
                while ($lo -lt $hi) { $mid = [int][math]::Floor(($lo + $hi) / 2); if ($cand[$tails[$mid]][1] -lt $cand[$k][1]) { $lo = $mid + 1 } else { $hi = $mid } }
                if ($lo -gt 0) { $prev[$k] = $tails[$lo - 1] } else { $prev[$k] = -1 }
                if ($lo -eq $tails.Count) { $tails.Add($k) } else { $tails[$lo] = $k }
            }
            $chain = New-Object System.Collections.Generic.List[object]
            if ($tails.Count -gt 0) { $k = $tails[$tails.Count - 1]; while ($k -ge 0) { $chain.Insert(0, $cand[$k]); $k = $prev[$k] } }
            foreach ($c in $chain) { $matches.Add($c) }
            Write-Verbose ("行数が多いため patience diff で対応付けました（{0} 行 × {1} 行、目印 {2} 行）" -f $on, $nn, $chain.Count)
        }
    }
    # 番兵として末尾を一致扱いにし、一致と一致の間（ギャップ）を処理する
    $matches.Add(@($oEnd, $nEnd))

    $po = $oi; $pn = $ni
    foreach ($mt in $matches) {
        $gapOld = @(); if ($mt[0] -gt $po) { $gapOld = @($po..($mt[0] - 1)) }
        $gapNew = @(); if ($mt[1] -gt $pn) { $gapNew = @($pn..($mt[1] - 1)) }
        $usedNew = @{}
        $next = 0
        foreach ($x in $gapOld) {
            $best = -1; $bestScore = 0.5
            $limit = [math]::Min($gapNew.Count, $next + $pairWindow)
            for ($k = $next; $k -lt $limit; $k++) {
                $score = Get-RowSimilarity $OldBook.Rows[$OldRows[$x]] $NewBook.Rows[$NewRows[$gapNew[$k]]]
                if ($score -ge $bestScore -and ($best -lt 0 -or $score -gt $bestScore)) { $best = $k; $bestScore = $score }
            }
            if ($best -ge 0) {
                $pairs.Add(@($OldRows[$x], $NewRows[$gapNew[$best]]))
                for ($k = $next; $k -lt $best; $k++) { $added.Add($NewRows[$gapNew[$k]]); $usedNew[$k] = $true }
                $usedNew[$best] = $true
                $next = $best + 1
            }
            else { $removed.Add($OldRows[$x]) }
        }
        for ($k = $next; $k -lt $gapNew.Count; $k++) { if (-not $usedNew.ContainsKey($k)) { $added.Add($NewRows[$gapNew[$k]]) } }
        if ($mt[0] -lt $oEnd) { $pairs.Add(@($OldRows[$mt[0]], $NewRows[$mt[1]])) }
        $po = $mt[0] + 1; $pn = $mt[1] + 1
    }
    for ($t = 0; $t -lt $tail; $t++) { $pairs.Add(@($OldRows[$oEnd + $t], $NewRows[$nEnd + $t])) }

    return [pscustomobject]@{ Pairs = $pairs; Removed = $removed; Added = $added }
}

# ---------------------------------------------------------------- 差分の出力


$script:Diffs = New-Object System.Collections.Generic.List[object]

function Add-Diff {
    #     DetailText を指定すると、Detail をその文字にする（書式の差分など）
    param([string] $Sheet, [string] $NewCell, [string] $OldCell, [string] $Kind, [string] $Change, [string] $OldText, [string] $NewText, [string] $DetailText)

    $cell = $NewCell
    if (-not $cell) { $cell = $OldCell }
    $location = $Sheet
    if ($cell) { $location = '{0}!{1}' -f $Sheet, $cell }
    # シート名を変更した場合、旧の場所は旧のシート名で表す
    $oldSheet = $script:OldSheetName
    if (-not $oldSheet) { $oldSheet = $Sheet }
    $oldLocation = ''
    if ($OldCell -and ($OldCell -ne $NewCell -or $oldSheet -ne $Sheet)) { $oldLocation = '{0}!{1}' -f $oldSheet, $OldCell }

    switch ($Change) {
        '変更' { $detail = '{0}: {1} → {2}' -f $Kind, (Get-Short $OldText), (Get-Short $NewText) }
        '追加' { $detail = '{0}を追加: {1}' -f $Kind, (Get-Short $NewText) }
        '削除' { $detail = '{0}を削除: {1}' -f $Kind, (Get-Short $OldText) }
        '移動' { $detail = '{0}を移動: {1}' -f $Kind, (Get-Short $NewText) }
        '名前変更' { $detail = 'シート名を変更: {0} → {1}' -f $OldText, $NewText }
    }
    # 長い文字列の変更は、変わった部分だけを示す（[-削除-]{+追加+}）。ほとんど書き換わっている場合は全体を示す
    $segments = $null
    if ($Change -eq '変更' -and @('値', '数式', 'コメント', '図形') -contains $Kind -and [math]::Max($OldText.Length, $NewText.Length) -gt 20) {
        $td = Get-TextDiff $OldText $NewText
        if ($td.Similarity -ge 0.3) {
            $segments = $td.Segments
            $detail = '{0}: {1}' -f $Kind, (Format-InlineDiff $segments)
        }
    }
    if ($DetailText) { $detail = $DetailText }
    if ($oldLocation -and $NewCell) { $detail += "（旧 $OldCell）" }

    $d = [pscustomobject]@{
        File        = $script:CurrentFile
        Location    = $location
        Sheet       = $Sheet
        Cell        = $cell
        OldLocation = $oldLocation
        OldSheet    = $oldSheet
        Kind        = $Kind
        Change      = $Change
        Old         = $OldText
        New         = $NewText
        Detail      = $detail
        Path        = $script:CurrentNewPath
        OldPath     = $script:CurrentOldPath
        Segments    = $segments
    }
    $d.PSObject.TypeNames.Insert(0, 'CompareExcel.Diff')
    $script:Diffs.Add($d)
}

function Get-Short {
    param([string] $Text, [int] $Max = 60)
    if ($null -eq $Text) { return '' }
    if ($Text.Length -gt $Max) { return $Text.Substring(0, $Max) + '…' }
    return $Text
}

# ---------------------------------------------------------------- 文字単位の差分


# 英数字の並びは 1 語、それ以外（日本語など）は 1 文字ずつを単位にする
$script:TokenPattern = New-Object System.Text.RegularExpressions.Regex('[A-Za-z0-9_]+|\s+|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\s\S]')
$maxTextDiffCells = 250000   # 文字単位の差分を LCS で求める「旧の語数 × 新の語数」の上限（超えたら全体の置き換えとみなす）

function Get-TextDiff {
    #     2 つの文字列の違いを、一致（=）・削除（-）・追加（+）の区間の並びで返す。
    #     返り値: Segments（Op と Text を持つオブジェクトの配列）と Similarity（長い方の文字数に対する、一致した文字数の割合）
    param([string] $A, [string] $B)

    $ta = @($script:TokenPattern.Matches($A) | ForEach-Object { $_.Value })
    $tb = @($script:TokenPattern.Matches($B) | ForEach-Object { $_.Value })
    $n = $ta.Count; $m = $tb.Count

    # 先頭と末尾の一致を除いた残りを LCS で比べる
    $pre = 0
    while ($pre -lt $n -and $pre -lt $m -and $ta[$pre] -ceq $tb[$pre]) { $pre++ }
    $suf = 0
    while ($suf -lt ($n - $pre) -and $suf -lt ($m - $pre) -and $ta[$n - 1 - $suf] -ceq $tb[$m - 1 - $suf]) { $suf++ }

    $ops = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $pre; $i++) { $ops.Add(@('=', $ta[$i])) }
    $na = $n - $pre - $suf; $nb = $m - $pre - $suf
    if ($na -gt 0 -and $nb -gt 0 -and $na * $nb -le $maxTextDiffCells) {
        $w = $nb + 1
        $dp = New-Object 'int[]' (($na + 1) * $w)
        for ($i = $na - 1; $i -ge 0; $i--) {
            $x = $ta[$pre + $i]
            for ($j = $nb - 1; $j -ge 0; $j--) {
                if ($x -ceq $tb[$pre + $j]) { $dp[$i * $w + $j] = $dp[($i + 1) * $w + $j + 1] + 1 }
                else {
                    $down = $dp[($i + 1) * $w + $j]; $right = $dp[$i * $w + $j + 1]
                    if ($down -ge $right) { $dp[$i * $w + $j] = $down } else { $dp[$i * $w + $j] = $right }
                }
            }
        }
        $i = 0; $j = 0
        while ($i -lt $na -and $j -lt $nb) {
            if ($ta[$pre + $i] -ceq $tb[$pre + $j]) { $ops.Add(@('=', $ta[$pre + $i])); $i++; $j++ }
            elseif ($dp[($i + 1) * $w + $j] -ge $dp[$i * $w + $j + 1]) { $ops.Add(@('-', $ta[$pre + $i])); $i++ }
            else { $ops.Add(@('+', $tb[$pre + $j])); $j++ }
        }
        for (; $i -lt $na; $i++) { $ops.Add(@('-', $ta[$pre + $i])) }
        for (; $j -lt $nb; $j++) { $ops.Add(@('+', $tb[$pre + $j])) }
    }
    else {
        for ($i = 0; $i -lt $na; $i++) { $ops.Add(@('-', $ta[$pre + $i])) }
        for ($j = 0; $j -lt $nb; $j++) { $ops.Add(@('+', $tb[$pre + $j])) }
    }
    for ($i = $n - $suf; $i -lt $n; $i++) { $ops.Add(@('=', $ta[$i])) }

    # 同じ種類の続きをつなげる
    $segs = New-Object System.Collections.Generic.List[object]
    foreach ($o in $ops) {
        if ($segs.Count -gt 0 -and $segs[$segs.Count - 1].Op -eq $o[0]) { [void]$segs[$segs.Count - 1].Text.Append($o[1]) }
        else { $segs.Add([pscustomobject]@{ Op = $o[0]; Text = (New-Object System.Text.StringBuilder($o[1])) }) }
    }
    # 変更に挟まれた短い一致（2 文字以下）は変更に含める（細切れの差分は読みにくいため）
    for ($k = 1; $k -lt $segs.Count - 1; $k++) {
        if ($segs[$k].Op -eq '=' -and $segs[$k].Text.Length -le 2 -and $segs[$k - 1].Op -ne '=' -and $segs[$k + 1].Op -ne '=') { $segs[$k].Op = '~' }
    }
    # 一致の間の変更を「削除 → 追加」の順にまとめる
    $result = New-Object System.Collections.Generic.List[object]
    $addSame = {
        param([string] $Text)
        if (-not $Text) { return }
        if ($result.Count -gt 0 -and $result[$result.Count - 1].Op -eq '=') { $result[$result.Count - 1].Text += $Text }
        else { $result.Add([pscustomobject]@{ Op = '='; Text = $Text }) }
    }
    $del = New-Object System.Text.StringBuilder
    $ins = New-Object System.Text.StringBuilder
    $segs.Add([pscustomobject]@{ Op = '='; Text = $null })   # 最後の変更を出力させるための番兵
    foreach ($s in $segs) {
        $t = [string]$s.Text
        switch ($s.Op) {
            '-' { [void]$del.Append($t) }
            '+' { [void]$ins.Append($t) }
            '~' { [void]$del.Append($t); [void]$ins.Append($t) }
            '=' {
                $dt = $del.ToString(); $it = $ins.ToString()
                [void]$del.Clear(); [void]$ins.Clear()
                if ($dt -and $it) {
                    # 削除と追加の先頭・末尾に共通の語があれば一致に戻す（[- -]{+, X +} → {+, X+} ）
                    $da = @($script:TokenPattern.Matches($dt) | ForEach-Object { $_.Value })
                    $ia = @($script:TokenPattern.Matches($it) | ForEach-Object { $_.Value })
                    $p = 0
                    while ($p -lt $da.Count -and $p -lt $ia.Count -and $da[$p] -ceq $ia[$p]) { $p++ }
                    $q = 0
                    while ($q -lt $da.Count - $p -and $q -lt $ia.Count - $p -and $da[$da.Count - 1 - $q] -ceq $ia[$ia.Count - 1 - $q]) { $q++ }
                    if ($p -gt 0 -or $q -gt 0) {
                        & $addSame (Join-Tokens $da 0 $p)
                        $t = (Join-Tokens $da ($da.Count - $q) $q) + $t
                        $dt = Join-Tokens $da $p ($da.Count - $p - $q)
                        $it = Join-Tokens $ia $p ($ia.Count - $p - $q)
                    }
                }
                if ($dt) { $result.Add([pscustomobject]@{ Op = '-'; Text = $dt }) }
                if ($it) { $result.Add([pscustomobject]@{ Op = '+'; Text = $it }) }
                & $addSame $t
            }
        }
    }
    $same = 0
    foreach ($s in $result) { if ($s.Op -eq '=') { $same += $s.Text.Length } }
    $longer = [math]::Max($A.Length, $B.Length)
    $similarity = 1.0
    if ($longer -gt 0) { $similarity = $same / $longer }
    return [pscustomobject]@{ Segments = $result.ToArray(); Similarity = $similarity }
}

function Join-Tokens {
    #     語の配列の Start 番目から Count 個をつなげる（Count が 0 以下なら ''）
    param($Tokens, [int] $Start, [int] $Count)
    if ($Count -le 0) { return '' }
    return (-join $Tokens[$Start..($Start + $Count - 1)])
}

function Format-InlineDiff {
    #     差分の区間の並びを「…受注を[-登録-]{+更新+}する…」の形の 1 行にする。一致した部分は変更の前後の数文字だけ残す
    param($Segments, [int] $Context = 10, [int] $Max = 160)

    $sb = New-Object System.Text.StringBuilder
    $last = $Segments.Count - 1
    for ($k = 0; $k -le $last; $k++) {
        $t = $Segments[$k].Text
        switch ($Segments[$k].Op) {
            '=' {
                if ($k -eq 0) { if ($t.Length -gt $Context) { $t = '…' + $t.Substring($t.Length - $Context) } }
                elseif ($k -eq $last) { if ($t.Length -gt $Context) { $t = $t.Substring(0, $Context) + '…' } }
                elseif ($t.Length -gt 2 * $Context + 1) { $t = $t.Substring(0, $Context) + '…' + $t.Substring($t.Length - $Context) }
                [void]$sb.Append($t)
            }
            '-' { [void]$sb.Append('[-').Append((Get-Short $t 40)).Append('-]') }
            '+' { [void]$sb.Append('{+').Append((Get-Short $t 40)).Append('+}') }
        }
    }
    return (Get-Short $sb.ToString() $Max)
}

function Format-RowContent {
    param($Row)
    $parts = @()
    foreach ($c in ($Row.V.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts += ('{0}:{1}' -f $c, $Row.V[$c]) }
    foreach ($c in ($Row.F.Keys | Sort-Object { (Get-ColumnIndex $_) })) { if (-not $Row.V.ContainsKey($c)) { $parts += ('{0}:{1}' -f $c, $Row.F[$c]) } }
    return ($parts -join ' | ')
}

function Get-ColumnProfile {
    #     列ごとの「値の出現回数」を数える（列の対応付けに使う）。数式は値と区別するため先頭に = を付ける
    param($Sheet)
    $prof = @{}
    foreach ($r in $Sheet.Rows.Values) {
        foreach ($c in $r.V.Keys) {
            if (-not $prof.ContainsKey($c)) { $prof[$c] = @{} }
            $prof[$c]['v' + $r.V[$c]] = 1 + [int]$prof[$c]['v' + $r.V[$c]]
        }
        foreach ($c in $r.F.Keys) {
            if (-not $prof.ContainsKey($c)) { $prof[$c] = @{} }
            $prof[$c]['f' + $r.F[$c]] = 1 + [int]$prof[$c]['f' + $r.F[$c]]
        }
    }
    return $prof
}

function Get-BagSimilarity {
    #     2 つの「値 → 出現回数」の似ている度合い（重複を数える Jaccard 係数、0〜1）
    param($A, $B)
    $inter = 0; $union = 0
    foreach ($k in $A.Keys) {
        # 変数名は大文字小文字を区別しないため、引数の $A / $B と別の名前にする
        $countA = $A[$k]; $countB = [int]$B[$k]
        $inter += [math]::Min($countA, $countB); $union += [math]::Max($countA, $countB)
    }
    foreach ($k in $B.Keys) { if (-not $A.ContainsKey($k)) { $union += $B[$k] } }
    if ($union -eq 0) { return 1.0 }
    return $inter / $union
}

function Get-ColumnMapping {
    #     列の挿入・削除を見つける。列ごとの値の集まりが似ている列どうしを、順序を保って対応付ける。
    #     同じ列番地どうしの対応より明らかに良い（類似度の合計が 0.5 以上高い）場合だけ採用し、$null なら対応付けなし
    param($OldSheet, $NewSheet)

    $po = Get-ColumnProfile $OldSheet
    $pn = Get-ColumnProfile $NewSheet
    $oc = @($po.Keys | Sort-Object { Get-ColumnIndex $_ })
    $nc = @($pn.Keys | Sort-Object { Get-ColumnIndex $_ })
    if ($oc.Count -eq 0 -or $nc.Count -eq 0) { return $null }

    # 同じ列番地どうしの類似度。列の並びが同じで、どの列もよく似ていれば対応付けは不要
    $identity = 0.0
    $allSimilar = ($oc.Count -eq $nc.Count)
    foreach ($c in $oc) {
        if ($pn.ContainsKey($c)) {
            $sim = Get-BagSimilarity $po[$c] $pn[$c]
            $identity += $sim
            if ($sim -lt 0.8) { $allSimilar = $false }
        }
        else { $allSimilar = $false }
    }
    if ($allSimilar) { return $null }

    $n = $oc.Count; $m = $nc.Count
    $sim = New-Object 'double[,]' $n, $m
    for ($i = 0; $i -lt $n; $i++) { for ($j = 0; $j -lt $m; $j++) { $sim[$i, $j] = Get-BagSimilarity $po[$oc[$i]] $pn[$nc[$j]] } }
    # 類似度 0.5 以上の組だけを一致とみなし、合計が最大になる順序どおりの対応付けを求める
    $dp = New-Object 'double[,]' ($n + 1), ($m + 1)
    for ($i = $n - 1; $i -ge 0; $i--) {
        for ($j = $m - 1; $j -ge 0; $j--) {
            # Windows PowerShell 5.1 はメソッドの引数の中の 2 次元の添字（$a[$i, $j]）を解釈できないため、いったん変数に入れる
            $down = $dp[($i + 1), $j]; $right = $dp[$i, ($j + 1)]
            $best = [math]::Max($down, $right)
            if ($sim[$i, $j] -ge 0.5) {
                $match = $dp[($i + 1), ($j + 1)] + $sim[$i, $j]
                $best = [math]::Max($best, $match)
            }
            $dp[$i, $j] = $best
        }
    }
    if ($dp[0, 0] -le $identity + 0.5) { return $null }

    $newToOld = @{}
    $i = 0; $j = 0
    while ($i -lt $n -and $j -lt $m) {
        if ($sim[$i, $j] -ge 0.5 -and $dp[$i, $j] -eq $dp[($i + 1), ($j + 1)] + $sim[$i, $j]) { $newToOld[$nc[$j]] = $oc[$i]; $i++; $j++ }
        elseif ($dp[$i, $j] -eq $dp[($i + 1), $j]) { $i++ }
        else { $j++ }
    }
    $identityMap = $true
    foreach ($k in $newToOld.Keys) { if ($k -ne $newToOld[$k]) { $identityMap = $false } }
    if ($identityMap) { return $null }

    $mappedOld = @{}
    foreach ($v in $newToOld.Values) { $mappedOld[$v] = $true }
    return [pscustomobject]@{
        NewToOld    = $newToOld
        AddedCols   = @($nc | Where-Object { -not $newToOld.ContainsKey($_) })
        RemovedCols = @($oc | Where-Object { -not $mappedOld.ContainsKey($_) })
    }
}

function Convert-SheetColumns {
    #     シートの列番地を読み替えた写しを作る。Map に無い列は取り除く
    param($Sheet, $Map)
    $rows = @{}
    foreach ($r in $Sheet.Rows.Keys) {
        $v = @{}; $f = @{}
        foreach ($c in $Sheet.Rows[$r].V.Keys) { if ($Map.ContainsKey($c)) { $v[$Map[$c]] = $Sheet.Rows[$r].V[$c] } }
        foreach ($c in $Sheet.Rows[$r].F.Keys) { if ($Map.ContainsKey($c)) { $f[$Map[$c]] = $Sheet.Rows[$r].F[$c] } }
        if ($v.Count -gt 0 -or $f.Count -gt 0) { $rows[$r] = @{ V = $v; F = $f } }
    }
    return [pscustomobject]@{ Rows = $rows; Comments = $Sheet.Comments; Shapes = $Sheet.Shapes }
}

function Get-KeyRegions {
    #     キー列の見出し（-KeyColumn の文字と一致するセル）を探し、その下の「キー列に値がある行」を表の範囲とする。
    #     返り値: 見出しの文字と出現順ごとの範囲（行の一覧と、行 → キーの値）
    param($Sheet, [string[]] $Keys)

    $regions = New-Object System.Collections.Generic.List[object]
    $rows = @($Sheet.Rows.Keys | Sort-Object { [int]$_ })
    $seenCount = @{}
    for ($i = 0; $i -lt $rows.Count; $i++) {
        $r = $Sheet.Rows[$rows[$i]]
        foreach ($c in ($r.V.Keys | Sort-Object { Get-ColumnIndex $_ })) {
            $text = ([string]$r.V[$c]).Trim()
            if ($Keys -notcontains $text) { continue }
            $seenCount[$text] = 1 + [int]$seenCount[$text]
            $region = [pscustomobject]@{ Name = '{0}#{1}' -f $text, $seenCount[$text]; Rows = New-Object System.Collections.Generic.List[string]; KeyOf = @{}; Duplicate = $false }
            $values = @{}
            # 見出しの次の行から、キー列が空の行（または空行）が出るまで
            for ($k = $i + 1; $k -lt $rows.Count; $k++) {
                if ([int]$rows[$k] -ne [int]$rows[$k - 1] + 1) { break }
                $key = ([string]$Sheet.Rows[$rows[$k]].V[$c]).Trim()
                if (-not $key) { break }
                if ($values.ContainsKey($key)) { $region.Duplicate = $true }
                $values[$key] = $true
                $region.Rows.Add($rows[$k])
                $region.KeyOf[$rows[$k]] = $key
            }
            if ($region.Rows.Count -gt 0) { $regions.Add($region) }
        }
    }
    return $regions
}

function Get-LisIndexes {
    #     数列の最長増加部分列に含まれる要素の位置（0 始まり）を返す
    param([int[]] $Values)
    $tails = New-Object System.Collections.Generic.List[int]
    $prev = New-Object int[] $Values.Count
    for ($k = 0; $k -lt $Values.Count; $k++) {
        $lo = 0; $hi = $tails.Count
        while ($lo -lt $hi) { $mid = [int][math]::Floor(($lo + $hi) / 2); if ($Values[$tails[$mid]] -lt $Values[$k]) { $lo = $mid + 1 } else { $hi = $mid } }
        if ($lo -gt 0) { $prev[$k] = $tails[$lo - 1] } else { $prev[$k] = -1 }
        if ($lo -eq $tails.Count) { $tails.Add($k) } else { $tails[$lo] = $k }
    }
    $result = @{}
    if ($tails.Count -gt 0) { $k = $tails[$tails.Count - 1]; while ($k -ge 0) { $result[$k] = $true; $k = $prev[$k] } }
    return $result
}

function Get-SheetSimilarity {
    #     2 つのシートの中身がどれだけ似ているか（行の内容の集まりの Jaccard 係数）。シート名の変更の判定に使う
    param($A, $B)
    $ba = @{}; $bb = @{}
    foreach ($r in $A.Rows.Values) { $s = Get-RowSignature $r; $ba[$s] = 1 + [int]$ba[$s] }
    foreach ($r in $B.Rows.Values) { $s = Get-RowSignature $r; $bb[$s] = 1 + [int]$bb[$s] }
    if ($ba.Count -eq 0 -and $bb.Count -eq 0) { return 0.0 }
    return Get-BagSimilarity $ba $bb
}

function Compare-Sheet {
    param([string] $Sheet, $OldSheet, $NewSheet, [string] $OldName)

    $script:OldSheetName = $OldName
    if (-not $script:OldSheetName) { $script:OldSheetName = $Sheet }
    $firstDiff = $script:Diffs.Count

    # ---- 列の挿入・削除: 新の列を旧の列番地に読み替えて比べ、出力は新の列番地に戻す
    $toNew = @{}
    $oldView = $OldSheet; $newView = $NewSheet
    $cm = $null
    if (-not $NoColumnAlign -and -not $NoRowAlign) { $cm = Get-ColumnMapping $OldSheet $NewSheet }
    if ($cm) {
        foreach ($k in $cm.NewToOld.Keys) { $toNew[$cm.NewToOld[$k]] = $k }
        $keepOld = @{}
        foreach ($v in $cm.NewToOld.Values) { $keepOld[$v] = $v }
        $newView = Convert-SheetColumns $NewSheet $cm.NewToOld
        $oldView = Convert-SheetColumns $OldSheet $keepOld
        foreach ($c in $cm.AddedCols) {
            $cells = @($NewSheet.Rows.Keys | Where-Object { $NewSheet.Rows[$_].V.ContainsKey($c) -or $NewSheet.Rows[$_].F.ContainsKey($c) } | Sort-Object { [int]$_ })
            $sample = @($cells | Select-Object -First 3 | ForEach-Object { $NewSheet.Rows[$_].V[$c] }) -join ', '
            Add-Diff $Sheet "$c$($cells[0])" $null '列' '追加' $null ("{0} 列（{1} 個: {2}）" -f $c, $cells.Count, $sample)
        }
        foreach ($c in $cm.RemovedCols) {
            $cells = @($OldSheet.Rows.Keys | Where-Object { $OldSheet.Rows[$_].V.ContainsKey($c) -or $OldSheet.Rows[$_].F.ContainsKey($c) } | Sort-Object { [int]$_ })
            $sample = @($cells | Select-Object -First 3 | ForEach-Object { $OldSheet.Rows[$_].V[$c] }) -join ', '
            Add-Diff $Sheet $null "$c$($cells[0])" '列' '削除' ("{0} 列（{1} 個: {2}）" -f $c, $cells.Count, $sample) $null
        }
    }
    $L = { param($c) if ($toNew.ContainsKey($c)) { $toNew[$c] } else { $c } }

    $oldRows = @($oldView.Rows.Keys | Sort-Object { [int]$_ })
    $newRows = @($newView.Rows.Keys | Sort-Object { [int]$_ })

    # ---- キー列のある表は、キーの値で行を対応付ける
    $pairs = New-Object System.Collections.Generic.List[object]
    $added = New-Object System.Collections.Generic.List[string]
    $removed = New-Object System.Collections.Generic.List[string]
    $moved = New-Object System.Collections.Generic.List[object]
    $keyedOld = @{}; $keyedNew = @{}
    if ($KeyColumn -and -not $NoRowAlign) {
        $ro = Get-KeyRegions $oldView $KeyColumn
        $rn = Get-KeyRegions $newView $KeyColumn
        foreach ($regionNew in $rn) {
            $regionOld = @($ro | Where-Object { $_.Name -eq $regionNew.Name })[0]
            if (-not $regionOld) { continue }
            if ($regionOld.Duplicate -or $regionNew.Duplicate) {
                Write-Warning ("{0}: キー列「{1}」の値が重複しているため、この表はキーで対応付けません" -f $Sheet, $regionNew.Name.Split('#')[0])
                continue
            }
            $byKeyOld = @{}
            foreach ($r in $regionOld.Rows) { $byKeyOld[$regionOld.KeyOf[$r]] = $r; $keyedOld[$r] = $true }
            $seq = New-Object System.Collections.Generic.List[object]
            foreach ($r in $regionNew.Rows) {
                $keyedNew[$r] = $true
                $key = $regionNew.KeyOf[$r]
                if ($byKeyOld.ContainsKey($key)) { $seq.Add(@($byKeyOld[$key], $r)); $byKeyOld.Remove($key) }
                else { $added.Add($r) }
            }
            foreach ($k in $byKeyOld.Keys) { $removed.Add($byKeyOld[$k]) }
            # 順番が入れ替わった行（旧の順に並べたとき、新の位置の最長増加部分列に入らない行）
            $ordered = @($seq | Sort-Object { [int]$_[0] })
            $lis = Get-LisIndexes ([int[]]@($ordered | ForEach-Object { [int]$_[1] }))
            for ($k = 0; $k -lt $ordered.Count; $k++) {
                $pairs.Add($ordered[$k])
                if (-not $lis.ContainsKey($k)) { $moved.Add($ordered[$k]) }
            }
        }
    }

    # ---- 残りの行は内容で対応付ける
    $restOld = @($oldRows | Where-Object { -not $keyedOld.ContainsKey($_) })
    $restNew = @($newRows | Where-Object { -not $keyedNew.ContainsKey($_) })
    $map = Get-RowMapping -OldRows $restOld -NewRows $restNew -OldBook $oldView -NewBook $newView
    foreach ($p in $map.Pairs) { $pairs.Add($p) }

    # 削除と追加に同じ内容の行があれば、行の移動とみなす
    $addedBySig = @{}
    foreach ($r in $map.Added) {
        $sig = Get-RowSignature $newView.Rows[$r]
        if (-not $addedBySig.ContainsKey($sig)) { $addedBySig[$sig] = New-Object System.Collections.Generic.List[string] }
        $addedBySig[$sig].Add($r)
    }
    $movedNew = @{}
    foreach ($r in $map.Removed) {
        $sig = Get-RowSignature $oldView.Rows[$r]
        if ($addedBySig.ContainsKey($sig) -and $addedBySig[$sig].Count -gt 0) {
            $to = $addedBySig[$sig][0]; $addedBySig[$sig].RemoveAt(0)
            $moved.Add(@($r, $to)); $movedNew[$to] = $true
            $pairs.Add(@($r, $to))
        }
        else { $removed.Add($r) }
    }
    foreach ($r in $map.Added) { if (-not $movedNew.ContainsKey($r)) { $added.Add($r) } }

    # ---- 対応付けた行のセルを比べる
    $rowMap = @{}
    foreach ($p in $pairs) {
        $rowMap[$p[0]] = $p[1]
        $o = $oldView.Rows[$p[0]]; $n = $newView.Rows[$p[1]]
        foreach ($kind in @(@('V', '値'), @('F', '数式'))) {
            $cols = @{}
            foreach ($c in @($o[$kind[0]].Keys) + @($n[$kind[0]].Keys)) { $cols[$c] = $true }
            foreach ($c in ($cols.Keys | Sort-Object { (Get-ColumnIndex $_) })) {
                $ov = $o[$kind[0]][$c]; $nv = $n[$kind[0]][$c]
                if ($ov -ceq $nv) { continue }
                $newCell = (& $L $c) + $p[1]; $oldCell = "$c$($p[0])"
                if ($null -eq $ov) { Add-Diff $Sheet $newCell $oldCell $kind[1] '追加' $null $nv }
                elseif ($null -eq $nv) { Add-Diff $Sheet $newCell $oldCell $kind[1] '削除' $ov $null }
                else { Add-Diff $Sheet $newCell $oldCell $kind[1] '変更' $ov $nv }
            }
        }
    }
    foreach ($mv in ($moved | Sort-Object { [int]$_[1] })) {
        Add-Diff $Sheet "A$($mv[1])" "A$($mv[0])" '行' '移動' (Format-RowContent $oldView.Rows[$mv[0]]) (Format-RowContent $newView.Rows[$mv[1]])
    }
    foreach ($r in ($added | Sort-Object { [int]$_ })) {
        Add-Diff $Sheet "A$r" $null '行' '追加' $null (Format-RowContent $NewSheet.Rows[$r])
    }
    foreach ($r in ($removed | Sort-Object { [int]$_ })) {
        Add-Diff $Sheet $null "A$r" '行' '削除' (Format-RowContent $OldSheet.Rows[$r]) $null
    }

    if ($CompareFormat) { Compare-SheetFormat -Sheet $Sheet -Pairs $pairs -OldSheet $OldSheet -NewSheet $NewSheet -OldView $oldView -NewView $newView -ColumnMap $cm -ToNew $toNew }

    # ---- コメント: 旧の位置を行・列の対応で新の位置に読み替えて比べる
    $oldComments = @{}
    foreach ($cell in $OldSheet.Comments.Keys) {
        $m = [regex]::Match($cell, '^([A-Z]+)(\d+)$')
        $row = $m.Groups[2].Value
        if ($rowMap.ContainsKey($row)) { $row = $rowMap[$row] }
        $oldComments[(& $L $m.Groups[1].Value) + $row] = @($cell, $OldSheet.Comments[$cell])
    }
    foreach ($cell in $NewSheet.Comments.Keys) {
        if ($oldComments.ContainsKey($cell)) {
            if ($oldComments[$cell][1] -cne $NewSheet.Comments[$cell]) { Add-Diff $Sheet $cell $oldComments[$cell][0] 'コメント' '変更' $oldComments[$cell][1] $NewSheet.Comments[$cell] }
            $oldComments.Remove($cell)
        }
        else { Add-Diff $Sheet $cell $null 'コメント' '追加' $null $NewSheet.Comments[$cell] }
    }
    foreach ($k in $oldComments.Keys) { Add-Diff $Sheet $null $oldComments[$k][0] 'コメント' '削除' $oldComments[$k][1] $null }

    # ---- 図形: 名前で対応付ける
    foreach ($name in $NewSheet.Shapes.Keys) {
        $ns = $NewSheet.Shapes[$name]
        if ($OldSheet.Shapes.ContainsKey($name)) {
            $os = $OldSheet.Shapes[$name]
            if ($os.Text -cne $ns.Text) { Add-Diff $Sheet $ns.Cell $os.Cell '図形' '変更' $os.Text $ns.Text }
        }
        else { Add-Diff $Sheet $ns.Cell $null '図形' '追加' $null $ns.Text }
    }
    foreach ($name in $OldSheet.Shapes.Keys) {
        if (-not $NewSheet.Shapes.ContainsKey($name)) { Add-Diff $Sheet $null $OldSheet.Shapes[$name].Cell '図形' '削除' $OldSheet.Shapes[$name].Text $null }
    }
    Merge-Renumbering -From $firstDiff
    Sort-SheetDiff -From $firstDiff
    $script:OldSheetName = $null
}

function Compare-SheetFormat {
    #     対応付けた行・列のセルの書式を比べる（-CompareFormat）。
    #     同じ変更が隣り合うセルは、範囲（A5:F5 など）にまとめて 1 件にする
    param([string] $Sheet, $Pairs, $OldSheet, $NewSheet, $OldView, $NewView, $ColumnMap, [hashtable] $ToNew)

    $found = New-Object System.Collections.Generic.List[object]
    foreach ($p in $Pairs) {
        $oldRow = [string]$p[0]; $newRow = [string]$p[1]
        $of = $OldSheet.Formats[$oldRow]; $nf = $NewSheet.Formats[$newRow]
        if (-not $of -and -not $nf) { continue }
        # 比べる列（旧の列番地 → 新の列番地）。列を対応付けた場合は、対応の付いた列だけ
        $cols = @{}
        if ($of) {
            foreach ($c in $of.Keys) {
                if (-not $ColumnMap) { $cols[$c] = $c }
                elseif ($ToNew.ContainsKey($c)) { $cols[$c] = $ToNew[$c] }
            }
        }
        if ($nf) {
            foreach ($c in $nf.Keys) {
                if (-not $ColumnMap) { $cols[$c] = $c }
                elseif ($ColumnMap.NewToOld.ContainsKey($c)) { $cols[$ColumnMap.NewToOld[$c]] = $c }
            }
        }
        foreach ($oc in $cols.Keys) {
            $nc = $cols[$oc]
            $ok = ''; $nk = ''
            if ($of -and $of.ContainsKey($oc)) { $ok = $of[$oc] }
            if ($nf -and $nf.ContainsKey($nc)) { $nk = $nf[$nc] }
            if ($ok -ceq $nk) { continue }
            $o = ConvertFrom-FormatKey $ok
            $n = ConvertFrom-FormatKey $nk
            # 文字列そのものが変わったセルの部分書式は比べない（変更は値の差分でわかる）
            if ($o.ContainsKey('部分書式') -or $n.ContainsKey('部分書式')) {
                $ov = $null; $nv = $null
                if ($OldView.Rows.ContainsKey($oldRow)) { $ov = $OldView.Rows[$oldRow].V[$oc] }
                if ($NewView.Rows.ContainsKey($newRow)) { $nv = $NewView.Rows[$newRow].V[$oc] }
                if ($ov -cne $nv) { $o.Remove('部分書式'); $n.Remove('部分書式') }
            }
            $changes = @()
            foreach ($k in $script:FormatOrder) {
                if ($o[$k] -cne $n[$k]) { $changes += ('{0} {1} → {2}' -f $k, (Get-FormatValueText $k $o[$k]), (Get-FormatValueText $k $n[$k])) }
            }
            if ($changes.Count -eq 0) { continue }
            $found.Add([pscustomobject]@{
                NewRow = [int]$newRow; NewCol = (Get-ColumnIndex $nc); OldRow = [int]$oldRow; OldCol = (Get-ColumnIndex $oc)
                Old = (Format-FormatText $o); New = (Format-FormatText $n); Detail = '書式: ' + ($changes -join '、')
            })
        }
    }
    if ($found.Count -eq 0) { return }

    $groups = [ordered]@{}
    foreach ($f in $found) {
        $k = $f.Old + $script:FormatSep + $f.New + $script:FormatSep + $f.Detail
        if (-not $groups.Contains($k)) { $groups[$k] = New-Object System.Collections.Generic.List[object] }
        $groups[$k].Add($f)
    }
    foreach ($g in $groups.Values) {
        # 横に並んだセル（旧でも並んでいるもの）をまとめる
        $spans = New-Object System.Collections.Generic.List[object]
        foreach ($f in ($g | Sort-Object NewRow, NewCol)) {
            $last = $null
            if ($spans.Count -gt 0) { $last = $spans[$spans.Count - 1] }
            if ($last -and $last.R1 -eq $f.NewRow -and $last.C1 + 1 -eq $f.NewCol -and $last.OR1 -eq $f.OldRow -and $last.OC1 + 1 -eq $f.OldCol) {
                $last.C1 = $f.NewCol; $last.OC1 = $f.OldCol
            }
            else {
                $spans.Add([pscustomobject]@{ R0 = $f.NewRow; R1 = $f.NewRow; C0 = $f.NewCol; C1 = $f.NewCol; OR0 = $f.OldRow; OR1 = $f.OldRow; OC0 = $f.OldCol; OC1 = $f.OldCol; F = $f })
            }
        }
        # 縦に並んだ、同じ列の範囲をまとめる
        $rects = New-Object System.Collections.Generic.List[object]
        foreach ($s in ($spans | Sort-Object C0, C1, R0)) {
            $last = $null
            if ($rects.Count -gt 0) { $last = $rects[$rects.Count - 1] }
            if ($last -and $last.C0 -eq $s.C0 -and $last.C1 -eq $s.C1 -and $last.OC0 -eq $s.OC0 -and $last.OC1 -eq $s.OC1 -and $last.R1 + 1 -eq $s.R0 -and $last.OR1 + 1 -eq $s.OR0) {
                $last.R1 = $s.R0; $last.OR1 = $s.OR0
            }
            else { $rects.Add($s) }
        }
        foreach ($r in $rects) {
            $newRef = (ConvertTo-ColumnName $r.C0) + $r.R0
            if ($r.C1 -ne $r.C0 -or $r.R1 -ne $r.R0) { $newRef += ':' + (ConvertTo-ColumnName $r.C1) + $r.R1 }
            $oldRef = (ConvertTo-ColumnName $r.OC0) + $r.OR0
            if ($r.OC1 -ne $r.OC0 -or $r.OR1 -ne $r.OR0) { $oldRef += ':' + (ConvertTo-ColumnName $r.OC1) + $r.OR1 }
            Add-Diff $Sheet $newRef $oldRef '書式' '変更' $r.F.Old $r.F.New -DetailText $r.F.Detail
        }
    }
}

function Merge-Renumbering {
    #     行の挿入・削除で No などの連番が振り直されると、番号のセルが 1 つずつ「変更」になってしまう。
    #     同じ列で、新の行が連続し、旧の番号も連続していて、どれも同じだけずれた整数のセルの変更（2 個以上）を 1 件にまとめる
    param([int] $From)

    $items = $script:Diffs.GetRange($From, $script:Diffs.Count - $From).ToArray()
    $isNum = { param($d) $d.Kind -eq '値' -and $d.Change -eq '変更' -and $d.Old -match '^-?\d+$' -and $d.New -match '^-?\d+$' }
    $runs = New-Object System.Collections.Generic.List[object]
    $byCol = @{}
    for ($i = 0; $i -lt $items.Count; $i++) {
        $d = $items[$i]
        if (-not (& $isNum $d)) { continue }
        $m = [regex]::Match($d.Cell, '^([A-Z]+)(\d+)$')
        $col = $m.Groups[1].Value
        if (-not $byCol.ContainsKey($col)) { $byCol[$col] = New-Object System.Collections.Generic.List[object] }
        $byCol[$col].Add([pscustomobject]@{ Index = $i; Diff = $d; Row = [int]$m.Groups[2].Value; OldNumber = [long]$d.Old; Delta = [long]$d.New - [long]$d.Old })
    }
    foreach ($col in $byCol.Keys) {
        $cur = New-Object System.Collections.Generic.List[object]
        foreach ($x in ($byCol[$col] | Sort-Object Row)) {
            # 行も番号も 1 ずつ増え、ずれ幅が同じあいだを 1 つの連番とみなす（金額などがたまたま同じだけ増えたものはまとめない）
            $prev = if ($cur.Count -gt 0) { $cur[$cur.Count - 1] } else { $null }
            if ($prev -and ($x.Row -ne $prev.Row + 1 -or $x.Delta -ne $prev.Delta -or $x.OldNumber -ne $prev.OldNumber + 1)) {
                if ($cur.Count -ge 2) { $runs.Add($cur) }
                $cur = New-Object System.Collections.Generic.List[object]
            }
            $cur.Add($x)
        }
        if ($cur.Count -ge 2) { $runs.Add($cur) }
    }
    if ($runs.Count -eq 0) { return }

    $drop = @{}        # まとめて消す差分の位置
    $summaryAt = @{}   # まとめた 1 件を置く位置（まとめた中で最初の差分の位置）
    foreach ($run in $runs) {
        $first = $run[0].Diff; $last = $run[$run.Count - 1].Diff
        foreach ($x in $run) { $drop[$x.Index] = $true }
        # 場所は範囲（A20:A26）で表す。旧の場所は、行がずれている場合だけ
        $range = '{0}:{1}' -f $first.Cell, $last.Cell
        $oldCell = { param($d) if ($d.OldLocation) { $d.OldLocation.Substring($d.OldLocation.LastIndexOf('!') + 1) } else { $d.Cell } }
        $oldRange = '{0}:{1}' -f (& $oldCell $first), (& $oldCell $last)
        $oldLocation = ''
        if ($oldRange -ne $range -or $first.OldSheet -ne $first.Sheet) { $oldLocation = '{0}!{1}' -f $first.OldSheet, $oldRange }
        $summary = [pscustomobject]@{
            File = $first.File; Location = ('{0}!{1}' -f $first.Sheet, $range); Sheet = $first.Sheet; Cell = $range
            OldLocation = $oldLocation; OldSheet = $first.OldSheet; Kind = '連番'; Change = '変更'
            Old = '{0}～{1}' -f $first.Old, $last.Old; New = '{0}～{1}' -f $first.New, $last.New
            Detail = '連番の振り直し: {0}～{1} → {2}～{3}（{4} 個）' -f $first.Old, $last.Old, $first.New, $last.New, $run.Count
            Path = $first.Path; OldPath = $first.OldPath; Segments = $null
        }
        $summary.PSObject.TypeNames.Insert(0, 'CompareExcel.Diff')
        $summaryAt[(@($run | Sort-Object Index)[0].Index)] = $summary
    }
    $kept = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $items.Count; $i++) {
        if ($summaryAt.ContainsKey($i)) { $kept.Add($summaryAt[$i]) }
        elseif (-not $drop.ContainsKey($i)) { $kept.Add($items[$i]) }
    }
    $script:Diffs.RemoveRange($From, $script:Diffs.Count - $From)
    foreach ($d in $kept) { $script:Diffs.Add($d) }
}

function Sort-SheetDiff {
    #     シートの中の差分を、上の行から順（同じ行なら左の列から）に並べ替える。
    #     削除した行・セルは旧の行番号の位置に置く。同じ位置どうしは見つけた順のまま
    param([int] $From)

    $count = $script:Diffs.Count - $From
    if ($count -lt 2) { return }
    $keyed = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $count; $i++) {
        $d = $script:Diffs[$From + $i]
        $row = 0; $col = 0
        $m = [regex]::Match([string]$d.Cell, '^([A-Z]+)(\d+)')
        if ($m.Success) { $row = [int]$m.Groups[2].Value; $col = Get-ColumnIndex $m.Groups[1].Value }
        $keyed.Add([pscustomobject]@{ Row = $row; Col = $col; Order = $i; Diff = $d })
    }
    $sorted = @($keyed | Sort-Object Row, Col, Order)
    for ($i = 0; $i -lt $count; $i++) { $script:Diffs[$From + $i] = $sorted[$i].Diff }
}

function Compare-Book {
    param([string] $OldPath, [string] $NewPath)

    $oldBook = Read-Book $OldPath
    $newBook = Read-Book $NewPath
    $firstDiff = $script:Diffs.Count

    # 片方にしか無いシートどうしで中身がよく似ていれば、シート名の変更とみなす
    $renamed = @{}   # 新のシート名 → 旧のシート名
    $onlyOld = @($oldBook.Keys | Where-Object { -not $newBook.Contains($_) })
    $onlyNew = @($newBook.Keys | Where-Object { -not $oldBook.Contains($_) })
    if ($onlyOld.Count -gt 0 -and $onlyNew.Count -gt 0) {
        $cands = New-Object System.Collections.Generic.List[object]
        foreach ($o in $onlyOld) { foreach ($n in $onlyNew) {
            $sim = Get-SheetSimilarity $oldBook[$o] $newBook[$n]
            if ($sim -ge 0.5) { $cands.Add([pscustomobject]@{ Old = $o; New = $n; Sim = $sim }) }
        } }
        $usedOld = @{}
        foreach ($c in ($cands | Sort-Object Sim -Descending)) {
            if ($usedOld.ContainsKey($c.Old) -or $renamed.ContainsKey($c.New)) { continue }
            $usedOld[$c.Old] = $true; $renamed[$c.New] = $c.Old
        }
    }

    foreach ($sheet in $newBook.Keys) {
        if ($oldBook.Contains($sheet)) { Compare-Sheet -Sheet $sheet -OldSheet $oldBook[$sheet] -NewSheet $newBook[$sheet] }
        elseif ($renamed.ContainsKey($sheet)) {
            $oldName = $renamed[$sheet]
            $script:OldSheetName = $oldName
            Add-Diff $sheet $null $null 'シート' '名前変更' $oldName $sheet
            $script:OldSheetName = $null
            Compare-Sheet -Sheet $sheet -OldSheet $oldBook[$oldName] -NewSheet $newBook[$sheet] -OldName $oldName
        }
        else { Add-Diff $sheet $null $null 'シート' '追加' $null ("{0} 行" -f $newBook[$sheet].Rows.Count) }
    }
    $renamedOld = @{}
    foreach ($v in $renamed.Values) { $renamedOld[$v] = $true }
    foreach ($sheet in $oldBook.Keys) {
        if (-not $newBook.Contains($sheet) -and -not $renamedOld.ContainsKey($sheet)) { Add-Diff $sheet $null $null 'シート' '削除' ("{0} 行" -f $oldBook[$sheet].Rows.Count) $null }
    }
    $diffs = $script:Diffs.GetRange($firstDiff, $script:Diffs.Count - $firstDiff).ToArray()
    $groups = $null
    if ($OutHistory -or $CheckHistory) { $groups = Get-ChangeGroups -OldBook $oldBook -NewBook $newBook -Diffs $diffs }
    if ($OutHistory) {
        foreach ($g in $groups.Values) {
            $text = $g.Texts -join '、'
            if ($g.Label) { $text = '{0}: {1}' -f $g.Label, $text }
            $location = $g.Sheet
            $cell = $g.NewCell
            if (-not $cell) { $cell = $g.OldCell }
            if ($cell) { $location = '{0}!{1}' -f $g.Sheet, $cell }
            $script:HistoryDrafts.Add([pscustomobject]@{ File = $script:CurrentFile; Section = $g.Sheet; Description = $text; Location = $location })
        }
    }
    if ($CheckHistory) { Test-HistoryCoverage -NewBook $newBook -Diffs $diffs -Groups $groups }
}

# ---------------------------------------------------------------- 改訂履歴の記載漏れ（-CheckHistory）


function Test-HistoryTerm {
    #     改訂履歴と突き合わせる語として使えるか（1 文字のもの、数字・日付・記号だけのものは使わない）
    param([string] $Text)
    if (-not $Text) { return $false }
    $t = $Text.Trim()
    if ($t.Length -lt 2) { return $false }
    if ($t -match '^[\d\s.,:/+%年月日()（）-]+$') { return $false }
    if ($t -match '^[\p{P}\p{S}\s○◯●×✓]+$') { return $false }
    return $true
}

function Get-RowLabelTerms {
    #     行を表す語として、左から 2 つの語を返す（例: 合計金額 / TOTAL_AMOUNT）
    param($Row)
    $terms = New-Object System.Collections.Generic.List[string]
    if (-not $Row) { return , $terms }
    foreach ($c in ($Row.V.Keys | Sort-Object { Get-ColumnIndex $_ })) {
        if (Test-HistoryTerm $Row.V[$c]) {
            $terms.Add($Row.V[$c].Trim())
            if ($terms.Count -ge 2) { break }
        }
    }
    return , $terms
}

function Get-ColumnHeaderText {
    #     セルの列の見出しを返す。上の行へたどり、数値のセルが無く文字のセルが 2 つ以上ある最初の行を見出しの行とみなす
    param($Sheet, [string] $Col, [int] $Row)
    for ($r = $Row - 1; $r -ge 1; $r--) {
        $data = $Sheet.Rows["$r"]
        if (-not $data -or $data.V.Count -lt 2) { continue }
        if (@($data.V.Values | Where-Object { $_ -match '^\s*-?[\d.,]+\s*$|^\s*\d{4}[/-]\d{1,2}[/-]\d{1,2}' }).Count -gt 0) { continue }
        if ($data.V.ContainsKey($Col) -and $data.V[$Col].Trim()) { return $data.V[$Col].Trim() }
        return $null
    }
    return $null
}

# 版を上げるたびに変わる欄（表紙やシートの上部の版数・更新日・更新者など）の見出し。改訂内容としては扱わない
$script:BookkeepingLabel = '^(版数|版|改訂版数|バージョン|Version|Ver\.?|Rev\.?|改版番号|作成日|更新日|最終更新日|改訂日|改版日|発行日|作成者|更新者|最終更新者|改訂者|承認者|承認日)$'
$script:HistoryDrafts = New-Object System.Collections.Generic.List[object]

function Test-BookkeepingCell {
    #     版数・更新日・更新者などの欄のセルか（同じ行で左にある、いちばん近い値を見出しとみなす）
    param($Sheet, [string] $Cell)
    $m = [regex]::Match([string]$Cell, '^([A-Z]+)(\d+)$')
    if (-not $m.Success -or -not $Sheet) { return $false }
    $row = $Sheet.Rows[$m.Groups[2].Value]
    if (-not $row) { return $false }
    $col = Get-ColumnIndex $m.Groups[1].Value
    $left = @($row.V.Keys | Where-Object { (Get-ColumnIndex $_) -lt $col } | Sort-Object { Get-ColumnIndex $_ })
    if ($left.Count -eq 0) { return $false }
    return (ConvertTo-Label $row.V[$left[-1]]) -match $script:BookkeepingLabel
}

function Get-ChangeGroups {
    #     改訂履歴のシート以外の変更（値・数式・行・列・シート）を、行（列・シートの変更はそれぞれ 1 件）ごとにまとめ、
    #     行を表す語（改訂履歴と突き合わせる語）と、改訂履歴に書く文をつくる。版数・更新日などの欄の変更は除く。
    #     返り値: [ordered] キー → @{ Sheet; OldSheet; NewCell; OldCell; Terms; Label; Texts }
    param($OldBook, $NewBook, $Diffs)

    $hist = @($NewBook.Keys | Where-Object { $_ -match $HistorySheet })[0]
    $groups = [ordered]@{}
    foreach ($d in $Diffs) {
        if ($d.Sheet -eq $hist -or @('値', '数式', '行', '列', 'シート') -notcontains $d.Kind) { continue }
        $useOld = $d.Change -eq '削除' -and $d.Kind -ne '値' -and $d.Kind -ne '数式'
        $sheetName = $d.Sheet
        $book = $NewBook
        if ($useOld) { $sheetName = $d.OldSheet; $book = $OldBook }
        $sheet = $book[$sheetName]
        if (($d.Kind -eq '値' -or $d.Kind -eq '数式') -and (Test-BookkeepingCell $sheet $d.Cell)) { continue }
        $m = [regex]::Match([string]$d.Cell, '^([A-Z]+)(\d+)')
        $col = $m.Groups[1].Value; $row = $m.Groups[2].Value
        $show = { param($t) if ($t) { Get-Short $t 30 } else { '（空）' } }
        switch ($d.Kind) {
            'シート' {
                $key = 'S|' + $sheetName
                $terms = @($d.Sheet)
                if ($d.Change -eq '名前変更') { $terms += $d.Old; $text = 'シート名を {0} → {1} に変更' -f $d.Old, $d.New }
                else { $text = 'シートを{0}' -f $d.Change }
                $label = ''
            }
            '列' {
                $key = 'C|{0}|{1}|{2}' -f $useOld, $sheetName, $col
                $first = @($sheet.Rows.Keys | Where-Object { $sheet.Rows[$_].V.ContainsKey($col) } | Sort-Object { [int]$_ })[0]
                $header = $null
                if ($first) { $header = $sheet.Rows[$first].V[$col] }
                $terms = @($header | Where-Object { Test-HistoryTerm $_ })
                $label = $header
                $text = '列を{0}' -f $d.Change
            }
            default {
                $key = 'R|{0}|{1}|{2}' -f $useOld, $sheetName, $row
                $data = $sheet.Rows[$row]
                # 行を表す語から、変わった値そのものは除く
                $rowTerms = Get-RowLabelTerms $data
                $terms = @($rowTerms | Where-Object { $_ -cne ([string]$d.New).Trim() -and $_ -cne ([string]$d.Old).Trim() })
                $header = $null
                # 見出しと値だけの行（表紙の「システム名 | 受注管理」など）は、左の見出しを項目名にする
                if ($d.Kind -ne '行' -and $data -and $data.V.Count -le 2) {
                    $left = @($data.V.Keys | Where-Object { (Get-ColumnIndex $_) -lt (Get-ColumnIndex $col) } | Sort-Object { Get-ColumnIndex $_ })
                    if ($left.Count -gt 0 -and (Test-HistoryTerm $data.V[$left[-1]])) { $header = $data.V[$left[-1]].Trim(); $terms = @($header) }
                }
                $label = ''
                if (-not $header) {
                    if ($terms.Count -gt 0) { $label = $terms[0] }
                    if ($terms.Count -gt 1) { $label += '（{0}）' -f $terms[1] }
                }
                if ($d.Kind -eq '行') { $text = '行を{0}' -f $d.Change }
                else {
                    if (-not $header) { $header = Get-ColumnHeaderText $sheet $col ([int]$row) }
                    if (-not $header) { $header = "$col 列" }
                    switch ($d.Change) {
                        '変更' {
                            if ($d.Kind -eq '数式') { $text = '{0}の数式を変更' -f $header }
                            else { $text = '{0}を {1} → {2} に変更' -f $header, (& $show $d.Old), (& $show $d.New) }
                        }
                        '追加' { $text = '{0}に「{1}」を追加' -f $header, (& $show $d.New) }
                        '削除' { $text = '{0}の「{1}」を削除' -f $header, (& $show $d.Old) }
                    }
                }
            }
        }
        if (-not $groups.Contains($key)) {
            $newCell = $d.Cell; $oldCell = $null
            if ($useOld) { $newCell = $null; $oldCell = $d.Cell }
            $groups[$key] = [pscustomobject]@{ Sheet = $d.Sheet; OldSheet = $d.OldSheet; NewCell = $newCell; OldCell = $oldCell; Terms = $terms; Label = $label; Texts = New-Object System.Collections.Generic.List[string] }
        }
        if ($text -and -not $groups[$key].Texts.Contains($text)) { $groups[$key].Texts.Add($text) }
    }
    return $groups
}

function Test-HistoryCoverage {
    #     改訂履歴のシートに追記・変更した行に、ほかのシートの変更が書かれているかを確かめる。
    #     書かれていない変更は、行（シート・列）ごとに「改訂履歴 / 記載なし」として、改訂履歴に書く文の下書きとともに出力する
    param($NewBook, $Diffs, $Groups)

    if ($Groups.Count -eq 0) { return }
    $hist = @($NewBook.Keys | Where-Object { $_ -match $HistorySheet })[0]
    if (-not $hist) {
        Add-Diff '' $null $null '改訂履歴' '記載なし' $null $null -DetailText ('改訂履歴のシートが見つかりません（-HistorySheet の「{0}」に一致する名前のシート）' -f $HistorySheet)
        return
    }
    $written = @($Diffs | Where-Object { $_.Sheet -eq $hist -and ($_.Change -eq '追加' -or $_.Change -eq '変更') -and $_.New } | ForEach-Object { $_.New })
    if ($written.Count -eq 0) {
        $count = @($Groups.Values | ForEach-Object { $_.Texts.Count } | Measure-Object -Sum).Sum
        Add-Diff $hist $null $null '改訂履歴' '記載なし' $null $null -DetailText ('改訂履歴に追記がありません（ほかのシートの変更: {0} 件）' -f $count)
        return
    }
    $nfkc = [System.Text.NormalizationForm]::FormKC
    $historyText = ($written -join "`n").Normalize($nfkc).ToLowerInvariant()

    foreach ($g in $Groups.Values) {
        # 行を表す語が無い（数値だけの行など）変更は判定できないので出さない
        if ($g.Terms.Count -eq 0) { continue }
        $mentioned = $false
        foreach ($t in $g.Terms) { if ($historyText.Contains($t.Normalize($nfkc).ToLowerInvariant())) { $mentioned = $true; break } }
        if ($mentioned) { continue }
        $subject = $g.Sheet
        if ($g.Label) { $subject = '{0} {1}' -f $g.Sheet, $g.Label }
        $draft = '{0}: {1}' -f $subject, ($g.Texts -join '、')
        $script:OldSheetName = $g.OldSheet
        Add-Diff $g.Sheet $g.NewCell $g.OldCell '改訂履歴' '記載なし' $null $draft -DetailText ('改訂履歴に記載が見当たりません: ' + $draft)
        $script:OldSheetName = $null
    }
}

# ---------------------------------------------------------------- 比べるファイルの組み合わせ


# ---------------------------------------------------------------- 差分の印を付けた写し（-OutMarkedDir）


# 印の塗りつぶしの色（Excel のセルのスタイル「どちらでもない」「良い」「悪い」と同じ色）と、入力時メッセージの見出し
$script:MarkKinds = [ordered]@{
    Changed = @{ Color = 'FFFFEB9C'; Title = '変更' }
    Added   = @{ Color = 'FFC6EFCE'; Title = '追加' }
    Deleted = @{ Color = 'FFFFC7CE'; Title = '削除' }
    Moved   = @{ Color = 'FFDDEBF7'; Title = '移動' }
}
# ワークシートの子要素のうち、dataValidations より後ろに置くもの（スキーマで順序が決まっている）
$script:AfterDataValidations = 'hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|customProperties|cellWatches|ignoredErrors|smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst'

function Get-DiffMarks {
    #     差分から、新のブックのシートごとに印を付けるセル・行・列と、セルを選んだときに出すメッセージを決める。
    #     返り値: シート名 → @{ Cells = 番地 →印; Rows = 行番号 → 印; Cols = 列 → 印 }（印: @{ Kind; Lines }）
    param($Diffs)

    $marks = @{}
    foreach ($d in $Diffs) {
        if (-not $d.Sheet -or -not $d.Cell) { continue }
        if (-not $marks.ContainsKey($d.Sheet)) { $marks[$d.Sheet] = @{ Cells = @{}; Rows = @{}; Cols = @{} } }
        $m = $marks[$d.Sheet]
        $table = $null; $keys = @(); $kind = $null; $line = $null
        switch ($d.Kind) {
            { $_ -eq '値' -or $_ -eq '数式' } {
                $table = $m.Cells; $keys = @($d.Cell)
                switch ($d.Change) {
                    '変更' { $kind = 'Changed'; $line = '変更前の{0}: {1}' -f $d.Kind, $d.Old }
                    '追加' { $kind = 'Added'; $line = '{0}を追加' -f $d.Kind }
                    '削除' { $kind = 'Deleted'; $line = '削除した{0}: {1}' -f $d.Kind, $d.Old }
                }
            }
            '行' {
                $table = $m.Rows; $keys = @([regex]::Match($d.Cell, '\d+').Value)
                if ($d.Change -eq '追加') { $kind = 'Added'; $line = '追加した行' }
                elseif ($d.Change -eq '移動') { $kind = 'Moved'; $line = '移動した行（旧 {0} 行目）' -f [regex]::Match($d.OldLocation, '\d+$').Value }
            }
            '列' {
                $table = $m.Cols; $keys = @([regex]::Match($d.Cell, '^[A-Z]+').Value)
                if ($d.Change -eq '追加') { $kind = 'Added'; $line = '追加した列' }
            }
            '連番' {
                $mm = [regex]::Match($d.Cell, '^([A-Z]+)(\d+):[A-Z]+(\d+)$')
                if ($mm.Success) {
                    $table = $m.Cells; $kind = 'Changed'; $line = $d.Detail
                    $keys = @([int]$mm.Groups[2].Value..[int]$mm.Groups[3].Value | ForEach-Object { $mm.Groups[1].Value + $_ })
                }
            }
        }
        if (-not $kind) { continue }
        foreach ($k in $keys) {
            if (-not $table.ContainsKey($k)) { $table[$k] = @{ Kind = $kind; Lines = New-Object System.Collections.Generic.List[string] } }
            elseif ($table[$k].Kind -ne $kind) { $table[$k].Kind = 'Changed' }   # 追加と削除が重なったセルなどは「変更」
            $table[$k].Lines.Add($line)
        }
    }
    return $marks
}

function Set-TagAttribute {
    #     要素の開始タグ（文字列の先頭のタグ）の属性を設定する。無ければ足す
    param([string] $Element, [string] $Name, [string] $Value)
    $end = $Element.IndexOf('>')
    $tag = $Element.Substring(0, $end + 1)
    $rest = $Element.Substring($end + 1)
    $rx = New-Object System.Text.RegularExpressions.Regex(('\s{0}="[^"]*"' -f [regex]::Escape($Name)))
    if ($rx.IsMatch($tag)) { $tag = $rx.Replace($tag, (' {0}="{1}"' -f $Name, $Value), 1) }
    elseif ($tag.EndsWith('/>')) { $tag = $tag.Substring(0, $tag.Length - 2) + (' {0}="{1}"/>' -f $Name, $Value) }
    else { $tag = $tag.Substring(0, $tag.Length - 1) + (' {0}="{1}">' -f $Name, $Value) }
    return $tag + $rest
}

function Get-TagAttribute {
    param([string] $Element, [string] $Name)
    $end = $Element.IndexOf('>')
    $m = [regex]::Match($Element.Substring(0, $end + 1), ('\s{0}="([^"]*)"' -f [regex]::Escape($Name)))
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Add-CountedChildren {
    #     <fills count="2">…</fills> のような要素の末尾に子要素を足し、count を直す。返り値: @(新しい XML, 足す前の子要素の数)
    param([string] $Xml, [string] $Prefix, [string] $Name, [string] $Child, [string[]] $NewItems)

    $rx = New-Object System.Text.RegularExpressions.Regex(('<{0}{1}\b[^>]*?(?:/>|>(.*?)</{0}{1}>)' -f $Prefix, $Name), [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $m = $rx.Match($Xml)
    if (-not $m.Success) { throw "styles.xml に $Name がありません" }
    $count = ([regex]::Matches($m.Groups[1].Value, ('<{0}{1}\b' -f $Prefix, $Child))).Count
    $element = $m.Value
    if ($element.EndsWith('/>') -and $m.Groups[1].Length -eq 0 -and -not $element.Contains('</')) {
        $element = $element.Substring(0, $element.Length - 2) + ('></{0}{1}>' -f $Prefix, $Name)
    }
    $close = '</{0}{1}>' -f $Prefix, $Name
    $element = $element.Substring(0, $element.Length - $close.Length) + ($NewItems -join '') + $close
    $element = Set-TagAttribute $element 'count' ($count + $NewItems.Count)
    return @(($Xml.Substring(0, $m.Index) + $element + $Xml.Substring($m.Index + $m.Length)), $count)
}

function ConvertTo-Sqref {
    #     セルの番地の一覧を、同じ行で隣り合うものを範囲にまとめた sqref（"A5:F5 H5"）にする
    param($Refs)
    $cells = @($Refs | ForEach-Object {
        $m = [regex]::Match($_, '^([A-Z]+)(\d+)$')
        [pscustomobject]@{ Row = [int]$m.Groups[2].Value; Col = (Get-ColumnIndex $m.Groups[1].Value) }
    } | Sort-Object Row, Col)
    $parts = @()
    $i = 0
    while ($i -lt $cells.Count) {
        $j = $i
        while ($j + 1 -lt $cells.Count -and $cells[$j + 1].Row -eq $cells[$i].Row -and $cells[$j + 1].Col -eq $cells[$j].Col + 1) { $j++ }
        $ref = (ConvertTo-ColumnName $cells[$i].Col) + $cells[$i].Row
        if ($j -gt $i) { $ref += ':' + (ConvertTo-ColumnName $cells[$j].Col) + $cells[$j].Row }
        $parts += $ref
        $i = $j + 1
    }
    return ($parts -join ' ')
}

function Get-SqrefAreas {
    #     sqref（"A1:B5 D7"）を、行・列の範囲の一覧にする
    param([string] $Sqref)
    $areas = @()
    foreach ($part in ($Sqref -split '\s+')) {
        $m = [regex]::Match($part, '^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$')
        if (-not $m.Success) { continue }
        $c0 = Get-ColumnIndex $m.Groups[1].Value; $r0 = [int]$m.Groups[2].Value
        $c1 = $c0; $r1 = $r0
        if ($m.Groups[3].Success) { $c1 = Get-ColumnIndex $m.Groups[3].Value; $r1 = [int]$m.Groups[4].Value }
        $areas += , @([math]::Min($r0, $r1), [math]::Max($r0, $r1), [math]::Min($c0, $c1), [math]::Max($c0, $c1))
    }
    return , $areas
}

function Get-MarkStyle {
    #     元のスタイル番号と印の種類から、印を付けたスタイルの番号を返す（無ければ割り当てる）
    param($State, [int] $Style, [string] $Kind)
    $key = '{0}|{1}' -f $Style, $Kind
    if (-not $State.Map.ContainsKey($key)) {
        $State.Map[$key] = $State.Next
        $State.Order.Add(@($Style, $Kind))
        $State.Next++
    }
    return $State.Map[$key]
}

function Add-SheetMarks {
    #     シートの XML のセルに印のスタイルを付け、印のメッセージを入力時メッセージ（データの入力規則）として足す。
    #     印を付けるセルが XML に無ければ、空のセルとして作る（値を削除したセルなど）
    param([string] $Xml, $Marks, $State)

    $px = [regex]::Match($Xml, '<(\w+:)?worksheet\b').Groups[1].Value
    $sd = [regex]::Match($Xml, ('<{0}sheetData\b[^>]*?(?:/>|>(.*?)</{0}sheetData>)' -f $px), [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $sd.Success) { return $Xml }

    $cellsByRow = @{}
    foreach ($ref in $Marks.Cells.Keys) {
        $row = [regex]::Match($ref, '\d+$').Value
        if (-not $cellsByRow.ContainsKey($row)) { $cellsByRow[$row] = New-Object System.Collections.Generic.List[string] }
        $cellsByRow[$row].Add($ref)
    }
    $messages = @{}         # 番地 → 印（メッセージを出すセル）
    $seenRows = @{}
    $cellRx = New-Object System.Text.RegularExpressions.Regex(('<{0}c\b[^>]*?(?:/>|>.*?</{0}c>)' -f $px), [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $rowRx = New-Object System.Text.RegularExpressions.Regex(('<{0}row\b[^>]*?(?:/>|>.*?</{0}row>)' -f $px), [System.Text.RegularExpressions.RegexOptions]::Singleline)

    $markCell = {
        param([string] $Cell, [string] $Ref, $Mark)
        $s = 0
        [void][int]::TryParse((Get-TagAttribute $Cell 's'), [ref]$s)
        $messages[$Ref] = $Mark
        return (Set-TagAttribute $Cell 's' (Get-MarkStyle $State $s $Mark.Kind))
    }
    $newCell = {
        param([string] $Ref, $Mark)
        $messages[$Ref] = $Mark
        return ('<{0}c r="{1}" s="{2}"/>' -f $px, $Ref, (Get-MarkStyle $State 0 $Mark.Kind))
    }

    $inner = $sd.Groups[1].Value
    $rowsXml = New-Object System.Collections.Generic.List[object]
    $allRows = $Marks.Cols.Count -gt 0
    foreach ($rm in $rowRx.Matches($inner)) {
        $rowXml = $rm.Value
        $r = Get-TagAttribute $rowXml 'r'
        if (-not $r) { return $Xml }   # 行番号を省略したシートは、位置が確かでないので印を付けない
        $seenRows[$r] = $true
        $rowMark = $Marks.Rows[$r]
        if (-not $allRows -and -not $rowMark -and -not $cellsByRow.ContainsKey($r)) {
            # 印を付けない行はそのまま（作る行を正しい位置に入れるため、位置だけ覚えておく）
            $rowsXml.Add([pscustomobject]@{ Row = [int]$r; Xml = $rowXml; Index = $rm.Index; Length = $rm.Length })
            continue
        }
        $found = @{}
        $endTag = $rowXml.IndexOf('>')
        if (-not $rowXml.Substring(0, $endTag + 1).EndsWith('/>')) {
            $rowXml = $cellRx.Replace($rowXml, {
                param($cm)
                $ref = Get-TagAttribute $cm.Value 'r'
                if (-not $ref) { return $cm.Value }
                $found[$ref] = $true
                $mark = $Marks.Cells[$ref]
                if (-not $mark) { $mark = $rowMark }
                if (-not $mark) { $mark = $Marks.Cols[[regex]::Match($ref, '^[A-Z]+').Value] }
                if (-not $mark) { return $cm.Value }
                return (& $markCell $cm.Value $ref $mark)
            })
        }
        # 印を付けるセルが無ければ作り、列の順に入れる
        $missing = @()
        if ($cellsByRow.ContainsKey($r)) { $missing = @($cellsByRow[$r] | Where-Object { -not $found.ContainsKey($_) }) }
        if ($missing.Count -gt 0) {
            $rowXml = [regex]::Replace($rowXml, '\sspans="[^"]*"', '')
            $tagEnd = $rowXml.IndexOf('>')
            if ($rowXml.Substring(0, $tagEnd + 1).EndsWith('/>')) {
                $rowXml = $rowXml.Substring(0, $tagEnd - 1) + ('></{0}row>' -f $px)
                $tagEnd = $rowXml.IndexOf('>')
            }
            $startTag = $rowXml.Substring(0, $tagEnd + 1)
            $close = '</{0}row>' -f $px
            $cells = New-Object System.Collections.Generic.List[object]
            foreach ($cm in $cellRx.Matches($rowXml)) { $cells.Add([pscustomobject]@{ Col = (Get-ColumnIndex (Get-TagAttribute $cm.Value 'r')); Xml = $cm.Value }) }
            foreach ($ref in $missing) { $cells.Add([pscustomobject]@{ Col = (Get-ColumnIndex $ref); Xml = (& $newCell $ref $Marks.Cells[$ref]) }) }
            $rowXml = $startTag + (($cells | Sort-Object Col | ForEach-Object { $_.Xml }) -join '') + $close
        }
        $rowsXml.Add([pscustomobject]@{ Row = [int]$r; Xml = $rowXml; Index = $rm.Index; Length = $rm.Length })
    }

    # 行そのものが無ければ作る
    $newRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $cellsByRow.Keys) {
        if ($seenRows.ContainsKey($r)) { continue }
        $cells = @($cellsByRow[$r] | Sort-Object { Get-ColumnIndex $_ } | ForEach-Object { & $newCell $_ $Marks.Cells[$_] })
        $newRows.Add([pscustomobject]@{ Row = [int]$r; Xml = ('<{0}row r="{1}">{2}</{0}row>' -f $px, $r, ($cells -join '')) })
    }
    # 書き換えた行を戻し、作った行を行番号の順に入れる（行と行の間の文字はそのまま残す）
    $sb = New-Object System.Text.StringBuilder
    $pos = 0
    $pending = @($newRows | Sort-Object Row)
    $p = 0
    foreach ($row in $rowsXml) {
        [void]$sb.Append($inner.Substring($pos, $row.Index - $pos))
        while ($p -lt $pending.Count -and $pending[$p].Row -lt $row.Row) { [void]$sb.Append($pending[$p].Xml); $p++ }
        [void]$sb.Append($row.Xml)
        $pos = $row.Index + $row.Length
    }
    [void]$sb.Append($inner.Substring($pos))
    while ($p -lt $pending.Count) { [void]$sb.Append($pending[$p].Xml); $p++ }
    $sheetData = '<{0}sheetData>{1}</{0}sheetData>' -f $px, $sb.ToString()
    $xmlOut = $Xml.Substring(0, $sd.Index) + $sheetData + $Xml.Substring($sd.Index + $sd.Length)

    # ---- 入力時メッセージ。既に入力規則のあるセルには付けない（重なると Excel が修復を求めるため）
    if ($messages.Count -eq 0) { return $xmlOut }
    $taken = @()
    foreach ($m in [regex]::Matches($xmlOut, '<(?:\w+:)?dataValidation\b[^>]*\ssqref="([^"]*)"')) { $taken += Get-SqrefAreas $m.Groups[1].Value }
    foreach ($m in [regex]::Matches($xmlOut, '<(?:\w+:)?sqref>([^<]*)</(?:\w+:)?sqref>')) { $taken += Get-SqrefAreas $m.Groups[1].Value }
    $groups = [ordered]@{}   # 見出しとメッセージ → 番地の一覧
    foreach ($ref in $messages.Keys) {
        $cm = [regex]::Match($ref, '^([A-Z]+)(\d+)$')
        $row = [int]$cm.Groups[2].Value; $col = Get-ColumnIndex $cm.Groups[1].Value
        $covered = $false
        foreach ($a in $taken) { if ($row -ge $a[0] -and $row -le $a[1] -and $col -ge $a[2] -and $col -le $a[3]) { $covered = $true; break } }
        if ($covered) { continue }
        $mark = $messages[$ref]
        $title = $script:MarkKinds[$mark.Kind].Title
        $prompt = (@($mark.Lines | Select-Object -Unique) -join "`n")
        if ($prompt.Length -gt 255) { $prompt = $prompt.Substring(0, 254) + '…' }
        $key = $title + "`t" + $prompt
        if (-not $groups.Contains($key)) { $groups[$key] = New-Object System.Collections.Generic.List[string] }
        $groups[$key].Add($ref)
    }
    if ($groups.Count -eq 0) { return $xmlOut }
    $items = @()
    foreach ($key in $groups.Keys) {
        $title, $prompt = $key -split "`t", 2
        $items += ('<{0}dataValidation allowBlank="1" showInputMessage="1" promptTitle="{1}" prompt="{2}" sqref="{3}"/>' -f $px, (ConvertTo-XmlText $title), ((ConvertTo-XmlText $prompt).Replace("`n", '&#10;')), (ConvertTo-Sqref $groups[$key]))
    }
    $dv = [regex]::Match($xmlOut, ('<{0}dataValidations\b[^>]*>.*?</{0}dataValidations>' -f $px), [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if ($dv.Success) {
        $count = ([regex]::Matches($dv.Value, ('<{0}dataValidation\b' -f $px))).Count
        $close = '</{0}dataValidations>' -f $px
        $element = $dv.Value.Substring(0, $dv.Value.Length - $close.Length) + ($items -join '') + $close
        $element = Set-TagAttribute $element 'count' ($count + $items.Count)
        return $xmlOut.Substring(0, $dv.Index) + $element + $xmlOut.Substring($dv.Index + $dv.Length)
    }
    $element = '<{0}dataValidations count="{1}">{2}</{0}dataValidations>' -f $px, $items.Count, ($items -join '')
    $sdEnd = $xmlOut.IndexOf(('</{0}sheetData>' -f $px)) + ('</{0}sheetData>' -f $px).Length
    $after = [regex]::Match($xmlOut.Substring($sdEnd), ('<{0}(?:{1})\b' -f $px, $script:AfterDataValidations))
    if ($after.Success) { $at = $sdEnd + $after.Index }
    else { $at = $xmlOut.LastIndexOf(('</{0}worksheet>' -f $px)) }
    return $xmlOut.Substring(0, $at) + $element + $xmlOut.Substring($at)
}

function New-DiffListSheetXml {
    #     差分一覧のシートの XML を作る。場所のセルは、この写しの中の該当セルへのリンクにする
    param($Diffs, [string] $Namespace, [int] $HeaderStyle, [int] $BodyStyle, [int] $LinkStyle)

    $columns = @(@('場所', 24), @('種類', 9), @('変更', 8), @('内容', 50), @('変更前', 40), @('変更後', 40), @('旧の場所', 18))
    $sb = New-Object System.Text.StringBuilder
    $links = New-Object System.Text.StringBuilder
    [void]$sb.Append(('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{0}">' -f $Namespace))
    [void]$sb.Append(('<dimension ref="A1:G{0}"/>' -f ($Diffs.Count + 1)))
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="18"/><cols>')
    for ($i = 0; $i -lt $columns.Count; $i++) { [void]$sb.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $columns[$i][1])) }
    [void]$sb.Append('</cols><sheetData><row r="1">')
    for ($i = 0; $i -lt $columns.Count; $i++) {
        [void]$sb.Append(('<c r="{0}1" s="{1}" t="inlineStr"><is><t>{2}</t></is></c>' -f (ConvertTo-ColumnName ($i + 1)), $HeaderStyle, $columns[$i][0]))
    }
    [void]$sb.Append('</row>')
    $r = 1
    foreach ($d in $Diffs) {
        $r++
        $oldValue = $d.Old; $newValue = $d.New
        if ($d.Segments) {
            $oldValue = @{ Runs = @($d.Segments | Where-Object { $_.Op -ne '+' } | ForEach-Object { @{ Text = $_.Text; Style = $(if ($_.Op -eq '-') { 'Deleted' } else { '' }) } }) }
            $newValue = @{ Runs = @($d.Segments | Where-Object { $_.Op -ne '-' } | ForEach-Object { @{ Text = $_.Text; Style = $(if ($_.Op -eq '+') { 'Inserted' } else { '' }) } }) }
        }
        # この写しにある場所だけリンクにする（削除した行・列・シート・コメント・図形は旧のブックにしかない）
        $hasNew = $d.Change -ne '削除' -or $d.Kind -eq '値' -or $d.Kind -eq '数式'
        $location = $d.Location
        if (-not $hasNew) { $location = '旧: ' + $location }
        $values = @($location, $d.Kind, $d.Change, $d.Detail, $oldValue, $newValue, $d.OldLocation)
        $target = ''
        if ($hasNew -and $d.Sheet) {
            $cell = $d.Cell
            if (-not $cell) { $cell = 'A1' }
            $target = "'{0}'!{1}" -f $d.Sheet.Replace("'", "''"), $cell
        }
        [void]$sb.Append(('<row r="{0}">' -f $r))
        for ($i = 0; $i -lt $values.Count; $i++) {
            $content = ConvertTo-InlineStringXml $values[$i]
            if (-not $content) { continue }
            $style = $BodyStyle
            if ($i -eq 0 -and $target) { $style = $LinkStyle }
            [void]$sb.Append(('<c r="{0}{1}" s="{2}" t="inlineStr">{3}</c>' -f (ConvertTo-ColumnName ($i + 1)), $r, $style, $content))
        }
        [void]$sb.Append('</row>')
        if ($target) { [void]$links.Append(('<hyperlink ref="A{0}" location="{1}" display="{2}"/>' -f $r, (ConvertTo-XmlText $target), (ConvertTo-XmlText $location))) }
    }
    [void]$sb.Append('</sheetData>')
    [void]$sb.Append(('<autoFilter ref="A1:G{0}"/>' -f $r))
    if ($links.Length -gt 0) { [void]$sb.Append('<hyperlinks>').Append($links.ToString()).Append('</hyperlinks>') }
    [void]$sb.Append('<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/></worksheet>')
    return $sb.ToString()
}

function Export-MarkedCopy {
    #     新のブックの写しを作り、差分のあったセルに色と入力時メッセージ（変更前の値など）を付け、末尾に差分一覧のシートを足す。
    #     元のブックのパーツは、印を付けるシート・styles.xml・ブックの構成を表すパーツ以外は 1 バイトも変えない
    param([string] $SourcePath, [string] $OutPath, $Diffs)

    $marks = Get-DiffMarks $Diffs
    $changes = @{}       # パーツ名（小文字）→ 新しい内容
    $newParts = [ordered]@{}
    $pkg = Open-ZipPackage -FilePath $SourcePath
    try {
        $read = {
            param([string] $Part)
            $e = Get-PackageEntry -Package $pkg -PartName $Part
            if ($null -eq $e) { return $null }
            $sr = New-Object System.IO.StreamReader($e.Open(), (New-Object System.Text.UTF8Encoding($false)), $true)
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        }
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) { if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break } }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        $relById = @{}
        $stylesPart = $null
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName $wbPart)) {
            $relById[$rel.Id] = $rel
            if ($rel.Kind -eq 'styles') { $stylesPart = $rel.Target }
        }
        if (-not $stylesPart) { throw 'styles.xml がありません' }
        $sheetPart = @{}
        $sheetNames = @()
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $sheetNames += $s.GetAttribute('name')
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            if ($rel -and $rel.Kind -eq 'worksheet') { $sheetPart[$s.GetAttribute('name')] = $rel.Target }
        }

        # ---- styles.xml の現在のスタイル数を数え、シートに印を付けながら必要なスタイルを割り当てる
        $styles = & $read $stylesPart
        $spx = [regex]::Match($styles, '<(\w+:)?styleSheet\b').Groups[1].Value
        $xfsMatch = [regex]::Match($styles, ('<{0}cellXfs\b[^>]*>(.*?)</{0}cellXfs>' -f $spx), [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if (-not $xfsMatch.Success) { throw 'styles.xml に cellXfs がありません' }
        $xfs = @([regex]::Matches($xfsMatch.Groups[1].Value, ('<{0}xf\b[^>]*?(?:/>|>.*?</{0}xf>)' -f $spx), [System.Text.RegularExpressions.RegexOptions]::Singleline) | ForEach-Object { $_.Value })
        $state = @{ Map = @{}; Order = New-Object System.Collections.Generic.List[object]; Next = $xfs.Count }

        foreach ($sheet in $marks.Keys) {
            if (-not $sheetPart.ContainsKey($sheet)) { continue }
            $xml = & $read $sheetPart[$sheet]
            if ($null -eq $xml) { continue }
            $new = Add-SheetMarks -Xml $xml -Marks $marks[$sheet] -State $state
            if ($new -cne $xml) { $changes[$sheetPart[$sheet].ToLowerInvariant()] = $new }
        }

        # ---- 塗りつぶし・フォント・スタイルを足す
        $fillXml = @($script:MarkKinds.Keys | ForEach-Object { '<{0}fill><{0}patternFill patternType="solid"><{0}fgColor rgb="{1}"/><{0}bgColor indexed="64"/></{0}patternFill></{0}fill>' -f $spx, $script:MarkKinds[$_].Color })
        $fillXml += '<{0}fill><{0}patternFill patternType="solid"><{0}fgColor rgb="FFD9E1F2"/><{0}bgColor indexed="64"/></{0}patternFill></{0}fill>' -f $spx
        $res = Add-CountedChildren $styles $spx 'fills' 'fill' $fillXml
        $styles = $res[0]; $fillBase = $res[1]
        $fillOf = @{}
        $k = 0
        foreach ($kind in $script:MarkKinds.Keys) { $fillOf[$kind] = $fillBase + $k; $k++ }
        $headerFill = $fillBase + $k

        $fontXml = @(
            '<{0}font><{0}b/><{0}sz val="10"/><{0}name val="游ゴシック"/><{0}family val="3"/><{0}charset val="128"/></{0}font>' -f $spx,
            '<{0}font><{0}sz val="10"/><{0}name val="游ゴシック"/><{0}family val="3"/><{0}charset val="128"/></{0}font>' -f $spx,
            '<{0}font><{0}u/><{0}sz val="10"/><{0}color rgb="FF0563C1"/><{0}name val="游ゴシック"/><{0}family val="3"/><{0}charset val="128"/></{0}font>' -f $spx)
        $res = Add-CountedChildren $styles $spx 'fonts' 'font' $fontXml
        $styles = $res[0]; $fontBase = $res[1]

        $newXfs = @()
        foreach ($o in $state.Order) {
            $x = $xfs[0]
            if ($o[0] -lt $xfs.Count) { $x = $xfs[$o[0]] }
            $x = Set-TagAttribute $x 'fillId' $fillOf[$o[1]]
            $newXfs += Set-TagAttribute $x 'applyFill' '1'
        }
        $listBase = $state.Next
        $newXfs += '<{0}xf numFmtId="0" fontId="{1}" fillId="{2}" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><{0}alignment vertical="center" wrapText="1"/></{0}xf>' -f $spx, $fontBase, $headerFill
        $newXfs += '<{0}xf numFmtId="0" fontId="{1}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><{0}alignment vertical="top" wrapText="1"/></{0}xf>' -f $spx, ($fontBase + 1)
        $newXfs += '<{0}xf numFmtId="0" fontId="{1}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><{0}alignment vertical="top" wrapText="1"/></{0}xf>' -f $spx, ($fontBase + 2)
        $res = Add-CountedChildren $styles $spx 'cellXfs' 'xf' $newXfs
        $changes[$stylesPart.ToLowerInvariant()] = $res[0]

        # ---- 差分一覧のシートを末尾に足す（シート名・パーツ名・リレーション ID は既存と重ならないようにする）
        $listName = '差分一覧'
        $n = 2
        while (@($sheetNames | Where-Object { $_ -ieq $listName }).Count -gt 0) { $listName = '差分一覧 ({0})' -f $n; $n++ }
        $wbDir = ''
        if ($wbPart.Contains('/')) { $wbDir = $wbPart.Substring(0, $wbPart.LastIndexOf('/') + 1) }
        $n = 1
        while ($pkg.Entries.ContainsKey(('{0}worksheets/sheet{1}.xml' -f $wbDir, $n).ToLowerInvariant())) { $n++ }
        $listPart = '{0}worksheets/sheet{1}.xml' -f $wbDir, $n

        $wbXml = & $read $wbPart
        $mainNs = [regex]::Match($wbXml, '<(?:\w+:)?workbook\b[^>]*\sxmlns(?::\w+)?="([^"]*spreadsheetml[^"]*)"').Groups[1].Value
        if (-not $mainNs) { $mainNs = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main' }
        $newParts[$listPart] = New-DiffListSheetXml -Diffs $Diffs -Namespace $mainNs -HeaderStyle $listBase -BodyStyle ($listBase + 1) -LinkStyle ($listBase + 2)

        $relsPart = '{0}_rels/{1}.rels' -f $wbDir, $wbPart.Substring($wbDir.Length)
        $relsXml = & $read $relsPart
        $ids = @{}
        foreach ($m in [regex]::Matches($relsXml, '\sId="([^"]*)"')) { $ids[$m.Groups[1].Value] = $true }
        $n = $ids.Count + 1
        while ($ids.ContainsKey("rId$n")) { $n++ }
        $rid = "rId$n"
        $wsType = [regex]::Match($relsXml, 'Type="([^"]*/worksheet)"').Groups[1].Value
        if (-not $wsType) { $wsType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' }
        $at = $relsXml.LastIndexOf('</Relationships>')
        $changes[$relsPart.ToLowerInvariant()] = $relsXml.Substring(0, $at) + ('<Relationship Id="{0}" Type="{1}" Target="{2}"/>' -f $rid, $wsType, $listPart.Substring($wbDir.Length)) + $relsXml.Substring($at)

        $wpx = [regex]::Match($wbXml, '<(\w+:)?workbook\b').Groups[1].Value
        # r:id の接頭辞は、ルート要素で宣言されているものだけ使う（無ければ sheet 要素で宣言する）
        $rootTag = [regex]::Match($wbXml, '<(?:\w+:)?workbook\b[^>]*>').Value
        $rPrefix = [regex]::Match($rootTag, 'xmlns:(\w+)="[^"]*/relationships"').Groups[1].Value
        $maxId = 0
        foreach ($m in [regex]::Matches($wbXml, '\ssheetId="(\d+)"')) { $maxId = [math]::Max($maxId, [int]$m.Groups[1].Value) }
        if ($rPrefix) { $sheetEl = '<{0}sheet name="{1}" sheetId="{2}" {3}:id="{4}"/>' -f $wpx, (ConvertTo-XmlText $listName), ($maxId + 1), $rPrefix, $rid }
        else { $sheetEl = '<{0}sheet name="{1}" sheetId="{2}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{3}"/>' -f $wpx, (ConvertTo-XmlText $listName), ($maxId + 1), $rid }
        $at = $wbXml.IndexOf(('</{0}sheets>' -f $wpx))
        $changes[$wbPart.ToLowerInvariant()] = $wbXml.Substring(0, $at) + $sheetEl + $wbXml.Substring($at)

        $ct = & $read '[Content_Types].xml'
        $at = $ct.LastIndexOf('</Types>')
        $changes['[content_types].xml'] = $ct.Substring(0, $at) + ('<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $listPart) + $ct.Substring($at)
    }
    finally { $pkg.Zip.Dispose() }

    # ---- 書き出す。変更しないパーツは同じ名前・日時・中身で書き直す
    $dir = [System.IO.Path]::GetDirectoryName($OutPath)
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir -Force }
    Save-ZipPackage -FilePath $SourcePath -Changes $changes -NewParts $newParts -OutPath $OutPath
}

function Get-ExcelFiles {
    param([string] $Dir)
    $map = [ordered]@{}
    $root = (Resolve-Path -LiteralPath $Dir).ProviderPath.TrimEnd('\', '/')
    foreach ($f in (Get-ChildItem -LiteralPath $root -File -Recurse:(-not $NoRecurse) | Sort-Object FullName)) {
        if (@('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb') -notcontains $f.Extension.ToLowerInvariant()) { continue }
        if ($f.Name.StartsWith('~$')) { continue }
        if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
        $map[$rel.ToLowerInvariant()] = [pscustomobject]@{ Relative = $rel; Path = $f.FullName }
    }
    return $map
}

if (-not (Test-Path -LiteralPath $Old)) { throw "見つかりません: $Old" }
if (-not (Test-Path -LiteralPath $New)) { throw "見つかりません: $New" }
$oldItem = Get-Item -LiteralPath $Old
$newItem = Get-Item -LiteralPath $New
if ($oldItem.PSIsContainer -ne $newItem.PSIsContainer) { throw '-Old と -New は、両方ともファイルか、両方ともフォルダにしてください。' }

$pairs = New-Object System.Collections.Generic.List[object]
if ($oldItem.PSIsContainer) {
    $oldFiles = Get-ExcelFiles $oldItem.FullName
    $newFiles = Get-ExcelFiles $newItem.FullName
    # 同じ相対パスが無いファイルは、版数（_v1.0 など）を除いた名前が 1 つずつ一致すれば対応付ける
    $byVersionless = @{}
    foreach ($side in @('Old', 'New')) {
        $files = $oldFiles; $other = $newFiles
        if ($side -eq 'New') { $files = $newFiles; $other = $oldFiles }
        foreach ($k in $files.Keys) {
            if ($other.Contains($k)) { continue }
            $vk = Get-VersionlessName $k
            if (-not $byVersionless.ContainsKey($vk)) { $byVersionless[$vk] = @{ Old = @(); New = @() } }
            $byVersionless[$vk][$side] += $k
        }
    }
    $renamedOld = @{}   # 旧の相対パス → 新の相対パス
    $renamedNew = @{}
    foreach ($v in $byVersionless.Values) {
        if ($v.Old.Count -eq 1 -and $v.New.Count -eq 1) { $renamedOld[$v.Old[0]] = $v.New[0]; $renamedNew[$v.New[0]] = $v.Old[0] }
    }
    foreach ($k in $newFiles.Keys) {
        $o = $null
        if ($oldFiles.Contains($k)) { $o = $oldFiles[$k].Path }
        elseif ($renamedNew.ContainsKey($k)) { $o = $oldFiles[$renamedNew[$k]].Path }
        $pairs.Add([pscustomobject]@{ Relative = $newFiles[$k].Relative; Old = $o; New = $newFiles[$k].Path })
    }
    foreach ($k in $oldFiles.Keys) {
        if (-not $newFiles.Contains($k) -and -not $renamedOld.ContainsKey($k)) { $pairs.Add([pscustomobject]@{ Relative = $oldFiles[$k].Relative; Old = $oldFiles[$k].Path; New = $null }) }
    }
}
else {
    $pairs.Add([pscustomobject]@{ Relative = $newItem.Name; Old = $oldItem.FullName; New = $newItem.FullName })
}

$summary = New-Object System.Collections.Generic.List[object]
$index = 0
$markedCount = 0
foreach ($p in $pairs) {
    $index++
    Write-Progress -Activity 'Excel を比較中' -Status ("{0}/{1}  {2}" -f $index, $pairs.Count, $p.Relative) -PercentComplete ([int](100 * $index / $pairs.Count))
    $script:CurrentFile = $p.Relative
    $script:CurrentNewPath = $p.New
    $script:CurrentOldPath = $p.Old
    $before = $script:Diffs.Count
    $state = ''
    try {
        if (-not $p.Old) {
            Add-Diff '' $null $null 'ファイル' '追加' $null $p.Relative; $state = '追加'
            if ($OutHistory) { $script:HistoryDrafts.Add([pscustomobject]@{ File = $p.Relative; Section = '全体'; Description = '新規作成'; Location = '' }) }
        }
        elseif (-not $p.New) { Add-Diff '' $null $null 'ファイル' '削除' $p.Relative $null; $state = '削除' }
        else {
            Compare-Book -OldPath $p.Old -NewPath $p.New
            if ($script:Diffs.Count -eq $before) { $state = '同じ' } else { $state = '差分あり' }
        }
    }
    catch {
        $state = '失敗'
        Write-Warning ("{0}: 比較できませんでした: {1}" -f $p.Relative, $_.Exception.Message)
    }
    if ($OutMarkedDir -and $state -eq '差分あり') {
        $markedPath = [System.IO.Path]::GetFullPath((Join-Path $OutMarkedDir $p.Relative))
        if (@('.xlsx', '.xlsm') -notcontains [System.IO.Path]::GetExtension($p.New).ToLowerInvariant()) {
            Write-Warning ("{0}: 印を付けた写しは .xlsx / .xlsm だけ作れます" -f $p.Relative)
        }
        elseif ($markedPath -ieq [System.IO.Path]::GetFullPath($p.New) -or $markedPath -ieq [System.IO.Path]::GetFullPath($p.Old)) {
            Write-Warning ("{0}: 印を付けた写しの保存先が比べたファイルと同じため、作りません: {1}" -f $p.Relative, $markedPath)
        }
        else {
            try {
                Export-MarkedCopy -SourcePath $p.New -OutPath $markedPath -Diffs $script:Diffs.GetRange($before, $script:Diffs.Count - $before).ToArray()
                $markedCount++
            }
            catch { Write-Warning ("{0}: 印を付けた写しを作れませんでした（同名のファイルを Excel で開いていませんか）: {1}" -f $p.Relative, $_.Exception.Message) }
        }
    }
    $summary.Add([pscustomobject]@{ File = $p.Relative; State = $state; Count = $script:Diffs.Count - $before; Old = $p.Old; New = $p.New })
}
Write-Progress -Activity 'Excel を比較中' -Completed

$script:Diffs

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($d in $script:Diffs) {
        $oldValue = $d.Old; $newValue = $d.New
        if ($d.Segments) {
            # 変わった部分に色を付ける（削除は赤の取り消し線、追加は赤の太字）
            $oldValue = @{ Runs = @($d.Segments | Where-Object { $_.Op -ne '+' } | ForEach-Object { @{ Text = $_.Text; Style = $(if ($_.Op -eq '-') { 'Deleted' } else { '' }) } }) }
            $newValue = @{ Runs = @($d.Segments | Where-Object { $_.Op -ne '-' } | ForEach-Object { @{ Text = $_.Text; Style = $(if ($_.Op -eq '+') { 'Inserted' } else { '' }) } }) }
        }
        $row = @{ Values = @($d.File, $d.Sheet, $d.Location, $d.OldLocation, $d.Kind, $d.Change, $oldValue, $newValue) }
        # 削除は旧のファイル、それ以外は新のファイルの該当セルを開く
        $target = $d.Path; $cell = $d.Cell
        if ($d.Change -eq '削除') {
            $target = $d.OldPath
            if ($d.OldLocation) { $cell = $d.OldLocation.Substring($d.OldLocation.LastIndexOf('!') + 1) }
        }
        if ($target) {
            $loc = ''
            $linkSheet = $d.Sheet
            if ($d.Change -eq '削除') { $linkSheet = $d.OldSheet }
            if ($linkSheet -and $d.Kind -ne 'シート' -and $cell) { $loc = "'{0}'!{1}" -f $linkSheet.Replace("'", "''"), $cell }
            $row.Link = @{ Column = 2; Path = $target; Location = $loc }
        }
        $rows.Add($row)
    }
    $fileRows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $summary) {
        $row = @{ Values = @($s.File, $s.State, $s.Count, $s.Old, $s.New) }
        $t = $s.New; if (-not $t) { $t = $s.Old }
        if ($t) { $row.Link = @{ Column = 0; Path = $t; Location = '' } }
        $fileRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{
                Name = '差分一覧'
                Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = 'シート'; Width = 16 }, @{ Header = '場所'; Width = 20 }, @{ Header = '旧の場所'; Width = 16 },
                    @{ Header = '種類'; Width = 9 }, @{ Header = '変更'; Width = 7 }, @{ Header = '変更前'; Width = 45 }, @{ Header = '変更後'; Width = 45 })
                Rows = $rows
            },
            @{
                Name = 'ファイル別'
                Columns = @(@{ Header = 'ファイル'; Width = 45 }, @{ Header = '結果'; Width = 10 }, @{ Header = '差分'; Width = 8 }, @{ Header = '旧'; Width = 60 }, @{ Header = '新'; Width = 60 })
                Rows = $fileRows
            })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

if ($OutHistory) {
    # 改訂履歴の下書き（New-DesignDocument.ps1 -Bump -CompareWith もこれを使う）
    $encoding = 'UTF8'
    if ($PSVersionTable.PSVersion.Major -ge 6) { $encoding = 'utf8BOM' }
    try {
        $script:HistoryDrafts | Select-Object @{ n = 'ファイル'; e = { $_.File } }, @{ n = '改訂箇所'; e = { $_.Section } }, @{ n = '改訂内容'; e = { $_.Description } }, @{ n = '場所'; e = { $_.Location } } |
            Export-Csv -LiteralPath $OutHistory -NoTypeInformation -Encoding $encoding
    }
    catch {
        Write-Warning ("改訂履歴の下書きを保存できませんでした: {0} : {1}" -f $OutHistory, $_.Exception.Message)
        $OutHistory = $null
    }
}

Write-Host ''
$states = $summary | Group-Object State | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
Write-Host ("{0} ファイル: {1} / 差分 {2} 件" -f $summary.Count, ($states -join ' / '), $script:Diffs.Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
if ($markedCount -gt 0) { Write-Host ("差分の印を付けた写しを保存しました: {0}（{1} ファイル）" -f $OutMarkedDir, $markedCount) }
if ($OutHistory) {
    Write-Host ("改訂履歴の下書きを保存しました: {0}（{1} 件）" -f $OutHistory, $script:HistoryDrafts.Count)
    foreach ($h in $script:HistoryDrafts) { Write-Host ('  {0}  {1}  {2}' -f $h.File, $h.Section, $h.Description) }
}

# .SYNOPSIS
#     2 つの Excel ファイル（またはフォルダ）を比べ、セル単位の差分を一覧にします（Excel 不要）。
#
# .DESCRIPTION
#     値・数式・コメント・図形の文字・シートの追加削除を比べます。
#     行や列を挿入・削除した場合でも、それより下（右）のセルがすべて「変更」にならないよう、
#     同じ内容の行・列どうしを対応付けてから比べます（-NoRowAlign で無効にできます）。
#     内容が同じまま場所が変わった行は「移動」、中身がよく似たシートの名前が変わった場合は「名前変更」として出力します。
#     行の挿入・削除で No などの連番が振り直された場合は、番号のセルの変更を「連番の振り直し」1 件にまとめます。
#     長い文字列の変更は、変わった部分だけを「…受注を[-登録-]{+更新+}する…」の形で示します。
#     -CompareFormat を付けると、文字の書式（太字・斜体・下線・取り消し線・文字色）と塗りつぶし、文字列の一部だけの書式も比べます。
#
#     フォルダどうしを指定すると、同じ相対パスのファイルを比べ、片方にしかないファイルは追加・削除として出力します。
#     版数だけが違う名前のファイル（_v1.0 と _v1.1、第1.0版 と 第1.1版 など）は、1 対 1 なら同じファイルとして比べます。
#     結果は 1 差分 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#     -OutMarkedDir を付けると、新のブックの写しに差分の印（セルの色と、選んだときに出る変更前の値）を付けて保存します。
#     -CheckHistory を付けると、改訂履歴に書かれていない変更を、改訂履歴に書く文の下書きとともに知らせます。
#     -OutHistory を付けると、すべての変更の改訂履歴の下書きを CSV に保存します。
#
# .PARAMETER Old
#     比較元（変更前）のファイルまたはフォルダ。
#
# .PARAMETER New
#     比較先（変更後）のファイルまたはフォルダ。
#
# .PARAMETER KeyColumn
#     行をキーの値で対応付ける表の、キー列の見出し（複数指定可）。例: -KeyColumn '物理名', '項目ID'
#     見出しと同じ文字のセルを探し、その下でキー列に値がある行を 1 つの表として、キーが同じ行どうしを比べます。
#     並べ替えや途中への挿入があっても正しく対応付けられ、順番が変わった行は「移動」として出力します。
#     キーの値が表の中で重複している場合は、その表は通常の方法（内容）で対応付けます。
#
# .PARAMETER NoRowAlign
#     行・列の対応付けを行わず、同じセル番地どうしを比べます。
#
# .PARAMETER NoColumnAlign
#     列の挿入・削除の対応付けを行いません。
#
# .PARAMETER CompareFormat
#     セルの書式も比べます（.xlsx / .xlsm のみ）。対象は太字・斜体・下線・取り消し線・文字色・塗りつぶしと、
#     文字列の一部だけに付けた書式です（フォントの種類・大きさ・罫線・表示形式は比べません）。
#     同じ変更が隣り合うセルは「A5:F5」のように範囲にまとめます。
#
# .PARAMETER CheckHistory
#     新のブックの改訂履歴のシートに追記・変更した行に、ほかのシートの変更が書かれているかを確かめます。
#     変更した行を表す語（左から 2 つ。例: 論理名と物理名）が改訂履歴に無い変更を、行ごとに「改訂履歴 / 記載なし」として、
#     「T_ORDER 合計金額（TOTAL_AMOUNT）: 桁数を 9 → 11 に変更」のような下書きとともに出力します（全角・半角は区別しません）。
#     改訂履歴に追記が無い場合や、改訂履歴のシートが無い場合も知らせます。
#
# .PARAMETER HistorySheet
#     -CheckHistory・-OutHistory で使う、改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#
# .PARAMETER OutHistory
#     改訂履歴に書く文の下書きを CSV（ファイル・改訂箇所・改訂内容・場所。UTF-8 BOM 付き）に保存します。
#     変更した行（列・シート）ごとに「合計金額（TOTAL_AMOUNT）: 桁数を 9 → 11 に変更」のような文にします。
#     改訂履歴のシートの変更と、版数・更新日・更新者などの欄（左隣の見出しで判断）の変更は含めません。
#     フォルダどうしを比べたとき、新しく足したファイルは「全体 / 新規作成」にします。
#     New-DesignDocument.ps1 -Bump -CompareWith は、この下書きを改訂履歴に書き込みます。
#
# .PARAMETER NoRecurse
#     フォルダを比べるとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は比較できなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（差分一覧・ファイル別。場所をクリックすると該当セルが開きます）。
#     長い文字列の変更は、変更前の削除した部分を赤の取り消し線、変更後の追加した部分を赤の太字で示します。
#
# .PARAMETER OutMarkedDir
#     差分のあった新のブック（.xlsx / .xlsm）の写しを、このフォルダーに同じ名前（フォルダーどうしなら同じ相対パス）で保存し、
#     差分の印を付けます。変更したセルは黄、追加したセル・行・列は緑、値を消したセルは赤、移動した行は青で塗り、
#     セルを選ぶと変更前の値などが出るようにします（データの入力規則の入力時メッセージ。既に入力規則のあるセルには付けません）。
#     末尾に、各セルへのリンク付きの「差分一覧」シートを足します。元のブックは変更しません。
#
# .EXAMPLE
#     .\Compare-Excel.ps1 .\v1.0\テーブル定義書.xlsx .\v1.1\テーブル定義書.xlsx
#
# .EXAMPLE
#     # 版のフォルダどうしを比べて保存する
#     .\Compare-Excel.ps1 D:\納品\v1.0 D:\納品\v1.1 -OutXlsx D:\差分_v1.0-v1.1.xlsx
#
# .EXAMPLE
#     # カラム定義を物理名で、インデックス定義をインデックス名で対応付ける
#     .\Compare-Excel.ps1 .\v1.0\テーブル定義書.xlsx .\v1.1\テーブル定義書.xlsx -KeyColumn '物理名', 'インデックス名'
#
# .EXAMPLE
#     # 書式（赤字・取り消し線・塗りつぶしなど）の変更も比べる
#     .\Compare-Excel.ps1 .\v1.0\画面設計書.xlsx .\v1.1\画面設計書.xlsx -CompareFormat
#
# .EXAMPLE
#     # 改訂履歴の記載漏れを確かめる
#     .\Compare-Excel.ps1 D:\納品\v1.0 D:\納品\v1.1 -CheckHistory -KeyColumn '物理名'
#
# .EXAMPLE
#     # 改訂履歴の下書きを作る
#     .\Compare-Excel.ps1 D:\納品\v1.0 D:\納品\v1.1 -OutHistory D:\改訂履歴_下書き.csv
#
# .EXAMPLE
#     # 新の版の写しに差分の印を付けて、レビュー用に保存する
#     .\Compare-Excel.ps1 D:\納品\v1.0 D:\納品\v1.1 -OutMarkedDir D:\差分レビュー
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     列の挿入・削除は、列ごとの値の集まりが似ている列どうしを対応付けて見つけます（値の少ない列は見つけにくい）。
