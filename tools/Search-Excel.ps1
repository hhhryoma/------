[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Pattern,

    [Parameter(Position = 1)]
    [string[]] $Path = @('.'),

    [switch] $SimpleMatch,
    [switch] $CaseSensitive,
    [switch] $IgnoreWidth,

    [ValidateSet('Value', 'Formula', 'Comment', 'Shape', 'SheetName')]
    [string[]] $Target = @('Value', 'Formula', 'Comment', 'Shape', 'SheetName'),

    [switch] $NoRecurse,
    [switch] $SkipHiddenSheets,

    [string] $SheetNamePattern,
    [string] $SheetNameExclude,

    [switch] $NoExcel,

    [switch] $Open,

    [string] $OutCsv,

    [string] $OutXlsx,

    [string[]] $Extensions = @('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb')
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml

# ---- Excel 定数（COM の列挙体を PowerShell から参照しないで済むよう定義） ----
$xlSheetVisible    = -1
$xlRangeValueDefault = 10
$msoGroup          = 6
$msoAutomationSecurityForceDisable = 3

$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$nfkc      = [System.Text.NormalizationForm]::FormKC

# 組み込みの表示形式のうち日付・時刻のもの（27〜36, 50〜58 は日本語ロケールの和暦・年月日）
$builtinDateFormatIds = @(14..22) + @(27..36) + @(45..47) + @(50..58)

$wantValue     = $Target -contains 'Value'
$wantFormula   = $Target -contains 'Formula'
$wantComment   = $Target -contains 'Comment'
$wantShape     = $Target -contains 'Shape'
$wantSheetName = $Target -contains 'SheetName'

# ---------------------------------------------------------------- 検索条件

$regexOptions = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
if (-not $CaseSensitive) {
    $regexOptions = $regexOptions -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
}

$patternText = $Pattern
if ($SimpleMatch) {
    if ($IgnoreWidth) { $patternText = $patternText.Normalize($nfkc) }
    $patternText = [regex]::Escape($patternText)
}

try {
    $rx = New-Object System.Text.RegularExpressions.Regex($patternText, $regexOptions)
}
catch {
    throw "検索パターンが正規表現として不正です: $Pattern （文字列として探すなら -SimpleMatch を付けてください） 詳細: $($_.Exception.InnerException.Message)"
}

foreach ($p in @($SheetNamePattern, $SheetNameExclude)) {
    if ($p) {
        try { [void][regex]::new($p) }
        catch { throw "シート名の正規表現が不正です: $p 詳細: $($_.Exception.InnerException.Message)" }
    }
}

# 出力先は PowerShell のカレントからの相対パスとして解決する（.NET のカレントとは異なるため）
if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') {
        throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx"
    }
}

if ($Open -and -not (Get-Command Out-GridView -ErrorAction SilentlyContinue)) {
    throw 'この環境では Out-GridView が使えないため、-Open は利用できません（Windows で実行してください）。'
}

$script:XmlSettings = New-Object System.Xml.XmlReaderSettings
$script:XmlSettings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
$script:XmlSettings.XmlResolver   = $null
$script:XmlSettings.IgnoreComments = $true
$script:XmlSettings.CloseInput     = $true

Update-TypeData -TypeName 'SearchExcel.Hit' -DefaultDisplayPropertySet File, Location, Kind, Text -Force

# ---------------------------------------------------------------- ヘルパー

function Test-TextMatch {
    param([string] $Text)

    if ([string]::IsNullOrEmpty($Text)) { return $false }
    if ($IgnoreWidth) { $Text = $Text.Normalize($nfkc) }
    return $rx.IsMatch($Text)
}

function Test-SheetNameMatch {
    #     シート名が -SheetNamePattern / -SheetNameExclude の条件を満たすか判定する。
    #     いずれも .NET 正規表現、大文字小文字は区別しない。
    param([string] $Name)

    if ($SheetNamePattern -and -not [regex]::IsMatch($Name, $SheetNamePattern, 'IgnoreCase')) { return $false }
    if ($SheetNameExclude -and [regex]::IsMatch($Name, $SheetNameExclude, 'IgnoreCase')) { return $false }
    return $true
}

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

function ConvertFrom-OoxmlEscape {
    # OOXML の文字列に含まれる _xHHHH_ 形式のエスケープ（改行の _x000D_ など）を戻す
    param([string] $Text)

    if ($Text.IndexOf('_x') -lt 0) { return $Text }
    return [regex]::Replace($Text, '_x([0-9A-Fa-f]{4})_', {
        param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16)
    })
}

function ConvertTo-OneLine {
    param([string] $Text)

    if ($null -eq $Text) { return '' }
    return $Text.Replace("`r`n", '↵').Replace("`n", '↵').Replace("`r", '↵')
}

function Format-ExcelNumber {
    param([double] $Number)
    return $Number.ToString('G15', $invariant)
}

function Format-ExcelDate {
    #     シリアル値を日付文字列にする。整数部が 0 なら時刻のみ、小数部が 0 なら日付のみ。
    param([double] $Serial, [bool] $Date1904)

    try {
        if ($Date1904) {
            $dt = (New-Object DateTime 1904, 1, 1).AddDays($Serial)
        }
        else {
            # Excel は 1900/2/29 を実在扱いするため、それより前は 1 日ずれる
            if ($Serial -lt 61 -and $Serial -ge 1) { $Serial += 1 }
            $dt = [DateTime]::FromOADate($Serial)
        }
        # 浮動小数点の誤差で 12:59:59.999 のようになるのを防ぐため秒に丸める
        $dt = New-Object DateTime ([long]([math]::Round($dt.Ticks / 1e7)) * 10000000L)

        if ([math]::Floor($Serial) -eq 0 -and -not $Date1904) { return $dt.ToString('HH:mm:ss', $invariant) }
        if ($dt.TimeOfDay.Ticks -eq 0) { return $dt.ToString('yyyy/MM/dd', $invariant) }
        return $dt.ToString('yyyy/MM/dd HH:mm:ss', $invariant)
    }
    catch {
        return (Format-ExcelNumber $Serial)
    }
}

function Test-DateFormatCode {
    #     ユーザー定義の表示形式が日付・時刻か判定する。
    #     文字列リテラル（"..."）・エスケープ（\x）・[色] や [$-411] などの [...] を除いてから
    #     y/m/d/h/s が残っていれば日付とみなす。経過時間 [h]:mm は日付扱いしない。
    param([string] $Code)

    if ([string]::IsNullOrEmpty($Code)) { return $false }
    if ($Code -match '\[(h+|m+|s+)\]') { return $false }

    $s = $Code.Split(';')[0]
    $s = [regex]::Replace($s, '"[^"]*"', '')
    $s = [regex]::Replace($s, '\\.', '')
    $s = [regex]::Replace($s, '[_*].', '')
    $s = [regex]::Replace($s, '\[[^\]]*\]', '')
    $s = [regex]::Replace($s, 'General', '', 'IgnoreCase')
    $s = [regex]::Replace($s, '[Ee][+-]', '')
    return ($s -match '[yYmMdDhHsS]')
}

# ---------------------------------------------------------------- xlsx（Open XML）の読み取り

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

function Read-RichText {
    #     XmlReader が <si> / <is> / <text> の開始位置にあるとき、その中の文字列を連結して返す。
    #     ふりがな（<rPh>）は本文ではないので除外する。終了時、リーダーは終了タグの上にある。
    param([System.Xml.XmlReader] $Reader)

    if ($Reader.IsEmptyElement) { return '' }

    $sb    = New-Object System.Text.StringBuilder
    $depth = $Reader.Depth
    $inPhonetic = $false
    $advance = $true
    while ($true) {
        if ($advance -and -not $Reader.Read()) { break }
        $advance = $true

        if ($Reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
            if ($Reader.LocalName -eq 'rPh' -and -not $Reader.IsEmptyElement) { $inPhonetic = $true }
            elseif ($Reader.LocalName -eq 't') {
                $t = $Reader.ReadElementContentAsString()
                if (-not $inPhonetic) { [void]$sb.Append($t) }
                $advance = $false
            }
        }
        elseif ($Reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
            if ($Reader.Depth -eq $depth) { break }
            if ($Reader.LocalName -eq 'rPh') { $inPhonetic = $false }
        }
    }
    return (ConvertFrom-OoxmlEscape $sb.ToString())
}

function Read-SharedStrings {
    #     共有文字列テーブルを読み、文字列の配列と「検索に一致するか」の配列を返す。
    #     同じ文字列を参照するセルが多いので、一致判定はここで 1 回だけ行う。
    param($Package, [string] $PartName)

    $strings = New-Object System.Collections.Generic.List[string]
    $hits    = New-Object System.Collections.Generic.List[bool]

    if ($PartName) {
        $reader = New-PartXmlReader -Package $Package -PartName $PartName
        if ($null -ne $reader) {
            # 文字列の数が多いブックもあるため、Read-RichText を呼ばずにここで直接読む
            $sb = New-Object System.Text.StringBuilder
            $normalize  = [bool]$IgnoreWidth
            $inPhonetic = $false
            $advance    = $true
            try {
                while ($true) {
                    if ($advance -and -not $reader.Read()) { break }
                    $advance = $true

                    $nodeType = $reader.NodeType
                    if ($nodeType -eq [System.Xml.XmlNodeType]::Element) {
                        $name = $reader.LocalName
                        if ($name -eq 't') {
                            $t = $reader.ReadElementContentAsString()
                            if (-not $inPhonetic) { [void]$sb.Append($t) }
                            $advance = $false
                        }
                        elseif ($name -eq 'si') {
                            [void]$sb.Clear()
                            if ($reader.IsEmptyElement) { $strings.Add(''); $hits.Add($false) }
                        }
                        elseif ($name -eq 'rPh' -and -not $reader.IsEmptyElement) { $inPhonetic = $true }
                    }
                    elseif ($nodeType -eq [System.Xml.XmlNodeType]::EndElement) {
                        $name = $reader.LocalName
                        if ($name -eq 'si') {
                            $s = $sb.ToString()
                            if ($s.IndexOf('_x') -ge 0) { $s = ConvertFrom-OoxmlEscape $s }
                            $strings.Add($s)
                            $probe = $s
                            if ($normalize) { $probe = $probe.Normalize($nfkc) }
                            $hits.Add($rx.IsMatch($probe))
                        }
                        elseif ($name -eq 'rPh') { $inPhonetic = $false }
                    }
                }
            }
            finally { $reader.Dispose() }
        }
    }

    return [pscustomobject]@{ Strings = $strings; Hits = $hits }
}

function Read-DateStyles {
    #     styles.xml の cellXfs を読み、スタイル番号ごとに日付書式かどうかを返す
    param($Package, [string] $PartName)

    $result = New-Object System.Collections.Generic.List[bool]
    if (-not $PartName) { return ,$result }

    $doc = Read-PartXml -Package $Package -PartName $PartName
    if ($null -eq $doc) { return ,$result }

    $customDate = @{}
    foreach ($nf in $doc.SelectNodes("//*[local-name()='numFmts']/*[local-name()='numFmt']")) {
        $customDate[[int]$nf.GetAttribute('numFmtId')] = Test-DateFormatCode $nf.GetAttribute('formatCode')
    }

    foreach ($xf in $doc.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
        $id = 0
        [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$id)
        if ($customDate.ContainsKey($id)) { $result.Add($customDate[$id]) }
        else { $result.Add($builtinDateFormatIds -contains $id) }
    }
    return ,$result
}

function New-Hit {
    param([string] $Kind, [string] $Sheet, [string] $Cell, [string] $Text, [string] $Name)

    $location = $Sheet
    if ($Cell) { $location = '{0}!{1}' -f $Sheet, $Cell }

    $hit = [pscustomobject]@{
        File     = $script:CurrentFile
        Location = $location
        Sheet    = $Sheet
        Cell     = $Cell
        Kind     = $Kind
        Text     = (ConvertTo-OneLine $Text)
        Name     = $Name
        Path     = $script:CurrentPath
    }
    $hit.PSObject.TypeNames.Insert(0, 'SearchExcel.Hit')
    return $hit
}

function Search-WorksheetCells {
    #     ワークシートの XML を先頭から順に読み、セルの値・数式を検索する。
    #     巨大なシートでもメモリに載せないよう XmlReader で逐次処理する。
    param($Package, [string] $PartName, [string] $SheetName, $Shared, $DateStyles, [bool] $Date1904)

    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return }

    $sstStrings = $Shared.Strings
    $sstHits    = $Shared.Hits
    $dateCount  = $DateStyles.Count
    $normalize  = [bool]$IgnoreWidth
    $numStyle   = [System.Globalization.NumberStyles]::Float
    $num        = 0.0
    $idx        = 0

    # セルごとの処理は回数が多いため、関数呼び出しを避けて直接書いている

    try {
        $row = 0; $col = 0
        $lastRef = $null
        $advance = $true
        $inCell  = $false
        $ref = $null; $type = $null; $style = 0
        $value = $null; $formula = $null; $inline = $null

        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true

            $nodeType = $reader.NodeType
            if ($nodeType -eq [System.Xml.XmlNodeType]::Element) {
                $name = $reader.LocalName

                if ($name -eq 'c') {
                    $ref = $reader.GetAttribute('r')
                    if ($ref) {
                        $lastRef = $ref
                    }
                    else {
                        # r 属性を省略するツールで作られたファイル向けに位置を数える
                        if ($lastRef) { $col = Get-ColumnIndex $lastRef; $lastRef = $null }
                        $col++
                        $ref = (ConvertTo-ColumnName $col) + $row
                    }
                    if ($reader.IsEmptyElement) { continue }

                    $type  = $reader.GetAttribute('t')
                    $s     = $reader.GetAttribute('s')
                    $style = 0
                    if ($s) { [void][int]::TryParse($s, [ref]$style) }
                    $value = $null; $formula = $null; $inline = $null
                    $inCell = $true
                }
                elseif ($name -eq 'row') {
                    $r = $reader.GetAttribute('r')
                    if ($r) { $row = [int]$r } else { $row++ }
                    $col = 0; $lastRef = $null
                }
                elseif ($inCell) {
                    if ($name -eq 'v') {
                        $value = $reader.ReadElementContentAsString()
                        $advance = $false
                    }
                    elseif ($name -eq 'f') {
                        $formula = $reader.ReadElementContentAsString()
                        $advance = $false
                    }
                    elseif ($name -eq 'is') {
                        $inline = Read-RichText -Reader $reader
                    }
                }
            }
            elseif ($nodeType -eq [System.Xml.XmlNodeType]::EndElement -and $inCell -and $reader.LocalName -eq 'c') {
                $inCell = $false

                if ($wantFormula -and $formula) {
                    $f = '=' + $formula
                    $probe = $f
                    if ($normalize) { $probe = $probe.Normalize($nfkc) }
                    if ($rx.IsMatch($probe)) { New-Hit -Kind '数式' -Sheet $SheetName -Cell $ref -Text $f }
                }

                if ($wantValue) {
                    $text = $null
                    if ($type -eq 's') {
                        if ([int]::TryParse($value, [ref]$idx) -and $idx -ge 0 -and $idx -lt $sstStrings.Count) {
                            if ($sstHits[$idx]) { New-Hit -Kind '値' -Sheet $SheetName -Cell $ref -Text $sstStrings[$idx] }
                        }
                        continue
                    }
                    elseif ($type -eq 'inlineStr') { $text = $inline }
                    elseif ($type -eq 'b') { if ($value -eq '1') { $text = 'TRUE' } elseif ($null -ne $value) { $text = 'FALSE' } }
                    elseif ($type -eq 'str' -or $type -eq 'e' -or $type -eq 'd') { $text = ConvertFrom-OoxmlEscape $value }
                    elseif ($value) {
                        if ([double]::TryParse($value, $numStyle, $invariant, [ref]$num)) {
                            if ($style -lt $dateCount -and $DateStyles[$style]) { $text = Format-ExcelDate -Serial $num -Date1904 $Date1904 }
                            else { $text = $num.ToString('G15', $invariant) }
                        }
                        else { $text = $value }
                    }

                    if ($text) {
                        $probe = $text
                        if ($normalize) { $probe = $probe.Normalize($nfkc) }
                        if ($rx.IsMatch($probe)) { New-Hit -Kind '値' -Sheet $SheetName -Cell $ref -Text $text }
                    }
                }
            }
        }
    }
    finally {
        $reader.Dispose()
    }
}

function Search-SheetComments {
    #     メモ（従来のコメント）とスレッド形式のコメントを検索する。
    #     スレッド形式のコメントには互換用のメモも付くため、同じセルのメモは重複として除く。
    param($Package, $Rels, [string] $SheetName)

    $threadedCells = @{}

    foreach ($rel in @($Rels | Where-Object { $_.Kind -eq 'threadedComment' })) {
        $doc = Read-PartXml -Package $Package -PartName $rel.Target
        if ($null -eq $doc) { continue }
        foreach ($tc in $doc.SelectNodes("//*[local-name()='threadedComment']")) {
            $cell = $tc.GetAttribute('ref')
            $threadedCells[$cell] = $true
            $textNode = $tc.SelectSingleNode("*[local-name()='text']")
            if ($null -eq $textNode) { continue }
            $text = ConvertFrom-OoxmlEscape $textNode.InnerText
            if (Test-TextMatch $text) { New-Hit -Kind 'コメント' -Sheet $SheetName -Cell $cell -Text $text }
        }
    }

    foreach ($rel in @($Rels | Where-Object { $_.Kind -eq 'comments' })) {
        $doc = Read-PartXml -Package $Package -PartName $rel.Target
        if ($null -eq $doc) { continue }

        $authors = @($doc.SelectNodes("//*[local-name()='authors']/*[local-name()='author']") | ForEach-Object { $_.InnerText })
        foreach ($c in $doc.SelectNodes("//*[local-name()='commentList']/*[local-name()='comment']")) {
            $cell = $c.GetAttribute('ref')
            if ($threadedCells.ContainsKey($cell)) { continue }

            $sb = New-Object System.Text.StringBuilder
            foreach ($t in $c.SelectNodes(".//*[local-name()='t'][not(ancestor::*[local-name()='rPh'])]")) {
                [void]$sb.Append($t.InnerText)
            }
            $text = ConvertFrom-OoxmlEscape $sb.ToString()

            $author = $null
            $aid = 0
            if ([int]::TryParse($c.GetAttribute('authorId'), [ref]$aid) -and $aid -lt $authors.Count) { $author = $authors[$aid] }

            if (Test-TextMatch $text) { New-Hit -Kind 'コメント' -Sheet $SheetName -Cell $cell -Text $text -Name $author }
        }
    }
}

function Search-SheetShapes {
    #     描画パーツ（drawingN.xml）の図形・テキストボックスの文字を検索する。
    #     セル番地は図形の左上が乗っているセル。
    param($Package, $Rels, [string] $SheetName)

    foreach ($rel in @($Rels | Where-Object { $_.Kind -eq 'drawing' })) {
        $doc = Read-PartXml -Package $Package -PartName $rel.Target
        if ($null -eq $doc) { continue }

        foreach ($sp in $doc.SelectNodes("//*[local-name()='sp']")) {
            # 互換用の代替表示（mc:Fallback）は同じ図形の重複なので除く
            if ($null -ne $sp.SelectSingleNode("ancestor::*[local-name()='Fallback']")) { continue }

            $paragraphs = New-Object System.Collections.Generic.List[string]
            foreach ($p in $sp.SelectNodes("*[local-name()='txBody']/*[local-name()='p']")) {
                $sb = New-Object System.Text.StringBuilder
                foreach ($t in $p.SelectNodes(".//*[local-name()='t']")) { [void]$sb.Append($t.InnerText) }
                $paragraphs.Add($sb.ToString())
            }
            $text = ($paragraphs -join "`n").TrimEnd("`n")
            if (-not (Test-TextMatch $text)) { continue }

            $shapeName = $null
            $nv = $sp.SelectSingleNode("*[local-name()='nvSpPr']/*[local-name()='cNvPr']")
            if ($null -ne $nv) { $shapeName = $nv.GetAttribute('name') }

            $cell = $null
            $from = $sp.SelectSingleNode("ancestor::*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor'][1]/*[local-name()='from']")
            if ($null -ne $from) {
                $c = $from.SelectSingleNode("*[local-name()='col']")
                $r = $from.SelectSingleNode("*[local-name()='row']")
                if ($null -ne $c -and $null -ne $r) {
                    $cell = (ConvertTo-ColumnName ([int]$c.InnerText + 1)) + ([int]$r.InnerText + 1)
                }
            }

            New-Hit -Kind '図形' -Sheet $SheetName -Cell $cell -Text $text -Name $shapeName
        }
    }
}

function Search-OpenXmlWorkbook {
    param([string] $FilePath)

    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        # ブック本体の場所はルートのリレーションから求める（通常は xl/workbook.xml）
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) {
            if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break }
        }

        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        if ($null -eq $wbDoc) {
            throw 'Excel ブックではありません（workbook.xml が見つかりません）'
        }

        $date1904 = $false
        $wbPr = $wbDoc.SelectSingleNode("//*[local-name()='workbookPr']")
        if ($null -ne $wbPr) { $date1904 = @('1', 'true') -contains $wbPr.GetAttribute('date1904') }

        $wbRels = Get-PartRelationships -Package $pkg -PartName $wbPart
        $relById = @{}
        $sstPart = $null; $stylesPart = $null
        foreach ($rel in $wbRels) {
            $relById[$rel.Id] = $rel
            if ($rel.Kind -eq 'sharedStrings') { $sstPart = $rel.Target }
            elseif ($rel.Kind -eq 'styles') { $stylesPart = $rel.Target }
        }

        $shared = $null
        $dateStyles = $null
        if ($wantValue) {
            $shared     = Read-SharedStrings -Package $pkg -PartName $sstPart
            $dateStyles = Read-DateStyles -Package $pkg -PartName $stylesPart
        }
        else {
            $shared = [pscustomobject]@{ Strings = @(); Hits = @() }
            $dateStyles = @()
        }

        foreach ($sheet in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $sheetName = $sheet.GetAttribute('name')
            $state     = $sheet.GetAttribute('state')

            if ($SkipHiddenSheets -and ($state -eq 'hidden' -or $state -eq 'veryHidden')) { continue }
            if (-not (Test-SheetNameMatch $sheetName)) { continue }

            if ($wantSheetName -and (Test-TextMatch $sheetName)) {
                New-Hit -Kind 'シート名' -Sheet $sheetName -Cell $null -Text $sheetName
            }

            $rel = $relById[(Get-RelationshipIdAttribute $sheet)]
            if ($null -eq $rel) { continue }
            # グラフシートなどセルを持たないシートは中身を見ない
            if ($rel.Kind -ne 'worksheet' -and $rel.Kind -ne 'macrosheet') { continue }

            if ($wantValue -or $wantFormula) {
                Search-WorksheetCells -Package $pkg -PartName $rel.Target -SheetName $sheetName `
                    -Shared $shared -DateStyles $dateStyles -Date1904 $date1904
            }

            if ($wantComment -or $wantShape) {
                $sheetRels = Get-PartRelationships -Package $pkg -PartName $rel.Target
                if ($wantComment) { Search-SheetComments -Package $pkg -Rels $sheetRels -SheetName $sheetName }
                if ($wantShape)   { Search-SheetShapes -Package $pkg -Rels $sheetRels -SheetName $sheetName }
            }
        }
    }
    finally {
        $pkg.Zip.Dispose()
    }
}

# ---------------------------------------------------------------- xls / xlsb（Excel COM）の読み取り

$script:Excel      = $null
$script:ExcelPids  = @()
$script:ExcelError = $null

function Get-ExcelApplication {
    #     .xls / .xlsb を最初に見つけた時点で Excel を起動する（xlsx だけなら起動しない）
    if ($null -ne $script:Excel) { return $script:Excel }
    if ($script:ExcelError) { throw $script:ExcelError }

    $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    try {
        $xl = New-Object -ComObject Excel.Application
    }
    catch {
        $script:ExcelError = "Excel を起動できないため .xls / .xlsb は検索できません。 詳細: $($_.Exception.Message)"
        throw $script:ExcelError
    }
    $after = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $script:ExcelPids = @($after | Where-Object { $before -notcontains $_ })

    $xl.Visible          = $false
    $xl.DisplayAlerts    = $false
    $xl.AskToUpdateLinks = $false
    $xl.EnableEvents     = $false
    $xl.ScreenUpdating   = $false
    try { $xl.AutomationSecurity = $msoAutomationSecurityForceDisable }
    catch { Write-Warning ("AutomationSecurity を設定できませんでした（マクロが無効化されません）: {0}" -f $_.Exception.Message) }

    $script:Excel = $xl
    return $xl
}

function Close-ExcelApplication {
    if ($null -ne $script:Excel) {
        try { $script:Excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:Excel) } catch { }
        $script:Excel = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }

    # Quit しても残ることがあるため、自分が起動したプロセスだけ確実に落とす
    foreach ($procId in $script:ExcelPids) {
        if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
            Start-Sleep -Seconds 2
            if (Get-Process -Id $procId -ErrorAction SilentlyContinue) {
                Write-Verbose "残存した Excel プロセスを終了します (PID: $procId)"
                Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
            }
        }
    }
    $script:ExcelPids = @()
}

function Format-ComValue {
    param($Value)

    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTime]) {
        if ($Value.Date -eq (New-Object DateTime 1899, 12, 30)) { return $Value.ToString('HH:mm:ss', $invariant) }
        if ($Value.TimeOfDay.Ticks -eq 0) { return $Value.ToString('yyyy/MM/dd', $invariant) }
        return $Value.ToString('yyyy/MM/dd HH:mm:ss', $invariant)
    }
    if ($Value -is [double] -or $Value -is [decimal]) { return Format-ExcelNumber ([double]$Value) }
    if ($Value -is [bool]) { if ($Value) { return 'TRUE' } else { return 'FALSE' } }
    if ($Value -is [int]) {
        # COM はセルのエラー値（#N/A など）を整数で返す
        switch ($Value) {
            -2146826281 { return '#DIV/0!' }
            -2146826246 { return '#N/A' }
            -2146826259 { return '#NAME?' }
            -2146826288 { return '#NULL!' }
            -2146826252 { return '#NUM!' }
            -2146826265 { return '#REF!' }
            -2146826273 { return '#VALUE!' }
        }
    }
    return [string]$Value
}

function Search-ComShape {
    param($Shape, [string] $SheetName)

    try {
        if ($Shape.Type -eq $msoGroup) {
            foreach ($child in $Shape.GroupItems) { Search-ComShape -Shape $child -SheetName $SheetName }
            return
        }
        if (-not $Shape.TextFrame2.HasText) { return }
        $text = [string]$Shape.TextFrame2.TextRange.Text
    }
    catch { return }

    $text = $text.Replace("`r", "`n")
    if (-not (Test-TextMatch $text)) { return }

    $cell = $null
    try { $cell = $Shape.TopLeftCell.Address($false, $false) } catch { }
    New-Hit -Kind '図形' -Sheet $SheetName -Cell $cell -Text $text -Name $Shape.Name
}

function Search-ComWorkbook {
    param([string] $FilePath)

    if ($FilePath.Length -gt 218) {
        throw ("パスが長すぎます（{0} 文字）。Excel はフルパス 218 文字までしか開けません。" -f $FilePath.Length)
    }

    $xl = Get-ExcelApplication
    $wb = $null
    try {
        # 読み取り専用・リンク更新なし。ダミーのパスワードを渡し、保護ファイルで入力待ちにならないようにする
        try {
            $wb = $xl.Workbooks.Open($FilePath, 0, $true, [Type]::Missing, 'x', 'x', $true)
        }
        catch {
            if ($_.Exception.Message -match 'パスワード|password') { throw }
            Write-Verbose ("省略可能引数つきの Open に失敗したため簡易形式で開き直します: {0}" -f $_.Exception.Message)
            $wb = $xl.Workbooks.Open($FilePath, 0, $true)
        }

        foreach ($ws in $wb.Worksheets) {
            $sheetName = [string]$ws.Name
            if ($SkipHiddenSheets -and $ws.Visible -ne $xlSheetVisible) { continue }
            if (-not (Test-SheetNameMatch $sheetName)) { continue }

            if ($wantSheetName -and (Test-TextMatch $sheetName)) {
                New-Hit -Kind 'シート名' -Sheet $sheetName -Cell $null -Text $sheetName
            }

            if ($wantValue -or $wantFormula) {
                $used = $ws.UsedRange
                $top  = [int]$used.Row
                $left = [int]$used.Column
                $rows = [int]$used.Rows.Count
                $cols = [int]$used.Columns.Count

                # 1 セルずつ COM を呼ぶと遅いので、範囲全体を配列で受け取る
                $values = $null; $formulas = $null
                if ($wantValue) {
                    try { $values = $used.Value($xlRangeValueDefault) } catch { $values = $used.Value2 }
                }
                if ($wantFormula) { $formulas = $used.Formula }

                for ($r = 1; $r -le $rows; $r++) {
                    for ($c = 1; $c -le $cols; $c++) {
                        $addr = $null

                        if ($wantFormula) {
                            if ($rows -eq 1 -and $cols -eq 1) { $f = [string]$formulas } else { $f = [string]$formulas[$r, $c] }
                            if ($f.StartsWith('=') -and (Test-TextMatch $f)) {
                                $addr = (ConvertTo-ColumnName ($left + $c - 1)) + ($top + $r - 1)
                                New-Hit -Kind '数式' -Sheet $sheetName -Cell $addr -Text $f
                            }
                        }

                        if ($wantValue) {
                            if ($rows -eq 1 -and $cols -eq 1) { $v = $values } else { $v = $values[$r, $c] }
                            $text = Format-ComValue $v
                            if ($text -and (Test-TextMatch $text)) {
                                if (-not $addr) { $addr = (ConvertTo-ColumnName ($left + $c - 1)) + ($top + $r - 1) }
                                New-Hit -Kind '値' -Sheet $sheetName -Cell $addr -Text $text
                            }
                        }
                    }
                }
            }

            if ($wantComment) {
                foreach ($cm in $ws.Comments) {
                    try {
                        $text = ([string]$cm.Text()).Replace("`r", "`n")
                        if (Test-TextMatch $text) {
                            New-Hit -Kind 'コメント' -Sheet $sheetName -Cell $cm.Parent.Address($false, $false) -Text $text -Name $cm.Author
                        }
                    }
                    catch { }
                }
            }

            if ($wantShape) {
                foreach ($sh in $ws.Shapes) { Search-ComShape -Shape $sh -SheetName $sheetName }
            }
        }
    }
    finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch { }
            try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
        }
    }
}

# ---------------------------------------------------------------- 検索結果を xlsx に保存

function ConvertTo-XmlText {
    # XML の特殊文字をエスケープし、XML に書けない制御文字を除く
    param([string] $Text)

    if ($null -eq $Text) { return '' }
    $Text = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]', '')
    return [System.Security.SecurityElement]::Escape($Text)
}

function Export-HitsToXlsx {
    #     検索結果を xlsx に書き出す。Excel は使わず、Open XML を直接組み立てる。
    #     「場所」列は、クリックすると該当ファイルのそのセルが開くハイパーリンクにする。
    param($Hits, [string] $FilePath)

    $maxLinks   = 65530   # Excel の 1 シートあたりのハイパーリンク上限
    $maxCellLen = 32767   # Excel の 1 セルあたりの文字数上限

    $headers = @('ファイル', '場所', 'シート', 'セル', '種類', '内容', '名前', 'フルパス')
    $widths  = @(40, 24, 18, 8, 9, 80, 18, 60)

    $rowCount = $Hits.Count + 1
    $lastCol  = ConvertTo-ColumnName $headers.Count

    $sheet = New-Object System.Text.StringBuilder
    $rels  = New-Object System.Text.StringBuilder
    $links = New-Object System.Text.StringBuilder

    [void]$sheet.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sheet.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
    [void]$sheet.Append(('<dimension ref="A1:{0}{1}"/>' -f $lastCol, $rowCount))
    [void]$sheet.Append('<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>')
    [void]$sheet.Append('<sheetFormatPr defaultRowHeight="18"/><cols>')
    for ($i = 0; $i -lt $widths.Count; $i++) {
        [void]$sheet.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), $widths[$i]))
    }
    [void]$sheet.Append('</cols><sheetData>')

    # 見出し行（スタイル 1: 太字・背景色）
    [void]$sheet.Append('<row r="1">')
    for ($i = 0; $i -lt $headers.Count; $i++) {
        [void]$sheet.Append(('<c r="{0}1" s="1" t="inlineStr"><is><t>{1}</t></is></c>' -f (ConvertTo-ColumnName ($i + 1)), $headers[$i]))
    }
    [void]$sheet.Append('</row>')

    [void]$rels.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$rels.Append('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')

    $r = 1
    $linkCount = 0
    foreach ($h in $Hits) {
        $r++

        # リンク先のセル。シート名のヒットや位置不明の図形は A1 にする
        $cell = $h.Cell
        if (-not $cell) { $cell = 'A1' }
        $location = "'{0}'!{1}" -f $h.Sheet.Replace("'", "''"), $cell

        $values = @($h.File, $h.Location, $h.Sheet, $h.Cell, $h.Kind, $h.Text, $h.Name, $h.Path)
        [void]$sheet.Append(('<row r="{0}">' -f $r))
        for ($i = 0; $i -lt $values.Count; $i++) {
            $v = [string]$values[$i]
            if ($v.Length -eq 0) { continue }
            if ($v.Length -gt $maxCellLen) { $v = $v.Substring(0, $maxCellLen) }
            $style = ''
            if ($i -eq 1 -and $linkCount -lt $maxLinks) { $style = ' s="2"' }
            [void]$sheet.Append(('<c r="{0}{1}"{2} t="inlineStr"><is><t xml:space="preserve">{3}</t></is></c>' -f (ConvertTo-ColumnName ($i + 1)), $r, $style, (ConvertTo-XmlText $v)))
        }
        [void]$sheet.Append('</row>')

        if ($linkCount -lt $maxLinks) {
            $linkCount++
            # Excel が他ブックへのリンクを保存するときと同じ file:/// 形式にする
            $target = $h.Path
            if ($target.StartsWith('/')) { $target = 'file://' + $target } else { $target = 'file:///' + $target }
            [void]$rels.Append(('<Relationship Id="rIdL{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="{1}" TargetMode="External"/>' -f $linkCount, (ConvertTo-XmlText $target)))
            [void]$links.Append(('<hyperlink ref="B{0}" r:id="rIdL{1}" location="{2}" tooltip="{3}"/>' -f $r, $linkCount, (ConvertTo-XmlText $location), (ConvertTo-XmlText ('{0} を開く' -f $h.File))))
        }
    }

    [void]$sheet.Append('</sheetData>')
    [void]$sheet.Append(('<autoFilter ref="A1:{0}{1}"/>' -f $lastCol, $rowCount))
    if ($linkCount -gt 0) { [void]$sheet.Append('<hyperlinks>').Append($links.ToString()).Append('</hyperlinks>') }
    [void]$sheet.Append('</worksheet>')
    [void]$rels.Append('</Relationships>')

    if ($Hits.Count -gt $maxLinks) {
        Write-Warning ("ヒットが {0} 件を超えたため、{1} 行目以降にはリンクを付けていません（Excel の上限）。" -f $maxLinks, $maxLinks + 2)
    }

    $font = '<sz val="11"/><name val="游ゴシック"/><family val="3"/><charset val="128"/>'
    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="検索結果" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">検索結果!$A$1:${0}${1}</definedName></definedNames></workbook>' -f $lastCol, $rowCount)
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/styles.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="3"><font>{0}</font><font><b/>{0}</font><font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>' -f $font)
        'xl/worksheets/sheet1.xml' = $sheet.ToString()
        'xl/worksheets/_rels/sheet1.xml.rels' = $rels.ToString()
    }
    if ($linkCount -eq 0) { $parts.Remove('xl/worksheets/_rels/sheet1.xml.rels') }

    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $stream = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create, $true)
        try {
            foreach ($name in $parts.Keys) {
                $entry  = $zip.CreateEntry($name, [System.IO.Compression.CompressionLevel]::Optimal)
                $writer = New-Object System.IO.StreamWriter($entry.Open(), $utf8)
                try { $writer.Write($parts[$name]) } finally { $writer.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $stream.Dispose() }
}

# ---------------------------------------------------------------- 検索結果を Excel で開く

$msoAutomationSecurityByUI = 2

function Get-RunningExcel {
    #     起動中の Excel を取得する。無ければ $null。
    #     Windows PowerShell 5.1 は Marshal.GetActiveObject を使う。
    #     PowerShell 7（.NET Core）には同メソッドが無いため oleaut32 を直接呼ぶ。
    try { return [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }

    try {
        if (-not ('SearchExcel.Ole' -as [type])) {
            Add-Type -Namespace SearchExcel -Name Ole -MemberDefinition @"
[DllImport("ole32.dll", PreserveSig = false)]
public static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);
[DllImport("oleaut32.dll", PreserveSig = false)]
[return: MarshalAs(UnmanagedType.IUnknown)]
public static extern object GetActiveObject(ref Guid clsid, IntPtr reserved);
"@
        }
        $clsid = [guid]::Empty
        [SearchExcel.Ole]::CLSIDFromProgID('Excel.Application', [ref]$clsid)
        return [SearchExcel.Ole]::GetActiveObject([ref]$clsid, [IntPtr]::Zero)
    }
    catch { return $null }
}

function Open-HitInExcel {
    #     ヒットしたファイルを（読み取り専用ではなく）通常どおり開き、該当セルを選択する。
    #     起動中の Excel があればそれを使い、既に開いているブックはそのまま切り替える。
    param($Hit)

    $xl = Get-RunningExcel
    if ($null -eq $xl) {
        $xl = New-Object -ComObject Excel.Application
    }
    $xl.Visible = $true
    $xl.UserControl = $true   # スクリプト終了後も Excel を閉じないようにする

    $wb = $null
    foreach ($w in $xl.Workbooks) {
        if ([string]::Equals([string]$w.FullName, $Hit.Path, [System.StringComparison]::OrdinalIgnoreCase)) { $wb = $w; break }
    }

    if ($null -eq $wb) {
        # 自動化で開いたブックは既定でマクロが有効になるため、
        # 手で開いたときと同じくセキュリティセンターの設定に従わせる
        $saved = $xl.AutomationSecurity
        try {
            $xl.AutomationSecurity = $msoAutomationSecurityByUI
            $wb = $xl.Workbooks.Open($Hit.Path)
        }
        finally {
            try { $xl.AutomationSecurity = $saved } catch { }
        }
    }

    $wb.Activate()
    $ws = $wb.Sheets.Item($Hit.Sheet)
    if ($ws.Visible -ne $xlSheetVisible) {
        Write-Warning ("シート「{0}」は非表示のため、ブックだけ開きました: {1}" -f $Hit.Sheet, $Hit.File)
        return
    }
    $ws.Activate()

    if ($Hit.Cell) {
        # Goto の第 2 引数 true で、セルが画面左上に来るようスクロールする
        $xl.Goto($ws.Range($Hit.Cell), $true)
    }
    if ($Hit.Kind -eq '図形' -and $Hit.Name) {
        try { $ws.Shapes.Item($Hit.Name).Select() } catch { }
    }
}

function Show-HitPicker {
    #     検索結果を一覧ウィンドウに出し、選んだ行を Excel で開く
    param($Hits)

    $selected = @($Hits |
        Select-Object File, Sheet, Cell, Kind, Text, Name, Path |
        Out-GridView -Title 'Excel で開く行を選んで OK を押してください（Ctrl / Shift で複数選択）' -PassThru)

    foreach ($h in $selected) {
        try { Open-HitInExcel -Hit $h }
        catch { Write-Warning ("開けませんでした: {0} {1}!{2} : {3}" -f $h.File, $h.Sheet, $h.Cell, $_.Exception.Message) }
    }
}

# ---------------------------------------------------------------- 対象ファイルの列挙

function Get-FileSignature {
    # 先頭 4 バイトで実体を判定する（拡張子と中身が違うファイル・暗号化ファイル対策）
    param([string] $FilePath)

    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $fs = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $buf = New-Object byte[] 4
        $n = $fs.Read($buf, 0, 4)
        if ($n -lt 4) { return 'Unknown' }
        if ($buf[0] -eq 0x50 -and $buf[1] -eq 0x4B) { return 'Zip' }
        if ($buf[0] -eq 0xD0 -and $buf[1] -eq 0xCF -and $buf[2] -eq 0x11 -and $buf[3] -eq 0xE0) { return 'Ole' }
        return 'Unknown'
    }
    finally { $fs.Dispose() }
}

$extSet = @{}
foreach ($e in $Extensions) {
    $k = $e.ToLowerInvariant()
    if (-not $k.StartsWith('.')) { $k = '.' + $k }
    $extSet[$k] = $true
}

$files = New-Object System.Collections.Generic.List[object]
$seen  = @{}

foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) {
        Write-Warning "見つかりません: $p"
        continue
    }

    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            $root  = $item.FullName.TrimEnd('\', '/')
            $found = Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue
        }
        else {
            $root  = [System.IO.Path]::GetDirectoryName($item.FullName)
            $found = @($item)
        }

        foreach ($f in $found) {
            if (-not $extSet.ContainsKey($f.Extension.ToLowerInvariant())) { continue }
            if ($f.Name.StartsWith('~$')) { continue }   # Excel の一時ファイル
            if ($seen.ContainsKey($f.FullName)) { continue }
            # 前回の -OutXlsx の出力が検索フォルダ内にあっても、それ自体は検索しない
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $seen[$f.FullName] = $true

            $rel = $f.FullName
            if ($rel.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
                $rel = $rel.Substring($root.Length).TrimStart('\', '/')
            }
            $files.Add([pscustomobject]@{ Info = $f; Relative = $rel })
        }
    }
}

if ($files.Count -eq 0) {
    Write-Warning '検索対象の Excel ファイルがありません。'
    return
}

# ---------------------------------------------------------------- 検索

$allHits    = New-Object System.Collections.Generic.List[object]
$failures   = New-Object System.Collections.Generic.List[object]
$hitFiles   = 0
$stopwatch  = [System.Diagnostics.Stopwatch]::StartNew()

try {
    $index = 0
    foreach ($entry in $files) {
        $index++
        $file = $entry.Info
        $script:CurrentFile = $entry.Relative
        $script:CurrentPath = $file.FullName

        Write-Progress -Activity 'Excel を検索中' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) `
            -PercentComplete ([int](100 * $index / $files.Count))

        try {
            $ext = $file.Extension.ToLowerInvariant()
            $sig = Get-FileSignature -FilePath $file.FullName

            if ($ext -eq '.xlsb' -or ($sig -eq 'Ole' -and $ext -eq '.xls') -or $sig -eq 'Unknown') {
                if ($NoExcel) {
                    Write-Verbose "スキップ（-NoExcel 指定のため .xls / .xlsb は対象外）: $($entry.Relative)"
                    continue
                }
                $hits = @(Search-ComWorkbook -FilePath $file.FullName)
            }
            elseif ($sig -eq 'Ole') {
                throw 'パスワードで保護されているため読めません'
            }
            else {
                $hits = @(Search-OpenXmlWorkbook -FilePath $file.FullName)
            }

            if ($hits.Count -gt 0) {
                $hitFiles++
                foreach ($h in $hits) {
                    $allHits.Add($h)
                    if (-not $Open) { $h }   # -Open 時は一覧ウィンドウに出すのでコンソールには流さない
                }
            }
        }
        catch {
            $ex = $_.Exception
            while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
            $msg = $ex.Message
            if ($ex -is [System.IO.InvalidDataException]) { $msg = "ファイルが壊れているか Excel 形式ではありません（$msg）" }
            $failures.Add([pscustomobject]@{ File = $entry.Relative; Error = $msg })
            Write-Warning ("読めませんでした: {0} : {1}" -f $entry.Relative, $msg)
        }
    }
}
finally {
    Write-Progress -Activity 'Excel を検索中' -Completed
    Close-ExcelApplication
}

$stopwatch.Stop()

# ---------------------------------------------------------------- 結果

if ($OutCsv) {
    $encoding = 'UTF8'
    if ($PSVersionTable.PSVersion.Major -ge 6) { $encoding = 'utf8BOM' }

    $allHits |
        Select-Object @{ n = 'ファイル'; e = { $_.File } },
                      @{ n = 'シート'; e = { $_.Sheet } },
                      @{ n = 'セル'; e = { $_.Cell } },
                      @{ n = '種類'; e = { $_.Kind } },
                      @{ n = '内容'; e = { $_.Text } },
                      @{ n = '名前'; e = { $_.Name } },
                      @{ n = 'フルパス'; e = { $_.Path } } |
        Export-Csv -LiteralPath $OutCsv -NoTypeInformation -Encoding $encoding
}

$xlsxSaved = $false
if ($OutXlsx) {
    try {
        Export-HitsToXlsx -Hits $allHits -FilePath $OutXlsx
        $xlsxSaved = $true
    }
    catch {
        Write-Warning ("xlsx を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
    }
}

Write-Host ''
Write-Host ("{0} ファイル中 {1} ファイルで {2} 件ヒット（{3:N1} 秒）" -f $files.Count, $hitFiles, $allHits.Count, $stopwatch.Elapsed.TotalSeconds)
if ($failures.Count -gt 0) {
    Write-Host ("読めなかったファイル {0} 件" -f $failures.Count) -ForegroundColor Yellow
    foreach ($f in $failures) { Write-Host ("  {0} : {1}" -f $f.File, $f.Error) -ForegroundColor Yellow }
}
if ($OutCsv) { Write-Host "結果を保存しました: $OutCsv" }
if ($xlsxSaved) { Write-Host "結果を保存しました: $OutXlsx" }

if ($Open -and $allHits.Count -gt 0) { Show-HitPicker -Hits $allHits }

# .SYNOPSIS
#     フォルダ配下の Excel ファイルを横断して文字列を検索します（Excel 版 grep）。
#
# .DESCRIPTION
#     セルの値・数式・コメント（メモ）・図形やテキストボックスの文字・シート名を検索し、
#     ヒットしたファイル・シート・セル番地・内容を一覧表示します。
#
#     xlsx / xlsm / xltx / xltm 形式は ZIP 内の XML を直接読むため Excel を起動せず高速です。
#     Excel がインストールされていない端末でも動作します。
#     xls / xlsb 形式はファイル形式上 XML で読めないため、見つかった場合のみ Excel を起動して読みます。
#
#     結果はオブジェクトとして出力されるので、Where-Object や Export-Csv にそのまま渡せます。
#     プロパティ: File（相対パス） / Location（シート!セル） / Sheet / Cell / Kind / Text / Name / Path
#
# .PARAMETER Pattern
#     検索する文字列。既定では .NET 正規表現として扱い、大文字小文字は区別しません。
#
# .PARAMETER Path
#     検索するフォルダまたはファイル（複数指定可、ワイルドカード可）。省略時はカレントフォルダ。
#
# .PARAMETER SimpleMatch
#     Pattern を正規表現ではなく、そのままの文字列として検索します。
#     「(株)」「1.5」「[済]」のように記号を含む文字列を探すときに使います。
#
# .PARAMETER CaseSensitive
#     大文字と小文字を区別します。
#
# .PARAMETER IgnoreWidth
#     全角と半角を区別しません（Unicode NFKC 正規化で比較）。「ＡＢＣ」「ABC」、「ｶﾅ」「カナ」が一致します。
#     正規表現で使う場合、パターンは半角で書いてください。
#
# .PARAMETER Target
#     検索対象。Value（セルの値） / Formula（数式） / Comment（コメント・メモ） / Shape（図形の文字）
#     / SheetName（シート名）から複数指定できます。既定はすべて。
#
# .PARAMETER NoRecurse
#     サブフォルダを検索しません。
#
# .PARAMETER SkipHiddenSheets
#     非表示シートを検索しません（既定では非表示シートも検索します）。
#
# .PARAMETER SheetNamePattern
#     検索対象とするシート名の正規表現（大文字小文字は区別しません）。
#
# .PARAMETER SheetNameExclude
#     除外するシート名の正規表現。SheetNamePattern より優先されます。
#
# .PARAMETER NoExcel
#     Excel を起動しません。.xls / .xlsb は検索対象から外れます。
#
# .PARAMETER Open
#     検索結果を一覧ウィンドウ（Out-GridView）に表示し、選んだ行のファイルを Excel で開いて
#     該当セルを選択します。図形のヒットは図形を選択します。複数行を選ぶとすべて開きます。
#     読み取り専用ではなく通常どおり開きます。Windows 専用です。
#
# .PARAMETER OutCsv
#     結果を CSV（UTF-8 BOM 付き、Excel でそのまま開ける）に保存します。
#
# .PARAMETER OutXlsx
#     結果を Excel ファイル（.xlsx）に保存します。Excel がなくても作成できます。
#     「場所」列はハイパーリンクになっており、クリックすると該当ファイルのそのセルが開きます。
#     見出し行の固定とオートフィルターを設定済みです。
#
# .PARAMETER Extensions
#     対象とする拡張子。既定は .xlsx .xlsm .xltx .xltm .xls .xlsb
#
# .EXAMPLE
#     .\Search-Excel.ps1 '顧客ID' D:\Docs
#
# .EXAMPLE
#     # 記号を含む文字列をそのまま探す
#     .\Search-Excel.ps1 '(株)山田' D:\Docs -SimpleMatch
#
# .EXAMPLE
#     # VLOOKUP を使っている数式だけを探して CSV に保存
#     .\Search-Excel.ps1 'VLOOKUP' D:\Docs -Target Formula -OutCsv D:\vlookup.csv
#
# .EXAMPLE
#     # 結果を一覧ウィンドウで絞り込みながら見る
#     .\Search-Excel.ps1 'TODO|要確認' D:\Docs | Out-GridView
#
# .EXAMPLE
#     # 結果一覧から選んだセルを Excel で開く
#     .\Search-Excel.ps1 'TODO' D:\Docs -Open
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xls / .xlsb を検索する場合のみデスクトップ版 Excel が必要です。
