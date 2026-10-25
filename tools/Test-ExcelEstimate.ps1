[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [ValidateSet('Sum', 'Total', 'Formula', 'TextNumber', 'Conversion', 'UnitPrice', 'Rounding', 'Error')]
    [string[]] $Check = @('Sum', 'Total', 'Formula', 'TextNumber', 'Conversion', 'UnitPrice', 'Rounding', 'Error'),

    [string] $TotalLabel = '^(合計|小計|総計|計|総額|合計金額|合計工数|総合計)$',

    [string] $RoleHeader = '要員|区分|ランク|担当|役割|職種|作業者',

    [switch] $NoRecurse,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

foreach ($p in @($TotalLabel, $RoleHeader)) {
    try { [void][regex]::new($p) }
    catch { throw "正規表現が不正です: $p 詳細: $($_.Exception.InnerException.Message)" }
}
if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'ExcelEstimate.Issue' -DefaultDisplayPropertySet File, Location, Level, Detail -Force -WhatIf:$false

$checkSet = @{}
foreach ($c in $Check) { $checkSet[$c] = $true }
$nfkc = [System.Text.NormalizationForm]::FormKC
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$numberStyle = [System.Globalization.NumberStyles]::Float

# ---------------------------------------------------------------- 指摘の記録


$script:Issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    param([ValidateSet('エラー', '警告', '情報')] [string] $Level, [string] $Item, [string] $Detail, [string] $Sheet = '', [string] $Cell = '')
    $location = $Sheet
    if ($Sheet -and $Cell) { $location = '{0}!{1}' -f $Sheet, $Cell }
    $issue = [pscustomobject]@{ File = $script:CurrentFile; Location = $location; Level = $Level; Item = $Item; Detail = $Detail; Sheet = $Sheet; Cell = $Cell; Path = $script:CurrentPath }
    $issue.PSObject.TypeNames.Insert(0, 'ExcelEstimate.Issue')
    $script:Issues.Add($issue)
}

# ---------------------------------------------------------------- 数式の相対表記（R1C1）


# セル参照（$A$1 など）。関数名（LOG10( など）や、名前の途中の文字には一致させない
$refRx = New-Object System.Text.RegularExpressions.Regex('"[^"]*"|(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])')

function ConvertTo-R1C1 {
    #     数式のセル参照を、セルからの相対位置（R[1]C[-2]）に直す。同じ計算をしている数式どうしは同じ文字になる
    param([string] $Formula, [int] $Row, [int] $Col)
    return $refRx.Replace($Formula, {
        param($m)
        if ($m.Value.StartsWith('"')) { return $m.Value }
        $c = Get-ColumnIndex $m.Groups[2].Value
        $r = [int]$m.Groups[4].Value
        $rp = 'R[{0}]' -f ($r - $Row)
        if ($m.Groups[3].Value) { $rp = "R$r" }
        $cp = 'C[{0}]' -f ($c - $Col)
        if ($m.Groups[1].Value) { $cp = "C$c" }
        return $rp + $cp
    })
}

# ---------------------------------------------------------------- 読み取り


function Read-SharedStringList {
    param($Package, [string] $PartName)
    $list = New-Object System.Collections.Generic.List[string]
    if (-not $PartName) { return , $list }
    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return , $list }
    try {
        $sb = New-Object System.Text.StringBuilder
        $inPhonetic = $false
        $advance = $true
        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                if ($reader.LocalName -eq 'si') { [void]$sb.Clear(); if ($reader.IsEmptyElement) { $list.Add('') } }
                elseif ($reader.LocalName -eq 'rPh' -and -not $reader.IsEmptyElement) { $inPhonetic = $true }
                elseif ($reader.LocalName -eq 't') {
                    # ReadElementContentAsString は次のノードまで進むので、次の Read を飛ばす
                    $t = $reader.ReadElementContentAsString()
                    $advance = $false
                    if (-not $inPhonetic) { [void]$sb.Append($t) }
                }
            }
            elseif ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
                if ($reader.LocalName -eq 'si') { $list.Add($sb.ToString()) }
                elseif ($reader.LocalName -eq 'rPh') { $inPhonetic = $false }
            }
        }
    }
    finally { $reader.Dispose() }
    return , $list
}

# 組み込みの表示形式のうち日付・時刻のもの（27〜36, 50〜58 は日本語ロケールの和暦・年月日。46 は経過時間の [h]:mm:ss なので除く）
$builtinDateFormatIds = @(14..22) + @(27..36) + @(45, 47) + @(50..58)

function Read-DateStyleList {
    #     styles.xml の cellXfs を読み、書式の番号ごとに日付・時刻の表示形式かを返す（Search-Excel.ps1 の Read-DateStyles と同じ判定）。
    #     ユーザー定義の形式は、文字（"..."・\x）や [赤]・[$-411] などと General（日本語の G/標準）を除いて
    #     y/m/d/h/s、日本語の曜日の a（aaa）、元号の g（ggge）、和暦の年の e が残れば日付とみなす。経過時間（[h]:mm）は除く
    param($Package, [string] $PartName)
    $list = New-Object System.Collections.Generic.List[bool]
    $doc = $null
    if ($PartName) { $doc = Read-PartXml -Package $Package -PartName $PartName }
    if ($null -eq $doc) { return , $list }
    $custom = @{}
    $id = 0
    foreach ($nf in $doc.SelectNodes("//*[local-name()='numFmts']/*[local-name()='numFmt']")) {
        if (-not [int]::TryParse($nf.GetAttribute('numFmtId'), [ref]$id)) { continue }
        $code = $nf.GetAttribute('formatCode')
        $s = $code.Split(';')[0]
        foreach ($rx in @('"[^"]*"', '\\.', '[_*].', '\[[^\]]*\]', '(?i)General|G/標準', '[Ee][+-]')) { $s = [regex]::Replace($s, $rx, '') }
        $custom[$id] = ($code -notmatch '\[(h+|m+|s+)\]') -and ($s -match '[ymdhsage]')
    }
    foreach ($xf in $doc.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
        $id = 0
        [void][int]::TryParse($xf.GetAttribute('numFmtId'), [ref]$id)
        if ($custom.ContainsKey($id)) { $list.Add($custom[$id]) } else { $list.Add($builtinDateFormatIds -contains $id) }
    }
    return , $list
}

function Read-SheetCells {
    #     シートのセルを読む。返り値: 行番号 → (列番号 → セル)。共有数式の写しのセルにも、元の数式から求めた数式を入れる
    param($Package, [string] $PartName, $Strings)

    $rows = @{}
    $shared = @{}   # 共有数式の番号 → 先頭のセル
    $pending = New-Object System.Collections.Generic.List[object]
    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return $rows }
    try {
        $cell = $null
        $open = $null   # 読んでいる途中のセル（値か数式が出てきたら、そこで記録する。方眼紙の空のセルは記録しない）
        $rowNo = 0; $colNo = 0
        $advance = $true
        $take = {
            if (-not $script:readCell) {
                $script:readCell = [pscustomobject]@{ Ref = $open[0]; Row = $open[1]; Col = $open[2]; Type = $open[3]; Style = $open[4]; Raw = $null; Text = $null; Formula = $null; R1C1 = $null; SharedIndex = $null }
                if (-not $rows.ContainsKey($open[1])) { $rows[$open[1]] = @{} }
                $rows[$open[1]][$open[2]] = $script:readCell
            }
            $script:readCell
        }
        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                switch ($reader.LocalName) {
                    'row' {
                        $r = $reader.GetAttribute('r')
                        if ($r) { $rowNo = [int]$r } else { $rowNo++ }
                        $colNo = 0
                    }
                    'c' {
                        $ref = $reader.GetAttribute('r')
                        if ($ref) { $colNo = Get-ColumnIndex $ref } else { $colNo++; $ref = (ConvertTo-ColumnName $colNo) + $rowNo }
                        $open = $null
                        $script:readCell = $null
                        if (-not $reader.IsEmptyElement) { $open = @($ref, $rowNo, $colNo, [string]$reader.GetAttribute('t'), [string]$reader.GetAttribute('s')) }
                    }
                    'f' {
                        if ($open) {
                            $cell = & $take
                            $ft = $reader.GetAttribute('t'); $si = $reader.GetAttribute('si')
                            $text = ''
                            if (-not $reader.IsEmptyElement) { $text = $reader.ReadElementContentAsString(); $advance = $false }
                            if ($text) {
                                $cell.Formula = $text
                                $cell.R1C1 = ConvertTo-R1C1 $text $cell.Row $cell.Col
                                if ($ft -eq 'shared' -and $si) { $shared[$si] = $cell }
                            }
                            elseif ($ft -eq 'shared' -and $si) { $cell.SharedIndex = $si; $pending.Add($cell) }
                        }
                    }
                    'v' {
                        if ($open) { $cell = & $take; $cell.Raw = $reader.ReadElementContentAsString(); $advance = $false }
                    }
                    'is' {
                        if ($open) { $cell = & $take; $cell.Raw = $reader.ReadInnerXml(); $cell.Text = ([xml]('<x>' + $cell.Raw + '</x>')).DocumentElement.InnerText; $advance = $false }
                    }
                }
            }
            elseif ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement -and $reader.LocalName -eq 'c') { $open = $null; $script:readCell = $null }
        }
    }
    finally { $reader.Dispose() }

    foreach ($c in $pending) {
        if ($shared.ContainsKey($c.SharedIndex)) {
            # 共有数式の写しのセルは、先頭のセルの数式を位置に合わせてずらしたもの
            $master = $shared[$c.SharedIndex]
            $c.R1C1 = $master.R1C1
            $c.Formula = Move-FormulaReference $master.Formula ($c.Row - $master.Row) ($c.Col - $master.Col)
        }
    }
    $idx = 0
    foreach ($r in $rows.Values) {
        foreach ($c in $r.Values) {
            if ($c.Type -eq 's' -and [int]::TryParse([string]$c.Raw, [ref]$idx) -and $idx -lt $Strings.Count) { $c.Text = $Strings[$idx] }
            elseif ($c.Type -ne 'inlineStr') { $c.Text = $c.Raw }
        }
    }
    return $rows
}

function Test-NumberCell {
    #     数値のセル（数値の定数、または文字以外の結果になる数式）か
    param($Cell)
    if ($null -eq $Cell) { return $false }
    if (@('s', 'inlineStr', 'str', 'b', 'e') -contains $Cell.Type) { return $false }
    if ($Cell.Formula) { return $true }
    $d = 0.0
    return [double]::TryParse([string]$Cell.Raw, $numberStyle, $invariant, [ref]$d)
}

function Test-TextCell {
    param($Cell)
    return $Cell -and (@('s', 'inlineStr') -contains $Cell.Type) -and ([string]$Cell.Text).Trim()
}

function Test-DateCell {
    #     日付・時刻の表示形式のセルか（月の見出しなど。DateStyles は Read-DateStyleList）。値は日付のシリアル値で読まれる
    param($Cell, $DateStyles)
    $i = 0
    if ($Cell.Style) { [void][int]::TryParse($Cell.Style, [ref]$i) }
    return $i -lt $DateStyles.Count -and $DateStyles[$i]
}

function Format-Number {
    param([string] $Raw)
    $d = 0.0
    if ([double]::TryParse($Raw, $numberStyle, $invariant, [ref]$d)) { return $d.ToString('#,0.##', $invariant) }
    return $Raw
}

# ---------------------------------------------------------------- シートごとのチェック


# 数式の中の、同じシートのセル・範囲の参照（A1・$A$1・A1:B5）。文字列（"..."）とシート名（'...'）は飛ばし、ほかのシートの参照（内訳!A1）には一致させない
$rangeRx = New-Object System.Text.RegularExpressions.Regex('"[^"]*"|''[^'']*''|(?<![A-Za-z0-9_.!:$])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(!])')

function Test-FormulaUsesCell {
    #     数式が、同じシートのセル（行・列）をセルの参照か範囲で使っているか（=SUM(D2:D4)-D5 の D5 など）
    param([string] $Formula, [int] $Row, [int] $Col)
    foreach ($m in $rangeRx.Matches($Formula)) {
        if (-not $m.Groups[1].Success) { continue }
        $r0 = [int]$m.Groups[2].Value; $c0 = Get-ColumnIndex $m.Groups[1].Value
        $r1 = $r0; $c1 = $c0
        if ($m.Groups[3].Success) { $r1 = [int]$m.Groups[4].Value; $c1 = Get-ColumnIndex $m.Groups[3].Value }
        if ($Row -ge [math]::Min($r0, $r1) -and $Row -le [math]::Max($r0, $r1) -and $Col -ge [math]::Min($c0, $c1) -and $Col -le [math]::Max($c0, $c1)) { return $true }
    }
    return $false
}

function Test-Sheet {
    param([string] $Sheet, $Rows, $Book)

    $rowNos = @($Rows.Keys | Sort-Object)
    $cellAt = { param($r, $c) if ($Rows.ContainsKey($r)) { $Rows[$r][$c] } else { $null } }

    foreach ($r in $rowNos) {
        # 上下の行（セルごとに関数を呼ばずに引けるよう、先に取っておく）
        $prevRow = $Rows[$r - 1]; $nextRow = $Rows[$r + 1]
        foreach ($c in @($Rows[$r].Keys | Sort-Object)) {
            $cell = $Rows[$r][$c]
            $up = $null; $down = $null
            if ($prevRow) { $up = $prevRow[$c] }
            if ($nextRow) { $down = $nextRow[$c] }

            # ---- エラー値
            if ($checkSet['Error'] -and $cell.Type -eq 'e') {
                Add-Issue 'エラー' 'エラー値' ("セルがエラー値（{0}）です" -f $cell.Raw) $Sheet $cell.Ref
            }

            # ---- 合計の範囲: SUM の範囲のすぐ外（合計のセルとの間、範囲の上）に数値が残っていないか
            if ($checkSet['Sum'] -and $cell.Formula) {
                foreach ($m in [regex]::Matches($cell.Formula, '(?i)(?<![A-Za-z0-9_.])SUM\(\s*\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)\s*\)')) {
                    $c0 = Get-ColumnIndex $m.Groups[1].Value; $r0 = [int]$m.Groups[2].Value
                    $c1 = Get-ColumnIndex $m.Groups[3].Value; $r1 = [int]$m.Groups[4].Value
                    $missing = @()
                    if ($c0 -eq $c1 -and $c0 -eq $cell.Col -and $r1 -lt $cell.Row) {
                        # 縦の合計
                        for ($k = $r1 + 1; $k -lt $cell.Row; $k++) { $x = & $cellAt $k $c0; if ((Test-NumberCell $x) -and $x.Formula -notmatch '(?i)SUM\(') { $missing += $x } }
                        # 範囲のすぐ上。表の見出しの行（列のいちばん上の値。年度など）と日付（月の見出しなど）は合計する値ではない
                        $above = & $cellAt ($r0 - 1) $c0
                        if ((Test-NumberCell $above) -and -not $above.Formula -and (Test-NumberCell (& $cellAt ($r0) $c0)) -and -not (Test-DateCell $above $Book.DateStyles) -and
                            @($rowNos | Where-Object { $_ -lt $above.Row } | ForEach-Object { & $cellAt $_ $c0 } | Where-Object { (Test-NumberCell $_) -or (Test-TextCell $_) }).Count -gt 0) { $missing += $above }
                    }
                    elseif ($r0 -eq $r1 -and $r0 -eq $cell.Row -and $c1 -lt $cell.Col) {
                        # 横の合計
                        for ($k = $c1 + 1; $k -lt $cell.Col; $k++) { $x = & $cellAt $r0 $k; if ((Test-NumberCell $x) -and $x.Formula -notmatch '(?i)SUM\(') { $missing += $x } }
                    }
                    # 合計の数式がほかで使っているセル（=SUM(D2:D4)-D5 で引く値引きなど）は入れ忘れではない
                    $missing = @($missing | Where-Object { -not (Test-FormulaUsesCell $cell.Formula $_.Row $_.Col) } | ForEach-Object { $_.Ref })
                    if ($missing.Count -gt 0) {
                        Add-Issue '警告' '合計の範囲' ("合計（={0}）の範囲に {1} が入っていません" -f $cell.Formula, ($missing -join '、')) $Sheet $cell.Ref
                    }
                }
            }

            # ---- 合計の行に、数式ではなく数値が入っている（文字のセルが 2 つまでの行だけ。項目の一覧の「合計金額」などを除く）
            if ($checkSet['Total'] -and (Test-TextCell $cell) -and ($cell.Text.Normalize($nfkc) -replace '\s', '') -match $TotalLabel -and
                @($Rows[$r].Values | Where-Object { Test-TextCell $_ }).Count -le 2) {
                foreach ($c2 in @($Rows[$r].Keys | Where-Object { $_ -gt $c } | Sort-Object)) {
                    $x = $Rows[$r][$c2]
                    if (-not (Test-NumberCell $x) -or $x.Formula) { continue }
                    # 上に数値が 2 つ以上ある列だけ（合計する対象がある列）
                    $above = @($rowNos | Where-Object { $_ -lt $r } | ForEach-Object { & $cellAt $_ $c2 } | Where-Object { Test-NumberCell $_ })
                    if ($above.Count -ge 2) {
                        Add-Issue '警告' '手入力の合計' ("「{0}」の行の {1} が数式ではなく数値（{2}）です" -f $cell.Text.Trim(), $x.Ref, (Format-Number $x.Raw)) $Sheet $x.Ref
                    }
                }
            }

            # ---- 文字列の数値: 数値の列の中で、数字が文字列として入っている（SUM に含まれない）。
            #      上下 3 行に数値のセルが 2 つ以上あり、文字のセルより多い場合だけ（表紙の「1.0」などの版数は除く）
            if ($checkSet['TextNumber'] -and (Test-TextCell $cell)) {
                $t = $cell.Text.Normalize($nfkc).Trim()
                if ($t -match '^[-+]?[\d,]+(\.\d+)?$' -and $t -match '\d') {
                    $numbers = 0; $texts = 0
                    for ($k = $r - 3; $k -le $r + 3; $k++) {
                        if ($k -eq $r) { continue }
                        $x = & $cellAt $k $c
                        if (Test-TextCell $x) { $texts++ } elseif (Test-NumberCell $x) { $numbers++ }
                    }
                    if ($numbers -ge 2 -and $numbers -gt $texts) {
                        Add-Issue '警告' '文字列の数値' ("「{0}」が文字列として入力されています（合計などの計算に含まれません）" -f $cell.Text.Trim()) $Sheet $cell.Ref
                    }
                }
            }

            # ---- 列の中で、上下と違う数式・数式の列に入った数値（Excel の「周りと矛盾する数式」と同じ考え方）
            if ($checkSet['Formula'] -and $up -and $down -and ($cell.Formula -or -not $cell.Type -or $cell.Type -eq 'n')) {
                if ($up -and $up.R1C1 -and $up.R1C1 -ceq $down.R1C1 -and $up.R1C1 -notmatch '(?i)SUM\(') {
                    if ($cell.R1C1 -and $cell.R1C1 -cne $up.R1C1) {
                        Add-Issue '警告' '周りと違う数式' ("上下のセル（={0}）と違う計算です（={1}）" -f $up.Formula, $cell.Formula) $Sheet $cell.Ref
                    }
                    elseif (-not $cell.Formula -and (Test-NumberCell $cell)) {
                        Add-Issue '警告' '数式の列の数値' ("上下のセルは数式（={0}）ですが、数値（{1}）が直接入力されています" -f $up.Formula, (Format-Number $cell.Raw)) $Sheet $cell.Ref
                    }
                }
            }

            # 換算・端数処理・単価はブック全体で比べるので集めておく（*20% などの率は換算ではない）
            if ($cell.Formula) {
                foreach ($m in [regex]::Matches($cell.Formula, '[*/]\s*(\d+(?:\.\d+)?)(?![\d.%])')) {
                    $n = [double]::Parse($m.Groups[1].Value, $invariant)
                    if (@(19, 20, 21, 22) -contains $n) { $Book.Days.Add([pscustomobject]@{ Value = $n; Where = ('{0}!{1}' -f $Sheet, $cell.Ref); Sheet = $Sheet; Cell = $cell.Ref }) }
                    elseif (@(7.5, 7.75, 8) -contains $n) { $Book.Hours.Add([pscustomobject]@{ Value = $n; Where = ('{0}!{1}' -f $Sheet, $cell.Ref); Sheet = $Sheet; Cell = $cell.Ref }) }
                }
                foreach ($m in [regex]::Matches($cell.Formula.ToUpperInvariant(), '(?<![A-Z0-9_.])(ROUNDDOWN|ROUNDUP|ROUND|INT|TRUNC|FLOOR|CEILING)\(')) {
                    $Book.Rounding.Add([pscustomobject]@{ Value = $m.Groups[1].Value; Where = ('{0}!{1}' -f $Sheet, $cell.Ref); Sheet = $Sheet; Cell = $cell.Ref })
                }
            }
        }
    }

    # ---- 単価: 「単価」の列と、要員（担当・区分など）の列がある表で、同じ要員の単価が違う
    if ($checkSet['UnitPrice']) {
        foreach ($r in $rowNos) {
            $priceCol = $null; $roleCol = $null
            foreach ($c in $Rows[$r].Keys) {
                $x = $Rows[$r][$c]
                if (-not (Test-TextCell $x)) { continue }
                $label = $x.Text.Normalize($nfkc) -replace '\s', ''
                if (-not $priceCol -and $label -match '単価') { $priceCol = $c }
                elseif (-not $roleCol -and $label -match $RoleHeader) { $roleCol = $c }
            }
            if (-not $priceCol -or -not $roleCol) { continue }
            foreach ($k in @($rowNos | Where-Object { $_ -gt $r })) {
                $role = & $cellAt $k $roleCol; $price = & $cellAt $k $priceCol
                if (-not (Test-TextCell $role) -or -not (Test-NumberCell $price) -or $price.Formula) { continue }
                $key = $role.Text.Trim()
                if (-not $Book.Prices.Contains($key)) { $Book.Prices[$key] = New-Object System.Collections.Generic.List[object] }
                $Book.Prices[$key].Add([pscustomobject]@{ Value = (Format-Number $price.Raw); Where = ('{0}!{1}' -f $Sheet, $price.Ref); Sheet = $Sheet; Cell = $price.Ref })
            }
        }
    }
}

function Add-MixedIssue {
    #     ブック全体で、同じ目的の値（換算の日数・端数処理の関数・単価）が混在していれば 1 件にまとめて指摘する
    param($Entries, [string] $Level, [string] $Item, [string] $Format)
    # 同じ数の値は、先に出てきた順（Index）に並べる
    $groups = @($Entries | Group-InOrder Value)
    if ($groups.Count -lt 2) { return }
    $parts = @($groups | Sort-Object @{ e = { $_.Count }; Descending = $true }, Index | ForEach-Object {
        $more = ''
        if ($_.Count -gt 1) { $more = ' ほか {0} か所' -f ($_.Count - 1) }
        '{0}: {1}{2}' -f $_.Name, $_.Group[0].Where, $more
    })
    $minor = @($groups | Sort-Object Count, Index)[0].Group[0]
    Add-Issue $Level $Item ($Format -f ($parts -join ' / ')) $minor.Sheet $minor.Cell
}

function Test-Book {
    param([string] $FilePath)
    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) { if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break } }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        if ($null -eq $wbDoc) { throw 'Excel ブックではありません（workbook.xml が見つかりません）' }
        $relById = @{}
        $sstPart = $null; $stylesPart = $null
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName $wbPart)) {
            $relById[$rel.Id] = $rel
            if ($rel.Kind -eq 'sharedStrings') { $sstPart = $rel.Target }
            elseif ($rel.Kind -eq 'styles') { $stylesPart = $rel.Target }
        }
        $strings = Read-SharedStringList $pkg $sstPart
        $book = [pscustomobject]@{
            Days = New-Object System.Collections.Generic.List[object]; Hours = New-Object System.Collections.Generic.List[object]
            Rounding = New-Object System.Collections.Generic.List[object]; Prices = [ordered]@{}
            DateStyles = (Read-DateStyleList $pkg $stylesPart)
        }
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            # マクロシート（Excel 4.0 マクロ。種類は xlMacrosheet・xlIntlMacrosheet）もセルの並びは同じなので見る
            if ($null -eq $rel -or ($rel.Kind -ne 'worksheet' -and $rel.Kind -notmatch 'macrosheet')) { continue }
            Test-Sheet -Sheet $s.GetAttribute('name') -Rows (Read-SheetCells $pkg $rel.Target $strings) -Book $book
        }
    }
    finally { $pkg.Zip.Dispose() }

    if ($checkSet['Conversion']) {
        Add-MixedIssue $book.Days '警告' '換算の日数' '人日と人月の換算に、違う日数が使われています（{0}）'
        Add-MixedIssue $book.Hours '警告' '換算の時間' '人時と人日の換算に、違う時間が使われています（{0}）'
    }
    if ($checkSet['Rounding']) { Add-MixedIssue $book.Rounding '情報' '端数処理' '端数処理の関数が混在しています（{0}）' }
    if ($checkSet['UnitPrice']) {
        foreach ($role in $book.Prices.Keys) { Add-MixedIssue $book.Prices[$role] '警告' '単価' ("「{0}」の単価が複数あります（{{0}}）" -f $role.Replace('{', '{{').Replace('}', '}}')) }
    }
}

# ---------------------------------------------------------------- 実行


$files = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($p in $Path) {
    $found = @(Get-PathItem $p)
    if ($found.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $found) {
        if ($item.PSIsContainer) {
            $root = $item.FullName.TrimEnd('\', '/')
            $list = Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue
        }
        else {
            $root = [System.IO.Path]::GetDirectoryName($item.FullName)
            $list = @($item)
        }
        foreach ($f in $list) {
            if (@('.xlsx', '.xlsm', '.xltx', '.xltm') -notcontains $f.Extension.ToLowerInvariant()) { continue }
            if ($f.Name.StartsWith('~$')) { continue }
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $files.Add([pscustomobject]@{ Info = $f; Relative = $f.FullName.Substring($root.Length).TrimStart('\', '/') })
        }
    }
}
if ($files.Count -eq 0) { Write-Warning 'チェックする Excel ファイル（.xlsx / .xlsm）がありません。'; return }

$index = 0
foreach ($entry in $files) {
    $index++
    $script:CurrentFile = $entry.Relative
    $script:CurrentPath = $entry.Info.FullName
    Write-Progress -Activity '見積書の検算' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) -PercentComplete ([int](100 * $index / $files.Count))
    try {
        # パスワード付きのファイルは中身が暗号化された別の形式（OLE）になっていて読めない
        # Excel で開いている最中のファイルも読めるよう、共有モードで開く
        $fs = New-Object System.IO.FileStream($entry.Info.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
        try { $head = New-Object byte[] 2; $n = $fs.Read($head, 0, 2) } finally { $fs.Dispose() }
        if ($n -eq 2 -and $head[0] -eq 0xD0 -and $head[1] -eq 0xCF) {
            Add-Issue '情報' 'パスワード' 'パスワードで保護（暗号化）されているためチェックしていません'
            continue
        }
        Test-Book $entry.Info.FullName
    }
    catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $msg = $ex.Message
        if ($ex -is [System.IO.InvalidDataException]) { $msg = "ファイルが壊れているか、パスワードで保護されています（$msg）" }
        Add-Issue 'エラー' '読み取り' "読み取れませんでした: $msg"
    }
}
Write-Progress -Activity '見積書の検算' -Completed

$levelOrder = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
$sorted = @($script:Issues | Sort-Object @{ e = { $_.File } }, @{ e = { $levelOrder[$_.Level] } }, @{ e = { $_.Sheet } }, @{ e = { $_.Cell } })
$sorted

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.File, $i.Level, $i.Item, $i.Location, $i.Detail) }
        if ($i.Sheet) {
            $cell = $i.Cell
            if (-not $cell) { $cell = 'A1' }
            $row.Link = @{ Column = 3; Path = $i.Path; Location = ("'{0}'!{1}" -f $i.Sheet.Replace("'", "''"), $cell) }
        }
        else { $row.Link = @{ Column = 0; Path = $i.Path; Location = '' } }
        $rows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(@{
            Name = '検算結果'
            Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 16 }, @{ Header = '場所'; Width = 20 }, @{ Header = '内容'; Width = 90 })
            Rows = $rows
        })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$ng = @($script:Issues | Where-Object { $_.Level -ne '情報' } | ForEach-Object { $_.Path } | Sort-Object -Unique).Count
Write-Host ("{0} ファイル中 {1} ファイルが要確認（エラー {2} 件 / 警告 {3} 件 / 情報 {4} 件）" -f $files.Count, $ng,
    @($script:Issues | Where-Object Level -eq 'エラー').Count, @($script:Issues | Where-Object Level -eq '警告').Count, @($script:Issues | Where-Object Level -eq '情報').Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     見積書（Excel）の計算のミスになりやすい箇所を見つけます（Excel 不要。ファイルは変更しません）。
#
# .DESCRIPTION
#     書式を決めずに、セルの値と数式から次の箇所を見つけます。
#
#       Sum         合計（SUM）の範囲のすぐ外に、数値のセルが残っている（最後の行を範囲に入れ忘れた など）
#       Total       「合計」「小計」の行に、数式ではなく数値が直接入力されている
#       Formula     列の中で、上下のセルと違う計算をしている数式、数式の列に直接入力された数値
#       TextNumber  数値の列の中で、数字が文字列として入力されている（SUM などの計算に含まれない）
#       Conversion  人日と人月（÷20・÷21 など）、人時と人日（÷8・÷7.5 など）の換算に違う数が使われている
#       UnitPrice   「単価」と要員（担当・区分など）の列がある表で、同じ要員の単価が複数ある
#       Rounding    端数処理の関数（ROUND・ROUNDDOWN・INT など）が混在している（情報）
#       Error       エラー値（#REF! など）のセル
#
#     結果は 1 指摘 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#
# .PARAMETER Path
#     見積書のファイルまたはフォルダ（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER Check
#     チェックする項目（Sum / Total / Formula / TextNumber / Conversion / UnitPrice / Rounding / Error）。既定はすべて。
#
# .PARAMETER TotalLabel
#     合計の行の見出しの正規表現。既定は「合計・小計・総計・計・総額・合計金額・合計工数・総合計」。
#
# .PARAMETER RoleHeader
#     単価のチェックで、要員を表す列の見出しの正規表現。既定は「要員・区分・ランク・担当・役割・職種・作業者」を含むもの。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（場所をクリックすると該当セルが開きます）。
#
# .EXAMPLE
#     .\Test-ExcelEstimate.ps1 D:\見積\見積書_v3.xlsx
#
# .EXAMPLE
#     .\Test-ExcelEstimate.ps1 D:\見積 -OutXlsx D:\見積の検算.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm が対象で、Excel は不要です。同じフォルダに ExcelTools.psm1 が必要です。
#     見つけるのは「間違いの可能性が高い箇所」です。意図どおりのものもあるので、確認してから直してください。
