[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $Rules,

    [string] $Coefficients,

    [string] $ExternalTables,

    [ValidateRange(1, 31)]
    [double] $DaysPerMonth = 20,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [switch] $Summary,

    [string[]] $Baseline,

    [string] $Actuals,

    [string] $OutCoefficients,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Rules) { $Rules = Join-Path $PSScriptRoot 'estimate/規模の数え方.csv' }
if (-not $Coefficients) { $Coefficients = Join-Path $PSScriptRoot 'estimate/係数表_サンプル.csv' }
foreach ($name in @('Rules', 'Coefficients', 'OutXlsx', 'Actuals', 'OutCoefficients')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($ExternalTables) {
    try { [void][regex]::new($ExternalTables) }
    catch { throw "-ExternalTables の正規表現が不正です: $ExternalTables 詳細: $($_.Exception.InnerException.Message)" }
}

Update-TypeData -TypeName 'DesignSize.Item' -DefaultDisplayPropertySet Kind, Name, Measures, Level, Days -Force -WhatIf:$false
Update-TypeData -TypeName 'DesignSize.Change' -DefaultDisplayPropertySet Kind, Name, Change, Detail, ChangeDays -Force -WhatIf:$false
if ($OutCoefficients -and -not $Actuals) { throw '-OutCoefficients は -Actuals（実績の工数）と一緒に指定してください。' }

$nfkc = [System.Text.NormalizationForm]::FormKC

# ---------------------------------------------------------------- 数え方と係数


$ruleList = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Rules @('種類', 'ファイル名', 'シート名', '単位', '見出し', '数える列', '数えるもの') '規模の数え方')) {
    if (-not $r.'種類') { continue }
    $unit = ([string]$r.'単位').Trim()
    if (@('ファイル', 'シート', '行') -notcontains $unit) { throw ("規模の数え方の「単位」は ファイル / シート / 行 のどれかにしてください: {0}（{1}）" -f $unit, $r.'種類') }
    $headers = @(([string]$r.'見出し') -split '/' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ })
    $ruleList.Add([pscustomobject]@{
        Kind = $r.'種類'.Trim(); FilePattern = [string]$r.'ファイル名'; SheetPattern = [string]$r.'シート名'; Unit = $unit
        Headers = $headers; KeyHeader = (ConvertTo-Label $r.'数える列'); Measure = ([string]$r.'数えるもの').Trim()
        NameLabel = (ConvertTo-Label $r.'名前'); IdLabel = (ConvertTo-Label $r.'ID'); CrudHeader = (ConvertTo-Label $r.'CRUD列')
    })
}
$coefList = New-Object System.Collections.Generic.List[object]
foreach ($c in (Read-CsvFile $Coefficients @('種類', '難易度', '基準', '上限', '工数') '係数表')) {
    if (-not $c.'種類') { continue }
    $limit = $null
    if (([string]$c.'上限').Trim()) { $limit = [double]$c.'上限' }
    $perItem = $false
    if (($c.PSObject.Properties | ForEach-Object { $_.Name }) -contains '単位') { $perItem = ([string]$c.'単位').Trim() -eq '件' }
    $coefList.Add([pscustomobject]@{ Kind = $c.'種類'.Trim(); Level = ([string]$c.'難易度').Trim(); Measure = ([string]$c.'基準').Trim(); Limit = $limit; Days = [double]$c.'工数'; PerItem = $perItem })
}

# ---------------------------------------------------------------- 設計書を読む


function Read-DesignBooks {
    param([string[]] $Paths)
    $searchArgs = @{ Pattern = '\S'; Path = $Paths; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
    return ConvertTo-DesignBooks $hits
}

$books = Read-DesignBooks $Path

function Get-IfDirection {
    #     IF の向き。「送受信区分」などに 受信／送信 があればそれ、無ければ「送信元」がこのシステム（システム名）なら送信
    param($Book)
    $system = $null; $from = $null; $kind = $null
    foreach ($sheet in $Book.Sheets.Values) {
        if (-not $system) { $system = Find-LabelValue $sheet 'システム名' }
        if (-not $from) { $from = Find-LabelValue $sheet '送信元' }
        if (-not $kind) {
            foreach ($l in @('送受信区分', '連携区分', '方向', '区分')) {
                $v = Find-LabelValue $sheet $l
                if ($v -match '受信|入力|取込') { $kind = '受信'; break }
                if ($v -match '送信|出力') { $kind = '送信'; break }
            }
        }
    }
    if ($kind) { return $kind }
    if ($from -and $system -and $from.Contains($system)) { return '送信' }
    if ($from -and $system) { return '受信' }
    return '送信'
}

# ---------------------------------------------------------------- 数える


# 行の内容を比べるとき（-Baseline）に除く列（行を挿入すると振り直される番号）
$numberHeader = '^(No|NO|№|項番|番号|行番号)\.?$'

function Get-SizeItems {
    #     設計書を数え、対象（画面・テーブル…）ごとの規模を返す。Rows には、数えた行の「数える列の値 → 行の内容」を入れる（-Baseline で比べる）
    param($Books)
    $items = [ordered]@{}
    foreach ($book in $Books.Values) {
        $fileName = [System.IO.Path]::GetFileName($book.Path)
        foreach ($rule in $ruleList) {
            if ($rule.FilePattern -and $fileName -notmatch $rule.FilePattern) { continue }
            $foundAny = $false
            foreach ($sheetName in $book.Sheets.Keys) {
                if ($rule.SheetPattern -and $sheetName -notmatch $rule.SheetPattern) { continue }
                $grid = $book.Sheets[$sheetName]
                foreach ($t in (Find-DesignTables $grid $rule.Headers $rule.KeyHeader)) {
                    $foundAny = $true
                    switch ($rule.Unit) {
                        'ファイル' { $key = '{0}|{1}' -f $rule.Kind, $book.Path }
                        default { $key = '{0}|{1}|{2}' -f $rule.Kind, $book.Path, $sheetName }
                    }
                    if (-not $items.Contains($key)) {
                        $name = Find-LabelValue $grid $rule.NameLabel
                        $id = Find-LabelValue $grid $rule.IdLabel
                        if ($rule.Unit -eq 'ファイル') {
                            foreach ($s in $book.Sheets.Values) { if (-not $name) { $name = Find-LabelValue $s $rule.NameLabel }; if (-not $id) { $id = Find-LabelValue $s $rule.IdLabel } }
                        }
                        if (-not $name -and -not $id) {
                            if ($rule.Unit -eq 'ファイル') { $name = [System.IO.Path]::GetFileNameWithoutExtension($fileName) } else { $name = $sheetName }
                        }
                        $items[$key] = [pscustomobject]@{
                            Kind = $rule.Kind; Id = $id; Name = $name; File = $book.File; Sheet = $sheetName; Cell = "A$($t.HeaderRow)"; Path = $book.Path
                            Unit = $rule.Unit; Counts = [ordered]@{}; EI = 0; EQ = 0; EO = 0; Rows = @{}
                        }
                    }
                    $item = $items[$key]
                    if (-not $item.Counts.Contains($rule.Measure)) { $item.Counts[$rule.Measure] = 0; $item.Rows[$rule.Measure] = @{} }
                    $item.Counts[$rule.Measure] += $t.Rows.Count
                    # 行の内容（番号の列を除く）。数える列の値が同じ行が複数あれば、2 つ目からは番号を付けて区別する
                    $keyCol = $t.ColumnOf[$rule.KeyHeader]
                    $skip = @($t.ColumnOf.Keys | Where-Object { $_ -match $numberHeader } | ForEach-Object { $t.ColumnOf[$_] })
                    foreach ($dr in $t.Rows) {
                        $rowKey = ConvertTo-Label ([string]$dr.Cells[$keyCol])
                        $n = 1
                        while ($item.Rows[$rule.Measure].ContainsKey($rowKey + $(if ($n -gt 1) { "#$n" } else { '' }))) { $n++ }
                        if ($n -gt 1) { $rowKey += "#$n" }
                        $item.Rows[$rule.Measure][$rowKey] = (@($dr.Cells.Keys | Where-Object { $skip -notcontains $_ } | Sort-Object | ForEach-Object { ConvertTo-Label ([string]$dr.Cells[$_]) }) -join "`t")
                    }
                    # FP の試算法: 使用テーブルに登録・更新・削除（C / U / D）があるイベントは外部入力、参照（R）だけなら外部照会
                    if ($rule.CrudHeader -and $t.ColumnOf.ContainsKey($rule.CrudHeader)) {
                        $crudCol = $t.ColumnOf[$rule.CrudHeader]
                        foreach ($dr in $t.Rows) {
                            $letters = (@([regex]::Matches(([string]$dr.Cells[$crudCol]).Normalize($nfkc), '\(([CRUDcrud]+)\)') | ForEach-Object { $_.Groups[1].Value }) -join '').ToUpperInvariant()
                            if ($letters -match '[CUD]') { $item.EI++ } elseif ($letters -match 'R') { $item.EQ++ }
                        }
                    }
                }
            }
            if (-not $foundAny -and $rule.FilePattern) {
                Write-Warning ("{0}: 「{1}」の見出し（{2}）が見つからないため、{3}の{4}を数えていません" -f $book.File, $rule.Kind, ($rule.Headers -join ' / '), $rule.Kind, $rule.Measure)
            }
        }
    }

    # IF の向きと帳票（FP の試算法では、送信の IF と帳票は外部出力、受信の IF は外部入力）
    foreach ($item in $items.Values) {
        if ($item.Kind -eq 'IF') {
            $item | Add-Member -NotePropertyName Direction -NotePropertyValue (Get-IfDirection $Books[$item.Path])
            if ($item.Direction -eq '受信') { $item.EI++ } else { $item.EO++ }
        }
        elseif ($item.Kind -eq '帳票') { $item.EO++ }
    }
    return $items
}

$items = Get-SizeItems $books

# ---------------------------------------------------------------- 難易度と工数


function ConvertTo-SizeResult {
    #     対象の規模から、係数表で難易度と工数を決めた結果をつくる
    param($Item, [switch] $Quiet)
    $coefs = @($coefList | Where-Object { $_.Kind -eq $Item.Kind })
    $level = ''; $days = $null
    $measure = $null
    if ($coefs.Count -gt 0) { $measure = $coefs[0].Measure }
    if (-not $measure -or -not $Item.Counts.Contains($measure)) { $measure = @($Item.Counts.Keys)[0] }
    $size = [double]$Item.Counts[$measure]
    foreach ($c in $coefs) {
        if ($null -eq $c.Limit -or $size -le $c.Limit) {
            $level = $c.Level
            $days = $c.Days
            if ($c.PerItem) { $days = $c.Days * $size }
            break
        }
    }
    $measures = @($Item.Counts.Keys | ForEach-Object { '{0} {1}' -f $_, $Item.Counts[$_] }) -join '・'
    $display = $Item.Name
    if ($Item.Id -and $Item.Name -and $Item.Id -ne $Item.Name) { $display = '{0} {1}' -f $Item.Id, $Item.Name } elseif ($Item.Id) { $display = $Item.Id }
    if ($Item.Unit -eq '行') { $display = '{0}（{1}）' -f $Item.File, $Item.Sheet }
    if (-not $Quiet -and @($Item.Counts.Values | Where-Object { $_ -eq 0 }).Count -gt 0) {
        Write-Warning ("{0}: {1}「{2}」の{3}が 0 件です（見出しの下に行がありません）" -f $Item.File, $Item.Kind, $display, (@($Item.Counts.Keys | Where-Object { $Item.Counts[$_] -eq 0 }) -join '・'))
    }
    $isExternal = $Item.Kind -eq 'テーブル' -and $ExternalTables -and (("{0} {1}" -f $Item.Id, $Item.Name) -match $ExternalTables)
    $r = [pscustomobject]@{
        Kind = $Item.Kind; Name = $display; Measures = $measures; Size = $size; Level = $level; Days = $days
        Id = $Item.Id; File = $Item.File; Location = '{0}!{1}' -f $Item.Sheet, $Item.Cell; Sheet = $Item.Sheet; Cell = $Item.Cell; Path = $Item.Path
        Counts = $Item.Counts; External = $isExternal; EI = $Item.EI; EQ = $Item.EQ; EO = $Item.EO; Title = $Item.Name; Rows = $Item.Rows
    }
    $r.PSObject.TypeNames.Insert(0, 'DesignSize.Item')
    return $r
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($item in $items.Values) { $results.Add((ConvertTo-SizeResult $item)) }
# 規模の数え方の種類の順、ファイル・名前の順に並べる
$kindOrder = @{}
foreach ($rule in $ruleList) { if (-not $kindOrder.ContainsKey($rule.Kind)) { $kindOrder[$rule.Kind] = $kindOrder.Count } }
$results = @($results | Sort-Object @{ e = { $kindOrder[$_.Kind] } }, File, Name)

# ---------------------------------------------------------------- 集計


$kindTotals = New-Object System.Collections.Generic.List[object]
$kinds = @($ruleList | ForEach-Object { $_.Kind } | Select-Object -Unique)
foreach ($k in $kinds) {
    $ofKind = @($results | Where-Object Kind -eq $k)
    if ($ofKind.Count -eq 0) { continue }
    foreach ($g in ($ofKind | Group-InOrder Level)) {
        $sizeSum = ($g.Group | Measure-Object Size -Sum).Sum
        $daySum = ($g.Group | Where-Object { $null -ne $_.Days } | Measure-Object Days -Sum).Sum
        $kindTotals.Add([pscustomobject]@{ Kind = $k; Level = $g.Name; Count = $g.Count; Size = $sizeSum; Days = $daySum })
    }
}
$totalDays = ($results | Where-Object { $null -ne $_.Days } | Measure-Object Days -Sum).Sum
if ($null -eq $totalDays) { $totalDays = 0 }

# FP（NESMA）: データファンクションは自システムのテーブル（ILF）と、ほかのシステムが管理するテーブル（EIF、-ExternalTables）
$ilf = @($results | Where-Object { $_.Kind -eq 'テーブル' -and -not $_.External }).Count
$eif = @($results | Where-Object { $_.Kind -eq 'テーブル' -and $_.External }).Count
$ei = ($results | Measure-Object EI -Sum).Sum
$eq = ($results | Measure-Object EQ -Sum).Sum
$eo = ($results | Measure-Object EO -Sum).Sum
$fpIndicative = 35 * $ilf + 15 * $eif
$fpEstimated = 7 * $ilf + 5 * $eif + 4 * $ei + 5 * $eo + 4 * $eq

# ---------------------------------------------------------------- 前の版との差（-Baseline）


function Get-MatchKeys {
    #     前の版と対応付けるキー。種類と「ID と名前」「ID」「名前」の順に試す（ID を直し忘れた写しなどで ID が重なっても対応付けられるように）
    param($Result)
    $id = (ConvertTo-Label $Result.Id).ToUpperInvariant(); $name = (ConvertTo-Label $Result.Title).ToUpperInvariant()
    $keys = New-Object System.Collections.Generic.List[string]
    if ($id -and $name) { $keys.Add(('{0}|{1}|{2}' -f $Result.Kind, $id, $name)) }
    if ($id) { $keys.Add(('{0}|{1}|' -f $Result.Kind, $id)) }
    if ($name) { $keys.Add(('{0}||{1}' -f $Result.Kind, $name)) }
    return , $keys
}

$changes = New-Object System.Collections.Generic.List[object]
if ($Baseline) {
    $baseList = New-Object System.Collections.Generic.List[object]
    $baseByKey = @{}
    foreach ($bi in (Get-SizeItems (Read-DesignBooks $Baseline)).Values) {
        $br = ConvertTo-SizeResult $bi -Quiet
        $baseList.Add($br)
        foreach ($k in (Get-MatchKeys $br)) { if (-not $baseByKey.ContainsKey($k)) { $baseByKey[$k] = New-Object System.Collections.Generic.List[object] }; $baseByKey[$k].Add($br) }
    }
    $matchedBase = New-Object 'System.Collections.Generic.HashSet[object]'
    foreach ($r in $results) {
        $b = $null
        foreach ($k in (Get-MatchKeys $r)) {
            if (-not $baseByKey.ContainsKey($k)) { continue }
            $b = @($baseByKey[$k] | Where-Object { -not $matchedBase.Contains($_) })[0]
            if ($b) { break }
        }
        $change = '新規'; $detail = $r.Measures; $changeDays = $r.Days
        if ($b) {
            [void]$matchedBase.Add($b)
            # 数える列の値で行を対応付け、追加・変更・削除の行を数える
            $parts = New-Object System.Collections.Generic.List[string]
            $changedRows = 0; $totalRows = 0
            foreach ($m in @(@($r.Counts.Keys) + @($b.Counts.Keys) | Select-Object -Unique)) {
                $newRows = $r.Rows[$m]; $oldRows = $b.Rows[$m]
                if (-not $newRows) { $newRows = @{} }
                if (-not $oldRows) { $oldRows = @{} }
                $added = @($newRows.Keys | Where-Object { -not $oldRows.ContainsKey($_) }).Count
                $deleted = @($oldRows.Keys | Where-Object { -not $newRows.ContainsKey($_) }).Count
                $modified = @($newRows.Keys | Where-Object { $oldRows.ContainsKey($_) -and $oldRows[$_] -cne $newRows[$_] }).Count
                $changedRows += $added + $deleted + $modified
                $totalRows += [math]::Max($newRows.Count, $oldRows.Count)
                if ($added + $deleted + $modified -gt 0) {
                    $parts.Add(('{0} {1} → {2}（追加 {3}・変更 {4}・削除 {5}）' -f $m, $oldRows.Count, $newRows.Count, $added, $modified, $deleted))
                }
            }
            if ($changedRows -eq 0) { $change = '変更なし'; $detail = ''; $changeDays = 0 }
            else {
                $change = '変更'
                $detail = $parts -join '・'
                # 変更の工数の目安: 新規の工数 × 変わった行の割合
                if ($null -ne $r.Days -and $totalRows -gt 0) { $changeDays = [math]::Round($r.Days * [math]::Min(1.0, [double]$changedRows / $totalRows), 2) }
            }
        }
        $c = [pscustomobject]@{ Kind = $r.Kind; Name = $r.Name; Change = $change; Detail = $detail; ChangeDays = $changeDays; Days = $r.Days; File = $r.File; Location = $r.Location; Sheet = $r.Sheet; Cell = $r.Cell; Path = $r.Path }
        $c.PSObject.TypeNames.Insert(0, 'DesignSize.Change')
        $changes.Add($c)
    }
    foreach ($b in $baseList) {
        if ($matchedBase.Contains($b)) { continue }
        $c = [pscustomobject]@{ Kind = $b.Kind; Name = $b.Name; Change = '削除'; Detail = $b.Measures; ChangeDays = 0; Days = $null; File = $b.File; Location = $b.Location; Sheet = $b.Sheet; Cell = $b.Cell; Path = $b.Path }
        $c.PSObject.TypeNames.Insert(0, 'DesignSize.Change')
        $changes.Add($c)
    }
    $changeOrder = @{ '新規' = 0; '変更' = 1; '削除' = 2; '変更なし' = 3 }
    $changes = @($changes | Sort-Object @{ e = { $kindOrder[$_.Kind] } }, @{ e = { $changeOrder[$_.Change] } }, File, Name)
}

# ---------------------------------------------------------------- 実績から係数表をつくる（-Actuals）


$calibration = New-Object System.Collections.Generic.List[object]
if ($Actuals) {
    if (-not (Test-Path -LiteralPath $Actuals)) { throw "実績の工数の CSV が見つかりません: $Actuals" }
    $actualRows = @(Read-TextFile $Actuals | ConvertFrom-Csv)
    $columns = @()
    if ($actualRows.Count -gt 0) { $columns = @($actualRows[0].PSObject.Properties | ForEach-Object { $_.Name }) }
    $daysCol = @('工数（人日）', '実績工数', '工数', '実績（人日）', '人日' | Where-Object { $columns -contains $_ })[0]
    if (-not $daysCol -or ($columns -notcontains 'ID' -and $columns -notcontains '名前')) { throw "実績の工数の CSV には「ID」か「名前」と、「工数」（または「工数（人日）」）の列が必要です: $Actuals" }
    $matched = New-Object System.Collections.Generic.List[object]
    foreach ($a in $actualRows) {
        $d = 0.0
        if (-not [double]::TryParse(([string]$a.$daysCol).Trim(), [ref]$d)) { continue }
        $idKey = ConvertTo-Label ([string]$a.'ID'); $nameKey = ConvertTo-Label ([string]$a.'名前'); $kind = ([string]$a.'種類').Trim()
        $r = @($results | Where-Object { (-not $kind -or $_.Kind -eq $kind) -and (($idKey -and (ConvertTo-Label $_.Id) -eq $idKey) -or ($nameKey -and ((ConvertTo-Label $_.Title) -eq $nameKey -or (ConvertTo-Label $_.Name) -eq $nameKey))) })[0]
        if (-not $r) { Write-Warning ("実績の「{0}」に当たる設計書が見つかりません" -f (@($a.'ID', $a.'名前') | Where-Object { $_ }) -join ' '); continue }
        $matched.Add([pscustomobject]@{ Result = $r; Actual = $d })
    }
    foreach ($c in $coefList) {
        $bucket = @($matched | Where-Object { $_.Result.Kind -eq $c.Kind -and $_.Result.Level -eq $c.Level })
        $avg = $null
        if ($bucket.Count -gt 0) {
            if ($c.PerItem) { $avg = ($bucket | ForEach-Object { if ($_.Result.Size -gt 0) { $_.Actual / $_.Result.Size } else { 0 } } | Measure-Object -Average).Average }
            else { $avg = ($bucket | Measure-Object Actual -Average).Average }
            $avg = [math]::Round($avg, 2)
        }
        $calibration.Add([pscustomobject]@{ Kind = $c.Kind; Level = $c.Level; Measure = $c.Measure; Limit = $c.Limit; PerItem = $c.PerItem; Days = $c.Days; Count = $bucket.Count; Average = $avg })
    }
    if ($OutCoefficients) {
        $encoding = 'UTF8'
        if ($PSVersionTable.PSVersion.Major -ge 6) { $encoding = 'utf8BOM' }
        # 実績の無い難易度は元の工数のまま。-Coefficients にそのまま渡せる（実績件数・元の工数の列は読むときには使わない）
        $calibration | ForEach-Object {
            $days = $_.Days
            if ($null -ne $_.Average) { $days = $_.Average }
            $limit = ''
            if ($null -ne $_.Limit) { $limit = $_.Limit }
            $unit = ''
            if ($_.PerItem) { $unit = '件' }
            [pscustomobject]@{ '種類' = $_.Kind; '難易度' = $_.Level; '基準' = $_.Measure; '上限' = $limit; '工数' = $days; '単位' = $unit; '実績件数' = $_.Count; '元の工数' = $_.Days }
        } | Export-Csv -LiteralPath $OutCoefficients -NoTypeInformation -Encoding $encoding
    }
}

if ($Summary) {
    [pscustomobject]@{
        TotalDays = [math]::Round($totalDays, 2); TotalMonths = [math]::Round($totalDays / $DaysPerMonth, 2)
        FpIndicative = $fpIndicative; FpEstimated = $fpEstimated; ILF = $ilf; EIF = $eif; EI = $ei; EO = $eo; EQ = $eq
        ByKind = $kindTotals.ToArray()
    }
}
elseif ($Baseline) { $changes }
else { $results }

Write-Host ''
Write-Host ('{0} ファイルを数えました（係数表: {1}）' -f $books.Count, [System.IO.Path]::GetFileName($Coefficients))
$table = $kindTotals | Format-Table @{ n = '種類'; e = { $_.Kind } }, @{ n = '難易度'; e = { $_.Level } }, @{ n = '数'; e = { $_.Count } }, @{ n = '規模'; e = { '{0:0.##}' -f $_.Size }; a = 'Right' }, @{ n = '工数（人日）'; e = { '{0:0.##}' -f $_.Days }; a = 'Right' } -AutoSize | Out-String
Write-Host $table.TrimEnd()
Write-Host ('合計 {0:0.##} 人日（{1:0.##} 人月、1 人月 = {2} 人日）' -f $totalDays, ($totalDays / $DaysPerMonth), $DaysPerMonth)
Write-Host ('FP（NESMA 概算法）{0} = 35 × ILF {1} + 15 × EIF {2}' -f $fpIndicative, $ilf, $eif)
Write-Host ('FP（NESMA 試算法）{0} = 7 × ILF {1} + 5 × EIF {2} + 4 × EI {3} + 5 × EO {4} + 4 × EQ {5}' -f $fpEstimated, $ilf, $eif, $ei, $eo, $eq)
if ($Baseline) {
    $byChange = { param($name) @($changes | Where-Object Change -eq $name) }
    $sumDays = { param($list) $v = ($list | Where-Object { $null -ne $_.ChangeDays } | Measure-Object ChangeDays -Sum).Sum; if ($null -eq $v) { 0 } else { $v } }
    $newList = & $byChange '新規'; $changedList = & $byChange '変更'
    Write-Host ''
    Write-Host ('前の版との差: 新規 {0} 件（{1:0.##} 人日）・変更 {2} 件（{3:0.##} 人日）・削除 {4} 件・変更なし {5} 件' -f $newList.Count, (& $sumDays $newList), $changedList.Count, (& $sumDays $changedList), (& $byChange '削除').Count, (& $byChange '変更なし').Count)
    Write-Host ('新規・変更の工数の目安 {0:0.##} 人日（変更は、新規の工数 × 変わった行の割合）' -f (& $sumDays $changes))
}
if ($Actuals) {
    Write-Host ''
    Write-Host '実績の工数と係数表（難易度ごとの平均）'
    $table = $calibration | Format-Table @{ n = '種類'; e = { $_.Kind } }, @{ n = '難易度'; e = { $_.Level } }, @{ n = '実績の件数'; e = { $_.Count }; a = 'Right' }, @{ n = '実績の平均'; e = { if ($null -ne $_.Average) { '{0:0.##}' -f $_.Average } else { '-' } }; a = 'Right' }, @{ n = '係数表'; e = { '{0:0.##}' -f $_.Days }; a = 'Right' }, @{ n = '単位'; e = { if ($_.PerItem) { '1 件あたり' } else { '1 つあたり' } } } -AutoSize | Out-String
    Write-Host $table.TrimEnd()
    if ($OutCoefficients) { Write-Host "実績から作った係数表を保存しました: $OutCoefficients （実績の無い難易度は元の工数のまま）" }
}

if ($OutXlsx) {
    $detailRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $days = ''
        if ($null -ne $r.Days) { $days = [double]$r.Days }
        $detailRows.Add(@{ Values = @($r.Kind, $r.Name, $r.Measures, [double]$r.Size, $r.Level, $days, ($r.EI + 0), ($r.EQ + 0), ($r.EO + 0), $r.File, $r.Location); Link = @{ Column = 1; Path = $r.Path; Location = ("'{0}'!{1}" -f $r.Sheet.Replace("'", "''"), $r.Cell) } })
    }
    $sumRows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $kindTotals) { $sumRows.Add(@{ Values = @($s.Kind, $s.Level, [int]$s.Count, [double]$s.Size, [double]$s.Days) }) }
    $sumRows.Add(@{ Values = @('合計（人日）', '', '', '', [double]$totalDays) })
    $sumRows.Add(@{ Values = @('合計（人月）', '', '', '', [math]::Round($totalDays / $DaysPerMonth, 2)) })
    $sumRows.Add(@{ Values = @('FP（NESMA 概算法）', '', '', '', [double]$fpIndicative) })
    $sumRows.Add(@{ Values = @('FP（NESMA 試算法）', '', '', '', [double]$fpEstimated) })
    foreach ($p in @(@('ILF（自システムのテーブル）', $ilf), @('EIF（ほかのシステムのテーブル）', $eif), @('EI（外部入力）', $ei), @('EO（外部出力）', $eo), @('EQ（外部照会）', $eq))) {
        $sumRows.Add(@{ Values = @($p[0], '', [double]$p[1], '', '') })
    }
    $coefRows = New-Object System.Collections.Generic.List[object]
    foreach ($c in $coefList) {
        $limit = ''
        if ($null -ne $c.Limit) { $limit = [double]$c.Limit }
        $unit = '1 つあたり'
        if ($c.PerItem) { $unit = '1 件あたり' }
        $coefRows.Add(@{ Values = @($c.Kind, $c.Level, $c.Measure, $limit, [double]$c.Days, $unit) })
    }
    $sheets = @(
        @{ Name = '集計'; Columns = @(@{ Header = '種類'; Width = 30 }, @{ Header = '難易度'; Width = 9 }, @{ Header = '数'; Width = 8 }, @{ Header = '規模'; Width = 10 }, @{ Header = '工数（人日）'; Width = 12 }); Rows = $sumRows },
        @{ Name = '規模一覧'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '名前'; Width = 30 }, @{ Header = '規模'; Width = 28 }, @{ Header = '基準の値'; Width = 9 },
                @{ Header = '難易度'; Width = 8 }, @{ Header = '工数（人日）'; Width = 11 }, @{ Header = 'EI'; Width = 5 }, @{ Header = 'EQ'; Width = 5 }, @{ Header = 'EO'; Width = 5 },
                @{ Header = 'ファイル'; Width = 40 }, @{ Header = '場所'; Width = 20 }); Rows = $detailRows },
        @{ Name = '係数表'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '難易度'; Width = 8 }, @{ Header = '基準'; Width = 10 }, @{ Header = '上限'; Width = 8 }, @{ Header = '工数（人日）'; Width = 11 }, @{ Header = '単位'; Width = 11 }); Rows = $coefRows })
    if ($Baseline) {
        $changeRows = New-Object System.Collections.Generic.List[object]
        foreach ($c in $changes) {
            $cd = ''
            if ($null -ne $c.ChangeDays) { $cd = [double]$c.ChangeDays }
            $changeRows.Add(@{ Values = @($c.Kind, $c.Name, $c.Change, $c.Detail, $cd, $c.File, $c.Location); Link = @{ Column = 1; Path = $c.Path; Location = ("'{0}'!{1}" -f $c.Sheet.Replace("'", "''"), $c.Cell) } })
        }
        $sheets += @{ Name = '前の版との差'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '名前'; Width = 30 }, @{ Header = '変更'; Width = 8 }, @{ Header = '内容'; Width = 60 }, @{ Header = '工数の目安（人日）'; Width = 12 }, @{ Header = 'ファイル'; Width = 40 }, @{ Header = '場所'; Width = 20 }); Rows = $changeRows }
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets $sheets
        Write-Host "結果を保存しました: $OutXlsx"
    }
    catch { Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message) }
}

# .SYNOPSIS
#     設計書（Excel）から画面・テーブル・IF などの規模を数え、係数表で工数の目安と FP（ファンクションポイント）を出します（Excel 不要）。
#
# .DESCRIPTION
#     設計書のフォルダを読み、規模の数え方（CSV）に沿って次のように数えます。
#
#       画面       画面設計書ごとに、画面項目定義の項目数とイベント定義のイベント数
#       テーブル   テーブル定義書のシートごとに、カラム定義のカラム数
#       IF         IF 定義書ごとに、レコード定義の項目数
#       帳票・バッチ・メッセージ・コード  それぞれの設計書の項目・ステップ・行の数
#
#     表は見出し（「項目名」「項目ID」など）がそろっている行を探し、その下で「数える列」に値のある行を、
#     次の節（「■」で始まる行）まで数えます。見出しが見つからない設計書や、0 件の表は警告します。
#
#     係数表（CSV）で、規模（項目数など）から難易度（簡易・普通・複雑）を決め、1 つあたりの工数を掛けて工数の目安を出します。
#     あわせて FP を NESMA の概算法（35 × ILF + 15 × EIF）と試算法（7 × ILF + 5 × EIF + 4 × EI + 5 × EO + 4 × EQ）で出します。
#     テーブルを ILF（-ExternalTables に一致するものは EIF）、画面のイベントのうち使用テーブルに C / U / D のあるものを EI、
#     R だけのものを EQ、送信の IF と帳票を EO、受信の IF を EI として数えます。
#
#     結果は対象（画面・テーブルなど）ごとに 1 行で出力し、集計はコンソールに表示します（-OutXlsx で保存）。
#
#     -Baseline で前の版を渡すと、対象ごとの新規・変更・削除と、項目・カラムなどの行の追加・変更・削除の数を出します
#     （仕様変更の追加見積もりの根拠）。-Actuals で実績の工数を渡すと、係数表の難易度ごとに実績の平均と比べ、
#     -OutCoefficients で実績から作った係数表を保存します。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER Rules
#     規模の数え方の CSV（種類・ファイル名・シート名・単位・見出し・数える列・数えるもの・名前・ID・CRUD列）。
#     既定は tools\estimate\規模の数え方.csv。社内の設計書の書式に合わせて直して使えます。
#
# .PARAMETER Coefficients
#     係数表の CSV（種類・難易度・基準・上限・工数・単位）。上から順に「基準の値が上限以下」の行の工数（人日）を使います。
#     単位が「件」の行は、基準の値（件数）に工数を掛けます。既定は tools\estimate\係数表_サンプル.csv（数値は例です）。
#
# .PARAMETER ExternalTables
#     ほかのシステムが管理するテーブル（FP の EIF）のテーブル ID・名前の正規表現。例: '^M_CUSTOMER'
#
# .PARAMETER DaysPerMonth
#     人月に換算するときの 1 人月の日数。既定は 20。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は数えなくなります）。
#
# .PARAMETER Summary
#     対象ごとの一覧の代わりに、集計（合計の人日・人月、FP、種類・難易度ごとの数と工数）を 1 つのオブジェクトで出力します。
#
# .PARAMETER Baseline
#     前の版の設計書のフォルダまたはファイル。対象（画面・テーブルなど）を種類と ID（無ければ名前）で対応付け、
#     新規・変更・削除・変更なしと、数えた行（項目・カラム・イベントなど）を「数える列」の値で対応付けた追加・変更・削除の数を出力します。
#     工数の目安は、新規は係数表の工数、変更は係数表の工数 × 変わった行の割合です（削除は 0）。番号の列（No など）は比べません。
#
# .PARAMETER Actuals
#     実績の工数の CSV（ID または 名前、工数（人日）。種類の列があれば種類も合わせる）。設計書の対象と対応付け、
#     係数表の難易度ごとに実績の平均を出します（単位が「件」の行は 1 件あたりの平均）。
#
# .PARAMETER OutCoefficients
#     -Actuals の実績の平均で工数を置き換えた係数表を、CSV（UTF-8 BOM 付き）に保存します。
#     実績の無い難易度は元の工数のままです。そのまま -Coefficients に渡せます。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（集計・規模一覧・使った係数表。規模一覧の名前をクリックすると設計書の表が開きます）。
#     -Baseline を付けたときは「前の版との差」のシートも入れます。
#
# .EXAMPLE
#     .\Measure-DesignSize.ps1 D:\設計書
#
# .EXAMPLE
#     # 自社の係数表で工数を出し、顧客マスタは他システム管理として FP を出す
#     .\Measure-DesignSize.ps1 D:\設計書 -Coefficients .\係数表.csv -ExternalTables '^M_CUSTOMER' -OutXlsx D:\規模.xlsx
#
# .EXAMPLE
#     # 前の版（納品物）と比べて、仕様変更の規模を出す
#     .\Measure-DesignSize.ps1 D:\設計書 -Baseline D:\納品\v1.0 -OutXlsx D:\変更規模.xlsx
#
# .EXAMPLE
#     # 前のプロジェクトの実績の工数から係数表を作る
#     .\Measure-DesignSize.ps1 D:\前のプロジェクト\設計書 -Actuals D:\実績工数.csv -OutCoefficients D:\係数表_実績.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     工数と FP は見積もりの目安です。係数表の数値は社内の実績に合わせて直してください。
