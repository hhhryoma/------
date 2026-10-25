[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $DesignPath,

    [string] $SizeCsv,

    [string[]] $ReviewLog,

    [string] $Limits,

    [string] $Headers,

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$sizeScript = Join-Path $PSScriptRoot 'Measure-DesignSize.ps1'

if (-not $Limits) { $Limits = Join-Path $PSScriptRoot 'review-quality/基準値_サンプル.csv' }
$defaultHeaders = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv'))
if (-not $Headers) { $Headers = $defaultHeaders }
$zoneCsv = Join-Path $PSScriptRoot 'review-quality/ゾーン分析.csv'
foreach ($name in @('Limits', 'Headers', 'SizeCsv', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($DesignPath -and $SizeCsv) { throw '-DesignPath と -SizeCsv は、どちらか一方を指定してください' }

Update-TypeData -TypeName 'ReviewQuality.Document' -DefaultDisplayPropertySet Document, Findings, Open, Overdue, Density, Judgement -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 日付は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 対応が終わった・指摘として数えない状態
$closedPattern = '完了|済|クローズ|close|done|対応不要|不要|対処なし|終了'
$excludedPattern = '取下げ|取り下げ|取消|取り消|却下|対象外|重複|誤指摘|キャンセル|cancel'

# ---------------------------------------------------------------- 設定


# 見出しの候補。-Headers に無い項目（レビュー記録の工数など）は、既定の CSV の候補を使う
$candidates = @{}
$headerFiles = @($defaultHeaders)
if ($Headers -ne $defaultHeaders) { $headerFiles += $Headers }
foreach ($file in $headerFiles) {
    if ($file -eq $defaultHeaders -and $Headers -ne $defaultHeaders -and -not (Test-Path -LiteralPath $file)) { continue }
    foreach ($r in (Read-CsvFile $file @('項目', '見出し') '指摘管理表の見出し')) {
        $field = ([string]$r.'項目').Trim()
        if ($field) { $candidates[$field] = @(([string]$r.'見出し') -split '\|' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ }) }
    }
}
$fields = @('設計書', '指摘内容', '分類', '重要度', '状態', '対応要否', '起票日', '期限', '完了日', '対応者', '工程')
$logFields = @('設計書', '工数', '時間', '人数', '参加者')
foreach ($f in ($fields + $logFields)) { if (-not $candidates.ContainsKey($f)) { $candidates[$f] = @() } }
if ($candidates['指摘内容'].Count -eq 0) { throw "指摘管理表の見出しに「指摘内容」の行がありません: $Headers" }

# 基準値。指標は「指摘密度」「レビュー工数密度」。指標が空の行と、「指標」の列が無い CSV（以前の形）は、指摘密度の基準値とみなす。
# New-QualityLimits.ps1 の CSV にあるテスト密度・バグ密度など、ほかの指標の行は読まない（分母が同じ名前でも、ほかの指標の基準値で判定しないように）
$limitRows = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Limits @('分母', '下限', '上限') '基準値')) {
    $unit = ([string]$r.'分母').Trim()
    if (-not $unit) { continue }
    $metricText = ''
    if ($r.PSObject.Properties['指標']) { $metricText = ([string]$r.'指標').Trim() }
    if ($metricText -match '工数') { $metric = 'Effort' }
    elseif (-not $metricText -or $metricText -match '指摘') { $metric = 'Findings' }
    else { continue }
    $process = ''
    if ($r.PSObject.Properties['工程']) { $process = ([string]$r.'工程').Trim() }
    $low = $null; $high = $null
    if (([string]$r.'下限').Trim()) { $low = [double]$r.'下限' }
    if (([string]$r.'上限').Trim()) { $high = [double]$r.'上限' }
    $limitRows.Add([pscustomobject]@{ Metric = $metric; Unit = $unit; Process = $process; Low = $low; High = $high })
}

function ConvertTo-DocKey {
    #     設計書の名前を比べる形（フォルダと拡張子を除き、全角・半角と空白をそろえる）
    param([string] $Name)
    $n = Get-LeafName $Name
    if ($n -match '\.(xlsx|xlsm|xls|xlsb|pdf|docx?|pptx?)$') { $n = [System.IO.Path]::GetFileNameWithoutExtension($n) }
    return (ConvertTo-Label $n).ToUpperInvariant()
}

# ---------------------------------------------------------------- 指摘管理表を読む


$findings = New-Object System.Collections.Generic.List[object]
$excluded = 0
$lastDoc = @{}   # 表ごとの上の行の設計書

function Add-Finding {
    #     指摘を 1 件足す。Table: 表（ファイル・シート・見出しの行）。日付は、年の無い日付・和暦・曜日つきも読む（ConvertTo-DateValue）
    param([hashtable] $Values, $Place, [string] $Table)
    if (-not ([string]$Values['指摘内容']).Trim()) { return }
    # 設計書はグループの最初の行だけに書くことがある（縦に結合したセルも左上だけに値がある）ので、空なら同じ表の上の行の設計書
    $doc = ([string]$Values['設計書']).Trim()
    if ($doc) { $lastDoc[$Table] = $doc } elseif ($lastDoc.ContainsKey($Table)) { $doc = $lastDoc[$Table] }
    $status = ([string]$Values['状態']).Trim()
    $required = ([string]$Values['対応要否']).Trim()
    if ($status -match $excludedPattern) { $script:excluded++; return }
    $closedOn = ConvertTo-DateValue $Values['完了日'] $AsOf
    $closed = (Test-ClosedStatus $status $closedPattern) -or ($required -match '^(不要|否|×|-)$') -or ((-not $status) -and $closedOn)
    $due = ConvertTo-DateValue $Values['期限'] $AsOf
    if (-not $doc) { $doc = '（設計書なし）' }
    $findings.Add([pscustomobject]@{
        Document = $doc; Key = (ConvertTo-DocKey $doc); Content = ([string]$Values['指摘内容']).Trim()
        Category = ([string]$Values['分類']).Trim(); Severity = ([string]$Values['重要度']).Trim(); Status = $status
        Owner = ([string]$Values['対応者']).Trim(); Process = ([string]$Values['工程']).Trim()
        Opened = (ConvertTo-DateValue $Values['起票日'] $AsOf); Due = $due; ClosedOn = $closedOn; Closed = [bool]$closed
        Overdue = ((-not $closed) -and $due -and $due -lt $AsOf); Place = $Place
    })
}

function Find-FirstLabel {
    param($Labels, [string[]] $Names)
    foreach ($n in $Names) { if ($Labels -contains $n) { return $n } }
    return $null
}

$excelFiles = New-Object System.Collections.Generic.List[string]
$csvFiles = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            if (@(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.xls[xmb]?$' -and -not $_.Name.StartsWith('~$') }).Count -gt 0) { $excelFiles.Add($item.FullName) }
            foreach ($c in @(Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.csv' -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { $csvFiles.Add($c) }
        }
        elseif ($item.Extension -eq '.csv') { $csvFiles.Add($item) }
        else { $excelFiles.Add([System.Management.Automation.WildcardPattern]::Escape($item.FullName)) }
    }
}

# 読み取り結果（どのファイル・シートをどの見出しで読んだか）
$readLog = New-Object System.Collections.Generic.List[object]
if ($excelFiles.Count -gt 0) {
    # 縦に結合したセル（状態・期限など）は、結合した各行の値として読む（Read-HeaderRecords と同じ）
    $searchArgs = @{ Pattern = '\S'; Path = $excelFiles.ToArray(); Target = 'Value'; FillMerged = $true }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
    $books = ConvertTo-DesignBooks $hits
    $ownResults = New-Object System.Collections.Generic.List[string]
    foreach ($book in $books.Values) {
        # 前に -OutXlsx で保存したこのツールの結果（設計書別・未完了・読み取り結果のシートがあるブック）は読まない（設計書別・未完了のシートを指摘に数えてしまう）
        if ($book.Sheets.Contains('設計書別') -and $book.Sheets.Contains('未完了') -and $book.Sheets.Contains('読み取り結果')) { $ownResults.Add($book.File); continue }
        $tableFound = $false
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # 見出しの行: 指摘内容と、設計書・状態・重要度・分類のどれかの見出しがある行
            $map = $null
            foreach ($row in $grid.Rows) {
                $labels = @($grid.Labels[$row].Values)
                $m = @{}
                foreach ($f in $fields) { $l = Find-FirstLabel $labels $candidates[$f]; if ($l) { $m[$f] = $l } }
                if ($m['指摘内容'] -and @(@('設計書', '状態', '重要度', '分類') | Where-Object { $m[$_] }).Count -gt 0) { $map = $m; break }
            }
            if (-not $map) { continue }
            $tableFound = $true
            # 一覧の表として読む（「■索引: …」のように「■」で始まる指摘内容の行で表を終えない）
            foreach ($t in (Find-DesignTables $grid @($map.Values) $map['指摘内容'] -List)) {
                Add-ReadLog $readLog '指摘管理表' $book.File $sheetName $t.HeaderRow $t.Rows.Count $map $fields ''
                foreach ($dr in $t.Rows) {
                    $v = @{}
                    foreach ($f in $map.Keys) { $v[$f] = [string]$dr.Cells[$t.ColumnOf[$map[$f]]] }
                    $col = $t.ColumnOf[$map['指摘内容']]
                    Add-Finding $v ([pscustomobject]@{ File = $book.File; Path = $book.Path; Sheet = $sheetName; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $col), $dr.Row) }) ('{0}|{1}|{2}' -f $book.Path, $sheetName, $t.HeaderRow)
                }
            }
        }
        if (-not $tableFound) { Add-ReadLog $readLog '指摘管理表' $book.File '' $null 0 $null $fields '表が見つかりません（必要な見出しの行がありません）' }
    }
    if ($ownResults.Count -gt 0) { Write-Warning ('このツールの結果のブック（設計書別・未完了・読み取り結果のシートがあるもの）は、指摘管理表として読みません: {0}' -f ($ownResults -join '・')) }
}
foreach ($c in $csvFiles) {
    $rows = Read-InputCsv $c.FullName '指摘管理表' $readLog $fields
    if ($rows.Count -eq 0) { continue }
    $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $map = @{}
    foreach ($f in $fields) { foreach ($n in $names) { if (-not $map[$f] -and $candidates[$f] -contains (ConvertTo-Label $n)) { $map[$f] = $n } } }
    if (-not $map['指摘内容']) {
        Write-Warning ("{0}: 指摘内容の列が見つかりません" -f $c.Name)
        Add-ReadLog $readLog '指摘管理表' $c.Name '' $null 0 $null $fields '表が見つかりません（指摘内容の列がありません）'
        continue
    }
    $labelMap = @{}
    foreach ($f in $map.Keys) { $labelMap[$f] = ConvertTo-Label $map[$f] }
    Add-ReadLog $readLog '指摘管理表' $c.Name '' 1 $rows.Count $labelMap $fields ''
    $line = 1
    foreach ($r in $rows) {
        $line++
        $v = @{}
        foreach ($f in $map.Keys) { $v[$f] = [string]$r.($map[$f]) }
        Add-Finding $v ([pscustomobject]@{ File = $c.Name; Path = $c.FullName; Sheet = ''; Cell = ('{0} 行目' -f $line) }) $c.FullName
    }
}

# ---------------------------------------------------------------- レビュー記録を読む（レビュー工数）


$efforts = New-Object System.Collections.Generic.List[object]

function ConvertTo-Hours {
    #     時間の文字（1.5・1.5h・2時間・90分・1時間30分・1時間半・3人時・1:30・10:00～11:30）を時間の数にする。読めなければ $null。
    #     時刻の書式のセルは Search-Excel が「01:30:00」の形で返すので、秒も読む
    #     Minutes: 数字だけのときに分とみなす（見出しが「レビュー時間(分)」など）
    param([string] $Text, [bool] $Minutes)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return $null }
    $m = [regex]::Match($t, '^(\d{1,2}):(\d{2})(?::\d{2})?\s*[~\-〜]\s*(\d{1,2}):(\d{2})(?::\d{2})?$')
    if ($m.Success) {
        $span = ([int]$m.Groups[3].Value * 60 + [int]$m.Groups[4].Value) - ([int]$m.Groups[1].Value * 60 + [int]$m.Groups[2].Value)
        if ($span -gt 0) { return $span / 60.0 }
        return $null
    }
    $m = [regex]::Match($t, '^(\d+):(\d{2})(?::(\d{2}))?$')
    if ($m.Success) {
        $hours = [int]$m.Groups[1].Value + [int]$m.Groups[2].Value / 60.0
        if ($m.Groups[3].Success) { $hours += [int]$m.Groups[3].Value / 3600.0 }
        return $hours
    }
    $m = [regex]::Match($t, '^(?:(\d+(?:\.\d+)?)\s*(?:人時|時間|hrs|hr|h)(半)?)?\s*(?:(\d+(?:\.\d+)?)\s*分)?$', 'IgnoreCase')
    if ($m.Success -and ($m.Groups[1].Success -or $m.Groups[3].Success)) {
        $hours = 0.0
        if ($m.Groups[1].Success) { $hours += [double]::Parse($m.Groups[1].Value, [System.Globalization.CultureInfo]::InvariantCulture) }
        if ($m.Groups[2].Success) { $hours += 0.5 }
        if ($m.Groups[3].Success) { $hours += [double]::Parse($m.Groups[3].Value, [System.Globalization.CultureInfo]::InvariantCulture) / 60 }
        return $hours
    }
    $n = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
        if ($Minutes) { return $n / 60 }
        return $n
    }
    return $null
}

function Get-PersonCount {
    #     参加人数（「3」「3名」）か、参加者の名前の数（「山田、佐藤」）
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return 0 }
    $m = [regex]::Match($t, '^(\d+)\s*(?:名|人)?$')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return @($t -split '[、,/・;\r\n↵]+' | Where-Object { $_.Trim() }).Count
}

$lastLogDoc = @{}   # 表ごとの上の行の設計書
$skippedLogs = New-Object System.Collections.Generic.List[string]   # 数えなかった行（場所と値）

function Add-Effort {
    #     レビュー 1 回分の工数（人時）を足す。工数が無ければ、時間 × 人数（参加者の数）。設計書が空なら、同じ表（Table）の上の行の設計書。
    #     工数・時間が書いてあるのに読めない行と、設計書の無い行は、数えずに skippedLogs に足す（工数・時間が空の行は、レビューの記録ではないので足さない）
    param([hashtable] $Values, [hashtable] $Labels, $Place, [string] $Table)
    $doc = ([string]$Values['設計書']).Trim()
    if ($doc) { $lastLogDoc[$Table] = $doc } elseif ($lastLogDoc.ContainsKey($Table)) { $doc = $lastLogDoc[$Table] }
    $where = ('{0} {1}' -f $Place.File, $Place.Sheet).Trim()
    if ($Place.Row) { $where += (' {0} 行目' -f $Place.Row) }
    $hours = ConvertTo-Hours $Values['工数'] ($Labels['工数'] -match '分')
    if ($null -eq $hours) {
        $time = ConvertTo-Hours $Values['時間'] ($Labels['時間'] -match '分')
        if ($null -eq $time) {
            $text = @(@($Values['工数'], $Values['時間']) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
            if ($text.Count -gt 0) { $skippedLogs.Add(('{0}「{1}」' -f $where, ($text -join '・'))) }
            return
        }
        $persons = 0
        if (([string]$Values['人数']).Trim()) { $persons = Get-PersonCount $Values['人数'] }
        if ($persons -eq 0) { $persons = Get-PersonCount $Values['参加者'] }
        if ($persons -eq 0) { $persons = 1 }
        $hours = $time * $persons
    }
    if (-not $doc) { $skippedLogs.Add(('{0}（設計書が無い）' -f $where)); return }
    $efforts.Add([pscustomobject]@{ Document = $doc; Key = (ConvertTo-DocKey $doc); Hours = [double]$hours; Place = $Place })
}

if ($ReviewLog) {
    # 記録票の見出し（参加者の右に並べた名前は、次の見出しの前まで）
    $logHeadings = @($logFields | ForEach-Object { $candidates[$_] })
    $logExcel = New-Object System.Collections.Generic.List[string]
    $logCsv = New-Object System.Collections.Generic.List[object]
    foreach ($p in $ReviewLog) {
        $items = @(Get-PathItem $p)
        if ($items.Count -eq 0) { Write-Warning "レビュー記録が見つかりません: $p"; continue }
        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                if (@(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.xls[xmb]?$' -and -not $_.Name.StartsWith('~$') }).Count -gt 0) { $logExcel.Add($item.FullName) }
                foreach ($c in @(Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.csv' -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { $logCsv.Add($c) }
            }
            elseif ($item.Extension -eq '.csv') { $logCsv.Add($item) }
            else { $logExcel.Add([System.Management.Automation.WildcardPattern]::Escape($item.FullName)) }
        }
    }
    if ($logExcel.Count -gt 0) {
        $searchArgs = @{ Pattern = '\S'; Path = $logExcel.ToArray(); Target = 'Value' }
        if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
        if ($NoExcel) { $searchArgs['NoExcel'] = $true }
        $books = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
        foreach ($book in $books.Values) {
            $logFound = $false
            foreach ($sheetName in $book.Sheets.Keys) {
                $grid = $book.Sheets[$sheetName]
                # 表の形: 設計書と、工数か時間の見出しがある行の下を、1 行 1 回のレビューとして読む
                $map = $null
                foreach ($row in $grid.Rows) {
                    $labels = @($grid.Labels[$row].Values)
                    $m = @{}
                    foreach ($f in $logFields) { $l = Find-FirstLabel $labels $candidates[$f]; if ($l) { $m[$f] = $l } }
                    if ($m['設計書'] -and ($m['工数'] -or $m['時間'])) { $map = $m; break }
                }
                if ($map) {
                    $logFound = $true
                    # 設計書はグループの最初の行だけに書くことがあるので、設計書か工数・時間のどれかに値のある行を読む。
                    # 表の見つけ方は同じなので、どの列で読んだ表も同じ順に並ぶ
                    $keyTables = @(@('設計書', '工数', '時間') | Where-Object { $map[$_] } | ForEach-Object { , (Find-DesignTables $grid @($map.Values) $map[$_]) })
                    for ($ti = 0; $ti -lt $keyTables[0].Count; $ti++) {
                        $t = $keyTables[0][$ti]
                        $rowOf = @{}
                        foreach ($keyed in $keyTables) { foreach ($dr in $keyed[$ti].Rows) { $rowOf[$dr.Row] = $dr } }
                        $rowNumbers = [int[]]@($rowOf.Keys)
                        [Array]::Sort($rowNumbers)
                        Add-ReadLog $readLog 'レビュー記録' $book.File $sheetName $t.HeaderRow $rowNumbers.Count $map $logFields ''
                        foreach ($rn in $rowNumbers) {
                            $dr = $rowOf[$rn]
                            $v = @{}
                            foreach ($f in $map.Keys) { $v[$f] = [string]$dr.Cells[$t.ColumnOf[$map[$f]]] }
                            Add-Effort $v $map ([pscustomobject]@{ File = $book.File; Sheet = $sheetName; Row = $dr.Row }) ('{0}|{1}|{2}' -f $book.Path, $sheetName, $t.HeaderRow)
                        }
                    }
                    continue
                }
                # 記録票の形: 「対象成果物」「レビュー時間」などの見出しの右の値を、1 シート 1 回のレビューとして読む
                $v = @{}; $found = @{}
                foreach ($f in $logFields) {
                    foreach ($cand in $candidates[$f]) {
                        $cell = Find-LabelCell $grid $cand
                        if (-not ($cell -and $cell.Value)) { continue }
                        $v[$f] = $cell.Value
                        # 参加者は、右に 1 人ずつ並べて書くことがある。同じ行の右にある見出し（欠席者・場所・レビュー時間など）の前までを読む
                        if ($f -eq '参加者') {
                            $names = New-Object System.Collections.Generic.List[string]
                            foreach ($col in @($grid.Cols[$cell.Row] | Where-Object { $_ -ge $cell.Col })) {
                                $label = $grid.Labels[$cell.Row][$col]
                                if ($logHeadings -contains $label -or $label -match '欠席|場所') { break }
                                $names.Add($grid.Cells[$cell.Row][$col])
                            }
                            $v[$f] = $names -join '、'
                        }
                        $found[$f] = $cand
                        break
                    }
                }
                if ($v['設計書'] -and ($v['工数'] -or $v['時間'])) {
                    $logFound = $true
                    Add-ReadLog $readLog 'レビュー記録' $book.File $sheetName $null 1 $found $logFields '記録票の形（見出しの右の値）'
                    Add-Effort $v $found ([pscustomobject]@{ File = $book.File; Sheet = $sheetName; Row = 0 }) ('{0}|{1}' -f $book.Path, $sheetName)
                }
            }
            if (-not $logFound) { Add-ReadLog $readLog 'レビュー記録' $book.File '' $null 0 $null $logFields '表・記録票が見つかりません（設計書と、工数か時間の見出しがありません）' }
        }
    }
    foreach ($c in $logCsv) {
        $rows = Read-InputCsv $c.FullName 'レビュー記録' $readLog $logFields
        if ($rows.Count -eq 0) { continue }
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $map = @{}
        foreach ($f in $logFields) { foreach ($n in $names) { if (-not $map[$f] -and $candidates[$f] -contains (ConvertTo-Label $n)) { $map[$f] = $n } } }
        if (-not $map['設計書'] -or -not ($map['工数'] -or $map['時間'])) {
            Write-Warning ("{0}: 設計書と、工数か時間の列が見つかりません" -f $c.Name)
            Add-ReadLog $readLog 'レビュー記録' $c.Name '' $null 0 $null $logFields '表が見つかりません（設計書と、工数か時間の列がありません）'
            continue
        }
        $labels = @{}
        foreach ($f in $map.Keys) { $labels[$f] = ConvertTo-Label $map[$f] }
        Add-ReadLog $readLog 'レビュー記録' $c.Name '' 1 $rows.Count $labels $logFields ''
        $line = 1
        foreach ($r in $rows) {
            $line++
            $v = @{}
            foreach ($f in $map.Keys) { $v[$f] = [string]$r.($map[$f]) }
            Add-Effort $v $labels ([pscustomobject]@{ File = $c.Name; Sheet = ''; Row = $line }) $c.FullName
        }
    }
    if ($skippedLogs.Count -gt 0) {
        $examples = @($skippedLogs | Select-Object -First 3) -join '、'
        if ($skippedLogs.Count -gt 3) { $examples += ' など' }
        Write-Warning ('レビュー記録の {0} 行は、工数・時間か設計書を読めないため数えていません: {1}' -f $skippedLogs.Count, $examples)
    }
    if ($efforts.Count -eq 0) { Write-Warning 'レビュー記録から工数を読めませんでした（設計書と、工数か時間の見出しを確かめてください）' }
}

# ---------------------------------------------------------------- 規模（分母）


$sizes = [ordered]@{}   # 設計書の比較用の名前 → @{ Name; Size; Unit }
if ($SizeCsv) {
    $rows = Read-CsvFile $SizeCsv @() '規模の CSV'
    if ($rows.Count -gt 0) {
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $nameCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(ファイル|設計書|成果物|ファイル名|名前|対象)$' })[0]
        $sizeCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(ページ数|規模|サイズ|件数|ページ)$' })[0]
        if (-not $nameCol -or -not $sizeCol) { throw "規模の CSV には、ファイル（設計書）の列と、ページ数（規模）の列が必要です: $SizeCsv" }
        foreach ($r in $rows) {
            # Export-ExcelToPdf.ps1 -LogPath の CSV は、失敗した行を使わない（-SkipExisting で「スキップ」した行にも、既存の PDF のページ数がある）
            if ($r.PSObject.Properties['結果'] -and ([string]$r.'結果').Trim() -eq '失敗') { continue }
            $n = 0.0
            if (-not [double]::TryParse(([string]$r.$sizeCol).Trim(), [ref]$n)) { continue }
            $k = ConvertTo-DocKey $r.$nameCol
            if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = @(([string]$r.$nameCol).Trim() -split '[\\/]')[-1]; Size = 0.0; Unit = $sizeCol } }
            $sizes[$k].Size += $n
        }
    }
}
elseif ($DesignPath) {
    $sizeArgs = @{ Path = $DesignPath }
    if ($NoRecurse) { $sizeArgs['NoRecurse'] = $true }
    if ($NoExcel) { $sizeArgs['NoExcel'] = $true }
    foreach ($item in @(& $sizeScript @sizeArgs 6>$null 3>$null)) {
        $k = ConvertTo-DocKey $item.File
        if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = $item.File; Size = 0.0; Unit = '規模' } }
        $sizes[$k].Size += [double]$item.Size
    }
}

function Find-SizeKey {
    #     指摘の設計書の名前に対応する規模（同じ名前、またはどちらかがもう一方を含む名前がただ 1 つ）
    param([string] $Key)
    if ($sizes.Contains($Key)) { return $Key }
    $hit = @($sizes.Keys | Where-Object { $_.Contains($Key) -or $Key.Contains($_) })
    if ($hit.Count -eq 1) { return $hit[0] }
    return $null
}

# ---------------------------------------------------------------- 集計


function Format-Breakdown {
    #     「高 3・中 5」のような内訳（多い順）
    param($Items, [string] $Property)
    $groups = @($Items | Where-Object { $_.$Property } | Group-Object $Property | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)
    return (@($groups | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '・')
}

function New-Doc {
    param([string] $Name, $SizeKey)
    return [pscustomobject]@{ Name = $Name; SizeKey = $SizeKey; Items = (New-Object System.Collections.Generic.List[object]); Effort = $null; Reviews = 0 }
}

$docs = [ordered]@{}
foreach ($f in $findings) {
    $sk = Find-SizeKey $f.Key
    $k = $f.Key
    if ($sk) { $k = $sk }
    if (-not $docs.Contains($k)) { $docs[$k] = New-Doc $f.Document $sk }
    $docs[$k].Items.Add($f)
}
foreach ($k in $sizes.Keys) { if (-not $docs.Contains($k)) { $docs[$k] = New-Doc $sizes[$k].Name $k } }
foreach ($e in $efforts) {
    # 規模の名前、指摘の設計書の名前（同じ名前か、どちらかがもう一方を含む名前がただ 1 つ）の順に対応付ける
    $sk = Find-SizeKey $e.Key
    $k = $sk
    if (-not $k) {
        if ($docs.Contains($e.Key)) { $k = $e.Key }
        else {
            $hit = @($docs.Keys | Where-Object { $_.Contains($e.Key) -or $e.Key.Contains($_) })
            if ($hit.Count -eq 1) { $k = $hit[0] } else { $k = $e.Key }
        }
    }
    if (-not $docs.Contains($k)) { $docs[$k] = New-Doc $e.Document $sk }
    if ($null -eq $docs[$k].Effort) { $docs[$k].Effort = 0.0 }
    $docs[$k].Effort += $e.Hours
    $docs[$k].Reviews++
}

function Get-Judgement {
    #     基準値（指標・分母・工程）と比べて「少ない」「基準内」「多い」。基準値が無ければ ''
    param([string] $Metric, [string] $Unit, [string] $Process, [double] $Value)
    $limit = @($limitRows | Where-Object { $_.Metric -eq $Metric -and $_.Unit -eq $Unit -and $Process -and $_.Process -eq $Process })[0]
    if (-not $limit) { $limit = @($limitRows | Where-Object { $_.Metric -eq $Metric -and $_.Unit -eq $Unit -and -not $_.Process })[0] }
    if (-not $limit) { return '' }
    if ($null -ne $limit.Low -and $Value -lt $limit.Low) { return '少ない' }
    if ($null -ne $limit.High -and $Value -gt $limit.High) { return '多い' }
    return '基準内'
}

$zones = @{}
if ($ReviewLog) { $zones = Read-ZoneTable $zoneCsv '工数密度' '指摘密度' }

$results = New-Object System.Collections.Generic.List[object]
foreach ($k in $docs.Keys) {
    $d = $docs[$k]
    $items = $d.Items
    $size = $null; $unit = ''; $density = $null; $judgement = ''
    $effortDensity = $null; $effortJudgement = ''; $zone = $null
    if ($d.SizeKey) { $size = $sizes[$d.SizeKey].Size; $unit = $sizes[$d.SizeKey].Unit }
    # 基準値に使う工程は、指摘の多い工程（同じ数なら先に出てきた工程。Group-Object は PowerShell 7 では工程の名前の順に並べるため、出てきた順のまま並べる Group-InOrder でまとめる）
    $process = ''; $most = 0
    foreach ($g in @($items | Where-Object { $_.Process } | Group-InOrder Process)) { if ($g.Count -gt $most) { $most = $g.Count; $process = $g.Name } }
    # レビューしたか: 指摘かレビュー記録があればした。レビュー記録を読んだときだけ、どちらも無い設計書を未レビューとする（読まなければ分からない）
    $wasReviewed = $null
    if ($items.Count -gt 0 -or $d.Reviews -gt 0) { $wasReviewed = $true } elseif ($ReviewLog) { $wasReviewed = $false }
    if ($null -ne $size -and $size -gt 0) {
        $density = [math]::Round($items.Count / $size, 3)
        $level = Get-Judgement 'Findings' $unit $process $density
        $judgement = @{ '少ない' = '少ない（レビュー不足の疑い）'; '多い' = '多い（品質に問題の疑い）'; '基準内' = '基準内'; '' = '' }[$level]
        if ($null -ne $d.Effort) {
            $effortDensity = [math]::Round($d.Effort / $size, 3)
            $effortJudgement = Get-Judgement 'Effort' $unit $process $effortDensity
            if ($level -and $effortJudgement) { $zone = $zones[('{0}|{1}' -f $effortJudgement, $level)] }
        }
    }
    elseif ($sizes.Count -gt 0) { $judgement = '規模なし' }
    if ($wasReviewed -eq $false) { $judgement = '未レビュー'; $zone = $null }
    $effort = $null
    if ($null -ne $d.Effort) { $effort = [math]::Round($d.Effort, 2) }
    $r = [pscustomobject]@{
        Document = $d.Name; Findings = $items.Count; Open = @($items | Where-Object { -not $_.Closed }).Count; Overdue = @($items | Where-Object Overdue).Count
        Size = $size; Unit = $unit; Density = $density; Judgement = $judgement; Process = $process
        Effort = $effort; Reviews = $d.Reviews; EffortDensity = $effortDensity; EffortJudgement = $effortJudgement
        Zone = $(if ($zone) { $zone.Zone } else { '' }); ZoneView = $(if ($zone) { $zone.View } else { '' }); ZoneAction = $(if ($zone) { $zone.Action } else { '' })
        NeedsAction = [bool]($zone -and $zone.NeedsAction); Reviewed = $wasReviewed
        BySeverity = (Format-Breakdown $items 'Severity'); ByCategory = (Format-Breakdown $items 'Category')
    }
    $r.PSObject.TypeNames.Insert(0, 'ReviewQuality.Document')
    $results.Add($r)
}
$results = @($results | Sort-Object @{ e = { $_.Findings }; Descending = $true }, Document)
if ($ReviewLog) { Update-TypeData -TypeName 'ReviewQuality.Document' -DefaultDisplayPropertySet Document, Findings, Open, Density, Effort, EffortDensity, Zone -Force -WhatIf:$false }
$results

# 日ごとの起票・完了（収束の確認）。未完了の残りが合うよう、起票日のある指摘だけを数え、完了は完了した指摘の完了日だけ数える
# （差し戻しなどで未完了に戻った指摘に残っている完了日は数えない）
$trend = New-Object System.Collections.Generic.List[object]
$dated = @($findings | Where-Object { $_.Opened })
$dates = @($dated | ForEach-Object { $_.Opened; if ($_.Closed) { $_.ClosedOn } } | Where-Object { $_ } | Sort-Object -Unique)
$cumOpen = 0; $cumClosed = 0
foreach ($day in $dates) {
    $o = @($dated | Where-Object { $_.Opened -eq $day }).Count
    $c = @($dated | Where-Object { $_.Closed -and $_.ClosedOn -eq $day }).Count
    $cumOpen += $o; $cumClosed += $c
    $trend.Add([pscustomobject]@{ Date = $day; Opened = $o; Closed = $c; TotalOpened = $cumOpen; TotalClosed = $cumClosed; Remaining = $cumOpen - $cumClosed })
}

if ($OutXlsx) {
    $docRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $size = ''; $density = ''
        if ($null -ne $r.Size) { $size = [double]$r.Size }
        if ($null -ne $r.Density) { $density = [double]$r.Density }
        $values = @($r.Document, [double]$r.Findings, [double]$r.Open, [double]$r.Overdue, $size, $r.Unit, $density, $r.Judgement, $r.BySeverity, $r.ByCategory)
        if ($ReviewLog) {
            $effort = ''; $effortDensity = ''
            if ($null -ne $r.Effort) { $effort = [double]$r.Effort }
            if ($null -ne $r.EffortDensity) { $effortDensity = [double]$r.EffortDensity }
            $values += @($effort, $effortDensity, $r.EffortJudgement, $r.Zone, $r.ZoneView, $r.ZoneAction)
        }
        $docRows.Add(@{ Values = $values })
    }
    $docColumns = @(@{ Header = '設計書'; Width = 40 }, @{ Header = '指摘'; Width = 7 }, @{ Header = '未完了'; Width = 7 }, @{ Header = '期限切れ'; Width = 8 }, @{ Header = '規模'; Width = 8 }, @{ Header = '分母'; Width = 9 }, @{ Header = '指摘密度'; Width = 9 }, @{ Header = '判定'; Width = 26 }, @{ Header = '重要度'; Width = 20 }, @{ Header = '分類'; Width = 30 })
    if ($ReviewLog) { $docColumns += @(@{ Header = 'レビュー工数の合計（人時）'; Width = 11 }, @{ Header = '工数密度'; Width = 9 }, @{ Header = '工数の判定'; Width = 9 }, @{ Header = 'ゾーン'; Width = 18 }, @{ Header = '見方'; Width = 40 }, @{ Header = '対応の例'; Width = 40 }) }
    # ゾーン分析: 縦が工数密度、横が指摘密度の 3 × 3 のマス。マスにはゾーンの名前と設計書を書く
    $zoneRows = New-Object System.Collections.Generic.List[object]
    if ($ReviewLog) {
        foreach ($effortLevel in @('多い', '基準内', '少ない')) {
            $values = @("工数密度が$effortLevel")
            foreach ($findingLevel in @('少ない', '基準内', '多い')) {
                $z = $zones[('{0}|{1}' -f $effortLevel, $findingLevel)]
                $lines = New-Object System.Collections.Generic.List[string]
                if ($z) { $lines.Add('【' + $z.Zone + '】') }
                foreach ($r in $results) {
                    if ($r.EffortJudgement -ne $effortLevel -or -not $r.Judgement.StartsWith($findingLevel)) { continue }
                    $lines.Add(('{0}（指摘 {1} / 工数 {2}）' -f $r.Document, $r.Density, $r.EffortDensity))
                }
                $values += ($lines -join "`n")
            }
            $zoneRows.Add(@{ Values = $values })
        }
    }
    $breakRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in @(@('分類', 'Category'), @('重要度', 'Severity'), @('対応者', 'Owner'), @('工程', 'Process'), @('状態', 'Status'))) {
        foreach ($g in @($findings | Where-Object { $_.($p[1]) } | Group-Object $p[1] | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)) {
            $breakRows.Add(@{ Values = @($p[0], $g.Name, [double]$g.Count, [double]@($g.Group | Where-Object { -not $_.Closed }).Count, [double]@($g.Group | Where-Object Overdue).Count) })
        }
    }
    $openRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in @($findings | Where-Object { -not $_.Closed } | Sort-Object @{ e = { -not $_.Overdue } }, @{ e = { if ($_.Due) { $_.Due } else { [datetime]::MaxValue } } })) {
        $due = ''
        if ($f.Due) { $due = $f.Due.ToString('yyyy/MM/dd', $invariant) }
        $overdue = ''
        if ($f.Overdue) { $overdue = '期限切れ' }
        $row = @{ Values = @($f.Document, $f.Content, $f.Severity, $f.Category, $f.Owner, $due, $overdue, $f.Status, ('{0} {1}' -f $f.Place.File, $f.Place.Cell).Trim()) }
        if ($f.Place.Sheet) { $row.Link = @{ Column = 8; Path = $f.Place.Path; Location = ("'{0}'!{1}" -f $f.Place.Sheet.Replace("'", "''"), $f.Place.Cell) } }
        $openRows.Add($row)
    }
    $trendRows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $trend) { $trendRows.Add(@{ Values = @($t.Date.ToString('yyyy/MM/dd', $invariant), [double]$t.Opened, [double]$t.Closed, [double]$t.TotalOpened, [double]$t.TotalClosed, [double]$t.Remaining) }) }
    try {
        # グラフ: 指摘の起票・完了の累計と未完了（推移）、レビュー工数密度 × 指摘密度（設計書別。-ReviewLog のとき、両方ある設計書を最大 50）
        $trendCharts = @(@{ Type = 'Line'; Title = '指摘の起票・完了（累計）と未完了'; Category = 0; Series = @(@{ Column = 3 }, @{ Column = 4 }, @{ Column = 5 }); YTitle = '指摘' })
        $docCharts = @()
        if ($ReviewLog) {
            $plotted = @(for ($i = 0; $i -lt $results.Count; $i++) { if ($null -ne $results[$i].EffortDensity -and $null -ne $results[$i].Density) { $i } }) | Select-Object -First 50
            if (@($plotted).Count -gt 0) {
                $docCharts = @(@{ Type = 'Scatter'; Title = 'レビュー工数密度 × 指摘密度'; Category = 11; Labels = $true; XTitle = 'レビュー工数密度'; YTitle = '指摘密度'; Width = 10; Height = 20
                        Series = @($plotted | ForEach-Object { @{ Column = 6; Rows = @($_); Name = $results[$_].Document } }) })
            }
        }
        $sheets = @(
            @{ Name = '設計書別'; Columns = $docColumns; Rows = $docRows; Charts = $docCharts },
            @{ Name = '内訳'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '値'; Width = 20 }, @{ Header = '指摘'; Width = 7 }, @{ Header = '未完了'; Width = 7 }, @{ Header = '期限切れ'; Width = 8 }); Rows = $breakRows },
            @{ Name = '未完了'; Columns = @(@{ Header = '設計書'; Width = 30 }, @{ Header = '指摘内容'; Width = 50 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '分類'; Width = 10 }, @{ Header = '対応者'; Width = 10 }, @{ Header = '期限'; Width = 11 }, @{ Header = '期限切れ'; Width = 8 }, @{ Header = '状態'; Width = 9 }, @{ Header = '指摘管理表'; Width = 36 }); Rows = $openRows },
            @{ Name = '推移'; Columns = @(@{ Header = '日付'; Width = 11 }, @{ Header = '起票'; Width = 7 }, @{ Header = '完了'; Width = 7 }, @{ Header = '起票の累計'; Width = 10 }, @{ Header = '完了の累計'; Width = 10 }, @{ Header = '未完了'; Width = 7 }); Rows = $trendRows; Charts = $trendCharts })
        if ($ReviewLog) {
            $sheets += @{ Name = 'ゾーン分析'; Columns = @(@{ Header = '工数密度 ＼ 指摘密度'; Width = 18 }, @{ Header = '指摘密度が少ない'; Width = 40 }, @{ Header = '指摘密度が基準内'; Width = 40 }, @{ Header = '指摘密度が多い'; Width = 40 }); Rows = $zoneRows }
        }
        $sheets += New-ReadLogSheet $readLog
        Export-Workbook -FilePath $OutXlsx -Sheets $sheets
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$open = @($findings | Where-Object { -not $_.Closed }).Count
$overdue = @($findings | Where-Object Overdue).Count
Write-Host ('指摘 {0} 件（未完了 {1} 件・うち期限切れ {2} 件。取下げなどで数えなかったもの {3} 件）/ 設計書 {4} 件（{5} 時点）' -f $findings.Count, $open, $overdue, $excluded, $results.Count, $AsOf.ToString('yyyy/MM/dd', $invariant))
foreach ($p in @(@('重要度', 'Severity'), @('分類', 'Category'))) {
    $text = Format-Breakdown $findings $p[1]
    if ($text) { Write-Host ('{0}: {1}' -f $p[0], $text) }
}
$flagged = @($results | Where-Object { $_.Judgement -like '少ない*' -or $_.Judgement -like '多い*' })
if ($flagged.Count -gt 0) { Write-Host ('基準値から外れた設計書: {0} 件' -f $flagged.Count) }
if ($ReviewLog) {
    $reviewed = @($results | Where-Object { $null -ne $_.Effort })
    Write-Host ('レビュー工数 {0} 人時（レビュー {1} 回・設計書 {2} 件）' -f [math]::Round((@($efforts | ForEach-Object { $_.Hours }) | Measure-Object -Sum).Sum, 2), $efforts.Count, $reviewed.Count)
    $zoned = @($results | Where-Object { $_.Zone })
    if ($zoned.Count -gt 0) { Write-Host ('ゾーン: {0}' -f (Format-Breakdown $zoned 'Zone')) }
    foreach ($r in @($results | Where-Object NeedsAction)) { Write-Host ('  要対応: {0}（{1}）{2}' -f $r.Document, $r.Zone, $r.ZoneAction) }
    if ($sizes.Count -gt 0 -and $results.Count -gt 0) {
        # レビュー実施率は、規模（設計書の一覧）があるときだけ出す（無ければ、レビューしていない設計書は結果に出てこない）
        $unreviewed = @($results | Where-Object { $_.Reviewed -eq $false })
        Write-Host ('レビュー実施率 {0:0.0}%（設計書 {1} 件のうち {2} 件）' -f (($results.Count - $unreviewed.Count) * 100.0 / $results.Count), $results.Count, ($results.Count - $unreviewed.Count))
        if ($unreviewed.Count -gt 0) { Write-Host ('  未レビュー: {0}' -f (@($unreviewed | ForEach-Object { $_.Document }) -join '・')) }
    }
    $noEffort = @($results | Where-Object { $null -eq $_.Effort -and $null -ne $_.Density -and $_.Reviewed -ne $false })
    if ($noEffort.Count -gt 0) { Write-Host ('レビューの記録が無い設計書: {0} 件（ゾーンを決められません）' -f $noEffort.Count) }
    if (@($limitRows | Where-Object { $_.Metric -eq 'Effort' }).Count -eq 0) { Write-Warning ('基準値の CSV にレビュー工数密度の行がありません（指標の列に「レビュー工数密度」と書いた行を足してください）: {0}' -f $Limits) }
}
else {
    $zeroDocs = @($results | Where-Object { $_.Findings -eq 0 })
    if ($zeroDocs.Count -gt 0) { Write-Host ('指摘 0 件の設計書: {0} 件（-ReviewLog でレビュー記録を渡すと、未レビューと区別できます）' -f $zeroDocs.Count) }
}
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     レビューの指摘管理表（Excel・CSV）を集計し、設計書ごとの指摘件数・指摘密度・未完了・期限切れを出します（Excel 不要）。
#
# .DESCRIPTION
#     指摘管理表から、指摘内容・設計書・分類・重要度・対応状況・起票日・期限・完了日・対応者などの列を見出しの文字で探して読み、
#     次を集計します。Invoke-DesignReview.ps1 のレビュー結果（指摘一覧）もそのまま読めます。
#
#       設計書別   指摘件数・未完了・期限切れ、規模あたりの指摘件数（指摘密度）と基準値との比較、重要度・分類の内訳
#       内訳       分類・重要度・対応者・工程・状態ごとの件数
#       未完了     対応が終わっていない指摘の一覧（期限切れから順。指摘管理表の該当行へのリンク付き）
#       推移       日ごとの起票・完了の件数と、未完了の残り（収束の確認）
#
#     指摘密度の分母は、-SizeCsv（Export-ExcelToPdf.ps1 -LogPath のページ数など）か、-DesignPath（Measure-DesignSize.ps1 の規模）で渡します。
#     基準値（下限・上限）から外れた設計書は「少ない（レビュー不足の疑い）」「多い（品質に問題の疑い）」と判定します。
#
#     -ReviewLog でレビュー記録（レビューの時間・参加者、または工数）を渡すと、レビュー工数密度（規模あたりの人時）も出し、
#     指摘密度との 2 軸で設計書を 9 つのゾーンに分けます（ゾーン分析）。指摘が少ないのが「レビューが足りない」のか
#     「十分にレビューして品質が良い」のかを分けられます。ゾーンの名前・見方・対応の例は tools\review-quality\ゾーン分析.csv で変えられます。
#
# .PARAMETER Path
#     指摘管理表のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#     前に -OutXlsx で保存したこのツールの結果のブック（設計書別・未完了・読み取り結果のシートがあるもの）は読みません。
#
# .PARAMETER DesignPath
#     設計書のフォルダ・ファイル。Measure-DesignSize.ps1 で数えた規模（項目数・カラム数など）を指摘密度の分母にします。
#
# .PARAMETER SizeCsv
#     設計書ごとの規模の CSV（ファイル・ページ数などの列）。Export-ExcelToPdf.ps1 -CountOnly -LogPath の CSV をそのまま使えます。
#
# .PARAMETER ReviewLog
#     レビュー記録のファイルまたはフォルダ（Excel・CSV。複数可）。1 行 1 回のレビューの表（対象成果物・レビュー時間・参加者など）か、
#     1 シート 1 回の記録票（「対象成果物」「レビュー時間」などの見出しの右に値）を読みます。工数の列があれば工数を、
#     無ければ時間 × 人数（参加者の数）をレビュー工数（人時）にします。時間は 1.5・1.5h・90分・1時間30分・1時間半・10:00～11:30 などを読めます。
#     設計書が空の行は、同じ表の上の行の設計書とみなします。読めない工数・時間の行は数えずに知らせます。
#
# .PARAMETER Limits
#     基準値の CSV（指標・分母・工程・下限・上限）。指標は「指摘密度」か「レビュー工数密度」（指標の列が無いか空なら指摘密度。ほかの指標の行は読みません）。
#     既定は tools\review-quality\基準値_サンプル.csv（数値は例です）。
#
# .PARAMETER Headers
#     指摘管理表・レビュー記録の見出しの CSV（項目と、見出しの候補）。既定は tools\review-quality\指摘管理表の見出し.csv。
#     指定した CSV に無い項目は、既定の CSV の候補を使います。
#
# .PARAMETER AsOf
#     期限切れを判定する日。既定は今日。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（設計書別・内訳・未完了・推移。-ReviewLog を付けるとゾーン分析も）。
#
# .EXAMPLE
#     .\Measure-ReviewQuality.ps1 D:\レビュー\指摘管理表.xlsx -DesignPath D:\設計書 -OutXlsx D:\品質分析.xlsx
#
# .EXAMPLE
#     # レビュー記録の工数も読み、ゾーン分析をする
#     .\Measure-ReviewQuality.ps1 D:\レビュー\指摘管理表.xlsx -ReviewLog D:\レビュー\レビュー記録 -DesignPath D:\設計書 -OutXlsx D:\品質分析.xlsx
#
# .EXAMPLE
#     # PDF のページ数を分母にする
#     .\Export-ExcelToPdf.ps1 -Path D:\設計書_PDF -CountOnly -LogPath D:\pages.csv
#     .\Measure-ReviewQuality.ps1 D:\レビュー -SizeCsv D:\pages.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1（-DesignPath では Measure-DesignSize.ps1 も）が必要です。
#     基準値のサンプルの数値は例です。社内の実績・標準値に合わせて直してください。
