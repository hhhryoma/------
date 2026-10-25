[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $Wbs,

    [datetime] $AsOf = (Get-Date).Date,

    [ValidateSet('人日', '時間')]
    [string] $Unit = '人日',

    [double] $HoursPerDay = 8,

    [double] $Tolerance = 0.1,

    [string] $Holidays,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
# 日付は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'progress/工数実績の見出し.csv' }
foreach ($name in @('Headers', 'Holidays', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($HoursPerDay -le 0) { throw '-HoursPerDay は 0 より大きい数にしてください' }

Update-TypeData -TypeName 'Effort.Task' -DefaultDisplayPropertySet WbsNo, Task, Owner, Planned, Actual, Progress, Eac, Judgement -Force -WhatIf:$false

$AsOf = $AsOf.Date
# 四捨五入（[math]::Round の既定は偶数への丸めのため、Excel と同じ丸めにする）
$away = [System.MidpointRounding]::AwayFromZero
# 時間 → 出力の単位
$perUnit = 1.0
if ($Unit -eq '人日') { $perUnit = $HoursPerDay }

function Format-Number {
    param($Value, [int] $Digits = 1)
    if ($null -eq $Value) { return '' }
    return [math]::Round([double]$Value, $Digits, $away).ToString([System.Globalization.CultureInfo]::InvariantCulture)
}

# ---------------------------------------------------------------- WBS（予定工数・進捗率）


$tasks = @()
$wbsFiles = $null
if ($Wbs) {
    $progressArgs = @{ Path = $Wbs; AsOf = $AsOf }
    if ($Holidays) { $progressArgs['Holidays'] = $Holidays }
    if ($NoExcel) { $progressArgs['NoExcel'] = $true }
    $tasks = @(& (Join-Path $PSScriptRoot 'Measure-Progress.ps1') @progressArgs 6>$null)
    # Measure-Progress は WBS のフォルダのサブフォルダを読まないため、同じ範囲を工数実績として読まない
    $wbsFiles = Get-FilePathSet $Wbs -NoRecurse
}

# ---------------------------------------------------------------- 工数実績を読む


$readLog = New-Object System.Collections.Generic.List[object]
$log = Read-EffortLog -Inputs (Get-InputFiles $Path '工数実績' -NoRecurse:$NoRecurse) -Headers $Headers -Tasks $tasks -AsOf $AsOf -HoursPerDay $HoursPerDay `
    -Skip $wbsFiles -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
$entries = $log.Entries
$unreadable = $log.Unreadable
$future = $log.Future
if ($unreadable -gt 0) { Write-Warning ('工数が数として読めない行 {0} 件は数えていません' -f $unreadable) }
if ($entries.Count -eq 0) { Write-Warning '工数実績が見つかりませんでした（工数と、WBS 番号かタスクの見出しがある表を探します）'; Write-ReadLogSummary $readLog; return }

# ---------------------------------------------------------------- タスクごと


$results = New-Object System.Collections.Generic.List[object]
$hoursOf = New-Object double[] ([math]::Max(1, $tasks.Count))
foreach ($e in $entries) { if ($e.Task -ge 0) { $hoursOf[$e.Task] += $e.Hours } }
for ($i = 0; $i -lt $tasks.Count; $i++) {
    $t = $tasks[$i]
    $actual = $hoursOf[$i] / $perUnit
    $planned = $null
    if ($null -ne $t.Effort -and $t.Effort -gt 0) { $planned = [double]$t.Effort }
    $eac = $null
    if ($t.Done) { $eac = $actual } elseif ($t.Progress -gt 0) { $eac = $actual / $t.Progress } elseif ($null -ne $planned) { $eac = [math]::Max($planned, $actual) }
    $rate = $null
    if ($null -ne $planned) { $rate = $actual / $planned }
    $r = [pscustomobject]@{
        WbsNo = $t.WbsNo; Task = $t.Task; Owner = $t.Owner; Planned = $planned; Actual = [math]::Round($actual, 2, $away); Rate = $rate; Progress = $t.Progress
        Eac = $(if ($null -ne $eac) { [math]::Round($eac, 2, $away) } else { $null }); Judgement = (Get-EffortJudgement $planned $actual $t.Progress ([bool]$t.Done) $Tolerance); Status = $t.Status
        EV = $(if ($null -ne $planned) { $planned * $t.Progress } else { $null }); Place = $t.Place
    }
    $r.PSObject.TypeNames.Insert(0, 'Effort.Task')
    $results.Add($r)
}

# WBS が無い・WBS に無い作業は、工数表の WBS 番号・タスクごとにまとめる
$others = New-Object System.Collections.Generic.List[object]
foreach ($g in @($entries | Where-Object { $_.Task -lt 0 } | Group-Object { ('{0} {1}' -f $_.WbsNo, $_.TaskName).Trim() })) {
    $sum = ($g.Group | Measure-Object Hours -Sum).Sum
    $others.Add([pscustomobject]@{
            Name = $g.Name; Actual = [math]::Round($sum / $perUnit, 2, $away); Rows = $g.Count
            Owners = (@($g.Group | Where-Object { $_.Owner } | Group-InOrder { Get-OwnerKey $_.Owner } | ForEach-Object { $_.Group[0].Owner }) -join '、'); Place = $g.Group[0].Place
        })
}
$others = @($others | Sort-Object @{ e = { $_.Actual }; Descending = $true }, Name)
if (-not $Wbs) {
    foreach ($o in $others) {
        $r = [pscustomobject]@{ WbsNo = ''; Task = $o.Name; Owner = $o.Owners; Planned = $null; Actual = $o.Actual; Rate = $null; Progress = $null; Eac = $null; Judgement = ''; Status = ''; EV = $null; Place = $o.Place }
        $r.PSObject.TypeNames.Insert(0, 'Effort.Task')
        $results.Add($r)
    }
}
$results

# ---------------------------------------------------------------- 担当者 × 週


$monday = { param($d) $d.Date.AddDays(-(([int]$d.DayOfWeek + 6) % 7)) }
$dated = @($entries | Where-Object { $_.Date })
$weekStarts = @()
if ($dated.Count -gt 0) {
    $w = & $monday (@($dated | Sort-Object Date)[0].Date)
    $lastWeek = & $monday $AsOf
    while ($w -le $lastWeek) { $weekStarts += $w; $w = $w.AddDays(7) }
}
# 担当者は、全角・半角、英字の大文字・小文字と空白の違いをそろえてまとめる（表示は最初に出てきた書き方。名前の順）
$ownerGroups = @($entries | Group-InOrder { Get-OwnerKey $_.Owner } | Sort-Object @{ e = { $_.Group[0].Owner } }, Index)
$ownerRows = New-Object System.Collections.Generic.List[object]
foreach ($g in $ownerGroups) {
    $byWeek = @{}
    foreach ($e in $g.Group) {
        if (-not $e.Date) { continue }
        $k = (& $monday $e.Date).ToString('yyyy/MM/dd', $invariant)
        if (-not $byWeek.ContainsKey($k)) { $byWeek[$k] = 0.0 }
        $byWeek[$k] += $e.Hours
    }
    $values = @($g.Group[0].Owner)
    foreach ($w in $weekStarts) {
        $h = $byWeek[$w.ToString('yyyy/MM/dd', $invariant)]
        $values += $(if ($h) { [double][math]::Round($h / $perUnit, 2, $away) } else { [double]0 })
    }
    $values += [double][math]::Round((($g.Group | Measure-Object Hours -Sum).Sum) / $perUnit, 2, $away)
    $ownerRows.Add(@{ Values = $values })
}

# ---------------------------------------------------------------- 保存・表示


$totalHours = ($entries | Measure-Object Hours -Sum).Sum
$matchedTotal = (@($entries | Where-Object { $_.Task -ge 0 }) | Measure-Object Hours -Sum).Sum
$withPlan = @($results | Where-Object { $null -ne $_.Planned })
$sumPlanned = ($withPlan | Measure-Object Planned -Sum).Sum
$sumActual = ($withPlan | Measure-Object Actual -Sum).Sum
$sumEv = ($withPlan | Measure-Object EV -Sum).Sum
$cpi = $null
if ($sumActual -gt 0) { $cpi = [math]::Round($sumEv / $sumActual, 2, $away) }
$flagged = @($results | Where-Object { $_.Judgement -in @('予定超過', '超過の見込み') } | Sort-Object @{ e = { if ($_.Judgement -eq '予定超過') { 0 } else { 1 } } }, @{ e = { $_.Eac - $_.Planned }; Descending = $true })

if ($OutXlsx) {
    $taskRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $num = { param($x, $d) if ($null -eq $x) { '' } else { [double][math]::Round($x, $d, $away) } }
        $row = @{ Values = @($r.WbsNo, $r.Task, $r.Owner, (& $num $r.Planned 2), (& $num $r.Actual 2), (& $num $r.Rate 3), (& $num $r.Progress 3), (& $num $r.Eac 2), $r.Judgement, $r.Status, $(if ($r.Place) { ('{0} {1}' -f $r.Place.File, $r.Place.Cell) } else { '' })) }
        if ($r.Place -and $r.Place.Sheet) { $row.Link = @{ Column = 10; Path = $r.Place.Path; Location = ("'{0}'!{1}" -f $r.Place.Sheet.Replace("'", "''"), $r.Place.Cell) } }
        $taskRows.Add($row)
    }
    $otherRows = New-Object System.Collections.Generic.List[object]
    foreach ($o in $others) {
        $row = @{ Values = @($o.Name, [double]$o.Actual, [double]$o.Rows, $o.Owners, ('{0} {1}' -f $o.Place.File, $o.Place.Cell)) }
        if ($o.Place.Sheet) { $row.Link = @{ Column = 4; Path = $o.Place.Path; Location = ("'{0}'!{1}" -f $o.Place.Sheet.Replace("'", "''"), $o.Place.Cell) } }
        $otherRows.Add($row)
    }
    $u = '（' + $Unit + '）'
    $sheets = @(
        @{ Name = 'タスク'; Columns = @(@{ Header = 'WBS'; Width = 8 }, @{ Header = 'タスク'; Width = 30 }, @{ Header = '担当'; Width = 12 }, @{ Header = '予定' + $u; Width = 10 }, @{ Header = '実績' + $u; Width = 10 }, @{ Header = '消化率'; Width = 8 }, @{ Header = '進捗率'; Width = 8 }, @{ Header = '完了見込' + $u; Width = 13 }, @{ Header = '判定'; Width = 12 }, @{ Header = '状態'; Width = 10 }, @{ Header = '場所'; Width = 30 }); Rows = $taskRows },
        @{ Name = '担当者×週'; Columns = @(@(@{ Header = '担当'; Width = 14 }) + @($weekStarts | ForEach-Object { @{ Header = $_.ToString('yyyy/MM/dd', $invariant); Width = 11 } }) + @(@{ Header = '合計'; Width = 9 })); Rows = $ownerRows })
    if ($Wbs) { $sheets += @{ Name = 'WBS に無い作業'; Columns = @(@{ Header = 'WBS 番号・作業'; Width = 36 }, @{ Header = '実績' + $u; Width = 10 }, @{ Header = '行'; Width = 6 }, @{ Header = '担当'; Width = 30 }, @{ Header = '場所'; Width = 30 }); Rows = $otherRows } }
    $sheets += New-ReadLogSheet $readLog
    try { Export-Workbook -FilePath $OutXlsx -Sheets $sheets }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-ReadLogSummary $readLog
$range = ''
if ($dated.Count -gt 0) { $range = '{0}～{1}・' -f (@($dated | Sort-Object Date)[0].Date.ToString('yyyy/MM/dd', $invariant)), (@($dated | Sort-Object Date)[-1].Date.ToString('yyyy/MM/dd', $invariant)) }
Write-Host ('工数の実績: {0} 時間（{1} 人日。{2}{3} 人・{4} 行）' -f (Format-Number $totalHours), (Format-Number ($totalHours / $HoursPerDay)), $range, $ownerGroups.Count, $entries.Count)
if ($future -gt 0) { Write-Host ('基準日（{0}）より後の日付の行 {1} 件は数えていません' -f $AsOf.ToString('yyyy/MM/dd', $invariant), $future) }
if ($Wbs) {
    $line = 'WBS のタスクに対応した工数: {0} 時間（{1}%）' -f (Format-Number $matchedTotal), (Format-Number ($(if ($totalHours -gt 0) { $matchedTotal / $totalHours * 100 } else { 0 })) 0)
    if ($others.Count -gt 0) { $line += ' / WBS に無い作業: ' + ((@($others | Select-Object -First 5 | ForEach-Object { '{0} {1}' -f $_.Name, (Format-Number $_.Actual 2) })) -join '・') + $(if ($others.Count -gt 5) { ' ほか' } else { '' }) + '（' + $Unit + '）' }
    Write-Host $line
    if ($withPlan.Count -gt 0) {
        $line = '予定工数のあるタスク {0} 件: 予定 {1} / 実績 {2}（{3}）' -f $withPlan.Count, (Format-Number $sumPlanned), (Format-Number $sumActual), $Unit
        if ($null -ne $cpi) { $line += ' / EV {0} / CPI {1:0.00}' -f (Format-Number $sumEv), $cpi }
        Write-Host $line
    }
    elseif ($tasks.Count -gt 0) { Write-Host 'WBS に予定工数の列が無いため、予定との比較はしていません' }
    if ($flagged.Count -gt 0) {
        Write-Host ('予定超過 {0} 件・超過の見込み {1} 件' -f @($flagged | Where-Object Judgement -eq '予定超過').Count, @($flagged | Where-Object Judgement -eq '超過の見込み').Count)
        foreach ($f in @($flagged | Select-Object -First 10)) {
            Write-Host ('  - [{0}] {1} {2}（{3}）: 予定 {4} / 実績 {5} / 進捗 {6}% → 完了見込 {7}' -f $f.Judgement, $f.WbsNo, $f.Task, $f.Owner, (Format-Number $f.Planned), (Format-Number $f.Actual), [math]::Round($f.Progress * 100, 0, $away), (Format-Number $f.Eac))
        }
    }
}
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     日報・工数実績表（Excel・CSV）の工数を、タスク・担当者・週ごとに集計し、WBS の予定工数と比べます（Excel 不要）。
#
# .DESCRIPTION
#     工数実績の表（1 行が「日付・担当・WBS 番号かタスク・工数」の形）を見出しの文字で探して読み、集計します。
#     見出しの候補は tools\progress\工数実績の見出し.csv です。
#
#       タスクごと   -Wbs の WBS のタスク（WBS 番号、無ければタスク名で対応付け）ごとの実績工数と、予定工数との比較
#                    予定超過      未完了で、実績が予定工数を超えた
#                    超過の見込み  実績 ÷ 進捗率（完了見込）が、予定工数 ×（1 + -Tolerance）を超える
#                    完了・超過    完了したが、実績が予定工数 ×（1 + -Tolerance）を超えた（次の見積もりの参考）
#       担当者 × 週  担当者ごと・週（月曜始まり）ごとの工数
#       WBS に無い作業  WBS のタスクに対応しない工数（会議・問い合わせ対応など）
#
#     予定工数・進捗率は WBS を Measure-Progress.ps1 で読みます。予定工数のあるタスクの EV（予定工数 × 進捗率）と実績から CPI も出します。
#
# .PARAMETER Path
#     工数実績の表のファイルまたはフォルダ（Excel・CSV。複数可）。担当の列が無い表は、シート名（ブックに表が複数あるとき）か
#     ファイル名を担当とみなします（人ごとの日報など）。
#
# .PARAMETER Wbs
#     WBS のファイルまたはフォルダ（予定工数・進捗率。読み方は Measure-Progress.ps1 と同じ）。
#
# .PARAMETER AsOf
#     集計する日（この日より後の日付の行は数えません）。既定は今日。
#
# .PARAMETER Unit
#     結果の単位と、WBS の予定工数の単位（人日・時間）。既定は人日。
#
# .PARAMETER HoursPerDay
#     1 人日の時間。既定は 8。工数実績は時間で読みます（見出しに「人日」があれば人日で読みます）。
#
# .PARAMETER Tolerance
#     超過の見込みとみなす、予定工数を超える割合。既定は 0.1（予定の 1.1 倍を超えると知らせる）。
#
# .PARAMETER Holidays
#     休日の CSV（WBS の予定の進捗率の計算に使う。Measure-Progress.ps1 と同じ）。
#
# .PARAMETER Headers
#     工数実績の見出しの候補の CSV。既定は tools\progress\工数実績の見出し.csv。
#
# .PARAMETER NoRecurse
#     工数実績のフォルダの、サブフォルダを読みません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（タスク・担当者×週・WBS に無い作業・読み取り結果。タスクから WBS の該当行へのリンク付き）。
#
# .EXAMPLE
#     .\Measure-Effort.ps1 D:\PJ\工数実績 -Wbs D:\PJ\WBS.xlsx -OutXlsx D:\PJ\工数集計.xlsx
#
# .EXAMPLE
#     # WBS の予定工数が時間のとき
#     .\Measure-Effort.ps1 D:\PJ\日報 -Wbs D:\PJ\WBS.xlsx -Unit 時間
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Measure-Progress.ps1 が必要です。
#     日付や担当が空の行は、同じ表の上の行の日付・担当とみなします（結合したセルなど）。
#     日付を列に並べた工数表（行がタスク、列が日付のマトリクス）は読めません。
