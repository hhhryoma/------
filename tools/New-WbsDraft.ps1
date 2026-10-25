[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [datetime] $Start = (Get-Date).Date,

    [string[]] $Members,

    [ValidateSet('Phase', 'Function')]
    [string] $GroupBy = 'Phase',

    [string[]] $Baseline,

    [string] $Phases,

    [string] $Holidays,

    [string] $Coefficients,

    [string] $Rules,

    [string] $OutXlsx,

    [string] $OutCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Phases) { $Phases = Join-Path $PSScriptRoot 'wbs/工程の配分.csv' }
foreach ($name in @('Phases', 'Holidays', 'Coefficients', 'Rules', 'OutXlsx', 'OutCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($Holidays -and -not (Test-Path -LiteralPath $Holidays)) { throw "休日の CSV が見つかりません: $Holidays" }

# 表示では日付を yyyy/MM/dd にする（Start・End は日付のまま）
Update-TypeData -TypeName 'WbsDraft.Task' -MemberType ScriptProperty -MemberName StartDate -Value { $this.Start.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } -Force -WhatIf:$false
Update-TypeData -TypeName 'WbsDraft.Task' -MemberType ScriptProperty -MemberName EndDate -Value { $this.End.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } -Force -WhatIf:$false
Update-TypeData -TypeName 'WbsDraft.Task' -DefaultDisplayPropertySet Wbs, Task, Member, StartDate, EndDate, Days -Force -WhatIf:$false
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 担当者（カンマ区切りでも書ける。無ければ担当者の決まっていない 1 人）
$memberList = @($Members | ForEach-Object { ([string]$_) -split '[,、]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
if ($memberList.Count -eq 0) { $memberList = @('') }

# ---------------------------------------------------------------- 工程の配分


$phaseList = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Phases @('工程', '配分', '種類') '工程の配分')) {
    $name = ([string]$r.'工程').Trim()
    if (-not $name) { continue }
    $ratio = 0.0
    if (-not [double]::TryParse(([string]$r.'配分').Trim().TrimEnd('%'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ratio) -or $ratio -le 0) {
        throw ("工程の配分の「配分」は 0 より大きい数にしてください: {0}（{1}）" -f $r.'配分', $name)
    }
    $kinds = @(([string]$r.'種類') -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $phaseList.Add([pscustomobject]@{ Name = $name; Ratio = $ratio; Kinds = $kinds; Index = $phaseList.Count })
}
if ($phaseList.Count -eq 0) { throw "工程の配分に工程がありません: $Phases" }

# ---------------------------------------------------------------- 休日・営業日


$holidaySet = New-Object 'System.Collections.Generic.HashSet[datetime]'
if ($Holidays) {
    # 内閣府の「国民の祝日」の CSV（Shift_JIS）など、1 列目が日付の CSV（Measure-Progress.ps1 と同じ）
    foreach ($line in ((Read-TextFile $Holidays) -split "\r?\n")) {
        $d = ConvertTo-DateValue (($line -split ',')[0].Trim('"'))
        if ($d) { [void]$holidaySet.Add($d.Date) }
    }
}
$businessDays = New-Object System.Collections.Generic.List[datetime]
function Get-BusinessDay {
    #     開始日から数えて Index 番目（0 始まり）の営業日（土日と休日を除く）
    param([int] $Index)
    while ($businessDays.Count -le $Index) {
        if ($businessDays.Count -eq 0) { $d = $Start.Date } else { $d = $businessDays[$businessDays.Count - 1].AddDays(1) }
        while ($d.DayOfWeek -eq [DayOfWeek]::Saturday -or $d.DayOfWeek -eq [DayOfWeek]::Sunday -or $holidaySet.Contains($d)) { $d = $d.AddDays(1) }
        $businessDays.Add($d)
    }
    return $businessDays[$Index]
}

# ---------------------------------------------------------------- 規模を数える（Measure-DesignSize.ps1）


$sizeArgs = @{ Path = $Path }
foreach ($name in @('Coefficients', 'Rules')) { $v = Get-Variable -Name $name -ValueOnly; if ($v) { $sizeArgs[$name] = $v } }
if ($Baseline) { $sizeArgs['Baseline'] = $Baseline }
if ($NoRecurse) { $sizeArgs['NoRecurse'] = $true }
if ($NoExcel) { $sizeArgs['NoExcel'] = $true }
$sizes = @(& (Join-Path $PSScriptRoot 'Measure-DesignSize.ps1') @sizeArgs 6>$null)

# 機能（WBS の 1 つの機能）: 種類・名前・工数・見積もりの根拠
$functions = New-Object System.Collections.Generic.List[object]
$noCoefficient = New-Object System.Collections.Generic.List[string]
foreach ($s in $sizes) {
    if ($Baseline) {
        # 改修: 新規と変更の機能だけを、変更の工数で
        if (@('新規', '変更') -notcontains $s.Change) { continue }
        $days = $s.ChangeDays
        $basis = '{0}（{1}）' -f $s.Change, $s.Detail
    }
    else {
        $days = $s.Days
        $basis = ('{0} {1}（{2}）' -f $s.Kind, $s.Level, $s.Measures) -replace '\s+（', '（'
    }
    if ($null -eq $days) { $noCoefficient.Add($s.Kind); $days = 0 }
    # 一覧の行を数えるもの（メッセージ一覧など）の名前「ファイル名（シート名）」は、拡張子の無いファイル名にする
    $name = $s.Name
    if ($s.File -and $name -eq ('{0}（{1}）' -f $s.File, $s.Sheet)) { $name = [System.IO.Path]::GetFileNameWithoutExtension($s.File) }
    $functions.Add([pscustomobject]@{ Kind = $s.Kind; Name = $name; Days = [double]$days; Basis = $basis; Seq = $functions.Count; Member = $null; Tasks = (New-Object System.Collections.Generic.List[object]) })
}
if ($noCoefficient.Count -gt 0) {
    Write-Warning ('係数表に無い種類の工数は 0 にしました（係数表に足してください）: {0}' -f (@($noCoefficient | Group-InOrder { $_ } | ForEach-Object { '{0} {1} 件' -f $_.Name, $_.Count }) -join '・'))
}
if ($functions.Count -eq 0) { Write-Warning ('WBS にする機能がありません（設計書の規模を数えられませんでした）: {0}' -f ($Path -join ', ')) }

# ---------------------------------------------------------------- タスクを作る


# 機能ごとに、その種類に当てはまる工程へ工数を配分する（当てはまる工程の配分の合計で割り、機能の工数の合計は変えない）
$noPhaseKinds = New-Object System.Collections.Generic.List[string]
foreach ($f in $functions) {
    $applicable = @($phaseList | Where-Object { $_.Kinds.Count -eq 0 -or $_.Kinds -contains $f.Kind })
    if ($applicable.Count -eq 0) { if (-not $noPhaseKinds.Contains($f.Kind)) { $noPhaseKinds.Add($f.Kind) }; continue }
    $ratioSum = ($applicable | Measure-Object Ratio -Sum).Sum
    $rest = [math]::Round($f.Days, 2)
    for ($i = 0; $i -lt $applicable.Count; $i++) {
        $p = $applicable[$i]
        # 丸めで合計がずれないよう、最後の工程に残りを入れる
        if ($i -eq $applicable.Count - 1) { $d = [math]::Round($rest, 2) } else { $d = [math]::Round($f.Days * $p.Ratio / $ratioSum, 2) }
        $rest -= $d
        $f.Tasks.Add([pscustomobject]@{ Function = $f; Phase = $p; Name = ('{0}（{1}）' -f $f.Name, $p.Name); Days = [math]::Max(0.0, $d); Member = ''; StartAt = 0.0; EndAt = 0.0; Start = $null; End = $null })
    }
}
if ($noPhaseKinds.Count -gt 0) { Write-Warning ('工程の配分に、次の種類の工程がありません（種類の列に足してください）: {0}' -f ($noPhaseKinds -join '・')) }
$functions = @($functions | Where-Object { $_.Tasks.Count -gt 0 })

# 担当者: 工数の大きい機能から、その時点で工数の合計がいちばん少ない担当者に割り当てる（機能の工程はすべて同じ担当者）
$load = @{}
foreach ($m in $memberList) { $load[$m] = 0.0 }
$byDays = New-Object System.Collections.Generic.List[object]
foreach ($f in $functions) { $byDays.Add($f) }
$byDays.Sort([System.Comparison[object]] { param($a, $b) if ($a.Days -ne $b.Days) { return $b.Days.CompareTo($a.Days) }; return $a.Seq.CompareTo($b.Seq) })
foreach ($f in $byDays) {
    $best = $memberList[0]
    foreach ($m in $memberList) { if ($load[$m] -lt $load[$best]) { $best = $m } }
    $f.Member = $best
    $load[$best] += $f.Days
    foreach ($t in $f.Tasks) { $t.Member = $best }
}

# 担当者ごとに順に並べて日程を決める（工程ごとの WBS は工程の順、機能ごとの WBS は機能の順。1 人日 = 1 営業日）
$allTasks = New-Object System.Collections.Generic.List[object]
foreach ($f in $functions) { foreach ($t in $f.Tasks) { $allTasks.Add($t) } }
$ordered = New-Object System.Collections.Generic.List[object]
foreach ($t in $allTasks) { $ordered.Add($t) }
if ($GroupBy -eq 'Phase') { $ordered.Sort([System.Comparison[object]] { param($a, $b) if ($a.Phase.Index -ne $b.Phase.Index) { return $a.Phase.Index.CompareTo($b.Phase.Index) }; return $a.Function.Seq.CompareTo($b.Function.Seq) }) }
else { $ordered.Sort([System.Comparison[object]] { param($a, $b) if ($a.Function.Seq -ne $b.Function.Seq) { return $a.Function.Seq.CompareTo($b.Function.Seq) }; return $a.Phase.Index.CompareTo($b.Phase.Index) }) }
$cursor = @{}
foreach ($m in $memberList) { $cursor[$m] = 0.0 }
$eps = 0.000001
foreach ($t in $ordered) {
    $t.StartAt = $cursor[$t.Member]
    $t.EndAt = $t.StartAt + $t.Days
    $cursor[$t.Member] = $t.EndAt
    $first = [int][math]::Floor($t.StartAt + $eps)
    $last = $first
    if ($t.Days -gt $eps) { $last = [math]::Max($first, [int][math]::Ceiling($t.EndAt - $eps) - 1) }
    $t.Start = Get-BusinessDay $first
    $t.End = Get-BusinessDay $last
}

# ---------------------------------------------------------------- WBS の行


$rows = New-Object System.Collections.Generic.List[object]
function Add-WbsRow {
    param([string] $Wbs, [string] $Task, $Children, [string] $Basis, [bool] $IsParent)
    $members = @($Children | ForEach-Object { $_.Member } | Select-Object -Unique)
    $member = ''
    if ($members.Count -eq 1) { $member = $members[0] }
    $days = [math]::Round(($Children | Measure-Object Days -Sum).Sum, 2)
    $r = [pscustomobject]@{
        Wbs = $Wbs; Task = $Task; Member = $member
        Start = (@($Children | Sort-Object Start)[0]).Start; End = (@($Children | Sort-Object End -Descending)[0]).End
        Days = $days; Basis = $Basis; IsParent = $IsParent
    }
    $r.PSObject.TypeNames.Insert(0, 'WbsDraft.Task')
    $rows.Add($r)
}
if ($GroupBy -eq 'Phase') {
    $n = 0
    foreach ($p in $phaseList) {
        $tasks = @($ordered | Where-Object { $_.Phase -eq $p } | Sort-Object { $_.Function.Seq })
        if ($tasks.Count -eq 0) { continue }
        $n++
        Add-WbsRow ([string]$n) $p.Name $tasks '' $true
        $c = 0
        foreach ($t in $tasks) { $c++; Add-WbsRow ('{0}.{1}' -f $n, $c) $t.Name @($t) $t.Function.Basis $false }
    }
}
else {
    $n = 0
    foreach ($f in $functions) {
        $n++
        Add-WbsRow ([string]$n) $f.Name $f.Tasks.ToArray() $f.Basis $true
        $c = 0
        foreach ($t in @($f.Tasks | Sort-Object { $_.Phase.Index })) { $c++; Add-WbsRow ('{0}.{1}' -f $n, $c) $t.Name @($t) '' $false }
    }
}

# ---------------------------------------------------------------- 結果


$rows

$memberRows = New-Object System.Collections.Generic.List[object]
foreach ($m in $memberList) {
    $mine = @($ordered | Where-Object { $_.Member -eq $m })
    if ($mine.Count -eq 0) { continue }
    $from = (@($mine | Sort-Object Start)[0]).Start
    $to = (@($mine | Sort-Object End -Descending)[0]).End
    $memberRows.Add([pscustomobject]@{ Member = $m; Tasks = $mine.Count; Days = [math]::Round(($mine | Measure-Object Days -Sum).Sum, 2); Start = $from; End = $to; BusinessDays = ($businessDays.IndexOf($to) - $businessDays.IndexOf($from) + 1) })
}

$header = @('WBS番号', 'タスク名', '担当者', '予定開始', '予定終了', '予定工数', '実績開始', '実績終了', '進捗率', '実績工数', '見積もりの根拠')
if ($OutXlsx) {
    $wbsRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) { $wbsRows.Add(@{ Values = @($r.Wbs, $r.Task, $r.Member, $r.Start, $r.End, [double]$r.Days, '', '', '', '', $r.Basis) }) }
    $loadRows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $memberRows) { $loadRows.Add(@{ Values = @($m.Member, [double]$m.Tasks, [double]$m.Days, $m.Start, $m.End, [double]$m.BusinessDays) }) }
    $sizeRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $functions) { $sizeRows.Add(@{ Values = @($f.Kind, $f.Name, $f.Basis, [double]$f.Days, $f.Member) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = 'WBS'; Columns = @(@{ Header = $header[0]; Width = 8 }, @{ Header = $header[1]; Width = 44 }, @{ Header = $header[2]; Width = 10 }, @{ Header = $header[3]; Width = 11 }, @{ Header = $header[4]; Width = 11 }, @{ Header = $header[5]; Width = 9 }, @{ Header = $header[6]; Width = 11 }, @{ Header = $header[7]; Width = 11 }, @{ Header = $header[8]; Width = 8 }, @{ Header = $header[9]; Width = 9 }, @{ Header = $header[10]; Width = 50 }); Rows = $wbsRows },
            # 担当者別・規模のシートの見出しは、WBS の見出し（タスク・開始・終了など）にしない（Measure-Progress.ps1 が WBS の表として読まないように）
            @{ Name = '担当者別'; Columns = @(@{ Header = '担当者'; Width = 12 }, @{ Header = 'タスクの数'; Width = 10 }, @{ Header = '工数の合計（人日）'; Width = 12 }, @{ Header = '最初の日'; Width = 11 }, @{ Header = '最後の日'; Width = 11 }, @{ Header = '営業日'; Width = 8 }); Rows = $loadRows },
            @{ Name = '規模'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '機能'; Width = 36 }, @{ Header = '見積もりの根拠'; Width = 50 }, @{ Header = '工数（人日）'; Width = 11 }, @{ Header = '担当者'; Width = 12 }); Rows = $sizeRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}
if ($OutCsv) {
    $dir = Split-Path -Parent $OutCsv
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $csvRows = foreach ($r in $rows) {
        $o = [ordered]@{}
        $values = @($r.Wbs, $r.Task, $r.Member, $r.Start.ToString('yyyy/MM/dd', $invariant), $r.End.ToString('yyyy/MM/dd', $invariant), $r.Days.ToString($invariant), '', '', '', '', $r.Basis)
        for ($i = 0; $i -lt $header.Count; $i++) { $o[$header[$i]] = $values[$i] }
        [pscustomobject]$o
    }
    $lines = @($csvRows | ConvertTo-Csv -NoTypeInformation)
    [System.IO.File]::WriteAllText($OutCsv, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Write-Host ''
$total = [math]::Round(($ordered | Measure-Object Days -Sum).Sum, 2)
Write-Host ('機能 {0}・工程 {1}・タスク {2}。予定工数 合計 {3} 人日' -f $functions.Count, @($ordered | ForEach-Object { $_.Phase.Name } | Select-Object -Unique).Count, $ordered.Count, $total)
if ($memberRows.Count -gt 0) {
    $named = @($memberRows | Where-Object { $_.Member })
    if ($named.Count -gt 0) {
        Write-Host ('担当者 {0} 人: {1}' -f $named.Count, (@($named | ForEach-Object { '{0} {1} 人日（{2}〜{3}）' -f $_.Member, $_.Days, $_.Start.ToString('yyyy/MM/dd', $invariant), $_.End.ToString('yyyy/MM/dd', $invariant) }) -join '・'))
    }
    $end = (@($memberRows | Sort-Object End -Descending)[0]).End
    Write-Host ('完了予定: {0}（{1} から {2} 営業日）' -f $end.ToString('yyyy/MM/dd', $invariant), (Get-BusinessDay 0).ToString('yyyy/MM/dd', $invariant), ($businessDays.IndexOf($end) + 1))
}
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
if ($OutCsv) { Write-Host "CSV を保存しました: $OutCsv" }

# .SYNOPSIS
#     設計書の規模の見積もり（Measure-DesignSize.ps1）から、機能ごとのタスク・担当者・予定日・予定工数の WBS の下書きを作ります（Excel 不要）。
#
# .DESCRIPTION
#     設計書の規模を Measure-DesignSize.ps1 で数え、機能（画面・テーブル・IF など）ごとの工数を、工程の配分の CSV の工程
#     （詳細設計・レビュー・製造・単体テストなど）に配分してタスクにします。
#
#       担当者: 工数の大きい機能から、工数の合計がいちばん少ない担当者に割り当てます（機能の工程はすべて同じ担当者）
#       予定日: 担当者ごとにタスクを順に並べ、1 人日を 1 営業日として、開始日から土日と休日を除いて決めます
#       WBS 番号: 工程ごと（-GroupBy Phase。工程 → 機能）か、機能ごと（-GroupBy Function。機能 → 工程）
#
#     出力の WBS は、Measure-Progress.ps1・Measure-Effort.ps1・New-WeeklyReport.ps1 がそのまま読める見出し
#     （WBS番号・タスク名・担当者・予定開始・予定終了・予定工数・実績開始・実績終了・進捗率・実績工数）です。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（規模を数える。複数可）。
#
# .PARAMETER Start
#     開始日（既定は今日。土日・休日なら次の営業日から）。
#
# .PARAMETER Members
#     担当者の名前（複数可。カンマ区切りでも書ける）。省略すると担当者の決まっていない 1 人で日程を決めます。
#
# .PARAMETER GroupBy
#     WBS の上の階層（Phase: 工程ごと（既定）/ Function: 機能ごと）。
#
# .PARAMETER Baseline
#     前の版の設計書のフォルダ。新規と変更の機能だけを、変更の工数（Measure-DesignSize.ps1 の -Baseline）で WBS にします（改修）。
#
# .PARAMETER Phases
#     工程の配分の CSV（工程・配分・種類）。既定は tools\wbs\工程の配分.csv。
#
# .PARAMETER Holidays
#     休日の CSV（1 列目が日付）。内閣府の「国民の祝日」の CSV（syukujitsu.csv）をそのまま使えます（Measure-Progress.ps1 と同じ）。
#
# .PARAMETER Coefficients
#     係数表の CSV（Measure-DesignSize.ps1 の -Coefficients）。
#
# .PARAMETER Rules
#     規模の数え方の CSV（Measure-DesignSize.ps1 の -Rules）。
#
# .PARAMETER OutXlsx
#     WBS を xlsx に保存します（WBS・担当者別・規模のシート）。
#
# .PARAMETER OutCsv
#     WBS を CSV（UTF-8・BOM 付き）に保存します。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の設計書は読まなくなります）。
#
# .EXAMPLE
#     .\New-WbsDraft.ps1 D:\設計書 -Start 2026-11-02 -Members 山田, 佐藤 -Holidays .\syukujitsu.csv -OutXlsx D:\WBS.xlsx
#
# .EXAMPLE
#     .\New-WbsDraft.ps1 D:\設計書_v1.1 -Baseline D:\設計書_v1.0 -Start 2026-12-01 -GroupBy Function -OutXlsx D:\改修のWBS.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Measure-DesignSize.ps1 が必要です。
#     下書きです。タスクの依存関係・レビューの日程・担当者の得意な機能などは、作ったあとに直してください。
#     機能が 10 以上あると、WBS 番号は 1.10・2.10 にもなります。Excel は標準の書式のセルに入れた 1.10 を 1.1 にするので、日報の WBS 番号の列は文字列の書式にしてください。
