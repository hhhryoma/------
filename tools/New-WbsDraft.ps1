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

    [ValidateRange(0, 1000)]
    [double] $BundleBelow = 1,

    [string] $CommonTasks,

    [string] $Coefficients,

    [string] $Rules,

    [string] $OutXlsx,

    [switch] $Gantt,

    [ValidateSet('Auto', 'Day', 'Week')]
    [string] $GanttUnit = 'Auto',

    [string] $OutCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Phases) { $Phases = Join-Path $PSScriptRoot 'wbs/工程の配分.csv' }
foreach ($name in @('Phases', 'Holidays', 'Coefficients', 'Rules', 'OutXlsx', 'OutCsv', 'CommonTasks')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($Holidays -and -not (Test-Path -LiteralPath $Holidays)) { throw "休日の CSV が見つかりません: $Holidays" }
if ($Gantt -and -not $OutXlsx) { throw '-Gantt は、-OutXlsx（保存先）と一緒に指定してください' }

# 表示では日付を yyyy/MM/dd にする（Start・End は日付のまま）
Update-TypeData -TypeName 'WbsDraft.Task' -MemberType ScriptProperty -MemberName StartDate -Value { $this.Start.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } -Force -WhatIf:$false
Update-TypeData -TypeName 'WbsDraft.Task' -MemberType ScriptProperty -MemberName EndDate -Value { $this.End.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } -Force -WhatIf:$false
Update-TypeData -TypeName 'WbsDraft.Task' -DefaultDisplayPropertySet Wbs, Task, Member, StartDate, EndDate, Days -Force -WhatIf:$false
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 担当者（カンマ区切りでも書ける。「佐藤:50%」で稼働率。無ければ担当者の決まっていない 1 人）
$memberList = New-Object System.Collections.Generic.List[string]
$memberRate = @{}
foreach ($entry in @($Members | ForEach-Object { ([string]$_) -split '[,、]' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
    $name = $entry
    $rate = 1.0
    $m = [regex]::Match($entry, '^(.+?)\s*[:：]\s*([0-9]+(?:\.[0-9]+)?)\s*([%％])?$')
    if ($m.Success) {
        $name = $m.Groups[1].Value.Trim()
        $rate = [double]::Parse($m.Groups[2].Value, $invariant)
        if ($m.Groups[3].Success -or $rate -gt 1) { $rate = $rate / 100 }
        if ($rate -le 0 -or $rate -gt 1) { throw ("担当者の稼働率は、0 より大きく 100% 以下にしてください: {0}" -f $entry) }
    }
    if ($memberList.Contains($name)) { continue }
    $memberList.Add($name)
    $memberRate[$name] = $rate
}
$namedMembers = $memberList.Count
if ($memberList.Count -eq 0) { $memberList.Add(''); $memberRate[''] = 1.0 }

# ---------------------------------------------------------------- 工程の配分


$phaseList = New-Object System.Collections.Generic.List[object]
$phaseRows = Read-CsvFile $Phases @('工程', '配分', '種類') '工程の配分'
# 担当の列: 「別の人」なら、その機能を作った人とは別の担当者に割り当てる（レビュー）。担当の列の無い CSV は、名前にレビューのある工程を別の人にする
$hasOwnerColumn = $phaseRows.Count -gt 0 -and $null -ne $phaseRows[0].PSObject.Properties['担当']
foreach ($r in $phaseRows) {
    $name = ([string]$r.'工程').Trim()
    if (-not $name) { continue }
    $ratio = 0.0
    if (-not [double]::TryParse(([string]$r.'配分').Trim().TrimEnd('%'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$ratio) -or $ratio -le 0) {
        throw ("工程の配分の「配分」は 0 より大きい数にしてください: {0}（{1}）" -f $r.'配分', $name)
    }
    $kinds = @(([string]$r.'種類') -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $optional = @{}
    foreach ($column in @('担当', '待ち合わせ', '先に作る種類', '成果物')) { $optional[$column] = ''; if ($r.PSObject.Properties[$column]) { $optional[$column] = ([string]$r.$column).Trim() } }
    if ($hasOwnerColumn) { $byOther = $optional['担当'] -match '別' } else { $byOther = $name -match 'レビュー' }
    $phaseList.Add([pscustomobject]@{
        Name = $name; Ratio = $ratio; Kinds = $kinds; Index = $phaseList.Count; ByOther = [bool]$byOther
        Wait = ($optional['待ち合わせ'] -match '^(する|あり|○|〇|◯|yes|true|1)$'); Deliverable = $optional['成果物']
        # 先に作る種類: この工程では、ほかの種類のタスクは、この種類のタスクがすべて終わってから始める（製造でテーブルを先に作るなど）
        FirstKinds = @($optional['先に作る種類'] -split '\|' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    })
}
if ($phaseList.Count -eq 0) { throw "工程の配分に工程がありません: $Phases" }

# 共通のタスク（-CommonTasks）: タスク名・時期（全体・最初・最後）・工数（人日か、機能の工数の合計に対する %）・担当・成果物
$commonRows = New-Object System.Collections.Generic.List[object]
if ($CommonTasks) {
    foreach ($r in (Read-CsvFile $CommonTasks @('タスク名', '時期', '工数') '共通のタスク')) {
        $name = ([string]$r.'タスク名').Trim()
        if (-not $name) { continue }
        $when = ([string]$r.'時期').Trim()
        if (@('全体', '最初', '最後') -notcontains $when) { throw ("共通のタスクの「時期」は 全体・最初・最後 のどれかにしてください: {0}（{1}）" -f $when, $name) }
        $text = ([string]$r.'工数').Normalize([System.Text.NormalizationForm]::FormKC).Trim()
        $value = 0.0
        if (-not [double]::TryParse($text.TrimEnd('%'), [System.Globalization.NumberStyles]::Float, $invariant, [ref]$value) -or $value -lt 0) { throw ("共通のタスクの「工数」は人日か % で書いてください: {0}（{1}）" -f $text, $name) }
        $owner = ''; if ($r.PSObject.Properties['担当']) { $owner = ([string]$r.'担当').Trim() }
        $deliverable = ''; if ($r.PSObject.Properties['成果物']) { $deliverable = ([string]$r.'成果物').Trim() }
        $commonRows.Add([pscustomobject]@{ Name = $name; When = $when; Value = $value; Percent = $text.EndsWith('%'); Owner = $owner; Deliverable = $deliverable })
    }
}

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
    $functions.Add([pscustomobject]@{ Kind = $s.Kind; Name = $name; Days = [double]$days; Basis = $basis; Seq = $functions.Count; Unit = $null })
}
if ($noCoefficient.Count -gt 0) {
    Write-Warning ('係数表に無い種類の工数は 0 にしました（係数表に足してください）: {0}' -f (@($noCoefficient | Group-InOrder { $_ } | ForEach-Object { '{0} {1} 件' -f $_.Name, $_.Count }) -join '・'))
}
if ($functions.Count -eq 0) { Write-Warning ('WBS にする機能がありません（設計書の規模を数えられませんでした）: {0}' -f ($Path -join ', ')) }

# ---------------------------------------------------------------- 小さい機能をまとめる


# 工数が -BundleBelow 人日に満たない機能は、同じ種類のものを順にまとめて 1 つのタスクにする（まとめた工数が -BundleBelow 以上になるまで。
# 最後に 1 つだけ残ったものは、その前のまとまりに入れる）。テーブルを 1 つずつ 0.04 人日のレビューにしないため
$eps = 0.000001
$bundleOf = @{}   # 機能の Seq → まとめた機能（2 つ以上）
$bundleNotes = New-Object System.Collections.Generic.List[string]
if ($BundleBelow -gt 0) {
    foreach ($g in @($functions | Where-Object { $_.Days -gt $eps -and $_.Days -lt $BundleBelow - $eps } | Group-InOrder Kind)) {
        $bundles = New-Object System.Collections.Generic.List[object]
        $current = New-Object System.Collections.Generic.List[object]
        $sum = 0.0
        foreach ($fn in $g.Group) {
            $current.Add($fn)
            $sum += $fn.Days
            if ($sum -ge $BundleBelow - $eps) { $bundles.Add($current); $current = New-Object System.Collections.Generic.List[object]; $sum = 0.0 }
        }
        if ($current.Count -eq 1 -and $bundles.Count -gt 0) { $bundles[$bundles.Count - 1].Add($current[0]) }
        elseif ($current.Count -gt 0) { $bundles.Add($current) }
        $merged = @($bundles | Where-Object { $_.Count -ge 2 })
        foreach ($bundle in $merged) { foreach ($fn in $bundle) { $bundleOf[$fn.Seq] = $bundle.ToArray() } }
        if ($merged.Count -gt 0) { $bundleNotes.Add(('{0} {1} 件 → {2} つ' -f $g.Name, ($merged | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum, $merged.Count)) }
    }
}

# WBS で割り当てる単位（機能、またはまとめた機能）
$units = New-Object System.Collections.Generic.List[object]
foreach ($fn in $functions) {
    if ($fn.Unit) { continue }
    $group = @($fn)
    if ($bundleOf.ContainsKey($fn.Seq)) { $group = @($bundleOf[$fn.Seq]) }
    $name = $fn.Name
    $basis = $fn.Basis
    # まとめたタスクの名前は「最初の機能 ほか N 件」（WBS の列に収まるように。まとめた機能は見積もりの根拠と規模のシートにある）
    if ($group.Count -gt 1) {
        $name = '{0} ほか {1} 件' -f $group[0].Name, ($group.Count - 1)
        $basis = @($group | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Basis }) -join ' / '
    }
    $unit = [pscustomobject]@{
        Kind = $fn.Kind; Name = $name; Days = [math]::Round((@($group | ForEach-Object { $_.Days }) | Measure-Object -Sum).Sum, 2); Basis = $basis
        Seq = $units.Count; Functions = $group; Member = $null; Tasks = (New-Object System.Collections.Generic.List[object])
    }
    foreach ($x in $group) { $x.Unit = $unit }
    $units.Add($unit)
}

# ---------------------------------------------------------------- タスクを作る


# 機能ごとに、その種類に当てはまる工程へ工数を配分する（当てはまる工程の配分の合計で割り、機能の工数の合計は変えない）
$noPhaseKinds = New-Object System.Collections.Generic.List[string]
foreach ($u in $units) {
    $applicable = @($phaseList | Where-Object { $_.Kinds.Count -eq 0 -or $_.Kinds -contains $u.Kind })
    if ($applicable.Count -eq 0) { if (-not $noPhaseKinds.Contains($u.Kind)) { $noPhaseKinds.Add($u.Kind) }; continue }
    $ratioSum = ($applicable | Measure-Object Ratio -Sum).Sum
    $rest = [math]::Round($u.Days, 2)
    $previous = $null
    for ($i = 0; $i -lt $applicable.Count; $i++) {
        $p = $applicable[$i]
        # 丸めで合計がずれないよう、最後の工程に残りを入れる
        if ($i -eq $applicable.Count - 1) { $d = [math]::Round($rest, 2) } else { $d = [math]::Round($u.Days * $p.Ratio / $ratioSum, 2) }
        $rest -= $d
        $task = [pscustomobject]@{
            Unit = $u; Phase = $p; Name = ('{0}（{1}）' -f $u.Name, $p.Name); Days = [math]::Max(0.0, $d); Member = ''
            Previous = $previous; Next = $null; Order = 0; ReadyAt = 0.0; StartAt = 0.0; EndAt = 0.0; Start = $null; End = $null; Wbs = ''
        }
        if ($previous) { $previous.Next = $task }
        $previous = $task
        $u.Tasks.Add($task)
    }
}
if ($noPhaseKinds.Count -gt 0) { Write-Warning ('工程の配分に、次の種類の工程がありません（種類の列に足してください）: {0}' -f ($noPhaseKinds -join '・')) }
$units = @($units | Where-Object { $_.Tasks.Count -gt 0 })
$functions = @($functions | Where-Object { $_.Unit.Tasks.Count -gt 0 })

# 担当者: 工数の大きい機能から、その時点でいちばん早く手の空く担当者（工数の合計 ÷ 稼働率）に割り当てる（機能の工程は同じ担当者）。
# 担当が「別の人」の工程（レビュー）は、作った人以外で、いちばん早く手の空く担当者にする
$load = @{}
foreach ($m in $memberList) { $load[$m] = 0.0 }
function Select-Member {
    param([double] $Days, [string] $Except)
    $best = $null
    $bestValue = 0.0
    foreach ($m in $memberList) {
        if ($memberList.Count -gt 1 -and $m -eq $Except) { continue }
        $value = ($load[$m] + $Days) / $memberRate[$m]
        if ($null -eq $best -or $value -lt $bestValue - $eps) { $best = $m; $bestValue = $value }
    }
    return $best
}
$byDays = New-Object System.Collections.Generic.List[object]
foreach ($u in $units) { $byDays.Add($u) }
$byDays.Sort([System.Comparison[object]] { param($a, $b) if ($a.Days -ne $b.Days) { return $b.Days.CompareTo($a.Days) }; return $a.Seq.CompareTo($b.Seq) })
foreach ($u in $byDays) {
    $ownDays = (@($u.Tasks | Where-Object { -not $_.Phase.ByOther } | ForEach-Object { $_.Days }) | Measure-Object -Sum).Sum
    $author = Select-Member ([double]$ownDays) ''
    $u.Member = $author
    $load[$author] += [double]$ownDays
    foreach ($t in $u.Tasks) {
        if ($t.Phase.ByOther) { $t.Member = Select-Member $t.Days $author; $load[$t.Member] += $t.Days }
        else { $t.Member = $author }
    }
}
if ($namedMembers -eq 1 -and @($units | ForEach-Object { $_.Tasks } | Where-Object { $_.Phase.ByOther }).Count -gt 0) {
    Write-Warning '担当が「別の人」の工程（レビューなど）がありますが、担当者が 1 人なので、作った人に割り当てました（-Members に 2 人以上を指定してください）'
}

# 先に作る種類: その工程のほかの種類のタスクは、先に作る種類のタスク（同じ工程）がすべて終わってから始める
foreach ($u in $units) { foreach ($t in $u.Tasks) { $t | Add-Member -NotePropertyName Gate -NotePropertyValue (New-Object System.Collections.Generic.List[object]); $t | Add-Member -NotePropertyName GateNext -NotePropertyValue (New-Object System.Collections.Generic.List[object]); $t | Add-Member -NotePropertyName Pending -NotePropertyValue 0 } }
foreach ($p in $phaseList) {
    if ($p.FirstKinds.Count -eq 0) { continue }
    $phaseTasks = @($units | ForEach-Object { $_.Tasks } | Where-Object { $_.Phase -eq $p })
    $firstTasks = @($phaseTasks | Where-Object { $p.FirstKinds -contains $_.Unit.Kind })
    if ($firstTasks.Count -eq 0) { continue }
    foreach ($t in @($phaseTasks | Where-Object { $p.FirstKinds -notcontains $_.Unit.Kind })) {
        foreach ($g in $firstTasks) { $t.Gate.Add($g); $g.GateNext.Add($t) }
    }
}

# 共通のタスク: 全体（期間の最初から最後まで。担当者の日程は使わない）・最初（担当者の最初のタスク）・最後（ほかのタスクがすべて終わってから）。
# 工数の % は、機能の工数の合計に対する割合。担当が空なら、最初・最後はいちばん早く手の空く担当者、全体は担当者なし
$functionDays = [double](($units | ForEach-Object { $_.Days } | Measure-Object -Sum).Sum)
$firstPhase = [pscustomobject]@{ Name = '共通'; Ratio = 0; Kinds = @(); Index = -1; ByOther = $false; Wait = $false; Deliverable = ''; FirstKinds = @() }
$lastPhase = [pscustomobject]@{ Name = '共通'; Ratio = 0; Kinds = @(); Index = $phaseList.Count; ByOther = $false; Wait = $true; Deliverable = ''; FirstKinds = @() }
$commonList = New-Object System.Collections.Generic.List[object]
foreach ($c in $commonRows) {
    $days = $c.Value
    if ($c.Percent) { $days = $functionDays * $c.Value / 100 }
    $days = [math]::Round($days, 2)
    $basis = '共通のタスク（{0}）' -f $c.When
    if ($c.Percent) { $basis = '共通のタスク（{0}。機能の工数の {1}%）' -f $c.When, $c.Value.ToString($invariant) }
    $member = $c.Owner
    if (-not $member -and $c.When -ne '全体') { $member = Select-Member $days ''; $load[$member] += $days }
    if ($member -and -not $memberRate.ContainsKey($member)) { $memberRate[$member] = 1.0; $memberList.Add($member) }
    $phase = $firstPhase; if ($c.When -eq '最後') { $phase = $lastPhase }
    $unit = [pscustomobject]@{ Kind = '共通'; Name = $c.Name; Days = $days; Basis = $basis; Seq = -1; Functions = @(); Member = $member; Tasks = (New-Object System.Collections.Generic.List[object]) }
    $task = [pscustomobject]@{
        Unit = $unit; Phase = $phase; Name = $c.Name; Days = [double]$days; Member = $member; Previous = $null; Next = $null; Order = 0; ReadyAt = 0.0; StartAt = 0.0; EndAt = 0.0
        Start = $null; End = $null; Wbs = ''; Gate = (New-Object System.Collections.Generic.List[object]); GateNext = (New-Object System.Collections.Generic.List[object]); Pending = 0
        When = $c.When; CommonDeliverable = $c.Deliverable
    }
    $unit.Tasks.Add($task)
    $commonList.Add($task)
}

# 並べる順（工程ごとの WBS は工程の順、機能ごとの WBS は機能の順）。担当者は、この順で、始められるタスクから手を付ける。
# 共通のタスクの最初のものは前に、最後のものは後ろに置く
$ordered = New-Object System.Collections.Generic.List[object]
foreach ($u in $units) { foreach ($t in $u.Tasks) { $ordered.Add($t) } }
if ($GroupBy -eq 'Phase') { $ordered.Sort([System.Comparison[object]] { param($a, $b) if ($a.Phase.Index -ne $b.Phase.Index) { return $a.Phase.Index.CompareTo($b.Phase.Index) }; return $a.Unit.Seq.CompareTo($b.Unit.Seq) }) }
else { $ordered.Sort([System.Comparison[object]] { param($a, $b) if ($a.Unit.Seq -ne $b.Unit.Seq) { return $a.Unit.Seq.CompareTo($b.Unit.Seq) }; return $a.Phase.Index.CompareTo($b.Phase.Index) }) }
foreach ($c in @($commonList | Where-Object { $_.When -eq '最初' })) { $ordered.Insert(0, $c) }
foreach ($c in @($commonList | Where-Object { $_.When -eq '最後' })) { $ordered.Add($c) }
# 最初のものは、CSV の順に前から並べる
$firstCount = @($commonList | Where-Object { $_.When -eq '最初' }).Count
if ($firstCount -gt 1) { $ordered.Reverse(0, $firstCount) }
for ($i = 0; $i -lt $ordered.Count; $i++) { $ordered[$i].Order = $i }

# 日程: 同じ機能の工程は順に（レビューが終わってから製造）、先に作る種類のタスクの後に、ほかの種類のタスク（同じ工程）を始める。
# 待ち合わせの工程は前の工程のタスクがすべて終わってから始める。
# 担当者ごとに、始められるタスクのうちいちばん早く始められるもの（同じなら並べる順の早いもの）から入れていく。
# 1 人日 = 1 営業日（稼働率 50% なら 2 営業日）。時刻は開始日からの営業日（小数）で数える
$cursor = @{}
$ready = @{}
foreach ($m in $memberList) { $cursor[$m] = 0.0; $ready[$m] = New-Object System.Collections.Generic.List[object] }
function Add-ReadyTask {
    param($Task, [double] $At)
    $Task.ReadyAt = $At
    $list = $ready[$Task.Member]
    $i = $list.Count
    while ($i -gt 0 -and $list[$i - 1].Order -gt $Task.Order) { $i-- }
    $list.Insert($i, $Task)
}
$remaining = @{}   # 工程 → まだ日程の決まっていないタスクの数
$phaseEnd = @{}    # 工程 → 決まったタスクのいちばん遅い終わり
$allPhases = $phaseList.ToArray() + @($firstPhase, $lastPhase)
foreach ($p in $allPhases) { $remaining[$p.Index] = 0; $phaseEnd[$p.Index] = 0.0 }
foreach ($t in $ordered) { $remaining[$t.Phase.Index]++ }
# 先行のまだ決まっていないタスクの数（同じ機能の前の工程と、先に作る種類のタスク。待ち合わせの工程は、前の工程の分を待ち合わせで見る）
foreach ($t in $ordered) { $t.Pending = $t.Gate.Count; if ($t.Previous -and -not $t.Phase.Wait) { $t.Pending++ } }
$waiting = New-Object System.Collections.Generic.List[object]
function Open-WaitingTasks {
    #     待ち合わせの工程のタスクのうち、前の工程のタスクがすべて決まり、先に作る種類のタスクも決まったものを始められるようにする
    for ($i = 0; $i -lt $waiting.Count; $i++) {
        $t = $waiting[$i]
        if ($t.Pending -gt 0) { continue }
        $before = @($allPhases | Where-Object { $_.Index -lt $t.Phase.Index })
        if (@($before | Where-Object { $remaining[$_.Index] -gt 0 }).Count -gt 0) { continue }
        $at = $t.ReadyAt
        foreach ($p in $before) { if ($phaseEnd[$p.Index] -gt $at) { $at = $phaseEnd[$p.Index] } }
        Add-ReadyTask $t $at
        $waiting.RemoveAt($i)
        $i--
    }
}
function Complete-Predecessor {
    #     先行のタスク（Task）の日程が決まった: 後のタスク（Next）の待ちを減らし、待ちが無くなれば始められるようにする
    param($Task, $Next)
    $Next.Pending--
    if ($Task.EndAt -gt $Next.ReadyAt) { $Next.ReadyAt = $Task.EndAt }
    if ($Next.Pending -eq 0 -and -not $Next.Phase.Wait) { Add-ReadyTask $Next $Next.ReadyAt }
}
foreach ($t in $ordered) {
    if ($t.Phase.Wait) { $waiting.Add($t) }
    elseif ($t.Pending -eq 0) { Add-ReadyTask $t 0.0 }
}
Open-WaitingTasks
for ($n = 0; $n -lt $ordered.Count; $n++) {
    $best = $null; $bestAt = 0.0; $bestIndex = 0
    foreach ($m in $memberList) {
        $list = $ready[$m]
        for ($i = 0; $i -lt $list.Count; $i++) {
            $at = [math]::Max($cursor[$m], $list[$i].ReadyAt)
            if ($null -eq $best -or $at -lt $bestAt - $eps -or ([math]::Abs($at - $bestAt) -le $eps -and $list[$i].Order -lt $best.Order)) { $best = $list[$i]; $bestAt = $at; $bestIndex = $i }
            # この担当者は、手の空く時刻より早くは始められない（並べる順で最初のものに決まる）
            if ($at -le $cursor[$m] + $eps) { break }
        }
    }
    if ($null -eq $best) { throw '日程を決められないタスクがあります（工程の配分の待ち合わせ・先に作る種類を確かめてください）' }
    $t = $best
    $ready[$t.Member].RemoveAt($bestIndex)
    $t.StartAt = $bestAt
    $t.EndAt = $bestAt + $t.Days / $memberRate[$t.Member]
    $cursor[$t.Member] = $t.EndAt
    $remaining[$t.Phase.Index]--
    if ($t.EndAt -gt $phaseEnd[$t.Phase.Index]) { $phaseEnd[$t.Phase.Index] = $t.EndAt }
    $openWaiting = $remaining[$t.Phase.Index] -eq 0
    if ($t.Next -and -not $t.Next.Phase.Wait) { Complete-Predecessor $t $t.Next }
    foreach ($g in $t.GateNext) { Complete-Predecessor $t $g; if ($g.Phase.Wait -and $g.Pending -eq 0) { $openWaiting = $true } }
    if ($openWaiting -and $waiting.Count -gt 0) { Open-WaitingTasks }
    $first = [int][math]::Floor($t.StartAt + $eps)
    $last = $first
    if ($t.EndAt - $t.StartAt -gt $eps) { $last = [math]::Max($first, [int][math]::Ceiling($t.EndAt - $eps) - 1) }
    $t.Start = Get-BusinessDay $first
    $t.End = Get-BusinessDay $last
}
# 全体の共通のタスク: 期間の最初の営業日から、いちばん遅いタスクの終わりまで
$allEnd = $null
foreach ($t in $ordered) { if ($null -eq $allEnd -or $t.End -gt $allEnd) { $allEnd = $t.End } }
foreach ($c in @($commonList | Where-Object { $_.When -eq '全体' })) {
    $c.Start = Get-BusinessDay 0
    $c.End = $allEnd
    if (-not $c.End) { $c.End = $c.Start }
}

# ---------------------------------------------------------------- WBS の行


$rows = New-Object System.Collections.Generic.List[object]
function Add-WbsRow {
    param([string] $Wbs, [string] $Task, $Children, [string] $Basis, [bool] $IsParent, [string] $Deliverable)
    $members = @($Children | ForEach-Object { $_.Member } | Select-Object -Unique)
    $member = ''
    if ($members.Count -eq 1) { $member = $members[0] }
    $days = [math]::Round(($Children | Measure-Object Days -Sum).Sum, 2)
    $r = [pscustomobject]@{
        Wbs = $Wbs; Task = $Task; Member = $member
        Start = (@($Children | Sort-Object Start)[0]).Start; End = (@($Children | Sort-Object End -Descending)[0]).End
        Days = $days; Basis = $Basis; IsParent = $IsParent; Deliverable = $Deliverable; Predecessors = ''
    }
    $r.PSObject.TypeNames.Insert(0, 'WbsDraft.Task')
    if (-not $IsParent) { $Children[0].Wbs = $Wbs; $taskRows.Add(@{ Row = $r; Task = $Children[0] }) }
    $rows.Add($r)
}
$taskRows = New-Object System.Collections.Generic.List[object]
$phaseWbs = @{}   # 工程 → 工程の行の WBS 番号（工程ごとの WBS）
if ($GroupBy -eq 'Phase') {
    $n = 0
    foreach ($p in $phaseList) {
        $tasks = @($ordered | Where-Object { $_.Phase -eq $p } | Sort-Object { $_.Unit.Seq })
        if ($tasks.Count -eq 0) { continue }
        $n++
        Add-WbsRow ([string]$n) $p.Name $tasks '' $true $p.Deliverable
        $phaseWbs[$p.Index] = [string]$n
        $c = 0
        foreach ($t in $tasks) { $c++; Add-WbsRow ('{0}.{1}' -f $n, $c) $t.Name @($t) $t.Unit.Basis $false $p.Deliverable }
    }
}
else {
    $n = 0
    foreach ($u in $units) {
        $n++
        Add-WbsRow ([string]$n) $u.Name $u.Tasks.ToArray() $u.Basis $true ''
        $c = 0
        foreach ($t in @($u.Tasks | Sort-Object { $_.Phase.Index })) { $c++; Add-WbsRow ('{0}.{1}' -f $n, $c) $t.Name @($t) '' $false $t.Phase.Deliverable }
    }
}
# 共通のタスクは、最後の「共通」のまとまりにする（CSV の順）
if ($commonList.Count -gt 0) {
    $n++
    Add-WbsRow ([string]$n) '共通' $commonList.ToArray() '' $true ''
    $c = 0
    foreach ($t in $commonList) { $c++; Add-WbsRow ('{0}.{1}' -f $n, $c) $t.Name @($t) $t.Unit.Basis $false $t.CommonDeliverable }
}

# 先行タスク（WBS 番号）: 同じ機能の前の工程と、先に作る種類のタスク（同じ工程）。待ち合わせの工程（と最後の共通のタスク）は、前の工程のすべて
# （工程ごとの WBS は工程の番号、機能ごとの WBS はタスクの番号）。Measure-Progress.ps1 は、この列から遅れを後のタスクに送って完了見込みを出す
foreach ($item in $taskRows) {
    $t = $item.Task
    $before = New-Object System.Collections.Generic.List[string]
    if ($t.Phase.Wait) {
        if ($GroupBy -eq 'Phase') { foreach ($p in $phaseList) { if ($p.Index -lt $t.Phase.Index -and $phaseWbs.ContainsKey($p.Index)) { $before.Add($phaseWbs[$p.Index]) } } }
        else { foreach ($o in $ordered) { if ($o.Phase.Index -lt $t.Phase.Index) { $before.Add($o.Wbs) } } }
    }
    elseif ($t.Previous) { $before.Add($t.Previous.Wbs) }
    foreach ($g in $t.Gate) { if (-not $before.Contains($g.Wbs)) { $before.Add($g.Wbs) } }
    $item.Row.Predecessors = $before -join ','
}

# ---------------------------------------------------------------- 結果


$rows

$memberRows = New-Object System.Collections.Generic.List[object]
foreach ($m in $memberList) {
    $mine = @($ordered | Where-Object { $_.Member -eq $m })
    if ($mine.Count -eq 0) { continue }
    $from = (@($mine | Sort-Object Start)[0]).Start
    $to = (@($mine | Sort-Object End -Descending)[0]).End
    $memberRows.Add([pscustomobject]@{
        Member = $m; Tasks = $mine.Count; Days = [math]::Round(($mine | Measure-Object Days -Sum).Sum, 2); Start = $from; End = $to; BusinessDays = ($businessDays.IndexOf($to) - $businessDays.IndexOf($from) + 1)
        OtherDays = [math]::Round((@($mine | Where-Object { $_.Phase.ByOther } | ForEach-Object { $_.Days }) | Measure-Object -Sum).Sum, 2); Rate = $memberRate[$m]
    })
}

$header = @('WBS番号', 'タスク名', '成果物', '担当者', '先行', '予定開始', '予定終了', '予定工数', '実績開始', '実績終了', '進捗率', '実績工数', '見積もりの根拠')
$widths = @(8, 44, 16, 10, 8, 11, 11, 9, 11, 11, 8, 9, 50)
if ($OutXlsx) {
    $wbsRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) { $wbsRows.Add(@{ Values = @($r.Wbs, $r.Task, $r.Deliverable, $r.Member, $r.Predecessors, $r.Start, $r.End, [double]$r.Days, '', '', '', '', $r.Basis) }) }
    $loadRows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $memberRows) { $loadRows.Add(@{ Values = @($m.Member, ('{0}%' -f [math]::Round($m.Rate * 100, 1).ToString($invariant)), [double]$m.Tasks, [double]$m.Days, [double]$m.OtherDays, $m.Start, $m.End, [double]$m.BusinessDays) }) }
    $sizeRows = New-Object System.Collections.Generic.List[object]
    foreach ($fn in $functions) {
        $bundle = ''
        if ($fn.Unit.Functions.Count -gt 1) { $bundle = $fn.Unit.Name }
        $sizeRows.Add(@{ Values = @($fn.Kind, $fn.Name, $fn.Basis, [double]$fn.Days, $fn.Unit.Member, $bundle) })
    }
    try {
        $wbsColumns = @(for ($i = 0; $i -lt $header.Count; $i++) { @{ Header = $header[$i]; Width = $widths[$i] } })
        $wbsSheet = @{ Name = 'WBS'; Columns = $wbsColumns; Rows = $wbsRows }
        $sheets = @($wbsSheet,
            # 担当者別・規模のシートの見出しは、WBS の見出し（タスク・開始・終了など）にしない（Measure-Progress.ps1 が WBS の表として読まないように）
            @{ Name = '担当者別'; Columns = @(@{ Header = '担当者'; Width = 12 }, @{ Header = '稼働率'; Width = 8 }, @{ Header = 'タスクの数'; Width = 10 }, @{ Header = '工数の合計（人日）'; Width = 12 }, @{ Header = 'うちレビューなど（人日）'; Width = 14 }, @{ Header = '最初の日'; Width = 11 }, @{ Header = '最後の日'; Width = 11 }, @{ Header = '営業日'; Width = 8 }); Rows = $loadRows },
            @{ Name = '規模'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '機能'; Width = 36 }, @{ Header = '見積もりの根拠'; Width = 50 }, @{ Header = '工数（人日）'; Width = 11 }, @{ Header = '担当者'; Width = 12 }, @{ Header = 'まとめた先'; Width = 36 }); Rows = $sizeRows })
        # 説明のシート（使い方・列の意味。-Gantt ならバーの色の凡例も）
        $guideExtra = @(
            @('使い方', '下書き', 'WBS の下書きです。機能をまたいだ順番（テーブルの後に画面など）・会議などの共通のタスク・担当者の得意な機能は、直してから使ってください', ''),
            @('シート', '担当者別', '担当者ごとの稼働率・タスクの数・工数の合計・うちレビューなどの工数・最初の日・最後の日', ''),
            @('シート', '規模', '機能ごとの種類・見積もりの根拠・工数・担当者・まとめた先（小さい機能をまとめたタスク）', ''))
        $guideSheet = New-WbsGuideSheet $header $false $guideExtra
        if ($Gantt) {
            # ガントチャート形式の進捗管理票: WBS のシートを、右に日付の列を並べたガントチャートにする（見積もりの根拠は規模のシートにある）。休日のシートも足す
            $ganttTasks = @(foreach ($r in $rows) { [pscustomobject]@{ Wbs = $r.Wbs; Task = $r.Task; Owner = $r.Member; Predecessors = $r.Predecessors; PlanStart = $r.Start; PlanEnd = $r.End; Effort = [double]$r.Days; Summary = $r.IsParent } })
            $ganttSheets = New-GanttSheets $ganttTasks (Get-Date).Date @($holidaySet) $GanttUnit 'WBS' '進捗管理票' @('ActualEffort', 'Forecast')
            $sheets = @($ganttSheets.Sheets[0]) + @($sheets | Select-Object -Skip 1) + @($ganttSheets.Sheets[1])
            $guideSheet = New-WbsGuideSheet @($ganttSheets.Sheets[0].Columns | ForEach-Object { $_.Header }) $true $guideExtra
        }
        Export-Workbook -FilePath $OutXlsx -Sheets (@($sheets) + @($guideSheet))
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
        $values = @($r.Wbs, $r.Task, $r.Deliverable, $r.Member, $r.Predecessors, $r.Start.ToString('yyyy/MM/dd', $invariant), $r.End.ToString('yyyy/MM/dd', $invariant), $r.Days.ToString($invariant), '', '', '', '', $r.Basis)
        for ($i = 0; $i -lt $header.Count; $i++) { $o[$header[$i]] = $values[$i] }
        [pscustomobject]$o
    }
    $lines = @($csvRows | ConvertTo-Csv -NoTypeInformation)
    [System.IO.File]::WriteAllText($OutCsv, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Write-Host ''
$functionTasks = @($ordered | Where-Object { $_.Unit.Kind -ne '共通' })
$total = [math]::Round(($functionTasks | Measure-Object Days -Sum).Sum, 2)
Write-Host ('機能 {0}・工程 {1}・タスク {2}。予定工数 合計 {3} 人日' -f $functions.Count, @($functionTasks | ForEach-Object { $_.Phase.Name } | Select-Object -Unique).Count, $functionTasks.Count, $total)
if ($commonList.Count -gt 0) {
    Write-Host ('共通のタスク {0} 件: {1}（全体のタスクは担当者の日程に入れていません。兼務の分は稼働率で指定してください）' -f $commonList.Count, (@($commonList | ForEach-Object { '{0} {1} 人日（{2}）' -f $_.Name, $_.Days, $_.When }) -join '・'))
}
if ($bundleNotes.Count -gt 0) { Write-Host ('{0} 人日に満たない機能は、同じ種類をまとめて 1 つのタスクにしました: {1}（-BundleBelow 0 でまとめません）' -f $BundleBelow.ToString($invariant), ($bundleNotes -join '・')) }
if ($memberRows.Count -gt 0) {
    $named = @($memberRows | Where-Object { $_.Member })
    if ($named.Count -gt 0) {
        Write-Host ('担当者 {0} 人: {1}' -f $named.Count, (@($named | ForEach-Object {
            $rateText = ''
            if ($_.Rate -lt 1) { $rateText = '・稼働率 {0}%' -f [math]::Round($_.Rate * 100, 1).ToString($invariant) }
            '{0} {1} 人日{2}（{3}〜{4}）' -f $_.Member, $_.Days, $rateText, $_.Start.ToString('yyyy/MM/dd', $invariant), $_.End.ToString('yyyy/MM/dd', $invariant)
        }) -join '・'))
    }
    $end = (@($memberRows | Sort-Object End -Descending)[0]).End
    Write-Host ('完了予定: {0}（{1} から {2} 営業日）' -f $end.ToString('yyyy/MM/dd', $invariant), (Get-BusinessDay 0).ToString('yyyy/MM/dd', $invariant), ($businessDays.IndexOf($end) + 1))
}
if (-not $Holidays) { Write-Host '休日の CSV（-Holidays）を指定していないので、土日だけを休みにしました（祝日・年末年始も営業日です）' }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
if ($OutCsv) { Write-Host "CSV を保存しました: $OutCsv" }

# .SYNOPSIS
#     設計書の規模の見積もり（Measure-DesignSize.ps1）から、機能ごとのタスク・担当者・予定日・予定工数の WBS の下書きを作ります（Excel 不要）。
#
# .DESCRIPTION
#     設計書の規模を Measure-DesignSize.ps1 で数え、機能（画面・テーブル・IF など）ごとの工数を、工程の配分の CSV の工程
#     （詳細設計・レビュー・製造・単体テストなど）に配分してタスクにします。
#
#       まとめる: 工数が -BundleBelow 人日（既定 1）に満たない機能（テーブルなど）は、同じ種類のものをまとめて 1 つのタスクにします
#       担当者: 工数の大きい機能から、いちばん早く手の空く担当者（工数の合計 ÷ 稼働率）に割り当てます（機能の工程は同じ担当者）。
#               工程の配分の担当が「別の人」の工程（レビュー）は、作った人以外の担当者に割り当てます
#       予定日: 同じ機能の工程は順に（レビューが終わってから製造）、待ち合わせの工程は前の工程がすべて終わってから始めます。
#               担当者ごとに、始められるタスクから順に、1 人日を 1 営業日（稼働率 50% なら 2 営業日）として、土日と休日を除いて決めます
#       WBS 番号: 工程ごと（-GroupBy Phase。工程 → 機能）か、機能ごと（-GroupBy Function。機能 → 工程）
#
#     出力の WBS は、Measure-Progress.ps1・Measure-Effort.ps1・New-WeeklyReport.ps1 がそのまま読める見出し
#     （WBS番号・タスク名・担当者・予定開始・予定終了・予定工数・実績開始・実績終了・進捗率・実績工数）です。工程の成果物と、
#     先行（同じ機能の前の工程の WBS 番号。Measure-Progress.ps1 が遅れを後のタスクに送って完了見込みを出すのに使う）の列も付けます。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（規模を数える。複数可）。
#
# .PARAMETER Start
#     開始日（既定は今日。土日・休日なら次の営業日から）。
#
# .PARAMETER Members
#     担当者の名前（複数可。カンマ区切りでも書ける）。「佐藤:50%」のように書くと稼働率（兼務）にし、その人のタスクの日数を延ばします。
#     省略すると担当者の決まっていない 1 人で日程を決めます。
#
# .PARAMETER GroupBy
#     WBS の上の階層（Phase: 工程ごと（既定）/ Function: 機能ごと）。
#
# .PARAMETER Baseline
#     前の版の設計書のフォルダ。新規と変更の機能だけを、変更の工数（Measure-DesignSize.ps1 の -Baseline）で WBS にします（改修）。
#
# .PARAMETER Phases
#     工程の配分の CSV（工程・配分・種類。省けるもの: 担当・待ち合わせ・先に作る種類・成果物）。既定は tools\wbs\工程の配分.csv。
#     担当が「別の人」の工程は作った人以外に割り当て（担当の列が無ければ、名前にレビューのある工程）、待ち合わせが「する」の工程は
#     前の工程のタスクがすべて終わってから始めます（結合テストなど）。先に作る種類（テーブルなど。| 区切り）を書いた工程は、ほかの種類の
#     タスクを、その種類のタスクがすべて終わってから始めます（既定は、製造でテーブルを先に作る）。成果物は WBS の成果物の列に書きます。
#
# .PARAMETER Holidays
#     休日の CSV（1 列目が日付）。内閣府の「国民の祝日」の CSV（syukujitsu.csv）をそのまま使えます（Measure-Progress.ps1 と同じ）。
#
# .PARAMETER CommonTasks
#     共通のタスク（会議・プロジェクト管理・環境の準備など）の CSV（タスク名・時期・工数・担当・成果物）。WBS の最後の「共通」のまとまりにします。
#     時期: 全体（期間の最初から最後まで。担当者の日程には入れない）・最初（担当者の最初のタスク）・最後（ほかのタスクがすべて終わってから）。
#     工数は人日か、機能の工数の合計に対する %（5%）。担当が空なら、最初・最後はいちばん早く手の空く担当者、全体は担当者なし。
#     ひな形は tools\wbs\共通のタスク_ひな形.csv。
#
# .PARAMETER BundleBelow
#     工数がこの人日に満たない機能は、同じ種類のものを順に、合わせてこの人日以上になるまでまとめて 1 つのタスクにします（既定 1。0 でまとめない）。
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
# .PARAMETER Gantt
#     -OutXlsx の WBS のシートを、ガントチャート形式の進捗管理票にします（右に日付の列を並べ、予定・実績・遅延のバー、見込終了までの
#     見込みの延び（赤）、今日の線、土日・休日の色を条件付き書式で描く。Excel で実績・見込終了を記入するとバーと状態が変わる。
#     Measure-Progress.ps1 -OutGantt と同じ形）。
#     休日のシートも足します（-Holidays の休日。足すと曜日の「祝」と状態が変わります）。
#
# .PARAMETER GanttUnit
#     -Gantt の日付の列の単位: Auto（既定。期間が 92 日までは 1 日 1 列、それより長ければ 1 週 1 列）・Day・Week。
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
#     .\New-WbsDraft.ps1 D:\設計書 -Start 2026-11-02 -Members 山田, 佐藤, 鈴木:50% -Holidays .\syukujitsu.csv -OutXlsx D:\WBS.xlsx -Gantt
#
# .EXAMPLE
#     .\New-WbsDraft.ps1 D:\設計書_v1.1 -Baseline D:\設計書_v1.0 -Start 2026-12-01 -GroupBy Function -OutXlsx D:\改修のWBS.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Measure-DesignSize.ps1 が必要です。
#     下書きです。機能をまたいだ依存関係（テーブルの後に画面など）・担当者の得意な機能・会議などの共通のタスクは、作ったあとに直してください。
#     機能が 10 以上あると、WBS 番号は 1.10・2.10 にもなります。Excel は標準の書式のセルに入れた 1.10 を 1.1 にするので、日報の WBS 番号の列は文字列の書式にしてください。
