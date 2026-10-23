[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string] $Config,

    [string[]] $Wbs,

    [string[]] $Findings,

    [string[]] $BugList,

    [datetime] $AsOf = (Get-Date).Date,

    [datetime] $Since,

    [string] $Holidays,

    [string] $OutDir,

    [string] $Title,

    [string] $HistoryCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

# 設定ファイル（-Config）の値は、コマンドで指定しなかったパラメーターにだけ使う
if ($Config) {
    $settings = Read-WeeklyReportConfig $Config
    foreach ($k in $settings.Keys) { if (-not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value $settings[$k] } }
}
if (-not ($Wbs -or $Findings -or $BugList)) { throw '-Wbs・-Findings・-BugList のどれか（または -Config の設定ファイル）を指定してください' }
foreach ($name in @('Holidays', 'OutDir', 'HistoryCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if (-not $Title) { $Title = '週次報告' }

Update-TypeData -TypeName 'WeeklyReport.Item' -DefaultDisplayPropertySet Owner, State, Kind, Title, Date -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$AsOf = $AsOf.Date
if ($PSBoundParameters.ContainsKey('Since')) { $Since = $Since.Date } else { $Since = $AsOf.AddDays(-6) }
if ($Since -gt $AsOf) { throw ('-Since（{0:yyyy/MM/dd}）が -AsOf（{1:yyyy/MM/dd}）より後です' -f $Since, $AsOf) }
# 来週: 基準日の翌日から 7 日
$nextStart = $AsOf.AddDays(1)
$nextEnd = $AsOf.AddDays(7)
$maxListed = 10

$reportName = '{0}_{1}' -f $Title, $AsOf.ToString('yyyyMMdd')
$outXlsx = $null; $outText = $null
if ($OutDir) {
    $outXlsx = Join-Path $OutDir ($reportName + '.xlsx')
    $outText = Join-Path $OutDir ($reportName + '.txt')
}

function Test-InRange {
    param($Date, [datetime] $From, [datetime] $To)
    return ($null -ne $Date) -and $Date -ge $From -and $Date -le $To
}

function Format-Short {
    #     一覧の文章に載せる長さにする（改行と続く空白は 1 つの空白に）
    param([string] $Text, [int] $Max = 40)
    $t = ([string]$Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + '…' }
    return $t
}

function Format-Day {
    param($Date)
    if ($null -eq $Date) { return '' }
    return $Date.ToString('M/d')
}

# ---------------------------------------------------------------- 担当者


# 担当者の名前は、全角・半角と空白の違いをそろえて同じ人とみなす（表示は最初に出てきた書き方）
$ownerNames = @{}
$noOwner = '（担当なし）'
function Get-OwnerKey {
    param([string] $Name)
    $key = (([string]$Name).Normalize($nfkc) -replace '\s', '')
    if (-not $key) { return '' }
    if (-not $ownerNames.ContainsKey($key)) { $ownerNames[$key] = ([string]$Name).Trim() }
    return $key
}
function Get-OwnerName {
    param([string] $Key)
    if (-not $Key) { return $noOwner }
    return $ownerNames[$Key]
}

# 結果のフォルダ（前の週の報告）・実績の CSV と、ほかの表のファイルは読まない
$outFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($OutDir) { $outFiles.UnionWith((Get-FilePathSet @($OutDir))) }
if ($HistoryCsv) { [void]$outFiles.Add($HistoryCsv) }

# ---------------------------------------------------------------- 進捗（WBS）


$tasks = @()
if ($Wbs) {
    $progressArgs = @{ Path = $Wbs; AsOf = $AsOf }
    if ($Holidays) { $progressArgs['Holidays'] = $Holidays }
    if ($NoExcel) { $progressArgs['NoExcel'] = $true }
    $tasks = @(& (Join-Path $PSScriptRoot 'Measure-Progress.ps1') @progressArgs 6>$null)
}
foreach ($t in $tasks) { $t | Add-Member -NotePropertyName OwnerKey -NotePropertyValue (Get-OwnerKey $t.Owner) -Force }
$delayed = @($tasks | Where-Object { $_.Status -eq '遅延' } | Sort-Object @{ e = { $_.Delay }; Descending = $true }, PlanEnd)
$doneInPeriod = @($tasks | Where-Object { $_.Done -and (Test-InRange $_.DoneOn $Since $AsOf) } | Sort-Object DoneOn, WbsNo)
$plannedInPeriod = @($tasks | Where-Object { Test-InRange $_.PlanEnd $Since $AsOf })
$nextTasks = @($tasks | Where-Object { -not $_.Done -and ((Test-InRange $_.PlanEnd $nextStart $nextEnd) -or (Test-InRange $_.PlanStart $nextStart $nextEnd)) } |
        Sort-Object @{ e = { if (Test-InRange $_.PlanEnd $nextStart $nextEnd) { $_.PlanEnd } else { $_.PlanStart } } }, WbsNo)
$spi = $null
if ($tasks.Count -gt 0) {
    $pv = ($tasks | Measure-Object PV -Sum).Sum
    if ($pv -gt 0) { $spi = [math]::Round(($tasks | Measure-Object EV -Sum).Sum / $pv, 2) }
}

function Get-NextLabel {
    #     来週の予定のタスクの区分（終了・開始）
    param($Task)
    $ends = Test-InRange $Task.PlanEnd $nextStart $nextEnd
    $starts = Test-InRange $Task.PlanStart $nextStart $nextEnd
    if ($ends -and $starts) { return '開始・終了' }
    if ($ends) { return '終了' }
    return '開始'
}

# ---------------------------------------------------------------- 指摘・障害


$issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    param([string] $Kind, $Record, [string] $Id, [string] $Target, [string] $Content, [string] $Severity, [string] $Owner, [string] $Status,
        $Opened, $Due, $ClosedOn, [bool] $Closed)
    # 基準日より後に起票したものは数えず、基準日より後に完了したものは基準日には未完了とする（前の週の報告を後から作るとき）
    if ($null -ne $Opened -and $Opened -gt $AsOf) { return }
    if ($Closed -and $null -ne $ClosedOn -and $ClosedOn -gt $AsOf) { $Closed = $false; $ClosedOn = $null }
    $cell = ''
    if ($Record.Sheet) { $cell = '{0}{1}' -f (ConvertTo-ColumnName $Record.Column), $Record.Row } else { $cell = '{0} 行目' -f $Record.Row }
    $open = -not $Closed
    $issues.Add([pscustomobject]@{
            Kind = $Kind; Id = $Id; Target = $Target; Content = $Content; Severity = $Severity; Owner = $Owner; OwnerKey = (Get-OwnerKey $Owner); Status = $Status
            Opened = $Opened; Due = $Due; ClosedOn = $ClosedOn; Closed = $Closed
            Overdue = ($open -and $null -ne $Due -and $Due -lt $AsOf); DueSoon = ($open -and (Test-InRange $Due $AsOf $nextEnd))
            OpenedInPeriod = (Test-InRange $Opened $Since $AsOf); ClosedInPeriod = ($Closed -and (Test-InRange $ClosedOn $Since $AsOf))
            Place = [pscustomobject]@{ File = $Record.File; Path = $Record.Path; Sheet = $Record.Sheet; Cell = $cell }
        })
}

$findingFiles = $null; $bugFiles = $null
if ($Findings) { $findingFiles = Get-FilePathSet $Findings -NoRecurse:$NoRecurse }
if ($BugList) { $bugFiles = Get-FilePathSet $BugList -NoRecurse:$NoRecurse }
function Join-Skip {
    param($Sets)
    $all = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($s in $Sets) { if ($s) { $all.UnionWith($s) } }
    return , $all
}
$readArgs = @{ NoRecurse = $NoRecurse; NoExcel = $NoExcel }

$hasFindings = $false
if ($Findings) {
    # 対応が終わった・指摘として数えない状態（Measure-ReviewQuality.ps1 と同じ）
    $closedPattern = '完了|済|クローズ|close|done|対応不要|不要|対処なし'
    $excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|誤指摘'
    $fields = @('設計書', '指摘内容', '分類', '重要度', '状態', '対応要否', '起票日', '期限', '完了日', '対応者')
    $candidates = Read-HeaderCandidates (Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv') $fields
    $records = Read-HeaderRecords @readArgs -Inputs (Get-InputFiles $Findings '指摘管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
        -Need @('指摘内容', @('設計書', '状態', '重要度', '分類')) -Key @('指摘内容') -What '指摘管理表' -Skip (Join-Skip @($outFiles, $bugFiles))
    if ($records.Count -eq 0) { Write-Warning '指摘管理表が見つかりませんでした（指摘内容と、設計書・状態・重要度・分類のどれかの見出しがある表を探します）' }
    foreach ($rec in $records) {
        $v = $rec.Values
        if (-not ([string]$v['指摘内容']).Trim()) { continue }
        $status = ([string]$v['状態']).Trim()
        if ($status -match $excludedPattern) { continue }
        $hasFindings = $true
        $closedOn = ConvertTo-DateValue $v['完了日'] $AsOf
        $closed = ($status -match $closedPattern) -or (([string]$v['対応要否']).Trim() -match '^(不要|否|×|-)$') -or ((-not $status) -and $closedOn)
        Add-Issue '指摘' $rec '' ([string]$v['設計書']).Trim() ([string]$v['指摘内容']).Trim() ([string]$v['重要度']).Trim() ([string]$v['対応者']).Trim() $status `
            (ConvertTo-DateValue $v['起票日'] $AsOf) (ConvertTo-DateValue $v['期限'] $AsOf) $closedOn ([bool]$closed)
    }
}

$hasBugs = $false
if ($BugList) {
    # 対応が終わった・障害として数えない状態（Measure-TestQuality.ps1 と同じ）
    $closedPattern = '完了|済|クローズ|close|done|解決|対応不要'
    $excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|仕様どおり|仕様通り|障害ではない|非障害|バグではない'
    $fields = @('障害ID', '機能', '障害内容', '起票日', '完了日', '状態', '重要度', '担当', '期限')
    $candidates = Read-HeaderCandidates (Join-Path $PSScriptRoot 'test-quality/見出し.csv') $fields
    $records = Read-HeaderRecords @readArgs -Inputs (Get-InputFiles $BugList '障害管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
        -Need @('障害内容', @('起票日', '状態', '重要度', '障害ID')) -Key @('障害内容') -What '障害管理表' -Skip (Join-Skip @($outFiles, $findingFiles))
    if ($records.Count -eq 0) { Write-Warning '障害管理表が見つかりませんでした（障害内容と、起票日・状態・重要度・障害番号のどれかの見出しがある表を探します）' }
    foreach ($rec in $records) {
        $v = $rec.Values
        $status = ([string]$v['状態']).Trim()
        if ($status -match $excludedPattern) { continue }
        $hasBugs = $true
        $closedOn = ConvertTo-DateValue $v['完了日'] $AsOf
        $closed = ($status -match $closedPattern) -or ((-not $status) -and $closedOn)
        Add-Issue '障害' $rec (@(ConvertTo-BugIds $v['障害ID'])[0]) ([string]$v['機能']).Trim() ([string]$v['障害内容']).Trim() ([string]$v['重要度']).Trim() ([string]$v['担当']).Trim() $status `
            (ConvertTo-DateValue $v['起票日'] $AsOf) (ConvertTo-DateValue $v['期限'] $AsOf) $closedOn ([bool]$closed)
    }
}

function Get-SeverityRank {
    #     重要度の並び（高いものから。知らない書き方は最後）
    param([string] $Severity)
    $t = ([string]$Severity).Normalize($nfkc).Trim().ToUpperInvariant()
    if ($t -match '致命|緊急|最重要|最高|^S$') { return 0 }
    if ($t -match '高|重大|大|^A$') { return 1 }
    if ($t -match '中|^B$') { return 2 }
    if ($t -match '低|軽微|小|^C$') { return 3 }
    return 4
}

function Get-IssueTitle {
    param($Issue, [int] $Max = 0)
    $content = $Issue.Content
    if ($Max -gt 0) { $content = Format-Short $content $Max }
    $head = $Issue.Id
    if (-not $head) { $head = $Issue.Target }
    if ($head) { return '{0}: {1}' -f $head, $content }
    return $content
}

# ---------------------------------------------------------------- 担当者ごとのやること


# 状態の順: 遅延・期限切れ → 来週に終わる・期限が来週まで → 来週に始まる → そのほかの未完了
$items = New-Object System.Collections.Generic.List[object]
function Add-Item {
    param([string] $OwnerKey, [string] $State, [int] $Rank, [string] $Kind, [string] $ItemTitle, [string] $DateLabel, $Date, [string] $Note, $Place, $Source)
    $item = [pscustomobject]@{
        Owner = (Get-OwnerName $OwnerKey); OwnerKey = $OwnerKey; State = $State; Rank = $Rank; Kind = $Kind; Title = $ItemTitle
        DateLabel = $DateLabel; Date = $Date; Note = $Note; Place = $Place; Source = $Source
    }
    $item.PSObject.TypeNames.Insert(0, 'WeeklyReport.Item')
    $items.Add($item)
}
foreach ($t in $tasks) {
    if ($t.Done) { continue }
    $name = ('{0} {1}' -f $t.WbsNo, $t.Task).Trim()
    $progress = '進捗 {0}%' -f [math]::Round($t.Progress * 100)
    if ($t.Status -eq '遅延') {
        $note = '{0} 営業日の遅れ・{1}' -f $t.Delay, $progress
        if ($t.ExpectedEnd) { $note += '・見込 ' + (Format-Day $t.ExpectedEnd) }
        Add-Item $t.OwnerKey '遅延' 0 'タスク' $name '予定終了' $t.PlanEnd $note $t.Place $t
    }
    elseif (Test-InRange $t.PlanEnd $nextStart $nextEnd) { Add-Item $t.OwnerKey '来週終了' 1 'タスク' $name '予定終了' $t.PlanEnd $progress $t.Place $t }
    elseif (Test-InRange $t.PlanStart $nextStart $nextEnd) { Add-Item $t.OwnerKey '来週開始' 2 'タスク' $name '予定開始' $t.PlanStart '' $t.Place $t }
}
foreach ($i in $issues) {
    if ($i.Closed) { continue }
    $note = $i.Status
    if ($i.Severity) { $note = (@($i.Severity, $i.Status) | Where-Object { $_ }) -join '・' }
    if ($i.Overdue) { Add-Item $i.OwnerKey '期限切れ' 0 $i.Kind (Get-IssueTitle $i) '期限' $i.Due $note $i.Place $i }
    elseif ($i.DueSoon) { Add-Item $i.OwnerKey '期限間近' 1 $i.Kind (Get-IssueTitle $i) '期限' $i.Due $note $i.Place $i }
    else { Add-Item $i.OwnerKey '未完了' 3 $i.Kind (Get-IssueTitle $i) '期限' $i.Due $note $i.Place $i }
}
$sortedItems = @($items | Sort-Object @{ e = { if ($_.OwnerKey) { 0 } else { 1 } } }, Owner, Rank, @{ e = { if ($_.Date) { $_.Date } else { [datetime]::MaxValue } } }, Kind)
# 担当者の順は $sortedItems の順（Group-Object は PowerShell の版によって並べ替えるため使わない）
$groups = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[object]]'
$ownerKeys = New-Object System.Collections.Generic.List[string]
foreach ($x in $sortedItems) {
    if (-not $groups.ContainsKey($x.OwnerKey)) { $groups[$x.OwnerKey] = New-Object System.Collections.Generic.List[object]; $ownerKeys.Add($x.OwnerKey) }
    $groups[$x.OwnerKey].Add($x)
}
$owners = New-Object System.Collections.Generic.List[object]
foreach ($key in $ownerKeys) {
    $list = @($groups[$key])
    $ownerIssues = @($issues | Where-Object { $_.OwnerKey -eq $key -and -not $_.Closed })
    $owners.Add([pscustomobject]@{
            Owner = $list[0].Owner; Items = $list
            Delayed = @($list | Where-Object { $_.State -eq '遅延' }).Count; EndingNext = @($list | Where-Object { $_.State -eq '来週終了' }).Count
            OpenFindings = @($ownerIssues | Where-Object { $_.Kind -eq '指摘' }).Count; OverdueFindings = @($ownerIssues | Where-Object { $_.Kind -eq '指摘' -and $_.Overdue }).Count
            OpenBugs = @($ownerIssues | Where-Object { $_.Kind -eq '障害' }).Count; OverdueBugs = @($ownerIssues | Where-Object { $_.Kind -eq '障害' -and $_.Overdue }).Count
        })
}

# ---------------------------------------------------------------- 数字（今回と前回）


$metrics = New-Object System.Collections.Generic.List[object]
function Add-Metric {
    param([string] $Area, [string] $Name, $Value)
    $metrics.Add([pscustomobject]@{ Area = $Area; Name = $Name; Value = $Value; Previous = $null })
}
if ($Wbs) {
    Add-Metric '進捗' 'タスク' ([double]$tasks.Count)
    foreach ($s in @('完了', '遅延', '予定どおり', '未着手')) { Add-Metric '進捗' $s ([double]@($tasks | Where-Object { $_.Status -eq $s }).Count) }
    Add-Metric '進捗' 'SPI' $(if ($null -ne $spi) { [double]$spi } else { $null })
    Add-Metric '進捗' '今週の完了' ([double]$doneInPeriod.Count)
    Add-Metric '進捗' '今週に終わる予定' ([double]$plannedInPeriod.Count)
    Add-Metric '進捗' '来週の予定' ([double]$nextTasks.Count)
}
foreach ($kind in @(@('指摘', [bool]$Findings), @('障害', [bool]$BugList))) {
    if (-not $kind[1]) { continue }
    $list = @($issues | Where-Object { $_.Kind -eq $kind[0] })
    Add-Metric $kind[0] '今週の起票' ([double]@($list | Where-Object OpenedInPeriod).Count)
    Add-Metric $kind[0] '今週の完了' ([double]@($list | Where-Object ClosedInPeriod).Count)
    Add-Metric $kind[0] '未完了' ([double]@($list | Where-Object { -not $_.Closed }).Count)
    Add-Metric $kind[0] '期限切れ' ([double]@($list | Where-Object Overdue).Count)
}

$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$dateText = $AsOf.ToString('yyyy/MM/dd')
$history = New-Object System.Collections.Generic.List[object]
$previousDate = $null
$historyDates = @()
if ($HistoryCsv -and (Test-Path -LiteralPath $HistoryCsv)) {
    foreach ($h in (Read-CsvFile $HistoryCsv @('時点', '区分', '項目', '値') '実績の CSV')) { $history.Add($h) }
    # 時点は yyyy/MM/dd なので、文字の順で日付の順になる
    $historyDates = @($history | ForEach-Object { [string]$_.'時点' } | Where-Object { $_ -lt $dateText } | Sort-Object -Unique)
    if ($historyDates.Count -gt 0) {
        $previousDate = $historyDates[-1]
        $before = @{}
        foreach ($h in @($history | Where-Object { ([string]$_.'時点') -eq $previousDate })) { $before[('{0}|{1}' -f $h.'区分', $h.'項目')] = [string]$h.'値' }
        foreach ($m in $metrics) {
            $n = 0.0
            if ([double]::TryParse([string]$before[('{0}|{1}' -f $m.Area, $m.Name)], [System.Globalization.NumberStyles]::Float, $invariant, [ref]$n)) { $m.Previous = $n }
        }
    }
}
function Get-Metric {
    param([string] $Area, [string] $Name)
    return @($metrics | Where-Object { $_.Area -eq $Area -and $_.Name -eq $Name })[0]
}
function Format-Count {
    #     件数（前回があれば「（前回 3）」を付ける）
    param([string] $Area, [string] $Name, [string] $Unit = ' 件')
    $m = Get-Metric $Area $Name
    $text = '{0}{1}' -f $m.Value, $Unit
    if ($null -ne $m.Previous) { $text += '（前回 {0}）' -f $m.Previous }
    return $text
}

# ---------------------------------------------------------------- 報告の文章


$lines = New-Object System.Collections.Generic.List[string]
$lines.Add(('{0}（{1}～{2}）' -f $Title, $Since.ToString('yyyy/MM/dd'), $AsOf.ToString('yyyy/MM/dd')))
function Add-List {
    #     一覧を最大 $maxListed 件まで書き、残りは「ほか N 件」にする
    param($List, [scriptblock] $Format, [string] $Indent = '  ')
    $i = 0
    foreach ($x in $List) {
        if ($i -ge $maxListed) { $lines.Add(('{0}- ほか {1} 件' -f $Indent, (@($List).Count - $maxListed))); break }
        $lines.Add($Indent + '- ' + (& $Format $x))
        $i++
    }
}
function Format-Owner {
    param([string] $Owner)
    if ($Owner) { return '（' + $Owner + '）' }
    return ''
}

if ($Wbs) {
    $lines.Add('')
    $lines.Add('■ 進捗（WBS）')
    if ($tasks.Count -eq 0) { $lines.Add('・WBS のタスクが読めませんでした') }
    else {
        $line = '・タスク {0} 件: 完了 {1}・遅延 {2}・予定どおり {3}・未着手 {4}' -f $tasks.Count, (Get-Metric '進捗' '完了').Value, $delayed.Count, (Get-Metric '進捗' '予定どおり').Value, (Get-Metric '進捗' '未着手').Value
        $spiMetric = Get-Metric '進捗' 'SPI'
        if ($null -ne $spi) {
            $line += ' / SPI {0:0.00}' -f $spi
            if ($null -ne $spiMetric.Previous) { $line += '（前回 {0:0.00}）' -f $spiMetric.Previous }
        }
        $lines.Add($line)
        $notDone = @($plannedInPeriod | Where-Object { -not $_.Done }).Count
        $lines.Add(('・今週の完了: {0} 件 / 今週に終わる予定: {1} 件（うち未完了 {2} 件）' -f $doneInPeriod.Count, $plannedInPeriod.Count, $notDone))
        Add-List $doneInPeriod { param($t) '{0} {1}{2}' -f $t.WbsNo, $t.Task, (Format-Owner $t.Owner) }
        $lines.Add(('・遅延: {0}' -f (Format-Count '進捗' '遅延')))
        Add-List $delayed {
            param($t)
            $text = '{0} {1}{2}: {3} 営業日の遅れ・進捗 {4}%・予定終了 {5}' -f $t.WbsNo, $t.Task, (Format-Owner $t.Owner), $t.Delay, [math]::Round($t.Progress * 100), (Format-Day $t.PlanEnd)
            if ($t.ExpectedEnd) { $text += '・見込 ' + (Format-Day $t.ExpectedEnd) }
            $text
        }
        $lines.Add(('・来週（{0}～{1}）の予定: {2} 件' -f (Format-Day $nextStart), (Format-Day $nextEnd), $nextTasks.Count))
        Add-List $nextTasks {
            param($t)
            $label = Get-NextLabel $t
            $day = $t.PlanStart
            if ($label -ne '開始') { $day = $t.PlanEnd }
            '{0} {1} {2} {3}{4}' -f $label, (Format-Day $day), $t.WbsNo, $t.Task, (Format-Owner $t.Owner)
        }
    }
}
foreach ($kind in @(@('指摘', 'レビューの指摘', $Findings, $hasFindings), @('障害', '障害', $BugList, $hasBugs))) {
    if (-not $kind[2]) { continue }
    $lines.Add('')
    $lines.Add(('■ {0}' -f $kind[1]))
    if (-not $kind[3]) { $lines.Add(('・{0}が読めませんでした' -f $(if ($kind[0] -eq '指摘') { '指摘管理表' } else { '障害管理表' }))); continue }
    $lines.Add(('・今週: 起票 {0} 件・完了 {1} 件' -f (Get-Metric $kind[0] '今週の起票').Value, (Get-Metric $kind[0] '今週の完了').Value))
    $open = @($issues | Where-Object { $_.Kind -eq $kind[0] -and -not $_.Closed })
    $line = '・未完了: {0}・うち期限切れ {1}' -f (Format-Count $kind[0] '未完了'), (Format-Count $kind[0] '期限切れ')
    $severities = @($open | Where-Object { $_.Severity } | Group-Object Severity | Sort-Object @{ e = { Get-SeverityRank $_.Name } }, Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
    if ($severities.Count -gt 0) { $line += ' / 重要度 {0}' -f ($severities -join '・') }
    $lines.Add($line)
    $overdueList = @($open | Where-Object Overdue | Sort-Object Due)
    Add-List $overdueList { param($i) '期限 {0} {1}{2}' -f (Format-Day $i.Due), (Get-IssueTitle $i 40), (Format-Owner $i.Owner) }
}
if ($owners.Count -gt 0) {
    $lines.Add('')
    $lines.Add('■ 担当者別（遅延・期限切れと、来週までのもの）')
    foreach ($o in $owners) {
        $urgent = @($o.Items | Where-Object { $_.Rank -lt 3 })
        $rest = $o.Items.Count - $urgent.Count
        $counts = New-Object System.Collections.Generic.List[string]
        foreach ($s in @('遅延', '期限切れ', '来週終了', '期限間近', '来週開始')) {
            $n = @($urgent | Where-Object { $_.State -eq $s }).Count
            if ($n -gt 0) { $counts.Add(('{0} {1}' -f $s, $n)) }
        }
        if ($rest -gt 0) { $counts.Add(('ほか未完了 {0}' -f $rest)) }
        $lines.Add(('【{0}】{1}' -f $o.Owner, ($counts -join '・')))
        Add-List $urgent {
            param($x)
            $text = '[{0}] {1}: {2}（{3} {4}' -f $x.State, $x.Kind, $(if ($x.Kind -eq 'タスク') { $x.Title } else { Format-Short $x.Title 50 }), $x.DateLabel, (Format-Day $x.Date)
            if ($x.Note) { $text += '・' + $x.Note }
            $text + '）'
        }
    }
}
$reportText = ($lines -join "`r`n") + "`r`n"

# ---------------------------------------------------------------- Excel と文章を保存する


function New-PlaceLink {
    param($Place, [int] $Column)
    if (-not $Place -or -not $Place.Sheet) { return $null }
    return @{ Column = $Column; Path = $Place.Path; Location = ("'{0}'!{1}" -f $Place.Sheet.Replace("'", "''"), $Place.Cell) }
}
function Format-PlaceText {
    param($Place)
    if (-not $Place) { return '' }
    return ('{0} {1}' -f $Place.File, $Place.Cell).Trim()
}
function Format-DateCell {
    param($Date)
    if ($null -eq $Date) { return '' }
    return $Date.ToString('yyyy/MM/dd')
}

if ($OutDir) {
    if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
    $previousHeader = '前回'
    if ($previousDate) { $previousHeader = '前回（' + $previousDate + '）' }
    $metricRows = New-Object System.Collections.Generic.List[object]
    foreach ($m in $metrics) {
        $value = $m.Value
        if ($null -eq $value) { $value = '' }
        $prev = ''; $change = ''
        if ($null -ne $m.Previous) {
            $prev = [double]$m.Previous
            if ($m.Value -is [double]) { $change = [double][math]::Round($m.Value - $m.Previous, 2) }
        }
        $metricRows.Add(@{ Values = @($m.Area, $m.Name, $value, $prev, $change) })
    }
    $ownerRows = New-Object System.Collections.Generic.List[object]
    foreach ($o in $owners) { $ownerRows.Add(@{ Values = @($o.Owner, [double]$o.Delayed, [double]$o.EndingNext, [double]$o.OpenFindings, [double]$o.OverdueFindings, [double]$o.OpenBugs, [double]$o.OverdueBugs) }) }
    $itemRows = New-Object System.Collections.Generic.List[object]
    foreach ($x in $sortedItems) {
        $row = @{ Values = @($x.Owner, $x.State, $x.Kind, $x.Title, $x.DateLabel, (Format-DateCell $x.Date), $x.Note, (Format-PlaceText $x.Place)) }
        $link = New-PlaceLink $x.Place 7
        if ($link) { $row.Link = $link }
        $itemRows.Add($row)
    }
    $taskRows = New-Object System.Collections.Generic.List[object]
    $taskGroups = @(@('遅延', $delayed), @('今週の完了', $doneInPeriod), @('来週の予定', $nextTasks))
    foreach ($g in $taskGroups) {
        foreach ($t in $g[1]) {
            $label = $g[0]
            if ($label -eq '来週の予定') { $label = '来週の' + (Get-NextLabel $t) }
            $row = @{ Values = @($label, $t.WbsNo, $t.Task, $t.Owner, (Format-DateCell $t.PlanStart), (Format-DateCell $t.PlanEnd), (Format-DateCell $t.ExpectedEnd), [double]$t.Progress, [double]$t.Delay, $t.Status, (Format-PlaceText $t.Place)) }
            $link = New-PlaceLink $t.Place 10
            if ($link) { $row.Link = $link }
            $taskRows.Add($row)
        }
    }
    $issueRows = New-Object System.Collections.Generic.List[object]
    $issueOrder = @($issues | Where-Object { -not $_.Closed -or $_.ClosedInPeriod } |
            Sort-Object Kind, @{ e = { if ($_.Overdue) { 0 } elseif ($_.DueSoon) { 1 } elseif (-not $_.Closed) { 2 } else { 3 } } }, @{ e = { if ($_.Due) { $_.Due } else { [datetime]::MaxValue } } })
    foreach ($i in $issueOrder) {
        $state = '未完了'
        if ($i.Overdue) { $state = '期限切れ' } elseif ($i.DueSoon) { $state = '期限間近' } elseif ($i.Closed) { $state = '今週の完了' }
        $row = @{ Values = @($i.Kind, $state, $i.Id, $i.Target, $i.Content, $i.Severity, (Get-OwnerName $i.OwnerKey), (Format-DateCell $i.Opened), (Format-DateCell $i.Due), (Format-DateCell $i.ClosedOn), $i.Status, (Format-PlaceText $i.Place)) }
        $link = New-PlaceLink $i.Place 11
        if ($link) { $row.Link = $link }
        $issueRows.Add($row)
    }
    $sheets = @(
        @{ Name = '概要'; Columns = @(@{ Header = '区分'; Width = 8 }, @{ Header = '項目'; Width = 18 }, @{ Header = '今回'; Width = 9 }, @{ Header = $previousHeader; Width = 16 }, @{ Header = '増減'; Width = 8 }); Rows = $metricRows },
        @{ Name = '担当者'; Columns = @(@{ Header = '担当'; Width = 14 }, @{ Header = '遅延タスク'; Width = 10 }, @{ Header = '来週終了のタスク'; Width = 15 }, @{ Header = '指摘の未完了'; Width = 12 }, @{ Header = 'うち期限切れ'; Width = 12 }, @{ Header = '障害の未完了'; Width = 12 }, @{ Header = 'うち期限切れ'; Width = 12 }); Rows = $ownerRows },
        @{ Name = '担当者別'; Columns = @(@{ Header = '担当'; Width = 14 }, @{ Header = '状態'; Width = 9 }, @{ Header = '種類'; Width = 6 }, @{ Header = '内容'; Width = 60 }, @{ Header = '日付の種類'; Width = 10 }, @{ Header = '日付'; Width = 11 }, @{ Header = '補足'; Width = 30 }, @{ Header = '場所'; Width = 30 }); Rows = $itemRows })
    if ($Wbs) { $sheets += @{ Name = 'タスク'; Columns = @(@{ Header = '区分'; Width = 12 }, @{ Header = 'WBS'; Width = 8 }, @{ Header = 'タスク'; Width = 30 }, @{ Header = '担当'; Width = 12 }, @{ Header = '予定開始'; Width = 11 }, @{ Header = '予定終了'; Width = 11 }, @{ Header = '実績／見込終了'; Width = 13 }, @{ Header = '進捗率'; Width = 7 }, @{ Header = '遅れ（営業日）'; Width = 12 }, @{ Header = '状態'; Width = 10 }, @{ Header = 'WBS の場所'; Width = 30 }); Rows = $taskRows } }
    if ($Findings -or $BugList) { $sheets += @{ Name = '指摘・障害'; Columns = @(@{ Header = '種類'; Width = 6 }, @{ Header = '区分'; Width = 10 }, @{ Header = '番号'; Width = 9 }, @{ Header = '設計書・機能'; Width = 24 }, @{ Header = '内容'; Width = 50 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '担当'; Width = 12 }, @{ Header = '起票日'; Width = 11 }, @{ Header = '期限'; Width = 11 }, @{ Header = '完了日'; Width = 11 }, @{ Header = '状態'; Width = 10 }, @{ Header = '場所'; Width = 30 }); Rows = $issueRows } }
    # 推移: 直近（最大 12 回）の数字を、時点ごとに 1 行で並べる
    if ($HistoryCsv) {
        $byDate = @{}
        foreach ($h in $history) { $byDate[('{0}|{1}|{2}' -f $h.'時点', $h.'区分', $h.'項目')] = [string]$h.'値' }
        $trendColumns = @(@{ Header = '時点'; Width = 11 })
        foreach ($m in $metrics) { $trendColumns += @{ Header = ('{0}の{1}' -f $m.Area, $m.Name); Width = 12 } }
        $trendRows = New-Object System.Collections.Generic.List[object]
        foreach ($d in @(@($historyDates | Select-Object -Last 11) + $dateText)) {
            $values = @($d)
            foreach ($m in $metrics) {
                if ($d -eq $dateText) { $values += $(if ($null -ne $m.Value) { $m.Value } else { '' }); continue }
                $n = 0.0
                if ([double]::TryParse([string]$byDate[('{0}|{1}|{2}' -f $d, $m.Area, $m.Name)], [System.Globalization.NumberStyles]::Float, $invariant, [ref]$n)) { $values += $n } else { $values += '' }
            }
            $trendRows.Add(@{ Values = $values })
        }
        # 残っているもの（遅延・未完了・期限切れ）の推移と、SPI の推移
        $remaining = @()
        for ($c = 0; $c -lt $metrics.Count; $c++) { if ($metrics[$c].Name -match '^(遅延|未完了|期限切れ)$') { $remaining += @{ Column = $c + 1 } } }
        $charts = @()
        if ($remaining.Count -gt 0) { $charts += @{ Type = 'Line'; Title = '遅延・未完了・期限切れの推移'; Category = 0; Series = $remaining; YTitle = '件' } }
        for ($c = 0; $c -lt $metrics.Count; $c++) { if ($metrics[$c].Name -eq 'SPI') { $charts += @{ Type = 'Line'; Title = 'SPI の推移'; Category = 0; Series = @(@{ Column = $c + 1 }) } } }
        $sheets += @{ Name = '推移'; Columns = $trendColumns; Rows = $trendRows; Charts = $charts }
    }
    try { Export-Workbook -FilePath $outXlsx -Sheets $sheets }
    catch {
        Write-Warning ("報告の Excel を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $outXlsx, $_.Exception.Message)
        $outXlsx = $null
    }
    [System.IO.File]::WriteAllText($outText, $reportText, (New-Object System.Text.UTF8Encoding($true)))
}

# 実績の CSV に今回の数字を足す（同じ時点の行は置き換える）
if ($HistoryCsv) {
    $kept = @($history | Where-Object { ([string]$_.'時点') -ne $dateText } | ForEach-Object { [pscustomobject][ordered]@{ '時点' = $_.'時点'; '区分' = $_.'区分'; '項目' = $_.'項目'; '値' = $_.'値' } })
    $added = @($metrics | Where-Object { $null -ne $_.Value } | ForEach-Object { [pscustomobject][ordered]@{ '時点' = $dateText; '区分' = $_.Area; '項目' = $_.Name; '値' = ([double]$_.Value).ToString($invariant) } })
    $dir = Split-Path -Path $HistoryCsv -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($HistoryCsv, (((@($kept + $added) | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

$sortedItems

Write-Host ''
foreach ($l in $lines) { Write-Host $l }
Write-Host ''
if ($outXlsx) { Write-Host "報告の Excel を保存しました: $outXlsx" }
if ($outText) { Write-Host "報告の文章を保存しました: $outText" }
if ($HistoryCsv) {
    $prevText = '前回の実績はありません'
    if ($previousDate) { $prevText = '前回 ' + $previousDate + ' と比べました' }
    Write-Host ('実績の CSV に {0} の数字を足しました（{1}）: {2}' -f $dateText, $prevText, $HistoryCsv)
}

# .SYNOPSIS
#     WBS・レビューの指摘管理表・障害管理表から、週次報告（進捗・指摘・障害の数字と、担当者ごとのやること）を作ります（Excel 不要）。
#
# .DESCRIPTION
#     毎週の進捗報告のために、表を見て数え、文章にまとめる作業を機械でします。
#
#       進捗      WBS を Measure-Progress.ps1 で集計し、タスクの状態・SPI、今週の完了、遅延（遅れの大きい順）、来週に終わる・始まるタスク
#       指摘      今週の起票・完了、未完了（重要度の内訳）・期限切れ
#       障害      今週の起票・完了、未完了（重要度の内訳）・期限切れ
#       担当者別  遅延したタスク・期限切れの指摘と障害、来週に終わるタスク・期限が来るもの、来週に始まるタスク（担当者ごと）
#
#     「今週」は -Since から -AsOf まで（既定は -AsOf までの 7 日）、「来週」は -AsOf の翌日から 7 日です。
#     -OutDir を指定すると、報告の Excel（概要・担当者・担当者別・タスク・指摘・障害。元の表へのリンク付き）と、
#     メールやチャットに貼る文章（.txt）を保存します。-HistoryCsv を指定すると毎週の数字を残し、前回と比べます。
#
# .PARAMETER Config
#     週次報告の設定ファイル（または、それを置いた案件のフォルダ）。WBS・指摘管理表・障害管理表・休日・出力先・実績を書いておきます。
#     ひな形は tools\weekly-report\週次報告の設定_ひな形.csv。コマンドで指定したパラメーターのほうを使います。
#
# .PARAMETER Wbs
#     WBS のファイルまたはフォルダ（Excel・CSV。複数可）。読み方は Measure-Progress.ps1 と同じです。
#
# .PARAMETER Findings
#     レビューの指摘管理表のファイルまたはフォルダ（Excel・CSV。複数可）。見出しは tools\review-quality\指摘管理表の見出し.csv で探します。
#
# .PARAMETER BugList
#     障害管理表のファイルまたはフォルダ（Excel・CSV。複数可）。見出しは tools\test-quality\見出し.csv で探します（担当・期限の列も読みます）。
#
# .PARAMETER AsOf
#     報告の基準日。既定は今日。
#
# .PARAMETER Since
#     「今週」の始まりの日。既定は -AsOf の 6 日前（-AsOf までの 7 日）。
#
# .PARAMETER Holidays
#     休日の CSV（1 列目が日付。内閣府の syukujitsu.csv をそのまま使えます）。遅れの営業日の計算に使います。
#
# .PARAMETER OutDir
#     報告の Excel（タイトル_yyyyMMdd.xlsx）と文章（タイトル_yyyyMMdd.txt）を保存するフォルダ。このフォルダのファイルは表として読みません。
#
# .PARAMETER Title
#     報告のタイトルとファイル名。既定は「週次報告」。
#
# .PARAMETER HistoryCsv
#     毎週の数字を残す CSV（時点・区分・項目・値）。前回の数字と比べ、報告の Excel に推移とグラフを足します。
#
# .PARAMETER NoRecurse
#     指摘管理表・障害管理表のフォルダの、サブフォルダを読みません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\New-WeeklyReport.ps1 -Wbs D:\PJ\WBS.xlsx -Findings D:\PJ\レビュー\指摘管理表.xlsx -BugList D:\PJ\テスト\障害管理表.xlsx -OutDir D:\PJ\週次報告
#
# .EXAMPLE
#     # 案件のフォルダの「週次報告の設定.csv」を読んで作る（金曜日の報告を、翌週の月曜日に作るときは -AsOf で金曜日を指定）
#     .\New-WeeklyReport.ps1 D:\PJ -AsOf 2026-10-23
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Measure-Progress.ps1 が必要です。
#     担当者の名前は、全角・半角と空白の違いをそろえて同じ人とみなします（「山田」と「山田 太郎」は別の人になります）。
#     今週の完了は、WBS の実績終了日・表の完了日で数えます（日付の無い完了は数えません）。
