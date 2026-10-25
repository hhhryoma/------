[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [datetime] $AsOf = (Get-Date).Date,

    [string] $Holidays,

    [string] $SheetNamePattern,

    [string] $Headers,

    [string] $FillTemplate,

    [string] $OutFile,

    [string] $OutXlsx,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'progress/WBSの見出し.csv' }
foreach ($name in @('Headers', 'Holidays', 'FillTemplate', 'OutFile', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($FillTemplate) {
    if (-not (Test-Path -LiteralPath $FillTemplate)) { throw "テンプレートが見つかりません: $FillTemplate" }
    if (-not $OutFile) { throw '-FillTemplate には、書き込んだテンプレートの保存先（-OutFile）も指定してください' }
    if ([System.IO.Path]::GetExtension($OutFile) -ne '.xlsx') { throw "-OutFile の拡張子は .xlsx にしてください: $OutFile" }
    if (Test-Path -LiteralPath $OutFile) { throw "-OutFile のファイルがすでにあります: $OutFile" }
}

Update-TypeData -TypeName 'Progress.Task' -DefaultDisplayPropertySet WbsNo, Task, Owner, Progress, PlannedProgress, Delay, Status -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 日付は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$AsOf = $AsOf.Date

# ---------------------------------------------------------------- 設定・休日・営業日


$candidates = @{}
foreach ($r in (Read-CsvFile $Headers @('項目', '見出し') 'WBS の見出し')) {
    $field = ([string]$r.'項目').Trim()
    if ($field) { $candidates[$field] = @(([string]$r.'見出し') -split '\|' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ }) }
}
$fields = @('WBS番号', 'タスク', '担当', '予定開始', '予定終了', '予定工数', '実績開始', '実績終了', '見込終了', '進捗率', '実績工数')
foreach ($f in $fields) { if (-not $candidates.ContainsKey($f)) { $candidates[$f] = @() } }
foreach ($f in @('タスク', '予定開始', '予定終了')) { if ($candidates[$f].Count -eq 0) { throw "WBS の見出しに「$f」の行がありません: $Headers" } }

function ConvertTo-Date {
    #     日付の文字（2026/10/1・2026-10-01・2026年10月1日・10/1(水)）を日付にする（無ければ $null）。年の無い日付は -AsOf の前後半年の年にする
    #     読めない文字は ExcelTools.psm1 の ConvertTo-DateValue でも読む（和暦の令和8年10月1日・R8.10.1 など）
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return $null }
    try {
        $m = [regex]::Match($t, '(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})')
        if ($m.Success) { return New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value) }
        $m = [regex]::Match($t, '^(\d{1,2})\s*[/月]\s*(\d{1,2})')
        if ($m.Success) {
            $d = New-Object DateTime $AsOf.Year, ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value)
            if (($d - $AsOf).TotalDays -gt 183) { $d = $d.AddYears(-1) } elseif (($AsOf - $d).TotalDays -gt 183) { $d = $d.AddYears(1) }
            return $d
        }
    }
    catch { }
    return ConvertTo-DateValue $Text $AsOf
}

function ConvertTo-Number {
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim() -replace '[,\s]', ''
    $n = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $n }
    return $null
}

function ConvertTo-Rate {
    #     進捗率（45%・0.45・45・完了・完了済み・済み）を 0～1 にする
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return $null }
    if ($t -match '^(完了(済み?)?|済み?)$') { return 1.0 }
    $percent = $t.EndsWith('%')
    $n = ConvertTo-Number ($t.TrimEnd('%'))
    if ($null -eq $n) { return $null }
    if ($percent -or $n -gt 1) { $n = $n / 100 }
    return [math]::Max(0.0, [math]::Min(1.0, $n))
}

function Test-TotalRow {
    #     合計・小計の行か（WBS 番号かタスクが「合計」「小計」「計」「Total」など。「設計」のように計を含むだけのものは違う）
    param($Values)
    foreach ($text in @($Values['WBS番号'], $Values['タスク'])) {
        if ((([string]$text).Normalize($nfkc) -replace '\s', '') -match '^(総?合計|小計|総計|計|(sub|grand)?total)$') { return $true }
    }
    return $false
}

$holidaySet = New-Object 'System.Collections.Generic.HashSet[datetime]'
if ($Holidays) {
    # 内閣府の「国民の祝日」の CSV（Shift_JIS）など、1 列目が日付の CSV
    foreach ($line in ((Read-TextFile $Holidays) -split "\r?\n")) {
        $d = ConvertTo-Date (($line -split ',')[0].Trim('"'))
        if ($d) { [void]$holidaySet.Add($d.Date) }
    }
}

function Test-BusinessDay {
    param([datetime] $Day)
    return $Day.DayOfWeek -ne [DayOfWeek]::Saturday -and $Day.DayOfWeek -ne [DayOfWeek]::Sunday -and -not $holidaySet.Contains($Day.Date)
}

function Get-BusinessDays {
    #     From から To までの営業日の数（両端を含む）
    param([datetime] $From, [datetime] $To)
    $n = 0
    for ($d = $From.Date; $d -le $To.Date; $d = $d.AddDays(1)) { if (Test-BusinessDay $d) { $n++ } }
    return $n
}

function Get-BusinessDayDiff {
    #     A から B まで何営業日後か（B が前なら負）
    param([datetime] $A, [datetime] $B)
    if ($B -gt $A) { return Get-BusinessDays $A.AddDays(1) $B }
    if ($B -lt $A) { return -(Get-BusinessDays $B.AddDays(1) $A) }
    return 0
}

function Add-BusinessDays {
    param([datetime] $Day, [int] $Count)
    $d = $Day.Date
    while ($Count -gt 0) { $d = $d.AddDays(1); if (Test-BusinessDay $d) { $Count-- } }
    return $d
}

# ---------------------------------------------------------------- WBS を読む


$rows = New-Object System.Collections.Generic.List[object]

function Resolve-HeaderMap {
    #     見出しの行（1 行、または「予定」「実績」などの見出しの下に「開始」「終了」がある 2 行）から、項目 → 列番号
    param($Grid, [int] $Row, [int] $Upper)
    $maps = New-Object System.Collections.Generic.List[hashtable]
    if ($Upper) {
        # 上の行の見出しは、結合されていることが多いので右へ続くものとして下の行の見出しとつなげる
        $joined = @{}
        $upperCols = @($Grid.Cols[$Upper])
        foreach ($c in $Grid.Cols[$Row]) {
            $u = @($upperCols | Where-Object { $_ -le $c })
            if ($u.Count -gt 0) { $joined[$Grid.Labels[$Upper][$u[-1]] + $Grid.Labels[$Row][$c]] = $c }
        }
        $maps.Add($joined)
    }
    $plain = @{}
    foreach ($c in $Grid.Cols[$Row]) { $l = $Grid.Labels[$Row][$c]; if (-not $plain.ContainsKey($l)) { $plain[$l] = $c } }
    $maps.Add($plain)
    if ($Upper) {
        $above = @{}
        foreach ($c in $Grid.Cols[$Upper]) { $l = $Grid.Labels[$Upper][$c]; if (-not $above.ContainsKey($l)) { $above[$l] = $c } }
        $maps.Add($above)
    }
    $map = @{}; $used = @{}
    foreach ($f in $fields) {
        foreach ($cand in $candidates[$f]) {
            $col = $null
            foreach ($m in $maps) { if ($m.ContainsKey($cand) -and -not $used.ContainsKey($m[$cand])) { $col = $m[$cand]; break } }
            if ($col) { $map[$f] = $col; $used[$col] = $true; break }
        }
    }
    return $map
}

$files = New-Object System.Collections.Generic.List[string]
$csvFiles = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) { $files.Add($item.FullName) }
        elseif ($item.Extension -eq '.csv') { $csvFiles.Add($item) }
        else { $files.Add([System.Management.Automation.WildcardPattern]::Escape($item.FullName)) }
    }
}
if ($files.Count -gt 0) {
    $searchArgs = @{ Pattern = '\S'; Path = $files.ToArray(); Target = 'Value'; NoRecurse = $true }
    if ($SheetNamePattern) { $searchArgs['SheetNamePattern'] = $SheetNamePattern }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not (($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) -or ($OutFile -and [string]::Equals($_.Path, $OutFile, [System.StringComparison]::OrdinalIgnoreCase)) -or ($FillTemplate -and [string]::Equals($_.Path, $FillTemplate, [System.StringComparison]::OrdinalIgnoreCase))) })
    $books = ConvertTo-DesignBooks $hits
    foreach ($book in $books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            $header = $null; $map = $null; $prev = 0
            foreach ($row in $grid.Rows) {
                $m = Resolve-HeaderMap $grid $row 0
                if (-not ($m['タスク'] -and $m['予定開始'] -and $m['予定終了']) -and $prev -eq $row - 1) { $m = Resolve-HeaderMap $grid $row $prev }
                if ($m['タスク'] -and $m['予定開始'] -and $m['予定終了']) { $header = $row; $map = $m; break }
                $prev = $row
            }
            if (-not $header) { continue }
            foreach ($row in @($grid.Rows | Where-Object { $_ -gt $header })) {
                $cells = $grid.Cells[$row]
                $v = @{}
                foreach ($f in $map.Keys) { $v[$f] = [string]$cells[$map[$f]] }
                if (-not $v['タスク'].Trim()) { continue }
                $rows.Add([pscustomobject]@{ Values = $v; Place = [pscustomobject]@{ File = $book.File; Path = $book.Path; Sheet = $sheetName; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $map['タスク']), $row) } })
            }
        }
    }
}
foreach ($c in $csvFiles) {
    $data = Read-InputCsv $c.FullName 'WBS'
    if ($data.Count -eq 0) { continue }
    $names = @($data[0].PSObject.Properties | ForEach-Object { $_.Name })
    $map = @{}
    foreach ($f in $fields) { foreach ($cand in $candidates[$f]) { $n = @($names | Where-Object { (ConvertTo-Label $_) -eq $cand -and @($map.Values) -notcontains $_ })[0]; if ($n) { $map[$f] = $n; break } } }
    if (-not ($map['タスク'] -and $map['予定開始'] -and $map['予定終了'])) { Write-Warning ("{0}: タスク・予定開始・予定終了の列が見つかりません" -f $c.Name); continue }
    $line = 1
    foreach ($r in $data) {
        $line++
        $v = @{}
        foreach ($f in $map.Keys) { $v[$f] = [string]$r.($map[$f]) }
        if (-not $v['タスク'].Trim()) { continue }
        $rows.Add([pscustomobject]@{ Values = $v; Place = [pscustomobject]@{ File = $c.Name; Path = $c.FullName; Sheet = ''; Cell = ('{0} 行目' -f $line) } })
    }
}
if ($rows.Count -eq 0) { Write-Warning 'WBS（タスク・予定開始・予定終了の見出しのある表）が見つかりませんでした'; return }

# ---------------------------------------------------------------- タスクごとの計算


# 親のタスク（下に 1.1 などの子のタスクがある 1）は数えない。工程ごとのシートで番号を振り直した WBS もあるので、ファイル・シートごとに見る
$parents = @{}
foreach ($r in $rows) {
    $n = (([string]$r.Values['WBS番号']).Normalize($nfkc).Trim() -replace '[\-_]', '.').TrimEnd('.')
    $r | Add-Member -NotePropertyName Wbs -NotePropertyValue $n
    $i = $n.LastIndexOf('.')
    while ($i -gt 0) { $parents['{0}|{1}|{2}' -f $r.Place.Path, $r.Place.Sheet, $n.Substring(0, $i)] = $true; $i = $n.LastIndexOf('.', $i - 1) }
}

$tasks = New-Object System.Collections.Generic.List[object]
$skipped = 0
foreach ($r in $rows) {
    if ($r.Wbs -and $parents.ContainsKey(('{0}|{1}|{2}' -f $r.Place.Path, $r.Place.Sheet, $r.Wbs))) { continue }
    $v = $r.Values
    if (Test-TotalRow $v) { continue }
    $start = ConvertTo-Date $v['予定開始']
    $end = ConvertTo-Date $v['予定終了']
    if (-not $start -or -not $end) { $skipped++; continue }
    if ($end -lt $start) { $t = $start; $start = $end; $end = $t }
    $actualStart = ConvertTo-Date $v['実績開始']
    $actualEnd = ConvertTo-Date $v['実績終了']
    $forecastEnd = ConvertTo-Date $v['見込終了']
    $rate = ConvertTo-Rate $v['進捗率']
    $done = [bool]$actualEnd -or ($rate -eq 1.0)
    if ($actualEnd -and $actualEnd -gt $AsOf) {
        # 基準日より後に終わったタスクは、基準日には未完了とし、実績終了日を見込終了日にする。記入された進捗率は後の状態のため、
        # 実績開始日（無ければ予定開始日）から実績終了日までの営業日のうち、基準日までの割合を進捗率にする
        $from = $actualStart
        if (-not $from) { $from = $start }
        $rate = 0.0
        if ($AsOf -ge $from) { $rate = (Get-BusinessDays $from $AsOf) / [math]::Max(1, (Get-BusinessDays $from $actualEnd)) }
        $done = $false; $forecastEnd = $actualEnd; $actualEnd = $null
    }
    if ($done) { $rate = 1.0 }
    if ($null -eq $rate) { $rate = 0.0 }
    # 基準日より前の見込終了日は古い（予定終了日を入れたままのものなど）ので、未完了のタスクには使わない
    if (-not $done -and $forecastEnd -and $forecastEnd -lt $AsOf) { $forecastEnd = $null }
    $planDays = [math]::Max(1, (Get-BusinessDays $start $end))
    $planned = 0.0
    if ($AsOf -ge $end) { $planned = 1.0 } elseif ($AsOf -ge $start) { $planned = (Get-BusinessDays $start $AsOf) / $planDays }
    # PV・EV の重みは予定工数（0 のマイルストーンなどは 0）。予定工数が空なら予定の営業日数
    $effort = ConvertTo-Number $v['予定工数']
    $weight = $effort
    if ($null -eq $weight -or $weight -lt 0) { $weight = [double]$planDays }

    # 遅れ（営業日）: 完了は予定終了日との差、見込終了日（基準日以降のもの）があればその差、予定終了日を過ぎていれば
    # 過ぎた日数に残りの作業の日数（（1 − 進捗率）× 予定の営業日）を足したもの（進まなければ毎日増える）、
    # それ以外は予定の進捗との差を日数にしたもの
    $delay = 0
    if ($done -and $actualEnd) { $delay = Get-BusinessDayDiff $end $actualEnd }
    elseif (-not $done) {
        if ($forecastEnd) { $delay = Get-BusinessDayDiff $end $forecastEnd }
        elseif ($AsOf -gt $end) { $delay = (Get-BusinessDayDiff $end $AsOf) + [int][math]::Round((1 - $rate) * $planDays) }
        elseif ($planned -gt $rate) { $delay = [int][math]::Round(($planned - $rate) * $planDays) }
    }
    if ($done) { $status = '完了' }
    elseif ($delay -gt 0) { $status = '遅延' }
    elseif ($AsOf -lt $start) { $status = '未着手' }
    else { $status = '予定どおり' }
    # 実績／見込の終了日（遅延タスクの一覧・テンプレート用）
    $expectedEnd = $actualEnd
    if (-not $expectedEnd) { $expectedEnd = $forecastEnd }
    if (-not $expectedEnd -and $delay -gt 0) { $expectedEnd = Add-BusinessDays $end $delay }
    $actualStartShown = $actualStart
    if (-not $actualStartShown) { $actualStartShown = $start }

    $task = [pscustomobject]@{
        WbsNo = $r.Wbs; Task = $v['タスク'].Trim(); Owner = ([string]$v['担当']).Trim(); PlanStart = $start; PlanEnd = $end
        ActualStart = $actualStartShown; ExpectedEnd = $expectedEnd; Progress = [math]::Round($rate, 3); PlannedProgress = [math]::Round($planned, 3)
        Delay = $delay; Status = $status; PV = ($weight * $planned); EV = ($weight * $rate); AC = (ConvertTo-Number $v['実績工数']); BAC = $weight; Effort = $effort
        Done = $done; DoneOn = $actualEnd; Place = $r.Place
    }
    $task.PSObject.TypeNames.Insert(0, 'Progress.Task')
    $tasks.Add($task)
}
if ($skipped -gt 0) { Write-Warning ('予定開始日・予定終了日が読めないタスク {0} 件は数えていません' -f $skipped) }
if ($tasks.Count -eq 0) { Write-Warning 'WBS に、予定開始日・予定終了日の読めるタスクがありませんでした'; return }
$tasks

# ---------------------------------------------------------------- 集計


function Get-Ratio { param($A, $B) if ($B -and $B -gt 0) { return [math]::Round($A / $B, 2) } return $null }

$sumPV = ($tasks | Measure-Object PV -Sum).Sum
$sumEV = ($tasks | Measure-Object EV -Sum).Sum
$sumBAC = ($tasks | Measure-Object BAC -Sum).Sum
$withAC = @($tasks | Where-Object { $null -ne $_.AC })
$sumAC = $null; $sumEVwithAC = 0
if ($withAC.Count -gt 0) { $sumAC = ($withAC | Measure-Object AC -Sum).Sum; $sumEVwithAC = ($withAC | Measure-Object EV -Sum).Sum }
$spi = Get-Ratio $sumEV $sumPV
$cpi = $null
if ($null -ne $sumAC) { $cpi = Get-Ratio $sumEVwithAC $sumAC }
$delayed = @($tasks | Where-Object { $_.Status -eq '遅延' } | Sort-Object @{ e = { $_.Delay }; Descending = $true }, PlanEnd)

$owners = New-Object System.Collections.Generic.List[object]
# 担当者は、全角・半角、英字の大文字・小文字と空白の違いをそろえてまとめる（Get-OwnerKey。表示は最初に出てきた書き方。名前の順）
foreach ($g in @($tasks | Group-InOrder { Get-OwnerKey $_.Owner } | Sort-Object @{ e = { $_.Group[0].Owner } }, Index)) {
    $pv = ($g.Group | Measure-Object PV -Sum).Sum; $ev = ($g.Group | Measure-Object EV -Sum).Sum
    $owners.Add([pscustomobject]@{ Owner = $g.Group[0].Owner; Tasks = $g.Count; PV = $pv; EV = $ev; SPI = (Get-Ratio $ev $pv); Delayed = @($g.Group | Where-Object Status -eq '遅延').Count })
}

# 週ごとの計画完了・実績完了（月曜始まり）
$weeks = New-Object System.Collections.Generic.List[object]
$first = @($tasks | Sort-Object PlanStart)[0].PlanStart
$monday = { param($d) $d.Date.AddDays(-(([int]$d.DayOfWeek + 6) % 7)) }
$w = & $monday $first
$last = & $monday $AsOf
$cumPlan = 0; $cumDone = 0; $no = 0
while ($w -le $last) {
    $no++
    $wEnd = $w.AddDays(6)
    $plan = @($tasks | Where-Object { $_.PlanEnd -ge $w -and $_.PlanEnd -le $wEnd }).Count
    $done = @($tasks | Where-Object { $_.DoneOn -and $_.DoneOn -ge $w -and $_.DoneOn -le $wEnd }).Count
    $cumPlan += $plan; $cumDone += $done
    $weeks.Add([pscustomobject]@{ No = ('W{0}' -f $no); Monday = $w; Planned = $plan; Done = $done; TotalPlanned = $cumPlan; TotalDone = $cumDone })
    $w = $w.AddDays(7)
}

# ---------------------------------------------------------------- スケジュール遅延リカバリ計画のテンプレートに書く


function Find-HeaderColumns {
    #     テンプレートの表の見出しの行と、見出し → 列番号
    param($Grid, [string[]] $Labels)
    foreach ($row in $Grid.Rows) {
        $colOf = @{}
        foreach ($c in $Grid.Cols[$row]) { $colOf[$Grid.Labels[$row][$c]] = $c }
        if (@($Labels | Where-Object { -not $colOf.ContainsKey((ConvertTo-Label $_)) }).Count -eq 0) {
            $map = @{}
            foreach ($l in $Labels) { $map[$l] = $colOf[(ConvertTo-Label $l)] }
            return [pscustomobject]@{ Row = $row; Cols = $map }
        }
    }
    return $null
}

if ($FillTemplate) {
    $tplBooks = ConvertTo-DesignBooks @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FillTemplate)) -Target Value -NoExcel 6>$null)
    $tplBook = @($tplBooks.Values)[0]
    $pkg = Open-ZipPackage $FillTemplate
    $changes = @{}
    try {
        $parts = Get-WorkbookSheetParts $pkg
        $written = @{}
        $plans = @(
            @{ Sheet = '02_遅延状況'; Labels = @('No', 'WBS No', '工程・タスク名', '担当', '当初開始日', '当初終了日', '実績／見込開始日', '実績／見込終了日', '進捗率'); Items = $delayed
                Values = { param($t, $i) @{ 'No' = [double]$i; 'WBS No' = $t.WbsNo; '工程・タスク名' = $t.Task; '担当' = $t.Owner; '当初開始日' = (ConvertTo-ExcelDate $t.PlanStart); '当初終了日' = (ConvertTo-ExcelDate $t.PlanEnd)
                        '実績／見込開始日' = (ConvertTo-ExcelDate $t.ActualStart); '実績／見込終了日' = $(if ($t.ExpectedEnd) { ConvertTo-ExcelDate $t.ExpectedEnd } else { $null }); '進捗率' = [double]$t.Progress } } },
            @{ Sheet = '08_週次モニタリング'; Labels = @('週No', '対象週（月曜）', '計画完了タスク数', '実績完了タスク数'); Items = $weeks
                Values = { param($wk, $i) @{ '週No' = $wk.No; '対象週（月曜）' = (ConvertTo-ExcelDate $wk.Monday); '計画完了タスク数' = [double]$wk.Planned; '実績完了タスク数' = [double]$wk.Done } } }
        )
        foreach ($plan in $plans) {
            if (-not $parts.Sheets.Contains($plan.Sheet) -or -not $tplBook.Sheets.Contains($plan.Sheet)) { Write-Warning ("テンプレートにシート「{0}」がありません" -f $plan.Sheet); continue }
            $grid = $tplBook.Sheets[$plan.Sheet]
            $h = Find-HeaderColumns $grid $plan.Labels
            if (-not $h) { Write-Warning ("テンプレートの「{0}」に見出し（{1}）が見つかりません" -f $plan.Sheet, ($plan.Labels -join '・')); continue }
            $doc = Read-PartXmlForEdit $pkg $parts.Sheets[$plan.Sheet]
            # 書く行: 見出しの下で、2 列目（WBS No・対象週）が空の行から（記入例の行は残す）
            $keyCol = $h.Cols[$plan.Labels[1]]
            $row = $h.Row + 1
            while ($grid.Cells.ContainsKey($row) -and $grid.Cells[$row].ContainsKey($keyCol)) { $row++ }
            $lastFormatted = $row
            foreach ($el in (Get-XmlChildElements (@(Get-XmlChildElements $doc.DocumentElement 'sheetData')[0]) 'row')) { $lastFormatted = [math]::Max($lastFormatted, [int]$el.GetAttribute('r')) }
            $i = 0
            foreach ($item in $plan.Items) {
                if ($row -gt $lastFormatted) { Write-Warning ("「{0}」の書式のある行が足りないため、{1} 件目からは書いていません" -f $plan.Sheet, ($i + 1)); break }
                $i++
                $vals = & $plan.Values $item $i
                $rowEl = Get-SheetRowElement $doc $row -Create
                foreach ($l in $plan.Labels) {
                    $cell = Get-SheetCellElement $doc $rowEl $row $h.Cols[$l] -Create
                    Set-SheetCellValue $doc $cell $vals[$l]
                }
                $row++
            }
            $written[$plan.Sheet] = $i
            $changes[$parts.Sheets[$plan.Sheet].ToLowerInvariant()] = ConvertTo-XmlString $doc
        }
        # 遅延日数・累計などの数式を、開いたときに計算させる
        $entry = Get-PackageEntry -Package $pkg -PartName $parts.WorkbookPart
        $sr = New-Object System.IO.StreamReader($entry.Open())
        try { $wbXml = $sr.ReadToEnd() } finally { $sr.Dispose() }
        $changes[$parts.WorkbookPart.ToLowerInvariant()] = Set-FullCalcOnLoad $wbXml
    }
    finally { $pkg.Zip.Dispose() }
    Save-ZipPackage -FilePath $FillTemplate -Changes $changes -OutPath $OutFile
}

if ($OutXlsx) {
    # 日付は日付のセルにする（Excel で並べ替え・計算ができる）
    $taskRows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $tasks) {
        $exp = ''
        if ($t.ExpectedEnd) { $exp = $t.ExpectedEnd.Date }
        $row = @{ Values = @($t.WbsNo, $t.Task, $t.Owner, $t.PlanStart.Date, $t.PlanEnd.Date, $exp, [double]$t.Progress, [double]$t.PlannedProgress, [double]$t.Delay, $t.Status, [double][math]::Round($t.PV, 2), [double][math]::Round($t.EV, 2), ('{0} {1}' -f $t.Place.File, $t.Place.Cell)) }
        if ($t.Place.Sheet) { $row.Link = @{ Column = 12; Path = $t.Place.Path; Location = ("'{0}'!{1}" -f $t.Place.Sheet.Replace("'", "''"), $t.Place.Cell) } }
        $taskRows.Add($row)
    }
    $ownerRows = New-Object System.Collections.Generic.List[object]
    foreach ($o in $owners) {
        $spiV = ''
        if ($null -ne $o.SPI) { $spiV = [double]$o.SPI }
        $ownerRows.Add(@{ Values = @($o.Owner, [double]$o.Tasks, [double][math]::Round($o.PV, 2), [double][math]::Round($o.EV, 2), $spiV, [double]$o.Delayed) })
    }
    $weekRows = New-Object System.Collections.Generic.List[object]
    foreach ($wk in $weeks) { $weekRows.Add(@{ Values = @($wk.No, $wk.Monday.Date, [double]$wk.Planned, [double]$wk.Done, [double]$wk.TotalPlanned, [double]$wk.TotalDone) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = 'タスク'; Columns = @(@{ Header = 'WBS'; Width = 8 }, @{ Header = 'タスク'; Width = 30 }, @{ Header = '担当'; Width = 10 }, @{ Header = '予定開始'; Width = 11 }, @{ Header = '予定終了'; Width = 11 }, @{ Header = '実績／見込終了'; Width = 13 }, @{ Header = '進捗率'; Width = 7 }, @{ Header = '予定の進捗率'; Width = 10 }, @{ Header = '遅れ（営業日）'; Width = 12 }, @{ Header = '状態'; Width = 10 }, @{ Header = 'PV'; Width = 8 }, @{ Header = 'EV'; Width = 8 }, @{ Header = 'WBS の場所'; Width = 30 }); Rows = $taskRows },
            @{ Name = '担当者別'; Columns = @(@{ Header = '担当'; Width = 12 }, @{ Header = 'タスク'; Width = 7 }, @{ Header = 'PV'; Width = 9 }, @{ Header = 'EV'; Width = 9 }, @{ Header = 'SPI'; Width = 7 }, @{ Header = '遅延タスク'; Width = 9 }); Rows = $ownerRows },
            @{ Name = '週次'; Columns = @(@{ Header = '週'; Width = 6 }, @{ Header = '対象週（月曜）'; Width = 13 }, @{ Header = '計画完了'; Width = 9 }, @{ Header = '実績完了'; Width = 9 }, @{ Header = '計画完了の累計'; Width = 13 }, @{ Header = '実績完了の累計'; Width = 13 }); Rows = $weekRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ('タスク {0} 件（{1} 時点。親のタスクは除く）: 完了 {2}・遅延 {3}・予定どおり {4}・未着手 {5}' -f $tasks.Count, $AsOf.ToString('yyyy/MM/dd', $invariant),
    @($tasks | Where-Object Status -eq '完了').Count, $delayed.Count, @($tasks | Where-Object Status -eq '予定どおり').Count, @($tasks | Where-Object Status -eq '未着手').Count)
$line = 'PV {0:0.##} / EV {1:0.##} / BAC {2:0.##}' -f $sumPV, $sumEV, $sumBAC
if ($null -ne $spi) { $line += ' / SPI {0:0.00}' -f $spi }
if ($null -ne $sumAC) { $line += ' / AC {0:0.##} / CPI {1:0.00}' -f $sumAC, $cpi }
Write-Host $line
if ($delayed.Count -gt 0) { Write-Host ('遅延 {0} 件（最大 {1} 営業日: {2}）' -f $delayed.Count, $delayed[0].Delay, $delayed[0].Task) }
if ($FillTemplate) { Write-Host ('テンプレートに書き込みました: {0}（02_遅延状況 {1} 件・08_週次モニタリング {2} 週）' -f $OutFile, $written['02_遅延状況'], $written['08_週次モニタリング']) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     WBS（Excel・CSV）から進捗を集計し、遅延タスク・PV／EV／SPI・週ごとの完了数を出します（Excel 不要）。
#
# .DESCRIPTION
#     WBS の表を見出しの文字で読み（「予定」「実績」の下に「開始」「終了」がある 2 行の見出しも読めます）、
#     子のタスクの無いタスクごとに、-AsOf の時点の予定の進捗率・遅れ（営業日）・状態を出します。
#
#       遅れ        完了したタスクは予定終了日との差、見込終了日（基準日以降のもの）があればその差、予定終了日を過ぎていれば
#                   その日数に残りの作業の日数（（1 − 進捗率）× 予定の営業日）を足したもの、それ以外は予定の進捗率との差を
#                   日数にしたもの（土日と -Holidays の休日を除く営業日）。基準日より後に終わったタスクは、基準日には未完了とします
#       PV / EV     予定工数（無ければ予定の営業日数）× 予定の進捗率 / 進捗率。SPI = EV / PV
#       AC / CPI    実績工数の列があれば、CPI = EV / AC
#
#     -FillTemplate にスケジュール遅延リカバリ計画のテンプレートを渡すと、遅延タスクを「02_遅延状況」に、
#     週ごとの計画完了・実績完了の数を「08_週次モニタリング」に書き込んで、-OutFile に保存します。
#
# .PARAMETER Path
#     WBS のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。
#
# .PARAMETER AsOf
#     進捗を見る日（基準日）。既定は今日。
#
# .PARAMETER Holidays
#     休日の CSV（1 列目が日付）。内閣府の「国民の祝日」の CSV（syukujitsu.csv）をそのまま使えます。
#
# .PARAMETER SheetNamePattern
#     WBS のシート名の正規表現（WBS の表のあるシートが複数あるときに絞る）。
#
# .PARAMETER Headers
#     WBS の見出しの CSV（項目と、見出しの候補）。既定は tools\progress\WBSの見出し.csv。
#
# .PARAMETER FillTemplate
#     スケジュール遅延リカバリ計画のテンプレート（docs\schedule-recovery のもの）。テンプレート自体は変更しません。
#
# .PARAMETER OutFile
#     -FillTemplate で書き込んだテンプレートの保存先（.xlsx）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（タスク・担当者別・週次。タスクから WBS の該当行へのリンク付き）。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\Measure-Progress.ps1 D:\PJ\WBS.xlsx -Holidays .\syukujitsu.csv
#
# .EXAMPLE
#     # 遅延タスクと週ごとの完了数を、スケジュール遅延リカバリ計画のテンプレートに書き込む
#     .\Measure-Progress.ps1 D:\PJ\WBS.xlsx -AsOf 2026-10-20 -FillTemplate ..\docs\schedule-recovery\スケジュール遅延リカバリ計画_テンプレート.xlsx -OutFile D:\PJ\リカバリ計画.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     親のタスクは、ファイル・シートごとに WBS 番号（1・1.1・1.1.1 など）で見分けます。WBS 番号の列が無いときは、すべての行をタスクとして数えます。
#     合計・小計の行（WBS 番号かタスクが「合計」「小計」「Total」など）は数えません。
