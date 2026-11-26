#     Measure-Progress.ps1 のテスト。WBS はテストの中で作る（Export-Workbook）

$ProgressTool = Join-Path $script:ToolsDir 'Measure-Progress.ps1'
$recoveryTemplate = Join-Path $script:RepoRoot 'docs/schedule-recovery/スケジュール遅延リカバリ計画_テンプレート.xlsx'

function New-Wbs {
    #     「予定」「実績」の下に「開始」「終了」がある 2 行の見出しの WBS と、Shift_JIS の祝日の CSV（10/12 スポーツの日）
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('WBS', 'タスク名', '担当者', '予定', '', '実績', '', '進捗率', '予定工数', '実績工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $rows = @(
        @('', '', '', '開始', '終了', '開始', '終了', '', '', ''),
        @('1', '基本設計', '', '2026/10/01', '2026/10/16', '', '', '', '', ''),
        @('1.1', '画面設計', '山田', '2026/10/01', '2026/10/07', '2026/10/01', '2026/10/09', '100%', '5', '7'),
        @('1.2', 'テーブル設計', '佐藤', '2026/10/05', '2026/10/09', '2026/10/05', '', '60%', '5', '4'),
        @('1.3', 'IF 設計', '佐藤', '2026/10/12', '2026/10/16', '', '', '0%', '4', ''),
        @('2', '詳細設計', '', '2026/10/19', '2026/10/30', '', '', '', '', ''),
        @('2.1', '画面詳細', '山田', '2026/10/19', '2026/10/30', '', '', '', '8', ''))
    $wbs = Join-Path $Dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $cols; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    $holidays = Join-Path $Dir 'syukujitsu.csv'
    [System.IO.File]::WriteAllText($holidays, "国民の祝日・休日月日,国民の祝日・休日名称`r`n2026/10/12,スポーツの日`r`n", [System.Text.Encoding]::GetEncoding(932))
    return [pscustomobject]@{ Wbs = $wbs; Holidays = $holidays }
}

function Read-GanttCells {
    #     ブックの Index 番目のシートのセル（セル番地 → @{ Style; Formula; Value（数式のセルは計算済みの値）}）と、条件付き書式の数式・シート名
    param([string] $Path, [int] $Index = 1)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $load = { param($n) $sr = New-Object System.IO.StreamReader($zip.GetEntry($n).Open()); try { $x = New-Object System.Xml.XmlDocument; $x.LoadXml($sr.ReadToEnd()); $x } finally { $sr.Dispose() } }
        $doc = & $load "xl/worksheets/sheet$Index.xml"
        $wb = & $load 'xl/workbook.xml'
    }
    finally { $fs.Dispose() }
    $cells = @{}
    foreach ($c in $doc.GetElementsByTagName('c')) {
        $f = @($c.GetElementsByTagName('f'))[0]; $v = @($c.GetElementsByTagName('v'))[0]
        $value = ''
        if ($c.GetAttribute('t') -eq 'inlineStr') { $value = $c.InnerText } elseif ($v) { $value = $v.InnerText }
        $formula = ''; if ($f) { $formula = $f.InnerText }
        $cells[$c.GetAttribute('r')] = [pscustomobject]@{ Style = $c.GetAttribute('s'); Formula = $formula; Value = $value }
    }
    return [pscustomobject]@{
        Cells = $cells
        Rules = @($doc.GetElementsByTagName('cfRule') | ForEach-Object { '{0}|{1}|{2}' -f $_.ParentNode.GetAttribute('sqref'), $_.GetAttribute('dxfId'), $_.InnerText })
        Sheets = @($wb.GetElementsByTagName('sheet') | ForEach-Object { $_.GetAttribute('name') })
        Names = @($wb.GetElementsByTagName('definedName') | Where-Object { -not $_.GetAttribute('localSheetId') } | ForEach-Object { '{0}={1}' -f $_.GetAttribute('name'), $_.InnerText })
        Frozen = @($doc.GetElementsByTagName('pane') | ForEach-Object { '{0}|{1}|{2}' -f $_.GetAttribute('xSplit'), $_.GetAttribute('ySplit'), $_.GetAttribute('topLeftCell') })[0]
    }
}

Test-Case 'Measure-Progress: 2 行の見出しの WBS から、子のタスクごとの予定の進捗・遅れ（営業日）・状態と PV／EV／SPI／CPI を出す' {
    $f = New-Wbs (New-TestDir)
    $r = @(& $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays 6>$null)
    Assert-Equal @(
        '1.1|画面設計|山田|1|1|2|完了',
        '1.2|テーブル設計|佐藤|0.6|1|4|遅延',
        '1.3|IF 設計|佐藤|0|0.5|2|遅延',
        '2.1|画面詳細|山田|0|0|0|未着手') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $_.WbsNo, $_.Task, $_.Owner, $_.Progress, $_.PlannedProgress, $_.Delay, $_.Status }) '親のタスク（1・2）は数えない。10/12 は祝日。1.2 は予定終了日を過ぎた 2 営業日と残りの作業の 2 営業日'
    Assert-Equal '2026/10/16|2026/10/20' ('{0:yyyy/MM/dd}|{1:yyyy/MM/dd}' -f $r[1].ExpectedEnd, $r[2].ExpectedEnd) '見込の終了日（予定終了日 + 遅れの営業日）'
    $text = Get-HostText { & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays | Out-Null }
    Assert-Match $text 'タスク 4 件（2026/10/14 時点。親のタスクは除く）: 完了 1・遅延 2・予定どおり 0・未着手 1'
    Assert-Match $text 'PV 12 / EV 8 / BAC 22 / SPI 0\.67 / AC 11 / CPI 0\.73'
    Assert-Match $text '遅延 2 件（最大 4 営業日: テーブル設計）'
}

Test-Case 'Measure-Progress: 完了見込み（遅れを、同じ担当者の後のタスクに送る）と -OutXlsx の見込の列' {
    $dir = New-TestDir
    $f = New-Wbs $dir
    $out = Join-Path $dir '進捗.xlsx'
    $r = @(& $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -OutXlsx $out 6>$null)
    Assert-Equal @('1.1|10/01|10/09|2', '1.2|10/05|10/16|4', '1.3|10/19|10/22|4', '2.1|10/19|10/30|0') @($r | ForEach-Object { '{0}|{1:MM/dd}|{2:MM/dd}|{3}' -f $_.WbsNo, $_.ForecastStart, $_.ForecastEnd, $_.ForecastDelay }) '1.3 は佐藤さんの 1.2（10/16 まで）の翌営業日から 4 営業日。2.1 は予定どおり'
    $text = Get-HostText { & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays | Out-Null }
    Assert-Match $text '完了見込み: 2026/10/30（予定どおり）'
    Assert-Match $text '予定終了より遅れる見込みのタスク 2 件(?!（)'
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal '見込開始|見込終了|見込の遅れ（営業日）' ((@('K', 'L', 'M') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal ([string](([datetime]'2026-10-22') - [datetime]'1899-12-30').TotalDays) $book.Rows[4]['L']
    Assert-True (@($book.Links | Where-Object { $_.Ref -like 'P*' }).Count -eq 4) 'WBS の場所へのリンクは P 列'
    # 予定・実績と完了見込みを並べたガントチャート（基準日は -AsOf のまま）。見込終了は、終わっていないタスクの完了見込み
    $g = Read-GanttCells $out 4
    Assert-Equal @('タスク', '担当者別', '週次', 'ガントチャート', '休日') $g.Sheets
    Assert-Equal @('進捗（2026/10/14 時点の見込み）', '', '46309') @($g.Cells['A1'].Value, $g.Cells['L1'].Formula, $g.Cells['L1'].Value)
    Assert-Equal '見込終了' $g.Cells['I3'].Value
    Assert-Equal @('', '46311', '46317', '46325') @(5, 6, 7, 9 | ForEach-Object { $g.Cells["I$_"].Value }) '1.1 は完了。1.2 は 10/16、1.3 は佐藤さんの 1.2 の後の 10/22、2.1 は予定どおり 10/30'
    Assert-Equal '遅延' $g.Cells['L7'].Value
}

Test-Case 'Measure-Progress: 先行の列（親の番号・3FS+2 の書き方・別の担当者）から遅れを後のタスクに送り、完了見込みを決めている流れを出す' {
    $dir = New-TestDir
    $csv = Join-Path $dir 'wbs.csv'
    @('WBS番号,タスク名,担当者,先行タスク,予定開始,予定終了,進捗率',
        '1,設計,,,2026/10/05,2026/10/12,',
        '1.1,画面設計,山田,,2026/10/05,2026/10/09,40%',
        '1.2,テーブル設計,佐藤,,2026/10/05,2026/10/07,100%',
        '1.3,画面レビュー,佐藤,1.1,2026/10/12,2026/10/12,',
        '2,製造,,,2026/10/13,2026/10/16,',
        '2.1,画面製造,山田,1FS+2日,2026/10/13,2026/10/16,',
        '2.2,テーブル製造,佐藤,9.9,2026/10/13,2026/10/14,') | Set-Content -LiteralPath $csv -Encoding UTF8
    $all = @(& $ProgressTool -Path $csv -AsOf '2026-10-09' 6>$null 3>&1)
    $r = @($all | Where-Object { $_ -isnot [System.Management.Automation.WarningRecord] })
    Assert-Equal @('1.1|10/05|10/14|3', '1.2|10/05|10/07|0', '1.3|10/15|10/15|3', '2.1|10/16|10/21|3', '2.2|10/16|10/19|3') @($r | ForEach-Object { '{0}|{1:MM/dd}|{2:MM/dd}|{3}' -f $_.WbsNo, $_.ForecastStart, $_.ForecastEnd, $_.ForecastDelay }) '1.1 は基準日 + 残り 3 営業日。1.3 は 1.1 の翌営業日（予定で後の日）。2.1 は親の 1 のすべて（1.3）の後。2.2 は佐藤さんの 1.3 の後'
    Assert-Match (@($all | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }) -join "`n") '先行の列の WBS 番号が見つかりません.*2\.2 の 9\.9'
    $text = Get-HostText { & $ProgressTool -Path $csv -AsOf '2026-10-09' 3>$null | Out-Null }
    Assert-Match $text '完了見込み: 2026/10/21（予定の 2026/10/16 より 3 営業日遅れ）'
    Assert-Match $text '予定終了より遅れる見込みのタスク 4 件（うち 3 件は、まだ遅れていないが'
    Assert-Match $text '完了見込みを決めている流れ: 1\.1 画面設計（山田） → 1\.3 画面レビュー（佐藤） → 2\.1 画面製造（山田）'

    # 親の番号を先行にしたタスクがたくさんある WBS（工程の待ち合わせ。くり返しが多くても計算できる）
    $lines = @('WBS番号,タスク名,担当者,先行,予定開始,予定終了,進捗率', '1,設計,,,2026/10/05,2026/10/16,', '2,テスト,,,2026/10/19,2026/10/30,')
    for ($i = 1; $i -le 30; $i++) { $lines += '1.{0},設計{0},担当{1},,2026/10/{2:00},2026/10/{2:00},' -f $i, ($i % 3), (5 + ($i % 5)) }
    for ($i = 1; $i -le 30; $i++) { $lines += '2.{0},テスト{0},担当{1},1,2026/10/19,2026/10/20,' -f $i, ($i % 3) }
    $lines[3] = '1.1,設計1,担当1,,2026/10/06,2026/10/06,0%'
    $lines | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = @(& $ProgressTool -Path $csv -AsOf '2026-10-14' 6>$null 3>$null)
    $lastDesign = @($r | Where-Object { $_.WbsNo -like '1.*' } | Sort-Object ForecastEnd -Descending)[0].ForecastEnd
    Assert-True ($lastDesign -ge [datetime]'2026-10-15') '遅れた設計がある'
    Assert-True (@($r | Where-Object { $_.WbsNo -like '2.*' -and $_.ForecastStart -le $lastDesign }).Count -eq 0) '親の 1 を先行にしたテストは、設計がすべて終わった翌営業日から'

    @('WBS番号,タスク名,担当者,先行,予定開始,予定終了', '1,A,,2,2026/10/05,2026/10/06', '2,B,,1,2026/10/07,2026/10/08') | Set-Content -LiteralPath $csv -Encoding UTF8
    $warnings = @(& $ProgressTool -Path $csv -AsOf '2026-10-01' 6>$null 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert-Match ($warnings -join "`n") '先行タスクが輪になっています'
}

Test-Case 'Measure-Progress: 日付の書式の無いセルの日付（Excel のシリアル値。値だけ貼り付けたもの）も日付として読む' {
    $csv = Join-Path (New-TestDir) 'wbs.csv'
    @('タスク名,担当者,予定開始,予定終了,実績開始,実績終了,進捗率', '画面設計,山田,46300,46304,46300,46307,100%', 'テーブル設計,佐藤,2026/10/05,46304.0,,,50%') | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = @(& $ProgressTool -Path $csv -AsOf '2026-10-14' 6>$null)
    Assert-Equal @('画面設計|10/05|10/09|10/12|1|完了', 'テーブル設計|10/05|10/09||5|遅延') @($r | ForEach-Object { '{0}|{1:MM/dd}|{2:MM/dd}|{3:MM/dd}|{4}|{5}' -f $_.Task, $_.PlanStart, $_.PlanEnd, $_.DoneOn, $_.Delay, $_.Status })
}

Test-Case 'Measure-Progress: CSV の WBS（WBS 番号なし）と見込終了日を読む' {
    $dir = New-TestDir
    $csv = Join-Path $dir 'wbs.csv'
    @('作業,担当,開始予定日,終了予定日,進捗,終了見込', '要件定義,田中,2026/10/01,2026/10/09,80,2026/10/13', '外部設計,田中,2026/10/12,2026/10/23,10,') | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = @(& $ProgressTool -Path $csv -AsOf '2026-10-12' 6>$null)
    Assert-Equal @('要件定義|0.8|2|遅延', '外部設計|0.1|0|予定どおり') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Task, $_.Progress, $_.Delay, $_.Status }) '見込終了日 10/13 は予定終了日 10/09 の 2 営業日後'
}

Test-Case 'Measure-Progress: -FillTemplate で遅延タスクと週ごとの完了数をスケジュール遅延リカバリ計画のテンプレートに書き込む' {
    $dir = New-TestDir
    $f = New-Wbs $dir
    $out = Join-Path $dir 'recovery.xlsx'
    $hash = (Get-FileHash -LiteralPath $recoveryTemplate).Hash
    $outXlsx = Join-Path $dir 'progress.xlsx'
    $null = & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -FillTemplate $recoveryTemplate -OutFile $out -OutXlsx $outXlsx 6>$null
    Assert-Equal $hash (Get-FileHash -LiteralPath $recoveryTemplate).Hash 'テンプレートは変えない'
    $cells = @{}
    foreach ($h in @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($out)) -Target Value -SheetNamePattern '^(02_遅延状況|08_週次モニタリング)$' 6>$null)) { $cells['{0}!{1}' -f $h.Sheet, $h.Cell] = $h.Text }
    Assert-Equal '3.2.1' $cells['02_遅延状況!B7'] '記入例の行は残す'
    Assert-Equal '1|1.2|テーブル設計|佐藤|2026/10/05|2026/10/09|2026/10/05|2026/10/16' ((@('A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8') | ForEach-Object { $cells["02_遅延状況!$_"] }) -join '|')
    Assert-Equal '2|1.3|IF 設計' ((@('A9', 'B9', 'C9') | ForEach-Object { $cells["02_遅延状況!$_"] }) -join '|')
    # テンプレートの遅延日数は暦日（10/09 → 10/16 は 7）。このツールの遅れ（営業日）を備考に書く
    Assert-Equal '遅延日数（暦日）|遅れ 4 営業日（土日・休日を除く）|遅れ 2 営業日（土日・休日を除く）' ((@('I6', 'N8', 'N9') | ForEach-Object { $cells["02_遅延状況!$_"] }) -join '|')
    Assert-Equal 'W2|2026/10/05|2|1' ((@('A7', 'B7', 'C7', 'D7') | ForEach-Object { $cells["08_週次モニタリング!$_"] }) -join '|')
    Assert-Equal 'W3|1|0' ((@('A8', 'C8', 'D8') | ForEach-Object { $cells["08_週次モニタリング!$_"] }) -join '|')
    # 遅延日数などの数式を開いたときに計算させる
    $fs = [System.IO.File]::OpenRead($out)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $sr = New-Object System.IO.StreamReader($zip.GetEntry('xl/workbook.xml').Open())
        $wbXml = $sr.ReadToEnd(); $sr.Dispose()
    }
    finally { $fs.Dispose() }
    Assert-Match $wbXml 'fullCalcOnLoad="1"'
    Assert-Throws { & $ProgressTool -Path $f.Wbs -FillTemplate $recoveryTemplate -OutFile $out 6>$null } 'すでにあります'
    $sheet = Read-XlsxSheet -Path $outXlsx -Index 1
    Assert-Equal @('タスク', '担当者別', '週次', 'ガントチャート', '休日') $sheet.Sheets
    Assert-Equal 5 $sheet.Rows.Count '見出しと 4 件'
}

Test-Case 'Measure-Progress: 和暦の設定（ja-JP の JapaneseCalendar）でも、日付は西暦で書き、-OutXlsx の日付は日付のセルにする' {
    $dir = New-TestDir
    $f = New-Wbs $dir
    $out = Join-Path $dir '進捗.xlsx'
    $text = Invoke-InWareki { Get-HostText { & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -OutXlsx $out | Out-Null } }
    Assert-Match $text 'タスク 4 件（2026/10/14 時点'
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-True (@($book.Rows.Values | ForEach-Object { $_.Values } | Where-Object { "$_" -match '^0?8/10/' }).Count -eq 0) '和暦の年（08/10/14）にしない'
    $serial = { param($Day) [string](([datetime]$Day) - [datetime]'1899-12-30').TotalDays }
    Assert-Equal ((@('2026-10-01', '2026-10-07', '2026-10-09') | ForEach-Object { & $serial $_ }) -join '|') ((@('D', 'E', 'F') | ForEach-Object { $book.Rows[2][$_] }) -join '|') 'タスクの予定開始・予定終了・実績／見込終了は日付のセル（シリアル値）'
    $weeks = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal (& $serial '2026-09-28') $weeks.Rows[2]['B'] '週次の対象週（月曜）も日付のセル'
}

Test-Case 'Measure-Progress: 予定終了日を過ぎた未完了のタスクは、過ぎた営業日に残りの作業の日数を足して遅れにし、基準日より前の見込終了日は使わない' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '見込終了日', '進捗率', '予定工数'),
            @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '2026/10/09', '60%', '5'),
            @('1.2', 'テーブル設計', '佐藤', '2026/10/05', '2026/10/16', '', '30%', '10'),
            @('1.3', 'IF 設計', '佐藤', '2026/10/05', '2026/10/09', '2026/10/23', '50%', '5')) })
    $delays = { param($Day) @(& $ProgressTool -Path $wbs -AsOf $Day 6>$null | ForEach-Object { '{0}|{1}|{2}|{3:yyyy/MM/dd}' -f $_.WbsNo, $_.Delay, $_.Status, $_.ExpectedEnd }) }
    Assert-Equal @('1.1|7|遅延|2026/10/20', '1.2|7|遅延|2026/10/27', '1.3|10|遅延|2026/10/23') (& $delays '2026-10-16') '見込終了日に予定終了日を入れたままの 1.1 も遅れる。基準日より後の見込終了日（1.3）は使う'
    Assert-Equal @('1.1|8|遅延|2026/10/21', '1.2|8|遅延|2026/10/28', '1.3|10|遅延|2026/10/23') (& $delays '2026-10-19') '進まなければ、予定終了日を過ぎても遅れは増える'
}

Test-Case 'Measure-Progress: 和暦の日付（令和8年10月5日・R8.10.12）を読む' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '実績終了日', '進捗率'),
            @('1.1', '画面設計', '山田', '令和8年10月5日', '令和8年10月9日', '令和8年10月9日', '100%'),
            @('1.2', 'テーブル設計', '佐藤', 'R8.10.12', 'R8.10.16', '', '40%')) })
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' 6>$null)
    Assert-Equal @('1.1|2026/10/05|2026/10/09|完了', '1.2|2026/10/12|2026/10/16|遅延') @($r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}|{3}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd, $_.Status })
}

Test-Case 'Measure-Progress: 日付の読めるタスクが無ければ警告して終わり、New-WeeklyReport・Measure-Effort は WBS のタスクが無いものとして続ける' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '予定工数'), @('1.1', '画面設計', '山田', '', '', '5'), @('1.2', 'テーブル設計', '佐藤', '未定', '未定', '5')) })
    $warnings = @()
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' -WarningVariable +warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal 0 $r.Count
    Assert-Match ($warnings -join ' ') '読めないタスク 2 件.*読めるタスクがありませんでした'
    $text = Get-HostText { & (Join-Path $script:ToolsDir 'New-WeeklyReport.ps1') -Wbs $wbs -AsOf '2026-10-14' 3>$null | Out-Null }
    Assert-Match $text '・WBS のタスクが読めませんでした'
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    @('日付,氏名,WBS,時間', '2026/10/05,山田,1.1,8') | Set-Content -LiteralPath (Join-Path $dir '工数/工数実績.csv') -Encoding UTF8
    $text = Get-HostText { & (Join-Path $script:ToolsDir 'Measure-Effort.ps1') (Join-Path $dir '工数') -Wbs $wbs -AsOf '2026-10-14' -Unit 時間 3>$null | Out-Null }
    Assert-Match $text '工数の実績: 8 時間'
}

Test-Case 'Measure-Progress: 合計・小計の行（WBS 番号かタスクが「合計」「Sub Total」など）は数えない' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数'),
            @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '100%', '5'),
            @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '50%', '5'),
            @('1.3', '設計', '鈴木', '2026/10/19', '2026/10/23', '0%', '4'),
            @('Sub Total', '基本設計', '', '2026/10/05', '2026/10/23', '', '14'),
            @('', '合　計', '', '2026/10/05', '2026/10/23', '', '14')) })
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14') }
    Assert-Equal @('1.1|画面設計', '1.2|テーブル設計', '1.3|設計') @($script:r | ForEach-Object { '{0}|{1}' -f $_.WbsNo, $_.Task }) '「計」を含むだけのタスク（設計）は数える'
    Assert-Match $text 'PV 8 / EV 7\.5 / BAC 14 / SPI 0\.94'
}

Test-Case 'Measure-Progress: 基準日より後に終わったタスクは基準日には未完了とし、実績終了日を見込終了日に、進捗率を基準日までの営業日の割合にする' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '実績開始日', '実績終了日', '進捗率'),
            @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/21', '2026/10/12', '2026/10/21', '100%'),
            @('1.3', 'IF 設計', '佐藤', '2026/10/19', '2026/10/23', '2026/10/20', '2026/10/26', '100%'),
            @('1.4', '帳票設計', '鈴木', '2026/10/21', '2026/10/22', '', '2026/10/27', '100%'),
            @('2.1', '画面詳細', '山田', '2026/10/26', '2026/10/30', '2026/10/26', '2026/10/28', '100%')) })
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-23') }
    Assert-Equal @('1.2|1|0|完了|2026/10/21|2026/10/21', '1.3|0.8|1|遅延|2026/10/26|', '1.4|0.6|3|遅延|2026/10/27|', '2.1|0|-2|未着手|2026/10/28|') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}|{4:yyyy/MM/dd}|{5:yyyy/MM/dd}' -f $_.WbsNo, $_.Progress, $_.Delay, $_.Status, $_.ExpectedEnd, $_.DoneOn }) '進捗率は実績開始日（1.4 は予定開始日）から実績終了日までの営業日のうち、基準日までの割合'
    Assert-Match $text 'タスク 4 件（2026/10/23 時点。親のタスクは除く）: 完了 1・遅延 2・予定どおり 0・未着手 1'
    Assert-Match $text 'PV 15 / EV 13\.2 / BAC 20 / SPI 0\.88'
}

Test-Case 'Measure-Progress: 実績開始日が基準日より後のタスクは、基準日には進捗率 0 にする（記入された進捗率・完了の進捗率は後の状態のため）' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '実績開始日', '実績終了日', '進捗率', '予定工数'),
            @('4.1', '受注一覧の単体テスト', '山田', '2026/10/15', '2026/10/16', '2026/10/16', '', '30%', '2'),
            @('4.2', '受注登録の単体テスト', '佐藤', '2026/10/16', '2026/10/19', '2026/10/22', '', '30%', '2'),
            @('4.3', '受注詳細の単体テスト', '佐藤', '2026/10/19', '2026/10/20', '2026/10/21', '', '100%', '2')) })
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-16') }
    Assert-Equal @('4.1|0.3|1|遅延|False', '4.2|0|1|遅延|False', '4.3|0|0|未着手|False') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.WbsNo, $_.Progress, $_.Delay, $_.Status, $_.Done }) '10/22 に始めた 4.2 は 10/16 には 0%、10/21 に始めた 4.3 は 10/16 には完了していない'
    Assert-Match $text 'PV 3 / EV 0\.6 / BAC 6 / SPI 0\.20'
}

Test-Case 'Measure-Progress: WBS のフォルダに置いた結果（-OutXlsx）・週次報告・リカバリ計画のテンプレート（書き込んだものも）は WBS として読まず、-Exclude のファイルも読まない' {
    $dir = New-TestDir
    $f = New-Wbs $dir
    $null = & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -FillTemplate $recoveryTemplate -OutFile (Join-Path $dir 'リカバリ計画_20261014.xlsx') -OutXlsx (Join-Path $dir '進捗_20261014.xlsx') 6>$null
    $null = & (Join-Path $script:ToolsDir 'New-WeeklyReport.ps1') -Wbs $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -OutDir $dir 6>$null 3>$null
    Copy-Item -LiteralPath $recoveryTemplate -Destination $dir
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $dir -AsOf '2026-10-14' -Holidays $f.Holidays) }
    Assert-Equal @('1.1', '1.2', '1.3', '2.1') @($script:r | ForEach-Object { $_.WbsNo }) 'WBS.xlsx のタスクだけ（記入例の 3.2.1 も読まない）'
    Assert-Match $text 'タスク 4 件（2026/10/14 時点。親のタスクは除く）: 完了 1・遅延 2・予定どおり 0・未着手 1'
    foreach ($name in @('リカバリ計画_20261014.xlsx', '進捗_20261014.xlsx', '週次報告_20261014.xlsx', 'スケジュール遅延リカバリ計画_テンプレート.xlsx')) {
        Assert-Match $text ('読まなかったブック（進捗・週次報告の結果とリカバリ計画のテンプレート）: .*' + [regex]::Escape($name))
    }

    # ほかの WBS の写し（読まないもの）は -Exclude で除く
    Copy-Item -LiteralPath $f.Wbs -Destination (Join-Path $dir 'WBS_旧.xlsx')
    Assert-Equal 8 @(& $ProgressTool -Path $dir -AsOf '2026-10-14' 6>$null).Count '写しも読む'
    Assert-Equal 4 @(& $ProgressTool -Path $dir -AsOf '2026-10-14' -Exclude (Join-Path $dir 'WBS_旧.xlsx') 6>$null).Count
    $warnings = @()
    $null = & $ProgressTool -Path (Join-Path $dir '進捗_20261014.xlsx') -AsOf '2026-10-14' -WarningVariable +warnings -WarningAction SilentlyContinue 6>$null
    Assert-Match ($warnings -join ' ') 'WBS（タスク・予定開始・予定終了の見出しのある表）が見つかりませんでした。進捗・週次報告の結果とリカバリ計画のテンプレートは読みません: 進捗_20261014\.xlsx'
}

Test-Case 'Measure-Progress: 親のタスクはシートごとに見分ける（工程ごとのシートで番号を振り直した WBS）' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    $h = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数')
    New-SheetsBook $wbs ([ordered]@{
            '基本設計' = @($h, @('1', '要件確認', '山田', '2026/10/05', '2026/10/07', '100%', '3'), @('2', '画面設計', '山田', '2026/10/08', '2026/10/14', '50%', '5'))
            '詳細設計' = @($h, @('1', '画面詳細', '', '2026/10/15', '2026/10/23', '', '8'), @('1.1', '一覧画面', '佐藤', '2026/10/15', '2026/10/19', '0%', '4'), @('1.2', '登録画面', '佐藤', '2026/10/20', '2026/10/23', '0%', '4'))
        })
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' 6>$null)
    Assert-Equal @('基本設計|1|要件確認', '基本設計|2|画面設計', '詳細設計|1.1|一覧画面', '詳細設計|1.2|登録画面') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Place.Sheet, $_.WbsNo, $_.Task }) '基本設計の 1 は、詳細設計のシートに 1.1 があっても子のないタスク'
}

Test-Case 'Measure-Progress: 予定工数が 0 のタスク（マイルストーン）は PV・EV の重みを 0 にし、空のときだけ予定の営業日数を重みにする' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数'),
            @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '100%', '5'),
            @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '40%', '5'),
            @('1.3', '◆基本設計レビュー完了', '', '2026/10/09', '2026/10/09', '0%', '0'),
            @('1.4', '◆お客様承認', '', '2026/10/13', '2026/10/13', '0%', '0'),
            @('2.1', '画面詳細', '山田', '2026/10/19', '2026/10/23', '', '')) })
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14') }
    Assert-Equal @('5', '5', '0', '0', '5') @($script:r | ForEach-Object { [string]$_.BAC }) '2.1 は予定工数が空のため、予定の営業日数'
    Assert-Match $text 'PV 8 / EV 7 / BAC 15 / SPI 0\.88'
}

Test-Case 'Measure-Progress: 進捗率の「済み」「完了済」「完了済み」も完了とし、「未完了」は完了にしない' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗'),
            @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '済み'),
            @('1.2', 'テーブル設計', '山田', '2026/10/05', '2026/10/09', '完了済'),
            @('1.3', 'IF 設計', '山田', '2026/10/05', '2026/10/09', '完了済み'),
            @('1.4', '帳票設計', '山田', '2026/10/05', '2026/10/09', '済'),
            @('1.5', 'バッチ設計', '山田', '2026/10/05', '2026/10/09', '未完了')) })
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' 6>$null)
    Assert-Equal @('1.1|1|完了', '1.2|1|完了', '1.3|1|完了', '1.4|1|完了', '1.5|0|遅延') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.WbsNo, $_.Progress, $_.Status })
}

Test-Case 'Measure-Progress: 担当者別の集計は、全角・半角、英字の大文字・小文字と空白の違い（山田 太郎・山田　太郎・山田太郎、Nguyen・NGUYEN）をそろえて 1 人にまとめる' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数'),
            @('1.1', '画面設計', '山田 太郎', '2026/10/05', '2026/10/09', '100%', '5'),
            @('1.2', 'テーブル設計', '山田　太郎', '2026/10/05', '2026/10/09', '100%', '5'),
            @('1.3', 'IF 設計', '山田太郎', '2026/10/12', '2026/10/16', '50%', '5'),
            @('1.4', '帳票設計', 'Nguyen', '2026/10/05', '2026/10/09', '100%', '5'),
            @('1.5', 'バッチ設計', 'NGUYEN', '2026/10/12', '2026/10/16', '0%', '5')) })
    $out = Join-Path $dir '進捗.xlsx'
    $null = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' -OutXlsx $out 6>$null)
    $owners = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('Nguyen|2', '山田 太郎|3') @(2..$owners.Rows.Count | ForEach-Object { '{0}|{1}' -f $owners.Rows[$_]['A'], $owners.Rows[$_]['B'] }) '表示は最初に出てきた書き方'
}

function Add-GanttFill {
    #     ブックの 1 枚目のシートのセルを塗りつぶす（テスト用）。Fills: @(@(範囲, 色), ...) を順に塗る（後のものが上書き）。
    #     範囲は E3:I3 のようなセルの範囲か、G:H のような列全体（列のスタイル。Excel で列を選んで塗ったとき）。色は RRGGBB。
    #     1 つだけなら @(, @(範囲, 色)) で渡す
    param([string] $Path, $Fills)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $colors = @($Fills | ForEach-Object { $_[1] } | Select-Object -Unique)
    # 色ごとの塗りつぶしとスタイルを、Export-Workbook のスタイルの後ろに足す
    $fs = [System.IO.File]::OpenRead($Path)
    try { $zip = New-Object System.IO.Compression.ZipArchive($fs); $sr = New-Object System.IO.StreamReader($zip.GetEntry('xl/styles.xml').Open()); $styles = $sr.ReadToEnd(); $sr.Dispose(); $zip.Dispose() }
    finally { $fs.Dispose() }
    $fillBase = [int][regex]::Match($styles, '<fills count="(\d+)">').Groups[1].Value
    $xfBase = [int][regex]::Match($styles, '<cellXfs count="(\d+)">').Groups[1].Value
    Edit-ZipText $Path 'xl/styles.xml' {
        param($xml)
        $fill = (@($colors | ForEach-Object { '<fill><patternFill patternType="solid"><fgColor rgb="FF{0}"/><bgColor indexed="64"/></patternFill></fill>' -f $_ })) -join ''
        $xf = (@(for ($i = 0; $i -lt $colors.Count; $i++) { '<xf numFmtId="0" fontId="0" fillId="{0}" borderId="0" xfId="0" applyFill="1"/>' -f ($i + $fillBase) })) -join ''
        $xml.Replace(('<fills count="{0}">' -f $fillBase), ('<fills count="{0}">' -f ($fillBase + $colors.Count))).Replace('</fills>', $fill + '</fills>').Replace(('<cellXfs count="{0}">' -f $xfBase), ('<cellXfs count="{0}">' -f ($xfBase + $colors.Count))).Replace('</cellXfs>', $xf + '</cellXfs>')
    }.GetNewClosure()
    $pkg = Open-ZipPackage $Path
    try {
        $part = @((Get-WorkbookSheetParts $pkg).Sheets.Values)[0]
        $doc = Read-PartXmlForEdit $pkg $part
        foreach ($f in $Fills) {
            $style = [string]($xfBase + [array]::IndexOf($colors, $f[1]))
            $m = [regex]::Match($f[0], '^([A-Z]+)(\d*):([A-Z]+)(\d*)$')
            if (-not $m.Groups[2].Value) {
                foreach ($col in @(Get-XmlChildElements (@(Get-XmlChildElements $doc.DocumentElement 'cols')[0]) 'col')) {
                    $c = [int]$col.GetAttribute('min')
                    if ($c -ge (Get-ColumnIndex $m.Groups[1].Value) -and $c -le (Get-ColumnIndex $m.Groups[3].Value)) { $col.SetAttribute('style', $style) }
                }
                continue
            }
            for ($r = [int]$m.Groups[2].Value; $r -le [int]$m.Groups[4].Value; $r++) {
                $rowEl = Get-SheetRowElement $doc $r -Create
                for ($c = (Get-ColumnIndex $m.Groups[1].Value); $c -le (Get-ColumnIndex $m.Groups[3].Value); $c++) { (Get-SheetCellElement $doc $rowEl $r $c -Create).SetAttribute('s', $style) }
            }
        }
        $changes = @{ $part.ToLowerInvariant() = (ConvertTo-XmlString $doc) }
    }
    finally { $pkg.Zip.Dispose() }
    Save-ZipPackage -FilePath $Path -Changes $changes
}

function New-GanttBook {
    #     ガントチャートの WBS を作る（テスト用）。Rows の 1 行目は見出し（Export-Workbook の 1 行目）
    param([string] $Path, $Rows, $Fills)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $Path -Sheets @(@{ Name = 'WBS'; Columns = @($Rows[0] | ForEach-Object { @{ Header = $_; Width = 6 } }); Rows = @($Rows | Select-Object -Skip 1 | ForEach-Object { @{ Values = $_ } }) })
    Add-GanttFill $Path $Fills
}

Test-Case 'Measure-Progress: 日付の列が無いガントチャートの WBS は、塗りつぶしたセル（バー）の最初と最後の日付を予定開始日・予定終了日にする。土日・祝日の列・今日の線・行の色はバーにしない' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'WBS.xlsx'
    # 2 行目の見出しの E〜T 列が 10/1（木）〜10/16（金）の日付のセル。G・H・N・O 列（土日）と P 列（10/12 スポーツの日）は列全体が灰色、
    # R 列（基準日 10/14）は今日の線（赤。バーのあるセルは上書きしない）、5 行目は行全体が薄い灰色（その上にバー）。1.4 は白で塗って消したバーだけ
    $days = @(0..15 | ForEach-Object { ([datetime]'2026-10-01').AddDays($_) })
    $blank = @($days | ForEach-Object { '' })
    New-GanttBook $wbs @(
        (@('プロジェクト計画', '', '', '') + $blank),
        (@('WBS', 'タスク名', '担当', '進捗率') + $days),
        (@('1', '基本設計', '', '') + $blank),
        (@('1.1', '画面設計', '山田', '100%') + $blank),
        (@('1.2', 'テーブル設計', '佐藤', '60%') + $blank),
        (@('1.3', 'IF 設計', '佐藤', '0%') + $blank),
        (@('1.4', '帳票設計', '鈴木', '') + $blank)) @(
        @('G:H', 'D9D9D9'), @('N:P', 'D9D9D9'), @('A5:T5', 'F2F2F2'),
        @('E3:F3', '1F4E79'), @('I3:M3', '1F4E79'), @('Q3:T3', '1F4E79'),
        @('E4:K4', '5B9BD5'), @('I5:M5', '5B9BD5'), @('Q6:T6', '5B9BD5'),
        @('R4:R5', 'FF0000'), @('R7:R7', 'FF0000'), @('E7:F7', 'FFFFFF'))
    $holidays = Join-Path $dir 'syukujitsu.csv'
    [System.IO.File]::WriteAllText($holidays, "国民の祝日・休日月日,国民の祝日・休日名称`r`n2026/10/12,スポーツの日`r`n", [System.Text.Encoding]::GetEncoding(932))
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' -Holidays $holidays) }
    Assert-Equal @(
        '1.1|画面設計|山田|2026/10/01|2026/10/07|1|1|0|完了',
        '1.2|テーブル設計|佐藤|2026/10/05|2026/10/09|0.6|1|4|遅延',
        '1.3|IF 設計|佐藤|2026/10/13|2026/10/16|0|0.5|2|遅延') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3:yyyy/MM/dd}|{4:yyyy/MM/dd}|{5}|{6}|{7}|{8}' -f $_.WbsNo, $_.Task, $_.Owner, $_.PlanStart, $_.PlanEnd, $_.Progress, $_.PlannedProgress, $_.Delay, $_.Status }) '1.1 は土日をまたぐバー。親のタスク（1）は数えない'
    Assert-Match $text '予定開始日・予定終了日が読めないタスク 1 件' 'バーの無いタスク（1.4。土日・今日の線・白の色だけ）は、日付の読めないタスク'
    Assert-Match $text '読み取り（ガントチャート）: WBS\.xlsx「WBS」の E〜T 列（1 列 1 日）の塗りつぶしから、予定開始・予定終了を読みました'
    Assert-Match $text 'タスク 3 件（2026/10/14 時点。親のタスクは除く）: 完了 1・遅延 2・予定どおり 0・未着手 0'
    # 休日の CSV が無くても、列全体を塗った 10/12 はバーにしない
    $r2 = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' 6>$null 3>$null)
    Assert-Equal @('1.1|2026/10/01|2026/10/07', '1.2|2026/10/05|2026/10/09', '1.3|2026/10/13|2026/10/16') @($r2 | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd })
}

Test-Case 'Measure-Progress: 土日の列をバーと同じ色で塗ったガントチャートも、土日の列はバーにしない' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    # E〜M 列: 10/1(木)〜10/9(金)。G・H 列（土日）は列全体がバーと同じ青
    New-GanttBook $wbs @(
        @('WBS', 'タスク名', '担当', '進捗率', '10/1(木)', '10/2(金)', '10/3(土)', '10/4(日)', '10/5(月)', '10/6(火)', '10/7(水)', '10/8(木)', '10/9(金)'),
        @('1.1', '画面設計', '山田', '100%', '', '', '', '', '', '', '', '', ''),
        @('1.2', 'テーブル設計', '佐藤', '50%', '', '', '', '', '', '', '', '', ''),
        @('1.3', 'IF 設計', '佐藤', '0%', '', '', '', '', '', '', '', '', '')) @(
        @('G:H', '5B9BD5'), @('E2:F2', '5B9BD5'), @('I3:K3', '5B9BD5'), @('L4:M4', '5B9BD5'))
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-07' 6>$null)
    Assert-Equal @('1.1|2026/10/01|2026/10/02', '1.2|2026/10/05|2026/10/07', '1.3|2026/10/08|2026/10/09') @($r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd })
}

Test-Case 'Measure-Progress: ガントチャートの日の数字の見出しは、上の月の行（9月・10月。年は基準日に近い年）で日付にし、土日の無い列も読む' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    # E〜N 列: 9/28・29・30・10/1・2・5・6・7・8・9（土日の列が無い）
    New-GanttBook $wbs @(
        @('WBS', 'タスク名', '担当', '進捗率', '9月', '', '', '10月', '', '', '', '', '', ''),
        @('', '', '', '', 28, 29, 30, 1, 2, 5, 6, 7, 8, 9),
        @('1.1', '画面設計', '山田', '100%', '', '', '', '', '', '', '', '', '', ''),
        @('1.2', 'テーブル設計', '佐藤', '50%', '', '', '', '', '', '', '', '', '', ''),
        @('1.3', 'IF 設計', '佐藤', '', '', '', '', '', '', '', '', '', '', '')) @(
        @('E3:H3', '70AD47'), @('I4:L4', '70AD47'), @('M5:N5', '70AD47'))
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-07') }
    Assert-Equal @('1.1|2026/09/28|2026/10/01|完了', '1.2|2026/10/02|2026/10/07|遅延', '1.3|2026/10/08|2026/10/09|未着手') @($script:r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}|{3}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd, $_.Status }) '金曜日の列（10/2）のバーは金曜日まで'
    Assert-Match $text '「WBS」の E〜N 列（1 列 1 日）の塗りつぶしから'
}

Test-Case 'Measure-Progress: ガントチャートの「予定」「実績」の 2 行のタスクは、予定の行のバーを予定に、実績の行のバーを実績開始日（完了したタスクは実績終了日も）にする' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    # F〜O 列: 10/5(月)〜10/9(金)・10/12(月)〜10/16(金)（文字の見出し）。1.3 は 2 行ともタスク名がある（予・実）
    $head = @('10/5(月)', '10/6(火)', '10/7(水)', '10/8(木)', '10/9(金)', '10/12(月)', '10/13(火)', '10/14(水)', '10/15(木)', '10/16(金)')
    $blank = @($head | ForEach-Object { '' })
    New-GanttBook $wbs @(
        (@('WBS', 'タスク名', '担当', '進捗率', '区分') + $head),
        (@('1.1', '画面設計', '山田', '100%', '予定') + $blank),
        (@('', '', '', '', '実績') + $blank),
        (@('1.2', 'テーブル設計', '佐藤', '40%', '予定') + $blank),
        (@('', '', '', '', '実績') + $blank),
        (@('1.3', 'IF 設計', '鈴木', '0%', '予') + $blank),
        (@('1.3', 'IF 設計', '鈴木', '', '実') + $blank)) @(
        @('F2:H2', '5B9BD5'), @('F3:J3', 'FFC000'), @('H4:K4', '5B9BD5'), @('I5:L5', 'FFC000'), @('L6:O6', '5B9BD5'))
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14') }
    Assert-Equal @('1.1|2026/10/05|2026/10/07|2026/10/05|2026/10/09|完了|2', '1.2|2026/10/07|2026/10/12|2026/10/08||遅延|4', '1.3|2026/10/13|2026/10/16|2026/10/13||遅延|2') @($script:r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}|{3:yyyy/MM/dd}|{4:yyyy/MM/dd}|{5}|{6}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd, $_.ActualStart, $_.DoneOn, $_.Status, $_.Delay }) '実績の行はタスクとして数えない。完了していない 1.2 は実績終了日にしない'
    Assert-Match $text 'F〜O 列（1 列 1 日）の塗りつぶしから、予定開始・予定終了を読みました（「実績」の行から実績開始・実績終了も）'
}

Test-Case 'Measure-Progress: 1 列 1 週（月曜日の日付）のガントチャートは、バーの最初の週の最初の営業日から、最後の週の最後の営業日までにする' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'WBS.xlsx'
    New-GanttBook $wbs @(
        @('WBS', 'タスク名', '担当', '進捗率', '10/5', '10/12', '10/19', '10/26'),
        @('1.1', '要件定義', '田中', '100%', '', '', '', ''),
        @('1.2', '基本設計', '田中', '30%', '', '', '', ''),
        @('1.3', '詳細設計', '佐藤', '0%', '', '', '', '')) @(
        @('E2:E2', '5B9BD5'), @('F3:G3', '5B9BD5'), @('H4:H4', '5B9BD5'))
    $holidays = Join-Path $dir 'syukujitsu.csv'
    [System.IO.File]::WriteAllText($holidays, "国民の祝日・休日月日,国民の祝日・休日名称`r`n2026/10/12,スポーツの日`r`n", [System.Text.Encoding]::GetEncoding(932))
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-14' -Holidays $holidays) }
    Assert-Equal @('1.1|2026/10/05|2026/10/09', '1.2|2026/10/13|2026/10/23', '1.3|2026/10/26|2026/10/30') @($script:r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd }) '10/12 の週は、祝日の次の 10/13 から'
    Assert-Match $text 'E〜H 列（1 列 1 週）'
}

Test-Case 'Measure-Progress: 予定開始・予定終了の列があれば、ガントチャートのバーより列の日付を使う（列の無い項目と空のセルだけバーから読む）' {
    $dir = New-TestDir
    $head = @('10/5', '10/6', '10/7', '10/8', '10/9')
    $both = Join-Path $dir 'WBS_列あり.xlsx'
    New-GanttBook $both @(
        (@('WBS', 'タスク名', '担当', '予定開始日', '予定終了日', '進捗率') + $head),
        @('1.1', '画面設計', '山田', '2026/10/06', '2026/10/08', '50%', '', '', '', '', ''),
        @('1.2', 'テーブル設計', '佐藤', '2026/10/05', '2026/10/09', '0%', '', '', '', '', '')) @(
        @('G2:K2', '5B9BD5'), @('G3:H3', '5B9BD5'))
    $startOnly = Join-Path $dir 'WBS_開始日だけ.xlsx'
    New-GanttBook $startOnly @(
        (@('WBS', 'タスク名', '担当', '開始日', '進捗率') + $head),
        @('2.1', '画面詳細', '山田', '2026/10/06', '0%', '', '', '', '', ''),
        @('2.2', 'テーブル詳細', '佐藤', '', '0%', '', '', '', '', '')) @(
        @('F2:J2', '5B9BD5'), @('F3:H3', '5B9BD5'))
    $text = Get-HostText { $script:r = @(& $ProgressTool -Path $both, $startOnly -AsOf '2026-10-07') }
    Assert-Equal @('1.1|2026/10/06|2026/10/08', '1.2|2026/10/05|2026/10/09', '2.1|2026/10/06|2026/10/09', '2.2|2026/10/05|2026/10/07') @($script:r | ForEach-Object { '{0}|{1:yyyy/MM/dd}|{2:yyyy/MM/dd}' -f $_.WbsNo, $_.PlanStart, $_.PlanEnd }) '2.1 の開始日は列の日付。開始日が空の 2.2 は、開始日もバーから読む'
    Assert-NotMatch $text 'WBS_列あり\.xlsx「WBS」の .*塗りつぶし'
    Assert-Match $text '読み取り（ガントチャート）: WBS_開始日だけ\.xlsx「WBS」の F〜J 列（1 列 1 日）の塗りつぶしから、予定終了を読みました'
}

Test-Case 'Measure-Progress: 日付の見出しの下に塗りつぶしのバーが無い表（条件付き書式のバー、列全体を塗っただけのものなど）は読まず、警告で知らせる' {
    $wbs = Join-Path (New-TestDir) 'WBS.xlsx'
    New-GanttBook $wbs @(
        @('WBS', 'タスク名', '担当', '10/5', '10/6', '10/7'),
        @('1.1', '画面設計', '山田', '', '', ''),
        @('1.2', 'テーブル設計', '佐藤', '', '', ''),
        @('1.3', 'IF 設計', '佐藤', '', '', '')) @(, @('D:E', 'D9D9D9'))
    $warnings = @()
    $r = @(& $ProgressTool -Path $wbs -AsOf '2026-10-07' -WarningVariable +warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal 0 $r.Count
    Assert-Match ($warnings -join ' ') '見出しのある表）が見つかりませんでした。日付の見出しの下に、セルを塗りつぶしたバーがありません（条件付き書式で描いたバーは読めません）: WBS\.xlsx「WBS」'
}

Test-Case 'Measure-Progress: Get-Help でヘルプが表示される' {
    $h = Get-Help $ProgressTool
    Assert-Match $h.Synopsis 'WBS'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'Measure-Progress: -OutGantt で、WBS からガントチャート形式の進捗管理票（条件付き書式のバー・状態の数式・休日のシート）を作り、そのまま WBS として読める' {
    $dir = New-TestDir
    $w = New-Wbs $dir
    $out = Join-Path $dir '進捗管理票.xlsx'
    $text = Get-HostText { $script:before = @(& $ProgressTool -Path $w.Wbs -AsOf '2026-10-14' -Holidays $w.Holidays -OutGantt $out 3>$null) }
    Assert-Match $text 'ガントチャートの進捗管理票を保存しました: .*（1 日 1 列・2026/09/28〜2026/11/15）'
    $g = Read-GanttCells $out
    Assert-Equal @('ガントチャート', '休日', '説明') $g.Sheets
    # 説明のシート: 使い方・列の意味（進捗管理票にある列だけ）・バーの色の凡例（バーと同じ色のセル）
    $guide = Read-GanttCells $out 3
    Assert-Equal '区分|列・色|説明|記入' ((@('A', 'B', 'C', 'D') | ForEach-Object { $guide.Cells["${_}1"].Value }) -join '|')
    $guideRows = @($guide.Cells.Keys | Where-Object { $_ -match '^A\d+$' } | ForEach-Object { [int]($_ -replace 'A') } | Sort-Object)
    Assert-True (@($guideRows | Where-Object { $guide.Cells["A$_"].Value -eq '列' -and $guide.Cells["B$_"].Value -eq '見込終了' }).Count -eq 1) '見込終了の列の説明'
    Assert-True (@($guideRows | Where-Object { $guide.Cells["A$_"].Value -eq '列' -and $guide.Cells["B$_"].Value -eq '先行' }).Count -eq 0) '進捗管理票に無い列（先行）は説明しない'
    Assert-Equal @('20', '23', '21', '25', '22', '24', '26') @($guideRows | Where-Object { $guide.Cells["A$_"].Value -eq '色' -and $guide.Cells["B$_"].Style } | ForEach-Object { $guide.Cells["B$_"].Style } | Where-Object { [int]$_ -ge 20 }) '色の凡例のセル（実績・予定・遅延・見込みの延び・親・土日休日・基準日の線）'
    Assert-Equal @("休日='休日'!`$A`$2:`$A`$300") $g.Names
    # 左の列（見込終了は WBS に無くても置く。実績工数はある）と、日付の列（9/28 の月曜から、最後の予定 10/30 の 2 週あとの日曜まで）
    Assert-Equal @('WBS番号', 'タスク名', '担当者', '予定開始', '予定終了', '予定工数', '実績開始', '実績終了', '見込終了', '進捗率', '実績工数', '状態') @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L' | ForEach-Object { $g.Cells["${_}3"].Value })
    Assert-Equal '12|3|M4' $g.Frozen
    Assert-Equal @('進捗管理票', '基準日', 'TODAY()', '46309') @($g.Cells['A1'].Value, $g.Cells['K1'].Value, $g.Cells['L1'].Formula, $g.Cells['L1'].Value)
    Assert-Equal @('46293', '5', '月', '2026年10月') @($g.Cells['M3'].Value, $g.Cells['M3'].Style, $g.Cells['M2'].Value, $g.Cells['P1'].Value)
    Assert-True (-not $g.Cells.ContainsKey('M1')) '月の最初の列が次の月の 3 日前（9/28）なら、そこには月を書かない'
    # 10/12（スポーツの日）は「祝」。曜日は休日の名前の範囲から数式で決める
    Assert-Equal @('祝', '日') @($g.Cells['AA2'].Value, $g.Cells['AG2'].Value)
    Assert-Equal 'IF(COUNTIF(休日,AA$3)>0,"祝",MID("日月火水木金土",WEEKDAY(AA$3),1))' $g.Cells['AA2'].Formula
    Assert-Equal 49 @($g.Cells.Keys | Where-Object { $_ -match '^[A-Z]+3$' -and $g.Cells[$_].Style -eq '5' }).Count
    # 状態（基準日 10/14）: 完了・遅延（予定終了を過ぎた）・遅延（予定の 2 営業日に進捗 0%）・未着手。親のタスクは太字で、状態の数式なし
    Assert-Equal @('|11', '完了|2', '遅延|2', '遅延|2', '|11', '未着手|2') @(4..9 | ForEach-Object { '{0}|{1}' -f $g.Cells["L$_"].Value, $g.Cells["L$_"].Style })
    Assert-Equal ('IF(OR(N(J6)>=1,H6<>""),"完了",IF(OR(D6="",E6=""),"",IF(IF(AND(I6<>"",I6>=$L$1),NETWORKDAYS(E6+1,I6,休日)>0,IF($L$1>E6,OR(NETWORKDAYS(E6+1,$L$1,休日)>0,(1-N(J6))*MAX(1,NETWORKDAYS(D6,E6,休日))>0.5),' +
        'IF($L$1>=D6,(IF($L$1>=E6,1,NETWORKDAYS(D6,$L$1,休日)/MAX(1,NETWORKDAYS(D6,E6,休日)))-N(J6))*MAX(1,NETWORKDAYS(D6,E6,休日))>0.5,FALSE))),"遅延",IF($L$1<D6,"未着手","予定どおり"))))') $g.Cells['L6'].Formula
    Assert-Equal @('11', '12', '14') @($g.Cells['B4'].Style, $g.Cells['D4'].Style, $g.Cells['J4'].Style)
    Assert-Equal @('0.6', '10') @($g.Cells['J6'].Value, $g.Cells['J6'].Style)
    # WBS 番号は文字列の書式（あとから 2.10 と入力しても数の 2.1 にしない）。空の実績開始・実績終了・見込終了も日付の書式
    Assert-Equal @('19', '18', '4', '4', '12', '4') @($g.Cells['A4'].Style, $g.Cells['A5'].Style, $g.Cells['H6'].Style, $g.Cells['G7'].Style, $g.Cells['H8'].Style, $g.Cells['I7'].Style)
    # 条件付き書式: 実績・遅延・見込みの延び・親のタスク・予定・土日休日（タスクの行）、見出しの土日休日、今日の線、進捗率のデータバー
    $rules = $g.Rules
    Assert-Equal 9 $rules.Count
    Assert-Equal 'M4:BI9|0|AND(AND($G4<>"",M$3>=$G4,M$3<=IF($H4<>"",$H4,$L$1)),AND(M$2<>"土",M$2<>"日",M$2<>"祝"))' $rules[0]
    Assert-Equal 'M4:BI9|1|AND(AND(M$3>=$D4,M$3<=$E4),AND(M$2<>"土",M$2<>"日",M$2<>"祝"),$L4="遅延")' $rules[1]
    Assert-Equal 'M4:BI9|6|AND(AND($I4<>"",$E4<>"",N($J4)<1,$H4="",M$3>$E4,M$3<=$I4),AND(M$2<>"土",M$2<>"日",M$2<>"祝"))' $rules[2] '見込みの延び: 終わっていないタスクの予定終了の翌日から見込終了まで'
    Assert-Match $rules[3] '^M4:BI9\|2\|AND\(AND\(M\$3>=\$D4,M\$3<=\$E4\),.*LEFT\(SUBSTITUTE\(\$A5,"-","\."\),LEN\(SUBSTITUTE\(\$A4,"-","\."\)\)\+1\)=SUBSTITUTE\(\$A4,"-","\."\)&"\."\)\)$'
    Assert-Equal @('M4:BI9|3', 'M4:BI9|4', 'M2:BI3|4', 'M2:BI9|5') @($rules[4..7] | ForEach-Object { ($_ -split '\|')[0..1] -join '|' })
    Assert-Equal 'M2:BI9|5|M$3=$L$1' $rules[7]
    Assert-Match $rules[8] '^J4:J9\|'

    # 進捗管理票をそのまま WBS として読むと、元の WBS と同じ結果になる
    $after = @(& $ProgressTool -Path $out -AsOf '2026-10-14' -Holidays $w.Holidays 6>$null 3>$null)
    $key = { param($t) '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $t.WbsNo, $t.Task, $t.Owner, $t.Progress, $t.PlannedProgress, $t.Delay, $t.Status }
    Assert-Equal @($script:before | ForEach-Object { & $key $_ }) @($after | ForEach-Object { & $key $_ })
    Assert-Equal @('完了', '遅延', '遅延', '未着手') @($after | ForEach-Object { $_.Status })

    # 記入した実績を消さないよう、すでにあるファイルには書かない
    $failed = $false
    try { $null = & $ProgressTool -Path $w.Wbs -AsOf '2026-10-14' -OutGantt $out 6>$null 3>$null } catch { $failed = $true; Assert-Match $_.Exception.Message '上書きしません' }
    Assert-True $failed '上書きしない'
    # 元の WBS と進捗管理票を同じフォルダに置くと、同じタスクが 2 つあると知らせる
    $warnings = (Get-HostText { $null = & $ProgressTool -Path $dir -AsOf '2026-10-14' 6>$null }) 
    Assert-Match $warnings '同じタスク（WBS 番号とタスク名）が、いくつものファイルにあります（6 件。'
}

Test-Case 'Measure-Progress: -OutGantt の日付の列は、期間が長ければ（-GanttUnit Week でも）1 週 1 列にする' {
    $dir = New-TestDir
    $w = New-Wbs $dir
    $out = Join-Path $dir 'week.xlsx'
    $null = & $ProgressTool -Path $w.Wbs -AsOf '2026-10-14' -OutGantt $out -GanttUnit Week 6>$null 3>$null
    $g = Read-GanttCells $out
    # 9/28 から 10/30 の 4 週あと（11/27）の週まで、月曜の日付（m/d）。曜日の行は無い
    $heads = @($g.Cells.Keys | Where-Object { $_ -match '^[A-Z]+3$' -and $g.Cells[$_].Style -eq '6' })
    Assert-Equal 9 $heads.Count
    Assert-Equal @('46293', '46300') @($g.Cells['M3'].Value, $g.Cells['N3'].Value)
    Assert-True (-not $g.Cells.ContainsKey('M2')) '週ごとは曜日の行を書かない'
    Assert-Equal 'M4:U9|0|AND(AND($G4<>"",$G4<=M$3+6,IF($H4<>"",$H4,$L$1)>=M$3),TRUE)' $g.Rules[0]
    Assert-Equal 'M4:U9|6|AND(AND($I4<>"",$E4<>"",N($J4)<1,$H4="",M$3+6>$E4,M$3<=$I4),TRUE)' $g.Rules[2] '週ごとの見込みの延びは、予定終了より後にかかる週'
    Assert-Equal 'M2:U9|5|AND($L$1>=M$3,$L$1<=M$3+6)' @($g.Rules | Where-Object { $_ -like '*|5|*' })[0]

    # 期間が 92 日より長い WBS は、Auto でも 1 週 1 列
    $long = Join-Path $dir 'long.csv'
    @('タスク,予定開始,予定終了', 'A,2026/10/01,2026/10/30', 'B,2026/11/02,2027/01/29') | Set-Content -LiteralPath $long -Encoding UTF8
    $text = Get-HostText { $null = & $ProgressTool -Path $long -AsOf '2026-10-14' -OutGantt (Join-Path $dir 'long.xlsx') 3>$null }
    Assert-Match $text '（1 週 1 列・2026/09/28〜'
}

Test-Case 'Measure-Progress: -OutGantt の状態（数式の計算済みの値）は、このツールの状態と同じ（休日だけのタスク・週末をはさむ遅れ・見込終了・ちょうど半分の遅れ）' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'wbs.csv'
    @('WBS番号,タスク名,担当者,予定開始,予定終了,実績開始,実績終了,見込終了,進捗率',
        '1,親,,2026/12/01,2026/12/31,,,,',
        '1.1,完了,A,2026/12/01,2026/12/03,2026/12/01,2026/12/03,,100%',
        '1.2,休日だけ,A,2026/12/07,2026/12/07,,,,0%',
        '1.3,金曜終わり90%,A,2026/12/07,2026/12/11,2026/12/07,,,90%',
        '1.4,金曜終わり50%,A,2026/12/07,2026/12/11,2026/12/07,,,50%',
        '1.5,途中ちょうど,A,2026/12/08,2026/12/15,2026/12/08,,,25%',
        '1.6,途中遅れ,A,2026/12/08,2026/12/15,2026/12/08,,,10%',
        '1.7,見込で遅れ,A,2026/12/14,2026/12/18,,,2026/12/21,0%',
        '1.8,見込で間に合う,A,2026/12/09,2026/12/18,2026/12/09,,2026/12/17,0%',
        '1.9,未着手,A,2026/12/16,2026/12/18,,,,',
        '1.10,未着手だが見込遅れ,A,2026/12/16,2026/12/18,,,2026/12/22,',
        '1.11,半分,A,2026/12/10,2026/12/11,2026/12/10,,,50%') | Set-Content -LiteralPath $wbs -Encoding UTF8
    $holidays = Join-Path $dir 'hol.csv'
    @('2026/12/07,休日') | Set-Content -LiteralPath $holidays -Encoding UTF8
    foreach ($asOf in @('2026-12-10', '2026-12-12', '2026-12-14')) {
        $out = Join-Path $dir "g-$asOf.xlsx"
        $tasks = @(& $ProgressTool -Path $wbs -AsOf $asOf -Holidays $holidays -OutGantt $out 6>$null 3>$null)
        $g = Read-GanttCells $out
        # 見込終了の列（I）があるので、状態は K。親のタスク（4 行目）は空
        Assert-Equal '見込終了' $g.Cells['I3'].Value
        Assert-Match $g.Cells['K6'].Formula '^IF\(OR\(N\(J6\)>=1,H6<>""\),"完了",IF\(OR\(D6="",E6=""\),"",IF\(IF\(AND\(I6<>"",I6>=\$K\$1\),NETWORKDAYS\(E6\+1,I6,休日\)>0,'
        Assert-Equal @($tasks | ForEach-Object { '{0}|{1}' -f $_.WbsNo, $_.Status }) @(5..15 | ForEach-Object { '{0}|{1}' -f $g.Cells["A$_"].Value, $g.Cells["K$_"].Value }) "基準日 $asOf"
    }
}
