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
    Assert-Equal @('タスク', '担当者別', '週次') $sheet.Sheets
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
    # 色ごとの塗りつぶしとスタイル（Export-Workbook のスタイルの後ろ、5 番から）を足す
    Edit-ZipText $Path 'xl/styles.xml' {
        param($xml)
        $fill = (@($colors | ForEach-Object { '<fill><patternFill patternType="solid"><fgColor rgb="FF{0}"/><bgColor indexed="64"/></patternFill></fill>' -f $_ })) -join ''
        $xf = (@(for ($i = 0; $i -lt $colors.Count; $i++) { '<xf numFmtId="0" fontId="0" fillId="{0}" borderId="0" xfId="0" applyFill="1"/>' -f ($i + 3) })) -join ''
        $xml.Replace('<fills count="3">', ('<fills count="{0}">' -f (3 + $colors.Count))).Replace('</fills>', $fill + '</fills>').Replace('<cellXfs count="5">', ('<cellXfs count="{0}">' -f (5 + $colors.Count))).Replace('</cellXfs>', $xf + '</cellXfs>')
    }.GetNewClosure()
    $pkg = Open-ZipPackage $Path
    try {
        $part = @((Get-WorkbookSheetParts $pkg).Sheets.Values)[0]
        $doc = Read-PartXmlForEdit $pkg $part
        foreach ($f in $Fills) {
            $style = [string](5 + [array]::IndexOf($colors, $f[1]))
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
