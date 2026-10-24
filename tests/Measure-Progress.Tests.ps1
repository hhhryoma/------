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
        '1.2|テーブル設計|佐藤|0.6|1|2|遅延',
        '1.3|IF 設計|佐藤|0|0.5|2|遅延',
        '2.1|画面詳細|山田|0|0|0|未着手') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $_.WbsNo, $_.Task, $_.Owner, $_.Progress, $_.PlannedProgress, $_.Delay, $_.Status }) '親のタスク（1・2）は数えない。10/12 は祝日'
    Assert-Equal '2026/10/14|2026/10/20' ('{0:yyyy/MM/dd}|{1:yyyy/MM/dd}' -f $r[1].ExpectedEnd, $r[2].ExpectedEnd) '見込の終了日（予定終了日 + 遅れの営業日）'
    $text = Get-HostText { & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays | Out-Null }
    Assert-Match $text 'タスク 4 件（2026/10/14 時点。親のタスクは除く）: 完了 1・遅延 2・予定どおり 0・未着手 1'
    Assert-Match $text 'PV 12 / EV 8 / BAC 22 / SPI 0\.67 / AC 11 / CPI 0\.73'
    Assert-Match $text '遅延 2 件（最大 2 営業日: テーブル設計）'
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
    Assert-Equal '1|1.2|テーブル設計|佐藤|2026/10/05|2026/10/09|2026/10/05|2026/10/14' ((@('A8', 'B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8') | ForEach-Object { $cells["02_遅延状況!$_"] }) -join '|')
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

Test-Case 'Measure-Progress: 和暦の設定（ja-JP の JapaneseCalendar）でも、日付は西暦で書く' {
    $dir = New-TestDir
    $f = New-Wbs $dir
    $out = Join-Path $dir '進捗.xlsx'
    $text = Invoke-InWareki { Get-HostText { & $ProgressTool -Path $f.Wbs -AsOf '2026-10-14' -Holidays $f.Holidays -OutXlsx $out | Out-Null } }
    Assert-Match $text 'タスク 4 件（2026/10/14 時点'
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-True (@($book.Rows.Values | ForEach-Object { $_.Values } | Where-Object { "$_" -match '^0?8/10/' }).Count -eq 0) '和暦の年（08/10/14）にしない'
    Assert-True (@($book.Rows.Values | ForEach-Object { $_.Values } | Where-Object { "$_" -match '^2026/10/' }).Count -gt 0) 'タスクのシートの予定の日付'
}

Test-Case 'Measure-Progress: Get-Help でヘルプが表示される' {
    $h = Get-Help $ProgressTool
    Assert-Match $h.Synopsis 'WBS'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
