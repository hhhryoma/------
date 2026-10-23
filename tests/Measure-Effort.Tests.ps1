#     Measure-Effort.ps1 のテスト。WBS と工数実績の表はテストの中で作る（Export-Workbook と CSV）

$EffortTool = Join-Path $script:ToolsDir 'Measure-Effort.ps1'

function New-EffortProject {
    #     予定工数（人日）のある WBS と、日付・氏名・WBS・作業内容・時間の工数実績（日付が空の行は上の行の日付）
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '実績終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(
        @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '2026/10/09', '100%', '4'),
        @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '', '50%', '5'),
        @('1.3', 'IF 設計', '佐藤', '2026/10/19', '2026/10/23', '', '20%', '4'),
        @('2.1', '画面詳細', '山田', '2026/10/26', '2026/10/30', '', '0%', '3'))
    $wbs = Join-Path $Dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })

    $logCols = @('日付', '氏名', 'WBS', '作業内容', '時間') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $logRows = @(
        @('2026/10/05', '山田', '1.1', '画面設計', '8'),
        @('2026/10/06', '山田', '1.1', '', '8'),
        @('2026/10/07', '山田', '1.1', '', '7:30'),
        @('2026/10/08', '山田', '1.1', '', '8'),
        @('2026/10/09', '山田', '1.1', '', '8'),
        @('2026/10/12', '佐藤', '1.2', 'テーブル設計', '8'),
        @('2026/10/13', '佐藤', '1.2', '', '8'),
        @('2026/10/14', '佐藤', '1.2', '', '8'),
        @('2026/10/15', '佐藤', '1.2', '', '8'),
        @('2026/10/16', '佐藤', '1.2', '', '8'),
        @('2026/10/19', '佐藤', '1.2', '', '4'),
        @('', '', '', 'IF 設計', '4'),
        @('', '', '', '定例会議', '1'),
        @('2026/10/20', '山田', '', '定例会議', '1:00'),
        @('2026/10/26', '山田', '2.1', '', '8'))
    $log = Join-Path $Dir '工数実績/工数実績.xlsx'
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $log) | Out-Null
    Export-Workbook -FilePath $log -Sheets @(@{ Name = '10月'; Columns = $logCols; Rows = @($logRows | ForEach-Object { @{ Values = $_ } }) })
    return [pscustomobject]@{ Wbs = $wbs; LogDir = (Split-Path -Parent $log) }
}

Test-Case 'Measure-Effort: 工数実績をタスクごとに集計し、WBS の予定工数と比べて予定超過・完了・超過を知らせ、担当者×週とWBS に無い作業も出す' {
    $dir = New-TestDir
    $p = New-EffortProject $dir
    $out = Join-Path $dir '工数集計.xlsx'
    $text = Get-HostText { $script:r = @(& $EffortTool $p.LogDir -Wbs $p.Wbs -AsOf '2026-10-23' -OutXlsx $out) }
    Assert-Equal @(
        '1.1|画面設計|4|4.94|4.94|完了・超過',
        '1.2|テーブル設計|5|5.5|11|予定超過',
        '1.3|IF 設計|4|0.5|2.5|',
        '2.1|画面詳細|3|0|3|') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.WbsNo, $_.Task, $_.Planned, $_.Actual, $_.Eac, $_.Judgement }) '日付・氏名が空の行は上の行のもの。タスク名でも対応付ける。7:30 は 7.5 時間'
    Assert-Match $text '工数の実績: 89\.5 時間（11\.2 人日。2026/10/05～2026/10/20・2 人・14 行）'
    Assert-Match $text '基準日（2026/10/23）より後の日付の行 1 件は数えていません'
    Assert-Match $text 'WBS のタスクに対応した工数: 87\.5 時間（98%） / WBS に無い作業: 定例会議 0\.25（人日）'
    Assert-Match $text '予定工数のあるタスク 4 件: 予定 16 / 実績 10\.9（人日） / EV 7\.3 / CPI 0\.67'
    Assert-Match $text '  - \[予定超過\] 1\.2 テーブル設計（佐藤）: 予定 5 / 実績 5\.5 / 進捗 50% → 完了見込 11'

    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('タスク', '担当者×週', 'WBS に無い作業', '読み取り結果') $sheet.Sheets
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'WBS'!B3" }).Count -eq 1) 'WBS の該当行へのリンク'
    $weeks = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal '担当|2026/10/05|2026/10/12|2026/10/19|合計' ((@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $weeks.Rows[1][$_] }) -join '|')
    $sato = @($weeks.Rows.Values | Where-Object { $_['A'] -eq '佐藤' })[0]
    Assert-Equal '0|5|1.13|6.13' ((@('B', 'C', 'D', 'E') | ForEach-Object { $sato[$_] }) -join '|')
}

Test-Case 'Measure-Effort: 担当の列が無い人ごとの工数表はファイル名を担当にし、1時間30分・0:45・2h・人日の見出しを読む（WBS なし）' {
    $dir = New-TestDir
    @('日付,作業,工数', '2026/10/01,設計,1時間30分', '2026/10/01,レビュー,0:45', '2026/10/02,設計,2h', '2026/10/02,設計,たくさん') | Set-Content -LiteralPath (Join-Path $dir '田中.csv') -Encoding UTF8
    @('日付,作業,工数(人日)', '2026/10/02,設計,0.5') | Set-Content -LiteralPath (Join-Path $dir '鈴木.csv') -Encoding UTF8
    $warnings = @()
    $r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間 -WarningVariable +warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @('設計|7.5|田中、鈴木', 'レビュー|0.75|田中') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner })
    Assert-Match ($warnings -join ' ') '数として読めない行 1 件'
    Assert-Throws { & $EffortTool $dir -HoursPerDay 0 6>$null } '0 より大きい数'
}

Test-Case 'Measure-Effort: Get-Help でヘルプが表示される' {
    $h = Get-Help $EffortTool
    Assert-Match $h.Synopsis '工数'
    Assert-Equal 11 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
