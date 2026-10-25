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

Test-Case 'Measure-Effort: 実績の合計と CPI は、タスクごとに丸める前の実績で出す（週次報告の実績・CPI と同じ値）' {
    $dir = New-TestDir
    $wbs = Join-Path $dir 'WBS.xlsx'
    New-SheetsBook $wbs ([ordered]@{ 'WBS' = @(@('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数'),
            @('1.1', '画面設計', '山田', '2026/10/19', '2026/10/23', '10%', '1'),
            @('1.2', 'テーブル設計', '山田', '2026/10/19', '2026/10/23', '10%', '1'),
            @('1.3', 'IF 設計', '山田', '2026/10/19', '2026/10/23', '10%', '1')) })
    $logDir = Join-Path $dir '工数'
    New-Item -ItemType Directory -Path $logDir | Out-Null
    @('日付,氏名,WBS,時間', '2026/10/19,山田,1.1,1', '2026/10/19,山田,1.2,1', '2026/10/19,山田,1.3,1') | Set-Content -LiteralPath (Join-Path $logDir '工数実績.csv') -Encoding UTF8
    $text = Get-HostText { $script:r = @(& $EffortTool $logDir -Wbs $wbs -AsOf '2026-10-23') }
    Assert-Equal @('0.13', '0.13', '0.13') @($script:r | ForEach-Object { [string]$_.Actual }) 'タスクごとの実績の表示は 0.01 に丸める（1 時間 = 0.125 人日）'
    Assert-Match $text '予定工数のあるタスク 3 件: 予定 3 / 実績 0\.4（人日） / EV 0\.3 / CPI 0\.80' 'EV 0.3 ÷ 実績 0.375（丸めた 0.39 で割ると 0.77）'
    $weekly = Get-HostText { & (Join-Path $script:ToolsDir 'New-WeeklyReport.ps1') -Wbs $wbs -EffortLog $logDir -AsOf '2026-10-23' | Out-Null }
    Assert-Match $weekly '・予定工数のあるタスク 3 件: 予定 3 / 実績 0\.38（人日）・CPI 0\.80'
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

Test-Case 'Measure-Effort: 時刻（h:mm）・経過時間（[h]:mm）の書式のセルに入れた工数（7:30・31:15）を時間で読む' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $log = Join-Path $dir '工数実績.xlsx'
    $logCols = @('日付', '氏名', '作業内容', '時間') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $logRows = @(@('2026/10/05', '佐藤', '設計', 0.3125), @('2026/10/06', '佐藤', '設計', 0.3125), @('2026/10/07', '佐藤', '設計', (31.25 / 24)))
    Export-Workbook -FilePath $log -Sheets @(@{ Name = '10月'; Columns = $logCols; Rows = @($logRows | ForEach-Object { @{ Values = $_ } }) })
    # Excel で 7:30 と入れたセル: D2 は時刻の h:mm（組み込みの 20）、D3・D4 は経過時間の [h]:mm
    Edit-ZipText $log 'xl/styles.xml' {
        param($x)
        $x = $x.Replace('</numFmts>', '<numFmt numFmtId="177" formatCode="[h]:mm"/></numFmts>')
        $x.Replace('<cellXfs count="5">', '<cellXfs count="7">').Replace('</cellXfs>', '<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="177" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>')
    }
    Edit-ZipText $log 'xl/worksheets/sheet1.xml' { param($x) $x.Replace('<c r="D2" s="2">', '<c r="D2" s="5">').Replace('<c r="D3" s="2">', '<c r="D3" s="6">').Replace('<c r="D4" s="2">', '<c r="D4" s="6">') }
    $warnings = @()
    $r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間 -WarningVariable +warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @('設計|46.25|佐藤') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner }) '7:30 + 7:30 + 31:15'
    Assert-Equal 0 $warnings.Count ($warnings -join ' ')
}

Test-Case 'Measure-Effort: 年の無い日付（12/28・1/5(火)）は基準日の前後半年の年とみなし、年明けの集計で年末の行を基準日より後にしない' {
    $dir = New-TestDir
    @('日付,氏名,作業内容,時間', '12/28,佐藤,設計,8', '12/29(火),佐藤,設計,8', '1/5（火）,佐藤,設計,8', '1/6,佐藤,設計,8') | Set-Content -LiteralPath (Join-Path $dir '工数.csv') -Encoding UTF8
    $text = Get-HostText { $null = & $EffortTool $dir -AsOf '2027-01-08' -Unit 時間 }
    Assert-Match $text '工数の実績: 32 時間（4 人日。2026/12/28～2027/01/06・1 人・4 行）'
    Assert-NotMatch $text '基準日（2027/01/08）より後の日付の行'
}

Test-Case 'Measure-Effort: 合計・小計の行（日付・担当・WBS・作業の欄が「小計」「Total」「合 計」など）は数えない（「設計」のように計を含むだけの作業は数える）' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $logCols = @('日付', '氏名', 'WBS', '作業内容', '時間') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $logRows = @(
        @('2026/10/19', '佐藤', '1.1', '画面設計', [double]8),
        @('2026/10/20', '佐藤', '1.1', '画面設計', [double]8),
        @('小計', '', '', '', [double]16),
        @('2026/10/21', '鈴木', '', '設計', [double]4),
        @('', 'Total', '', '', [double]4),
        @('', '', 'Grand Total', '', [double]20),
        @('', '', '', '合 計', [double]20))
    Export-Workbook -FilePath (Join-Path $dir '工数実績.xlsx') -Sheets @(@{ Name = '10月'; Columns = $logCols; Rows = @($logRows | ForEach-Object { @{ Values = $_ } }) })
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間) }
    Assert-Equal @('1.1 画面設計|16|佐藤', '設計|4|鈴木') @($script:r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner })
    Assert-Match $text '工数の実績: 20 時間（2\.5 人日。2026/10/19～2026/10/21・2 人・3 行）'
}

Test-Case 'Measure-Effort: 工数実績の WBS 番号は、書き方（ＷＢＳ１．１・No.1.3・1.1 画面設計・1-1）が WBS と違っても番号の部分で対応付ける' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(@('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '50%', '4'), @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '0%', '4'), @('1.3', 'IF設計', '佐藤', '2026/10/12', '2026/10/16', '0%', '4'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    @('日付,氏名,WBS,時間', '2026/10/05,山田,1-1,8', '2026/10/06,山田,ＷＢＳ１．１,8', '2026/10/07,山田,1.1 画面設計,8', '2026/10/12,佐藤,1．2,8', '2026/10/13,佐藤,No.1.3,8') | Set-Content -LiteralPath (Join-Path $dir '工数/工数.csv') -Encoding UTF8
    $text = Get-HostText { $script:r = @(& $EffortTool (Join-Path $dir '工数') -Wbs $wbs -AsOf '2026-10-23' -Unit 時間) }
    Assert-Equal @('1.1|24', '1.2|8', '1.3|8') @($script:r | ForEach-Object { '{0}|{1}' -f $_.WbsNo, $_.Actual })
    Assert-Match $text 'WBS のタスクに対応した工数: 40 時間（100%）'
    Assert-NotMatch $text 'WBS に無い作業'
}

Test-Case 'Measure-Effort: シートごとに番号を振り直した WBS（同じ WBS 番号のタスクが複数）は、作業の名前でタスクを選び、決まらない工数は対応付けずに警告する' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $basic = @(@('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '50%', '16'), @('1.2', '帳票設計', '山田', '2026/10/12', '2026/10/16', '0%', '16'))
    $detail = @(@('1.1', '画面詳細', '佐藤', '2026/10/19', '2026/10/23', '0%', '16'), @('1.2', '帳票詳細', '佐藤', '2026/10/26', '2026/10/30', '0%', '16'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = '基本設計'; Columns = $wbsCols; Rows = @($basic | ForEach-Object { @{ Values = $_ } }) }, @{ Name = '詳細設計'; Columns = $wbsCols; Rows = @($detail | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    @('日付,氏名,WBS,作業内容,時間', '2026/10/05,山田,1.1,画面設計,8', '2026/10/19,佐藤,1.1,画面詳細,4', '2026/10/12,山田,1.2,,2', '2026/10/13,山田,1.2,打合せ,1') | Set-Content -LiteralPath (Join-Path $dir '工数/工数.csv') -Encoding UTF8
    $text = Get-HostText { $script:r = @(& $EffortTool (Join-Path $dir '工数') -Wbs $wbs -AsOf '2026-10-23' -Unit 時間) }
    $actual = @{}
    foreach ($t in $script:r) { $actual[$t.Task] = $t.Actual }
    Assert-Equal '8|0|4|0' ('{0}|{1}|{2}|{3}' -f $actual['画面設計'], $actual['帳票設計'], $actual['画面詳細'], $actual['帳票詳細']) '同じ番号のタスクは作業の名前で選ぶ。決まらなければ、最後のシートのタスクにしない'
    Assert-Match $text '同じ WBS 番号のタスクが複数あり、タスク名でも決まらない工数 2 行は、WBS のタスクに対応付けていません（WBS 番号 1\.2）'
}

Test-Case 'Measure-Effort: 担当者の名前は、全角・半角、英字の大文字・小文字と空白の違い（山田 太郎・山田太郎・山田　太郎、Nguyen・NGUYEN）をそろえて同じ人とみなす' {
    $dir = New-TestDir
    @('日付,氏名,作業内容,時間', '2026/10/19,山田 太郎,設計,8', '2026/10/20,山田太郎,設計,8', '2026/10/21,山田　太郎,設計,8', '2026/10/21,Nguyen,設計,4', '2026/10/22,NGUYEN,設計,4') | Set-Content -LiteralPath (Join-Path $dir '工数.csv') -Encoding UTF8
    $out = Join-Path $dir '工数集計.xlsx'
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間 -OutXlsx $out) }
    Assert-Match $text '工数の実績: 32 時間（4 人日。2026/10/19～2026/10/22・2 人・5 行）'
    Assert-Equal @('設計|山田 太郎、Nguyen') @($script:r | ForEach-Object { '{0}|{1}' -f $_.Task, $_.Owner }) '表示は最初に出てきた書き方'
    $weeks = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('Nguyen|8', '山田 太郎|24') @(2..$weeks.Rows.Count | ForEach-Object { '{0}|{1}' -f $weeks.Rows[$_]['A'], $weeks.Rows[$_]['C'] }) '担当者×週は 1 人 1 行'
}

Test-Case 'Measure-Effort: 日付を列に並べた工数表（日付のセルの見出し・担当の列）を読み、WBS 番号で WBS のタスクに対応付ける（合計の行と列、休・- のセルは数えない）' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(@('1.2', 'テーブル設計', '佐藤', '2026/10/19', '2026/10/23', '50%', '24'), @('1.3', 'IF 設計', '鈴木', '2026/10/19', '2026/10/23', '20%', '16'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })
    # 1 行目は表題、2 行目が見出し（日付は日付のセル。表示形式は m/d(aaa)）。10/26 は基準日より後
    $day = { param($d) [datetime]('2026-10-' + $d) }
    $rows = @(
        @('WBS番号', 'タスク名', '担当者', (& $day 19), (& $day 20), (& $day 21), (& $day 22), (& $day 23), (& $day 26), '合計'),
        @('1.2', 'テーブル設計', '佐藤', [double]4, [double]4, '', '休', [double]2, [double]8, [double]18),
        @('1.3', 'IF 設計', '鈴木', '', [double]3.5, [double]4, '-', '1:30', '', [double]9),
        @('', '定例会議', '山田', [double]1, '', '', '', '', '', [double]1),
        @('合計', '', '', [double]5, [double]7.5, [double]4, [double]0, [double]3.5, [double]8, [double]28))
    $log = Join-Path $dir '工数/工数表.xlsx'
    New-Item -ItemType Directory -Path (Split-Path -Parent $log) | Out-Null
    $title = @('2026年10月 工数実績', '', '', '', '', '', '', '', '', '') | ForEach-Object { @{ Header = $_; Width = 10 } }
    Export-Workbook -FilePath $log -Sheets @(@{ Name = '10月'; Columns = $title; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    Edit-ZipText $log 'xl/styles.xml' { param($x) $x.Replace('formatCode="yyyy/mm/dd"', 'formatCode="m/d\(aaa\)"') }
    $out = Join-Path $dir '工数集計.xlsx'
    $warnings = @()
    $text = Get-HostText { $script:r = @(& $EffortTool (Join-Path $dir '工数') -Wbs $wbs -AsOf '2026-10-23' -Unit 時間 -OutXlsx $out -WarningVariable +warnings -WarningAction SilentlyContinue) }
    Assert-Equal @('1.2|テーブル設計|24|10', '1.3|IF 設計|16|9') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.WbsNo, $_.Task, $_.Planned, $_.Actual }) '4 + 4 + 2・3.5 + 4 + 1:30'
    Assert-Match $text '読み取り（工数実績）: 1 ファイル・3 行'
    Assert-Match $text '工数の実績: 20 時間（2\.5 人日。2026/10/19～2026/10/23・3 人・7 行）'
    Assert-Match $text '基準日（2026/10/23）より後の日付の行 1 件は数えていません'
    Assert-Match $text 'WBS のタスクに対応した工数: 19 時間（95%） / WBS に無い作業: 定例会議 1（時間）'
    Assert-Equal 0 $warnings.Count ('休・- のセルは数として読めない行にしない: ' + ($warnings -join ' '))
    $others = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal '定例会議|1|1|山田' ((@('A', 'B', 'C', 'D') | ForEach-Object { $others.Rows[2][$_] }) -join '|')
    Assert-True (@($others.Links | Where-Object { $_.Location -eq "'10月'!D5" }).Count -eq 1) '場所は工数のセル'
    $readLog = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal '工数表.xlsx|10月|2|3|日付（2026/10/19～2026/10/26 の列）、担当（担当者）、WBS番号、タスク（タスク名）、工数（日付の列のセル）' ((@('B', 'C', 'D', 'E', 'F') | ForEach-Object { $readLog.Rows[2][$_] }) -join '|')
}

Test-Case 'Measure-Effort: 日付を列に並べた工数表の日だけの見出し（1(木)・2↵金・5）は、表の上の「2026年10月」の月とし、担当の列が無ければファイル名を担当にする' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir '山田.xlsx') ([ordered]@{
            'Sheet1' = @(
                @('2026年10月 工数実績（山田）', '', '', '', '', '', '', '', ''),
                @('WBS', '作業内容', '1(木)', "2`n金", '5', '6', '7', '8', '合計'),
                @('1.1', '画面設計', [double]8, [double]7.5, '休', '', '', '', [double]15.5),
                @('1.2', 'テーブル設計', '', '0:30', [double]8, [double]8, '-', [double]4, [double]20.5),
                @('合計', '', [double]8, [double]8, [double]8, [double]8, '', [double]4, [double]36))
        })
    $warnings = @()
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2026-10-07' -Unit 時間 -WarningVariable +warnings -WarningAction SilentlyContinue) }
    Assert-Equal @('1.2 テーブル設計|16.5|山田', '1.1 画面設計|15.5|山田') @($script:r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner })
    Assert-Match $text '工数の実績: 32 時間（4 人日。2026/10/01～2026/10/06・1 人・5 行）'
    Assert-Match $text '基準日（2026/10/07）より後の日付の行 1 件は数えていません'
    Assert-Equal 0 $warnings.Count ($warnings -join ' ')
}

Test-Case 'Measure-Effort: 日付を列に並べた工数表の日だけの見出しは、シート名（12月）か見出しの行の「2027年1月」の月とし、年が無ければ基準日の前後半年の年にする（日付と曜日の 2 行の見出しも読む。月の名前のシートは担当にしない。月が分からなければ警告する）' {
    $dir = New-TestDir
    # 12月: 日付の行の下に、タスクの見出しと曜日の行。1月: タスクの見出しと「2027年1月」の行の下に日付の行
    New-SheetsBook (Join-Path $dir '佐藤.xlsx') ([ordered]@{
            '12月' = @(
                @('', '24', '25', '28', '29', '30'),
                @('タスク', '木', '金', '月', '火', '水'),
                @('設計', [double]8, [double]8, '', '', ''),
                @('レビュー', '', '', [double]2, '', ''))
            '1月'  = @(
                @('タスク', '2027年1月', '', '', '', ''),
                @('', [double]4, [double]5, [double]6, [double]7, [double]8),
                @('設計', '', [double]8, [double]4, [double]8, ''))
            'メモ' = @(
                @('タスク', '1', '2', '3', '4', '5'),
                @('設計', [double]1, '', '', '', ''))
        })
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2027-01-06' -Unit 時間) }
    Assert-Equal @('設計|28|佐藤', 'レビュー|2|佐藤') @($script:r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner }) 'シート名の 12月・1月 は担当にしない'
    Assert-Match $text '工数の実績: 30 時間（3\.8 人日。2026/12/24～2027/01/06・1 人・5 行）'
    Assert-Match $text '基準日（2027/01/06）より後の日付の行 1 件は数えていません'
    Assert-Match $text '佐藤\.xlsx メモ: 日付の列が日だけで月が分からないため、工数を読んでいません'
}

Test-Case 'Measure-Effort: 日付を列に並べた人日の工数表（表題に「（人日）」）は時間にし、予定と実績を 2 行に書く表は実績の行だけをタスク名で WBS に対応付ける' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(@('1.1', '画面設計', '山田', '2026/10/19', '2026/10/23', '50%', '3'), @('1.2', 'テーブル設計', '佐藤', '2026/10/19', '2026/10/23', '50%', '2'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    New-SheetsBook (Join-Path $dir '工数/チーム工数.xlsx') ([ordered]@{
            '工数' = @(
                @('工数実績（人日）', '', '', '', '', '', '', '', ''),
                @('タスク名', '担当', '区分', '10/19(月)', '10/20(火)', '10/21(水)', '10/22(木)', '10/23(金)', '合計'),
                @('画面設計', '山田', '予定', [double]1, [double]1, [double]1, '', '', [double]3),
                @('', '', '実績', [double]1, [double]0.5, [double]1, [double]0.25, '', [double]2.75),
                @('テーブル設計', '佐藤', '予定', '', '', [double]1, [double]1, [double]1, [double]3),
                @('テーブル設計', '佐藤', '実績', '', '', [double]0.5, [double]1, '', [double]1.5),
                @('合計', '', '実績', [double]1, [double]0.5, [double]1.5, [double]1.25, '', [double]4.25))
        })
    $text = Get-HostText { $script:r = @(& $EffortTool (Join-Path $dir '工数') -Wbs $wbs -AsOf '2026-10-23') }
    Assert-Equal @('1.1|画面設計|山田|3|2.75', '1.2|テーブル設計|佐藤|2|1.5') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.WbsNo, $_.Task, $_.Owner, $_.Planned, $_.Actual })
    Assert-Match $text '工数の実績: 34 時間（4\.3 人日。2026/10/19～2026/10/22・2 人・6 行）'
    Assert-Match $text 'WBS のタスクに対応した工数: 34 時間（100%）'
}

Test-Case 'Measure-Effort: 日付を列に並べた工数表の CSV（WBS・氏名・10/1 … の列）も読む' {
    $dir = New-TestDir
    @('WBS,氏名,10/1,10/2,10/5,10/6,10/7,合計', '1.1,田中,4,4,,2,,10', '1.2,田中,,1:30,,,,1.5') | Set-Content -LiteralPath (Join-Path $dir '工数.csv') -Encoding UTF8
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間) }
    Assert-Equal @('1.1|10|田中', '1.2|1.5|田中') @($script:r | ForEach-Object { '{0}|{1}|{2}' -f $_.Task, $_.Actual, $_.Owner })
    Assert-Match $text '工数の実績: 11\.5 時間（1\.4 人日。2026/10/01～2026/10/06・1 人・4 行）'
}

Test-Case 'Measure-Effort: 1 行が 1 件の工数実績の表のあるブックは、日付を列に並べた集計のシートを読まない。日付を並べたガントチャートの WBS（予定開始日などの列・■ だけのセル）からは工数を読まない' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir '日報.xlsx') ([ordered]@{
            '日報' = @(
                @('日付', '氏名', 'WBS', '作業内容', '時間'),
                @('2026/10/19', '佐藤', '1.2', '', [double]8),
                @('2026/10/20', '佐藤', '1.2', '', [double]6))
            '集計' = @(
                @('WBS', '作業内容', '10/19', '10/20', '10/21', '10/22', '10/23'),
                @('1.2', '', [double]8, [double]6, '', '', ''))
        })
    New-SheetsBook (Join-Path $dir 'WBS_ガント.xlsx') ([ordered]@{
            'WBS' = @(
                @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '10/19', '10/20', '10/21', '10/22', '10/23'),
                @('1.2', 'テーブル設計', '佐藤', '2026/10/19', '2026/10/21', [double]1, [double]1, [double]1, '', ''),
                @('1.3', 'IF 設計', '佐藤', '2026/10/22', '2026/10/23', '', '', '', [double]1, [double]1))
        })
    New-SheetsBook (Join-Path $dir '予定表.xlsx') ([ordered]@{
            '予定' = @(
                @('WBS', 'タスク名', '10/19', '10/20', '10/21', '10/22', '10/23'),
                @('1.2', 'テーブル設計', '■', '■', '■', '', ''),
                @('1.3', 'IF 設計', '', '', '', '■', '■'))
        })
    $warnings = @()
    $text = Get-HostText { $script:r = @(& $EffortTool $dir -AsOf '2026-10-23' -Unit 時間 -WarningVariable +warnings -WarningAction SilentlyContinue) }
    Assert-Equal @('1.2|14') @($script:r | ForEach-Object { '{0}|{1}' -f $_.Task, $_.Actual }) '集計のシートと WBS の 1 は数えない'
    Assert-Match $text '工数の実績: 14 時間（1\.8 人日。2026/10/19～2026/10/20・1 人・2 行）'
    Assert-Match $text '  表が見つからないファイル: WBS_ガント\.xlsx'
    Assert-Equal 0 $warnings.Count ('■ のセルは数として読めない行にしない: ' + ($warnings -join ' '))
}

Test-Case 'Measure-Effort: WBS に 1.1 と 1.10 があるとき、数のセル（1.10 が 1.1 になる）と CSV の 1.1 はタスク名、次に担当で選び、決まらなければ警告する（文字のセルはそのまま。担当者と違う人の工数を知らせる）' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS番号', 'タスク名', '担当者', '予定開始', '予定終了', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(@('1.1', '画面設計', '山田', '2026/10/05', '2026/10/09', '50%', '16'), @('1.2', 'テーブル設計', '佐藤', '2026/10/05', '2026/10/09', '0%', '8'), @('1.10', '帳票設計', '佐藤', '2026/10/12', '2026/10/16', '0%', '16'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '日報') | Out-Null
    # 標準の書式のセルに 1.10 と入れると数の 1.1 になる（[double]1.1）。文字の書式のセルは '1.1'・'1.10' のまま
    New-SheetsBook (Join-Path $dir '日報/日報.xlsx') ([ordered]@{
            '日報' = @(
                @('日付', '氏名', 'WBS番号', '作業内容', '時間'),
                @('2026/10/05', '佐藤', [double]1.1, '', [double]2),
                @('2026/10/05', '山田', [double]1.1, '', [double]3),
                @('2026/10/06', '鈴木', [double]1.1, '', [double]1),
                @('2026/10/06', '鈴木', [double]1.1, '帳票設計', [double]1.5),
                @('2026/10/07', '鈴木', '1.1', '', [double]4),
                @('2026/10/07', '山田', '1.10', '', [double]0.5))
        })
    @('日付,氏名,WBS,時間', '2026/10/08,佐藤,1.1,1') | Set-Content -LiteralPath (Join-Path $dir '日報/日報.csv') -Encoding UTF8
    # 日付を列に並べた佐藤の工数表（担当はファイル名）の数の 1.1
    New-SheetsBook (Join-Path $dir '日報/佐藤.xlsx') ([ordered]@{ '2026年10月' = @(@('WBS番号', '5(月)', '6(火)', '7(水)', '8(木)', '9(金)'), @([double]1.1, '', [double]2, '', '', '')) })
    $text = Get-HostText { $script:r = @(& $EffortTool (Join-Path $dir '日報') -Wbs $wbs -AsOf '2026-10-23' -Unit 時間) }
    Assert-Equal @('1.1|画面設計|7', '1.2|テーブル設計|0', '1.10|帳票設計|7') @($script:r | ForEach-Object { '{0}|{1}|{2}' -f $_.WbsNo, $_.Task, $_.Actual }) '1.1: 山田の数の 1.1 3 + 文字の 1.1 4。1.10: 佐藤の数の 1.1 2 + 帳票設計 1.5 + 文字の 1.10 0.5 + CSV の佐藤 1 + 佐藤の工数表 2'
    Assert-Match $text '1\.1 か 1\.10 か分からず、タスク名・担当でも決まらない工数 1 行は、WBS のタスクに対応付けていません（WBS 番号 1\.1/1\.10）'
    Assert-Match $text 'WBS のタスクに対応した工数: 14 時間（93%） / WBS に無い作業: 1\.1 1（時間）'
    Assert-Match $text '担当者と違う人がつけた工数 1 件: 1\.10 帳票設計（佐藤）に 山田 0\.5（時間）。手伝いでなければ、WBS 番号を確かめてください'
}

Test-Case 'Measure-Effort: 担当の列が無い人ごと・月ごとのファイル（工数実績_佐藤_202610.xlsx・日報_山田_2026年10月.csv）と月ごとのシートは、年月と工数実績などのことばを除いたファイル名（WBS にあれば WBS の書き方）を担当にし、読み取り結果に書く' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS番号', 'タスク名', '担当者', '予定開始', '予定終了', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(@('1.1', '画面設計', '山田', '2026/10/26', '2026/11/06', '50%', '16'), @('1.2', 'テーブル設計', '佐藤 花子', '2026/10/26', '2026/11/06', '50%', '16'))
    $wbs = Join-Path $dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })
    $log = Join-Path $dir '工数実績'
    New-Item -ItemType Directory -Path $log | Out-Null
    New-SheetsBook (Join-Path $log '工数実績_佐藤花子_202610.xlsx') ([ordered]@{
            '2026年10月' = @(
                @('WBS番号', 'タスク名', '26(月)', '27(火)', '28(水)', '29(木)', '30(金)'),
                @('1.2', 'テーブル設計', [double]4, [double]4, '', [double]2, ''))
        })
    New-SheetsBook (Join-Path $log '工数実績_佐藤花子_202611.xlsx') ([ordered]@{
            '2026年11月' = @(
                @('WBS番号', 'タスク名', '2(月)', '4(水)', '5(木)', '6(金)', '9(月)'),
                @('1.2', 'テーブル設計', [double]3, '', [double]1, '', ''))
        })
    @('日付,WBS,時間', '2026/10/26,1.1,8', '2026/11/02,1.1,6') | Set-Content -LiteralPath (Join-Path $log '日報_山田_2026年10月.csv') -Encoding UTF8
    # 1 行が 1 件の表を、月ごとのシートに分けた鈴木のブック（シート名は担当にしない）
    New-SheetsBook (Join-Path $log '作業日報（鈴木）.xlsx') ([ordered]@{ '10月' = @(@('日付', 'WBS', '時間'), @('2026/10/30', '1.1', [double]1)); '11月' = @(@('日付', 'WBS', '時間'), @('2026/11/04', '1.1', [double]2)) })
    $out = Join-Path $dir '工数集計.xlsx'
    $text = Get-HostText { $script:r = @(& $EffortTool $log -Wbs $wbs -AsOf '2026-11-06' -Unit 時間 -OutXlsx $out) }
    Assert-Match $text '工数の実績: 31 時間（3\.9 人日。2026/10/26～2026/11/05・3 人・9 行）'
    Assert-True (-not $text.Contains('担当者と違う人')) ('WBS に無い鈴木は知らせない: ' + $text)
    $weeks = Read-XlsxSheet -Path $out -Index 2
    $byOwner = @{}
    foreach ($n in @($weeks.Rows.Keys | Where-Object { $_ -gt 1 })) { $row = $weeks.Rows[$n]; $byOwner[$row['A']] = (@('B', 'C', 'D') | ForEach-Object { $row[$_] }) -join '|' }
    Assert-Equal @('佐藤 花子', '山田', '鈴木') @($byOwner.Keys | Sort-Object { [int][char]$_[0] }) '月ごとのファイル・シートを 1 人にまとめ、WBS にあれば WBS の書き方にする'
    Assert-Equal @('10|4|14', '8|6|14', '1|2|3') @($byOwner['佐藤 花子'], $byOwner['山田'], $byOwner['鈴木'])
    $readLog = Read-XlsxSheet -Path $out -Index 4
    $found = @{}
    foreach ($n in @($readLog.Rows.Keys | Where-Object { $_ -gt 1 })) { $row = $readLog.Rows[$n]; $found['{0}|{1}' -f $row['B'], $row['C']] = '{0}|{1}' -f $row['F'], $row['G'] }
    Assert-Equal '日付（2026/10/26～2026/10/30 の列）、担当（ファイル名から「佐藤 花子」）、WBS番号、タスク（タスク名）、工数（日付の列のセル）|' $found['工数実績_佐藤花子_202610.xlsx|2026年10月']
    Assert-Equal '日付（2026/11/02～2026/11/09 の列）、担当（ファイル名から「佐藤 花子」）、WBS番号、タスク（タスク名）、工数（日付の列のセル）|' $found['工数実績_佐藤花子_202611.xlsx|2026年11月']
    Assert-Equal '日付、担当（ファイル名から「山田」）、WBS番号（WBS）、工数（時間）|タスク' $found['日報_山田_2026年10月.csv|']
    Assert-Equal '日付、担当（ファイル名から「鈴木」）、WBS番号（WBS）、工数（時間）|タスク' $found['作業日報（鈴木）.xlsx|11月']
}

Test-Case 'Measure-Effort: Get-Help でヘルプが表示される' {
    $h = Get-Help $EffortTool
    Assert-Match $h.Synopsis '工数'
    Assert-Equal 11 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
