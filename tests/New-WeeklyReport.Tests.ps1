#     New-WeeklyReport.ps1 のテスト。WBS・指摘管理表・障害管理表はテストの中で作る（Export-Workbook と CSV）

$WeeklyReport = Join-Path $script:ToolsDir 'New-WeeklyReport.ps1'

function New-WeeklyProject {
    #     2026/10/23（金）の報告に使う WBS（xlsx）・指摘管理表（xlsx）・障害管理表（CSV）
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $wbsCols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '実績開始日', '実績終了日', '進捗率') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $wbsRows = @(
        @('1', '基本設計', '', '2026/10/05', '2026/10/23', '', '', ''),
        @('1.1', '画面設計', '山田', '2026/10/05', '2026/10/16', '2026/10/05', '2026/10/16', '100%'),
        @('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/21', '2026/10/12', '2026/10/22', '100%'),
        @('1.3', 'IF 設計', '佐藤', '2026/10/19', '2026/10/23', '2026/10/19', '', '40%'),
        @('2.1', '画面詳細', '山田', '2026/10/26', '2026/10/30', '', '', ''),
        @('2.2', 'バッチ詳細', '鈴木', '2026/10/28', '2026/11/06', '', '', ''),
        @('2.3', '帳票詳細', '', '2026/10/19', '2026/10/28', '2026/10/19', '', '60%'))
    $wbs = Join-Path $Dir 'WBS.xlsx'
    Export-Workbook -FilePath $wbs -Sheets @(@{ Name = 'WBS'; Columns = $wbsCols; Rows = @($wbsRows | ForEach-Object { @{ Values = $_ } }) })

    $findingCols = @('No', '設計書', '指摘内容', '重要度', '対応状況', '指摘日', '対応期限', '完了日', '対応者') | ForEach-Object { @{ Header = $_; Width = 12 } }
    $findingRows = @(
        @('1', 'DD03 テーブル定義書', '型が誤り', '高', '未対応', '2026/10/19', '2026/10/21', '', '佐藤'),
        @('2', 'DD03 テーブル定義書', '桁数が違う', '中', '対応中', '2026/10/20', '2026/10/27', '', '佐藤'),
        @('3', 'DD02 画面設計書', '文言の誤り', '低', '完了', '2026/10/14', '2026/10/20', '2026/10/22', '山田'),
        @('4', 'DD02 画面設計書', '並び順の記載なし', '中', '未対応', '2026/10/22', '', '', '山田'),
        @('5', 'DD02 画面設計書', '重複した指摘', '中', '取下げ', '2026/10/22', '', '', '山田'))
    $findings = Join-Path $Dir 'レビュー/指摘管理表.xlsx'
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $findings) | Out-Null
    Export-Workbook -FilePath $findings -Sheets @(@{ Name = '指摘'; Columns = $findingCols; Rows = @($findingRows | ForEach-Object { @{ Values = $_ } }) })

    # 担当者の書き方の違い（全角の空白）は同じ人とみなす
    $bugs = Join-Path $Dir 'テスト/障害管理表.csv'
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $bugs) | Out-Null
    @('障害番号,機能,障害内容,重要度,状態,起票日,完了日,担当者,修正予定日',
        'B-001,受注一覧,並び順が逆,中,修正中,2026/10/20,,佐藤　,2026/10/22',
        'B-002,受注登録,登録できない,高,完了,2026/10/16,2026/10/21,山田,2026/10/21',
        'B-003,受注登録,合計が違う,高,未対応,2026/10/23,,鈴木,2026/10/29',
        'B-004,受注登録,税の端数,低,仕様どおり,2026/10/21,,鈴木,') | Set-Content -LiteralPath $bugs -Encoding UTF8
    return [pscustomobject]@{ Wbs = $wbs; Findings = $findings; Bugs = $bugs }
}

Test-Case 'New-WeeklyReport: WBS・指摘・障害から、今週の実績・遅延・来週の予定と担当者ごとのやることをまとめ、Excel と文章に保存する' {
    $dir = New-TestDir
    $p = New-WeeklyProject $dir
    $out = Join-Path $dir '週次報告'
    $history = Join-Path $out '実績.csv'
    $text = Get-HostText { $script:items = @(& $WeeklyReport -Wbs $p.Wbs -Findings $p.Findings -BugList $p.Bugs -AsOf '2026-10-23' -OutDir $out -HistoryCsv $history) }

    # 担当者ごとのやること（遅延・期限切れ → 来週に終わる・期限が来週まで → 来週に始まる → そのほかの未完了）
    $byOwner = { param($o) @($script:items | Where-Object { $_.Owner -eq $o } | ForEach-Object { '{0}|{1}|{2}|{3:MM/dd}' -f $_.State, $_.Kind, $_.Title, $_.Date }) }
    Assert-Equal @('期限切れ|指摘|DD03 テーブル定義書: 型が誤り|10/21', '期限切れ|障害|B-001: 並び順が逆|10/22', '遅延|タスク|1.3 IF 設計|10/23', '期限間近|指摘|DD03 テーブル定義書: 桁数が違う|10/27') (& $byOwner '佐藤') '全角の空白のある「佐藤　」も同じ人'
    Assert-Equal @('来週終了|タスク|2.1 画面詳細|10/30', '未完了|指摘|DD02 画面設計書: 並び順の記載なし|') (& $byOwner '山田') '完了・取下げは載せない'
    Assert-Equal @('期限間近|障害|B-003: 合計が違う|10/29', '来週開始|タスク|2.2 バッチ詳細|10/28') (& $byOwner '鈴木') '仕様どおりの障害は数えない'
    Assert-Equal @('来週終了|タスク|2.3 帳票詳細|10/28') (& $byOwner '（担当なし）')
    Assert-Equal '（担当なし）' $script:items[-1].Owner '担当なしは最後'
    Assert-Equal '3 営業日の遅れ・進捗 40%・見込 10/28' (@($script:items | Where-Object { $_.State -eq '遅延' })[0].Note)

    # 報告の文章
    Assert-Match $text '週次報告（2026/10/17～2026/10/23）'
    Assert-Match $text '・タスク 6 件: 完了 2・遅延 1・予定どおり 1・未着手 2 / SPI '
    Assert-Match $text '・今週の完了: 1 件 / 今週に終わる予定: 2 件（うち未完了 1 件）\n  - 1\.2 テーブル設計（佐藤）\n'
    Assert-Match $text '・遅延: 1 件\n  - 1\.3 IF 設計（佐藤）: 3 営業日の遅れ・進捗 40%・予定終了 10/23・見込 10/28\n'
    Assert-Match $text '・来週（10/24～10/30）の予定: 3 件\n  - 開始 10/28 2\.2 バッチ詳細（鈴木）\n  - 終了 10/28 2\.3 帳票詳細\n  - 開始・終了 10/30 2\.1 画面詳細（山田）\n'
    Assert-Match $text '■ レビューの指摘\n・今週: 起票 3 件・完了 1 件\n・未完了: 3 件・うち期限切れ 1 件 / 重要度 高 1・中 2\n  - 期限 10/21 DD03 テーブル定義書: 型が誤り（佐藤）\n'
    Assert-Match $text '■ 障害\n・今週: 起票 2 件・完了 1 件\n・未完了: 2 件・うち期限切れ 1 件 / 重要度 高 1・中 1\n  - 期限 10/22 B-001: 並び順が逆（佐藤）\n'
    Assert-Match $text '【佐藤】遅延 1・期限切れ 2・期限間近 1\n  - \[期限切れ\] 指摘: DD03 テーブル定義書: 型が誤り（期限 10/21・高・未対応）\n'
    Assert-Match $text '  - \[遅延\] タスク: 1\.3 IF 設計（予定終了 10/23・3 営業日の遅れ・進捗 40%・見込 10/28）'
    Assert-Match $text '【山田】来週終了 1・ほか未完了 1\n  - \[来週終了\] タスク: 2\.1 画面詳細（予定終了 10/30・進捗 0%）\n'
    Assert-Match $text '【（担当なし）】来週終了 1'

    # 保存したもの
    $txt = Join-Path $out '週次報告_20261023.txt'
    Assert-True (Test-Path -LiteralPath $txt) '文章'
    Assert-Match ([System.IO.File]::ReadAllText($txt)) '^週次報告（2026/10/17～2026/10/23）\r\n'
    $xlsx = Join-Path $out '週次報告_20261023.xlsx'
    $summary = Read-XlsxSheet -Path $xlsx -Index 1
    Assert-Equal @('概要', '担当者', '担当者別', 'タスク', '指摘・障害', '推移') $summary.Sheets
    $rows = @{}
    foreach ($r in $summary.Rows.Keys) { $rows[('{0}|{1}' -f $summary.Rows[$r]['A'], $summary.Rows[$r]['B'])] = $summary.Rows[$r]['C'] }
    Assert-Equal '1' $rows['進捗|遅延']
    Assert-Equal '3' $rows['進捗|来週の予定']
    Assert-Equal '3' $rows['指摘|今週の起票']
    Assert-Equal '1' $rows['障害|期限切れ']
    $people = Read-XlsxSheet -Path $xlsx -Index 2
    $sato = @($people.Rows.Values | Where-Object { $_['A'] -eq '佐藤' })[0]
    Assert-Equal '1|0|2|1|1|1' ((@('B', 'C', 'D', 'E', 'F', 'G') | ForEach-Object { $sato[$_] }) -join '|') '遅延・来週終了・指摘の未完了・うち期限切れ・障害の未完了・うち期限切れ'
    $list = Read-XlsxSheet -Path $xlsx -Index 3
    Assert-Equal 10 $list.Rows.Count '見出しと 9 件'
    Assert-True (@($list.Links | Where-Object { $_.Location -eq "'指摘'!C2" }).Count -eq 1) '指摘管理表の該当セルへのリンク'
    Assert-True (@($list.Links | Where-Object { $_.Location -eq "'WBS'!B5" }).Count -eq 1) 'WBS の該当セルへのリンク'
    $issues = Read-XlsxSheet -Path $xlsx -Index 5
    Assert-Equal 8 $issues.Rows.Count '見出しと、未完了 5 件・今週の完了 2 件'
    # 担当者ごとの文章（本人に送るリマインド）
    $ownerDir = Join-Path $out '週次報告_20261023_担当者別'
    $names = [string[]]@(Get-ChildItem -LiteralPath $ownerDir -File | ForEach-Object { $_.Name })
    [Array]::Sort($names, [StringComparer]::Ordinal)
    Assert-Equal @('佐藤.txt', '山田.txt', '担当なし.txt', '鈴木.txt') $names
    $sato = [System.IO.File]::ReadAllText((Join-Path $ownerDir '佐藤.txt'))
    Assert-Match $sato '^佐藤 さん\r\n\r\n週次報告（2026/10/17～2026/10/23）の、担当の分です。\r\n\r\n■ 遅延・期限切れ（3 件）\r\n  - \[期限切れ\] 指摘: DD03 テーブル定義書: 型が誤り（期限 10/21・高・未対応） … 指摘管理表\.xlsx C2\r\n'
    Assert-Match $sato '■ 来週まで（1 件）\r\n  - \[期限間近\] 指摘: DD03 テーブル定義書: 桁数が違う'
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $ownerDir '担当なし.txt'))) '^週次報告（2026/10/17～2026/10/23）の、担当者が決まっていないものの分です。'
    Assert-Equal @('時点,区分,項目,値', '2026/10/23,進捗,タスク,6') @(Get-Content -LiteralPath $history -Encoding UTF8 | Select-Object -First 2 | ForEach-Object { $_ -replace '"', '' })
}

Test-Case 'New-WeeklyReport: 実績の CSV で前回と比べ、推移とグラフを足す。同じ日に作り直すと、その日の数字を置き換える' {
    $dir = New-TestDir
    $p = New-WeeklyProject $dir
    $out = Join-Path $dir 'out'
    $history = Join-Path $dir '実績.csv'
    $null = & $WeeklyReport -Wbs $p.Wbs -Findings $p.Findings -BugList $p.Bugs -AsOf '2026-10-16' -OutDir $out -HistoryCsv $history 6>$null
    $null = & $WeeklyReport -Wbs $p.Wbs -Findings $p.Findings -BugList $p.Bugs -AsOf '2026-10-23' -OutDir $out -HistoryCsv $history 6>$null
    $text = Get-HostText { $null = & $WeeklyReport -Wbs $p.Wbs -Findings $p.Findings -BugList $p.Bugs -AsOf '2026-10-23' -Since '2026-10-19' -OutDir $out -HistoryCsv $history }
    Assert-Match $text '週次報告（2026/10/19～2026/10/23）'
    Assert-Match $text '・遅延: 1 件（前回 1）' '10/16 には、10/22 に終わった 1.2 は未完了で 1 営業日の遅れ'
    Assert-Match $text '・未完了: 3 件（前回 1）・うち期限切れ 1 件（前回 0）'
    Assert-Match $text '前回 2026/10/16 と比べました'
    $dates = @(Get-Content -LiteralPath $history -Encoding UTF8 | Select-Object -Skip 1 | ForEach-Object { ($_ -replace '"', '').Split(',')[0] } | Sort-Object -Unique)
    Assert-Equal @('2026/10/16', '2026/10/23') $dates '同じ日の行は置き換える'
    $xlsx = Join-Path $out '週次報告_20261023.xlsx'
    $summary = Read-XlsxSheet -Path $xlsx -Index 1
    Assert-Equal '前回（2026/10/16）' $summary.Rows[1]['D']
    $delay = @($summary.Rows.Values | Where-Object { $_['A'] -eq '進捗' -and $_['B'] -eq '遅延' })[0]
    Assert-Equal '1|1|0' ('{0}|{1}|{2}' -f $delay['C'], $delay['D'], $delay['E']) '今回・前回・増減'
    $trend = Read-XlsxSheet -Path $xlsx -Index 6
    Assert-Equal 3 $trend.Rows.Count '見出しと 2 回'
    Assert-Equal '2026/10/16' $trend.Rows[2]['A']
    Assert-Equal @('遅延・未完了・期限切れの推移', 'SPI・CPI の推移') @(Read-XlsxCharts $xlsx | ForEach-Object { $_.Title })
}

Test-Case 'New-WeeklyReport: 案件のフォルダの「週次報告の設定.csv」を読む。コマンドの指定が優先し、結果のフォルダは読まない' {
    $dir = New-TestDir
    $p = New-WeeklyProject $dir
    @('項目,値', 'WBS,WBS.xlsx', '指摘管理表,レビュー', '障害管理表,テスト\障害管理表.csv', 'タイトル,受注システム 週報') | Set-Content -LiteralPath (Join-Path $dir '週次報告の設定.csv') -Encoding UTF8
    $null = & $WeeklyReport $dir -AsOf '2026-10-23' 6>$null
    $first = Join-Path $dir '週次報告/受注システム 週報_20261023.xlsx'
    Assert-True (Test-Path -LiteralPath $first) '出力先の既定は設定ファイルのフォルダの「週次報告」'
    # 前の週の報告（指摘・障害のシートがある）を指摘管理表のフォルダに置いても、出力先なら読まない
    $out = Join-Path $dir 'レビュー/報告'
    $null = & $WeeklyReport $dir -AsOf '2026-10-16' -OutDir $out 6>$null
    $text = Get-HostText { $null = & $WeeklyReport $dir -AsOf '2026-10-23' -OutDir $out }
    Assert-Match $text '・今週: 起票 3 件・完了 1 件'
    Assert-Match $text '受注システム 週報（2026/10/17～2026/10/23）'

    Assert-Throws { & $WeeklyReport -AsOf '2026-10-23' 6>$null } '-Wbs・-Findings・-BugList・-IssueList・-EffortLog のどれか'
    Assert-Throws { & $WeeklyReport -Wbs $p.Wbs -AsOf '2026-10-23' -Since '2026-10-24' 6>$null } 'より後です'
    @('項目,値', '設計書,設計') | Set-Content -LiteralPath (Join-Path $dir '週次報告の設定.csv') -Encoding UTF8
    Assert-Throws { & $WeeklyReport $dir 6>$null } '週次報告の設定に知らない項目があります: 設計書'
    Assert-Throws { & $WeeklyReport (Join-Path $dir 'テスト') 6>$null } '週次報告の設定ファイルが見つかりません'
}

Test-Case 'New-WeeklyReport: 課題・QA の管理表の未完了・期限切れを報告し、担当者ごとのやることに入れる' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('課題No', '区分', '件名', '起票日', '回答期限', '回答日', 'ステータス', 'ボール') | ForEach-Object { @{ Header = $_; Width = 12 } }
    $rows = @(
        @('Q-01', 'QA', '締め日の扱いを確認したい', '2026/10/19', '2026/10/22', '', '回答待ち', '田中'),
        @('Q-02', 'QA', '税の端数処理', '2026/10/15', '2026/10/20', '2026/10/21', '回答済', '田中'),
        @('K-01', '課題', '帳票の出力先が未定', '2026/10/20', '2026/10/28', '', '対応中', '佐藤'),
        @('K-02', '課題', '性能の目標値', '2026/10/21', '', '', '取下げ', '佐藤'))
    Export-Workbook -FilePath (Join-Path $dir '課題管理表.xlsx') -Sheets @(@{ Name = '課題'; Columns = $cols; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    @('項目,値', '課題管理表,課題管理表.xlsx') | Set-Content -LiteralPath (Join-Path $dir '週次報告の設定.csv') -Encoding UTF8
    $text = Get-HostText { $script:items = @(& $WeeklyReport $dir -AsOf '2026-10-23') }
    Assert-Equal @('佐藤|期限間近|課題|K-01: 帳票の出力先が未定|10/28', '田中|期限切れ|課題|Q-01: 締め日の扱いを確認したい|10/22') @($script:items | ForEach-Object { '{0}|{1}|{2}|{3}|{4:MM/dd}' -f $_.Owner, $_.State, $_.Kind, $_.Title, $_.Date } | Sort-Object) '回答済み・取下げは載せない'
    Assert-Match $text '■ 課題・QA\n・今週: 起票 2 件・完了 1 件\n・未完了: 2 件・うち期限切れ 1 件\n  - 期限 10/22 Q-01: 締め日の扱いを確認したい（田中）\n'
    $xlsx = Join-Path $dir '週次報告/週次報告_20261023.xlsx'
    $people = Read-XlsxSheet -Path $xlsx -Index 2
    Assert-Equal @('概要', '担当者', '担当者別', '課題') $people.Sheets
    Assert-Equal '担当|遅延タスク|来週終了のタスク|課題の未完了|うち期限切れ' ((@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $people.Rows[1][$_] }) -join '|')
}

Test-Case 'New-WeeklyReport: 工数実績から今週・累計の工数を出し、WBS の予定工数を超えた・超えそうなタスクを担当者ごとのやることに入れる' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $rows = @(@('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '50%', '5'), @('1.3', 'IF 設計', '佐藤', '2026/10/19', '2026/10/23', '20%', '4'))
    Export-Workbook -FilePath (Join-Path $dir 'WBS.xlsx') -Sheets @(@{ Name = 'WBS'; Columns = $cols; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    @('日付,氏名,WBS,作業内容,時間', '2026/10/12,佐藤,1.2,,8', '2026/10/13,佐藤,1.2,,8', '2026/10/14,佐藤,1.2,,8', '2026/10/15,佐藤,1.2,,8', '2026/10/16,佐藤,1.2,,8',
        '2026/10/19,佐藤,1.2,,4', '2026/10/19,佐藤,1.3,,4', '2026/10/20,山田,,定例会議,1:30', '2026/10/22,佐藤,1.3,,6') | Set-Content -LiteralPath (Join-Path $dir '工数/工数実績.csv') -Encoding UTF8
    @('項目,値', 'WBS,WBS.xlsx', '工数実績,工数') | Set-Content -LiteralPath (Join-Path $dir '週次報告の設定.csv') -Encoding UTF8
    $text = Get-HostText { $script:items = @(& $WeeklyReport $dir -AsOf '2026-10-23') }
    Assert-Match $text '■ 工数\n・今週: 15\.5 時間（1\.94 人日）・累計 55\.5 時間\n・担当者別（今週）: 佐藤 14・山田 1\.5（時間）\n'
    Assert-Match $text '・予定工数のあるタスク 2 件: 予定 9 / 実績 6\.75（人日）・CPI 0\.49\n・予定超過 1 件・超過の見込み 1 件\n'
    Assert-Match $text '  - \[予定超過\] 1\.2 テーブル設計（佐藤）: 予定 5 / 実績 5\.5 / 進捗 50% → 完了見込 11\n  - \[超過の見込み\] 1\.3 IF 設計（佐藤）: 予定 4 / 実績 1\.25 / 進捗 20% → 完了見込 6\.25\n'
    Assert-Equal @('予定超過|1.2 テーブル設計|予定 5 / 実績 5.5 / 進捗 50% → 完了見込 11（人日）', '超過の見込み|1.3 IF 設計|予定 4 / 実績 1.25 / 進捗 20% → 完了見込 6.25（人日）') @($script:items | Where-Object { $_.Kind -eq '工数' } | ForEach-Object { '{0}|{1}|{2}' -f $_.State, $_.Title, $_.Note })
    Assert-Match $text '【佐藤】遅延 2・予定超過 1・超過の見込み 1'
    $xlsx = Join-Path $dir '週次報告/週次報告_20261023.xlsx'
    $sheet = Read-XlsxSheet -Path $xlsx -Index 1
    Assert-Equal @('概要', '担当者', '担当者別', 'タスク', '工数') $sheet.Sheets
    $effort = Read-XlsxSheet -Path $xlsx -Index 5
    Assert-Equal '佐藤|14|54' ((@('A', 'B', 'C') | ForEach-Object { $effort.Rows[2][$_] }) -join '|')
}

Test-Case 'New-WeeklyReport: 日付を列に並べた工数表（行がタスク、列が日付）からも今週・累計の工数と予定超過を出す' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('WBS', 'タスク名', '担当者', '予定開始日', '予定終了日', '進捗率', '予定工数') | ForEach-Object { @{ Header = $_; Width = 10 } }
    $rows = @(@('1.2', 'テーブル設計', '佐藤', '2026/10/12', '2026/10/16', '50%', '5'), @('1.3', 'IF 設計', '佐藤', '2026/10/19', '2026/10/23', '20%', '4'))
    Export-Workbook -FilePath (Join-Path $dir 'WBS.xlsx') -Sheets @(@{ Name = 'WBS'; Columns = $cols; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    New-Item -ItemType Directory -Path (Join-Path $dir '工数') | Out-Null
    New-SheetsBook (Join-Path $dir '工数/工数表.xlsx') ([ordered]@{
            '10月' = @(
                @('WBS番号', 'タスク名', '担当者', '10/12', '10/13', '10/14', '10/15', '10/16', '10/19', '10/20', '10/21', '10/22', '10/23', '合計'),
                @('1.2', '', '佐藤', [double]8, [double]8, [double]8, [double]8, [double]8, [double]4, '', '', '', '', [double]44),
                @('1.3', '', '佐藤', '', '', '', '', '', [double]4, '', '', [double]6, '', [double]10),
                @('', '定例会議', '山田', '', '', '', '', '', '', '1:30', '', '', '', [double]1.5))
        })
    @('項目,値', 'WBS,WBS.xlsx', '工数実績,工数') | Set-Content -LiteralPath (Join-Path $dir '週次報告の設定.csv') -Encoding UTF8
    $text = Get-HostText { $script:items = @(& $WeeklyReport $dir -AsOf '2026-10-23') }
    Assert-Match $text '■ 工数\n・今週: 15\.5 時間（1\.94 人日）・累計 55\.5 時間\n・担当者別（今週）: 佐藤 14・山田 1\.5（時間）\n'
    Assert-Match $text '  - \[予定超過\] 1\.2 テーブル設計（佐藤）: 予定 5 / 実績 5\.5 / 進捗 50% → 完了見込 11\n  - \[超過の見込み\] 1\.3 IF 設計（佐藤）: 予定 4 / 実績 1\.25 / 進捗 20% → 完了見込 6\.25\n'
}

Test-Case 'New-WeeklyReport: 和暦の設定（ja-JP の JapaneseCalendar）でも、日付は西暦で書く' {
    $message = ''
    try { Invoke-InWareki { & $WeeklyReport -Wbs (Join-Path (New-TestDir) 'WBS.xlsx') -Since '2026-10-20' -AsOf '2026-10-14' 6>$null } }
    catch { $message = $_.Exception.Message }
    Assert-Equal '-Since（2026/10/20）が -AsOf（2026/10/14）より後です' $message
}

Test-Case 'New-WeeklyReport: 地域の設定で日付の区切りを「-」にしていても、報告の日付は 10/23 の形で書く' {
    $dir = New-TestDir
    $p = New-WeeklyProject $dir
    $text = Invoke-InDashDate { Get-HostText { & $WeeklyReport -Wbs $p.Wbs -Findings $p.Findings -BugList $p.Bugs -AsOf '2026-10-23' -OutDir (Join-Path $dir '週次報告') | Out-Null } }
    Assert-Match $text '予定終了 10/23・見込 10/28'
    Assert-Match $text '・来週（10/24～10/30）の予定: 3 件'
}

Test-Case 'New-WeeklyReport: 指摘の状態の「未完了」「完了予定」は、完了を含んでも未完了として担当者ごとのやることに入れる' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('No', '設計書', '指摘内容', '対応状況', '対応期限', '対応者') | ForEach-Object { @{ Header = $_; Width = 12 } }
    $rows = @(@('1', 'DD02 画面設計書', '桁数が違う', '未完了', '2026/10/30', '山田'), @('2', 'DD02 画面設計書', '誤記', '完了', '', '山田'), @('3', 'DD03 テーブル定義書', '型が違う', '完了予定', '2026/10/30', '佐藤'))
    Export-Workbook -FilePath (Join-Path $dir '指摘管理表.xlsx') -Sheets @(@{ Name = '指摘'; Columns = $cols; Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    $items = @(& $WeeklyReport -Findings (Join-Path $dir '指摘管理表.xlsx') -AsOf '2026-10-23' -OutDir (Join-Path $dir '週次報告') 6>$null 3>$null)
    Assert-Equal @('佐藤|期限間近|DD03 テーブル定義書: 型が違う', '山田|期限間近|DD02 画面設計書: 桁数が違う') @($items | ForEach-Object { '{0}|{1}|{2}' -f $_.Owner, $_.State, $_.Title })
}

Test-Case 'New-WeeklyReport: Redmine の CSV（見出しの「#」）を課題・障害の管理表として読み、「終了」は完了、「キャンセル」は数えない' {
    $dir = New-TestDir
    @('#,トラッカー,ステータス,題名,担当者,開始日,期日,終了日',
        '101,タスク,終了,一覧の並び順を決める,佐藤,2026/10/05,2026/10/09,2026/10/08',
        '102,タスク,終了,登録の権限を決める,佐藤,2026/10/05,2026/10/12,2026/10/12',
        '103,タスク,キャンセル,税の端数の扱い,鈴木,2026/10/05,2026/10/12,',
        '104,タスク,進行中,合計の出し方,鈴木,2026/10/19,2026/10/30,') | Set-Content -LiteralPath (Join-Path $dir 'redmine_issues.csv') -Encoding UTF8
    @('#,トラッカー,ステータス,優先度,題名,担当者,期日,作成日,終了日',
        '201,バグ,終了,通常,一覧の並び順が逆,佐藤,2026/10/21,2026/10/19 10:15,2026/10/21 17:30',
        '202,バグ,キャンセル,通常,税の端数,鈴木,2026/10/20,2026/10/19 11:00,',
        '203,バグ,新規,高め,登録できない,鈴木,2026/10/22,2026/10/20 09:00,') | Set-Content -LiteralPath (Join-Path $dir 'redmine_bugs.csv') -Encoding UTF8
    $text = Get-HostText { $script:items = @(& $WeeklyReport -IssueList (Join-Path $dir 'redmine_issues.csv') -BugList (Join-Path $dir 'redmine_bugs.csv') -AsOf '2026-10-23') }
    Assert-Equal @('鈴木|期限切れ|障害|203: 登録できない|10/22', '鈴木|期限間近|課題|104: 合計の出し方|10/30') @($script:items | ForEach-Object { '{0}|{1}|{2}|{3}|{4:MM/dd}' -f $_.Owner, $_.State, $_.Kind, $_.Title, $_.Date } | Sort-Object) '終了・キャンセルは載せない'
    Assert-Match $text '■ 課題・QA\n・今週: 起票 0 件・完了 0 件\n・未完了: 1 件・うち期限切れ 0 件\n'
    Assert-Match $text '・今週: 起票 2 件・完了 1 件'
}

Test-Case 'New-WeeklyReport: 英字の大文字・小文字だけが違う担当者（Nguyen・NGUYEN・nguyen）は同じ人として 1 つにまとめる' {
    $dir = New-TestDir
    @('障害番号,障害内容,状態,起票日,担当者,修正予定日', 'B-1,登録できない,未対応,2026/10/20,Nguyen,2026/10/21', 'B-2,合計が違う,未対応,2026/10/20,NGUYEN,2026/10/21') | Set-Content -LiteralPath (Join-Path $dir '障害管理表.csv') -Encoding UTF8
    @('課題No,件名,状態,回答期限,担当', 'Q-1,締め日の確認,回答待ち,2026/10/22,nguyen') | Set-Content -LiteralPath (Join-Path $dir '課題管理表.csv') -Encoding UTF8
    $out = Join-Path $dir '週次報告'
    $text = Get-HostText { $null = & $WeeklyReport -BugList (Join-Path $dir '障害管理表.csv') -IssueList (Join-Path $dir '課題管理表.csv') -AsOf '2026-10-23' -OutDir $out }
    Assert-Equal @('【Nguyen】期限切れ 3') @($text -split "`n" | Where-Object { $_ -match '^【' }) '表示は最初に出てきた書き方'
    $people = Read-XlsxSheet -Path (Join-Path $out '週次報告_20261023.xlsx') -Index 2
    Assert-Equal @('Nguyen|0|0|2|2|1|1') @(2..$people.Rows.Count | ForEach-Object { $row = $people.Rows[$_]; (@('A', 'B', 'C', 'D', 'E', 'F', 'G') | ForEach-Object { $row[$_] }) -join '|' }) '担当・遅延・来週終了・障害の未完了・うち期限切れ・課題の未完了・うち期限切れ'
    $ownerDir = Join-Path $out '週次報告_20261023_担当者別'
    Assert-Equal @('Nguyen.txt') @(Get-ChildItem -LiteralPath $ownerDir -File | ForEach-Object { $_.Name })
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $ownerDir 'Nguyen.txt'))) '■ 遅延・期限切れ（3 件）'
}

Test-Case 'New-WeeklyReport: 担当に 2 人を書いた課題（山田、佐藤・山田さん、佐藤・山田・佐藤さん）は、1 人ずつの担当者のやることに入れる（2 人で 1 人の担当者にしない）' {
    $dir = New-TestDir
    @('No,区分,件名,担当,期限,状態,起票日,完了日',
        '1,宿題,メッセージ一覧の文言をそろえる,"山田さん、佐藤",2026/10/20,未対応,2026/10/13,',
        '2,宿題,IF 定義書を直す,山田・佐藤さん,2026/10/28,未対応,2026/10/13,',
        '3,宿題,移行手順を書く,ジョン・スミス,2026/10/28,未対応,2026/10/13,',
        '4,宿題,帳票を直す,"山田、佐藤",2026/10/16,完了,2026/10/13,2026/10/20') | Set-Content -LiteralPath (Join-Path $dir '宿題.csv') -Encoding UTF8
    $out = Join-Path $dir '週次報告'
    $text = Get-HostText { $script:items = @(& $WeeklyReport -IssueList (Join-Path $dir '宿題.csv') -AsOf '2026-10-23' -OutDir $out) }
    $heads = @($text -split "`n" | Where-Object { $_ -match '^【' })
    Assert-Equal 3 $heads.Count ($heads -join ' / ')
    foreach ($h in @('【山田】期限切れ 1・期限間近 1', '【佐藤】期限切れ 1・期限間近 1', '【ジョン・スミス】期限間近 1')) { Assert-True ($heads -contains $h) ('{0} がありません: {1}' -f $h, ($heads -join ' / ')) }
    Assert-Match $text '・今週: 起票 0 件・完了 1 件\n・未完了: 3 件・うち期限切れ 1 件\n  - 期限 10/20 1: メッセージ一覧の文言をそろえる（山田、佐藤）\n'
    $ownerDir = Join-Path $out '週次報告_20261023_担当者別'
    Assert-Equal 3 @(Get-ChildItem -LiteralPath $ownerDir -File).Count
    foreach ($n in @('山田', '佐藤', 'ジョン・スミス')) { Assert-True (Test-Path -LiteralPath (Join-Path $ownerDir ($n + '.txt'))) "$n.txt がありません" }
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $ownerDir '佐藤.txt'))) '\[期限切れ\] 課題: 1: メッセージ一覧の文言をそろえる'
    $people = Read-XlsxSheet -Path (Join-Path $out '週次報告_20261023.xlsx') -Index 2
    $rows = @(2..$people.Rows.Count | ForEach-Object { $row = $people.Rows[$_]; (@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $row[$_] }) -join '|' })
    foreach ($r in @('山田|0|0|2|1', '佐藤|0|0|2|1')) { Assert-True ($rows -contains $r) ('担当者のシートに {0} がありません: {1}' -f $r, ($rows -join ' / ')) }
    $issues = Read-XlsxSheet -Path (Join-Path $out '週次報告_20261023.xlsx') -Index 4
    Assert-Equal '山田、佐藤' $issues.Rows[2]['G'] '課題のシートの担当は 1 人ずつの名前をつなぐ'
}

Test-Case 'New-WeeklyReport: Excel のセル内で改行した指摘（↵）は、報告の一覧では空白にして 1 行に書く' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $cols = @('No', '設計書', '指摘内容', '対応状況', '対応期限', '対応者') | ForEach-Object { @{ Header = $_; Width = 12 } }
    Export-Workbook -FilePath (Join-Path $dir '指摘管理表.xlsx') -Sheets @(@{ Name = '指摘'; Columns = $cols; Rows = @(@{ Values = @('1', 'DD03 テーブル定義書', "型が誤り`n桁数も違う", '未対応', '2026/10/21', '佐藤') }) })
    $text = Get-HostText { $null = & $WeeklyReport -Findings (Join-Path $dir '指摘管理表.xlsx') -AsOf '2026-10-23' }
    Assert-Match $text '  - 期限 10/21 DD03 テーブル定義書: 型が誤り 桁数も違う（佐藤）\n'
    Assert-Match $text '【佐藤】期限切れ 1\n  - \[期限切れ\] 指摘: DD03 テーブル定義書: 型が誤り 桁数も違う（期限 10/21・未対応）'
}

Test-Case 'New-WeeklyReport: 年の無い期限（12/25・1/12(火)）は基準日の前後半年の年とみなす（年明けの報告で、前の年の 12/25 は期限切れ）' {
    $dir = New-TestDir
    @('課題No,件名,状態,回答期限,担当', 'Q-1,締め日の確認,回答待ち,12/25,田中', 'Q-2,端数処理,回答待ち,1/12(火),田中') | Set-Content -LiteralPath (Join-Path $dir '課題管理表.csv') -Encoding UTF8
    $text = Get-HostText { $script:items = @(& $WeeklyReport -IssueList (Join-Path $dir '課題管理表.csv') -AsOf '2027-01-08') }
    Assert-Equal @('期限切れ|Q-1: 締め日の確認|2026/12/25', '期限間近|Q-2: 端数処理|2027/01/12') @($script:items | ForEach-Object { '{0}|{1}|{2}' -f $_.State, $_.Title, $(if ($_.Date) { $_.Date.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) }) })
    Assert-Match $text '・未完了: 2 件・うち期限切れ 1 件\n  - 期限 12/25 Q-1: 締め日の確認（田中）\n'
}

Test-Case 'New-WeeklyReport: Get-Help でヘルプが表示される' {
    $h = Get-Help $WeeklyReport
    Assert-Match $h.Synopsis '週次報告'
    Assert-Equal 16 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
