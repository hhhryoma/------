#     New-QualityReport.ps1 のテスト。設計書は docs\design-sample、テスト仕様書・障害管理表・指摘管理表・要件一覧はテストの中で作る

$ReportTool = Join-Path $script:ToolsDir 'New-QualityReport.ps1'
$reportDesignDir = Join-Path $script:RepoRoot 'docs/design-sample'

function New-ReportBook {
    param([string] $FilePath, [string] $Sheet, [string[]] $Header, $Rows)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $FilePath -Sheets @(@{
            Name = $Sheet; Columns = @($Header | ForEach-Object { @{ Header = $_; Width = 12 } }); Rows = @($Rows | ForEach-Object { @{ Values = $_ } })
        })
}

function New-ReportData {
    #     テスト仕様書（6 ケース。NG 1）・障害管理表（2 件。未完了 1・設計で入りテストで見つかったもの 1）・指摘管理表（3 件。期限切れ 1）・要件一覧
    param([string] $Dir)
    New-Item -ItemType Directory -Force -Path (Join-Path $Dir 'テスト'), (Join-Path $Dir 'レビュー') | Out-Null
    $texts = @('初期表示で受注日の初期値を確かめる', '検索ボタンで受注を検索する（REQ-001）', '受注番号に全角を入れると E0003 が出る', '登録ボタンで受注を登録する', '登録で必須の未入力はエラーになる', 'キャンセルで一覧に戻る')
    $functions = @('SCR001', '', '', 'SCR002', '', 'SCR003')
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le 6; $i++) {
        $result = 'OK'; $bug = ''
        if ($i -eq 4) { $result = 'NG'; $bug = 'B002' }
        $rows.Add(@([string]$i, $functions[$i - 1], $texts[$i - 1], '画面を開く', '一覧が表示される', ('2026/10/0' + $i), ('2026/10/0' + $i), $result, $bug))
    }
    New-ReportBook (Join-Path $Dir 'テスト\UT_受注.xlsx') 'ケース' @('No', '機能', '確認内容', '手順', '期待結果', '予定日', '実施日', '結果', '障害番号') $rows
    New-ReportBook (Join-Path $Dir '障害管理表.xlsx') '障害' @('障害番号', '障害内容', '起票日', '機能', '混入工程', '摘出工程', '状態') @(
        , @('B001', '一覧の件数が違う', '2026/10/01', 'SCR001', '製造', '単体テスト', '完了')
        , @('B002', '登録できない', '2026/10/04', 'SCR002', '詳細設計', '単体テスト', '対応中'))
    New-ReportBook (Join-Path $Dir 'レビュー\指摘管理表.xlsx') '指摘' @('No', '指摘日', '対象成果物', '指摘内容', '指摘区分', '工程', '対応期限', '対応状況') @(
        , @('1', '2026/10/01', 'DD02_画面設計書_SCR001_受注一覧', '桁数が違う', '誤り', '詳細設計レビュー', '2026/10/10', '完了')
        , @('2', '2026/10/01', 'DD02_画面設計書_SCR001_受注一覧', 'イベントが足りない', '漏れ', '詳細設計レビュー', '2026/10/05', '対応中')
        , @('3', '2026/10/02', 'DD03_テーブル定義書', '型が違う', '誤り', '詳細設計レビュー', '2026/10/30', '未対応'))
    @('要件ID,要件名', 'REQ-001,受注を検索する', 'REQ-003,出荷指示を送る') | Set-Content -LiteralPath (Join-Path $Dir '要件一覧.csv') -Encoding UTF8
    @('設計書,レビュー工数', 'DD02_画面設計書_SCR001_受注一覧,2', 'DD03_テーブル定義書,1') | Set-Content -LiteralPath (Join-Path $Dir 'レビュー記録.csv') -Encoding UTF8
}

Test-Case 'New-QualityReport: 各ツールを実行し、指標ごとの実績と判定（○△×）・要対応を 1 冊にまとめる' {
    $dir = New-TestDir
    New-ReportData $dir
    $out = Join-Path $dir '報告'
    $r = @(& $ReportTool -OutDir $out -DesignPath $reportDesignDir -Findings (Join-Path $dir 'レビュー') -ReviewLog (Join-Path $dir 'レビュー記録.csv') -TestPath (Join-Path $dir 'テスト') -BugList (Join-Path $dir '障害管理表.xlsx') -Requirement (Join-Path $dir '要件一覧.csv') -AsOf '2026-10-20' 3>$null 6>$null)
    $got = @{}
    foreach ($m in $r) { $got[$m.Name] = '{0}|{1}' -f $m.Value, $m.Judgement }
    Assert-Equal '3|' $got['指摘件数']
    Assert-Equal '2|△' $got['未完了の指摘']
    Assert-Equal '1|×' $got['期限切れの指摘'] '期限切れが残っていれば ×'
    Assert-Equal '1|○' $got['消化率']
    Assert-Match $got['レビュー実施率'] '^0\.\d+\|×$' 'レビュー記録か指摘があるのは DD02 の SCR001 と DD03 だけ'
    Assert-Equal '1|△' $got['未完了のバグ']
    Assert-Match $got['バグの収束（信頼度成長曲線）'] '^まだ判定できない（バグ 2 件.*\|$' 'バグが少ないうちは判定しない'
    Assert-Equal '0|' $got['直近 7 日の新規バグ']
    Assert-Equal '0|○' $got['テスト仕様書と障害管理表の食い違い']
    Assert-Equal '0|○' $got['エラー（期待結果が空など）']
    Assert-Equal '1|△' $got['後の工程への流出'] '詳細設計で入り単体テストで見つかった B002'
    Assert-Equal '1|○' $got['テストのある画面の割合']
    Assert-Equal '0|×' $got['設計とテストのある要件の割合'] 'docs\design-sample には要件 ID が書かれていない'
    Assert-Equal @('レビュー', 'テスト', 'テスト仕様書', '工程', '網羅', '設計書') @($r | ForEach-Object { $_.Area } | Select-Object -Unique)

    $report = Join-Path $out '品質評価_20261020.xlsx'
    Assert-True (Test-Path -LiteralPath $report) '品質評価のブック'
    # 設計で入ったバグ B002（SCR002・詳細設計）は、SCR002 の設計書のレビューの見逃し
    Assert-Equal '1|△' $got['レビューで見逃した設計書']
    Assert-Equal '1|○' $got['設計工程の摘出率'] '指摘 3 件 ÷ (3 + 0)。B002 はレビューしていない SCR002 のバグなので入れない（レビューで見逃した設計書で知らせる）'
    foreach ($name in @('レビュー.xlsx', 'テスト.xlsx', 'テスト仕様書のチェック.xlsx', '工程別の欠陥.xlsx', 'レビューの効き目.xlsx', '網羅.xlsx', '設計書の整合性.xlsx')) {
        Assert-True (Test-Path -LiteralPath (Join-Path $out "詳細\$name")) "詳細\$name"
    }
    $book = Read-XlsxSheet -Path $report
    Assert-Equal @('品質評価', '要対応', '対象') $book.Sheets
    Assert-Equal '区分|指標|実績|判定|判定の考え方|定性分析（原因）|対策|詳細' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal $r.Count ($book.Rows.Count - 1) '指標の行'
    # 詳細の列は、各ツールの結果のブックを開くリンク
    Assert-Equal $r.Count @($book.Links | Where-Object { $_.ref -like 'H*' }).Count
    Assert-Match (@($book.Links)[0].target) 'レビュー\.xlsx$'

    $todo = Read-XlsxSheet -Path $report -Index 2
    $todoText = @($todo.Rows.Values | ForEach-Object { '{0}|{1}|{2}' -f $_['A'], $_['B'], $_['C'] })
    Assert-True (@($todoText | Where-Object { $_ -like '工程|詳細設計|*1 件が後の工程で見つかった*' }).Count -eq 1) ($todoText -join "`n")
    Assert-True (@($todoText | Where-Object { $_ -like '網羅|テストの無い要素*' }).Count -eq 1) ($todoText -join "`n")
    Assert-True (@($todoText | Where-Object { $_ -like 'レビュー|DD05_メッセージ一覧*|未レビュー*' }).Count -eq 1) ($todoText -join "`n")
}

Test-Case 'New-QualityReport: 失敗したツールがあっても、ほかのツールの結果でまとめる' {
    $dir = New-TestDir
    New-ReportData $dir
    $out = Join-Path $dir '報告'
    $warnings = $null
    $text = Get-HostText { & $ReportTool -OutDir $out -Findings (Join-Path $dir 'レビュー') -ReviewLimits (Join-Path $dir '無い基準値.csv') -BugList (Join-Path $dir '障害管理表.xlsx') -AsOf '2026-10-20' -WarningVariable script:reportWarnings -WarningAction SilentlyContinue | Out-Null }
    Assert-Match $text '失敗したツール: Measure-ReviewQuality'
    Assert-True (@($script:reportWarnings | Where-Object { "$_" -match 'Measure-ReviewQuality が失敗しました' }).Count -eq 1) '失敗を警告で知らせる'
    $book = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx')
    Assert-Equal @('工程') @($book.Rows.Keys | Where-Object { $_ -gt 1 } | ForEach-Object { $book.Rows[$_]['A'] } | Select-Object -Unique) 'レビューの行は無く、工程の行はある'
    $scope = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx') -Index 3
    Assert-True (@($scope.Rows.Values | Where-Object { $_['A'] -eq '失敗したツール' -and $_['B'] -like 'Measure-ReviewQuality（*' }).Count -eq 1) '対象のシートに失敗したツール'

    Assert-Throws { & $ReportTool -OutDir $out 6>$null } '-Findings'
}

Test-Case 'New-QualityReport: ツールの結果をまとめる途中でエラーになっても、読み取り結果の使い回しを残さない' {
    $dir = New-TestDir
    # tools を写し、詳細は保存するが、指摘件数が数でない結果を返す Measure-ReviewQuality.ps1 に置き換える（集計の Measure-Object で止まる）
    $tools = Join-Path $dir 'tools'
    Copy-Item -LiteralPath $script:ToolsDir -Destination $tools -Recurse
    Set-Content -LiteralPath (Join-Path $tools 'Measure-ReviewQuality.ps1') -Value "param(`$Path, `$AsOf, `$OutXlsx) Set-Content -LiteralPath `$OutXlsx -Value 'x'; [pscustomobject]@{ Document = 'x'; Findings = 'abc'; Open = 0; Overdue = 0 }" -Encoding UTF8
    New-Item -ItemType Directory -Path (Join-Path $dir 'レビュー') | Out-Null
    Assert-Throws { & (Join-Path $tools 'New-QualityReport.ps1') -OutDir (Join-Path $dir '報告') -Findings (Join-Path $dir 'レビュー') 3>$null 6>$null } 'abc'
    Assert-True ($null -eq (Get-Variable -Name ExcelToolsSearchCache -Scope Global -ErrorAction SilentlyContinue)) '読み取り結果を残している'
}

Test-Case 'New-QualityReport: 判定基準の CSV で ○△× の基準を変えられ、判定の考え方に基準を書く' {
    $dir = New-TestDir
    New-ReportData $dir
    $criteria = Join-Path $dir '判定基準.csv'
    @('指標,向き,○,△', '未完了の指摘,以下,0,1', '摘出すべき工程までに見つかった割合,以上,80%,') | Set-Content -LiteralPath $criteria -Encoding UTF8
    $out = Join-Path $dir '報告'
    $r = @(& $ReportTool -OutDir $out -Findings (Join-Path $dir 'レビュー') -BugList (Join-Path $dir '障害管理表.xlsx') -Criteria $criteria -AsOf '2026-10-20' 3>$null 6>$null)
    $got = @{}
    foreach ($m in $r) { $got[$m.Name] = '{0}|{1}' -f $m.Value, $m.Judgement }
    Assert-Equal '2|×' $got['未完了の指摘'] '未完了 2 件は △ の上限（1 件）を超える'
    Assert-Equal '0.8|○' $got['摘出すべき工程までに見つかった割合']
    Assert-Equal '1|' $got['期限切れの指摘'] '判定基準に無い指標は判定しない'
    $book = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx')
    $basis = @{}
    foreach ($row in $book.Rows.Values) { $basis[$row['B']] = $row['E'] }
    Assert-Equal '0 件以下なら ○、1 件以下なら △、それより多ければ ×' $basis['未完了の指摘']
    Assert-Equal '4 / 5 件。80% 以上なら ○、それ未満は △' $basis['摘出すべき工程までに見つかった割合']

    @('指標,向き,○,△', '未完了の指摘,多い,0,') | Set-Content -LiteralPath $criteria -Encoding UTF8
    Assert-Throws { & $ReportTool -OutDir $out -Findings (Join-Path $dir 'レビュー') -Criteria $criteria 6>$null } '向きは「以上」か「以下」'
}

Test-Case 'New-QualityReport: -AiReview で AI のレビュー結果と人の指摘を突き合わせ、一致率と観点表への反映案を出す' {
    $dir = New-TestDir
    New-ReportData $dir
    $scr = 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    New-ReportBook (Join-Path $dir 'レビュー結果.xlsx') '指摘一覧' @('No', '判定', '観点ID', '設計書', '場所', '指摘内容') @(
        , @('1', 'NG', 'S01', $scr, '画面項目定義!AB11', '桁数が違う')
        , @('2', 'NG', 'S02', $scr, '画面項目定義!AK7', '数量のエラー時のメッセージが無い')
        , @('3', 'NG', 'S02', $scr, '画面項目定義!AK8', '単価のエラー時のメッセージが無い'))
    $out = Join-Path $dir '報告'
    $r = @(& $ReportTool -OutDir $out -Findings (Join-Path $dir 'レビュー') -AiReview (Join-Path $dir 'レビュー結果.xlsx') -AsOf '2026-10-20' 3>$null 6>$null)
    $got = @{}
    foreach ($m in @($r | Where-Object Area -eq 'AI レビュー')) { $got[$m.Name] = $m.Value }
    Assert-Equal 0.333 $got['人の指摘のうち AI も見つけた割合'] '人の指摘 3 件のうち 1 件'
    Assert-Equal 0.333 $got['AI の指摘のうち人と一致した割合'] 'AI の指摘 3 件のうち 1 件'
    Assert-Equal 3 $got['観点表への反映案'] '判定基準を具体的にする（S04）・観点を足す（型が違う）・判定基準を見直す（S02）'
    Assert-True (Test-Path -LiteralPath (Join-Path $out '詳細/AIと人の指摘.xlsx')) '突き合わせの詳細'
    $actions = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx') -Index 2
    $s02 = @($actions.Rows.Values | Where-Object { $_['A'] -eq 'AI レビュー' -and $_['B'] -eq 'S02' })
    Assert-Equal 1 $s02.Count 'AI だけの指摘が多い S02 を要対応に並べる'
    Assert-Match $s02[0]['C'] '^判定基準を見直す（2 件）: 数量のエラー時のメッセージが無い'

    # 観点別の一致率: 実績の CSV に残し、前の時点と並べる
    $history = Join-Path $dir '実績.csv'
    foreach ($asOf in @('2026-10-06', '2026-10-20')) { $null = & $ReportTool -OutDir (Join-Path $dir '報告3') -Findings (Join-Path $dir 'レビュー') -AiReview (Join-Path $dir 'レビュー結果.xlsx') -HistoryCsv $history -Project 'A' -AsOf $asOf 3>$null 6>$null }
    $book = Read-XlsxSheet -Path (Join-Path $dir '報告3/品質評価_20261020.xlsx') -Index 5
    Assert-Equal @('品質評価', '要対応', '対象', '推移', '観点別の一致率') $book.Sheets
    Assert-Equal @('観点ID|AI の指摘|人と一致|2026/10/06|2026/10/20', 'S01|1|1|1|1', 'S02|2|0|0|0') @(1..3 | ForEach-Object { $r = $book.Rows[$_]; (@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $r[$_] }) -join '|' })
    Assert-Equal 2 @(Import-Csv -LiteralPath $history | Where-Object { $_.'区分' -eq 'AI レビュー（観点別）' -and $_.'時点' -eq '2026/10/20' }).Count '観点ごとの一致率を実績の CSV に残す'

    # -Findings が無ければ突き合わせない
    $text = Get-HostText { & $ReportTool -OutDir (Join-Path $dir '報告2') -TestPath (Join-Path $dir 'テスト') -AiReview (Join-Path $dir 'レビュー結果.xlsx') -AsOf '2026-10-20' -WarningVariable script:aiWarnings -WarningAction SilentlyContinue | Out-Null }
    Assert-Match ($script:aiWarnings -join ' ') '-Findings（人の指摘管理表）と一緒に'
}

Test-Case 'New-QualityReport: -HistoryCsv に実績を足し、前回の値・増減と推移を出す（New-QualityLimits にそのまま渡せる）' {
    $dir = New-TestDir
    New-ReportData $dir
    $pages = Join-Path $dir 'pages.csv'
    @('"ファイル","ページ数"', '"DD02_画面設計書_SCR001_受注一覧.pdf","4"', '"DD03_テーブル定義書.pdf","10"') | Set-Content -LiteralPath $pages -Encoding UTF8
    $history = Join-Path $dir '実績.csv'
    $base = @{ Findings = (Join-Path $dir 'レビュー'); PageCsv = $pages; BugList = (Join-Path $dir '障害管理表.xlsx'); HistoryCsv = $history; Project = 'A' }
    & $ReportTool @base -OutDir (Join-Path $dir '1') -AsOf '2026-10-06' 3>$null 6>$null | Out-Null
    $text = Get-HostText { & $ReportTool @base -OutDir (Join-Path $dir '2') -AsOf '2026-10-20' 3>$null | Out-Null }
    Assert-Match $text '実績の CSV に \d+ 件を足しました（プロジェクト「A」・2026/10/20。前回 2026/10/06 と比べました）'
    Assert-Match $text '同じ Excel の読み直し \d+ 回を省きました' '指摘管理表は 3 つのツールが読む'
    Assert-True ($null -eq (Get-Variable -Name ExcelToolsSearchCache -Scope Global -ErrorAction SilentlyContinue)) '読み取り結果を残している'

    $rows = @(Import-Csv -LiteralPath $history)
    Assert-Equal @('2026/10/06', '2026/10/20') @($rows | ForEach-Object { $_.'時点' } | Sort-Object -Unique)
    $density = @($rows | Where-Object { $_.'指標' -eq '指摘密度' -and $_.'時点' -eq '2026/10/20' })[0]
    Assert-Equal 'ページ数|0.214' ('{0}|{1}' -f $density.'分母', $density.'値') '指摘 3 件 ÷ 14 ページ'

    $book = Read-XlsxSheet -Path (Join-Path $dir '2\品質評価_20261020.xlsx')
    Assert-Equal @('品質評価', '要対応', '対象', '推移') $book.Sheets
    Assert-Equal '実績|前回（2026/10/06）|増減|判定' ((@('C', 'D', 'E', 'F') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    # 期限切れの指摘: 10/06 は 1 件（10/05 期限）、10/20 も 1 件で増減 0
    $overdue = @($book.Rows.Values | Where-Object { $_['B'] -eq '期限切れの指摘' })[0]
    Assert-Equal '1|1|0' ('{0}|{1}|{2}' -f $overdue['C'], $overdue['D'], $overdue['E'])
    Assert-True (@($book.Links | Where-Object { $_.Ref -like 'J*' }).Count -gt 0) '詳細のリンクは右にずれる'
    $trend = Read-XlsxSheet -Path (Join-Path $dir '2\品質評価_20261020.xlsx') -Index 4
    Assert-Equal '区分|指標|2026/10/06|2026/10/20' ((@('A', 'B', 'C', 'D') | ForEach-Object { $trend.Rows[1][$_] }) -join '|')

    # 同じ時点で実行し直すと置き換える
    $count = $rows.Count
    & $ReportTool @base -OutDir (Join-Path $dir '2') -AsOf '2026-10-20' 3>$null 6>$null | Out-Null
    Assert-Equal $count @(Import-Csv -LiteralPath $history).Count

    # New-QualityLimits はプロジェクトごとに最後の時点だけを使う
    $limits = @(& (Join-Path $script:ToolsDir 'New-QualityLimits.ps1') $history 6>$null)
    $l = @($limits | Where-Object { $_.Metric -eq '指摘密度' })[0]
    Assert-Equal 'ページ数|1|0.214' ('{0}|{1}|{2}' -f $l.Unit, $l.Count, $l.Mean)
}

Test-Case 'New-QualityReport: 設計工程の摘出率に、未レビューの設計書のバグを入れない（Measure-ReviewEffectiveness と同じ）' {
    $dir = New-TestDir
    @('設計書,指摘内容,分類,状態', 'SCR001_受注一覧,桁数が違う,誤り,完了', 'SCR001_受注一覧,型が違う,誤り,完了', 'SCR001_受注一覧,イベントの抜け,漏れ,完了', 'SCR001_受注一覧,遷移先の誤り,誤り,完了') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    @('ファイル,ページ数', 'SCR001_受注一覧.pdf,4', 'SCR002_受注登録.pdf,4') | Set-Content -LiteralPath (Join-Path $dir 'pages.csv') -Encoding UTF8
    @('設計書,工数', 'SCR001_受注一覧,2') | Set-Content -LiteralPath (Join-Path $dir 'reviewlog.csv') -Encoding UTF8
    @('障害番号,障害内容,設計書,混入工程,状態', 'B1,一覧のソート順,SCR001_受注一覧,詳細設計,完了', 'B2,登録の桁あふれ,SCR002_受注登録,詳細設計,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告') -Findings (Join-Path $dir 'findings.csv') -BugList (Join-Path $dir 'bugs.csv') -PageCsv (Join-Path $dir 'pages.csv') -ReviewLog (Join-Path $dir 'reviewlog.csv') -AsOf '2026-10-20' 3>$null 6>$null)
    $rate = @($r | Where-Object Name -eq '設計工程の摘出率')
    Assert-Equal 1 $rate.Count
    Assert-Equal 0.8 $rate[0].Value '指摘 4 件 ÷ (4 + レビューした SCR001 のバグ 1 件)。レビューしていない SCR002 のバグ B2 は入れない'
}

Test-Case 'New-QualityReport: -Config の設定ファイル（相対パス・| 区切り）で資料を指定し、コマンドの指定を優先する' {
    $dir = Join-Path (New-TestDir) '受注システム'
    New-Item -ItemType Directory -Path $dir | Out-Null
    New-ReportData $dir
    @('項目,値,説明', '指摘管理表,レビュー\指摘管理表.xlsx,', '障害管理表,障害管理表.xlsx|無い.csv,', '実績,品質報告\実績.csv,', 'タイトル,週次品質,') | Set-Content -LiteralPath (Join-Path $dir '品質評価の設定.csv') -Encoding UTF8
    $r = @(& $ReportTool -Config $dir -AsOf '2026-10-20' 3>$null 6>$null)
    Assert-True ($r.Count -gt 0) '指標'
    $report = Join-Path $dir '品質報告/週次品質_20261020.xlsx'
    Assert-True (Test-Path -LiteralPath $report) '設定の出力先とタイトル'
    $history = @(Import-Csv -LiteralPath (Join-Path $dir '品質報告/実績.csv'))
    Assert-Equal '受注システム' $history[0].'プロジェクト' 'プロジェクトの既定は設定ファイルのフォルダの名前'

    # コマンドで指定した出力先を優先する
    $other = Join-Path $dir '別の出力先'
    & $ReportTool -Config (Join-Path $dir '品質評価の設定.csv') -OutDir $other -AsOf '2026-10-20' 3>$null 6>$null | Out-Null
    Assert-True (Test-Path -LiteralPath (Join-Path $other '週次品質_20261020.xlsx')) '-OutDir を優先'

    @('項目,値', '知らない項目,x') | Set-Content -LiteralPath (Join-Path $dir '品質評価の設定.csv') -Encoding UTF8
    Assert-Throws { & $ReportTool -Config $dir 6>$null } '知らない項目'
    Assert-Throws { & $ReportTool -Config (Join-Path $dir '無い') 6>$null } '設定ファイルが見つかりません'
}

Test-Case 'New-QualityReport: 前回の詳細を Excel で開いていて保存できなかったツールは失敗にし、前回の詳細の値を今回の結果にしない' {
    $dir = New-TestDir
    New-Item -ItemType Directory -Path (Join-Path $dir 'テスト') | Out-Null
    # 10 ケースのうち、前から $done 件を実施した（OK）テスト仕様書
    $writeCases = { param([int] $done) @(@('No,機能,確認内容,結果') + @(1..10 | ForEach-Object { '{0},SCR001,確認{0},{1}' -f $_, $(if ($_ -le $done) { 'OK' } else { '' }) })) | Set-Content -LiteralPath (Join-Path $dir 'テスト/cases.csv') -Encoding UTF8 }
    $out = Join-Path $dir '報告'
    & $writeCases 5
    $r = @(& $ReportTool -OutDir $out -TestPath (Join-Path $dir 'テスト') -AsOf '2026-10-13' 3>$null 6>$null)
    Assert-Equal 0.5 @($r | Where-Object Name -eq '消化率')[0].Value
    & $writeCases 10
    # Excel で前回の詳細を開いたときと同じく、書き込みは拒み、読み取りは許す
    $lock = [System.IO.File]::Open((Join-Path $out '詳細/テスト.xlsx'), [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    try { $text = Get-HostText { $script:staleResult = @(& $ReportTool -OutDir $out -TestPath (Join-Path $dir 'テスト') -AsOf '2026-10-20' 3>$null) } }
    finally { $lock.Dispose() }
    Assert-Equal 0 @($script:staleResult | Where-Object Name -eq '消化率').Count '前回の詳細の消化率（0.5）を今回の値にしている'
    Assert-Match $text '失敗したツール: Measure-TestQuality（詳細を保存できませんでした'
    Assert-True (@($script:staleResult | Where-Object Area -eq 'テスト仕様書').Count -gt 0) '保存できたツールの結果はまとめる'
}

Test-Case 'New-QualityReport: 資料のパスが見つからない・資料から何も読めなかったツールは失敗にし、0 件として ○ にしない' {
    $dir = New-TestDir
    # 設定ファイルの書き間違いなどで、資料のフォルダが無い
    $text = Get-HostText { $script:unreadResult = @(& $ReportTool -OutDir (Join-Path $dir '報告1') -Findings (Join-Path $dir '無いレビュー') -TestPath (Join-Path $dir '無いテスト') -AsOf '2026-10-20' 3>$null) }
    Assert-Equal '' ((@($script:unreadResult | ForEach-Object { $_.Name })) -join '・') '読まなかった資料の指標'
    foreach ($tool in @('Measure-ReviewQuality', 'Measure-TestQuality', 'Test-TestSpec')) { Assert-Match $text ('{0}（見つかりません: [^）]*無い' -f $tool) }

    # 見出しが読めない指摘管理表・テスト仕様書、定義も参照も無い設計書のフォルダ
    New-Item -ItemType Directory -Path (Join-Path $dir 'テスト'), (Join-Path $dir '設計書') | Out-Null
    @('項目番号,確認すること,OK/NG', '1,一覧が表示される,NG', '2,検索できる,') | Set-Content -LiteralPath (Join-Path $dir 'テスト/cases.csv') -Encoding UTF8
    @('番号,コメント,担当', '1,桁数が違う,山田') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    New-ReportBook (Join-Path $dir '設計書/メモ.xlsx') 'メモ' @('項目', '内容') @(, @('打合せ', '来週に決める'))
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告2') -Findings (Join-Path $dir 'findings.csv') -TestPath (Join-Path $dir 'テスト') -DesignPath (Join-Path $dir '設計書') -AsOf '2026-10-20' 3>$null 6>$null)
    Assert-Equal '' ((@($r | ForEach-Object { '{0} {1}' -f $_.Name, $_.Judgement })) -join '・') '読めなかった資料の指標'
    $scope = Read-XlsxSheet -Path (Join-Path $dir '報告2/品質評価_20261020.xlsx') -Index 3
    $failedTools = [string]@($scope.Rows.Values | Where-Object { $_['A'] -eq '失敗したツール' })[0]['B']
    $ranTools = [string]@($scope.Rows.Values | Where-Object { $_['A'] -eq '実行したツール' })[0]['B']
    foreach ($tool in @('Measure-ReviewQuality（指摘を読めませんでした', 'Measure-TestQuality（テストケースを読めませんでした', 'Test-TestSpec（テストケースを読めませんでした', 'Test-DesignConsistency（設計書から定義も参照も読めませんでした')) {
        Assert-True ($failedTools.Contains($tool)) $failedTools
        Assert-True (-not $ranTools.Contains($tool.Split('（')[0])) $ranTools
    }

    # 見出しを読めた指摘管理表に指摘が 0 件なら、失敗にせず 0 件として判定する
    New-ReportBook (Join-Path $dir '指摘0件.xlsx') '指摘' @('No', '設計書', '指摘内容', '状態') @()
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告3') -Findings (Join-Path $dir '指摘0件.xlsx') -AsOf '2026-10-20' 3>$null 6>$null)
    Assert-Equal '0|○' ((@($r | Where-Object Name -eq '未完了の指摘' | ForEach-Object { '{0}|{1}' -f $_.Value, $_.Judgement })) -join ',') '指摘が 0 件の指摘管理表'
    $scope = Read-XlsxSheet -Path (Join-Path $dir '報告3/品質評価_20261020.xlsx') -Index 3
    Assert-Equal 0 @($scope.Rows.Values | Where-Object { $_['A'] -eq '失敗したツール' -and $_['B'] }).Count '失敗したツール'
}

Test-Case 'New-QualityReport: 割合は丸める前の値で判定し、実績に 100% 未満を 1 と書かない（判定基準ちょうどの 99.9% は ○）' {
    $dir = New-TestDir
    New-Item -ItemType Directory -Path (Join-Path $dir 'テスト') | Out-Null
    # 2000 ケースのうち、前から $done 件を実施した（OK）テスト仕様書
    $writeCases = {
        param([int] $done)
        $lines = New-Object System.Collections.Generic.List[string]
        $lines.Add('No,機能,確認内容,結果')
        for ($i = 1; $i -le 2000; $i++) { $lines.Add(('{0},SCR001,確認{0},{1}' -f $i, $(if ($i -le $done) { 'OK' } else { '' }))) }
        [System.IO.File]::WriteAllLines((Join-Path $dir 'テスト/cases.csv'), $lines, (New-Object System.Text.UTF8Encoding($true)))
    }
    & $writeCases 1999
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告') -TestPath (Join-Path $dir 'テスト') -AsOf '2026-10-20' 3>$null 6>$null)
    $m = @($r | Where-Object Name -eq '消化率')[0]
    Assert-Equal '0.999|△' ('{0}|{1}' -f $m.Value, $m.Judgement) '1999 / 2000 件（99.95%）を 1（100%）に丸めて ○ にしている'
    & $writeCases 1998
    $criteria = Join-Path $dir '判定基準.csv'
    @('指標,向き,○,△', '消化率,以上,99.9%,90%') | Set-Content -LiteralPath $criteria -Encoding UTF8
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告') -TestPath (Join-Path $dir 'テスト') -Criteria $criteria -AsOf '2026-10-20' 3>$null 6>$null)
    $m = @($r | Where-Object Name -eq '消化率')[0]
    Assert-Equal '0.999|○' ('{0}|{1}' -f $m.Value, $m.Judgement) '1998 / 2000 件は 99.9% ちょうど'
}

Test-Case 'New-QualityReport: 実績の CSV の時点を日付として読み、Excel で保存した 2026/10/6 の形でも前回と比べ、同じ時点の行を置き換える' {
    $dir = New-TestDir
    @('設計書,指摘内容,分類,状態', 'SCR001_受注一覧,桁数が違う,誤り,完了', 'SCR001_受注一覧,型が違う,誤り,対応中') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    $history = Join-Path $dir '実績.csv'
    $base = @{ Findings = (Join-Path $dir 'findings.csv'); HistoryCsv = $history; Project = 'A'; OutDir = (Join-Path $dir '報告') }
    & $ReportTool @base -AsOf '2026-10-06' 3>$null 6>$null | Out-Null
    # Excel で開いて保存すると、時点が 0 埋めの無い形になる
    [System.IO.File]::WriteAllText($history, ([System.IO.File]::ReadAllText($history) -replace '"2026/10/06"', '"2026/10/6"'), (New-Object System.Text.UTF8Encoding($true)))
    $text = Get-HostText { & $ReportTool @base -AsOf '2026-10-20' 3>$null | Out-Null }
    Assert-Match $text '前回 2026/10/06 と比べました'
    $count = @(Import-Csv -LiteralPath $history).Count
    & $ReportTool @base -AsOf '2026-10-06' 3>$null 6>$null | Out-Null
    $rows = @(Import-Csv -LiteralPath $history)
    Assert-Equal $count $rows.Count '同じ時点の行を置き換えずに足している'
    Assert-Equal @('2026/10/06', '2026/10/20') @($rows | ForEach-Object { $_.'時点' } | Sort-Object -Unique)
}

Test-Case 'New-QualityReport: 同じ時点のブックを作り直すと、記入した定性分析（原因）・対策と要対応の担当・期限を引き継ぐ（引き継げないものは知らせて前のブックを残す）' {
    $dir = New-TestDir
    New-ReportData $dir
    $out = Join-Path $dir '報告'
    $base = @{ OutDir = $out; Findings = (Join-Path $dir 'レビュー'); BugList = (Join-Path $dir '障害管理表.xlsx'); AsOf = '2026-10-20' }
    $null = & $ReportTool @base -TestPath (Join-Path $dir 'テスト') 3>$null 6>$null
    $report = Join-Path $out '品質評価_20261020.xlsx'
    $rowOf = { param($sheet, $col, $text) @($sheet.Rows.Keys | Where-Object { $sheet.Rows[$_][$col] -eq $text })[0] }
    $book = Read-XlsxSheet -Path $report
    $overdueRow = & $rowOf $book 'B' '期限切れの指摘'
    $progressRow = & $rowOf $book 'B' '消化率'
    # 会議の前に記入した（セルの中の改行も）
    Set-CellValues $report '品質評価' @{ "F$overdueRow" = "レビューの遅れ`n担当の不在"; "G$overdueRow" = '週次で確認する'; "F$progressRow" = '環境の遅れ' }
    $todo = Read-XlsxSheet -Path $report -Index 2
    $leakRow = & $rowOf $todo 'A' '工程'
    Assert-Match $todo.Rows[$leakRow]['C'] '1 件が後の工程で見つかった'
    Set-CellValues $report '要対応' @{ "E$leakRow" = '山田'; "F$leakRow" = '2026/10/30' }

    # 障害管理表を直し（流出が 2 件になり、要対応の内容が変わる）、テスト仕様書を渡さずに同じ日に作り直す
    New-ReportBook (Join-Path $dir '障害管理表.xlsx') '障害' @('障害番号', '障害内容', '起票日', '機能', '混入工程', '摘出工程', '状態') @(
        , @('B001', '一覧の件数が違う', '2026/10/01', 'SCR001', '製造', '単体テスト', '完了')
        , @('B002', '登録できない', '2026/10/04', 'SCR002', '詳細設計', '単体テスト', '対応中')
        , @('B003', '検索できない', '2026/10/05', 'SCR001', '詳細設計', '単体テスト', '対応中'))
    $text = Get-HostText { & $ReportTool @base -WarningVariable script:carryWarnings -WarningAction SilentlyContinue | Out-Null }
    Assert-Match $text '同じ時点の前の品質評価から、定性分析（原因）・対策 1 件（期限切れの指摘）と、要対応の担当・期限 1 件を引き継ぎました'
    $book = Read-XlsxSheet -Path $report
    $overdueRow = & $rowOf $book 'B' '期限切れの指摘'
    Assert-Equal "レビューの遅れ`n担当の不在|週次で確認する" ('{0}|{1}' -f $book.Rows[$overdueRow]['F'], $book.Rows[$overdueRow]['G']) '定性分析・対策を引き継ぐ'
    $todo = Read-XlsxSheet -Path $report -Index 2
    $leakRow = & $rowOf $todo 'A' '工程'
    Assert-Match $todo.Rows[$leakRow]['C'] '2 件が後の工程で見つかった'
    Assert-Equal ('山田|{0}' -f ([datetime]'2026-10-30').ToOADate()) ('{0}|{1}' -f $todo.Rows[$leakRow]['E'], $todo.Rows[$leakRow]['F']) '内容の数が変わっても、区分と対象で担当・期限（日付のセル）を引き継ぐ'

    # 今回の品質評価に無い指標（テスト仕様書を渡さなかったので消化率が無い）の記入は、知らせて前のブックを残す
    $kept = Join-Path $out '品質評価_20261020_前の版.xlsx'
    Assert-Match ($script:carryWarnings -join ' ') '引き継げませんでした: 定性分析・対策（消化率）。前のブックを残しました: .*品質評価_20261020_前の版\.xlsx'
    $old = Read-XlsxSheet -Path $kept
    Assert-Equal '環境の遅れ' $old.Rows[(& $rowOf $old 'B' '消化率')]['F']
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Assert-Equal $report (Find-QualityBook $out) '定性分析には、前の版ではなく品質評価のブックを使う'
}

Test-Case 'New-QualityReport: -Project を省くと、出力先が「品質報告」なら案件のフォルダの名前にし（-Config と同じ）、実績の CSV に別の名前の行だけがあれば知らせる' {
    $root = Join-Path (New-TestDir) '受注システム'
    New-Item -ItemType Directory -Path $root | Out-Null
    New-ReportData $root
    $history = Join-Path $root '品質報告/実績.csv'
    & $ReportTool -OutDir (Join-Path $root '品質報告') -Findings (Join-Path $root 'レビュー') -HistoryCsv $history -AsOf '2026-10-06' 3>$null 6>$null | Out-Null
    Assert-Equal @('受注システム') @(Import-Csv -LiteralPath $history | ForEach-Object { $_.'プロジェクト' } | Sort-Object -Unique)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Assert-Equal '受注システム' (Get-ProjectName (Join-Path $root 'Output')) '出力先によく付ける名前（大文字・小文字は問わない）'
    Assert-Equal '週次品質' (Get-ProjectName (Join-Path $root '週次品質'))
    # 設定ファイルで作っても同じプロジェクトになり、前回と比べる
    @('項目,値', '指摘管理表,レビュー', '実績,品質報告\実績.csv') | Set-Content -LiteralPath (Join-Path $root '品質評価の設定.csv') -Encoding UTF8
    $text = Get-HostText { & $ReportTool -Config $root -AsOf '2026-10-20' 3>$null | Out-Null }
    Assert-Match $text 'プロジェクト「受注システム」・2026/10/20。前回 2026/10/06 と比べました'

    # 以前の既定の名前（品質報告）で残した実績だけがあれば、知らせる
    $old = Join-Path $root '以前の実績.csv'
    [System.IO.File]::WriteAllText($old, ([System.IO.File]::ReadAllText($history) -replace '"受注システム"', '"品質報告"'), (New-Object System.Text.UTF8Encoding($true)))
    $null = & $ReportTool -OutDir (Join-Path $root '品質報告') -Findings (Join-Path $root 'レビュー') -HistoryCsv $old -AsOf '2026-10-27' -WarningVariable script:projectWarnings -WarningAction SilentlyContinue 6>$null
    Assert-Match ($script:projectWarnings -join ' ') 'プロジェクト「品質報告」の行だけがあり、今回のプロジェクト「受注システム」の行がありません'
}

Test-Case 'New-QualityReport: 全体のテスト密度・バグ密度で除いた規模の無い機能のケース・バグの件数を判定の考え方に書き、要対応に並べる。設計書の FP は「FP（トランザクション）」と書き、以前の「FP」の実績と比べる' {
    $dir = New-TestDir
    New-Item -ItemType Directory -Path (Join-Path $dir 'テスト') | Out-Null
    @('No,機能,確認内容,結果', '1,SCR001,確認 1,OK', '2,SCR001,確認 2,OK', '3,SCR002,確認 3,OK', '4,SCR003,確認 4,NG') | Set-Content -LiteralPath (Join-Path $dir 'テスト/cases.csv') -Encoding UTF8
    @('障害番号,障害内容,機能,起票日,状態', 'B1,一覧の誤り,SCR001,2026/10/01,完了', 'B2,登録の誤り,SCR003,2026/10/02,対応中', 'B3,共通部品の誤り,,2026/10/03,対応中') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    @('機能,KLOC', 'SCR001,1.0', 'SCR002,1.0') | Set-Content -LiteralPath (Join-Path $dir 'kloc.csv') -Encoding UTF8
    $out = Join-Path $dir '報告'
    $r = @(& $ReportTool -OutDir $out -TestPath (Join-Path $dir 'テスト') -BugList (Join-Path $dir 'bugs.csv') -SizeCsv (Join-Path $dir 'kloc.csv') -AsOf '2026-10-20' 3>$null 6>$null)
    Assert-Equal '1.5|0.5' ('{0}|{1}' -f @($r | Where-Object Name -eq 'テスト密度（全体）')[0].Value, @($r | Where-Object Name -eq 'バグ密度（全体）')[0].Value) '規模のある SCR001・SCR002（2 KLOC）のケース 3 件・バグ 1 件'
    $book = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx')
    $basis = @{}
    foreach ($row in $book.Rows.Values) { $basis[$row['B']] = $row['E'] }
    Assert-Equal 'ケース 3 件 ÷ 2 KLOC。規模の無い機能のケース 1 件は除く' $basis['テスト密度（全体）']
    Assert-Equal 'バグ 1 件 ÷ 2 KLOC。規模の無い機能のバグ 2 件は除く' $basis['バグ密度（全体）']
    $todo = Read-XlsxSheet -Path (Join-Path $out '品質評価_20261020.xlsx') -Index 2
    $todoText = @($todo.Rows.Values | ForEach-Object { '{0}|{1}|{2}' -f $_['A'], $_['B'], $_['C'] })
    Assert-True ($todoText -contains 'テスト|SCR003|規模が無いため、全体のテスト密度・バグ密度に入れていない（ケース 1 件・バグ 1 件）') ($todoText -join "`n")
    Assert-True ($todoText -contains 'テスト|（機能なし）|規模が無いため、全体のテスト密度・バグ密度に入れていない（ケース 0 件・バグ 1 件）') ($todoText -join "`n")

    # 設計書の FP（4 × EI + 5 × EO + 4 × EQ）は「FP（トランザクション）」。以前の実績の「FP」の行を前回として比べる
    $history = Join-Path $dir '実績.csv'
    @('プロジェクト,時点,区分,指標,分母,工程,値', 'A,2026/10/06,テスト,テスト密度,FP,,0.05') | Set-Content -LiteralPath $history -Encoding UTF8
    $r = @(& $ReportTool -OutDir (Join-Path $dir '報告2') -TestPath (Join-Path $dir 'テスト') -DesignPath $reportDesignDir -HistoryCsv $history -Project 'A' -AsOf '2026-10-20' 3>$null 6>$null)
    $book = Read-XlsxSheet -Path (Join-Path $dir '報告2/品質評価_20261020.xlsx')
    $density = @($book.Rows.Values | Where-Object { $_['B'] -eq 'テスト密度（全体）' })[0]
    Assert-Match $density['G'] '÷ \d+ FP（トランザクション）'
    Assert-Equal '0.05' $density['D'] '以前の分母「FP」の値を前回にする'
    Assert-Equal 'FP（トランザクション）' @(Import-Csv -LiteralPath $history | Where-Object { $_.'指標' -eq 'テスト密度' -and $_.'時点' -eq '2026/10/20' })[0].'分母'
}
