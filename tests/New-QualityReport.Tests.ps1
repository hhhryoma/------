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
}

Test-Case 'New-QualityReport: 各ツールを実行し、指標ごとの実績と判定（○△×）・要対応を 1 冊にまとめる' {
    $dir = New-TestDir
    New-ReportData $dir
    $out = Join-Path $dir '報告'
    $r = @(& $ReportTool -OutDir $out -DesignPath $reportDesignDir -Findings (Join-Path $dir 'レビュー') -TestPath (Join-Path $dir 'テスト') -BugList (Join-Path $dir '障害管理表.xlsx') -Requirement (Join-Path $dir '要件一覧.csv') -AsOf '2026-10-20' 3>$null 6>$null)
    $got = @{}
    foreach ($m in $r) { $got[$m.Name] = '{0}|{1}' -f $m.Value, $m.Judgement }
    Assert-Equal '3|' $got['指摘件数']
    Assert-Equal '2|△' $got['未完了の指摘']
    Assert-Equal '1|×' $got['期限切れの指摘'] '期限切れが残っていれば ×'
    Assert-Equal '1|○' $got['消化率']
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
    foreach ($name in @('レビュー.xlsx', 'テスト.xlsx', 'テスト仕様書のチェック.xlsx', '工程別の欠陥.xlsx', '網羅.xlsx', '設計書の整合性.xlsx')) {
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
