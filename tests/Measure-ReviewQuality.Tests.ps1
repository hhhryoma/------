#     Measure-ReviewQuality.ps1 のテスト。指摘管理表はテストの中で作る（Export-Workbook）

$QualityTool = Join-Path $script:ToolsDir 'Measure-ReviewQuality.ps1'
$sampleDir = Join-Path $script:RepoRoot 'docs/design-sample'

function New-FindingBook {
    param([string] $FilePath, [string] $Sheet, [string[]] $Header, $Rows)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $FilePath -Sheets @(@{
        Name = $Sheet; Columns = @($Header | ForEach-Object { @{ Header = $_; Width = 12 } }); Rows = @($Rows | ForEach-Object { @{ Values = $_ } })
    })
}

function New-Findings {
    #     指摘管理表（指摘日・対象成果物・指摘区分・担当者など、ツールの既定とは違う見出し）と、ページ数の CSV
    param([string] $Dir)
    New-FindingBook (Join-Path $Dir '指摘管理表.xlsx') '指摘' @('No', '指摘日', '対象成果物', '指摘内容', '指摘区分', '重要度', '担当者', '対応期限', '対応状況', '完了日') @(
        @('1', '2026/10/01', 'DD02_画面設計書_SCR001_受注一覧', '桁数が違う', '誤り', '高', '山田', '2026/10/10', '完了', '2026/10/05'),
        @('2', '2026/10/01', 'DD02_画面設計書_SCR001_受注一覧', 'イベントが足りない', '漏れ', '中', '山田', '2026/10/15', '対応中', ''),
        @('3', '2026/10/02', 'DD03_テーブル定義書', '型が違う', '誤り', '低', '佐藤', '2026/10/30', '未対応', ''),
        @('4', '2026/10/02', 'DD03_テーブル定義書', '誤記', '表記', '低', '佐藤', '', '取下げ', ''),
        @('5', '2026/10/03', 'DD02_画面設計書_SCR001_受注一覧', 'メッセージの誤り', '誤り', '中', '山田', '', '対応済', '2026/10/05'))
    $csv = Join-Path $Dir 'pages.csv'
    @(
        '"ファイル","出力先","ページ数","除外ページ","結果","詳細","秒"',
        '"DD02_画面設計書_SCR001_受注一覧.pdf","","4","0","成功","","0"',
        '"sub\DD03_テーブル定義書.pdf","","10","0","成功","","0"',
        '"DD05_メッセージ一覧.pdf","","2","0","成功","","0"',
        '"broken.pdf","","","0","失敗","読めません","0"'
    ) | Set-Content -LiteralPath $csv -Encoding UTF8
    return $csv
}

Test-Case 'Measure-ReviewQuality: 設計書ごとの指摘件数・未完了・期限切れと、ページ数あたりの指摘密度を基準値と比べる' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -AsOf '2026-10-20' 6>$null)
    $got = @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Document, $_.Findings, $_.Open, $_.Overdue, $_.Density, $_.Judgement })
    Assert-Equal @(
        'DD02_画面設計書_SCR001_受注一覧|3|1|1|0.75|基準内',
        'DD03_テーブル定義書|1|1|0|0.1|少ない（レビュー不足の疑い）',
        'DD05_メッセージ一覧.pdf|0|0|0|0|少ない（レビュー不足の疑い）') $got '取下げは数えない、失敗した PDF は使わない'
    Assert-Equal '中 2・高 1' $r[0].BySeverity
    Assert-Equal '誤り 2・漏れ 1' $r[0].ByCategory
    Assert-Equal 'ページ数' $r[0].Unit
    $text = Get-HostText { & $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -AsOf '2026-10-20' | Out-Null }
    Assert-Match $text '指摘 4 件（未完了 2 件・うち期限切れ 1 件。取下げなどで数えなかったもの 1 件）/ 設計書 3 件（2026/10/20 時点）'
    Assert-Match $text '基準値から外れた設計書: 2 件'
}

Test-Case 'Measure-ReviewQuality: Invoke-DesignReview の指摘一覧を読み、-DesignPath の規模（Measure-DesignSize）を分母にする' {
    $dir = New-TestDir
    $book = Join-Path $dir 'レビュー結果.xlsx'
    New-FindingBook $book '指摘一覧' @('No', '判定', '重要度', '観点ID', '分類', '観点', '設計書', '場所', '指摘内容', '根拠', '依頼No', '対応要否', '対応状況', '対応者', '対応内容', '備考') @(
        @('1', 'NG', '高', 'V01', '整合性', '観点1', 'DD02_画面設計書_SCR002_受注登録.xlsx', '画面項目定義!C6', '桁数が違う', '', '1', '', '', '', '', ''),
        @('2', '要確認', '中', 'V02', '網羅性', '観点2', 'DD02_画面設計書_SCR002_受注登録.xlsx', 'イベント定義!A6', 'イベントの漏れ', '', '1', '', '', '', '', ''),
        @('3', 'NG', '低', 'V03', '表記', '観点3', 'DD05_メッセージ一覧.xlsx', 'メッセージ一覧!K6', '句読点', '', '2', '不要', '', '', '', ''))
    $r = @(& $QualityTool -Path $book -DesignPath $sampleDir -AsOf '2026-10-20' 6>$null)
    $scr2 = @($r | Where-Object { $_.Document -like '*SCR002*' })[0]
    Assert-Equal '2|2|19|0.105|基準内|規模' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $scr2.Findings, $scr2.Open, $scr2.Size, $scr2.Density, $scr2.Judgement, $scr2.Unit)
    $msg = @($r | Where-Object { $_.Document -like '*メッセージ*' })[0]
    Assert-Equal '1|0|16' ('{0}|{1}|{2}' -f $msg.Findings, $msg.Open, $msg.Size) '対応要否が「不要」は完了'
    Assert-Equal 7 $r.Count '指摘の無い設計書も規模があれば出す'
    Assert-Equal '少ない（レビュー不足の疑い）' @($r | Where-Object { $_.Document -like '*テーブル定義書*' })[0].Judgement
}

Test-Case 'Measure-ReviewQuality: CSV の指摘管理表も読み、-OutXlsx に設計書別・内訳・未完了・推移を保存する' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    $list = Join-Path $dir 'list.csv'
    @('設計書,指摘内容,重要度,状態,期限', 'DD05_メッセージ一覧,文言の誤り,高,未対応,2026/10/01') | Set-Content -LiteralPath $list -Encoding UTF8
    $out = Join-Path $dir 'quality.xlsx'
    $r = @(& $QualityTool -Path $dir -SizeCsv $csv -AsOf '2026-10-20' -OutXlsx $out 6>$null)
    Assert-Equal '1|1|1' @($r | Where-Object { $_.Document -like 'DD05*' } | ForEach-Object { '{0}|{1}|{2}' -f $_.Findings, $_.Open, $_.Overdue })[0]
    $docs = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('設計書別', '内訳', '未完了', '推移') $docs.Sheets
    $open = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal 4 $open.Rows.Count '見出しと未完了 3 件'
    Assert-Equal '期限切れ' $open.Rows[2]['G'] '期限切れから順'
    Assert-True (@($open.Links | Where-Object { $_.Location -eq "'指摘'!D3" }).Count -eq 1) '指摘管理表の行へのリンク'
    $trend = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal @('2026/10/01|2|0|2', '2026/10/02|1|0|3', '2026/10/03|1|0|4', '2026/10/05|0|2|2') @(2..5 | ForEach-Object { '{0}|{1}|{2}|{3}' -f $trend.Rows[$_]['A'], $trend.Rows[$_]['B'], $trend.Rows[$_]['C'], $trend.Rows[$_]['F'] })
}

Test-Case 'Measure-ReviewQuality: Get-Help でヘルプが表示される' {
    $h = Get-Help $QualityTool
    Assert-Match $h.Synopsis '指摘'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
