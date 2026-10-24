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
    Assert-Equal @('設計書別', '内訳', '未完了', '推移', '読み取り結果') $docs.Sheets
    $open = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal 4 $open.Rows.Count '見出しと未完了 3 件'
    Assert-Equal '期限切れ' $open.Rows[2]['G'] '期限切れから順'
    Assert-True (@($open.Links | Where-Object { $_.Location -eq "'指摘'!D3" }).Count -eq 1) '指摘管理表の行へのリンク'
    $trend = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal @('2026/10/01|2|0|2', '2026/10/02|1|0|3', '2026/10/03|1|0|4', '2026/10/05|0|2|2') @(2..5 | ForEach-Object { '{0}|{1}|{2}|{3}' -f $trend.Rows[$_]['A'], $trend.Rows[$_]['B'], $trend.Rows[$_]['C'], $trend.Rows[$_]['F'] })
}

function New-ReviewLogs {
    #     レビュー記録（表の形の Excel・記録票の形の Excel・CSV）
    param([string] $Dir)
    $logDir = Join-Path $Dir 'レビュー記録'
    New-Item -ItemType Directory -Path $logDir | Out-Null
    New-FindingBook (Join-Path $logDir 'レビュー記録一覧.xlsx') '記録' @('No', 'レビュー日', '対象成果物', 'レビュー時間（分）', '参加者') @(
        @('1', '2026/10/01', 'DD02_画面設計書_SCR001_受注一覧', '60', '山田、佐藤'),
        @('2', '2026/10/02', 'DD03_テーブル定義書.xlsx', '60', '鈴木'))
    # 記録票: 見出しの右に値。参加者は右に 1 人ずつ
    New-FindingBook (Join-Path $logDir 'DD05記録票.xlsx') '記録票' @('レビュー記録票', '', '') @(
        @('対象成果物', 'DD05_メッセージ一覧', ''),
        @('レビュー工数', '3人時', ''),
        @('参加者', '山田', '佐藤'))
    # 時刻の書式のセルは「01:30:00」の形で読まれる
    @('設計書,レビュー時間,人数', 'DD02_画面設計書_SCR001_受注一覧,10:00～10:30,2', 'DD09_未知の設計書,01:30:00,1') | Set-Content -LiteralPath (Join-Path $logDir 'log.csv') -Encoding UTF8
    return $logDir
}

Test-Case 'Measure-ReviewQuality: -ReviewLog でレビュー工数密度を出し、指摘密度との 2 軸でゾーンに分ける' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    $logDir = New-ReviewLogs $dir
    $out = Join-Path $dir 'quality.xlsx'
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -AsOf '2026-10-20' -OutXlsx $out 6>$null)
    $got = @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Document, $_.Density, $_.Effort, $_.EffortDensity, $_.EffortJudgement, $_.Zone })
    Assert-Equal @(
        'DD02_画面設計書_SCR001_受注一覧|0.75|3|0.75|基準内|良好',
        'DD03_テーブル定義書|0.1|1|0.1|少ない|レビュー不足',
        'DD05_メッセージ一覧.pdf|0|3|1.5|多い|良好',
        'DD09_未知の設計書||1.5|||') $got '分の時間 × 参加者・時刻の範囲 × 人数・記録票の人時'
    Assert-Equal 2 @($r | Where-Object { $_.Document -like 'DD02*' })[0].Reviews
    $dd03 = @($r | Where-Object { $_.Document -like 'DD03*' })[0]
    Assert-True $dd03.NeedsAction
    Assert-Equal '観点を決めて追加のレビューをする' $dd03.ZoneAction

    $text = Get-HostText { & $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -AsOf '2026-10-20' | Out-Null }
    Assert-Match $text 'レビュー工数 8\.5 人時（レビュー 5 回・設計書 4 件）'
    Assert-Match $text 'ゾーン: 良好 2・レビュー不足 1'
    Assert-Match $text '要対応: DD03_テーブル定義書（レビュー不足）観点を決めて追加のレビューをする'

    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('設計書別', '内訳', '未完了', '推移', 'ゾーン分析', '読み取り結果') $book.Sheets
    Assert-Equal 'レビュー不足' @($book.Rows.Values | Where-Object { $_['A'] -eq 'DD03_テーブル定義書' })[0]['N']
    $charts = @(Read-XlsxCharts $out)
    Assert-Equal @('レビュー工数密度 × 指摘密度', '指摘の起票・完了（累計）と未完了') @($charts | ForEach-Object { $_.Title })
    Assert-Equal 3 @([regex]::Matches($charts[0].Xml, '<c:ser>')).Count '工数密度と指摘密度の両方がある設計書（DD09 は規模なし）'
    $zone = Read-XlsxSheet -Path $out -Index 5
    Assert-Equal '工数密度が多い' $zone.Rows[2]['A']
    Assert-Match $zone.Rows[2]['B'] '^【良好】\nDD05_メッセージ一覧\.pdf（指摘 0 / 工数 1\.5）$'
    Assert-Match $zone.Rows[3]['C'] '^【良好】\nDD02_画面設計書_SCR001_受注一覧（指摘 0\.75 / 工数 0\.75）$'
    Assert-Match $zone.Rows[4]['B'] '^【レビュー不足】\nDD03_テーブル定義書（指摘 0\.1 / 工数 0\.1）$'
    Assert-Equal '【品質に問題】' $zone.Rows[4]['D']

    # 結果のブックを同じフォルダに置いて読み直しても、レビュー記録として数えない
    Copy-Item -LiteralPath $out -Destination (Join-Path $logDir 'old-result.xlsx')
    $again = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -AsOf '2026-10-20' 6>$null)
    Assert-Equal 3 @($again | Where-Object { $_.Document -like 'DD02*' })[0].Effort
}

Test-Case 'Measure-ReviewQuality: レビュー記録も指摘も無い設計書を未レビューとし、レビュー実施率を出す' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    $logDir = New-ReviewLogs $dir
    Add-Content -LiteralPath $csv -Value '"DD04_コード定義書.pdf","","3","0","成功","","0"' -Encoding UTF8
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -AsOf '2026-10-20' 6>$null)
    $dd04 = @($r | Where-Object { $_.Document -like 'DD04*' })[0]
    Assert-Equal '未レビュー|False|' ('{0}|{1}|{2}' -f $dd04.Judgement, $dd04.Reviewed, $dd04.Zone)
    # 指摘が 0 件でもレビュー記録があればレビューした（DD05）
    Assert-Equal $true @($r | Where-Object { $_.Document -like 'DD05*' })[0].Reviewed
    $text = Get-HostText { & $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -AsOf '2026-10-20' 3>$null | Out-Null }
    Assert-Match $text 'レビュー実施率 80\.0%（設計書 5 件のうち 4 件）'
    Assert-Match $text '未レビュー: DD04_コード定義書\.pdf'
    Assert-NotMatch $text 'レビューの記録が無い設計書'

    # レビュー記録が無ければ、指摘 0 件の設計書を未レビューとは決めない
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -AsOf '2026-10-20' 6>$null)
    $dd04 = @($r | Where-Object { $_.Document -like 'DD04*' })[0]
    Assert-Equal '少ない（レビュー不足の疑い）|' ('{0}|{1}' -f $dd04.Judgement, $dd04.Reviewed)
    $text = Get-HostText { & $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -AsOf '2026-10-20' 3>$null | Out-Null }
    Assert-Match $text '指摘 0 件の設計書: 2 件（-ReviewLog でレビュー記録を渡すと、未レビューと区別できます）'
}

Test-Case 'Measure-ReviewQuality: 指標の列が無い基準値（以前の形）は指摘密度の基準値とし、ゾーンは決めずに知らせる' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    $logDir = New-ReviewLogs $dir
    $limits = Join-Path $dir 'limits.csv'
    @('分母,下限,上限', 'ページ数,0.5,2.0') | Set-Content -LiteralPath $limits -Encoding UTF8
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -Limits $limits -AsOf '2026-10-20' 6>$null 3>$null)
    $dd02 = @($r | Where-Object { $_.Document -like 'DD02*' })[0]
    Assert-Equal '基準内|0.75||' ('{0}|{1}|{2}|{3}' -f $dd02.Judgement, $dd02.EffortDensity, $dd02.EffortJudgement, $dd02.Zone)
    $warn = (& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -ReviewLog $logDir -Limits $limits -AsOf '2026-10-20' 6>$null 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message }) -join ' '
    Assert-Match $warn 'レビュー工数密度の行がありません'
}

Test-Case 'Measure-ReviewQuality: 指摘の数が同じ工程は、先に出てきた工程の基準値を使う（PowerShell の版によらない）' {
    $dir = New-TestDir
    $csv = New-Findings $dir
    # 先に出てくる「詳細設計」は、名前の順では「基本設計」のあと
    New-FindingBook (Join-Path $dir '工程別.xlsx') '指摘' @('No', '対象成果物', '指摘内容', '工程', '対応状況') @(
        @('1', 'DD02_画面設計書_SCR001_受注一覧', '桁数が違う', '詳細設計', '完了'),
        @('2', 'DD02_画面設計書_SCR001_受注一覧', 'イベントが足りない', '基本設計', '完了'))
    $limits = Join-Path $dir 'limits.csv'
    @('指標,分母,工程,下限,上限', '指摘密度,ページ数,詳細設計,0.1,0.3', '指摘密度,ページ数,基本設計,1.0,2.0') | Set-Content -LiteralPath $limits -Encoding UTF8
    $r = @(& $QualityTool -Path (Join-Path $dir '工程別.xlsx') -SizeCsv $csv -Limits $limits -AsOf '2026-10-20' 6>$null 3>$null)
    $dd02 = @($r | Where-Object { $_.Document -like 'DD02*' })[0]
    Assert-Equal '詳細設計|0.5' ('{0}|{1}' -f $dd02.Process, $dd02.Density)
    Assert-Match $dd02.Judgement '^多い' '詳細設計の基準値（0.1〜0.3）と比べる'
}

Test-Case 'Measure-ReviewQuality: 状態の「未完了」「未済」「完了予定」「差し戻し」は、完了・済を含んでも未完了に数える' {
    $dir = New-TestDir
    New-FindingBook (Join-Path $dir '指摘管理表.xlsx') '指摘' @('No', '指摘日', '対象成果物', '指摘内容', '対応状況') @(
        @('1', '2026/10/01', 'DD03_テーブル定義書', '型が違う', '完了'),
        @('2', '2026/10/01', 'DD03_テーブル定義書', '桁数が違う', '未完了'),
        @('3', '2026/10/02', 'DD03_テーブル定義書', '誤記', '未済'),
        @('4', '2026/10/02', 'DD03_テーブル定義書', '名前の揺れ', '完了予定'),
        @('5', '2026/10/03', 'DD03_テーブル定義書', '説明が足りない', '対応済（差し戻し）'),
        @('6', '2026/10/03', 'DD03_テーブル定義書', '索引が足りない', '済'))
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -AsOf '2026-10-20' 6>$null 3>$null)
    Assert-Equal 'DD03_テーブル定義書|6|4' ('{0}|{1}|{2}' -f $r[0].Document, $r[0].Findings, $r[0].Open)
}

Test-Case 'Measure-ReviewQuality: 設計書の名前のフォルダ（D:\納品物\…・sub\…）を除いて同じ設計書にまとめ、「外部I/F設計書」と「内部I/F設計書」は分ける' {
    $dir = New-TestDir
    New-FindingBook (Join-Path $dir '指摘管理表.xlsx') '指摘' @('No', '指摘日', '対象成果物', '指摘内容', '対応状況') @(
        @('1', '2026/10/01', 'DD03_テーブル定義書', '型が違う', '対応中'),
        @('2', '2026/10/02', 'D:\納品物\DD03_テーブル定義書.xlsx', '桁数が違う', '対応中'),
        @('3', '2026/10/02', '外部I/F設計書', '項目が足りない', '対応中'),
        @('4', '2026/10/03', '内部I/F設計書', '誤記', '対応中'))
    $csv = Join-Path $dir 'pages.csv'
    # DD03_テーブル定義書 を含む名前が 2 つあるので、フォルダを除かないと規模と突き合わせられない
    @('"ファイル","出力先","ページ数","除外ページ","結果","詳細","秒"', '"sub\DD03_テーブル定義書.pdf","","10","0","成功","","0"', '"DD03_テーブル定義書_別紙.pdf","","5","0","成功","","0"') | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = @(& $QualityTool -Path (Join-Path $dir '指摘管理表.xlsx') -SizeCsv $csv -AsOf '2026-10-20' 6>$null 3>$null)
    Assert-Equal @('DD03_テーブル定義書|2|10|0.2', '内部I/F設計書|1||', '外部I/F設計書|1||', 'DD03_テーブル定義書_別紙.pdf|0|5|0') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Document, $_.Findings, $_.Size, $_.Density })
}

Test-Case 'Measure-ReviewQuality: Get-Help でヘルプが表示される' {
    $h = Get-Help $QualityTool
    Assert-Match $h.Synopsis '指摘'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
