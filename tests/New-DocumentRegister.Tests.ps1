#     New-DocumentRegister.ps1 のテスト。成果物（Excel・Word・PDF）はテストの中で作る。詳細設計サンプルの表紙も読む

$RegisterTool = Join-Path $script:ToolsDir 'New-DocumentRegister.ps1'

function New-CoverBook {
    #     表紙と改訂履歴のブック。Cover: @(@(見出し, 値)...)（1 行目も 2 列にする。Export-Workbook は 1 行目の列の数だけ書く）。
    #     History: @(@(版数, 改訂日, 内容, 改訂者)...)
    param([string] $Path, $Cover, $History)
    $sheets = [ordered]@{ '表紙' = $Cover }
    if ($History) { $sheets['改訂履歴'] = @(, @('版数', '改訂日', '改訂内容', '改訂者')) + @($History) }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    New-SheetsBook $Path $sheets
}

function New-Deliverables {
    #     成果物のフォルダ（古い版・版数の食い違い・Word・PDF・サブフォルダ・一時ファイル）
    param([string] $Dir)
    New-CoverBook (Join-Path $Dir '設計書_v1.0.xlsx') @(@('受注管理システム', ''), @('文書名', '設計書'), @('文書番号', 'X-01'), @('版数', '1.0'), @('最終更新日', '2026/10/01')) @(, @('1.0', '2026/10/01', '新規作成', '佐藤'))
    New-CoverBook (Join-Path $Dir '設計書_v1.1.xlsx') @(@('受注管理システム', ''), @('文書名', '設計書'), @('文書番号', 'X-01'), @('版数', '1.1'), @('最終更新日', '2026/10/15'), @('作成者', '佐藤')) @(@('1.0', '2026/10/01', '新規作成', '佐藤'), @('1.1', '2026/10/15', '画面の追加', '山田'))
    # 表紙の版数と更新日を直し忘れた
    New-CoverBook (Join-Path $Dir '基本設計書_v1.2.xlsx') @(@('基本設計書', ''), @('文書番号', 'B-01'), @('版数', '1.1'), @('最終更新日', '2026/10/10')) @(@('1.1', '2026/10/10', '見直し', '佐藤'), @('1.2', '2026/10/20', '指摘の反映', '佐藤'))
    # 文書名の見出しの無い表紙（「…書」で終わる文字を文書名にする）
    New-CoverBook (Join-Path $Dir '02_テスト/テスト計画書.xlsx') @(@('受注管理システム', ''), @('テスト計画書（結合テスト）', ''), @('文書番号', 'T-01'), @('版数', '第1.0版'), @('最終更新日', '2026年10月5日')) $null
    New-WordDoc (Join-Path $Dir '議事録_v1.0.docx') @(
        , @(@('文書名', '議事録'), @('文書番号', 'M-01'), @('版数', '1.0'))
        '本文です'
        , @(@('版数', '改訂日', '改訂内容', '改訂者'), @('1.0', '2026/10/03', '新規作成', '鈴木')))
    $pdf = "%PDF-1.4`n1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj`n2 0 obj << /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >> endobj`n3 0 obj << /Type /Page /Parent 2 0 R >> endobj`n4 0 obj << /Type/Page /Parent 2 0 R >> endobj`ntrailer << /Root 1 0 R >>`n%%EOF`n"
    [System.IO.File]::WriteAllText((Join-Path $Dir '報告書.pdf'), $pdf, (New-Object System.Text.UTF8Encoding($false)))
    [System.IO.File]::WriteAllText((Join-Path $Dir '~$設計書_v1.1.xlsx'), 'lock', (New-Object System.Text.UTF8Encoding($false)))
}

function Get-RowKeys {
    param($Docs)
    return @($Docs | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $_.Folder, $_.File, $_.DocNo, $_.Title, $_.Version, $_.Updated, $_.Note })
}

Test-Case 'New-DocumentRegister: 表紙と改訂履歴から文書番号・文書名・版数・更新日を読み、古い版と版数・更新日の食い違いを知らせる' {
    $dir = New-TestDir
    New-Deliverables $dir
    $docs = @(& $RegisterTool -Path $dir -AsOf '2026-10-24' 6>$null)
    Assert-Equal @(
        '02_テスト|テスト計画書.xlsx|T-01|テスト計画書（結合テスト）|第1.0版|2026/10/05|',
        '|基本設計書_v1.2.xlsx|B-01|基本設計書|1.1|2026/10/10|版数が違う（ファイル名 1.2・表紙 1.1・改訂履歴 1.2）・更新日が違う（表紙 2026/10/10・改訂履歴 2026/10/20）',
        '|報告書.pdf||報告書|||',
        '|設計書_v1.1.xlsx|X-01|設計書|1.1|2026/10/15|',
        '|議事録_v1.0.docx|M-01|議事録|1.0||') (Get-RowKeys $docs) '最新の版だけ。一時ファイル（~$）は載せない'
    $design = @($docs | Where-Object { $_.File -eq '設計書_v1.1.xlsx' })[0]
    Assert-Equal '1.1|2026/10/15|画面の追加|山田|佐藤' ('{0}|{1}|{2}|{3}|{4}' -f $design.History.Version, $design.History.Date, $design.History.Content, $design.History.Author, $design.Author) '改訂履歴の最後の行'
    $minutes = @($docs | Where-Object { $_.File -eq '議事録_v1.0.docx' })[0]
    Assert-Equal 'Word|1.0|2026/10/03|鈴木' ('{0}|{1}|{2}|{3}' -f $minutes.Kind, $minutes.History.Version, $minutes.History.Date, $minutes.History.Author) 'Word の表の表紙と改訂履歴'
    Assert-Equal 2 @($docs | Where-Object { $_.File -eq '報告書.pdf' })[0].Pages 'PDF のページ数（/Pages は数えない）'

    $text = Get-HostText { & $RegisterTool -Path $dir -AsOf '2026-10-24' | Out-Null }
    Assert-Match $text '成果物 5 件（Excel 3・Word 1・PDF 1）。古い版 1 件'
    Assert-Match $text '版数・更新日などの確認が要るもの 1 件: 基本設計書_v1.2\.xlsx'

    # -AllVersions: 古い版も載せる
    $all = @(& $RegisterTool -Path $dir -AllVersions 6>$null)
    Assert-Equal '古い版（最新: 設計書_v1.1.xlsx）' @($all | Where-Object { $_.File -eq '設計書_v1.0.xlsx' })[0].Note
}

Test-Case 'New-DocumentRegister: -Planned の成果物の予定と、文書番号か名前で突き合わせ、未提出と予定に無い成果物を知らせる' {
    $dir = New-TestDir
    $docsDir = Join-Path $dir '納品物'
    New-Item -ItemType Directory -Path $docsDir | Out-Null
    New-Deliverables $docsDir
    $planned = Join-Path $dir '成果物の予定.csv'
    @('成果物名,文書番号,予定日,担当', '設計書,X-01,2026/10/15,山田', '基本設計書,,2026/10/20,佐藤', '議事録.docx,,,', 'テスト計画書,,2026/10/05,', '運用手順書,,2026/10/01,', '移行計画書,,2026/12/01,') |
        Set-Content -LiteralPath $planned -Encoding UTF8
    $out = Join-Path $dir '成果物一覧.xlsx'
    $csv = Join-Path $dir '成果物一覧.csv'
    $text = Get-HostText { & $RegisterTool -Path $docsDir -Planned $planned -AsOf '2026-10-24' -OutXlsx $out -OutCsv $csv | Out-Null }
    Assert-Match $text '予定 6 件: 提出済み 4・未提出 2（予定日を過ぎたもの 1）。予定に無い成果物 1 件'
    Assert-Match $text '未提出（予定日を過ぎています）: 運用手順書'

    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('成果物一覧', '予定との突き合わせ', '古い版', '読み取り結果') $book.Sheets
    Assert-Equal 'No|フォルダ|ファイル|種類|文書番号|文書名|版数' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal '予定に無い' $book.Rows[4]['O'] '報告書.pdf'
    Assert-True (@($book.Links | Where-Object { $_.Target -like '*基本設計書_v1.2.xlsx' }).Count -eq 1) 'ファイルへのリンク'
    $plan = Read-XlsxSheet -Path $out -Index 2
    $rows = @(2..($plan.Rows.Count) | ForEach-Object { $r = $plan.Rows[$_]; '{0}|{1}|{2}|{3}' -f $r['A'], $r['C'], $r['E'], ([string]$r['F']).Replace('\', '/') })
    Assert-Equal @(
        '設計書|2026/10/15|提出済み|設計書_v1.1.xlsx',
        '基本設計書|2026/10/20|提出済み|基本設計書_v1.2.xlsx',
        '議事録.docx||提出済み|議事録_v1.0.docx',
        'テスト計画書|2026/10/05|提出済み|02_テスト/テスト計画書.xlsx',
        '運用手順書|2026/10/01|未提出（予定日を過ぎています）|',
        '移行計画書|2026/12/01|未提出|',
        '||予定に無い|報告書.pdf') $rows '文書番号が無ければ名前（文書名か、版数を除いたファイル名）で突き合わせる'
    $old = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal '設計書_v1.0.xlsx|1.0|設計書_v1.1.xlsx' ('{0}|{1}|{2}' -f $old.Rows[2]['B'], $old.Rows[2]['C'], $old.Rows[2]['E'])
    $bytes = [System.IO.File]::ReadAllBytes($csv)
    Assert-Equal '239|187|191' (@($bytes[0..2]) -join '|') 'CSV は BOM 付き'
    Assert-Match ([System.IO.File]::ReadAllText($csv)) '^﻿?No,フォルダ,ファイル,種類,文書番号,文書名,版数,作成日,更新日,作成者,最新の改訂,ページ,サイズ（KB）,ファイルの更新日時,確認'
}

Test-Case 'New-DocumentRegister: 詳細設計サンプルの表紙（見出しの無い文書名・後ろの（ ））と改訂履歴を読む' {
    $docs = @(& $RegisterTool -Path (Join-Path $script:RepoRoot 'docs/design-sample') -NoRecurse 6>$null | Where-Object { $_.Kind -eq 'Excel' })
    Assert-Equal 8 $docs.Count
    $table = @($docs | Where-Object { $_.File -eq 'DD03_テーブル定義書.xlsx' })[0]
    Assert-Equal 'ODR-DD-03|テーブル定義書|1.1|2026/10/12|山田 太郎' ('{0}|{1}|{2}|{3}|{4}' -f $table.DocNo, $table.Title, $table.Version, $table.Updated, $table.Author)
    Assert-Equal '1.1|2026/10/12' ('{0}|{1}' -f $table.History.Version, $table.History.Date)
    Assert-Equal '画面設計書（SCR001 受注一覧）' @($docs | Where-Object { $_.File -like 'DD02*SCR001*' })[0].Title
}

Test-Case 'New-DocumentRegister: Get-Help でヘルプが表示される' {
    $h = Get-Help $RegisterTool
    Assert-Match $h.Synopsis '成果物'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
