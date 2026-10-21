#     Measure-DefectRemoval.ps1 のテスト。障害管理表・指摘管理表はテストの中で作る

$DefectTool = Join-Path $script:ToolsDir 'Measure-DefectRemoval.ps1'

function New-DefectData {
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $rows = @(
        , @('B001', '一覧の件数が違う', '製造', '単体テスト', 'コーディングミス', '中', '完了')
        , @('B002', '桁あふれ', '製造', '単体テスト', 'コーディングミス', '高', '完了')
        , @('B003', '初期値が違う', '製造', 'UT', 'コーディングミス', '低', '完了')
        , @('B004', '排他制御の記載漏れ', '詳細設計', '単体テスト', '設計誤り', '高', '完了')
        , @('B005', 'チェックの順番が仕様と違う', '詳細設計', 'UT', '設計誤り', '中', '対応中')
        , @('B006', 'IF の項目の抜け', '基本設計', 'IT', '設計漏れ', '高', '完了')
        , @('B007', '操作ミス', '製造', '単体テスト', '', '低', '取下げ')
        , @('B008', '原因不明', '謎の工程', '単体テスト', 'コーディングミス', '低', '完了'))
    Export-Workbook -FilePath (Join-Path $Dir '障害管理表.xlsx') -Sheets @(@{
            Name = '障害'; Columns = @(@('障害番号', '障害内容', '混入工程', '摘出工程', '原因分類', '重要度', '状態') | ForEach-Object { @{ Header = $_; Width = 12 } })
            Rows = @($rows | ForEach-Object { @{ Values = $_ } })
        })
    $findings = Join-Path $Dir '指摘.csv'
    @('指摘内容,工程,分類,状態',
        '桁数が無い,詳細設計レビュー,誤り,完了',
        '型が違う,詳細設計レビュー,誤り,完了',
        'メッセージIDの誤り,詳細設計レビュー,誤り,完了',
        'イベントの抜け,詳細設計レビュー,漏れ,完了',
        '画面遷移の抜け,基本設計レビュー,漏れ,完了',
        '用語の揺れ,基本設計レビュー,誤り,完了') | Set-Content -LiteralPath $findings -Encoding UTF8
    return $findings
}

Test-Case 'Measure-DefectRemoval: 混入工程 × 摘出工程から、工程ごとの摘出率と流出を出す（別名の工程・レビューの指摘も読む）' {
    $dir = New-TestDir
    $findings = New-DefectData $dir
    $out = Join-Path $dir 'defect.xlsx'
    $warnings = $null
    $r = @(& $DefectTool -BugList (Join-Path $dir '障害管理表.xlsx') -Findings $findings -OutXlsx $out -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @(
        '基本設計|3|2|1|0.667|2|0',
        '詳細設計|6|4|2|0.667|4|0',
        '製造|3|3|0|1|0|0',
        '単体テスト|0|0|0||6|5',
        '結合テスト|0|0|0||1|1') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $_.Phase, $_.Injected, $_.FoundInTime, $_.Leaked, $_.RemovalRate, $_.Detected, $_.FromEarlier })
    Assert-True (@($warnings | Where-Object { "$_" -match '工程として読めない値があります.*謎の工程' }).Count -eq 1) '読めない工程を知らせる'

    $text = Get-HostText { & $DefectTool -BugList (Join-Path $dir '障害管理表.xlsx') -Findings $findings 3>$null | Out-Null }
    Assert-Match $text '欠陥 13 件（バグ 7・レビューの指摘 6。取下げなどで数えなかったもの 1 件）'
    Assert-Match $text '摘出すべき工程までに見つかった欠陥: 9 / 12 件（75\.0%）。流出 3 件'
    Assert-Match $text '設計で入り、テストで見つかった欠陥: 3 件（詳細設計 2・基本設計 1）'
    Assert-Match $text '原因の上位: コーディングミス 4（31%）・誤り 4（31%）・漏れ 2（15%）'

    $book = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('工程別', '工程マトリクス', '原因', '流出') $book.Sheets
    # 読めない混入工程（謎の工程）は、行の最後の「（不明）」に数える
    Assert-Equal '混入工程 ＼ 摘出工程|基本設計|詳細設計|単体テスト|結合テスト|合計' ((@('A', 'B', 'C', 'D', 'E', 'F') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal '詳細設計|0|4|2|0|6' ((@('A', 'B', 'C', 'D', 'E', 'F') | ForEach-Object { $book.Rows[3][$_] }) -join '|')
    Assert-Equal '（不明）|0|0|1|0|1' ((@('A', 'B', 'C', 'D', 'E', 'F') | ForEach-Object { $book.Rows[5][$_] }) -join '|')
    $leak = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal 4 $leak.Rows.Count '見出しと流出 3 件'
    Assert-Equal '基本設計|結合テスト' ('{0}|{1}' -f $leak.Rows[2]['B'], $leak.Rows[2]['C'])
    Assert-True (@($leak.Links | Where-Object { $_.Location -like "'障害'!B*" }).Count -eq 3) '障害管理表の行へのリンク'
}

Test-Case 'Measure-DefectRemoval: Get-Help でヘルプが表示される' {
    $h = Get-Help $DefectTool
    Assert-Match $h.Synopsis '混入工程'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
