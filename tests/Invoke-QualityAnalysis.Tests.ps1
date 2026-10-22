#     Invoke-QualityAnalysis.ps1 のテスト。品質評価のブックはテストの中で作る（Export-Workbook）

$QualityAnalysis = Join-Path $script:ToolsDir 'Invoke-QualityAnalysis.ps1'

function New-QualityBook {
    param([string] $FilePath)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $metrics = @(
        , @('レビュー', '指摘密度（件/ページ）', [double]0.4, '×', '0.8 以上が ○、0.5 以上が △', '', '', '指摘.xlsx')
        , @('レビュー', '未完了の指摘', [double]0, '○', '0 件が ○', '', '', '')
        , @('テスト', 'テスト密度（件/FP）', [double]3.1, '△', '4 以上が ○、3 以上が △', '', '', '')
        , @('テスト', '未完了のバグ', [double]5, '×', '0 件が ○', '担当の割り当てが遅れた', '日次で確認する', ''))
    Export-Workbook -FilePath $FilePath -Sheets @(
        @{ Name = '品質評価'; Columns = @(@('区分', '指標', '実績', '判定', '判定の考え方', '定性分析（原因）', '対策', '詳細') | ForEach-Object { @{ Header = $_; Width = 14 } }); Rows = @($metrics | ForEach-Object { @{ Values = $_ } }) },
        @{ Name = '要対応'; Columns = @(@('区分', '対象', '内容', '対応の例', '担当', '期限') | ForEach-Object { @{ Header = $_; Width = 14 } }); Rows = @(@{ Values = @('設計書', 'DD02_画面設計書_SCR002_受注登録', '指摘密度 0.2 件/ページ（基準 0.5 以上）', '再レビュー', '', '') }) })
}

Test-Case 'Invoke-QualityAnalysis: × と △ の指標の依頼を作り、回答の定性分析・対策を品質評価のシートに書き込む' {
    $dir = New-TestDir
    $book = Join-Path $dir '品質評価_20261022.xlsx'
    New-QualityBook $book
    $text = Get-HostText { & $QualityAnalysis $dir -PromptOnly | Out-Null }
    Assert-Match $text '判定が × か △ の指標 2 件の定性分析・対策の依頼を作りました'
    $prompt = [System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))
    Assert-Match $prompt '\| 1 \| レビュー \| 指摘密度（件/ページ） \| 0\.4 \|  \| × \| 0\.8 以上が ○、0\.5 以上が △ \|'
    Assert-Match $prompt '\| 2 \| テスト \| テスト密度（件/FP） \| 3\.1 \|  \| △ \|'
    Assert-True (-not $prompt.Contains('| 3 |')) '記入済みの × の指標は依頼に含めない'
    Assert-Match $prompt '\| 設計書 \| DD02_画面設計書_SCR002_受注登録 \| 指摘密度 0\.2 件/ページ'
    Assert-Match $prompt '# そのほかの指標（参考）[\s\S]*\| レビュー \| 未完了の指摘 \| 0 \| ○ \|'

    # Markdown の表と、Copilot の表をコピーしたタブ区切りのどちらも読む
    $answer = @('| No | 定性分析（原因） | 対策 |', '| --- | --- | --- |', '| 1 | レビュー時間が短かった可能性がある（レビュー記録で確かめる） | 設計リーダーが再レビューする（今週中） |', "2`tテストケースの洗い出しが画面単位に偏っている<br>可能性がある`t異常系のケースを追加する")
    [System.IO.File]::WriteAllText((Join-Path $dir '品質評価_20261022_定性分析の回答.txt'), ($answer -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    $result = @(& $QualityAnalysis $book -ImportOnly 6>$null)
    Assert-Equal @('1|指摘密度（件/ページ）', '2|テスト密度（件/FP）') @($result | ForEach-Object { '{0}|{1}' -f $_.No, $_.Name })

    $sheet = Read-XlsxSheet -Path $book -Index 1
    Assert-Equal @('品質評価', '要対応') $sheet.Sheets
    Assert-Equal 'レビュー時間が短かった可能性がある（レビュー記録で確かめる）|設計リーダーが再レビューする（今週中）' ('{0}|{1}' -f $sheet.Rows[2]['F'], $sheet.Rows[2]['G'])
    Assert-Equal 'テストケースの洗い出しが画面単位に偏っている 可能性がある|異常系のケースを追加する' ('{0}|{1}' -f $sheet.Rows[4]['F'], $sheet.Rows[4]['G'])
    Assert-Equal '担当の割り当てが遅れた|日次で確認する' ('{0}|{1}' -f $sheet.Rows[5]['F'], $sheet.Rows[5]['G']) '記入済みの行はそのまま'
    Assert-Equal '0.4|指摘.xlsx' ('{0}|{1}' -f $sheet.Rows[2]['C'], $sheet.Rows[2]['H']) 'ほかのセルはそのまま'
    Assert-Equal 'DD02_画面設計書_SCR002_受注登録' (Read-XlsxSheet -Path $book -Index 2).Rows[2]['B']

    # 記入済みになったので、もう依頼は無い。-Force なら記入済みも含める
    $text = Get-HostText { & $QualityAnalysis $book -PromptOnly | Out-Null }
    Assert-Match $text '判定が × か △ の指標 3 件は、定性分析・対策が記入済みです'
    $text = Get-HostText { & $QualityAnalysis $book -PromptOnly -Force | Out-Null }
    Assert-Match $text '判定が × か △ の指標 3 件の定性分析・対策の依頼を作りました'
}

Test-Case 'Invoke-QualityAnalysis: 回答が無い・読めないときと、品質評価のブックが無いときに知らせる' {
    $dir = New-TestDir
    Assert-Throws { & $QualityAnalysis $dir -PromptOnly 6>$null } '品質評価のブック'
    $book = Join-Path $dir '品質評価_20261022.xlsx'
    New-QualityBook $book
    Assert-Throws { & $QualityAnalysis $book -ImportOnly 6>$null } '回答のファイルがありません'
    [System.IO.File]::WriteAllText((Join-Path $dir '品質評価_20261022_定性分析の回答.txt'), 'わかりました。', (New-Object System.Text.UTF8Encoding($true)))
    Assert-Throws { & $QualityAnalysis $book -ImportOnly 6>$null } '読み取れませんでした'
}

Test-Case 'Invoke-QualityAnalysis: Get-Help でヘルプが表示される' {
    $h = Get-Help $QualityAnalysis
    Assert-Match $h.Synopsis '定性分析（原因）と対策の下書き'
    Assert-Equal 5 @($h.parameters.parameter).Count
}
