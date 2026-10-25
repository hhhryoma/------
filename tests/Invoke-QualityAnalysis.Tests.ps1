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

Test-Case 'Invoke-QualityAnalysis: 同じフォルダの前回の品質評価の定性分析・対策を、今回の実績と並べて依頼に入れる' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $columns = @(@('区分', '指標', '実績', '判定', '判定の考え方', '定性分析（原因）', '対策') | ForEach-Object { @{ Header = $_; Width = 14 } })
    foreach ($old in @(@('20260924', '古い原因', '古い対策'), @('20261008', 'レビュー時間が短い', '再レビューする'))) {
        Export-Workbook -FilePath (Join-Path $dir ('品質評価_{0}.xlsx' -f $old[0])) -Sheets @(@{ Name = '品質評価'; Columns = $columns; Rows = @(
                    @{ Values = @('レビュー', '指摘密度（件/ページ）', [double]0.3, '×', '', $old[1], $old[2]) },
                    @{ Values = @('テスト', '消化率', [double]0.5, '×', '', 'テスト環境の遅れ', '環境を先に作る') },
                    @{ Values = @('テスト', 'テスト密度（件/FP）', [double]3.5, '△', '', '', '') }) })
    }
    $book = Join-Path $dir '品質評価_20261022.xlsx'
    New-QualityBook $book
    $text = Get-HostText { & $QualityAnalysis $book -PromptOnly | Out-Null }
    Assert-Match $text '前回の品質評価の定性分析・対策 2 件も入れました: .*品質評価_20261008\.xlsx'
    $prompt = [System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))
    Assert-Match $prompt '# 前回の定性分析と対策（品質評価_20261008\.xlsx）'
    Assert-Match $prompt '\| レビュー \| 指摘密度（件/ページ） \| 0\.3 \| × \| 0\.4 \| × \| レビュー時間が短い \| 再レビューする \|'
    Assert-Match $prompt '\| テスト \| 消化率 \| 0\.5 \| × \| （今回なし） \|  \| テスト環境の遅れ \| 環境を先に作る \|'
    Assert-True (-not $prompt.Contains('古い対策')) '一番近い前回だけを使う'
    Assert-Match $prompt '対策が効いたかを定性分析に書いてください'

    # -Previous で前回を指定できる
    $null = & $QualityAnalysis $book -PromptOnly -Previous (Join-Path $dir '品質評価_20260924.xlsx') 6>$null
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))) '古い対策'
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

Test-Case 'Invoke-QualityAnalysis: 依頼の表のセルの、セル内の改行は空白にする' {
    $dir = New-TestDir
    $book = Join-Path $dir '品質評価_20261022.xlsx'
    New-QualityBook $book
    # 判定の考え方のセルの中で改行する（Alt+Enter）
    Edit-ZipText $book 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('0.8 以上が ○、0.5 以上が △', "0.8 以上が ○、`n0.5 以上が △") }
    $null = Get-HostText { & $QualityAnalysis $dir -PromptOnly | Out-Null }
    $prompt = [System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))
    Assert-Match $prompt '\| × \| 0\.8 以上が ○、 0\.5 以上が △ \|'
    Assert-True (-not $prompt.Contains('↵')) '↵ を書かない'
}

Test-Case 'Invoke-QualityAnalysis: -ImportOnly は、依頼を作ったときの No の指標に回答を書く（記入済みの行は -Force のときだけ）' {
    $dir = New-TestDir
    $book = Join-Path $dir '品質評価_20261022.xlsx'
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    # 記入済みの × の指標が、空欄の × の指標より上にある
    $columns = @(@('区分', '指標', '実績', '判定', '判定の考え方', '定性分析（原因）', '対策') | ForEach-Object { @{ Header = $_; Width = 14 } })
    Export-Workbook -FilePath $book -Sheets @(@{ Name = '品質評価'; Columns = $columns; Rows = @(
                @{ Values = @('テスト', '未完了のバグ', [double]5, '×', '0 件が ○', '担当の割り当てが遅れた', '日次で確認する') },
                @{ Values = @('レビュー', '指摘密度（件/ページ）', [double]0.4, '×', '0.8 以上が ○', '', '') }) })
    # -Force で記入済みの指標も依頼に入れる（No 1 が未完了のバグ、No 2 が指摘密度）
    $null = & $QualityAnalysis $book -PromptOnly -Force 6>$null
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))) '\| 1 \| テスト \| 未完了のバグ \|[\s\S]*\| 2 \| レビュー \| 指摘密度'
    $answer = @('| No | 定性分析（原因） | 対策 |', '| --- | --- | --- |', '| 1 | 修正の担当が 1 人に集中した | 担当を 2 人にする |', '| 2 | レビュー時間が短かった | 再レビューする |')
    [System.IO.File]::WriteAllText((Join-Path $dir '品質評価_20261022_定性分析の回答.txt'), ($answer -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))

    # -Force を付けずに取り込むと、No 2 の回答は指摘密度に書き、記入済みの未完了のバグは書き換えずに知らせる
    $text = Get-HostText { & $QualityAnalysis $book -ImportOnly | Out-Null }
    $sheet = Read-XlsxSheet -Path $book -Index 1
    Assert-Equal 'レビュー時間が短かった|再レビューする' ('{0}|{1}' -f $sheet.Rows[3]['F'], $sheet.Rows[3]['G']) 'No 2 の回答は指摘密度の行に書く'
    Assert-Equal '担当の割り当てが遅れた|日次で確認する' ('{0}|{1}' -f $sheet.Rows[2]['F'], $sheet.Rows[2]['G']) '記入済みの行は書き換えない'
    Assert-Match $text '記入済みの指標には書き込んでいません（書き直すときは -Force）: 1\. 未完了のバグ'

    # -Force を付けると、記入済みの行にも書く
    $result = @(& $QualityAnalysis $book -ImportOnly -Force 6>$null)
    Assert-Equal @('1|未完了のバグ', '2|指摘密度（件/ページ）') @($result | ForEach-Object { '{0}|{1}' -f $_.No, $_.Name })
    $sheet = Read-XlsxSheet -Path $book -Index 1
    Assert-Equal '修正の担当が 1 人に集中した|担当を 2 人にする' ('{0}|{1}' -f $sheet.Rows[2]['F'], $sheet.Rows[2]['G'])
    Assert-Equal 'レビュー時間が短かった|再レビューする' ('{0}|{1}' -f $sheet.Rows[3]['F'], $sheet.Rows[3]['G'])
}

Test-Case 'Invoke-QualityAnalysis: 要対応の行がいくつあっても、すべて依頼に入れる。フォルダでは、名前の日付が一番新しい品質評価のブックを使う' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $metricColumns = @(@('区分', '指標', '実績', '判定', '判定の考え方', '定性分析（原因）', '対策') | ForEach-Object { @{ Header = $_; Width = 14 } })
    $actionColumns = @(@('区分', '対象', '内容', '対応の例', '担当', '期限') | ForEach-Object { @{ Header = $_; Width = 14 } })
    foreach ($b in @(@('20261022', [double]0.8), @('20261015', [double]0.5))) {
        Export-Workbook -FilePath (Join-Path $dir ('品質評価_{0}.xlsx' -f $b[0])) -Sheets @(
            @{ Name = '品質評価'; Columns = $metricColumns; Rows = @(@{ Values = @('テスト', '消化率', $b[1], '×', '100% 以上なら ○', '', '') }) },
            @{ Name = '要対応'; Columns = $actionColumns; Rows = @(
                    @{ Values = @('テスト', 'SCR001', 'ゾーン: テスト不足', 'テストケースを追加する', '', '') },
                    @{ Values = @('テスト', 'SCR002', 'ゾーン: 品質に問題', '原因を分析する', '', '') },
                    @{ Values = @('工程', '詳細設計', '流出 2 件', 'レビューの観点を見直す', '', '') }) })
    }
    # 前の週のブックを、あとで直して保存した（更新日時は今週のブックより新しい）
    (Get-Item -LiteralPath (Join-Path $dir '品質評価_20261015.xlsx')).LastWriteTime = (Get-Date).AddMinutes(5)
    $text = Get-HostText { & $QualityAnalysis $dir -PromptOnly | Out-Null }
    Assert-Match $text '依頼を作りました: .*品質評価_20261022_定性分析の依頼\.txt' '名前の日付が一番新しいブック'
    $prompt = [System.IO.File]::ReadAllText((Join-Path $dir '品質評価_20261022_定性分析の依頼.txt'))
    Assert-Match $prompt '\| テスト \| 消化率 \| 0\.8 \|'
    Assert-Match $prompt '# 要対応'
    foreach ($row in @('\| テスト \| SCR001 \| ゾーン: テスト不足 \|', '\| テスト \| SCR002 \| ゾーン: 品質に問題 \|', '\| 工程 \| 詳細設計 \| 流出 2 件 \|')) { Assert-Match $prompt $row '要対応のすべての行' }
}

Test-Case 'Invoke-QualityAnalysis: Get-Help でヘルプが表示される' {
    $h = Get-Help $QualityAnalysis
    Assert-Match $h.Synopsis '定性分析（原因）と対策の下書き'
    Assert-Equal 6 @($h.parameters.parameter).Count
}
