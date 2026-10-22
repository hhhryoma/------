#     Invoke-SmokeTest.ps1（Windows の手元での動作確認）のテスト。全ツールが動き、確認リストができることだけを確かめる

Test-Case 'Invoke-SmokeTest: サンプルの資料で全ツールを実行し、結果のファイルと確認リストを作る' {
    $dir = Join-Path (New-TestDir) '動作確認'
    $text = Get-HostText { & (Join-Path $script:RepoRoot 'tests/Invoke-SmokeTest.ps1') -OutDir $dir -NoOpen 3>$null | Out-Null }
    Assert-Match $text '実行 \d+ 件: 成功 \d+・失敗 0・スキップ \d+'
    $list = Read-XlsxSheet -Path (Join-Path $dir '00_確認リスト.xlsx')
    Assert-Equal @('No', 'ツール', '実行', '結果のファイル', 'Excel で開いて見るところ', '確認（OK / NG）', 'エラー') @('A', 'B', 'C', 'D', 'E', 'F', 'G' | ForEach-Object { $list.Rows[1][$_] })
    Assert-True (@($list.Links | Where-Object { $_.Target -like '*12_*.xlsx' }).Count -eq 1) '結果のファイルへのリンク'
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir '17_品質評価/品質評価_20261020_定性分析の依頼.txt'))) '# 前回の定性分析と対策（品質評価_20261006\.xlsx）'
    Assert-True (Test-Path -LiteralPath (Join-Path $dir '18_AIレビュー/prompts/R001.txt')) 'AI レビューの依頼'
}
