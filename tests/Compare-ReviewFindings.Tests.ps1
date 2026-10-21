#     Compare-ReviewFindings.ps1 のテスト。AI のレビュー結果と人の指摘管理表はテストの中で作る

$CompareFindings = Join-Path $script:ToolsDir 'Compare-ReviewFindings.ps1'

function New-FindingSheets {
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $scr = 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $aiRows = @(
        , @('1', 'NG', '高', 'S01', '網羅性', '入力項目', $scr, '画面項目定義!AB11', '受注ステータスの桁数が無い')
        , @('2', 'NG', '中', 'S02', '網羅性', 'メッセージ', $scr, '画面項目定義!AK7', 'エラー時のメッセージが無い')
        , @('3', '要確認', '中', 'S06', '異常時', '異常時の動き', $scr, 'イベント定義!P8', 'クリア時の動きが不明')
        , @('4', 'NG', '低', 'C03', '表記', '用語', $scr, 'イベント定義!J9', '「押下」と「クリック」が混在している')
        , @('5', 'OK', '低', 'C01', '表記', '表紙', $scr, '', '-')
        , @('6', 'NG', '高', 'T01', '整合性', '型', 'DD03_テーブル定義書.xlsx', 'T_ORDER!E7', '型の誤り'))
    Export-Workbook -FilePath (Join-Path $Dir 'レビュー結果.xlsx') -Sheets @(@{
            Name = '指摘一覧'; Columns = @(@('No', '判定', '重要度', '観点ID', '分類', '観点', '設計書', '場所', '指摘内容') | ForEach-Object { @{ Header = $_; Width = 12 } })
            Rows = @($aiRows | ForEach-Object { @{ Values = $_ } })
        })
    $humanRows = @(
        , @('1', 'DD02_画面設計書_SCR001_受注一覧', '画面項目定義', 'AB12', '受注ステータスの桁数が未記入', '完了')
        , @('2', 'DD02_画面設計書_SCR001_受注一覧', 'イベント定義', 'J20', '「押下」と「クリック」が混在', '完了')
        , @('3', 'DD02_画面設計書_SCR001_受注一覧', '画面概要', 'C5', '画面の目的が不明確', '対応中')
        , @('4', 'DD03_テーブル定義書', 'T_ORDER', 'E7', '型の誤り', '取下げ')
        , @('5', 'DD05_メッセージ一覧', 'メッセージ一覧', 'K6', '句読点の揺れ', '完了'))
    Export-Workbook -FilePath (Join-Path $Dir '指摘管理表.xlsx') -Sheets @(@{
            Name = '指摘'; Columns = @(@('No', '対象成果物', 'シート', '箇所', '指摘内容', '対応状況') | ForEach-Object { @{ Header = $_; Width = 12 } })
            Rows = @($humanRows | ForEach-Object { @{ Values = $_ } })
        })
}

Test-Case 'Compare-ReviewFindings: AI と人の指摘を、場所の近さと内容の似ている度合いで突き合わせる' {
    $dir = New-TestDir
    New-FindingSheets $dir
    $out = Join-Path $dir 'compare.xlsx'
    $r = @(& $CompareFindings (Join-Path $dir 'レビュー結果.xlsx') (Join-Path $dir '指摘管理表.xlsx') -OutXlsx $out 6>$null)
    Assert-Equal @(
        '一致|場所|画面項目定義!AB11|画面項目定義!AB12',
        '一致|内容|イベント定義!J9|イベント定義!J20',
        'AI だけ||画面項目定義!AK7|',
        'AI だけ||イベント定義!P8|',
        '人だけ|||画面概要!C5',
        'AI だけ||T_ORDER!E7|',
        '人だけ|||メッセージ一覧!K6') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Status, $_.Basis, $_.AiLocation, $_.HumanLocation })
    Assert-Equal 'DD02_画面設計書_SCR001_受注一覧' $r[2].Document '拡張子は付けない' '判定が OK の行と取下げの指摘は数えない'

    $text = Get-HostText { & $CompareFindings (Join-Path $dir 'レビュー結果.xlsx') (Join-Path $dir '指摘管理表.xlsx') | Out-Null }
    Assert-Match $text 'AI の指摘 5 件: 人と一致 2（40\.0%）・AI だけ 3'
    Assert-Match $text '人の指摘 4 件: AI と一致 2（50\.0%）・人だけ 2'

    $points = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('突き合わせ', '観点別', '設計書別', '読み取り結果') $points.Sheets
    Assert-Equal 'C03|1|1|1' ('{0}|{1}|{2}|{3}' -f $points.Rows[2]['A'], $points.Rows[2]['B'], $points.Rows[2]['C'], $points.Rows[2]['D'])
    $pairsSheet = Read-XlsxSheet -Path $out -Index 1
    Assert-True (@($pairsSheet.Links | Where-Object { $_.Location -eq "'指摘'!E2" }).Count -eq 1) '人の指摘管理表の行へのリンク'
    Assert-True (@($pairsSheet.Links | Where-Object { $_.Location -eq "'指摘一覧'!I3" }).Count -eq 1) 'AI だけの行はレビュー結果の行へのリンク'
}

Test-Case 'Compare-ReviewFindings: 場所を「シート名 セル」の形で書いた指摘も、場所の近さで対応付ける' {
    $dir = New-TestDir
    @('設計書,指摘内容,場所', 'DD02_画面設計書_SCR001_受注一覧,全く違う書き方の指摘,画面項目定義 AB10') | Set-Content -LiteralPath (Join-Path $dir 'human.csv') -Encoding UTF8
    @('設計書,指摘内容,場所', 'DD02_画面設計書_SCR001_受注一覧.xlsx,受注ステータスの桁数が無い,画面項目定義!AB11') | Set-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8
    $r = @(& $CompareFindings (Join-Path $dir 'ai.csv') (Join-Path $dir 'human.csv') 6>$null)
    Assert-Equal '一致|場所' ('{0}|{1}' -f $r[0].Status, $r[0].Basis)
}

Test-Case 'Compare-ReviewFindings: Get-Help でヘルプが表示される' {
    $h = Get-Help $CompareFindings
    Assert-Match $h.Synopsis 'AI'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
