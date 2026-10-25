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
    Assert-Equal @('突き合わせ', '観点別', '設計書別', '観点表への反映案', '読み取り結果') $points.Sheets
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

Test-Case 'Compare-ReviewFindings: 人だけの指摘と AI だけの指摘から、観点表に足す観点・具体的にする判定基準・見直す観点の案を出す' {
    $dir = New-TestDir
    $checklist = Join-Path $dir '観点表.csv'
    @('観点ID,分類,観点,判定基準,重要度,対象設計書,有効',
        'S02,画面項目,入力チェックのエラー時のメッセージが決まっている,入力チェックごとにエラー時のメッセージIDが書かれている,中,*画面設計書*,○',
        'S03,画面項目,選択項目・日付項目の初期値が決まっている,選択項目と日付項目の初期値が書かれている,中,*画面設計書*,○',
        'X01,記述,ログ出力の方針が書かれている,使わない観点,低,*,×') | Set-Content -LiteralPath $checklist -Encoding UTF8
    $scr = 'DD02_画面設計書_SCR001_受注一覧'
    @('設計書,指摘内容,場所,観点ID',
        "$scr,受注ステータスの桁数が無い,画面項目定義!AB11,S01",
        "$scr,数量のエラー時のメッセージが無い,画面項目定義!AK7,S02",
        "$scr,単価のエラー時のメッセージが無い,画面項目定義!AK8,S02") | Set-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8
    @('設計書,指摘内容,場所,分類,重要度',
        "$scr,受注ステータスの桁数が未記入,画面項目定義!AB12,,",
        "$scr,受注日の初期値が書かれていない,画面項目定義!M20,画面項目,高",
        "$scr,ログ出力の方針が書かれていない,画面概要!C5,非機能,中",
        'DD05_メッセージ一覧,句読点の揺れ,メッセージ一覧!K6,表記,低') | Set-Content -LiteralPath (Join-Path $dir 'human.csv') -Encoding UTF8
    $out = Join-Path $dir 'compare.xlsx'
    $null = & $CompareFindings (Join-Path $dir 'ai.csv') (Join-Path $dir 'human.csv') -Checklist $checklist -OutXlsx $out 6>$null
    $book = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal '提案|観点ID|分類|観点|判定基準|重要度|対象設計書|件数' (@('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H' | ForEach-Object { $book.Rows[1][$_] }) -join '|') 'B〜G 列は観点表と同じ並び'
    Assert-Equal @(
        '判定基準を具体的にする|S03|画面項目|選択項目・日付項目の初期値が決まっている|中|*画面設計書*|1',
        '観点を足す||非機能||中|*画面設計書*|1',
        '観点を足す||表記||低|*|1',
        '判定基準を見直す|S02|画面項目|入力チェックのエラー時のメッセージが決まっている|中|*画面設計書*|2') @(2..5 | ForEach-Object { $r = $book.Rows[$_]; (@('A', 'B', 'C', 'D', 'F', 'G', 'H' | ForEach-Object { $r[$_] }) -join '|') })
    Assert-Equal '受注日の初期値が書かれていない（DD02_画面設計書_SCR001_受注一覧）' $book.Rows[2]['I']
    Assert-Match $book.Rows[5]['J'] 'AI の指摘 2 件のうち 2 件が AI だけ'
    $text = Get-HostText { & $CompareFindings (Join-Path $dir 'ai.csv') (Join-Path $dir 'human.csv') -Checklist $checklist | Out-Null }
    Assert-Match $text '観点表への反映案: .*観点を足す 2'

    # Invoke-DesignReview のツールの確認の指摘（判定元がツール）は、判定基準を見直す案に数えない
    $lines = @(Get-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8 | ForEach-Object {
            if ($_ -like '設計書,*') { "$_,判定元" } elseif ($_ -like '*,S02') { "$_,ツール（入力チェックのメッセージ）" } else { "$_,AI" } })
    $lines | Set-Content -LiteralPath (Join-Path $dir 'ai-tool.csv') -Encoding UTF8
    $null = & $CompareFindings (Join-Path $dir 'ai-tool.csv') (Join-Path $dir 'human.csv') -Checklist $checklist -OutXlsx $out 6>$null
    $book = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal @('判定基準を具体的にする', '観点を足す', '観点を足す') @(2..4 | ForEach-Object { $book.Rows[$_]['A'] })
    Assert-Equal 4 $book.Rows.Count '判定基準を見直す案が無い'
}

Test-Case 'Compare-ReviewFindings: 判定が OK の行を数えないのは AI のレビュー結果だけで、人の指摘管理表の判定（修正の確認）が OK の指摘は数える' {
    $dir = New-TestDir
    $scr = 'DD02_画面設計書_SCR001_受注一覧'
    @('設計書,指摘内容,場所,判定', "$scr.xlsx,受注ステータスの桁数が無い,画面項目定義!AB11,NG", "$scr.xlsx,-,,OK") | Set-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8
    @('設計書,指摘内容,場所,対応状況,判定',
        "$scr,受注ステータスの桁数が未記入,画面項目定義!AB12,完了,OK",
        "$scr,画面の目的が不明確,画面概要!C5,完了,OK") | Set-Content -LiteralPath (Join-Path $dir 'human.csv') -Encoding UTF8
    $r = @(& $CompareFindings (Join-Path $dir 'ai.csv') (Join-Path $dir 'human.csv') 6>$null)
    Assert-Equal @('一致|場所|画面項目定義!AB11|画面項目定義!AB12', '人だけ|||画面概要!C5') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Status, $_.Basis, $_.AiLocation, $_.HumanLocation }) 'AI の判定が OK の行は数えない'
}

Test-Case 'Compare-ReviewFindings: 設計書が空欄の行（結合したセルなど）は、同じシートの上の行の設計書とみなす' {
    $dir = New-TestDir
    $scr = 'DD02_画面設計書_SCR001_受注一覧'
    # 対象成果物を B2:B4 で結合する（値は左上の B2 にだけある）。別のシートの空欄は、前のシートの設計書を引き継がない
    $book = Join-Path $dir '指摘管理表.xlsx'
    New-SheetsBook -Path $book -Sheets ([ordered]@{
            '指摘' = @(
                @('No', '対象成果物', 'シート', '箇所', '指摘内容', '対応状況'),
                @('1', $scr, '画面項目定義', 'AB12', '受注ステータスの桁数が未記入', '完了'),
                @('2', '', 'イベント定義', 'J20', '「押下」と「クリック」が混在', '完了'),
                @('3', '', '画面概要', 'C5', '画面の目的が不明確', '完了'))
            '指摘2' = @(
                @('No', '対象成果物', 'シート', '箇所', '指摘内容', '対応状況'),
                @('1', '', 'イベント定義', 'J21', 'ボタン名の誤り', '完了'))
        })
    Edit-ZipText -Path $book -Name 'xl/worksheets/sheet1.xml' -Change { param($t) $t -replace '</sheetData>', '</sheetData><mergeCells count="1"><mergeCell ref="B2:B4"/></mergeCells>' }
    @('設計書,指摘内容,場所', "$scr.xlsx,受注ステータスの桁数が無い,画面項目定義!AB11", "$scr.xlsx,「押下」と「クリック」が混在している,イベント定義!J20") | Set-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8
    $r = @(& $CompareFindings (Join-Path $dir 'ai.csv') $book 6>$null)
    Assert-Equal @(
        '|人だけ||イベント定義!J21',
        "$scr|一致|イベント定義!J20|イベント定義!J20",
        "$scr|一致|画面項目定義!AB11|画面項目定義!AB12",
        "$scr|人だけ||画面概要!C5") @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Document, $_.Status, $_.AiLocation, $_.HumanLocation })
}

Test-Case 'Compare-ReviewFindings: 「IF001 B6」「画面項目定義（AB12）」「画面項目定義：AB12」や、Word の「本文!段落 12」「本文!表 1 の 3 行目」の場所も、場所の近さで対応付ける' {
    $dir = New-TestDir
    # 設計書ごとに AI と人の指摘を 1 件ずつ。指摘内容は似せず、場所だけで対応付くかを見る
    $cases = @(
        , @('01_IF定義書', 'IF001!B5', 'IF001 B6')                           # 0 で始まる ID はセル番地ではない
        , @('02_テーブル定義書', 'T01!E7', 'T01 E8')
        , @('03_IF定義書', 'IF1!B5', 'IF1 B6')                               # シート名の後の、最後のセル番地
        , @('04_画面設計書', '画面項目定義!AB11', '画面項目定義（AB12）')
        , @('05_画面設計書', '画面項目定義!AB11', '画面項目定義：AB12')
        , @('06_画面設計書', '画面項目定義!AB11', '画面項目定義 AB12:AB15')   # 範囲は先頭のセル
        , @('07_基本設計書', '本文!段落 12', '本文!段落12')
        , @('08_基本設計書', '本文!表 1 の 3 行目', '本文!表 1 の 4 行目')
        , @('09_基本設計書', '本文!表 2 の 3 行目', '本文!表 1 の 3 行目'))   # 別の表の行は近くない
    $aiLines = @('設計書,指摘内容,場所'); $humanLines = @('設計書,指摘内容,場所')
    $n = 0
    foreach ($c in $cases) { $n++; $aiLines += ('{0},AI の指摘その{1},{2}' -f $c[0], $n, $c[1]); $humanLines += ('{0},全く違う書き方{1},{2}' -f $c[0], $n, $c[2]) }
    $aiLines | Set-Content -LiteralPath (Join-Path $dir 'ai.csv') -Encoding UTF8
    $humanLines | Set-Content -LiteralPath (Join-Path $dir 'human.csv') -Encoding UTF8
    $r = @(& $CompareFindings (Join-Path $dir 'ai.csv') (Join-Path $dir 'human.csv') 6>$null)
    Assert-Equal @(
        '01_IF定義書|一致|場所', '02_テーブル定義書|一致|場所', '03_IF定義書|一致|場所', '04_画面設計書|一致|場所', '05_画面設計書|一致|場所', '06_画面設計書|一致|場所',
        '07_基本設計書|一致|場所', '08_基本設計書|一致|場所', '09_基本設計書|AI だけ|', '09_基本設計書|人だけ|') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Document, $_.Status, $_.Basis })
}

Test-Case 'Compare-ReviewFindings: 設計書の名前の版数（_v1.0・_v1.1）は除いて同じ設計書とみなす（版を上げたあとの名前で書いた指摘管理表）' {
    $dir = New-TestDir
    New-FindingSheets $dir
    # AI は v1.0 のファイルをレビューし、人は版を上げたあとのファイル名で指摘管理表を書いた
    $ai = Join-Path $dir 'レビュー結果.xlsx'
    Edit-ZipText $ai 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('DD02_画面設計書_SCR001_受注一覧.xlsx', 'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx') }
    $human = Join-Path $dir '指摘管理表.xlsx'
    Edit-ZipText $human 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('DD02_画面設計書_SCR001_受注一覧', 'DD02_画面設計書_SCR001_受注一覧_v1.1.xlsx') }
    $text = Get-HostText { $script:r = @(& $CompareFindings $ai $human) }
    Assert-Match $text 'AI の指摘 5 件: 人と一致 2（40\.0%）・AI だけ 3'
    Assert-Equal 2 @($script:r | Where-Object Status -eq '一致').Count
}

Test-Case 'Compare-ReviewFindings: Get-Help でヘルプが表示される' {
    $h = Get-Help $CompareFindings
    Assert-Match $h.Synopsis 'AI'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
