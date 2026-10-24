#     New-WbsDraft.ps1 のテスト。設計書は docs\design-sample（規模: 画面 3・テーブル 4・IF 1・メッセージ・コード。合計 20.1 人日）

$WbsTool = Join-Path $script:ToolsDir 'New-WbsDraft.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Get-WbsKeys {
    param($Rows)
    return @($Rows | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Wbs, $_.Task, $_.Member, $_.Start.ToString('MM/dd'), $_.End.ToString('MM/dd'), $_.Days })
}

Test-Case 'New-WbsDraft: 規模の見積もりの工数を工程に配分し、担当者に割り当てて、土日・休日を除いた予定日の WBS を作る' {
    $dir = New-TestDir
    $holidays = Join-Path $dir 'syukujitsu.csv'
    @('国民の祝日・休日月日,国民の祝日・休日名称', '2026/11/3,文化の日', '2026/11/23,勤労感謝の日') | Set-Content -LiteralPath $holidays -Encoding UTF8
    $rows = @(& $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤' -Holidays $holidays 6>$null)
    $keys = Get-WbsKeys $rows
    Assert-Equal 36 $keys.Count '工程 4 + タスク 32（テーブル・メッセージ・コードは単体テストが無い）'
    Assert-Equal @(
        '1|詳細設計||11/02|11/06|5.89',
        '1.1|SCR001 受注一覧（詳細設計）|山田|11/02|11/04|1.25',
        '1.2|SCR002 受注登録（詳細設計）|佐藤|11/02|11/04|1.25',
        '1.3|SCR003 受注詳細（詳細設計）|山田|11/04|11/05|1.25',
        '1.4|M_CUSTOMER 顧客マスタ（詳細設計）|佐藤|11/04|11/04|0.18') @($keys | Select-Object -First 5) '11/03（文化の日）は営業日にしない。工数の大きい機能から、工数の合計の少ない担当者へ'
    Assert-True ($keys -contains '1.9|DD05_メッセージ一覧（詳細設計）|佐藤|11/05|11/06|0.67') 'メッセージ一覧は詳細設計・レビューに配分する（25:5）'
    Assert-True ($keys -contains '3.7|T_ORDER_DETAIL 受注明細（製造）|佐藤|11/11|11/12|0.28') 'テーブルは詳細設計・レビュー・製造に配分する'
    Assert-Equal @('4|単体テスト||11/12|11/17|5.1', '4.1|SCR001 受注一覧（単体テスト）|山田|11/12|11/13|1.5') @($keys | Where-Object { $_ -match '^4(\.1)?\|' })
    $tables = @($rows | Where-Object { $_.Task -like 'M_PRODUCT*' })
    Assert-Equal '0.18|0.04|0.28' ((@($tables | ForEach-Object { $_.Days })) -join '|') '配分を丸めた残りは最後の工程に入れ、機能の工数（0.5）を変えない'
    Assert-Equal 'テーブル 簡易（カラム数 10）' $tables[0].Basis

    $text = Get-HostText { & $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤' -Holidays $holidays | Out-Null }
    Assert-Match $text '機能 10・工程 4・タスク 32。予定工数 合計 20.1 人日'
    Assert-Match $text '担当者 2 人: 山田 10 人日（2026/11/02〜2026/11/16）・佐藤 10.1 人日（2026/11/02〜2026/11/17）'
    Assert-Match $text '完了予定: 2026/11/17（2026/11/02 から 11 営業日）'
}

Test-Case 'New-WbsDraft: -GroupBy Function で機能 → 工程の順にし、-Baseline で新規と変更の機能だけを変更の工数で WBS にする' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1.0'
    $new = Join-Path $dir 'v1.1'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    Get-ChildItem -LiteralPath $designDir -Filter '*.xlsx' | Copy-Item -Destination $old
    Get-ChildItem -LiteralPath $designDir -Filter '*.xlsx' | Copy-Item -Destination $new
    $scr2 = @(Get-ChildItem -LiteralPath $new -Filter '*SCR002*.xlsx')[0].FullName
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -Pattern '銀行振込' -Replacement 'クレジットカード' -SimpleMatch -Path $scr2 -Force 6>$null

    $rows = @(& $WbsTool $new -Baseline $old -Start '2026-12-01' -GroupBy Function 6>$null)
    Assert-Equal @('1|SCR002 受注登録', '1.1|SCR002 受注登録（詳細設計）', '1.2|SCR002 受注登録（詳細設計レビュー）', '1.3|SCR002 受注登録（製造）', '1.4|SCR002 受注登録（単体テスト）') @($rows | ForEach-Object { '{0}|{1}' -f $_.Wbs, $_.Task }) '変わった画面だけ'
    Assert-Match $rows[0].Basis '^変更（項目数 19 → 19（追加 0・変更 1・削除 0）'
    Assert-Equal '' $rows[1].Member '担当者を指定しなければ空'
    Assert-Equal ([datetime]'2026-12-01') $rows[1].Start
    Assert-Equal ([math]::Round(($rows | Select-Object -Skip 1 | Measure-Object Days -Sum).Sum, 2)) $rows[0].Days '機能の行は工程の合計'
}

Test-Case 'New-WbsDraft: -OutXlsx の WBS は Measure-Progress.ps1 がそのまま読め、-OutCsv は BOM 付き' {
    $dir = New-TestDir
    $out = Join-Path $dir 'WBS.xlsx'
    $csv = Join-Path $dir 'WBS.csv'
    $null = & $WbsTool $designDir -Start '2026-11-02' -Members '山田', '佐藤' -OutXlsx $out -OutCsv $csv 6>$null
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('WBS', '担当者別', '規模') $book.Sheets
    Assert-Equal 'WBS番号|タスク名|担当者|予定開始|予定終了|予定工数|実績開始|実績終了|進捗率|実績工数|見積もりの根拠' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal ([string](([datetime]'2026-11-02') - [datetime]'1899-12-30').TotalDays) ([string]$book.Rows[3]['D']) '予定開始は日付のセル'

    $progress = @(& (Join-Path $script:ToolsDir 'Measure-Progress.ps1') $out -AsOf '2026-11-05' 6>$null)
    Assert-Equal 32 $progress.Count '親のタスク（工程）を除くタスク。担当者別・規模のシートは WBS として読まない'
    Assert-Equal '1.1|SCR001 受注一覧（詳細設計）|山田' ('{0}|{1}|{2}' -f $progress[0].WbsNo, $progress[0].Task, $progress[0].Owner)
    $text = Get-HostText { & (Join-Path $script:ToolsDir 'Measure-Progress.ps1') $out -AsOf '2026-11-05' | Out-Null }
    Assert-Match $text 'BAC 20\.1'

    $bytes = [System.IO.File]::ReadAllBytes($csv)
    Assert-Equal '239|187|191' (@($bytes[0..2]) -join '|') 'CSV は BOM 付き'
    $lines = [System.IO.File]::ReadAllLines($csv)
    Assert-Equal '"1.1","SCR001 受注一覧（詳細設計）","山田","2026/11/02","2026/11/03","1.25","","","","","画面 普通（項目数 17・イベント数 6）"' $lines[2]
}

Test-Case 'New-WbsDraft: Get-Help でヘルプが表示される' {
    $h = Get-Help $WbsTool
    Assert-Match $h.Synopsis 'WBS'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
