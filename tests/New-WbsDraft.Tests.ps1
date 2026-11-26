#     New-WbsDraft.ps1 のテスト。設計書は docs\design-sample（規模: 画面 3・テーブル 4・IF 1・メッセージ・コード。合計 20.1 人日）

$WbsTool = Join-Path $script:ToolsDir 'New-WbsDraft.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Get-WbsKeys {
    param($Rows)
    return @($Rows | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Wbs, $_.Task, $_.Member, $_.Start.ToString('MM/dd'), $_.End.ToString('MM/dd'), $_.Days })
}

function Test-WbsOrder {
    #     同じ機能の工程は順に（次の工程の開始日は、前の工程の終了日より前にならない）、別の人の工程（レビュー）は作った人と違う担当者
    param($Rows, [string] $ReviewPattern = 'レビュー）$')
    $tasks = @($Rows | Where-Object { -not $_.IsParent })
    foreach ($g in @($tasks | Group-Object { $_.Task -replace '（[^（]+）$' })) {
        $chain = @($g.Group | Sort-Object { [int](($_.Wbs -split '\.')[0]) })
        for ($i = 1; $i -lt $chain.Count; $i++) { Assert-True ($chain[$i].Start -ge $chain[$i - 1].End) ('{0} は {1} が終わってから' -f $chain[$i].Task, $chain[$i - 1].Task) }
        $author = @($chain | Where-Object { $_.Task -notmatch $ReviewPattern })[0].Member
        foreach ($r in @($chain | Where-Object { $_.Task -match $ReviewPattern })) { Assert-True ($r.Member -ne $author) ('{0} は作った人（{1}）以外' -f $r.Task, $author) }
    }
}

Test-Case 'New-WbsDraft: 規模の見積もりの工数を工程に配分し、担当者に割り当てて、土日・休日を除いた予定日の WBS を作る（レビューは別の人・工程の順・小さい機能をまとめる）' {
    $dir = New-TestDir
    $holidays = Join-Path $dir 'syukujitsu.csv'
    @('国民の祝日・休日月日,国民の祝日・休日名称', '2026/11/3,文化の日', '2026/11/23,勤労感謝の日') | Set-Content -LiteralPath $holidays -Encoding UTF8
    $rows = @(& $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤' -Holidays $holidays 6>$null)
    $keys = Get-WbsKeys $rows
    Assert-Equal 30 $keys.Count '工程 4 + タスク 26（テーブル 4 つを 2 つにまとめる。テーブル・メッセージ・コードは単体テストが無い）'
    Assert-Equal @(
        '1|詳細設計||11/02|11/06|5.89',
        '1.1|SCR001 受注一覧（詳細設計）|山田|11/02|11/04|1.25',
        '1.2|SCR002 受注登録（詳細設計）|佐藤|11/02|11/04|1.25',
        '1.3|SCR003 受注詳細（詳細設計）|山田|11/04|11/05|1.25',
        '1.4|M_CUSTOMER 顧客マスタ ほか 1 件（詳細設計）|佐藤|11/04|11/04|0.36') @($keys | Select-Object -First 5) '11/03（文化の日）は営業日にしない。工数の大きい機能から、手の空く担当者へ'
    Assert-True ($keys -contains '1.7|DD05_メッセージ一覧（詳細設計）|佐藤|11/05|11/06|0.67') 'メッセージ一覧は詳細設計・レビューに配分する（25:5）。種類の違うコードとはまとめない'
    # レビューは作った人以外。製造はレビューが終わってから（SCR001 のレビューは 11/06 に終わり、製造は 11/06 から）
    Assert-True ($keys -contains '2.1|SCR001 受注一覧（詳細設計レビュー）|佐藤|11/06|11/06|0.25')
    Assert-True ($keys -contains '3.1|SCR001 受注一覧（製造）|山田|11/06|11/10|2')
    Test-WbsOrder $rows
    Assert-Equal '|1.1|2.1|3.1|' ((@('1.1', '2.1', '3.1', '4.1', '1') | ForEach-Object { $w = $_; @($rows | Where-Object { $_.Wbs -eq $w })[0].Predecessors }) -join '|') '先行は同じ機能の前の工程の WBS 番号（工程の行は空）'
    Assert-Equal @('4|単体テスト||11/12|11/17|5.1', '4.1|SCR001 受注一覧（単体テスト）|山田|11/12|11/16|1.5') @($keys | Where-Object { $_ -match '^4(\.1)?\|' })
    $tables = @($rows | Where-Object { $_.Task -like 'M_CUSTOMER*' })
    Assert-Equal '0.36|0.07|0.57' ((@($tables | ForEach-Object { $_.Days })) -join '|') '0.5 人日のテーブル 2 つを 1 人日のタスクにまとめ、配分を丸めた残りは最後の工程に入れる'
    Assert-Equal 'M_CUSTOMER 顧客マスタ: テーブル 簡易（カラム数 14） / M_PRODUCT 商品マスタ: テーブル 簡易（カラム数 10）' $tables[0].Basis
    Assert-Equal @('詳細設計書', '詳細設計書', 'レビュー記録票', 'ソースコード', '単体テスト仕様書・成績書') @($rows[0], $rows[1], ($rows | Where-Object { $_.Wbs -eq '2.1' }), ($rows | Where-Object { $_.Wbs -eq '3.1' }), ($rows | Where-Object { $_.Wbs -eq '4.1' }) | ForEach-Object { $_.Deliverable }) '工程の成果物'

    $text = Get-HostText { & $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤' -Holidays $holidays | Out-Null }
    Assert-Match $text '機能 10・工程 4・タスク 26。予定工数 合計 20.1 人日'
    Assert-Match $text '1 人日に満たない機能は、同じ種類をまとめて 1 つのタスクにしました: テーブル 4 件 → 2 つ'
    Assert-Match $text '担当者 2 人: 山田 10.17 人日（2026/11/02〜2026/11/17）・佐藤 9.93 人日（2026/11/02〜2026/11/16）'
    Assert-Match $text '完了予定: 2026/11/17（2026/11/02 から 11 営業日）'
    Assert-True ($text -notmatch '休日の CSV') '休日を指定したときは知らせない'
}

Test-Case 'New-WbsDraft: 稼働率（佐藤:50%）・待ち合わせの工程・担当の列の無い工程の配分（名前にレビュー）・-BundleBelow 0' {
    $dir = New-TestDir
    $phases = Join-Path $dir 'phases.csv'
    @('工程,配分,種類,待ち合わせ', '設計,30,,', '設計レビュー,10,,', '製造,40,画面|IF|テーブル,', '結合テスト,20,画面|IF,する') | Set-Content -LiteralPath $phases -Encoding UTF8
    $rows = @(& $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤：50%' -Phases $phases -BundleBelow 0 6>$null)
    $keys = Get-WbsKeys $rows
    Assert-Equal 36 $keys.Count '工程 4 + タスク 32（まとめない）'
    Assert-True ($keys -contains '1.3|SCR003 受注詳細（設計）|佐藤|11/02|11/04|1.5') '稼働率 50% の佐藤は、1.5 人日に 3 営業日'
    Assert-True ($keys -contains '1.4|M_CUSTOMER 顧客マスタ（設計）|山田|11/05|11/05|0.19') 'まとめない'
    Test-WbsOrder $rows
    $tasks = @($rows | Where-Object { -not $_.IsParent })
    $lastBefore = @($tasks | Where-Object { $_.Task -notmatch '結合テスト' } | Sort-Object End -Descending)[0].End
    $firstIt = @($tasks | Where-Object { $_.Task -match '結合テスト' } | Sort-Object Start)[0].Start
    Assert-Equal ([datetime]'2026-11-17') $lastBefore
    Assert-True ($firstIt -ge $lastBefore) '待ち合わせの結合テストは、前の工程がすべて終わってから'
    Assert-Equal '1,2,3' @($rows | Where-Object { $_.Wbs -eq '4.1' })[0].Predecessors '待ち合わせの工程の先行は、前の工程の番号すべて'
    Assert-Equal '' $rows[1].Deliverable '成果物の列が無ければ空'

    $text = Get-HostText { & $WbsTool $designDir -Start '2026-11-02' -Members '山田, 佐藤：50%' -Phases $phases -BundleBelow 0 | Out-Null }
    Assert-Match $text '担当者 2 人: 山田 13\.36 人日（2026/11/02〜2026/11/19）・佐藤 6\.74 人日・稼働率 50%（2026/11/02〜2026/11/19）'
    Assert-Match $text '休日の CSV（-Holidays）を指定していないので'
    Assert-True ($text -notmatch 'まとめて') '-BundleBelow 0 はまとめない'

    $warnings = @(& $WbsTool $designDir -Start '2026-11-02' -Members '山田' 6>$null 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] })
    Assert-Match ($warnings -join "`n") '担当者が 1 人なので、作った人に割り当てました'
    $failed = $false
    try { $null = & $WbsTool $designDir -Members '山田, 佐藤:150%' 6>$null } catch { $failed = $true; Assert-Match $_.Exception.Message '稼働率' }
    Assert-True $failed '稼働率は 100% まで'
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
    Assert-Equal 'WBS番号|タスク名|成果物|担当者|先行|予定開始|予定終了|予定工数|実績開始|実績終了|進捗率|実績工数|見積もりの根拠' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal ([string](([datetime]'2026-11-02') - [datetime]'1899-12-30').TotalDays) ([string]$book.Rows[3]['F']) '予定開始は日付のセル'
    $load = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal '担当者|稼働率|タスクの数|工数の合計（人日）|うちレビューなど（人日）' ((@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $load.Rows[1][$_] }) -join '|')
    Assert-Equal '100%' $load.Rows[2]['B']
    $size = Read-XlsxSheet -Path $out -Index 3
    $customer = @($size.Rows.Values | Where-Object { $_['B'] -eq 'M_CUSTOMER 顧客マスタ' })[0]
    Assert-Equal 'M_CUSTOMER 顧客マスタ ほか 1 件' $customer['F'] '規模のシートは機能ごと。まとめた先のタスク'

    $progress = @(& (Join-Path $script:ToolsDir 'Measure-Progress.ps1') $out -AsOf '2026-11-05' 6>$null)
    Assert-Equal 26 $progress.Count '親のタスク（工程）を除くタスク。担当者別・規模のシートは WBS として読まない'
    Assert-Equal '1.1|SCR001 受注一覧（詳細設計）|山田' ('{0}|{1}|{2}' -f $progress[0].WbsNo, $progress[0].Task, $progress[0].Owner)
    $text = Get-HostText { & (Join-Path $script:ToolsDir 'Measure-Progress.ps1') $out -AsOf '2026-11-05' | Out-Null }
    Assert-Match $text 'BAC 20\.1'

    $bytes = [System.IO.File]::ReadAllBytes($csv)
    Assert-Equal '239|187|191' (@($bytes[0..2]) -join '|') 'CSV は BOM 付き'
    $lines = [System.IO.File]::ReadAllLines($csv)
    Assert-Equal '"1.1","SCR001 受注一覧（詳細設計）","詳細設計書","山田","","2026/11/02","2026/11/03","1.25","","","","","画面 普通（項目数 17・イベント数 6）"' $lines[2]
    Assert-Match ($lines -join "`n") '"2\.1","SCR001 受注一覧（詳細設計レビュー）","レビュー記録票","佐藤","1\.1",' '先行は同じ機能の前の工程'
}

Test-Case 'New-WbsDraft: Get-Help でヘルプが表示される' {
    $h = Get-Help $WbsTool
    Assert-Match $h.Synopsis 'WBS'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'New-WbsDraft: -Gantt で WBS のシートをガントチャート形式の進捗管理票にし、休日のシートを足す（Measure-Progress.ps1 がそのまま読める）' {
    $dir = New-TestDir
    $holidays = Join-Path $dir 'holidays.csv'
    @('2026/11/03,文化の日') | Set-Content -LiteralPath $holidays -Encoding UTF8
    $out = Join-Path $dir 'WBS.xlsx'
    $rows = @(& $WbsTool $designDir -Start '2026-11-02' -Members '山田', '佐藤' -Holidays $holidays -OutXlsx $out -Gantt 6>$null 3>$null)
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('WBS', '担当者別', '規模', '休日') $book.Sheets
    # 3 行目が見出し（見積もりの根拠は規模のシート。実績工数は空でも置く）。日付の列は 11/02 の月曜から
    Assert-Equal @('WBS番号', 'タスク名', '担当者', '先行', '予定開始', '予定終了', '予定工数', '実績開始', '実績終了', '進捗率', '実績工数', '状態') @('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L' | ForEach-Object { $book.Rows[3][$_] })
    Assert-Equal '46328' $book.Rows[3]['M']
    Assert-Equal '46329' (Read-XlsxSheet -Path $out -Index 4).Rows[2]['A'] '休日のシートの 11/03'
    $tasks = @(& (Join-Path $script:ToolsDir 'Measure-Progress.ps1') -Path $out -AsOf '2026-11-02' -Holidays $holidays 6>$null 3>$null)
    Assert-Equal 26 $tasks.Count
    Assert-Equal @($rows | Where-Object { -not $_.IsParent } | ForEach-Object { '{0}|{1}' -f $_.Wbs, $_.Task }) @($tasks | ForEach-Object { '{0}|{1}' -f $_.WbsNo, $_.Task })

    $failed = $false
    try { $null = & $WbsTool $designDir -Gantt 6>$null 3>$null } catch { $failed = $true; Assert-Match $_.Exception.Message '-OutXlsx' }
    Assert-True $failed '-Gantt は -OutXlsx と一緒に'
}
