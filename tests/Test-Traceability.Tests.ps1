#     Test-Traceability.ps1 のテスト。設計書は docs\design-sample、テスト仕様書・要件一覧はテストの中で作る

$TraceTool = Join-Path $script:ToolsDir 'Test-Traceability.ps1'
$traceDesignDir = Join-Path $script:RepoRoot 'docs/design-sample'

function New-TraceBook {
    param([string] $FilePath, $Sheets)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $FilePath -Sheets @($Sheets | ForEach-Object {
            @{ Name = $_.Name; Columns = @($_.Header | ForEach-Object { @{ Header = $_; Width = 12 } }); Rows = @($_.Rows | ForEach-Object { @{ Values = $_ } }) }
        })
}

function New-TraceTests {
    #     テスト仕様書: 機能の列（空なら上の行）で画面を分けるシートと、上に「画面ID」を書いたシート
    param([string] $Dir)
    New-TraceBook (Join-Path $Dir 'UT_受注.xlsx') @(
        @{ Name = 'ケース'; Header = @('No', '機能', '確認内容', '結果'); Rows = @(
                , @('1', 'SCR001', '初期表示で受注日の初期値を確かめる', 'OK')
                , @('2', '', '検索ボタンで受注を検索する', 'OK')
                , @('3', '', '受注番号に全角を入れると E0003 が出る', 'OK')
                , @('4', 'SCR002', '登録ボタンで受注を登録する（REQ-002）', 'OK')
                , @('5', '', '要件 REQ-010 の確認', 'OK')) },
        @{ Name = '受注詳細'; Header = @('画面ID', 'SCR003', ''); Rows = @(
                , @('No', '確認内容', '結果')
                , @('1', 'キャンセルで一覧に戻る', 'OK')) })
}

Test-Case 'Test-Traceability: 画面・イベント・メッセージ・IF にテストがあるかを、画面ごとの範囲で確かめる' {
    $dir = New-TestDir
    New-TraceTests $dir
    $gaps = @(& $TraceTool -DesignPath $traceDesignDir -TestPath $dir 6>$null)
    $keys = @($gaps | ForEach-Object { '{0}|{1}|{2}' -f $_.Kind, $_.Scope, $_.Id })
    # 画面: SCR001・SCR002 は機能の列、SCR003 はシートの上の「画面ID」で、どれもテストがある
    Assert-Equal 0 @($keys | Where-Object { $_ -like '画面|*' }).Count ($keys -join ', ')
    # イベント: SCR001 の初期表示・検索はある（3 行目は機能が空なので上の行の SCR001）。クリアは無い
    Assert-True ($keys -notcontains 'イベント|SCR001|EV01') '初期表示'
    Assert-True ($keys -notcontains 'イベント|SCR001|EV02') '検索'
    Assert-True ($keys -contains 'イベント|SCR001|EV03') 'クリア'
    Assert-True ($keys -notcontains 'イベント|SCR003|EV03') 'キャンセル（シートの上の画面ID）'
    Assert-True ($keys -notcontains 'イベント|SCR003|EV04') '戻る（「一覧に戻る」に名前が出てくる）'
    # メッセージ: E0003 は SCR001 のテストにはあるが、SCR002 のテストには無い
    Assert-True ($keys -notcontains 'メッセージ|SCR001|E0003') 'SCR001 の E0003'
    Assert-True ($keys -contains 'メッセージ|SCR002|E0003') 'SCR002 の E0003'
    Assert-True ($keys -contains 'IF||IF001') 'IF'
    $clear = @($gaps | Where-Object { $_.Id -eq 'EV03' -and $_.Scope -eq 'SCR001' })[0]
    Assert-Equal 'テストなし（名前が出てこない）|クリア' ('{0}|{1}' -f $clear.Judgement, $clear.Name)
    Assert-Match $clear.Place '^DD02_画面設計書_SCR001_受注一覧\.xlsx イベント定義!A8$'

    $text = Get-HostText { & $TraceTool -DesignPath $traceDesignDir -TestPath $dir | Out-Null }
    Assert-Match $text 'テストの網羅: 画面 3/3（100%）・イベント 5/17（29%）・メッセージ 1/24（4%）・IF 0/1（0%）'
}

Test-Case 'Test-Traceability: 要件一覧の要件が設計書とテスト仕様書に出てくるかを確かめ、一覧に無い要件 ID も知らせる' {
    $root = New-TestDir
    $design = Join-Path $root 'design'
    Copy-Item -LiteralPath $traceDesignDir -Destination $design -Recurse
    New-TraceBook (Join-Path $design 'DD00_要件対応表.xlsx') @(@{ Name = '対応表'; Header = @('要件', '画面ID'); Rows = @(
                , @('REQ-001', 'SCR001')
                , @('REQ-002', 'SCR002')
                , @('REQ-009', 'SCR003')) })
    # 要件一覧は設計書のフォルダに置いても、設計書としては読まない
    $reqCsv = Join-Path $design '要件一覧.csv'
    @('要件ID,要件名', 'REQ-001,受注を検索する', 'REQ-002,受注を登録する', 'REQ-003,出荷指示を送る') | Set-Content -LiteralPath $reqCsv -Encoding UTF8
    $tests = Join-Path $root 'tests'
    New-Item -ItemType Directory -Path $tests | Out-Null
    New-TraceTests $tests
    $out = Join-Path $root 'trace.xlsx'
    $gaps = @(& $TraceTool -DesignPath $design -TestPath $tests -Requirement $reqCsv -OutXlsx $out 6>$null)
    Assert-Equal @('REQ-001|テストなし', 'REQ-003|設計・テストなし', 'REQ-009|要件一覧に無い要件 ID', 'REQ-010|要件一覧に無い要件 ID') @($gaps | Where-Object Kind -eq '要件' | ForEach-Object { '{0}|{1}' -f $_.Id, $_.Judgement })

    $text = Get-HostText { & $TraceTool -DesignPath $design -TestPath $tests -Requirement $reqCsv | Out-Null }
    Assert-Match $text '要件 3 件: 設計とテストあり 1・設計なし 0・テストなし 1・設計・テストなし 1。要件一覧に無い要件 ID 2 件'

    $book = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('抜け', '網羅', '要件') $book.Sheets
    Assert-Equal 'REQ-002|受注を登録する|DD00_要件対応表.xlsx|UT_受注.xlsx|1|○' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $book.Rows[3]['A'], $book.Rows[3]['B'], $book.Rows[3]['C'], $book.Rows[3]['D'], $book.Rows[3]['E'], $book.Rows[3]['F'])
    $cover = Read-XlsxSheet -Path $out -Index 2
    $first = @($cover.Rows.Values | Where-Object { $_['C'] -eq 'SCR001' -and $_['A'] -eq '画面' })[0]
    Assert-Equal '○' $first['H']
    Assert-True (@($cover.Links | Where-Object { $_.Location -like "'ケース'!*" }).Count -gt 0) 'テストの場所へのリンク'
}

Test-Case 'Test-Traceability: 行の機能は表の見出しの行（表の上の「テスト対象」の行ではない）の列から読み、機能名より画面ID の列を使う' {
    $dir = New-TestDir
    # 表の上に「テスト対象: 受注管理システム」の行があり、画面ID は最初の行だけに書いたシート
    New-TraceBook (Join-Path $dir 'UT_受注一覧.xlsx') @(@{ Name = 'ケース'; Header = @('テスト対象', '受注管理システム', '', ''); Rows = @(
                , @('No', '画面ID', '確認内容', '結果')
                , @('1', 'SCR001', '初期表示で受注日の初期値を確かめる', 'OK')
                , @('2', '', '検索ボタンで受注を検索する', 'OK')
                , @('3', '', 'クリアボタンで検索条件を消す', 'OK')) })
    # 機能（機能名。どの行にもある）と画面ID（最初の行だけ）の列があるシート
    New-TraceBook (Join-Path $dir 'UT_受注詳細.xlsx') @(@{ Name = 'ケース'; Header = @('No', '機能', '画面ID', '確認内容', '結果'); Rows = @(
                , @('1', '受注詳細', 'SCR003', '初期表示で受注の内容を表示する', 'OK')
                , @('2', '受注詳細', '', '更新ボタンで受注を更新する', 'OK')
                , @('3', '受注詳細', '', 'キャンセルボタンで変更を取り消す', 'OK')) })
    $gaps = @(& $TraceTool -DesignPath $traceDesignDir -TestPath $dir 6>$null)
    $keys = @($gaps | Where-Object Kind -eq 'イベント' | ForEach-Object { '{0}|{1}' -f $_.Scope, $_.Id })
    Assert-True ($keys -notcontains 'SCR001|EV02') ('検索（表の画面ID の列の、上の行の SCR001）: ' + ($keys -join ', '))
    Assert-True ($keys -notcontains 'SCR001|EV03') 'クリア'
    Assert-True ($keys -contains 'SCR001|EV04') '新規登録（テストが無い）'
    Assert-True ($keys -notcontains 'SCR003|EV02') ('更新（機能の列ではなく画面ID の列の、上の行の SCR003）: ' + ($keys -join ', '))
    Assert-True ($keys -notcontains 'SCR003|EV03') 'キャンセル'
    Assert-True ($keys -contains 'SCR003|EV04') '戻る（テストが無い）'
}

Test-Case 'Test-Traceability: 結果が削除・対象外・欠番の行は、テストとみなさない' {
    $dir = New-TestDir
    New-TraceBook (Join-Path $dir 'UT_受注.xlsx') @(@{ Name = 'ケース'; Header = @('No', '機能', '確認内容', '結果'); Rows = @(
                , @('1', 'SCR001', '初期表示で受注日の初期値を確かめる', 'OK')
                , @('2', '', '検索ボタンで受注を検索する', '削除')
                , @('3', '', 'クリアボタンで検索条件を消す', '対象外')
                , @('4', '', '新規登録ボタンで登録画面を開く', '欠番')
                , @('5', 'SCR003', 'キャンセルボタンで変更を取り消す', '削除')
                , @('6', '', '戻るボタンで一覧に戻る', '')) })
    $gaps = @(& $TraceTool -DesignPath $traceDesignDir -TestPath $dir 6>$null)
    $keys = @($gaps | Where-Object Kind -eq 'イベント' | ForEach-Object { '{0}|{1}' -f $_.Scope, $_.Id })
    Assert-True ($keys -notcontains 'SCR001|EV01') '初期表示（OK）'
    Assert-True ($keys -contains 'SCR001|EV02') ('検索（削除）: ' + ($keys -join ', '))
    Assert-True ($keys -contains 'SCR001|EV03') 'クリア（対象外）'
    Assert-True ($keys -contains 'SCR001|EV04') '新規登録（欠番）'
    Assert-True ($keys -contains 'SCR003|EV03') 'キャンセル（削除）'
    # 削除の行の機能（SCR003）は、下の行に引き継ぐ
    Assert-True ($keys -notcontains 'SCR003|EV04') '戻る（未実施）'
}

Test-Case 'Test-Traceability: 要件一覧は ID の見出しがいちばん詳しい行の表を読み、ID が数字だけの表（改訂履歴など）は読まない' {
    $root = New-TestDir
    $tests = Join-Path $root 'tests'
    New-Item -ItemType Directory -Path $tests | Out-Null
    New-TraceTests $tests
    # 改訂履歴（No・内容）のシートと、2 行の見出し（No・要件 の下に ID・内容）の要件一覧のシート
    $req = Join-Path $root '要件一覧.xlsx'
    New-SheetsBook $req ([ordered]@{
            '改訂履歴' = @(@('No', '改訂日', '内容', '担当'), @('1', '2026/09/01', '新規作成', '山田'), @('2', '2026/10/01', 'REQ-003 を追加', '山田'))
            '要件一覧' = @(@('No', '要件', '', '優先度'), @('', 'ID', '内容', ''), @('1', 'REQ-001', '受注を検索する', '高'), @('2', 'REQ-002', '受注を登録する', '高'), @('3', 'REQ-003', '出荷指示を送る', '中')) })
    # 要件の表の無いブックは知らせる
    $memo = Join-Path $root '要件メモ.xlsx'
    New-SheetsBook $memo ([ordered]@{ 'メモ' = @(@('No', '内容'), @('1', '打ち合わせで決めたこと')) })
    $out = Join-Path $root 'trace.xlsx'
    $text = Get-HostText { $script:traceReqGaps = @(& $TraceTool -DesignPath $traceDesignDir -TestPath $tests -Requirement $req, $memo -OutXlsx $out) }
    Assert-Match $text '要件 3 件: 設計とテストあり 0・設計なし 1・テストなし 0・設計・テストなし 2。要件一覧に無い要件 ID 1 件'
    Assert-Match $text '要件メモ\.xlsx: 要件の表が見つかりません'
    Assert-NotMatch $text '要件一覧\.xlsx: 要件の表が見つかりません'
    $book = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('REQ-001|受注を検索する', 'REQ-002|受注を登録する', 'REQ-003|出荷指示を送る') @(2..4 | ForEach-Object { '{0}|{1}' -f $book.Rows[$_]['A'], $book.Rows[$_]['B'] })
    # 一覧の ID の形（REQ- と 3 桁の数字）で、一覧に無い要件 ID も探す
    Assert-Equal @('REQ-010') @($script:traceReqGaps | Where-Object Judgement -eq '要件一覧に無い要件 ID' | ForEach-Object { $_.Id })
}

Test-Case 'Test-Traceability: 要件一覧の CSV も、ID の見出しがいちばん詳しい列（No より要件ID）を読む' {
    $root = New-TestDir
    $tests = Join-Path $root 'tests'
    New-Item -ItemType Directory -Path $tests | Out-Null
    New-TraceTests $tests
    $req = Join-Path $root '要件一覧.csv'
    @('No,要件ID,要件名', '1,REQ-001,受注を検索する', '2,REQ-002,受注を登録する') | Set-Content -LiteralPath $req -Encoding UTF8
    $text = Get-HostText { $script:traceCsvGaps = @(& $TraceTool -DesignPath $traceDesignDir -TestPath $tests -Requirement $req 3>$null) }
    Assert-Match $text '要件 2 件'
    Assert-Equal @() @($script:traceCsvGaps | Where-Object { $_.Kind -eq '要件' -and $_.Id -notmatch '^REQ-' } | ForEach-Object { $_.Id }) '要件 ID は No（1・2）ではなく REQ-001・REQ-002'
    Assert-Equal @('REQ-010') @($script:traceCsvGaps | Where-Object Judgement -eq '要件一覧に無い要件 ID' | ForEach-Object { $_.Id })
}

Test-Case 'Test-Traceability: Get-Help でヘルプが表示される' {
    $h = Get-Help $TraceTool
    Assert-Match $h.Synopsis 'テスト'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
