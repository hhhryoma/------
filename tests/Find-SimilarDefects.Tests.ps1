#     Find-SimilarDefects.ps1 のテスト。設計書は docs\design-sample、障害管理表はテストの中で作る

$SimilarTool = Join-Path $script:ToolsDir 'Find-SimilarDefects.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function New-BugCsv {
    #     障害管理表（CSV）。Rows: @(障害番号, 機能, 障害内容, 原因, 対策, 横展開, 状態)
    param([string] $Path, $Rows)
    $lines = @('障害番号,機能,障害内容,原因,対策,横展開,状態') + @($Rows | ForEach-Object { ($_ | ForEach-Object { '"{0}"' -f $_ }) -join ',' })
    [System.IO.File]::WriteAllText($Path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

function Get-CandidateKeys {
    param($Results)
    return @($Results | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.No, $_.Bug, $_.Kind, $_.Target, $_.Score })
}

$bugRows = @(
    , @('B001', 'SCR002 受注登録', '受注日に 2028/02/29 を入れると E0004 になる', '日付形式のチェックでうるう年を考慮していなかった', '日付の共通チェックを直す', '要', '完了')
    , @('B002', '受注詳細', '更新で、ほかの人が先に更新していても上書きされる', 'バージョン番号の比較が漏れていた', '更新の条件にバージョン番号を足す', '', '対応中')
    , @('B003', 'SCR001', '画面の背景色が仕様と違う', 'CSS の指定誤り', '色を直す', '', '完了')
    , @('B004', 'SCR001', '受注番号の検索で前方一致にならない', 'LIKE の条件の書き間違い', '', '不要', '完了')
    , @('B005', 'SCR003', 'キャンセルの確認で Q0002 の文言が違う', '', '', '', '取下げ'))

Test-Case 'Find-SimilarDefects: 障害の項目・イベントと同じ作り（型・入力チェック・メッセージ・処理）の箇所を、点数の高い順に集める' {
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs $bugRows
    $results = @(& $SimilarTool -BugList $bugs -DesignPath $designDir 6>$null)
    Assert-Equal @(
        '1|B001|項目|SCR001 受注日（From）|6',
        '2|B001|項目|SCR001 受注日（To）|6',
        '3|B001|項目|SCR002 希望納品日|6',
        '4|B001|項目|SCR003 希望納品日|5',
        '5|B002|イベント|SCR003 キャンセル（EV03）|8',
        '6|B002|イベント|SCR002 登録（EV06）|5',
        '7|B002|イベント|SCR001 検索（EV02）|4') (Get-CandidateKeys $results) '入力の項目には、表示だけの項目（SCR001 受注日 など）を出さない。横展開が不要の B004・取下げの B005 は探さない'
    Assert-Equal 'SCR002 受注日' $results[0].Anchor '障害の画面（機能の SCR002）の項目'
    Assert-Equal '型 日付・入力チェック 日付形式・桁数 10・備考が同じ（yyyy/MM/dd）' $results[0].Common
    Assert-Equal 'SCR003 更新（EV02）' $results[4].Anchor '画面名（受注詳細）から画面を決め、「更新」はその画面のイベント'
    Assert-Equal 'メッセージ E0006・メッセージ I0002・処理 画面表示時・処理 再表示・原因の言葉 バージョン' $results[4].Common '原因の言葉（バージョン）。2 文字の言葉（受注・更新）と、対策の「条件」は「検索条件」の中に探さない'

    $text = Get-HostText { & $SimilarTool -BugList $bugs -DesignPath $designDir | Out-Null }
    Assert-Match $text '障害 5 件のうち、横展開を探した障害 3 件（横展開が不要・済みのもの、障害ではないもの 2 件を除く）'
    Assert-Match $text '候補あり 2 件（候補 7 か所。項目 4・イベント 3）'
    Assert-Match $text '障害の箇所が設計書に見つからない障害 1 件: B003'
}

Test-Case 'Find-SimilarDefects: 表示の項目・カラム・メッセージ ID の障害と、-Id・-MaxCandidates' {
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs @(
        , @('B006', 'SCR001', '受注一覧の合計金額にカンマが付かない', '金額の表示書式の指定漏れ', '', '', '完了')
        , @('B007', 'バッチ', 'T_ORDER.TOTAL_AMOUNT が 11 桁を超えると登録でエラーになる', '桁数の確認漏れ', '', '', '完了')
        , @('B008', '', '在庫が足りないのに E0008 が出ない', '在庫の引当の条件誤り', '', '', '完了')
        , @('B009', 'SCR001', '受注番号の検索で前方一致にならない', 'LIKE の条件の書き間違い', '', '不要', '完了'))
    $results = @(& $SimilarTool -BugList $bugs -DesignPath $designDir -Id 'B00[678]' -MaxCandidates 3 6>$null)
    Assert-Equal @(
        '1|B006|項目|SCR002 合計金額|5',
        '2|B006|項目|SCR003 合計金額|5',
        '3|B006|項目|SCR002 金額|3',
        '4|B007|カラム|T_ORDER_DETAIL.AMOUNT（金額）|5',
        '5|B007|カラム|M_PRODUCT.UNIT_PRICE（単価）|3',
        '6|B007|カラム|M_PRODUCT.STOCK_QTY（在庫数）|3',
        '7|B008|イベント|SCR002 登録（EV06）|3') (Get-CandidateKeys $results) '表示の項目には表示する項目。画面の分からない障害の「登録」はイベントとみなさない。障害ごとに 3 か所まで'
    Assert-Equal 'T_ORDER.TOTAL_AMOUNT（合計金額）' $results[3].Anchor
    Assert-Equal '型 NUMERIC・桁数 11・NOT NULL・論理名が「金額」で終わる' $results[3].Common
    Assert-Equal 'メッセージ E0008' $results[6].Common 'メッセージ ID の障害は、そのメッセージを使う箇所'

    # -Id で指定すれば、横展開が不要の障害も探す
    $only = @(& $SimilarTool -BugList $bugs -DesignPath $designDir -Id 'b009' 6>$null)
    $anchors = @($only | ForEach-Object { $_.Anchor } | Select-Object -Unique)
    Assert-Equal 2 $anchors.Count ($anchors -join ' / ')
    Assert-True ($anchors -contains 'SCR001 受注番号' -and $anchors -contains 'SCR001 検索（EV02）') '障害内容に出てくる項目とイベント（入力の受注番号。表示の受注番号ではない）'
}

Test-Case 'Find-SimilarDefects: 空白で区切ったメッセージ ID・カラム名・画面 ID も、前後の英数字とつなげずに見つける' {
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs @(
        , @('B011', '', 'メッセージID E0006 が出ない', '', '', '', '')
        , @('B012', '', 'E0006 E0008 が出ない', '', '', '', '')
        , @('B013', 'バッチ', 'UPDATE T_ORDER SET TOTAL_AMOUNT = ? が 11 桁を超えるとエラー', '', '', '', '')
        , @('B014', '画面ID SCR003', '希望納品日に 2028/02/29 を入れると E0004 になる', '', '', '', '')
        , @('B015', 'バッチ', 'T_ORDER_DETAIL . ORDER_NO が重複する', '', '', '', ''))
    $results = @(& $SimilarTool -BugList $bugs -DesignPath $designDir -MaxCandidates 2 6>$null)
    Assert-Equal @(
        '1|B011|イベント|SCR003 更新（EV02）|3',
        '2|B011|イベント|SCR003 キャンセル（EV03）|3',
        '3|B012|イベント|SCR002 登録（EV06）|3',
        '4|B012|イベント|SCR003 更新（EV02）|3',
        '5|B013|カラム|T_ORDER_DETAIL.AMOUNT（金額）|5',
        '6|B013|カラム|M_PRODUCT.UNIT_PRICE（単価）|3',
        '7|B014|項目|SCR002 希望納品日|9',
        '8|B014|項目|SCR001 受注日（From）|5',
        '9|B015|カラム|T_ORDER.ORDER_NO（受注番号）|8',
        '10|B015|カラム|M_CUSTOMER.CUSTOMER_CD（顧客コード）|3') (Get-CandidateKeys $results) '「メッセージID E0006」の ID と E0006、「E0006 E0008」の 2 つの ID をつなげない'
    Assert-Match $results[0].Anchor '^E0006「'
    Assert-Equal 'T_ORDER.TOTAL_AMOUNT（合計金額）' $results[4].Anchor 'SQL の文の中のカラム名'
    Assert-Equal 'SCR003 希望納品日' $results[6].Anchor '機能の「画面ID SCR003」の画面の項目'
    Assert-Equal 'T_ORDER_DETAIL.ORDER_NO（受注番号）' $results[8].Anchor 'テーブル名.カラム名の「.」の前後の空白は、これまでどおり許す'
}

Test-Case 'Find-SimilarDefects: 画面名を含む長い項目名（画面「ログイン」の「ログインID」）は、項目として見つける' {
    $design = New-TestDir
    # 1 行目（画面ID・画面名）は、下の表の列の数だけ書く（New-SheetsBook は 1 行目の数の列を書く）
    $header = { param($id, $name) @('画面ID', $id, '画面名', $name, '') }
    $table = @(@(), @('項目名', '入出力', '型', '桁数', '入力チェック'), @('ログインID', 'I', '半角英数', '20', '必須(E0001)、半角英数(E0003)'))
    New-SheetsBook (Join-Path $design '画面設計書.xlsx') ([ordered]@{
            'SCR001' = @(, (& $header 'SCR001' 'ログイン')) + $table + @(, @('ログイン', '-', '-', '', ''))
            'SCR002' = @(, (& $header 'SCR002' 'ユーザー登録')) + $table
            'SCR003' = @(, (& $header 'SCR003' 'パスワード変更')) + $table })
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs @(
        , @('B021', 'SCR002', 'ログインIDに全角を入れても登録できてしまう', '半角英数のチェック漏れ', '', '', '')
        , @('B022', 'SCR001', 'ログイン画面でログインIDを空にしてもエラーにならない', '', '', '', ''))
    $out = Join-Path $dir '横展開.xlsx'
    $results = @(& $SimilarTool -BugList $bugs -DesignPath $design -OutXlsx $out 6>$null)
    Assert-Equal @(
        '1|B021|項目|SCR001 ログインID|11',
        '2|B021|項目|SCR003 ログインID|11',
        '3|B022|項目|SCR002 ログインID|11',
        '4|B022|項目|SCR003 ログインID|11') (Get-CandidateKeys $results)
    $sheet = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal 'B021|SCR002 ログインID / B022|SCR001 ログインID' ((@(2, 3) | ForEach-Object { '{0}|{1}' -f $sheet.Rows[$_]['A'], $sheet.Rows[$_]['E'] }) -join ' / ') '「ログイン画面」は画面名で、ボタンの「ログイン」の項目ではない'
}

Test-Case 'Find-SimilarDefects: 桁数の無いカラム（DATE など）どうしは、桁数を共通点にしない' {
    $design = New-TestDir
    $header = @('論理名', '物理名', 'データ型', '桁数', '必須')
    New-SheetsBook (Join-Path $design 'テーブル定義書.xlsx') ([ordered]@{
            'T_SHIP' = @($header, @('出荷予定', 'SHIP_PLAN', 'DATE', '', '○'))
            'T_EMP' = @($header, @('入社', 'HIRED', 'DATE', '', '○'))
            'T_ITEM' = @($header, @('発売', 'RELEASED', 'DATE', '', '')) })
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs @(, @('B031', 'バッチ', 'T_SHIP.SHIP_PLAN が NULL で落ちる', '', '', '', ''))
    $results = @(& $SimilarTool -BugList $bugs -DesignPath $design 6>$null)
    Assert-Equal @('1|B031|カラム|T_EMP.HIRED（入社）|3') (Get-CandidateKeys $results) '型だけが同じカラム（T_ITEM.RELEASED。2 点）は候補にしない'
    Assert-Equal '型 DATE・NOT NULL' $results[0].Common
    Assert-Equal '論理名 入社・型 DATE・NOT NULL' $results[0].Detail
}

Test-Case 'Find-SimilarDefects: -OutPrompt で Copilot への依頼を作り、-Answer の回答（表）の判定と理由を一覧と -OutXlsx に入れる' {
    $dir = New-TestDir
    $bugs = Join-Path $dir '障害管理表.csv'
    New-BugCsv $bugs $bugRows
    $prompt = Join-Path $dir '依頼/横展開の依頼.txt'
    $null = & $SimilarTool -BugList $bugs -DesignPath $designDir -OutPrompt $prompt 6>$null
    $p = [System.IO.File]::ReadAllText($prompt)
    Assert-Match $p '# 障害 B001'
    Assert-Match $p '- 障害の箇所: SCR002 受注日（型 日付・桁数 10・必須・入出力 IO・入力チェック 必須\(E0001\)、日付形式\(E0004\)・備考 yyyy/MM/dd）'
    Assert-Match $p '\| 1 \| SCR001 受注日（From） \| 型 日付・桁数 10・入出力 I・入力チェック 日付形式\(E0004\)・備考 yyyy/MM/dd \| 型 日付'
    Assert-Match $p '\| No \| 判定 \| 理由 \|'
    Assert-True ($p -notmatch '↵') '処理の改行は空白にする'
    Assert-True ($p -notmatch 'B003') '候補の無い障害は入れない'

    $answer = Join-Path $dir '横展開の回答.txt'
    [System.IO.File]::WriteAllText($answer, (@(
                '| No | 判定 | 理由 |', '| --- | --- | --- |',
                '| 1 | 要確認 | 同じ日付形式のチェック |',
                '| 2 | **要確認** | 同じ日付形式のチェック |',
                "4`t不要`t変更できる場合が少ない",
                '| 5 | 要確認 | 同じバージョン番号の比較 |') -join "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
    $out = Join-Path $dir '横展開.xlsx'
    $text = Get-HostText { $script:answered = @(& $SimilarTool -BugList $bugs -DesignPath $designDir -Answer $answer -OutXlsx $out) }
    Assert-Equal '要確認|要確認||不要|要確認' ((@($script:answered | Select-Object -First 5 | ForEach-Object { $_.Judgement })) -join '|')
    Assert-Equal '変更できる場合が少ない' $script:answered[3].Reason 'タブ区切り（Copilot の表をコピーしたもの）も読む'
    Assert-Match $text 'AI の判定: 要確認 3・不要 1（回答の無い候補 3）'

    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('横展開の候補', '障害', '読み取り結果') $book.Sheets
    Assert-Equal 'No|障害ID|障害の箇所|種類|候補の箇所|設計の内容|共通点|点数|AI の判定|AI の理由|確認結果|設計書' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal 'SCR001 受注日（From）|要確認|同じ日付形式のチェック' ('{0}|{1}|{2}' -f $book.Rows[2]['E'], $book.Rows[2]['I'], $book.Rows[2]['J'])
    Assert-True (@($book.Links | Where-Object { $_.Target -like '*DD02_*SCR001*.xlsx' -and $_.Location -like "'画面項目定義'!A*" }).Count -gt 0) '設計書の行へのリンク'
    $sheet = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal 'B003|0|障害の箇所（画面の項目・イベント・カラム・メッセージ ID）が、障害内容・原因・対策に見つかりません' ('{0}|{1}|{2}' -f $sheet.Rows[4]['A'], $sheet.Rows[4]['F'], $sheet.Rows[4]['G'])
}

Test-Case 'Find-SimilarDefects: Get-Help でヘルプが表示される' {
    $h = Get-Help $SimilarTool
    Assert-Match $h.Synopsis '横展開'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
