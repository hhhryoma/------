#     Select-RegressionTests.ps1 のテスト。旧版は docs\design-sample、新版はその写しを Replace-Excel で変えて作る

$RegressionTool = Join-Path $script:ToolsDir 'Select-RegressionTests.ps1'

function New-RegressionData {
    #     新版: SCR002 の支払方法の初期値・E0004 の文言・改訂履歴（除く）を変える。テスト仕様書: 5 ケース
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $design = Join-Path $script:RepoRoot 'docs/design-sample'
    $v2 = Join-Path $Dir 'v2'
    New-Item -ItemType Directory -Path $v2 | Out-Null
    Get-ChildItem -LiteralPath $design -Filter '*.xlsx' | Copy-Item -Destination $v2
    $replace = Join-Path $script:ToolsDir 'Replace-Excel.ps1'
    $null = & $replace -Pattern '銀行振込' -Replacement 'クレジットカード' -SimpleMatch -Path (Join-Path $v2 'DD02_画面設計書_SCR002_受注登録.xlsx') -SheetNamePattern '画面項目定義' -Force 6>$null
    $null = & $replace -Pattern '日付の形式が正しくありません' -Replacement '日付の形式が違います' -SimpleMatch -Path (Join-Path $v2 'DD05_メッセージ一覧.xlsx') -Force 6>$null
    $null = & $replace -Pattern '新規作成' -Replacement 'E0004 の文言を変更' -SimpleMatch -Path (Join-Path $v2 'DD05_メッセージ一覧.xlsx') -SheetNamePattern '改訂履歴' -Force 6>$null
    $rows = @(
        @('1', 'SCR002', '正常', '', '支払方法の初期値を確かめる', '', '銀行振込が選ばれている', 'OK'),
        @('2', '', '異常', '', '受注日を未入力で登録する', '', 'E0001 が表示される', 'OK'),
        @('3', '', '異常', '', '受注日に存在しない日付を入れて登録する', '', 'E0004 が表示される', 'OK'),
        @('4', 'SCR001', '正常', '', '検索ボタンで検索する', '', '一覧が表示される', 'OK'),
        @('5', 'SCR003', '異常', '', '出荷予定日に不正な日付を入れる', '', 'E0004 が表示される', 'OK'))
    $spec = Join-Path $Dir 'UT_受注.xlsx'
    Export-Workbook -FilePath $spec -Sheets @(@{ Name = 'テストケース'; Columns = @(@('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') | ForEach-Object { @{ Header = $_; Width = 14 } }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    return [pscustomobject]@{ Old = $design; New = $v2; Spec = $spec }
}

function New-RegressionScreen {
    #     小さな画面設計書を作る。画面概要（画面ID・改訂日・画面名）と、画面項目定義（Header の見出しの表。Items は行の配列）。
    #     Sheets（順序付き辞書）を指定すると、そのシートも後ろに足す
    param([string] $Path, [string] $Screen, [object[]] $Items, $Sheets, [string] $Date = '2026/10/01', [string[]] $Header = @('No', '項目名', '項目ID', '初期値'))
    $book = [ordered]@{ '画面概要' = @(@('画面ID', $Screen, '改訂日', $Date), @('画面名', '受注登録')); '画面項目定義' = @(, $Header) + $Items }
    if ($Sheets) { foreach ($name in $Sheets.Keys) { $book[$name] = $Sheets[$name] } }
    New-SheetsBook $Path $book
}

function New-RegressionSpec {
    #     テスト仕様書（No・機能・区分・手順・期待結果・結果。Rows は行の配列）を作る
    param([string] $Path, [object[]] $Rows)
    New-SheetsBook $Path ([ordered]@{ 'テストケース' = @(, @('No', '機能', '区分', '手順', '期待結果', '結果')) + $Rows })
}

Test-Case 'Select-RegressionTests: 設計書の変更した行の名前が出てくるケース（高）と、変更した画面のケース（中）を選ぶ' {
    $dir = New-TestDir
    $d = New-RegressionData $dir
    $out = Join-Path $dir 'regression.xlsx'
    $r = @(& $RegressionTool -Old $d.Old -New $d.New -TestPath $d.Spec -OutXlsx $out 6>$null)
    Assert-Equal @('高|SCR002|1', '高|SCR002|3', '高|SCR003|5', '中|SCR002|2') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id }) '変更の無い画面（SCR001）のケースは選ばない'
    Assert-Equal '[高] DD02_画面設計書_SCR002_受注登録.xlsx 画面項目定義!AF9（値）「支払方法・paymentMethod」の初期値を変更' $r[0].Reason
    Assert-Equal '[中] DD02_画面設計書_SCR002_受注登録.xlsx 画面項目定義!AF9（値）「支払方法・paymentMethod」の初期値を変更 ／ [高] DD05_メッセージ一覧.xlsx メッセージ一覧!K9（値）「E0004」のメッセージを変更' $r[1].Reason
    Assert-Equal 'テストケース!A2' $r[0].Location

    $text = Get-HostText { & $RegressionTool -Old $d.Old -New $d.New -TestPath $d.Spec | Out-Null }
    Assert-Match $text '設計書の変更 2 件（差分 3 件。表紙・改訂履歴・作成日などは除く）: 再テストするケース 4 件（高 3・中 1）/ テストケース 5 件'
    $book = Read-XlsxSheet -Path $out
    Assert-Equal @('再テストするケース', '設計書の変更', '読み取り結果') $book.Sheets
    Assert-True (@($book.Links | Where-Object { $_.Location -eq "'テストケース'!A2" }).Count -eq 1) 'テスト仕様書の該当セルへのリンク'
}

Test-Case 'Select-RegressionTests: テストが見つからない変更を知らせる' {
    $dir = New-TestDir
    $d = New-RegressionData $dir
    @('No,機能,確認内容,期待結果', '1,SCR001,検索する,一覧が表示される') | Set-Content -LiteralPath (Join-Path $dir 'only-scr001.csv') -Encoding UTF8
    $text = Get-HostText { & $RegressionTool -Old $d.Old -New $d.New -TestPath (Join-Path $dir 'only-scr001.csv') | Out-Null }
    Assert-Match $text '再テストするケース 0 件'
    Assert-Match $text 'テストが見つからない変更 2 件: DD02_画面設計書_SCR002_受注登録\.xlsx 画面項目定義!AF9 「支払方法・paymentMethod」の初期値・DD05_メッセージ一覧\.xlsx メッセージ一覧!K9 「E0004」のメッセージ'
}

Test-Case 'Select-RegressionTests: 設計書のファイルの追加・削除（名前の変更）も、その画面の変更として選び、テストの無い画面を知らせる' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $items = @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', '銀行振込'))
    # 名前の変更（受注登録 → 受注入力）は、旧のファイルの削除と新のファイルの追加になる。SCR010 は新しい画面。
    # 補足資料には画面ID の欄が無いので、ファイル名に出てくる画面 ID の画面とみなす
    New-RegressionScreen (Join-Path $old 'DD02_画面設計書_SCR009_受注登録.xlsx') 'SCR009' $items
    New-RegressionScreen (Join-Path $new 'DD02_画面設計書_SCR009_受注入力.xlsx') 'SCR009' $items
    New-RegressionScreen (Join-Path $new 'DD02_画面設計書_SCR010_受注取消.xlsx') 'SCR010' $items
    New-SheetsBook (Join-Path $new 'DD09_補足資料_SCR009.xlsx') ([ordered]@{ '補足' = @(@('No', '説明'), @('1', '支払方法の選び方')) })
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR001', '正常', '検索ボタンで検索する', '一覧が表示される', 'OK'))
    $text = Get-HostText { $script:reg = @(& $RegressionTool -Old $old -New $new -TestPath $spec) }
    Assert-Equal @('中|SCR009|1') @($script:reg | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id }) '変更の無い画面（SCR001）のケースは選ばない'
    Assert-Equal '[中] DD02_画面設計書_SCR009_受注入力.xlsx（ファイル）を追加 ／ [中] DD09_補足資料_SCR009.xlsx（ファイル）を追加 ／ [中] DD02_画面設計書_SCR009_受注登録.xlsx（ファイル）を削除' $script:reg[0].Reason
    Assert-Match $text '設計書の変更 4 件（差分 4 件。表紙・改訂履歴・作成日などは除く）: 再テストするケース 1 件（高 0・中 1）/ テストケース 2 件'
    Assert-Match $text 'テストが見つからない変更 1 件: DD02_画面設計書_SCR010_受注取消\.xlsx(\n|$)'
}

Test-Case 'Select-RegressionTests: スクリーンショットの貼り替え（画像）・図形・列の変更も、その画面の変更として選ぶ' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $file = 'DD02_画面設計書_SCR009_受注登録.xlsx'
    # 画面項目定義に列（区分）を挿入し、画面概要のスクリーンショットを貼り替えて、注記の図形を足す
    New-RegressionScreen (Join-Path $old $file) 'SCR009' @(1..12 | ForEach-Object { , @("$_", "項目$_", "item$_", "初期値$_") })
    New-RegressionScreen (Join-Path $new $file) 'SCR009' @(1..12 | ForEach-Object { , @("$_", "項目$_", "区分$($_ % 3)", "item$_", "初期値$_") }) -Header @('No', '項目名', '区分', '項目ID', '初期値')
    Add-TestDrawing -Path (Join-Path $old $file) -Items @(@{ Kind = 'pic'; Name = 'レイアウト'; Png = 'Old'; From = @(0, 3); To = @(4, 12) })
    Add-TestDrawing -Path (Join-Path $new $file) -Items @(@{ Kind = 'pic'; Name = 'レイアウト'; Png = 'New'; From = @(0, 3); To = @(4, 12) }, @{ Kind = 'sp'; Name = '注記'; Text = '支払方法は必須'; From = @(5, 3); To = @(8, 5) })
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR001', '正常', '検索ボタンで検索する', '一覧が表示される', 'OK'))
    $r = @(& $RegressionTool -Old $old -New $new -TestPath $spec 6>$null)
    Assert-Equal @('中|SCR009|1') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id })
    # 列の見出しのセルや、図形を置いたセルの行の名前・左の見出しは、変えたものではないので理由に書かない
    Assert-Equal ('[中] {0} 画面概要!A4（画像）を変更 ／ [中] {0} 画面概要!F4（図形）を追加 ／ [中] {0} 画面項目定義!C1（列）を追加' -f $file) $r[0].Reason
}

Test-Case 'Select-RegressionTests: 削除したセル・シートも、旧版の行の名前・シートの名前で理由を示す' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $file = 'DD02_画面設計書_SCR009_受注登録.xlsx'
    # 支払方法の初期値（D3）を空にし（行は動かない）、補足のシートを除く
    New-RegressionScreen (Join-Path $old $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', '銀行振込')) ([ordered]@{ '補足' = @(@('No', '説明'), @('1', '画面の補足')) })
    New-RegressionScreen (Join-Path $new $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', ''))
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR009', '異常', '受注日を未入力で登録する', 'E0001 が表示される', 'OK'))
    $r = @(& $RegressionTool -Old $old -New $new -TestPath $spec 6>$null)
    Assert-Equal @('高|SCR009|1', '中|SCR009|2') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id }) '消した初期値の項目名が出てくるケースは高'
    Assert-Equal ('[高] {0} 画面項目定義!D3（値）「支払方法・paymentMethod」の初期値を削除 ／ [中] {0} 補足（シート）を削除' -f $file) $r[0].Reason
}

Test-Case 'Select-RegressionTests: ハイフン付き・全角の画面 ID（SCR-009 など）も、その画面のケースだけを選ぶ' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $file = 'DD02_画面設計書_SCR-009_受注登録.xlsx'
    # 設計書の画面ID は全角（ＳＣＲ－００９）、テスト仕様書は半角（SCR-009）で書く。支払方法の初期値を変える
    New-RegressionScreen (Join-Path $old $file) 'ＳＣＲ－００９' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', '銀行振込'))
    New-RegressionScreen (Join-Path $new $file) 'ＳＣＲ－００９' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', 'クレジットカード'))
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR-009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR-009', '異常', '受注日を未入力で登録する', 'E0001 が表示される', 'OK'), @('3', 'SCR-010', '正常', '支払方法の一覧を表示する', '一覧が表示される', 'OK'))
    $out = Join-Path $dir 'regression.xlsx'
    $r = @(& $RegressionTool -Old $old -New $new -TestPath $spec -OutXlsx $out 6>$null)
    Assert-Equal @('高|SCR-009|1', '中|SCR-009|2') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id }) 'ほかの画面（SCR-010）のケースは、項目名が出てきても選ばない'
    Assert-Equal 'SCR-009' (Read-XlsxSheet -Path $out -Index 2).Rows[2]['B'] '設計書の変更の画面'
}

Test-Case 'Select-RegressionTests: 再テストするケースの「場所」のセルに、テスト仕様書の該当セルへのリンクを付ける' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $file = 'DD02_画面設計書_SCR009_受注登録.xlsx'
    New-RegressionScreen (Join-Path $old $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', '銀行振込'))
    New-RegressionScreen (Join-Path $new $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', 'クレジットカード'))
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR009', '異常', '受注日を未入力で登録する', 'E0001 が表示される', 'OK'))
    $out = Join-Path $dir 'regression.xlsx'
    $null = & $RegressionTool -Old $old -New $new -TestPath $spec -OutXlsx $out 6>$null
    $book = Read-XlsxSheet -Path $out
    Assert-Equal '場所' $book.Rows[1]['G']
    Assert-Equal @('G2', 'G3') @($book.Links | ForEach-Object { $_.Ref }) '場所（G 列）のセルにリンク'
    Assert-Equal 'テストケース!A2' $book.Rows[2]['G']
}

Test-Case 'Select-RegressionTests: 改版履歴の追記・改訂日の変更（版情報）は、設計書の変更に数えない' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1'; $new = Join-Path $dir 'v2'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    $file = 'DD02_画面設計書_SCR009_受注登録.xlsx'
    # 支払方法の初期値を変え、改版履歴に 1.1 の行を足して、画面概要の改訂日を新しくする
    $history = @(@('版', '日付', '内容'), @('1.0', '2026/10/01', '新規作成'))
    New-RegressionScreen (Join-Path $old $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', '銀行振込')) ([ordered]@{ '改版履歴' = $history })
    New-RegressionScreen (Join-Path $new $file) 'SCR009' @(@('1', '受注日', 'orderDate', 'システム日付'), @('2', '支払方法', 'paymentMethod', 'クレジットカード')) ([ordered]@{ '改版履歴' = $history + @(, @('1.1', '2026/10/20', '支払方法の初期値を変更')) }) -Date '2026/10/20'
    $spec = Join-Path $dir 'UT.xlsx'
    New-RegressionSpec $spec @(@('1', 'SCR009', '正常', '支払方法の初期値を確かめる', '銀行振込が選ばれている', 'OK'), @('2', 'SCR009', '異常', '受注日を未入力で登録する', 'E0001 が表示される', 'OK'))
    $text = Get-HostText { $script:reg = @(& $RegressionTool -Old $old -New $new -TestPath $spec) }
    Assert-Equal @('高|SCR009|1', '中|SCR009|2') @($script:reg | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id })
    Assert-Equal ('[高] {0} 画面項目定義!D3（値）「支払方法・paymentMethod」の初期値を変更' -f $file) $script:reg[0].Reason '改版履歴・改訂日の変更は理由に入れない'
    Assert-Match $text '設計書の変更 1 件（差分 3 件。表紙・改訂履歴・作成日などは除く）'
}

Test-Case 'Select-RegressionTests: テスト仕様書の隣のエビデンスのブック（テストID の名前のブック・フォルダ）の表は、ケースとして選ばない' {
    $dir = New-TestDir
    $d = New-RegressionData $dir
    $spec = Join-Path $dir 'spec'
    New-Item -ItemType Directory -Path (Join-Path $spec 'UT-001') | Out-Null
    New-RegressionSpec (Join-Path $spec 'UT_受注登録.xlsx') @(
        @('UT-001', 'SCR002', '正常', '画面を開く', '支払方法に銀行振込が選ばれている', 'OK'),
        @('UT-002', 'SCR001', '正常', '検索ボタンを押す', '一覧が表示される', 'OK'))
    # UT-001 のエビデンスのブック（支払方法の初期値を確かめた表）
    New-RegressionSpec (Join-Path $spec 'UT-001/エビデンス.xlsx') @(, @('1', '', '正常', '画面を開く', '支払方法: 銀行振込', 'OK'))
    $r = @(& $RegressionTool -Old $d.Old -New $d.New -TestPath $spec 6>$null)
    Assert-Equal @('高|SCR002|UT-001') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Priority, $_.Function, $_.Id })
    Assert-Match (Get-HostText { & $RegressionTool -Old $d.Old -New $d.New -TestPath $spec | Out-Null }) '/ テストケース 2 件'
}

Test-Case 'Select-RegressionTests: Get-Help でヘルプが表示される' {
    $h = Get-Help $RegressionTool
    Assert-Match $h.Synopsis '再テストするテストケースを選びます'
    Assert-Equal 7 @($h.parameters.parameter).Count
}
