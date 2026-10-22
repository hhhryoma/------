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

Test-Case 'Select-RegressionTests: Get-Help でヘルプが表示される' {
    $h = Get-Help $RegressionTool
    Assert-Match $h.Synopsis '再テストするテストケースを選びます'
    Assert-Equal 7 @($h.parameters.parameter).Count
}
