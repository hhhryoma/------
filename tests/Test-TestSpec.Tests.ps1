#     Test-TestSpec.ps1 のテスト。テスト仕様書はテストの中で作る（Export-Workbook）。設計書は docs\design-sample

$TestSpecTool = Join-Path $script:ToolsDir 'Test-TestSpec.ps1'

function New-SpecBook {
    param([string] $FilePath, [string[]] $Header, $Rows)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $FilePath -Sheets @(@{ Name = 'テストケース'; Columns = @($Header | ForEach-Object { @{ Header = $_; Width = 14 } }); Rows = @($Rows | ForEach-Object { @{ Values = $_ } }) })
}

function New-Spec {
    param([string] $Dir)
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add(@('1', 'SCR001', '正常', '', '検索ボタンを押す', '受注番号: A001', '受注番号 A001 の受注が 1 件表示される', 'OK'))
    $rows.Add(@('2', '', '正常', '', '検索ボタンを押す', '', '正しく表示されること', 'OK'))                       # あいまい
    $rows.Add(@('3', '', '異常', '', '', '', '', ''))                                                             # 期待結果・手順が空
    $rows.Add(@('4', '', '異常', '', '受注番号に 13 桁を入れて検索する', '1234567890123', 'E0002 が表示される', 'OK'))  # 受注番号の境界値
    $rows.Add(@('5', '', '正常', '', '検索ボタンを押す', '受注番号: A001', '受注番号 A001 の受注が 1 件表示される', 'たぶんOK'))  # 1 と同じ内容、読めない結果
    $rows.Add(@('5', '', '正常', '', 'クリアボタンを押す', '', '検索条件が空になる', 'OK'))                         # テストIDの重複
    for ($i = 7; $i -le 12; $i++) { $rows.Add(@([string]$i, $(if ($i -eq 7) { 'SCR002' } else { '' }), '正常', '', "登録の確認 $i", '', "登録した受注 $i が一覧に表示される", 'OK')) }
    $path = Join-Path $Dir 'UT_受注.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') $rows
    return $path
}

Test-Case 'Test-TestSpec: 期待結果の空欄・あいまいさ、手順の空欄、ID と内容の重複、読めない結果、異常系の少ない機能を見つける' {
    $path = New-Spec (New-TestDir)
    $issues = @(& $TestSpecTool -Path $path 6>$null)
    Assert-Equal @(
        'エラー|期待結果が空|SCR001|3|テストケース!G4',
        'エラー|テストIDの重複|SCR001|5|テストケース!A7',
        '警告|あいまいな期待結果|SCR001|2|テストケース!G3',
        '警告|手順が空|SCR001|3|テストケース!E4',
        '警告|同じ内容のケース|SCR001|5|テストケース!A6',
        '警告|結果の値が読めない|SCR001|5|テストケース!H6',
        '警告|異常系が少ない|SCR002|7|テストケース!A8') @($issues | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Level, $_.Item, $_.Function, $_.Id, $_.Location })
    Assert-Match @($issues | Where-Object Item -eq 'あいまいな期待結果')[0].Detail '^「正しく」では'
    Assert-Match @($issues | Where-Object Item -eq '異常系が少ない')[0].Detail '^異常系のケースが 0 件 / 6 件（0%）で、20% 未満です（区分の列）'
    $text = Get-HostText { & $TestSpecTool -Path $path | Out-Null }
    Assert-Match $text 'テストケース 12 件（機能 2）: エラー 2 件 / 警告 5 件'
}

Test-Case 'Test-TestSpec: -DesignPath の画面項目定義の桁数から、境界値のケースが無い入力項目を知らせる' {
    $dir = New-TestDir
    $path = New-Spec $dir
    $out = Join-Path $dir 'spec-check.xlsx'
    $issues = @(& $TestSpecTool -Path $path -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') -OutXlsx $out 6>$null)
    $boundary = @($issues | Where-Object Item -eq '境界値のテストが無い')
    $scr001 = @($boundary | Where-Object { $_.Detail -like 'SCR001 *' } | ForEach-Object { [regex]::Match($_.Detail, '「(.+?)」').Groups[1].Value })
    Assert-True ($scr001 -notcontains '受注番号') '13 桁を入れるケースがある受注番号'
    Assert-True ($scr001 -contains '顧客コード') ($scr001 -join ', ')
    Assert-Match @($boundary | Where-Object { $_.Detail -like '*「顧客コード」*' })[0].Detail '（8 桁。DD02_画面設計書_SCR001_受注一覧\.xlsx 画面項目定義 \d+ 行目）の桁数の境界（8 桁・9 桁）'
    Assert-Equal 0 @($boundary | Where-Object { $_.Detail -like 'SCR003 *' }).Count 'テストの無い画面は確かめない'
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('指摘', '機能別', '読み取り結果') $book.Sheets
    Assert-True (@($book.Links | Where-Object { $_.Location -eq "'テストケース'!G4" }).Count -eq 1) '期待結果のセルへのリンク'
    $functions = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal 'SCR001|6|2|0.333|区分の列' ('{0}|{1}|{2}|{3}|{4}' -f $functions.Rows[2]['A'], $functions.Rows[2]['B'], $functions.Rows[2]['C'], $functions.Rows[2]['D'], $functions.Rows[2]['E'])
}

Test-Case 'Test-TestSpec: -DesignPath の必須・型・コード定義書から、未入力・不正な値・選ぶ値ごとのケースが無い入力項目を知らせる' {
    $dir = New-TestDir
    $rows = @(
        @('1', 'SCR002', '異常', '', '受注日を未入力で登録する', '', 'E0001 が表示される', 'OK'),
        @('2', '', '異常', '', '受注日に存在しない日付を入れて登録する', '2025/02/30', 'E0004 が表示される', 'OK'),
        @('3', '', '正常', '', '支払方法に 銀行振込 を選んで登録する', '', '登録される', 'OK'))
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') $rows
    $issues = @(& $TestSpecTool -Path $path -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') 6>$null)
    $of = { param($item) @($issues | Where-Object Item -eq $item | ForEach-Object { [regex]::Match($_.Detail, '「(.+?)」').Groups[1].Value }) }
    $required = & $of '必須のテストが無い'
    Assert-True ($required -notcontains '受注日') '未入力のケースがある受注日'
    Assert-True ($required -contains '顧客コード') ($required -join ', ')
    $type = & $of '型のテストが無い'
    Assert-True ($type -notcontains '受注日') '存在しない日付のケースがある受注日'
    Assert-True ($type -contains '希望納品日') ($type -join ', ')
    Assert-Match @($issues | Where-Object { $_.Item -eq '型のテストが無い' -and $_.Detail -like '*「顧客コード」*' })[0].Detail '（型 半角英数。DD02_画面設計書_SCR002_受注登録\.xlsx 画面項目定義 \d+ 行目）に、半角として正しくない値'
    $code = @($issues | Where-Object { $_.Item -eq 'コード値のテストが無い' -and $_.Detail -like '*「支払方法」*' })
    Assert-Equal 1 $code.Count
    Assert-Match $code[0].Detail '値のうち、「クレジットカード」・「代金引換」 を選ぶケースが見つかりません（コード定義書の「支払方法」の 3 値）'
    $boundary = & $of '境界値のテストが無い'
    Assert-True ($boundary -notcontains '支払方法' -and $boundary -notcontains '受注日') '選ぶ項目と日付は桁数の境界を見ない'
    Assert-True ($boundary -contains '備考') ($boundary -join ', ')
    $text = Get-HostText { & $TestSpecTool -Path $path -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') | Out-Null }
    Assert-Match $text '入力項目 \d+ 件（必須 \d+・桁数あり \d+）とコード区分 \d+ 件を確かめました'
}

Test-Case 'Test-TestSpec: 区分の列が無ければ、期待結果の言葉から異常系を推定する' {
    $dir = New-TestDir
    $csv = Join-Path $dir 'cases.csv'
    $lines = @('テストID,確認内容,判定')
    for ($i = 1; $i -le 10; $i++) { $lines += ('T{0},登録した受注 {0} が表示される,OK' -f $i) }
    $lines += 'T11,数量に 0 を入れると E0005 のエラーが表示される,OK'
    $lines | Set-Content -LiteralPath $csv -Encoding UTF8
    $issues = @(& $TestSpecTool -Path $csv 6>$null)
    Assert-Equal 1 $issues.Count
    Assert-Match $issues[0].Detail '1 件 / 11 件（9%）.*（期待結果の言葉から推定）'
    Assert-Equal 'cases.csv|2 行目' ('{0}|{1}' -f $issues[0].File, $issues[0].Location) '機能の最初のケース'
}

Test-Case 'Test-TestSpec: Get-Help でヘルプが表示される' {
    $h = Get-Help $TestSpecTool
    Assert-Match $h.Synopsis 'テスト仕様書'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
