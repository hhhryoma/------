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
    Assert-Match $text '入力項目 \d+ 件（必須 \d+・桁数あり \d+）・コード区分 \d+ 件・メッセージ \d+ 件を確かめました'
}

Test-Case 'Test-TestSpec: -DesignPath のメッセージ一覧に無いメッセージ ID と、一覧と違う文言を知らせる' {
    $dir = New-TestDir
    $rows = @(
        @('1', 'SCR002', '異常', '', '受注日を未入力で登録する', '', 'E0001「受注日は必須入力です。」が表示される', 'OK'),
        @('2', '', '異常', '', '受注日に存在しない日付を入れて登録する', '2025/02/30', 'E0004「日付が正しくありません。」が表示される', 'OK'),
        @('3', '', '異常', '', '顧客コードに「ABC」を入れて登録する（REQ-001）', '', 'E9999 が表示される', 'OK'))
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') $rows
    $issues = @(& $TestSpecTool -Path $path -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') 6>$null | Where-Object { $_.Item -match 'メッセージ|あいまい' })
    Assert-Equal @(
        'メッセージの文言が違う|テストケース!G3|「日付が正しくありません。」が、メッセージ一覧の E0004「{0}の日付の形式が正しくありません。」（DD05_メッセージ一覧.xlsx メッセージ一覧 9 行目） と合いません',
        'メッセージIDが一覧に無い|テストケース!G4|メッセージ ID E9999 がメッセージ一覧にありません（誤記か、一覧への追加漏れ）') @($issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Item, $_.Location, $_.Detail })
}

Test-Case 'Test-TestSpec: -DesignPath のメッセージ一覧の、- や _ で区切ったメッセージ ID（MSG-E001 など）も確かめる' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    New-SheetsBook (Join-Path $design 'メッセージ一覧.xlsx') ([ordered]@{ 'メッセージ一覧' = @(@('メッセージID', 'メッセージ'), @('MSG-E001', '{0}は必須入力です。'), @('E0003', '形式が違います。')) })
    $rows = @(
        @('1', 'SCR002', '異常', '', '受注日を未入力で登録する', '', 'MSG-E001「受注日は必須入力です。」が表示される', 'OK'),
        @('2', '', '異常', '', '数量に 0 を入れて登録する', '', 'MSG-E099 が表示される', 'OK'))
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') $rows
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null | Where-Object { $_.Item -match 'メッセージ' })
    Assert-Equal @('メッセージIDが一覧に無い|テストケース!G3|メッセージ ID MSG-E099 がメッセージ一覧にありません（誤記か、一覧への追加漏れ）') @($issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Item, $_.Location, $_.Detail })
}

Test-Case 'Test-TestSpec: -DesignPath の項目名と違う書き方（顧客CD・受注日付など）を、項目名の揺れとして知らせる' {
    $dir = New-TestDir
    $rows = @(
        @('1', 'SCR002', '異常', '', '顧客CD を未入力で登録する', '', 'E0001 が表示される', 'OK'),
        @('2', '', '異常', '', '受注日付に存在しない日付を入れて登録する', '', 'E0004 が表示される', 'OK'),
        @('3', '', '異常', '', '顧客CD に全角を入れて登録する', '', 'E0003 が表示される', 'OK'),
        @('4', '', '正常', '', '顧客コードを入れて顧客名を表示する', '', '顧客名が表示される', 'OK'))
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') $rows
    $issues = @(& $TestSpecTool -Path $path -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') 6>$null | Where-Object Item -eq '項目名の揺れ')
    Assert-Equal @(
        'テストケース!A2|SCR002 のケースに「顧客CD」と書かれています。設計書の項目名は「顧客コード」です（DD02_画面設計書_SCR002_受注登録.xlsx 画面項目定義 7 行目。ほかに 1 件）',
        'テストケース!A3|SCR002 のケースに「受注日付」と書かれています。設計書の項目名は「受注日」です（DD02_画面設計書_SCR002_受注登録.xlsx 画面項目定義 6 行目）') @($issues | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Detail }) '設計書どおりの「顧客コード」「顧客名」は揺れにしない'
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

Test-Case 'Test-TestSpec: -DesignPath の必須の印（●・◎・✓）と、桁数の「全角20」「40バイト」「7,2」（全体の 7 桁）からも、必須・境界値のケースの抜けを知らせる' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    $hdr = @('画面ID', 'SCR010', '画面名', '顧客登録', '', '')
    New-SheetsBook (Join-Path $design '画面設計書_SCR010.xlsx') ([ordered]@{ '画面項目定義' = @($hdr, @(), @('No', '項目名', '入出力', '型', '桁数', '必須'),
                @('1', '顧客名', 'I', '全角', '全角20', '●'), @('2', '顧客名カナ', 'I', '全角カナ', '40バイト', '◎'), @('3', '単価', 'I', '数値', '7,2', '✓')) })
    $path = Join-Path $dir 'UT_顧客登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') @(, @('1', 'SCR010', '正常', '', '顧客を登録する', '', '登録される', 'OK'))
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null | Where-Object { $_.Item -match '^(必須|境界値)' })
    Assert-Equal @(
        '必須のテストが無い|SCR010 の必須の入力項目「顧客名」（画面設計書_SCR010.xlsx 画面項目定義 4 行目）を未入力にするケースが見つかりません',
        '境界値のテストが無い|SCR010 の入力項目「顧客名」（20 桁。画面設計書_SCR010.xlsx 画面項目定義 4 行目）の桁数の境界（20 桁・21 桁）を確かめるケースが見つかりません',
        '必須のテストが無い|SCR010 の必須の入力項目「顧客名カナ」（画面設計書_SCR010.xlsx 画面項目定義 5 行目）を未入力にするケースが見つかりません',
        '境界値のテストが無い|SCR010 の入力項目「顧客名カナ」（40 桁。画面設計書_SCR010.xlsx 画面項目定義 5 行目）の桁数の境界（40 桁・41 桁）を確かめるケースが見つかりません',
        '必須のテストが無い|SCR010 の必須の入力項目「単価」（画面設計書_SCR010.xlsx 画面項目定義 6 行目）を未入力にするケースが見つかりません',
        '境界値のテストが無い|SCR010 の入力項目「単価」（7 桁。画面設計書_SCR010.xlsx 画面項目定義 6 行目）の桁数の境界（7 桁・8 桁）を確かめるケースが見つかりません') @($issues | ForEach-Object { '{0}|{1}' -f $_.Item, $_.Detail })
}

Test-Case 'Test-TestSpec: -DesignPath の画面ID が全角（ＳＣＲ０１０）でもケースの画面と突き合わせ、半角で書いたシートと同じ画面として項目名の揺れを見る' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    $columns = @('No', '項目名', '入出力', '型', '桁数', '必須')
    New-SheetsBook (Join-Path $design '画面設計書_SCR010.xlsx') ([ordered]@{
            '画面項目定義' = @(@('画面ID', 'ＳＣＲ０１０', '画面名', '顧客登録', '', ''), @(), $columns, @('1', '顧客名', 'I', '全角', '20', '○'))
            '画面項目定義2' = @(@('画面ID', 'SCR010', '画面名', '顧客登録', '', ''), @(), $columns, @('2', '顧客名称', 'I', '全角', '40', ''))
        })
    $path = Join-Path $dir 'UT_顧客登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '前提条件', '手順', '入力値', '期待結果', '結果') @(
        @('1', 'SCR010', '正常', '', '顧客名を入れて登録する', '', '登録される', 'OK'),
        @('2', '', '正常', '', '顧客名称を入れて登録する', '', '登録される', 'OK'))
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null)
    Assert-Equal @(
        '必須のテストが無い|ＳＣＲ０１０ の必須の入力項目「顧客名」（画面設計書_SCR010.xlsx 画面項目定義 4 行目）を未入力にするケースが見つかりません',
        '境界値のテストが無い|ＳＣＲ０１０ の入力項目「顧客名」（20 桁。画面設計書_SCR010.xlsx 画面項目定義 4 行目）の桁数の境界（20 桁・21 桁）を確かめるケースが見つかりません',
        '境界値のテストが無い|SCR010 の入力項目「顧客名称」（40 桁。画面設計書_SCR010.xlsx 画面項目定義2 4 行目）の桁数の境界（40 桁・41 桁）を確かめるケースが見つかりません') @($issues | ForEach-Object { '{0}|{1}' -f $_.Item, $_.Detail }) '同じ画面の「顧客名」「顧客名称」は揺れにしない'
}

Test-Case 'Test-TestSpec: 区分として読んだ列の値が正常・異常などでない（観点の列の「必須チェック」など）か空欄の機能は、期待結果の言葉から異常系を推定する' {
    $dir = New-TestDir
    $rows = @(
        @('1', 'SCR002', '必須チェック', '受注日を空欄にして登録する', 'E0001 が表示される'),
        @('2', '', '必須チェック', '顧客コードを空欄にして登録する', 'E0001 が表示される'),
        @('3', '', '桁数チェック', '備考に 501 文字を入れて登録する', 'E0002 が表示される'),
        @('4', '', '型チェック', '顧客コードに全角を入れて登録する', 'E0003 が表示される'),
        @('5', '', '正常登録', '全項目を入れて登録する', 'I0001 が表示され、受注が登録される'),
        @('6', 'SCR003', '', '備考に 501 文字を入れて更新する', 'E0002 が表示される'),
        @('7', '', '', '希望納品日に存在しない日付を入れて更新する', 'E0004 が表示される'),
        @('8', '', '', '他のユーザーが更新した後で更新する', 'E0006 が表示される'),
        @('9', '', '', '配送時間帯を変えて更新する', 'I0002 が表示される'),
        @('10', '', '', 'キャンセルボタンを押す', 'Q0002 が表示される'))
    $path = Join-Path $dir 'UT_受注.xlsx'
    New-SpecBook $path @('No', '機能', '観点', '手順', '期待結果') $rows
    $out = Join-Path $dir 'spec-check.xlsx'
    $issues = @(& $TestSpecTool -Path $path -OutXlsx $out 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    $functions = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('SCR002|5|4|期待結果の言葉から推定', 'SCR003|5|3|期待結果の言葉から推定') @(2, 3 | ForEach-Object { '{0}|{1}|{2}|{3}' -f $functions.Rows[$_]['A'], $functions.Rows[$_]['B'], $functions.Rows[$_]['C'], $functions.Rows[$_]['E'] })
}

Test-Case 'Test-TestSpec: -DesignPath の型・境界値のケースは確認内容・期待結果・手順などの列から探し、テストID や予定日の数字を桁数・日付と読まない' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    New-SheetsBook (Join-Path $design '画面設計書_SCR010.xlsx') ([ordered]@{ '画面項目定義' = @(@('画面ID', 'SCR010', '画面名', '顧客登録', '', ''), @(), @('No', '項目名', '入出力', '型', '桁数', '必須'),
                @('1', '登録日', 'I', '日付', '10', ''), @('2', '顧客コード', 'I', '半角英数', '10', '')) })
    $path = Join-Path $dir 'UT_顧客登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '手順', '期待結果', '予定日') @(, @('10', 'SCR010', '正常', '登録日と顧客コードを入れて登録する', '登録される', '2026/12/30'))
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null | Where-Object { $_.Item -match '^(型|境界値)' })
    Assert-Equal @(
        '型のテストが無い|SCR010 の入力項目「登録日」（型 日付。画面設計書_SCR010.xlsx 画面項目定義 4 行目）に、日付として正しくない値を入れるケースが見つかりません',
        '型のテストが無い|SCR010 の入力項目「顧客コード」（型 半角英数。画面設計書_SCR010.xlsx 画面項目定義 5 行目）に、半角として正しくない値を入れるケースが見つかりません',
        '境界値のテストが無い|SCR010 の入力項目「顧客コード」（10 桁。画面設計書_SCR010.xlsx 画面項目定義 5 行目）の桁数の境界（10 桁・11 桁）を確かめるケースが見つかりません') @($issues | ForEach-Object { '{0}|{1}' -f $_.Item, $_.Detail }) 'No の 10、予定日の 12/30（2/30）'
}

Test-Case 'Test-TestSpec: -DesignPath のメッセージの文言は、セル内の改行の違いを見ない' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    New-SheetsBook (Join-Path $design 'メッセージ一覧.xlsx') ([ordered]@{ 'メッセージ一覧' = @(@('メッセージID', 'メッセージ'), @('E0001', "{0}は必須入力です。`n入力してください。"), @('E0002', '{0}は{1}文字以内で入力してください。')) })
    $rows = @(
        @('1', 'SCR002', '異常', '受注日を未入力で登録する', 'E0001「受注日は必須入力です。入力してください。」が表示される'),
        @('2', '', '異常', '備考に 501 文字を入れて登録する', "E0002「備考は500文字以内で`n入力してください。」が表示される"),
        @('3', '', '異常', '数量を未入力で登録する', 'E0001「数量は必須入力です。」が表示される'))
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '区分', '手順', '期待結果') $rows
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null | Where-Object { $_.Item -match 'メッセージ' })
    Assert-Equal @('メッセージの文言が違う|テストケース!E4|「数量は必須入力です。」が、メッセージ一覧の E0001「{0}は必須入力です。↵入力してください。」（メッセージ一覧.xlsx メッセージ一覧 2 行目） と合いません') @($issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Item, $_.Location, $_.Detail })
}

Test-Case 'Test-TestSpec: -DesignPath のメッセージの指摘は、ケースの列を項目の順（確認内容・期待結果・手順・前提条件・入力値）に見て並べる（実行ごとに変わらない）' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $null = New-Item -ItemType Directory -Path $design
    New-SheetsBook (Join-Path $design 'メッセージ一覧.xlsx') ([ordered]@{ 'メッセージ一覧' = @(@('メッセージID', 'メッセージ'), @('E0001', '{0}は必須入力です。'), @('E0006', '他のユーザーにより更新されています。再度検索してから操作してください。')) })
    $path = Join-Path $dir 'UT_受注登録.xlsx'
    New-SpecBook $path @('No', '機能', '確認内容', '期待結果', '手順', '前提条件', '入力値') @(, @('1', 'SCR002', '受注日の必須チェック（E9991）', 'E0001「受注日を入力してください。」が表示される（E9992）', '受注日を空にして登録する（E9993）', 'E0006「他のユーザーにより更新されています。」が出た状態（E9994）', '空欄（E9995）'))
    $issues = @(& $TestSpecTool -Path $path -DesignPath $design 6>$null | Where-Object { $_.Item -match 'メッセージ' })
    $defined = 'E0001「{0}は必須入力です。」（メッセージ一覧.xlsx メッセージ一覧 2 行目）・E0006「他のユーザーにより更新されています。再度検索してから操作してください。」（メッセージ一覧.xlsx メッセージ一覧 3 行目）'
    Assert-Equal @(@(1..5 | ForEach-Object { 'メッセージIDが一覧に無い|メッセージ ID E999{0} がメッセージ一覧にありません（誤記か、一覧への追加漏れ）' -f $_ }) + @(
            ('メッセージの文言が違う|「受注日を入力してください。」が、メッセージ一覧の {0} と合いません' -f $defined),
            ('メッセージの文言が違う|「他のユーザーにより更新されています。」が、メッセージ一覧の {0} と合いません' -f $defined))) @($issues | ForEach-Object { '{0}|{1}' -f $_.Item, $_.Detail })
}

Test-Case 'Test-TestSpec: Get-Help でヘルプが表示される' {
    $h = Get-Help $TestSpecTool
    Assert-Match $h.Synopsis 'テスト仕様書'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
