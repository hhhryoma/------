#     Test-DesignConsistency.ps1 のテスト。設計書は docs\design-sample と tests\fixtures\consistency（make_fixtures.py で生成）

$ConsistencyTool = Join-Path $script:ToolsDir 'Test-DesignConsistency.ps1'
$consistencyDir = Join-Path $script:Fixtures 'consistency'
$sampleDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Invoke-Consistency {
    param([hashtable] $Options)
    return , @(& $ConsistencyTool @Options 6>$null 3>$null)
}

function Get-ConsistencyKeys {
    param($Issues)
    return @($Issues | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Level, $_.Item, $_.File, $_.Location } | Sort-Object)
}

Test-Case 'Test-DesignConsistency: 参照が整合している詳細設計サンプルでは指摘しない' {
    $issues = Invoke-Consistency @{ Path = $sampleDir }
    Assert-Equal 0 $issues.Count ((Get-ConsistencyKeys $issues) -join ', ')
}

Test-Case 'Test-DesignConsistency: 参照先なし・重複定義・一覧との違い・ヘッダの違い・組み合わせの違い・表記の違いを見つける' {
    $issues = Invoke-Consistency @{ Path = $consistencyDir }
    $scr1 = 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $scr2 = 'DD02_画面設計書_SCR002_受注登録.xlsx'
    $scr3 = 'DD02_画面設計書_SCR003_受注詳細.xlsx'
    $scr4 = 'DD02_画面設計書_SCR004_出荷状況照会.xlsx'
    Assert-Equal (@(
        "エラー|参照先なし|$scr1|画面項目定義!AS6",
        "エラー|参照先なし|$scr1|画面項目定義!AK9",
        "エラー|参照先なし|$scr2|画面項目定義!AS11",
        "エラー|参照先なし|$scr2|画面項目定義!BA9",
        "エラー|参照先なし|$scr3|イベント定義!AM7",
        "エラー|重複定義|$scr4|画面概要!S2",
        'エラー|重複定義|DD05_メッセージ一覧.xlsx|メッセージ一覧!A22',
        '警告|一覧との違い|DD01_画面一覧.xlsx|画面一覧!C8',
        '警告|一覧との違い|DD01_画面一覧.xlsx|画面一覧!C9',
        '警告|一覧との違い|DD03_テーブル定義書.xlsx|T_ORDER_DETAIL!S2',
        "警告|ヘッダの違い|$scr3|画面項目定義!S2",
        "警告|ヘッダの違い|$scr4|画面概要!S2",
        "警告|組み合わせの違い|$scr1|イベント定義!BB10",
        '警告|組み合わせの違い|DD01_画面一覧.xlsx|画面遷移!Z12',
        '警告|組み合わせの違い|DD01_画面一覧.xlsx|画面遷移!Z12',
        "情報|表記の違い|$scr2|画面項目定義!AS6",
        '情報|表記の違い|DD06_IF定義書_IF001_出荷指示データ.xlsx|レコード定義!AG6') | Sort-Object) (Get-ConsistencyKeys $issues)

    $details = @($issues | ForEach-Object { $_.Detail })
    foreach ($expected in @(
            'テーブルID「T_ORDERX」が定義されていません（T_ORDERX.ORDER_NO）',
            'メッセージ「E0099」が定義されていません',
            'カラム「T_ORDER.DELIVERY_TIME」が定義されていません（対応項目）',
            'コード「C009」が定義されていません',
            'テーブル「T_ORDERS」が定義されていません（使用テーブル）',
            "画面「SCR001」は、ほかの場所でも定義されています（$scr1 画面概要!S2）",
            'メッセージ「E0001」は、ほかの場所でも定義されています（メッセージ一覧!A6）',
            "画面「SCR003」の名前が違います（一覧: 受注照会 / ${scr3}: 受注詳細）",
            '一覧にある画面「SCR004 出荷状況照会」の定義が見つかりません（設計書が無いか、ID が違います）',
            'テーブル「T_ORDER_DETAIL 受注明細」が一覧（DD03_テーブル定義書.xlsx（テーブル一覧））にありません',
            '画面IDが「SCR002」です（このファイルのほかのシートは「SCR003」）',
            'ファイル名は「SCR004」ですが、画面IDは「SCR001」です',
            'SCR001 の遷移先「SCR003」が、DD01_画面一覧.xlsx（画面遷移）の遷移先画面IDにありません',
            "SCR003 の遷移先画面ID「SCR002」が、${scr3}（イベント定義）の遷移先にありません",
            "SCR003 の遷移先画面ID「SCR002」が、${scr3}（画面概要）の遷移先画面IDにありません",
            '「T_ORDER.order_date」は、定義では「T_ORDER.ORDER_DATE」と書かれています（全角・半角や大文字・小文字）',
            '「Ｔ_ORDER.ORDER_NO」は、定義では「T_ORDER.ORDER_NO」と書かれています（全角・半角や大文字・小文字）')) {
        Assert-True ($details -contains $expected) "指摘がありません: $expected"
    }
    # 指摘の Kind は定義・参照の種類（Invoke-DesignReview の「参照先なし（メッセージ）」などで使う）
    Assert-Equal @('カラム', 'メッセージ', 'コード', 'カラム', 'テーブル') @($issues | Where-Object Item -eq '参照先なし' | ForEach-Object { $_.Kind })
    Assert-Equal @('画面', '画面', 'テーブル') @($issues | Where-Object Item -eq '一覧との違い' | ForEach-Object { $_.Kind })
}

Test-Case 'Test-DesignConsistency: 属性の違い（桁数・型・必須と既定値）と、どこからも参照されていない定義を見つける' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    $tables = Join-Path $dir 'DD03_テーブル定義書.xlsx'
    # 配送時間帯の既定値を消す・備考を 200 桁にする・顧客コードを任意にする・顧客マスタの顧客コードの型を変える
    Set-CellValues $tables 'T_ORDER' @{ AG19 = $null; X21 = 200; AE15 = $null }
    Set-CellValues $tables 'M_CUSTOMER' @{ S13 = 'VARCHAR' }
    Set-CellValues (Join-Path $dir 'DD05_メッセージ一覧.xlsx') 'メッセージ一覧' @{ A22 = 'E0099'; G22 = 'エラー'; K22 = 'どこからも使わないメッセージです。' }
    $issues = Invoke-Consistency @{ Path = $dir }
    $got = @($issues | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Level, $_.Item, $_.File, $_.Location, $_.Detail })
    foreach ($expected in @(
            '警告|属性の違い|DD02_画面設計書_SCR002_受注登録.xlsx|画面項目定義!AS11|入力が任意ですが、T_ORDER.DELIVERY_TIME_KBN は必須で既定値がありません（未入力では登録できません）',
            '警告|属性の違い|DD02_画面設計書_SCR002_受注登録.xlsx|画面項目定義!AS12|桁数が 500 ですが、T_ORDER.REMARKS は 200 です（入力した値がカラムに入りきりません）',
            '警告|属性の違い|DD06_IF定義書_IF001_出荷指示データ.xlsx|レコード定義!AG9|必須の項目ですが、T_ORDER.CUSTOMER_CD は必須ではありません（値が無いことがあります）',
            '警告|属性の違い|DD03_テーブル定義書.xlsx|T_ORDER!AO15|型が「CHAR」ですが、M_CUSTOMER.CUSTOMER_CD は「VARCHAR」です',
            '情報|未使用|DD05_メッセージ一覧.xlsx|メッセージ一覧!A22|メッセージ「E0099 どこからも使わないメッセージです。」は、どこからも参照されていません')) {
        Assert-True ($got -contains $expected) ("指摘がありません: {0}`n実際: {1}" -f $expected, ($got -join "`n"))
    }
    # 検索画面（SCR001 は受注を参照するだけ）の任意の検索条件は、必須のカラムと比べない
    Assert-Equal 0 @($issues | Where-Object { $_.File -like '*SCR001*' -and $_.Detail -like '入力が任意*' }).Count
    Assert-Equal 1 @($issues | Where-Object Item -eq '未使用').Count
}

Test-Case 'Test-DesignConsistency: -OutXlsx に CRUD 図を足し、-Where で ID を使っている場所（影響範囲）を一覧にする' {
    $dir = New-TestDir
    $out = Join-Path $dir 'check.xlsx'
    $null = Invoke-Consistency @{ Path = $sampleDir; OutXlsx = $out }
    $crud = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('指摘一覧', '定義と参照', 'CRUD') $crud.Sheets
    Assert-Equal '画面・IF|名前|M_CUSTOMER（顧客マスタ）|M_PRODUCT（商品マスタ）|T_ORDER（受注）|T_ORDER_DETAIL（受注明細）' (@('A', 'B', 'C', 'D', 'E', 'F' | ForEach-Object { $crud.Rows[1][$_] }) -join '|')
    Assert-Equal @('IF001|出荷指示データ|R||R|R', 'SCR001|受注一覧|R||R|', 'SCR002|受注登録|R|RU|C|C', 'SCR003|受注詳細|R|RU|RU|R') @(2..5 | ForEach-Object { $r = $_; (@('A', 'B', 'C', 'D', 'E', 'F' | ForEach-Object { $crud.Rows[$r][$_] }) -join '|') }) 'イベント定義の使用テーブルと、IF の設定元（読み取り）'

    $usages = Invoke-Consistency @{ Path = $sampleDir; Where = @('T_ORDER.DELIVERY_TIME_KBN', 'C005') }
    Assert-Equal @(
        'C005|定義|DD04_コード定義書.xlsx|コード一覧!A16',
        'C005|参照|DD02_画面設計書_SCR002_受注登録.xlsx|画面項目定義!BA11',
        'C005|参照|DD02_画面設計書_SCR003_受注詳細.xlsx|画面項目定義!BA13',
        'C005|参照|DD03_テーブル定義書.xlsx|T_ORDER!AX19',
        'C005|参照|DD06_IF定義書_IF001_出荷指示データ.xlsx|レコード定義!AR17',
        'T_ORDER.DELIVERY_TIME_KBN|定義|DD03_テーブル定義書.xlsx|T_ORDER!J19',
        'T_ORDER.DELIVERY_TIME_KBN|参照|DD02_画面設計書_SCR002_受注登録.xlsx|画面項目定義!AS11',
        'T_ORDER.DELIVERY_TIME_KBN|参照|DD02_画面設計書_SCR003_受注詳細.xlsx|画面項目定義!AS13',
        'T_ORDER.DELIVERY_TIME_KBN|参照|DD06_IF定義書_IF001_出荷指示データ.xlsx|レコード定義!AG17') @($usages | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Target, $_.Usage, $_.File, $_.Location })
    # テーブルだけを指定すると、そのカラムの参照も含める
    $usages = Invoke-Consistency @{ Path = $sampleDir; Where = 'M_PRODUCT' }
    Assert-True (@($usages | Where-Object { $_.Text -eq 'M_PRODUCT.PRODUCT_NAME' }).Count -ge 1) 'カラムの参照'
    Assert-True (@($usages | Where-Object Usage -eq '一覧').Count -eq 1) 'テーブル一覧'

    $bad = Join-Path $dir 'bad.csv'
    @('区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複,属性,未使用,CRUD', '定義,カラム,,,物理名,物理名,,,,,,長さ=桁数,,') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $sampleDir -Rules $bad 6>$null } '属性'
    @('区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複,属性,未使用,CRUD', '参照,テーブル,,,,,,,@画面ID,(?<id>T_\w+),,,,X') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $sampleDir -Rules $bad 6>$null } 'CRUD'
}

Test-Case 'Test-DesignConsistency: -Exclude で確かめない参照を除き、定義の無い種類は「未確認」として知らせる' {
    $file = Join-Path $consistencyDir 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $issues = Invoke-Consistency @{ Path = $file; Exclude = '^T_ORDERX\.' }
    $unchecked = @($issues | Where-Object Item -eq '未確認' | ForEach-Object { $_.Detail })
    Assert-True ($unchecked -contains 'カラムの定義が見つからないため、カラムの参照（11 か所）は確かめていません') ($unchecked -join ' / ')
    Assert-True ($unchecked -contains 'メッセージの定義が見つからないため、メッセージの参照（15 か所）は確かめていません') ($unchecked -join ' / ')
    Assert-Equal '' (@($issues | Where-Object Item -eq '未確認')[0].File) '場所のない指摘'
    # 画面は 1 つ（SCR001）だけ定義されているので、遷移先の SCR002・SCR003 は参照先なし
    Assert-Equal 4 @($issues | Where-Object Item -eq '参照先なし').Count
}

Test-Case 'Test-DesignConsistency: 参照のルールを差し替えられる（区分の ID を最初の行にだけ書く表、重複を許す定義）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    Export-Workbook -FilePath (Join-Path $dir 'コード定義書.xlsx') -Sheets @(@{
        Name = 'コード一覧'
        Columns = @(@{ Header = 'コード区分ID'; Width = 10 }, @{ Header = 'コード区分名'; Width = 10 }, @{ Header = 'コード値'; Width = 10 }, @{ Header = '表示名'; Width = 10 })
        Rows = @(@{ Values = @('C001', '状態', '10', '受付') }, @{ Values = @('', '', '20', '出荷') }, @{ Values = @('C002', '支払', '1', '振込') })
    })
    Export-Workbook -FilePath (Join-Path $dir '画面設計書_S01.xlsx') -Sheets @(@{
        Name = '項目'; Columns = @(@{ Header = '備考'; Width = 30 }); Rows = @(@{ Values = @('C001-20 を表示する') }, @{ Values = @('C001-30 は表示しない') })
    })
    $rules = Join-Path $dir 'rules.csv'
    @(
        '区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複',
        '定義,コード,コード定義書,,コード区分ID/コード値,コード区分ID,コード区分名,,,,可',
        '定義,コード値,コード定義書,,コード区分ID/コード値,コード値,表示名,コード区分ID,,,',
        '参照,コード値,画面設計書,,,,,,,(?<parent>C\d{3})-(?<id>\d+),'
    ) | Set-Content -LiteralPath $rules -Encoding UTF8
    $issues = Invoke-Consistency @{ Path = $dir; Rules = $rules }
    Assert-Equal 'エラー|参照先なし|画面設計書_S01.xlsx|項目!A3' ((Get-ConsistencyKeys $issues) -join ', ') 'C001-20 は 2 行目（区分の ID が空）の定義'
    Assert-Equal 'コード値「C001.30」が定義されていません' $issues[0].Detail

    $bad = Join-Path $dir 'bad.csv'
    @('区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複', '参照,メッセージ,,,,,,,,[EW]\d{4},') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $dir -Rules $bad 6>$null } '\(\?<id>'
    @('区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複', '索引,メッセージ,,,,,,,,,') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $dir -Rules $bad 6>$null } '定義 / 一覧 / 参照'
}

Test-Case 'Test-DesignConsistency: -OutXlsx で指摘一覧（該当セルへのリンク付き）と種類ごとの数を保存し、設計書は変更しない' {
    $dir = Copy-Fixture 'consistency'
    $file = Join-Path $dir 'DD05_メッセージ一覧.xlsx'
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $out = Join-Path $dir 'check.xlsx'
    $null = Invoke-Consistency @{ Path = $dir; OutXlsx = $out }
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal 18 $sheet.Rows.Count '見出しと 17 件（保存先のファイル自体は読まない）'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'メッセージ一覧'!A22" }).Count -eq 1) '重複定義のセルへのリンク'
    $stats = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('指摘一覧', '定義と参照', 'CRUD') $stats.Sheets
    $row = @($stats.Rows.Values | Where-Object { $_['A'] -eq 'メッセージ' })[0]
    Assert-Equal '16|1' ('{0}|{1}' -f $row['B'], $row['E']) 'メッセージの定義 16・参照先なし 1'
}

Test-Case 'Test-DesignConsistency: 画面遷移の遷移元が空の行（縦に結合したセル）は上の行の値を使い、1 つのセルに書いた複数の遷移元は ID ごとに比べる' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    # SCR001 の 2 行目の遷移元を空にする（縦に結合したセルは、左上のセルにだけ値がある）。SCR002 から SCR001 への行の遷移元に SCR003 も書く
    Set-CellValues (Join-Path $dir 'DD01_画面一覧.xlsx') '画面遷移' @{ C7 = $null; I7 = $null; C9 = 'SCR002、SCR003'; I9 = '受注登録・受注詳細' }
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal 0 $issues.Count ((Get-ConsistencyKeys $issues) -join ', ')

    # CRUD 図も同じ（元の列が空の行は上の行の値、複数の ID は ID ごとの行）
    $dir = New-TestDir
    $book = Join-Path $dir 'CRUD表.xlsx'
    New-SheetsBook $book ([ordered]@{ 'CRUD' = @(@('画面ID', '使用テーブル'), @('S01、S02', 'T_A'), @('', 'T_B')) })
    $rules = Join-Path $dir 'rules.csv'
    @('区分,種類,ファイル名,シート名,見出し,列,名前,親,元,書き方,重複,属性,未使用,CRUD', '参照,テーブル,CRUD表,,画面ID/使用テーブル,使用テーブル,,,画面ID,(?<id>T_\w+),,,,R') | Set-Content -LiteralPath $rules -Encoding UTF8
    $out = Join-Path $dir 'check.xlsx'
    $null = Invoke-Consistency @{ Path = $book; Rules = $rules; OutXlsx = $out }
    $crud = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('S01||R|R', 'S02||R|R') @(2, 3 | ForEach-Object { $r = $_; (@('A', 'B', 'C', 'D' | ForEach-Object { $crud.Rows[$r][$_] }) -join '|') })
}

Test-Case 'Test-DesignConsistency: 「コード定義 C001、C002」の 2 つ目からのコードも読み、コード区分 ID を桁数で切らない（C0001・C0010）' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir 'コード定義書.xlsx') ([ordered]@{ 'コード一覧' = @(@('コード区分ID', 'コード区分名', 'コード値', '表示名'), @('C001', '受注ステータス', '10', '受付'), @('C002', '支払方法', '1', '銀行振込'), @('C0001', '配送区分', '1', '通常')) })
    New-SheetsBook (Join-Path $dir '画面設計書_SCR001.xlsx') ([ordered]@{ '画面項目定義' = @(@('項目名', '備考'), @('状態・支払', 'コード定義 C001、C002'), @('配送', 'コード定義 C0001'), @('区分', 'コード定義 C0010')) })
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal 'エラー|参照先なし|画面設計書_SCR001.xlsx|画面項目定義!B4' ((Get-ConsistencyKeys $issues) -join ', ') 'C002・C0001 は未使用ではなく、C0010 は C001 ではない'
    Assert-Equal 'コード「C0010」が定義されていません' $issues[0].Detail
    $usages = Invoke-Consistency @{ Path = $dir; Where = 'C002' }
    Assert-Equal @('定義|コード一覧!A3', '参照|画面項目定義!B2') @($usages | ForEach-Object { '{0}|{1}' -f $_.Usage, $_.Location })
}

Test-Case 'Test-DesignConsistency: ファイル名に同じ長さの ID が 2 つあれば、左に書かれたほうとヘッダの ID を比べる' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir '画面設計書_SCR001_SCR002.xlsx') ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', 'SCR002')) })
    New-SheetsBook (Join-Path $dir '画面設計書_SCR002_SCR001.xlsx') ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', 'SCR001')) })
    # ハッシュテーブルの順（PowerShell 7 では実行ごとに変わる）で決めると、どちらか一方しか知らせない
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal @('画面設計書_SCR001_SCR002.xlsx|ファイル名は「SCR001」ですが、画面IDは「SCR002」です', '画面設計書_SCR002_SCR001.xlsx|ファイル名は「SCR002」ですが、画面IDは「SCR001」です') @($issues | ForEach-Object { '{0}|{1}' -f $_.File, $_.Detail })
}

Test-Case 'Test-DesignConsistency: カラムの必須の「要」などの印と、既定値の「－」「なし」を、テーブル定義書から DDL を作るとき（Export-TableDdl）と同じに読む' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    # 顧客コードの必須を「要」（DDL では NOT NULL）に、配送時間帯の既定値を「－」（既定値なし）にする
    Set-CellValues (Join-Path $dir 'DD03_テーブル定義書.xlsx') 'T_ORDER' @{ AE15 = '要'; AG19 = '－' }
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal (@(
        '警告|属性の違い|DD02_画面設計書_SCR002_受注登録.xlsx|画面項目定義!AS11',
        '警告|属性の違い|DD02_画面設計書_SCR003_受注詳細.xlsx|画面項目定義!AS13') | Sort-Object) (Get-ConsistencyKeys $issues)
    Assert-Equal 2 @($issues | Where-Object Detail -eq '入力が任意ですが、T_ORDER.DELIVERY_TIME_KBN は必須で既定値がありません（未入力では登録できません）').Count
}

Test-Case 'Test-DesignConsistency: Get-Help でヘルプが表示される' {
    $h = Get-Help $ConsistencyTool
    Assert-Match $h.Synopsis '参照'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'Test-DesignConsistency: 版を上げたあとの古い版、雛形のシート（非表示・差し込み欄のまま）、設計書のフォルダに置いたツールの出力・差分の印の写し・書き換える前のファイルを読まない' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    # テーブル定義書の版を上げ（前の版のファイルも残る）、新しい版で備考を 1000 桁にする
    $old = Join-Path $dir 'DD03_テーブル定義書_v1.0.xlsx'
    $new = Join-Path $dir 'DD03_テーブル定義書_v1.1.xlsx'
    Move-Item -LiteralPath (Join-Path $dir 'DD03_テーブル定義書.xlsx') -Destination $old
    Copy-Item -LiteralPath $old -Destination $new
    Set-CellValues $new 'T_ORDER' @{ X21 = 1000 }
    # 非表示のシートのテーブルと、テーブル ID が差し込み欄のままの雛形のシート
    $extra = Join-Path $dir 'DD03_テーブル定義書_追加.xlsx'
    $columns = @(@('論理名', '物理名', 'データ型'), @('番号', 'NO', 'CHAR'))
    New-SheetsBook $extra ([ordered]@{
            'T_EXTRA' = @(, @('テーブルID', 'T_EXTRA', 'テーブル名', '追加')) + $columns
            '雛形'    = @(, @('テーブルID', '{{テーブルID}}', 'テーブル名', '{{テーブル名}}')) + $columns
        })
    Edit-ZipText $extra 'xl/workbook.xml' { param($t) $t -replace '(<sheet name="T_EXTRA")', '$1 state="hidden"' }
    # 設計書のフォルダに置いたツールの出力と、差分の印を付けた写し・書き換える前のファイル
    $null = & $ConsistencyTool -Path $dir -OutXlsx (Join-Path $dir '整合性.xlsx') 6>$null 3>$null
    $null = & (Join-Path $script:ToolsDir 'Compare-Excel.ps1') $old $new -OutMarkedDir (Join-Path $dir '差分の印') 6>$null 3>$null
    Assert-True (Test-Path -LiteralPath (Join-Path $dir '差分の印/DD03_テーブル定義書_v1.1.xlsx')) '差分の印を付けた写し'
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -Pattern '受注管理システム' -Replacement '受注管理システム（改）' -SimpleMatch -Path (Join-Path $dir 'DD05_メッセージ一覧.xlsx') -Force -BackupDir (Join-Path $dir '置換前') 6>$null 3>$null
    Assert-True (Test-Path -LiteralPath (Join-Path $dir '置換前/DD05_メッセージ一覧.xlsx')) '書き換える前のファイル'

    $all = @(& $ConsistencyTool -Path $dir 6>&1 3>&1)
    $issues = @($all | Where-Object { $_ -isnot [System.Management.Automation.InformationRecord] -and $_ -isnot [System.Management.Automation.WarningRecord] })
    Assert-Match (($all | Where-Object { $_ -is [System.Management.Automation.InformationRecord] } | ForEach-Object { "$_" }) -join "`n") '8 ファイルを読みました（定義: テーブル 4・'
    Assert-Equal 0 @($issues | Where-Object Level -eq 'エラー').Count ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')
    Assert-Equal @('DD02_画面設計書_SCR002_受注登録.xlsx|桁数が 500 ですが、T_ORDER.REMARKS は 1000 です', 'DD02_画面設計書_SCR003_受注詳細.xlsx|桁数が 500 ですが、T_ORDER.REMARKS は 1000 です') `
        @($issues | Where-Object Item -eq '属性の違い' | ForEach-Object { '{0}|{1}' -f $_.File, $_.Detail } | Sort-Object) '新しい版のテーブル定義書と比べる'
    Assert-Equal 0 @($issues | Where-Object { $_.Detail -match 'T_EXTRA|\{\{' }).Count '非表示のシート・雛形のシート'
    $warnings = ($all | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message }) -join "`n"
    Assert-Match $warnings '古い版 1 件は読みませんでした（DD03_テーブル定義書_v1\.0\.xlsx → DD03_テーブル定義書_v1\.1\.xlsx。'
    Assert-Match $warnings '雛形（差し込み欄のまま）のシート・ブック 1 件は読みませんでした（DD03_テーブル定義書_追加\.xlsx'
    Assert-Match $warnings 'ツールが書き出したファイル 2 件は、設計書ではないので読みませんでした（'
    Assert-Match $warnings '整合性\.xlsx（Test-DesignConsistency\.ps1 の出力）'
    Assert-Match $warnings '差分の印[\\/]DD03_テーブル定義書_v1\.1\.xlsx（Compare-Excel\.ps1 の差分の印を付けた写し）'
    Assert-Match $warnings '書き換える前のファイルを残したフォルダの 1 件は読みませんでした（置換前。'
}

function New-BasicDesignBook {
    #     1 つのブックにまとめた基本設計書（機能一覧・画面一覧・テーブル一覧・IF一覧・画面遷移のシート）
    param([string] $Path)
    New-SheetsBook $Path ([ordered]@{
            '機能一覧'     = @(@('機能ID', '機能名', '画面ID'), @('F001', '受注一覧', 'SCR001'), @('F002', '受注登録', 'SCR002'), @('F003', '受注照会', 'SCR003'), @('F004', '受注取消', 'SCR004'), @('F005', '出荷状況', 'SCR009'))
            '画面一覧'     = @(@('画面ID', '画面名'), @('SCR001', '受注一覧'), @('SCR002', '受注登録'), @('SCR003', '受注照会'), @('SCR004', '受注取消'))
            'テーブル一覧' = @(@('テーブルID', 'テーブル名'), @('T_ORDER', '受注'), @('T_ORDER_DETAIL', '受注明細'), @('M_CUSTOMER', '顧客マスタ'), @('M_STAFF', '担当者マスタ'))
            'IF一覧'       = @(@('IF ID', 'IF名'), @('IF001', '出荷指示データ'), @('IF002', '請求データ'))
            '画面遷移'     = @(@('遷移元画面ID', '遷移先画面ID'), @('SCR001', 'SCR002'), @('SCR001', 'SCR003'), @('SCR002', 'SCR003'), @('SCR002', 'SCR001'), @('SCR003', 'SCR001'), @('SCR003', 'SCR004'))
        })
}

Test-Case 'Test-DesignConsistency: 基本設計書の一覧は詳細設計書の一覧と別に比べ（重複定義にしない）、基本設計書と詳細設計書の過不足・名前・画面遷移の違いを見つける' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    New-BasicDesignBook (Join-Path $dir '基本設計書.xlsx')
    $out = Join-Path $dir '整合性.xlsx'
    $text = Get-HostText { $script:basicIssues = @(& $ConsistencyTool -Path $dir -OutXlsx $out 3>$null) }
    $issues = $script:basicIssues
    Assert-Equal (@(
        'エラー|参照先なし|基本設計書.xlsx|機能一覧!C6',
        '警告|基本設計との違い|基本設計書.xlsx|画面一覧!A4',
        '警告|基本設計との違い|基本設計書.xlsx|画面一覧!A5',
        '警告|基本設計との違い|基本設計書.xlsx|テーブル一覧!A5',
        '警告|基本設計との違い|基本設計書.xlsx|IF一覧!A3',
        '警告|基本設計との違い|DD03_テーブル定義書.xlsx|M_PRODUCT!S2',
        '警告|組み合わせの違い|基本設計書.xlsx|画面遷移!B7',
        '警告|組み合わせの違い|基本設計書.xlsx|画面遷移!B7',
        '警告|組み合わせの違い|基本設計書.xlsx|画面遷移!B7') | Sort-Object) (Get-ConsistencyKeys $issues)
    $details = @($issues | ForEach-Object { $_.Detail })
    foreach ($expected in @(
            '画面「SCR009」が定義されていません（画面ID）',
            '画面「SCR003」の名前が違います（基本設計: 受注照会 / DD02_画面設計書_SCR003_受注詳細.xlsx: 受注詳細）',
            '基本設計にある画面「SCR004 受注取消」が、詳細設計にありません（設計書が無いか、ID が違います）',
            '基本設計にあるテーブル「M_STAFF 担当者マスタ」が、詳細設計にありません（設計書が無いか、ID が違います）',
            '基本設計にあるIF「IF002 請求データ」が、詳細設計にありません（設計書が無いか、ID が違います）',
            'テーブル「M_PRODUCT 商品マスタ」が基本設計（基本設計書.xlsx（テーブル一覧））にありません',
            'SCR003 の遷移先画面ID「SCR004」が、DD01_画面一覧.xlsx（画面遷移）の遷移先画面IDにありません',
            'SCR003 の遷移先画面ID「SCR004」が、DD02_画面設計書_SCR003_受注詳細.xlsx（イベント定義）の遷移先にありません')) {
        Assert-True ($details -contains $expected) "指摘がありません: $expected"
    }
    Assert-Match $text 'うち基本設計書 1 ファイル（基本設計書\.xlsx）'
    # CRUD 図は詳細設計書から作る（README の CRUD 図と同じ 4 行）
    $crud = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('IF001', 'SCR001', 'SCR002', 'SCR003') @(2..$crud.Rows.Count | ForEach-Object { $crud.Rows[$_]['A'] })

    # 基本設計書だけなら、これまでどおり基本設計書の中で比べる（詳細設計書が無いので、基本設計との違いは出さない）
    $only = Join-Path (New-TestDir) '基本設計書.xlsx'
    Copy-Item -LiteralPath (Join-Path $dir '基本設計書.xlsx') -Destination $only
    $text = Get-HostText { $script:basicIssues = @(& $ConsistencyTool -Path $only 3>$null) }
    Assert-Equal @('エラー|参照先なし|基本設計書.xlsx|機能一覧!C6') (Get-ConsistencyKeys $script:basicIssues)
    Assert-True ($text -notmatch 'うち基本設計書') '詳細設計書が無ければ、別に比べたとは書かない'
}

Test-Case 'Test-DesignConsistency: 基本設計の画面項目一覧（ID の無い表）と詳細設計の画面項目を、画面ごとに名前で比べ、過不足と桁数・入出力の違いを見つける。ID だけ違う画面も見つける' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    New-SheetsBook (Join-Path $dir '基本設計書.xlsx') ([ordered]@{
            '画面一覧' = @(@('画面ID', '画面名'), @('SCR001', '受注一覧'), @('SCR02', '受注登録'))
            '画面項目' = @(@('画面ID', '画面名', '項目名', '入出力', '桁数', '必須'),
                @('SCR001', '受注一覧', '受注番号', 'I', '10', ''),
                @('', '', '受注日(From)', '入力', '10', ''),
                @('', '', '受注日（To）', 'O', '10', ''),
                @('', '', '担当者', 'I', '10', ''),
                @('', '', '受注番号', 'O', '12', ''),
                @('SCR009', '出荷状況', '出荷番号', 'I', '5', '○'))
        })
    $issues = Invoke-Consistency @{ Path = $dir }
    $items = @($issues | Where-Object { $_.Kind -eq '画面項目' })
    Assert-Equal @(
        '情報|基本設計との違い|DD02_画面設計書_SCR001_受注一覧.xlsx|画面項目定義!I9',
        '警告|基本設計との違い|基本設計書.xlsx|画面項目!C2',
        '警告|基本設計との違い|基本設計書.xlsx|画面項目!C4',
        '警告|基本設計との違い|基本設計書.xlsx|画面項目!C5') (Get-ConsistencyKeys $items) '同じ名前の「受注番号」は、検索条件と一覧の順に比べる（重複にしない）。詳細設計の無い SCR009 の項目は比べない'
    $details = @($items | ForEach-Object { $_.Detail })
    Assert-True ($details -contains '画面項目「SCR001.orderNo」の桁数が違います（基本設計: 桁数 10 / DD02_画面設計書_SCR001_受注一覧.xlsx: 桁数 12）') '桁数'
    Assert-True ($details -contains '画面項目「SCR001.orderDateTo」の入出力が違います（基本設計: 入出力 O / DD02_画面設計書_SCR001_受注一覧.xlsx: 入出力 I）') '入出力（「入力」と I は同じ）'
    Assert-True ($details -contains '基本設計にある画面項目「SCR001.担当者」が、詳細設計にありません（名前が違うか、無くした項目です）')
    Assert-Match ($details -join "`n") '基本設計（基本設計書\.xlsx（画面項目））に無い画面項目が 13 件あります（SCR001: 顧客コード・顧客名・受注ステータス・検索・'
    Assert-True (@($issues | Where-Object { $_.Detail -eq '画面「受注登録」の ID が違います（基本設計: SCR02 / DD02_画面設計書_SCR002_受注登録.xlsx: SCR002）' }).Count -eq 1) 'ID で見つからず、名前が同じ画面は ID の違い'
    Assert-True (@($issues | Where-Object { $_.Detail -match 'SCR02.*詳細設計にありません|SCR002.*基本設計.*にありません' }).Count -eq 0) 'ID の違いは過不足にしない'
}

Test-Case 'Test-DesignConsistency: 一覧ごとのファイル（BD01_・DD01_）、フォルダの名前、-BasicDesign で基本設計書を見分け、基本設計の画面設計書も基本設計として比べる' {
    # BD01_画面一覧.xlsx（基本設計。SCR002 が無く、画面遷移も少ない）と DD01_画面一覧.xlsx（詳細設計）
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    New-SheetsBook (Join-Path $dir 'BD01_画面一覧.xlsx') ([ordered]@{
            '画面一覧' = @(@('画面ID', '画面名'), @('SCR001', '受注一覧'), @('SCR003', '受注詳細'))
            '画面遷移' = @(@('遷移元画面ID', '遷移先画面ID'), @('SCR001', 'SCR003'), @('SCR002', 'SCR003'), @('SCR002', 'SCR001'), @('SCR003', 'SCR001'))
        })
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal (@(
        '警告|基本設計との違い|DD02_画面設計書_SCR002_受注登録.xlsx|画面概要!S2',
        '警告|組み合わせの違い|DD01_画面一覧.xlsx|画面遷移!Z6') | Sort-Object) (Get-ConsistencyKeys $issues)
    Assert-True (@($issues | ForEach-Object { $_.Detail }) -contains '画面「SCR002 受注登録」が基本設計（BD01_画面一覧.xlsx（画面一覧））にありません')

    # 名前で見分けられないフォルダ（上流）: -BasicDesign が無ければ、これまでどおり 1 つの一覧として重複を知らせる（-BasicDesign を勧める）
    $root = New-TestDir
    $detail = Copy-Fixture (Join-Path $sampleDir '.')
    $upstream = Join-Path $root '上流'
    New-Item -ItemType Directory -Path $upstream | Out-Null
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD01_画面一覧.xlsx') -Destination (Join-Path $upstream '画面一覧.xlsx')
    New-SheetsBook (Join-Path $upstream '一覧.xlsx') ([ordered]@{ 'テーブル一覧' = @(@('テーブルID', 'テーブル名'), @('T_ORDER', '受注'), @('T_ORDER_DETAIL', '受注明細'), @('M_CUSTOMER', '顧客'), @('M_PRODUCT', '商品マスタ')) })
    $issues = Invoke-Consistency @{ Path = @($detail, $upstream) }
    Assert-Equal @('重複定義') @($issues | ForEach-Object { $_.Item } | Select-Object -Unique)
    Assert-Match @($issues)[0].Detail '一覧のほかの行にも書かれています.*-BasicDesign で基本設計書を指定してください'
    # -BasicDesign: 名前に「基本設計」の無いファイル（一覧.xlsx）も、基本設計書の一覧のルールで読む
    $issues = Invoke-Consistency @{ Path = $detail; BasicDesign = $upstream }
    Assert-Equal @('警告|基本設計との違い|一覧.xlsx|テーブル一覧!A4') (Get-ConsistencyKeys $issues)
    Assert-Equal 'テーブル「M_CUSTOMER」の名前が違います（基本設計: 顧客 / DD03_テーブル定義書.xlsx: 顧客マスタ）' @($issues)[0].Detail

    # 基本設計の画面設計書（「基本設計」のフォルダ。詳細設計と同じ画面 ID）: 重複定義にせず、基本設計として比べる
    $root = New-TestDir
    $basic = Join-Path $root '01_基本設計'
    $detailDir = Join-Path $root '02_詳細設計'
    New-Item -ItemType Directory -Path $basic, $detailDir | Out-Null
    Copy-Item -Path (Join-Path $sampleDir '*.xlsx') -Destination $detailDir
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD02_画面設計書_SCR001_受注一覧.xlsx') -Destination (Join-Path $basic '画面設計書_SCR001_受注一覧.xlsx')
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD02_画面設計書_SCR002_受注登録.xlsx') -Destination (Join-Path $basic '画面設計書_SCR002_受注登録.xlsx')
    $issues = Invoke-Consistency @{ Path = $root }
    Assert-Equal @('警告|基本設計との違い|02_詳細設計/DD02_画面設計書_SCR003_受注詳細.xlsx|画面概要!S2') @(Get-ConsistencyKeys $issues | ForEach-Object { $_.Replace('\', '/') })
    Assert-Match @($issues)[0].Detail '^画面「SCR003 受注詳細」が基本設計（01_基本設計[\\/]画面設計書_SCR001_受注一覧\.xlsx ほか 1 ファイル）にありません$'
}

Test-Case 'Test-DesignConsistency: 帳票設計書・バッチ設計書を定義として読み、基本設計書の帳票一覧・バッチ一覧・機能一覧の帳票ID・バッチID と比べる' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    New-SheetsBook (Join-Path $dir 'DD07_帳票設計書_RPT001_受注伝票.xlsx') ([ordered]@{ '帳票概要' = @(@('帳票ID', 'RPT001', '帳票名', '受注伝票'), @('出力', 'PDF')) })
    New-SheetsBook (Join-Path $dir 'DD08_バッチ設計書_BAT001_日次集計.xlsx') ([ordered]@{ 'バッチ概要' = @(@('バッチID', 'BAT001', 'バッチ名', '日次集計'), @('起動', '毎日')) })
    # 詳細設計書だけなら、帳票・バッチの設計書があっても指摘しない
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal 0 $issues.Count ((Get-ConsistencyKeys $issues) -join ', ')

    New-SheetsBook (Join-Path $dir '基本設計書.xlsx') ([ordered]@{
            '機能一覧'   = @(@('機能ID', '機能名', '画面ID', '帳票ID', 'バッチID'), @('F001', '受注一覧', 'SCR001', '', ''), @('F002', '受注伝票', '', 'RPT001', ''), @('F003', '請求', '', 'RPT003', ''), @('F004', '日次集計', '', '', 'BAT001'))
            '帳票一覧'   = @(@('帳票ID', '帳票名'), @('RPT001', '受注伝票'), @('RPT002', '請求書'))
            'バッチ一覧' = @(@('バッチID', 'バッチ名'), @('BAT001', '日次受注集計'))
        })
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal (@(
        'エラー|参照先なし|基本設計書.xlsx|機能一覧!D4',
        '警告|基本設計との違い|基本設計書.xlsx|帳票一覧!A3',
        '警告|基本設計との違い|基本設計書.xlsx|バッチ一覧!A2') | Sort-Object) (Get-ConsistencyKeys $issues)
    $details = @($issues | ForEach-Object { $_.Detail })
    foreach ($expected in @(
            '帳票「RPT003」が定義されていません（帳票ID）',
            '基本設計にある帳票「RPT002 請求書」が、詳細設計にありません（設計書が無いか、ID が違います）',
            'バッチ「BAT001」の名前が違います（基本設計: 日次受注集計 / DD08_バッチ設計書_BAT001_日次集計.xlsx: 日次集計）')) {
        Assert-True ($details -contains $expected) "指摘がありません: $expected"
    }
}

Test-Case 'Test-DesignConsistency: 基本設計書の CRUD 図（行は画面 ID か機能 ID、列はテーブルの ID か名前）を、詳細設計の使用テーブルから作った CRUD 図と比べる' {
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    $crudSheet = @(@('画面ID', '画面名', '顧客マスタ', '商品マスタ', '受注', 'T_ORDER_DETAIL（受注明細）', '担当者マスタ'),
        @('SCR001', '受注一覧', 'R', '', 'R', '○', ''),
        @('SCR002', '受注登録', 'R', 'RU', 'CR', 'C', ''),
        @('SCR003', '受注詳細', 'R', 'RU', 'RU', '', ''))
    New-SheetsBook (Join-Path $dir '基本設計書.xlsx') ([ordered]@{
            '機能一覧' = @(@('機能ID', '機能名', '画面ID'), @('F001', '受注一覧', 'SCR001'), @('F002', '受注登録', 'SCR002'), @('F003', '受注詳細', 'SCR003'))
            '画面一覧' = @(@('画面ID', '画面名'), @('SCR001', '受注一覧'), @('SCR002', '受注登録'), @('SCR003', '受注詳細'))
            'CRUD図'   = $crudSheet
            '機能CRUD' = @(@('機能ID', '機能名', '顧客マスタ', '受注'), @('F002', '受注登録', 'R', 'C'))
        })
    $text = Get-HostText { $script:crudIssues = @(& $ConsistencyTool -Path $dir 3>$null) }
    $issues = $script:crudIssues
    $scr2 = 'DD02_画面設計書_SCR002_受注登録.xlsx'
    Assert-Equal (@(
        '情報|未確認|基本設計書.xlsx|CRUD図!G1',
        '警告|CRUD の違い|基本設計書.xlsx|CRUD図!F2',
        '警告|CRUD の違い|基本設計書.xlsx|CRUD図!E3',
        '警告|CRUD の違い|基本設計書.xlsx|CRUD図!F4',
        "警告|CRUD の違い|$scr2|イベント定義!AM8",
        "警告|CRUD の違い|$scr2|イベント定義!AM11") | Sort-Object) (Get-ConsistencyKeys $issues)
    $details = @($issues | ForEach-Object { $_.Detail })
    foreach ($expected in @(
            '基本設計の CRUD 図の列「担当者マスタ」は、テーブルの ID・名前と合わないため比べていません',
            '基本設計の CRUD 図には SCR001 → T_ORDER_DETAIL（○）がありますが、詳細設計（使用テーブルなど）にはありません',
            'SCR002 → T_ORDER の CRUD が違います（基本設計: CR / 詳細設計: C）',
            '詳細設計には SCR003 → T_ORDER_DETAIL（R）がありますが、基本設計の CRUD 図（基本設計書.xlsx（CRUD図））は空欄です',
            'M_PRODUCT は詳細設計で使っています（F002（SCR002）: RU）が、基本設計の CRUD 図（基本設計書.xlsx（機能CRUD））に列がありません',
            'T_ORDER_DETAIL は詳細設計で使っています（F002（SCR002）: C）が、基本設計の CRUD 図（基本設計書.xlsx（機能CRUD））に列がありません')) {
        Assert-True ($details -contains $expected) "指摘がありません: $expected"
    }
    Assert-Match $text '基本設計の CRUD 図（基本設計書\.xlsx（CRUD図）・基本設計書\.xlsx（機能CRUD））を、詳細設計の使用テーブルから作った CRUD 図と比べました'

    # 詳細設計の画面の行が基本設計の CRUD 図に無ければ知らせる（IF の行の無い CRUD 図では、IF001 の行は比べない）
    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    New-SheetsBook (Join-Path $dir '基本設計書.xlsx') ([ordered]@{ 'CRUD' = @($crudSheet[0], $crudSheet[1], $crudSheet[2]) })
    $all = Invoke-Consistency @{ Path = $dir }
    $issues = @($all | Where-Object { $_.Detail -like '*の行が*' })
    Assert-Equal 1 $issues.Count
    Assert-Equal 'SCR003 の行が、基本設計の CRUD 図（基本設計書.xlsx（CRUD））にありません（詳細設計: M_CUSTOMER: R・M_PRODUCT: RU・T_ORDER: RU・T_ORDER_DETAIL: R）' $issues[0].Detail
    Assert-Equal 'DD02_画面設計書_SCR003_受注詳細.xlsx' $issues[0].File
}
