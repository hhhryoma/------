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

Test-Case 'Test-DesignConsistency: 表示制御のシートの空欄・書き方・過不足と、画面項目定義・イベント定義との食い違い、戻す処理の抜けを見つける' {
    # 詳細設計サンプルの表示制御は整合している（画面 3 件。状態 8・項目 57）
    $all = @(& $ConsistencyTool -Path $sampleDir 6>&1 3>$null)
    Assert-Match (($all | Where-Object { $_ -is [System.Management.Automation.InformationRecord] } | ForEach-Object { "$_" }) -join "`n") '表示制御: 画面 3 件（状態 8・項目 57）'
    Assert-Equal 0 @($all | Where-Object { $_.PSObject.TypeNames -contains 'DesignConsistency.Issue' }).Count

    $dir = Copy-Fixture (Join-Path $sampleDir '.')
    $scr1 = 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $scr2 = 'DD02_画面設計書_SCR002_受注登録.xlsx'
    $scr3 = 'DD02_画面設計書_SCR003_受注詳細.xlsx'
    Set-CellValues (Join-Path $dir $scr1) '表示制御' @{
        R12 = $null                                   # 空欄
        W13 = '入力OK'                                # 書き方に無い
        AB23 = '○'                                    # 一覧（受注一覧）が非表示の状態で、一覧の項目を表示する
        R20 = '×'; W20 = '△'; AB20 = '×'             # どの状態でも押せないボタン
        AO7 = 'EV02、EV09'                            # イベント定義に無いイベント
    }
    # 行追加を非活性にする状態（明細20行。EV05）から、表示に戻る状態（入力中）が初期表示（EV01）だけになる
    Set-CellValues (Join-Path $dir $scr2) '表示制御' @{ AO6 = 'EV01' }
    Set-CellValues (Join-Path $dir $scr3) '表示制御' @{
        W29 = '○'; AB29 = '○'                         # EV01 で非表示にする更新ボタンを、どの状態でも表示する
        R32 = '○'                                     # 隠し項目を表示する
        R18 = '表示のみ'                              # どの状態でも入力できない入力項目
        A33 = 22; C33 = '出荷日'; J33 = 'shipDate'; R33 = '○'; W33 = '○'; AB33 = '○'           # 画面項目定義に無い項目
        A31 = $null; C31 = $null; J31 = $null; R31 = $null; W31 = $null; AB31 = $null      # 戻るボタンの行を消す
        C17 = '支払い方法'                            # 項目名が違う
        AO7 = $null                                   # 状態になるイベントが空欄
    }
    $issues = Invoke-Consistency @{ Path = $dir }
    Assert-Equal (@(
        "エラー|参照先なし|$scr1|表示制御!AO7",
        "警告|表示制御の空欄|$scr1|表示制御!R12",
        "警告|表示制御の書き方|$scr1|表示制御!W13",
        "警告|項目定義との矛盾|$scr1|表示制御!C20",
        "警告|項目定義との矛盾|$scr1|表示制御!C23",
        "警告|戻す処理なし|$scr2|表示制御!C26",
        "警告|表示制御の過不足|$scr3|画面項目定義!C25",
        "警告|イベントとの食い違い|$scr3|イベント定義!P6",
        "警告|表示制御の空欄|$scr3|表示制御!AO7",
        "警告|表示制御の過不足|$scr3|表示制御!C17",
        "警告|項目定義との矛盾|$scr3|表示制御!C18",
        "警告|項目定義との矛盾|$scr3|表示制御!C32",
        "警告|表示制御の過不足|$scr3|表示制御!C33") | Sort-Object) (Get-ConsistencyKeys $issues)

    $details = @($issues | ForEach-Object { $_.Detail })
    foreach ($expected in @(
            'イベント「EV09」が定義されていません（表示制御の状態「検索後」）',
            'No.1「受注番号」の状態「初期表示」が空欄です',
            '「入力OK」は表示制御の書き方にありません（○・△・表示のみ・× などで書きます）',
            'No.9「新規登録」はボタンですが、どの状態でも表示（押せる状態）になっていません',
            'No.12「受注番号」は一覧「受注一覧」の項目ですが、一覧が非表示の状態「該当なし」で非表示になっていません',
            'No.16「行追加」は状態「明細20行」（EV05）で非活性になりますが、表示になる状態「入力中」には画面を開いたときしかなりません（元に戻すイベントがありません）',
            'No.20「戻る」が、表示制御の表にありません（どの状態で表示するかが決まっていません）',
            'EV01 で「更新」を非表示にしていますが、表示制御の状態「受付・キャンセル」のどれも非表示になっていません',
            '状態「出荷指示済」になるイベントが空欄です',
            'No.6「支払い方法」の項目名が、画面項目定義（No.6「支払方法」）と違います',
            'No.7「希望納品日」は入力項目（入出力「IO」）ですが、どの状態でも表示（入力できる状態）になっていません',
            'No.21「バージョン番号」は隠し項目ですが、状態「受付」で非表示になっていません',
            'No.22「出荷日」（項目ID「shipDate」）は、画面項目定義にありません')) {
        Assert-True ($details -contains $expected) ("指摘がありません: {0}`n実際: {1}" -f $expected, ($details -join "`n"))
    }
    Assert-Equal 'イベント' @($issues | Where-Object Item -eq '参照先なし')[0].Kind
    Assert-Equal 12 @($issues | Where-Object Kind -eq '表示制御').Count
}

Test-Case 'Test-DesignConsistency: 表示制御のシートの名前と書き方を変えられ、2 段の見出し・状態の表の無い表も読み、条件や打ち消しの文は表示を変える処理として読まない' {
    $dir = New-TestDir
    # New-SheetsBook は 1 行目の列の数だけ書くので、1 行目（画面ID の行）を表の幅まで空の文字で埋める
    $items = { param($id, $rows) @(, @('画面ID', $id, '画面名', 'テスト', '', '')) + @(, @('No', '項目名', '項目ID', '区分', '入出力', '桁数')) + $rows }
    $events = { param($id, $rows) @(, @('画面ID', $id, '', '')) + @(, @('イベントID', 'イベント名', 'トリガー', '処理概要')) + $rows }
    New-SheetsBook (Join-Path $dir 'DD02_画面設計書_S01.xlsx') ([ordered]@{
            '画面項目定義' = (& $items 'S01' @(@(1, '検索条件', 'cond', 'テキスト', 'I', '10'), @(2, '一覧', 'list', '一覧', 'O', ''), @(3, '名前', 'row.name', 'ラベル', 'O', '10'), @(4, '登録', 'btnRegist', 'ボタン', '－', '')))
            'イベント定義' = (& $events 'S01' @(
                    @('EV01', '初期表示', '画面表示時', '一覧は非表示とする。登録ボタンが非活性の場合は何もしない。'),
                    @('EV02', '検索', '検索ボタン押下', '一覧を表示する。検索条件を非活性にしない。'),
                    @('EV03', '登録', '登録ボタン押下', '登録したあと、登録ボタンをグレーアウトする。')))
            # 状態の表（イベントの「なし」は、その状態になるイベントが無い）と、2 段の見出し（上の段は「表示制御」、下の段に状態の名前）の表
            '画面状態'     = @(@('画面ID', 'S01', '', '', ''), @('■ 状態'), @('状態', 'イベント'), @('開いたとき', 'EV01'), @('検索後', 'EV02'), @('確認中', 'なし'), @('■ 項目'),
                @('項目名', '項目ID', '表示制御'), @('', '', '開いたとき', '検索後', '確認中'),
                @('検索条件', 'cond', '可', '可', '可'), @('一覧', 'list', '不可視', '可', '可'), @('名前', 'row.name', '可', '可', '可'), @('登録', 'btnRegist', '可', '可', '可'))
        })
    # 表示制御のシートの無い画面設計書
    New-SheetsBook (Join-Path $dir 'DD02_画面設計書_S02.xlsx') ([ordered]@{ '画面項目定義' = (& $items 'S02' @(, @(1, '戻る', 'btnBack', 'ボタン', '－', ''))) })
    # 状態の表の無い表示制御（イベントの文は、どれかの状態と合えばよい）
    New-SheetsBook (Join-Path $dir 'DD02_画面設計書_S03.xlsx') ([ordered]@{
            '画面項目定義' = (& $items 'S03' @(, @(1, '詳細', 'btnDetail', 'ボタン', '－', '')))
            'イベント定義' = (& $events 'S03' @(, @('EV01', '初期表示', '画面表示時', '詳細ボタンを非表示にする。')))
            '画面状態'     = @(@('画面ID', 'S03', '', ''), @('No', '項目名', '項目ID', '通常'), @(1, '詳細', 'btnDetail', '可'))
        })
    $words = Join-Path $dir 'words.csv'
    @('状態,書き方', '表示,可', '非活性,グレーアウト', '非表示,不可視') | Set-Content -LiteralPath $words -Encoding UTF8
    $issues = Invoke-Consistency @{ Path = $dir; DisplaySheet = '^画面状態$'; DisplayWords = $words }
    $got = @($issues | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Item, $_.File, $_.Location, $_.Level, $_.Detail })
    # S03 の「非表示」は書き方の CSV に無いが、状態の名前なので読む
    Assert-Equal @(
        'イベントとの食い違い|DD02_画面設計書_S01.xlsx|イベント定義!D5|警告|EV03 で「登録」を非活性にしていますが、EV03 が表示制御の状態の表にありません（どの状態になるかが書かれていません）',
        '項目定義との矛盾|DD02_画面設計書_S01.xlsx|画面状態!A12|警告|「名前」は一覧「一覧」の項目ですが、一覧が非表示の状態「開いたとき」で非表示になっていません',
        '表示制御の過不足|DD02_画面設計書_S02.xlsx|画面項目定義!B2|警告|表示制御のシートがありません（ほかの画面設計書にはあります。シートの名前: ^画面状態$）',
        'イベントとの食い違い|DD02_画面設計書_S03.xlsx|イベント定義!D3|警告|EV01 で「詳細」を非表示にしていますが、表示制御のどの状態でも非表示になっていません') $got

    # 表示制御のシートがどの画面設計書にも無ければ、表示制御は確かめない
    Assert-Equal 0 (Invoke-Consistency @{ Path = $dir; DisplaySheet = '^無いシート$'; DisplayWords = $words }).Count

    $bad = Join-Path $dir 'bad.csv'
    @('状態,書き方', '見えない,○') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $dir -DisplaySheet '^画面状態$' -DisplayWords $bad 6>$null 3>$null } '表示 / 非活性 / 表示のみ / 非表示'
    @('状態,書き方', '表示,○', '非表示,○') | Set-Content -LiteralPath $bad -Encoding UTF8
    Assert-Throws { & $ConsistencyTool -Path $dir -DisplaySheet '^画面状態$' -DisplayWords $bad 6>$null 3>$null } '書き方にもあります'
    Assert-Throws { & $ConsistencyTool -Path $dir -DisplaySheet ' ' 6>$null 3>$null } '-DisplaySheet'
}
