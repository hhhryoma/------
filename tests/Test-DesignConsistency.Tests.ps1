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
    Assert-Equal @('指摘一覧', '定義と参照') $stats.Sheets
    $row = @($stats.Rows.Values | Where-Object { $_['A'] -eq 'メッセージ' })[0]
    Assert-Equal '16|1' ('{0}|{1}' -f $row['B'], $row['E']) 'メッセージの定義 16・参照先なし 1'
}

Test-Case 'Test-DesignConsistency: Get-Help でヘルプが表示される' {
    $h = Get-Help $ConsistencyTool
    Assert-Match $h.Synopsis '参照'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
