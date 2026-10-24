#     New-TestData.ps1 のテスト。テーブル定義書は docs\design-sample と、テストの中で作るもの
#     （作った INSERT 文は、Export-TableDdl.ps1 の DDL と一緒に PostgreSQL 16 で登録できることを手元で確かめた）

$DataTool = Join-Path $script:ToolsDir 'New-TestData.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Read-DataCsv {
    #     作った CSV を読む（1 行目はカラムの物理名）
    param([string] $Dir, [string] $Table)
    return @(Import-Csv -LiteralPath (Join-Path $Dir "$Table.csv") -Encoding UTF8)
}

function Get-SqlLines {
    param([string] $Dir, [string] $Pattern)
    return @(([System.IO.File]::ReadAllText((Join-Path $Dir 'テストデータ.sql')) -split "`r`n") | Where-Object { $_ -match $Pattern })
}

function New-ItemTableBook {
    #     いろいろな型のテーブル（T_ITEM）と、コード定義（区分名を最初の行だけに書いた表）のブック
    param([string] $Path)
    New-SheetsBook $Path ([ordered]@{
            'T_ITEM' = @(
                @('論理名', '物理名', 'データ型', '桁数', '小数', 'PK', '必須', '外部キー参照先', '備考'),
                @('品目ID', 'ITEM_ID', 'INTEGER', '', '', '1', '○', '', ''),
                @('品目の名称', 'ITEM_NAME', 'VARCHAR', '5', '', '', '○', '', ''),
                @('単価', 'PRICE', 'NUMERIC', '5', '2', '', '○', '', ''),
                @('発売日', 'RELEASE_DATE', 'DATE', '', '', '', '', '', ''),
                @('更新日時', 'UPDATED_AT', 'TIMESTAMP', '', '', '', '○', '', ''),
                @('開始時刻', 'START_TIME', 'TIME', '', '', '', '○', '', ''),
                @('公開', 'IS_PUBLIC', 'BOOLEAN', '', '', '', '○', '', ''),
                @('画像', 'IMAGE', 'BLOB', '', '', '', '○', '', ''),
                @('親品目ID', 'PARENT_ID', 'INTEGER', '', '', '', '', 'T_ITEM.ITEM_ID', ''),
                @('状態', 'STATUS', 'CHAR', '1', '', '', '○', '', 'コード区分 K01 の値'),
                @('表示フラグ', 'DISPLAY_FLAG', 'CHAR', '1', '', '', '○', '', ''))
            'コード' = @(
                @('コード区分ID', 'コード区分名', 'コード値', '表示名'),
                @('K01', '品目状態', 'A', '有効'),
                @('', '', 'B', '停止'),
                @('K02', '色', 'R', '赤'))
        })
}

Test-Case 'New-TestData: 詳細設計サンプルのテーブル定義書から、参照先の行を参照する CSV と INSERT 文を作る（主キー・外部キー・コード値・境界値）' {
    $out = Join-Path (New-TestDir) 'data'
    $text = Get-HostText { & $DataTool -Path $designDir -OutDir $out -BaseDate '2026-04-01' | Out-Null }
    Assert-Match $text 'テストデータを作りました: テーブル 4 件・行 24 件（標準 12・境界値 12）'
    Assert-Match $text 'コード値から作ったカラム: .*T_ORDER\.ORDER_STATUS'
    Assert-Equal @(@('M_CUSTOMER.csv', 'M_PRODUCT.csv', 'T_ORDER.csv', 'T_ORDER_DETAIL.csv', 'テストデータ.sql') | Sort-Object) @(Get-ChildItem -LiteralPath $out | ForEach-Object { $_.Name } | Sort-Object)
    Assert-Equal @(0xEF, 0xBB, 0xBF) @([System.IO.File]::ReadAllBytes((Join-Path $out 'T_ORDER.csv'))[0..2]) '既定は BOM 付き（Excel でそのまま開ける）'

    $orders = Read-DataCsv $out 'T_ORDER'
    Assert-Equal @('000000000001', '000000000002', '000000000003', '999999999999', '000000000000', '000000000004') @($orders | ForEach-Object { $_.ORDER_NO }) '主キーは番号（固定長は桁数まで 0 で埋める）。最大・最小・NULL の行も重ならない'
    Assert-Equal @('10', '20', '90', '90', '10', '10') @($orders | ForEach-Object { $_.ORDER_STATUS }) 'コード定義書のコード値を順に。最大は最後、最小と NULL の行は最初'
    Assert-Equal '2026-04-01|2026-04-02|9999-12-31|1900-01-01|' ('{0}|{1}|{2}|{3}|{4}' -f $orders[0].DELIVERY_DATE, $orders[1].DELIVERY_DATE, $orders[3].DELIVERY_DATE, $orders[4].DELIVERY_DATE, $orders[5].DELIVERY_DATE) 'NULL の行は、必須でないカラムを空にする'
    Assert-Equal 500 $orders[3].REMARKS.Length '最大の行の文字は桁数いっぱい'
    Assert-Equal '100|99999999999|0' ('{0}|{1}|{2}' -f $orders[0].TOTAL_AMOUNT, $orders[3].TOTAL_AMOUNT, $orders[4].TOTAL_AMOUNT) '数値の桁数（11 桁）の最大'
    $customers = @(Read-DataCsv $out 'M_CUSTOMER' | ForEach-Object { $_.CUSTOMER_CD })
    Assert-Equal @('00000001', '00000002', '00000003', '99999999', '00000000', '00000004') @($orders | ForEach-Object { $_.CUSTOMER_CD }) '外部キーは参照先の同じ種類の行の値'
    Assert-True (@($orders | Where-Object { $customers -notcontains $_.CUSTOMER_CD }).Count -eq 0) '参照先に無い値は作らない'
    $details = Read-DataCsv $out 'T_ORDER_DETAIL'
    Assert-Equal @('000000000001|1', '000000000002|2', '000000000003|3', '999999999999|999', '000000000000|0', '000000000004|4') @($details | ForEach-Object { '{0}|{1}' -f $_.ORDER_NO, $_.LINE_NO }) '主キーの一部が外部キーでも、組み合わせは重ならない'

    $sql = Get-SqlLines $out '^INSERT INTO T_ORDER '
    Assert-Equal 6 $sql.Count
    Assert-Match $sql[0] "^INSERT INTO T_ORDER \(ORDER_NO, ORDER_DATE, CUSTOMER_CD, ORDER_STATUS, .*\) VALUES \('000000000001', DATE '2026-04-01', '00000001', '10', .*, 1\); -- 標準 1$"
    Assert-Match $sql[5] "NULL, '00', 400, NULL, TIMESTAMP '2026-04-04 09:00:00'.*; -- NULL$"
    $order = @(Get-SqlLines $out '^-- [A-Z_]+（' | ForEach-Object { ($_ -split '（')[0].Substring(3) })
    Assert-Equal @('M_CUSTOMER', 'M_PRODUCT', 'T_ORDER', 'T_ORDER_DETAIL') $order '参照先のテーブルから順に'
}

Test-Case 'New-TestData: 型ごとの値（短い文字・小数・日付・日時・時刻・真偽・バイナリ）と、備考のコード区分・フラグ・自分を参照する外部キー' {
    $dir = New-TestDir
    New-ItemTableBook (Join-Path $dir 'テーブル定義書.xlsx')
    $out = Join-Path $dir 'data'
    $warnings = @(& $DataTool -Path $dir -OutDir $out -BaseDate '2026-04-01' -Database PostgreSQL 6>$null 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })
    Assert-Equal @('T_ITEM: IMAGE はバイナリのため値を作りません（必須のため、登録の前に値を入れてください）') $warnings
    $rows = Read-DataCsv $out 'T_ITEM'
    $show = { param($r) '{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}' -f $r.ITEM_ID, $r.ITEM_NAME, $r.PRICE, $r.RELEASE_DATE, $r.UPDATED_AT, $r.START_TIME, $r.IS_PUBLIC, $r.IMAGE, $r.PARENT_ID, $r.STATUS, $r.DISPLAY_FLAG }
    Assert-Equal @(
        '1|00001|100.50|2026-04-01|2026-04-01 09:00:00|09:00:00|1|||A|0',
        '2|00002|200.50|2026-04-02|2026-04-02 09:00:00|09:01:00|0|||B|1',
        '3|00003|300.50|2026-04-03|2026-04-03 09:00:00|09:02:00|1|||A|0',
        '2147483647|あああああ|999.99|9999-12-31|9999-12-31 23:59:59|23:59:59|1|||B|1',
        '0|あ|0|1900-01-01|1900-01-01 00:00:00|00:00:00|0|||A|0',
        '4|00004|400.50||2026-04-04 09:00:00|09:03:00|0|||A|0') @($rows | ForEach-Object { & $show $_ }) '論理名と番号が桁数に入らなければ、番号を 0 で埋める。区分名を最初の行だけに書いたコード定義も読む'
    $sql = Get-SqlLines $out '^INSERT INTO T_ITEM '
    Assert-Match $sql[0] "VALUES \(1, '00001', 100\.50, DATE '2026-04-01', TIMESTAMP '2026-04-01 09:00:00', TIME '09:00:00', TRUE, NULL, NULL, 'A', '0'\); -- 標準 1$" 'PostgreSQL の書き方'
    Assert-Match $sql[1] ", FALSE, NULL, NULL, 'B', '1'\); -- 標準 2$"

    # SQL Server・MySQL・Oracle の書き方
    & $DataTool -Path $dir -OutDir $out -BaseDate '2026-04-01' -Database SQLServer -Rows 1 -NoBoundary 6>$null 3>$null | Out-Null
    Assert-Match @(Get-SqlLines $out '^INSERT INTO T_ITEM ')[0] "VALUES \(1, N'00001', 100\.50, '2026-04-01', '2026-04-01 09:00:00', '09:00:00', 1, NULL, NULL, N'A', N'0'\);"
    & $DataTool -Path $dir -OutDir $out -BaseDate '2026-04-01' -Database Oracle -Rows 1 -NoBoundary 6>$null 3>$null | Out-Null
    Assert-Match @(Get-SqlLines $out '^INSERT INTO T_ITEM ')[0] "TIMESTAMP '2026-04-01 09:00:00', TIMESTAMP '2026-04-01 09:00:00', 1, NULL" 'Oracle には時刻だけの型が無いので、基準日の日時'
}

Test-Case 'New-TestData: 整数部の無い数値（NUMBER(3,3)・NUMERIC(4,4) など）は、桁あふれしない 0.xxx の値にする' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir 'テーブル定義書.xlsx') ([ordered]@{
            'M_TAX' = @(
                @('論理名', '物理名', 'データ型', '桁数', '小数', 'PK', '必須'),
                @('税区分', 'TAX_CD', 'CHAR', '2', '', '1', '○'),
                @('税率', 'TAX_RATE', 'NUMBER', '3', '3', '', '○'),
                @('割引率', 'DISCOUNT_RATE', 'NUMERIC(4,4)', '', '', '', ''),
                @('換算係数', 'RATE_COEF', 'NUMBER', '2', '4', '', ''))
        })
    $out = Join-Path $dir 'data'
    & $DataTool -Path $dir -OutDir $out -BaseDate '2026-04-01' -Database PostgreSQL 6>$null 3>$null | Out-Null
    Assert-Equal @('0.001|0.0001|0.0001', '0.002|0.0002|0.0002', '0.003|0.0003|0.0003', '0.999|0.9999|0.0099', '0|0|0', '0.004||') @(Read-DataCsv $out 'M_TAX' | ForEach-Object { '{0}|{1}|{2}' -f $_.TAX_RATE, $_.DISCOUNT_RATE, $_.RATE_COEF }) '標準は番号を小数の下の桁に。最大は 0.999（小数が桁数より大きい Oracle の NUMBER(2,4) は 0.0099）'
    Assert-Match @(Get-SqlLines $out '^INSERT INTO M_TAX ')[3] "VALUES \('99', 0\.999, 0\.9999, 0\.0099\); -- 最大$"
}

Test-Case 'New-TestData: 和暦の設定（ja-JP の JapaneseCalendar）でも、日付・日時は西暦で書く' {
    $dir = New-TestDir
    New-ItemTableBook (Join-Path $dir 'テーブル定義書.xlsx')
    $out = Join-Path $dir 'data'
    $baseDate = [datetime]'2026-04-01'
    $wareki = New-Object System.Globalization.CultureInfo('ja-JP')
    $wareki.DateTimeFormat.Calendar = New-Object System.Globalization.JapaneseCalendar
    $saved = [System.Threading.Thread]::CurrentThread.CurrentCulture
    try {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $wareki
        & $DataTool -Path $dir -OutDir $out -BaseDate $baseDate -Rows 1 -NoBoundary 6>$null 3>$null | Out-Null
    }
    finally { [System.Threading.Thread]::CurrentThread.CurrentCulture = $saved }
    Assert-Match @(Get-SqlLines $out '^INSERT INTO T_ITEM ')[0] "DATE '2026-04-01', TIMESTAMP '2026-04-01 09:00:00', TIMESTAMP '2026-04-01 09:00:00', 1, NULL" 'Oracle の時刻は基準日の日時'
    $row = @(Read-DataCsv $out 'T_ITEM')[0]
    Assert-Equal '2026-04-01|2026-04-01 09:00:00|09:00:00' ('{0}|{1}|{2}' -f $row.RELEASE_DATE, $row.UPDATED_AT, $row.START_TIME)
}

Test-Case 'New-TestData: -Table で選んだテーブルの参照先も作り、-Rows・-NoBoundary・-IncludeDelete・-OutXlsx を使える' {
    $dir = New-TestDir
    $out = Join-Path $dir 'data'
    $xlsx = Join-Path $dir 'テストデータ.xlsx'
    $text = Get-HostText { & $DataTool -Path $designDir -OutDir $out -Table 'T_ORDER_D*' -Rows 2 -NoBoundary -IncludeDelete -BaseDate '2026-04-01' -OutXlsx $xlsx | Out-Null }
    Assert-Match $text 'テーブル 4 件・行 8 件（標準 8・境界値 0）'
    Assert-Match $text '参照先のテーブルも作りました: T_ORDER・M_PRODUCT・M_CUSTOMER'
    Assert-Equal @('DELETE FROM T_ORDER_DETAIL;', 'DELETE FROM T_ORDER;', 'DELETE FROM M_CUSTOMER;', 'DELETE FROM M_PRODUCT;') (Get-SqlLines $out '^DELETE ') '参照する側から消す'
    Assert-Equal @('000000000001|0000000001', '000000000002|0000000002') @(Read-DataCsv $out 'T_ORDER_DETAIL' | ForEach-Object { '{0}|{1}' -f $_.ORDER_NO, $_.PRODUCT_CD })
    $book = Read-XlsxSheet -Path $xlsx -Index 1
    Assert-Equal @('一覧', 'M_PRODUCT', 'M_CUSTOMER', 'T_ORDER', 'T_ORDER_DETAIL') $book.Sheets
    Assert-Equal 'T_ORDER_DETAIL|受注明細|2|T_ORDER・M_PRODUCT' ('{0}|{1}|{2}|{3}' -f $book.Rows[5]['A'], $book.Rows[5]['B'], $book.Rows[5]['C'], $book.Rows[5]['D'])
    $detail = Read-XlsxSheet -Path $xlsx -Index 5
    Assert-Equal 'パターン|ORDER_NO|LINE_NO' ('{0}|{1}|{2}' -f $detail.Rows[1]['A'], $detail.Rows[1]['B'], $detail.Rows[1]['C'])
    Assert-Equal '標準 1|000000000001|1' ('{0}|{1}|{2}' -f $detail.Rows[2]['A'], $detail.Rows[2]['B'], $detail.Rows[2]['C']) '0 で始まる値も文字のまま'
    $w = @(& $DataTool -Path $designDir -OutDir $out -Table 'NOTHING' 6>$null 3>&1 | ForEach-Object { "$_" })
    Assert-Match ($w -join ' ') '-Table に合うテーブルがありません: NOTHING'
}

Test-Case 'New-TestData: 1 シートにカラム定義の表が 2 つあれば最初の表だけを使い（警告する）、ページごとにくり返した見出しの行は読まない（カラム・コード値）' {
    $dir = New-TestDir
    $header = @('論理名', '物理名', 'データ型', '桁数', 'PK', '必須', '備考')
    $codeHeader = @('コード区分ID', 'コード区分名', 'コード値', '表示名')
    New-SheetsBook (Join-Path $dir 'テーブル定義書.xlsx') ([ordered]@{
            'T_ITEM' = @($header,
                @('品目ID', 'ITEM_ID', 'INTEGER', '', '1', '○', ''),
                @('品目名', 'ITEM_NAME', 'VARCHAR', '20', '', '○', ''),
                @('■履歴', '', '', '', '', '', ''),
                $header,
                @('履歴番号', 'HIST_NO', 'INTEGER', '', '1', '○', ''))
            'T_ORDER' = @($header,
                @('受注ID', 'ORDER_ID', 'INTEGER', '', '1', '○', ''),
                $header,
                @('状態', 'STATUS', 'CHAR', '1', '', '○', 'コード区分 K01 の値'))
            'コード' = @($codeHeader,
                @('K01', '受注状態', 'A', '受付'),
                $codeHeader,
                @('', '', 'B', '出荷'))
        })
    $out = Join-Path $dir 'data'
    $warnings = @(& $DataTool -Path $dir -OutDir $out -Rows 2 -NoBoundary 6>$null 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })
    Assert-Equal @('テーブル定義書.xlsx T_ITEM!B5: 1 つのシートにカラム定義の表が 2 つ以上あるため、この表は使いません（テーブル T_ITEM には最初の表を使います）') $warnings
    Assert-Equal @(
        "INSERT INTO T_ITEM (ITEM_ID, ITEM_NAME) VALUES (1, '品目名1'); -- 標準 1",
        "INSERT INTO T_ITEM (ITEM_ID, ITEM_NAME) VALUES (2, '品目名2'); -- 標準 2",
        "INSERT INTO T_ORDER (ORDER_ID, STATUS) VALUES (1, 'A'); -- 標準 1",
        "INSERT INTO T_ORDER (ORDER_ID, STATUS) VALUES (2, 'B'); -- 標準 2") (Get-SqlLines $out '^INSERT ') 'くり返した見出しの下の行は、上の表（カラム定義・コード区分 K01）の続き'
}

Test-Case 'New-TestData: Get-Help でヘルプが表示される' {
    $h = Get-Help $DataTool
    Assert-Match $h.Synopsis 'テストデータ'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
