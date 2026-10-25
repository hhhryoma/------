#     Export-TableDdl.ps1 のテスト。テーブル定義書は docs\design-sample、DB のカタログは tests\fixtures\ddl
#     （postgresql_columns.csv は、作った DDL を PostgreSQL 16 で実行し、列を変えてから information_schema.columns を出力したもの）

$DdlTool = Join-Path $script:ToolsDir 'Export-TableDdl.ps1'
$tableDoc = Join-Path $script:RepoRoot 'docs/design-sample/DD03_テーブル定義書.xlsx'
$ddlDir = Join-Path $script:Fixtures 'ddl'

function Invoke-Ddl {
    #     DDL（文字列）と警告を返す
    param([hashtable] $Options)
    $out = @(& $DdlTool @Options 6>$null 3>&1)
    return [pscustomobject]@{
        Text = (@($out | Where-Object { $_ -is [string] }) -join "`n").Replace("`r`n", "`n")
        Warnings = @($out | Where-Object { $_ -is [System.Management.Automation.WarningRecord] } | ForEach-Object { $_.Message })
        Objects = @($out | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.WarningRecord] })
    }
}

Test-Case 'Export-TableDdl: 詳細設計サンプルのテーブル定義書から Oracle の DDL を作る（型・DEFAULT・主キー・コメント・インデックス・外部キー）' {
    $r = Invoke-Ddl @{ Path = $tableDoc }
    Assert-Equal 0 $r.Warnings.Count ($r.Warnings -join ' / ')
    $t = $r.Text
    Assert-Match $t '(?m)^CREATE TABLE T_ORDER \($'
    Assert-Match $t '(?m)^    ORDER_NO\s+CHAR\(12 CHAR\)\s+NOT NULL,$'
    Assert-Match $t '(?m)^    TOTAL_AMOUNT\s+NUMBER\(11,0\)\s+DEFAULT 0 NOT NULL,$'
    Assert-Match $t '(?m)^    ORDER_STATUS\s+CHAR\(2 CHAR\)\s+DEFAULT ''10'' NOT NULL,$'
    Assert-Match $t '(?m)^    REMARKS\s+VARCHAR2\(500 CHAR\),$' 'NULL 可のカラム'
    Assert-Match $t '(?m)^    VERSION_NO\s+NUMBER\(10\)\s+DEFAULT 1 NOT NULL,$'
    Assert-Match $t '(?m)^    CONSTRAINT PK_T_ORDER_DETAIL PRIMARY KEY \(ORDER_NO, LINE_NO\)$' '主キーの順番'
    Assert-Match $t "(?m)^COMMENT ON COLUMN T_ORDER\.ORDER_NO IS '受注番号';$"
    Assert-Match $t "(?m)^COMMENT ON TABLE M_CUSTOMER IS '顧客マスタ';$"
    Assert-Match $t '(?m)^CREATE INDEX IX_T_ORDER_01 ON T_ORDER \(CUSTOMER_CD, ORDER_DATE\);$'
    Assert-Match $t '(?m)^ALTER TABLE T_ORDER_DETAIL ADD CONSTRAINT FK_T_ORDER_DETAIL_T_ORDER FOREIGN KEY \(ORDER_NO\) REFERENCES T_ORDER \(ORDER_NO\);$'
    Assert-Equal 3 ([regex]::Matches($t, '(?m)^ALTER TABLE .* FOREIGN KEY').Count) '外部キー'
    # 参照先のテーブルが先
    $order = @([regex]::Matches($t, '(?m)^CREATE TABLE (\S+)') | ForEach-Object { $_.Groups[1].Value })
    Assert-True ([array]::IndexOf($order, 'M_CUSTOMER') -lt [array]::IndexOf($order, 'T_ORDER')) ($order -join ',')
    Assert-True ([array]::IndexOf($order, 'T_ORDER') -lt [array]::IndexOf($order, 'T_ORDER_DETAIL')) ($order -join ',')
}

Test-Case 'Export-TableDdl: PostgreSQL・SQL Server・MySQL の型とコメント、-IncludeDrop（参照しているテーブルから削除）と -Table' {
    $pg = (Invoke-Ddl @{ Path = $tableDoc; Database = 'PostgreSQL'; IncludeDrop = $true }).Text
    Assert-Match $pg '(?m)^    REMARKS\s+VARCHAR\(500\),$'
    Assert-Match $pg '(?m)^    CREATED_AT\s+TIMESTAMP\s+DEFAULT CURRENT_TIMESTAMP NOT NULL,$'
    $drops = @([regex]::Matches($pg, '(?m)^DROP TABLE IF EXISTS (\S+) CASCADE;$') | ForEach-Object { $_.Groups[1].Value })
    Assert-Equal 4 $drops.Count
    Assert-True ([array]::IndexOf($drops, 'T_ORDER_DETAIL') -lt [array]::IndexOf($drops, 'T_ORDER')) ($drops -join ',')

    $ss = (Invoke-Ddl @{ Path = $tableDoc; Database = 'SQLServer'; Table = 'T_ORDER' }).Text
    Assert-Match $ss '(?m)^    REMARKS\s+NVARCHAR\(500\),$'
    Assert-Match $ss "@value = N'受注番号'.*@level1name = N'T_ORDER', @level2type = N'COLUMN', @level2name = N'ORDER_NO';"
    Assert-Equal 1 ([regex]::Matches($ss, '(?m)^CREATE TABLE ').Count) '-Table で 1 テーブル'

    $my = (Invoke-Ddl @{ Path = $tableDoc; Database = 'MySQL'; Table = 'M_*' }).Text
    Assert-Match $my "(?m)^    CUSTOMER_CD\s+CHAR\(8\)\s+NOT NULL COMMENT '顧客コード',$"
    Assert-Match $my "(?m)^\) COMMENT = '顧客マスタ';$"
    Assert-Equal 2 ([regex]::Matches($my, '(?m)^CREATE TABLE ').Count) 'ワイルドカード'
}

Test-Case 'Export-TableDdl: 定義書の気になる点（予約語・名前の文字と長さ・未知の型・桁数の抜け・主キーなし・外部キー）を警告する' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 },
        @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 }, @{ Header = '外部キー参照先'; Width = 20 })
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'M_PARENT'; Columns = $cols; Rows = @(@{ Values = @('コード', 'CODE', 'CHAR', '8', '1', '○', '') }) },
        @{ Name = 'ORDER'; Columns = $cols; Rows = @(
                @{ Values = @('ID', 'ID', 'NUMERIC', '10', '', '○', '') },
                @{ Values = @('名前', 'NAME', 'VARCHAR', '', '', '', '') },
                @{ Values = @('種別', 'KIND', 'MYTYPE', '', '', '', '') },
                @{ Values = @('親', 'PARENT_CD', 'CHAR', '10', '', '', 'M_PARENT.CODE') },
                @{ Values = @('他', 'OTHER_ID', 'INTEGER', '', '', '', 'M_NONE.ID') },
                @{ Values = @('名称', '名称', 'VARCHAR', '10', '', '', '') },
                @{ Values = @('長い', ('A' * 31), 'DATE', '', '', '', '') }) })
    $r = Invoke-Ddl @{ Path = $dir }
    $file = 'テーブル定義書.xlsx'
    Assert-Equal (@(
        "${file} ORDER!B1: テーブル名「ORDER」は SQL の予約語です",
        "${file} ORDER!B3: ORDER.NAME のデータ型 VARCHAR に桁数がありません",
        "${file} ORDER!B4: ORDER.KIND のデータ型「MYTYPE」は型の対応（型の対応.csv）にないため、そのまま書きます",
        "${file} ORDER!B7: カラム名「名称」に英数字と _ 以外の文字が含まれています",
        "${file} ORDER!B8: カラム名「$('A' * 31)」が 30 バイトを超えています（Oracle 12.1 以前では使えません）",
        "${file} ORDER!B1: テーブル ORDER に主キーがありません",
        "${file} ORDER!B5: ORDER.PARENT_CD（CHAR(10 CHAR)）と外部キーの参照先 M_PARENT.CODE（CHAR(8 CHAR)）の型・桁数が違います",
        "${file} ORDER!B6: ORDER.OTHER_ID の外部キーの参照先テーブル M_NONE がテーブル定義書にありません") | Sort-Object) ($r.Warnings | Sort-Object)
    Assert-Match $r.Text '(?m)^    NAME\s+VARCHAR2,$' '桁数の無い型は括弧を省く'
    Assert-Match $r.Text '(?m)^    KIND\s+MYTYPE,$' '未知の型はそのまま'
    Assert-Match $r.Text 'ALTER TABLE ORDER ADD CONSTRAINT FK_ORDER_M_PARENT FOREIGN KEY \(PARENT_CD\) REFERENCES M_PARENT \(CODE\);'
}

Test-Case 'Export-TableDdl: -OutFile に Shift_JIS・CRLF で保存できる' {
    $out = Join-Path (New-TestDir) 'create.sql'
    $r = Invoke-Ddl @{ Path = $tableDoc; OutFile = $out; Encoding = 'ShiftJIS' }
    Assert-Equal '' $r.Text '-OutFile のときは出力しない'
    $bytes = [System.IO.File]::ReadAllBytes($out)
    $text = [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
    Assert-Match $text "COMMENT ON TABLE T_ORDER_DETAIL IS '受注明細';\r\n"
    Assert-True ($bytes[0] -ne 0xEF) 'BOM なし'
}

Test-Case 'Export-TableDdl: -CompareCatalog で DB のカタログ（PostgreSQL の information_schema）との差異を出す' {
    $r = Invoke-Ddl @{ Path = $tableDoc; Database = 'PostgreSQL'; CompareCatalog = (Join-Path $ddlDir 'postgresql_columns.csv') }
    Assert-Equal '' $r.Text '比べるときは DDL を出力しない'
    $got = @($r.Objects | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Table, $_.Column, $_.Item, $_.Definition, $_.Database })
    Assert-Equal @(
        'M_PRODUCT|extra_note|定義書にないカラム||text',
        'T_ORDER|PAYMENT_METHOD|型|CHAR(1)|character varying',
        'T_ORDER|DELIVERY_DATE|NULL|NULL 可|NOT NULL',
        'T_ORDER|REMARKS|桁数|500|400',
        'T_ORDER_DETAIL|QUANTITY|桁数|5,0|7,0',
        'm_other||定義書にないテーブル||m_other') $got
    Assert-Equal 'T_ORDER!J21' @($r.Objects | Where-Object Column -eq 'REMARKS')[0].Location '定義書のセル'
}

Test-Case 'Export-TableDdl: Oracle のカタログの列名（DATA_PRECISION・CHAR_LENGTH・NULLABLE）でも比べられ、-OutXlsx に保存する' {
    $dir = New-TestDir
    $csv = Join-Path $dir 'oracle_columns.csv'
    @(
        'TABLE_NAME,COLUMN_NAME,DATA_TYPE,CHAR_LENGTH,DATA_PRECISION,DATA_SCALE,NULLABLE',
        'M_PRODUCT,PRODUCT_CD,CHAR,10,,,N',
        'M_PRODUCT,PRODUCT_NAME,VARCHAR2,80,,,N',
        'M_PRODUCT,UNIT_PRICE,NUMBER,0,9,0,N',
        'M_PRODUCT,STOCK_QTY,NUMBER,0,7,0,Y',
        'M_PRODUCT,DELETE_FLG,CHAR,1,,,N',
        'M_PRODUCT,CREATED_AT,TIMESTAMP(6),0,,6,N',
        'M_PRODUCT,CREATED_BY,VARCHAR2,20,,,N',
        'M_PRODUCT,UPDATED_AT,TIMESTAMP(6),0,,6,N',
        'M_PRODUCT,UPDATED_BY,VARCHAR2,20,,,N',
        'M_PRODUCT,VERSION_NO,NUMBER,0,10,0,N'
    ) | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Join-Path $dir 'diff.xlsx'
    $r = Invoke-Ddl @{ Path = $tableDoc; Table = 'M_PRODUCT'; CompareCatalog = $csv; OutXlsx = $out }
    $got = @($r.Objects | ForEach-Object { '{0}|{1}|{2}' -f $_.Column, $_.Item, $_.Database })
    Assert-Equal @('PRODUCT_NAME|桁数|80', 'STOCK_QTY|NULL|NULL 可') $got
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('DB との差異', '定義書の指摘', 'テーブル') $sheet.Sheets
    Assert-Equal 3 $sheet.Rows.Count '見出しと 2 件'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'M_PRODUCT'!J14" }).Count -eq 1) '定義書のセルへのリンク'
}

Test-Case 'Export-TableDdl: 1 行に 1 カラムずつ書いたインデックス（2 行目からインデックス名が空）とセル内の改行で区切ったカラムを読み、コメントの改行は空白にする' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 }, @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 })
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'T_ORDER'; Columns = $cols; Rows = @(
                @{ Values = @('受注番号', 'ORDER_NO', 'CHAR', '12', '1', '○') },
                @{ Values = @("顧客`nコード", 'CUSTOMER_CD', 'CHAR', '8', '', '○') },
                @{ Values = @('受注日', 'ORDER_DATE', 'DATE', '', '', '○') },
                @{ Values = @('受注ステータス', 'ORDER_STATUS', 'CHAR', '2', '', '○') },
                @{ Values = @('■インデックス定義') },
                @{ Values = @('No', 'インデックス名', 'カラム', 'ユニーク') },
                # インデックス名・No を縦に結合したセル（左上のセルのほかは空）
                @{ Values = @('1', 'IX_T_ORDER_01', 'CUSTOMER_CD', '') },
                @{ Values = @('', '', 'ORDER_DATE DESC', '') },
                @{ Values = @('2', 'IX_T_ORDER_02', "ORDER_STATUS`nORDER_DATE", '') }) })
    $r = Invoke-Ddl @{ Path = $dir; Database = 'PostgreSQL' }
    Assert-Equal 0 $r.Warnings.Count ($r.Warnings -join ' / ')
    Assert-Match $r.Text '(?m)^CREATE INDEX IX_T_ORDER_01 ON T_ORDER \(CUSTOMER_CD, ORDER_DATE DESC\);$'
    Assert-Match $r.Text '(?m)^CREATE INDEX IX_T_ORDER_02 ON T_ORDER \(ORDER_STATUS, ORDER_DATE\);$'
    Assert-Match $r.Text "(?m)^COMMENT ON COLUMN T_ORDER\.CUSTOMER_CD IS '顧客 コード';$"
    Assert-Match (Invoke-Ddl @{ Path = $dir; Database = 'MySQL' }).Text "(?m)^    CUSTOMER_CD\s+CHAR\(8\)\s+NOT NULL COMMENT '顧客 コード',$"
    Assert-Match (Invoke-Ddl @{ Path = $dir; Database = 'SQLServer' }).Text "@value = N'顧客 コード'.*@level2name = N'CUSTOMER_CD';"
}

Test-Case 'Export-TableDdl: NULL許可の列の「可」「許可」「○」は NULL 可、「不可」「×」「NOT NULL」と空は NOT NULL にする' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 }, @{ Header = 'PK'; Width = 5 }, @{ Header = 'NULL許可'; Width = 5 })
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'M_ITEM'; Columns = $cols; Rows = @(
                @{ Values = @('商品コード', 'ITEM_CD', 'CHAR', '10', '1', '不可') },
                @{ Values = @('備考', 'NOTE', 'VARCHAR', '200', '', '可') },
                @{ Values = @('廃止日', 'END_DATE', 'DATE', '', '', '○') },
                @{ Values = @('メモ', 'MEMO', 'VARCHAR', '100', '', '許可') },
                @{ Values = @('商品名', 'ITEM_NAME', 'VARCHAR', '100', '', '不可') },
                @{ Values = @('区分', 'ITEM_KBN', 'CHAR', '1', '', '×') },
                @{ Values = @('単価', 'UNIT_PRICE', 'NUMERIC', '9', '', 'NOT NULL') },
                @{ Values = @('在庫数', 'STOCK_QTY', 'NUMERIC', '7', '', '') }) })
    $r = Invoke-Ddl @{ Path = $dir; Database = 'PostgreSQL' }
    Assert-Equal 0 $r.Warnings.Count ($r.Warnings -join ' / ')
    $notNull = @([regex]::Matches($r.Text, '(?m)^    (\w+) .*NOT NULL,$') | ForEach-Object { $_.Groups[1].Value })
    Assert-Equal @('ITEM_CD', 'ITEM_NAME', 'ITEM_KBN', 'UNIT_PRICE', 'STOCK_QTY') $notNull
}

Test-Case 'Export-TableDdl: 日時・時刻の桁数（秒の小数部）を DB の型に付け、DB の上限を超える桁数と、型に無い桁数（INTEGER などを除く）は使わずに警告する' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 }, @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 })
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'T_LOG'; Columns = $cols; Rows = @(
                @{ Values = @('ID', 'LOG_ID', 'NUMERIC', '10', '1', '○') },
                @{ Values = @('受信日時', 'RECEIVED_AT', 'TIMESTAMP(3)', '', '', '○') },
                @{ Values = @('処理日時', 'PROCESSED_AT', 'DATETIME', '6', '', '○') },
                @{ Values = @('開始時刻', 'START_TIME', 'TIME', '0', '', '') },
                @{ Values = @('登録日時', 'CREATED_AT', 'TIMESTAMP', '', '', '○') },
                @{ Values = @('送信日時', 'SENT_AT', 'TIMESTAMP', '19', '', '') },
                @{ Values = @('件数', 'ROW_COUNT', 'INTEGER', '10', '', '') },
                @{ Values = @('有効', 'ACTIVE_FLG', 'BOOLEAN', '1', '', '') },
                @{ Values = @('基準日', 'BASE_DATE', 'DATE', '8', '', '') },
                @{ Values = @('本文', 'BODY', 'TEXT', '4000', '', '') }) })
    $expected = @{
        Oracle     = 'NUMBER(10)|TIMESTAMP(3)|TIMESTAMP(6)|TIMESTAMP(0)|TIMESTAMP|TIMESTAMP|NUMBER(10)|NUMBER(1)|DATE|CLOB'
        PostgreSQL = 'NUMERIC(10)|TIMESTAMP(3)|TIMESTAMP(6)|TIME(0)|TIMESTAMP|TIMESTAMP|INTEGER|BOOLEAN|DATE|TEXT'
        SQLServer  = 'DECIMAL(10)|DATETIME2(3)|DATETIME2(6)|TIME(0)|DATETIME2|DATETIME2|INT|BIT|DATE|NVARCHAR(MAX)'
        MySQL      = 'DECIMAL(10)|DATETIME(3)|DATETIME(6)|TIME(0)|DATETIME|DATETIME|INT|TINYINT(1)|DATE|LONGTEXT'
    }
    $limit = @{ Oracle = 9; PostgreSQL = 6; SQLServer = 7; MySQL = 6 }
    $sqlType = @{ Oracle = 'CLOB'; PostgreSQL = 'TEXT'; SQLServer = 'NVARCHAR(MAX)'; MySQL = 'LONGTEXT' }
    foreach ($db in @('Oracle', 'PostgreSQL', 'SQLServer', 'MySQL')) {
        $r = Invoke-Ddl @{ Path = $dir; Database = $db }
        $types = @([regex]::Matches($r.Text, '(?m)^    (?!CONSTRAINT )\S+\s+([^\s,]+)') | ForEach-Object { $_.Groups[1].Value })
        Assert-Equal $expected[$db] ($types -join '|') $db
        Assert-Equal (@(
            "テーブル定義書.xlsx T_LOG!B7: T_LOG.SENT_AT の桁数 19 は、$db の秒の小数部の上限（$($limit[$db]) 桁）を超えるため使いません",
            "テーブル定義書.xlsx T_LOG!B11: T_LOG.BODY の桁数 4000 は、データ型 TEXT（$db の $($sqlType[$db])）では使いません") | Sort-Object) ($r.Warnings | Sort-Object) $db
    }
}

Test-Case 'Export-TableDdl: -CompareCatalog は、桁数を使わない型（INTEGER・BIGINT・BOOLEAN）の桁数を DB の精度と比べない' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 }, @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 })
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'T_ITEM'; Columns = $cols; Rows = @(
                @{ Values = @('ID', 'ITEM_ID', 'INTEGER', '10', '1', '○') },
                @{ Values = @('数量', 'QTY', 'BIGINT', '19', '', '○') },
                @{ Values = @('有効', 'ACTIVE_FLG', 'BOOLEAN', '1', '', '○') },
                @{ Values = @('単価', 'UNIT_PRICE', 'NUMERIC', '9', '', '○') }) })
    # 作った DDL どおりの DB（PostgreSQL の integer・bigint の精度は 2 進の 32・64、MySQL の TINYINT(1) は 3）。単価だけ桁数を変えてある
    $pg = Join-Path $dir 'pg.csv'
    @('table_name,column_name,data_type,character_maximum_length,numeric_precision,numeric_scale,is_nullable',
        't_item,item_id,integer,,32,0,NO', 't_item,qty,bigint,,64,0,NO', 't_item,active_flg,boolean,,,,NO', 't_item,unit_price,numeric,,7,0,NO') | Set-Content -LiteralPath $pg -Encoding UTF8
    $my = Join-Path $dir 'my.csv'
    @('TABLE_NAME,COLUMN_NAME,DATA_TYPE,CHARACTER_MAXIMUM_LENGTH,NUMERIC_PRECISION,NUMERIC_SCALE,IS_NULLABLE',
        'T_ITEM,ITEM_ID,int,,10,0,NO', 'T_ITEM,QTY,bigint,,19,0,NO', 'T_ITEM,ACTIVE_FLG,tinyint,,3,0,NO', 'T_ITEM,UNIT_PRICE,decimal,,7,0,NO') | Set-Content -LiteralPath $my -Encoding UTF8
    foreach ($c in @(@('PostgreSQL', $pg), @('MySQL', $my))) {
        $r = Invoke-Ddl @{ Path = $dir; Database = $c[0]; CompareCatalog = $c[1] }
        Assert-Equal 'T_ITEM|UNIT_PRICE|桁数|9,0|7,0' (@($r.Objects | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Table, $_.Column, $_.Item, $_.Definition, $_.Database }) -join ', ') $c[0]
    }
}

Test-Case 'Export-TableDdl: MySQL のコメントでは \ を \\ にする（ほかの DB の文字列では \ はそのまま）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 }, @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 })
    # 半角の円記号（日本語のフォントでは \ が ¥ に見える）
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'M_PRICE'; Columns = $cols; Rows = @(
                @{ Values = @('商品コード', 'PRODUCT_CD', 'CHAR', '10', '1', '○') },
                @{ Values = @('単価(\)', 'UNIT_PRICE', 'NUMERIC', '9', '', '○') },
                @{ Values = @('税込金額\', 'TAX_INCLUDED', 'NUMERIC', '9', '', '○') }) })
    $my = (Invoke-Ddl @{ Path = $dir; Database = 'MySQL' }).Text
    Assert-Match $my ([regex]::Escape("NOT NULL COMMENT '単価(\\)',"))
    Assert-Match $my ([regex]::Escape("NOT NULL COMMENT '税込金額\\'"))
    $pg = (Invoke-Ddl @{ Path = $dir; Database = 'PostgreSQL' }).Text
    Assert-Match $pg ([regex]::Escape("COMMENT ON COLUMN M_PRICE.TAX_INCLUDED IS '税込金額\';"))
}

Test-Case 'Export-TableDdl: 主キー・ユニークインデックスでないカラムへの外部キー、30 バイトを超える制約名（Oracle）、スキーマの中で重なるインデックス名（Oracle・PostgreSQL）を警告する' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $cols = @(@{ Header = '論理名'; Width = 10 }, @{ Header = '物理名'; Width = 10 }, @{ Header = 'データ型'; Width = 10 }, @{ Header = '桁数'; Width = 5 },
        @{ Header = 'PK'; Width = 5 }, @{ Header = '必須'; Width = 5 }, @{ Header = '外部キー参照先'; Width = 20 })
    $ixRows = { param([object[]] $rows) @(@{ Values = @('■インデックス定義') }, @{ Values = @('インデックス名', 'カラム', 'ユニーク') }) + @($rows | ForEach-Object { @{ Values = $_ } }) }
    Export-Workbook -FilePath (Join-Path $dir 'テーブル定義書.xlsx') -Sheets @(
        @{ Name = 'M_CODE'; Columns = $cols; Rows = @(
                @{ Values = @('コード区分', 'CODE_KBN', 'CHAR', '4', '1', '○', '') },
                @{ Values = @('コード値', 'CODE_VALUE', 'CHAR', '2', '2', '○', '') },
                @{ Values = @('名称', 'CODE_NAME', 'VARCHAR', '20', '', '○', '') }) + (& $ixRows @(, @('IX_01', 'CODE_NAME', ''))) },
        @{ Name = 'M_WAREHOUSE'; Columns = $cols; Rows = @(
                @{ Values = @('倉庫ID', 'WAREHOUSE_ID', 'INTEGER', '', '1', '○', '') },
                @{ Values = @('倉庫コード', 'WAREHOUSE_CD', 'CHAR', '4', '', '○', '') },
                @{ Values = @('倉庫名', 'WAREHOUSE_NAME', 'VARCHAR', '20', '', '○', '') }) + (& $ixRows @(@('UX_M_WAREHOUSE_01', 'WAREHOUSE_CD', '○'), @('IX_01', 'WAREHOUSE_NAME', ''))) },
        @{ Name = 'T_SHIPMENT_INSTRUCTION_DETAIL'; Columns = $cols; Rows = @(
                @{ Values = @('出荷指示番号', 'SHIPMENT_NO', 'CHAR', '12', '1', '○', '') },
                @{ Values = @('出荷状態', 'SHIP_STATUS', 'CHAR', '2', '', '○', 'M_CODE.CODE_VALUE') },
                @{ Values = @('倉庫コード', 'WAREHOUSE_CD', 'CHAR', '4', '', '○', 'M_WAREHOUSE.WAREHOUSE_CD') }) })
    $file = 'テーブル定義書.xlsx'
    $detail = 'T_SHIPMENT_INSTRUCTION_DETAIL'
    $fkTarget = "${file} ${detail}!B3: ${detail} (SHIP_STATUS) の外部キーの参照先 M_CODE (CODE_VALUE) が、M_CODE の主キーでもユニークインデックスでもありません"
    $r = Invoke-Ddl @{ Path = $dir; Database = 'Oracle' }
    Assert-Equal (@(
        "${file} M_WAREHOUSE!A8: インデックス名「IX_01」はテーブル M_CODE のインデックスにも使っています（Oracle ではスキーマの中で同じインデックス名を使えません）",
        "${file} ${detail}!B1: 主キーの制約名「PK_${detail}」が 30 バイトを超えています（Oracle 12.1 以前では使えません）",
        "${file} ${detail}!B1: 外部キーの制約名「FK_${detail}_M_CODE」が 30 バイトを超えています（Oracle 12.1 以前では使えません）",
        "${file} ${detail}!B1: 外部キーの制約名「FK_${detail}_M_WAREHOUSE」が 30 バイトを超えています（Oracle 12.1 以前では使えません）",
        $fkTarget) | Sort-Object) ($r.Warnings | Sort-Object)
    $r = Invoke-Ddl @{ Path = $dir; Database = 'PostgreSQL' }
    Assert-Equal (@(
        "${file} M_WAREHOUSE!A8: インデックス名「IX_01」はテーブル M_CODE のインデックスにも使っています（PostgreSQL ではスキーマの中で同じインデックス名を使えません）",
        $fkTarget) | Sort-Object) ($r.Warnings | Sort-Object)
    # SQL Server・MySQL のインデックス名はテーブルごと
    Assert-Equal $fkTarget ((Invoke-Ddl @{ Path = $dir; Database = 'SQLServer' }).Warnings -join ' / ')
}

Test-Case 'Export-TableDdl: Get-Help でヘルプが表示される' {
    $h = Get-Help $DdlTool
    Assert-Match $h.Synopsis 'DDL'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'Export-TableDdl: 同じテーブル定義書の古い版と、雛形のシート（テーブル ID が差し込み欄のまま）・非表示のシートからは DDL を作らない' {
    $dir = New-TestDir
    # 1 行目（表の見出しの行）の数だけ列を書くので、1 行目も 5 列にする
    $columns = @('論理名', '物理名', 'データ型', '桁数', 'PK')
    foreach ($v in @(@('v1.0', '500'), @('v1.1', '1000'))) {
        $book = Join-Path $dir ('テーブル定義書_{0}.xlsx' -f $v[0])
        New-SheetsBook $book ([ordered]@{
                'T_ORDER' = @(@('テーブルID', 'T_ORDER', 'テーブル名', '受注', ''), $columns, @('受注番号', 'ORDER_NO', 'CHAR', '12', '1'), @('備考', 'REMARKS', 'VARCHAR', $v[1], ''))
                'T_OLD'   = @(@('テーブルID', 'T_OLD', 'テーブル名', '使わなくなったテーブル', ''), $columns, @('番号', 'NO', 'CHAR', '1', '1'))
                '雛形'    = @(@('テーブルID', '{{テーブルID}}', 'テーブル名', '{{テーブル名}}', ''), $columns, @('（主キー）', '{{主キー}}', 'CHAR', '10', '1'))
            })
        Edit-ZipText $book 'xl/workbook.xml' { param($t) $t -replace '(<sheet name="T_OLD")', '$1 state="hidden"' }
    }
    $r = Invoke-Ddl @{ Path = $dir }
    Assert-Equal @('T_ORDER') @([regex]::Matches($r.Text, '(?m)^CREATE TABLE (\S+)') | ForEach-Object { $_.Groups[1].Value })
    Assert-Match $r.Text '(?m)^    REMARKS\s+VARCHAR2\(1000 CHAR\),?$' '新しい版（v1.1）の桁数'
    Assert-Match $r.Text '元のファイル: テーブル定義書_v1\.1\.xlsx'
    Assert-Equal 2 $r.Warnings.Count ($r.Warnings -join ' / ')
    Assert-Match $r.Warnings[0] '^古い版 1 件は読みませんでした（テーブル定義書_v1\.0\.xlsx → テーブル定義書_v1\.1\.xlsx。'
    Assert-Match $r.Warnings[1] '^雛形（差し込み欄のまま）のシート・ブック 1 件は読みませんでした（テーブル定義書_v1\.1\.xlsx 雛形（テーブルID が \{\{テーブルID\}\}））'
}
