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

Test-Case 'Export-TableDdl: Get-Help でヘルプが表示される' {
    $h = Get-Help $DdlTool
    Assert-Match $h.Synopsis 'DDL'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
