#     Test-CodeTrace.ps1 のテスト。設計書は docs\design-sample、ソースはテストの中で作る

$TraceTool = Join-Path $script:ToolsDir 'Test-CodeTrace.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function New-SourceTree {
    #     ソースのフォルダ（Java・SQL・MyBatis・メッセージのリソース。コメントの行と、見ないフォルダの中のファイルも）
    param([string] $Dir)
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $write = { param($rel, [string[]] $lines) $p = Join-Path $Dir $rel; $d = Split-Path -Parent $p; if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null }; [System.IO.File]::WriteAllText($p, (($lines -join "`n") + "`n"), $utf8) }
    & $write 'src/OrderList.java' @(
        'package sample;',
        '// E0003 はもう使わない',
        '/* SCR003 の画面は',
        '   別のクラス */',
        'public class OrderList {',
        '    static final String SCREEN = "SCR001";',
        '    void search() {',
        '        if (orderNo == null) { error("E0001"); }',
        '        error("E0099");',
        '        String sql = "SELECT * FROM T_ORDER o JOIN T_ORDR x ON o.ORDER_NO = x.ORDER_NO";',
        '        String note = "copy from source";',
        '    }',
        '}')
    & $write 'src/messages.properties' @('E0001={0}は必須入力です。', 'E0002={0}は{1}文字以内で入力してください。', '# E0004 は使わない')
    & $write 'src/sql/order.sql' @('select d.order_no, p.product_name', '  from t_order_detail d', '  join m_product p on p.product_cd = d.product_cd;', '-- from m_customer', 'select extract(year from order_date) from t_order;')
    & $write 'src/OrderMapper.xml' @('<mapper namespace="sample.OrderMapper">', '  <select id="history">select * from t_order_hist where order_no = #{orderNo}</select>', '</mapper>')
    & $write 'src/config/app.xml' @('<app>', '  <description>copy from source</description>', '</app>')
    & $write 'src/node_modules/lib/a.js' @('var x = "E0005";')
}

Test-Case 'Test-CodeTrace: 設計書のメッセージ・画面・テーブルの ID がソースに無いものと、ソースにあって設計書に無いものを見つける（コメントの行は見ない）' {
    $dir = New-TestDir
    New-SourceTree $dir
    $issues = @(& $TraceTool -DesignPath $designDir -SourcePath (Join-Path $dir 'src') 6>$null)
    $keys = @($issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Item, $_.Kind, $_.Id })
    Assert-Equal 20 $keys.Count ($keys -join "`n")
    Assert-Equal @('実装に無い|メッセージ|E0003', '実装に無い|メッセージ|E0004', '実装に無い|メッセージ|E0005') @($keys | Select-Object -First 3) 'コメントの行（// /* # --）と、見ないフォルダ（node_modules）の ID は数えない'
    Assert-True ($keys -notcontains '実装に無い|メッセージ|E0002') 'メッセージのリソース（.properties）にある'
    Assert-True ($keys -contains '実装に無い|画面|SCR003') 'ブロックコメントの中の画面 ID は数えない'
    Assert-True ($keys -notcontains '実装に無い|画面|SCR001')
    Assert-True ($keys -contains '実装に無い|テーブル|M_CUSTOMER') '.sql のコメントの行（-- from m_customer）'
    Assert-Equal @('設計書に無い|メッセージ|E0099', '設計書に無い|テーブル|T_ORDR', '設計書に無い|テーブル|t_order_hist') @($keys | Select-Object -Last 3) '.sql と MyBatis の .xml は小文字の SQL も見る。ほかのファイルの小文字の from（copy from source）と、FROM の後ろのカラム名（extract(year from order_date)）はテーブルにしない'
    $e0099 = @($issues | Where-Object Id -eq 'E0099')[0]
    Assert-Equal 'ソースのメッセージ ID E0099（OrderList.java 9 行目）が、メッセージ一覧にありません' ($e0099.Detail -replace '\\', '/' -replace 'src/', '')
    Assert-Equal '警告' $e0099.Level

    $text = Get-HostText { & $TraceTool -DesignPath $designDir -SourcePath (Join-Path $dir 'src') | Out-Null }
    Assert-Match $text 'ソース 5 ファイルで、設計書の ID が見つかったもの: メッセージ 2/16・画面 1/3・テーブル 3/4'
    Assert-Match $text '設計書に無い ID: メッセージ 1・テーブル 2'
}

Test-Case 'Test-CodeTrace: -Kind Column で、カラムの物理名を ORM の名前（orderNo など）でも探し、見つからないカラムを情報で知らせる' {
    $dir = New-TestDir
    New-SourceTree $dir
    $issues = @(& $TraceTool -DesignPath (Join-Path $designDir 'DD03_テーブル定義書.xlsx') -SourcePath (Join-Path $dir 'src') -Kind Column 6>$null)
    $ids = @($issues | ForEach-Object { $_.Id })
    Assert-True ($ids -notcontains 'T_ORDER.ORDER_NO') 'orderNo（Java）と ORDER_NO（SQL）'
    Assert-True ($ids -notcontains 'M_PRODUCT.PRODUCT_NAME') 'product_name（.sql）'
    Assert-True ($ids -contains 'T_ORDER.TOTAL_AMOUNT')
    Assert-Equal @('情報') @($issues | ForEach-Object { $_.Level } | Select-Object -Unique)
}

Test-Case 'Test-CodeTrace: -OutXlsx に指摘と、ID ごとのソースの箇所の対応表を保存する' {
    $dir = New-TestDir
    New-SourceTree $dir
    $out = Join-Path $dir 'trace.xlsx'
    $null = & $TraceTool -DesignPath $designDir -SourcePath (Join-Path $dir 'src') -OutXlsx $out 6>$null
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('指摘', '対応表') $book.Sheets
    Assert-True (@($book.Links | Where-Object { $_.Target -like '*DD05_*.xlsx' -and $_.Location -like "'メッセージ一覧'!A*" }).Count -gt 0) '設計書の行へのリンク'
    $trace = Read-XlsxSheet -Path $out -Index 2
    $row = @(2..($trace.Rows.Count) | Where-Object { $trace.Rows[$_]['B'] -eq 'E0001' })[0]
    Assert-Equal 'メッセージ|E0001|2|2' ('{0}|{1}|{2}|{3}' -f $trace.Rows[$row]['A'], $trace.Rows[$row]['B'], $trace.Rows[$row]['E'], $trace.Rows[$row]['F']) 'Java とメッセージのリソースの 2 か所'
    Assert-Equal 'messages.properties 1 行目 ほか 1 か所' $trace.Rows[$row]['G'] 'ソースのフォルダからの相対パスの順で、最初の箇所'
}

Test-Case 'Test-CodeTrace: 文字列の中の /*（"/api/**" など）と行コメントの中の /* でコメントを始めず、行の途中から始まるブロックコメントの中の ID は数えない' {
    $dir = New-TestDir
    $src = Join-Path $dir 'src'
    $null = New-Item -ItemType Directory -Path $src
    [System.IO.File]::WriteAllText((Join-Path $src 'WebConfig.java'), ((@(
                'public class WebConfig {',
                '    void addCorsMappings(CorsRegistry registry) { registry.addMapping("/api/**"); }',
                '    void search() {',
                '        if (orderNo == null) { error("E0001"); }',
                '        forward("SCR001"); // e.g. /api/*',
                '        String sql = "SELECT * FROM T_ORDER";',
                '        int x = 1; /* a */ int y = 2; /* old:',
                '        error("E0003");',
                '        */',
                '    }',
                '}') -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
    $issues = @(& $TraceTool -DesignPath $designDir -SourcePath $src 6>$null)
    $missing = @($issues | Where-Object Item -eq '実装に無い' | ForEach-Object { $_.Id })
    Assert-Equal '' (@($missing | Where-Object { @('E0001', 'SCR001', 'T_ORDER') -contains $_ }) -join ', ') '"/api/**" と // e.g. /api/* の後ろの行の ID'
    Assert-True ($missing -contains 'E0003') 'int y = 2; /* old: から次の行に続くブロックコメントの中の E0003'
}

Test-Case 'Test-CodeTrace: SELECT … INTO の変数・FOR UPDATE NOWAIT・MyBatis のタグ（<update id=…>）を SQL のテーブルにしない' {
    $dir = New-TestDir
    $src = Join-Path $dir 'src'
    $null = New-Item -ItemType Directory -Path $src
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $src 'OrderMapper.xml'), ((@(
                '<mapper namespace="sample.OrderMapper">',
                '  <update id="updateStatus">',
                '    UPDATE T_ORDER SET ORDER_STATUS = #{status} WHERE ORDER_NO = #{orderNo}',
                '  </update>',
                '  <insert id="addHistory">INSERT INTO T_ORDER_HIST SELECT * FROM T_ORDER</insert>',
                '  <select id="countLarge"><![CDATA[ SELECT COUNT(*) FROM T_ORDER_BIG WHERE TOTAL_AMOUNT > 100 ]]></select>',
                '</mapper>') -join "`n") + "`n"), $utf8)
    [System.IO.File]::WriteAllText((Join-Path $src 'order_pkg.pkb'), ((@(
                'CREATE OR REPLACE PACKAGE BODY order_pkg AS',
                '  PROCEDURE cancel(p_no IN VARCHAR2) IS',
                '    v_count NUMBER;',
                '  BEGIN',
                '    SELECT COUNT(*) INTO v_count FROM T_ORDER WHERE ORDER_NO = p_no;',
                '    SELECT * INTO r_order FROM T_ORDER WHERE ORDER_NO = p_no FOR UPDATE NOWAIT;',
                '    SELECT ORDER_NO',
                '      INTO v_no',
                '      FROM T_ORDER WHERE ORDER_NO = p_no FOR UPDATE OF ORDER_STATUS SKIP LOCKED;',
                '    INSERT /*+ APPEND */ INTO T_ORDER_LOG (ORDER_NO) VALUES (p_no);',
                '    MERGE INTO T_ORDER_SUM s USING T_ORDER o ON (s.ORDER_NO = o.ORDER_NO)',
                '      WHEN MATCHED THEN UPDATE SET s.CNT = s.CNT + 1;',
                '    UPDATE T_ORDR SET ORDER_STATUS = ''90'' WHERE ORDER_NO = p_no;',
                '  END;',
                'END order_pkg;') -join "`n") + "`n"), $utf8)
    $issues = @(& $TraceTool -DesignPath (Join-Path $designDir 'DD03_テーブル定義書.xlsx') -SourcePath $src -Kind Table 6>$null)
    Assert-Equal @('T_ORDER_HIST', 'T_ORDER_BIG', 'T_ORDER_LOG', 'T_ORDER_SUM', 'T_ORDR') @($issues | Where-Object Item -eq '設計書に無い' | ForEach-Object { $_.Id }) 'INSERT・MERGE の INTO（ヒントの後ろも）・UPDATE の後ろと、CDATA の中の SQL のテーブルは見る'
}

Test-Case 'Test-CodeTrace: 行をまたぐ SQL のテーブルも見て、ON UPDATE CASCADE・ON DUPLICATE KEY UPDATE・トリガーの UPDATE ON はテーブルにせず、行の後ろのコメントの ID は数えない' {
    $dir = New-TestDir
    $src = Join-Path $dir 'src'
    $null = New-Item -ItemType Directory -Path $src
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText((Join-Path $src 'order.sql'), ((@(
                'SELECT o.ORDER_NO',
                '  FROM',
                '    T_ORDR o',
                '  JOIN /* 明細 */ T_ORDER_DETAL d ON d.ORDER_NO = o.ORDER_NO;',
                'INSERT',
                '  INTO T_ORDER_LOGG (ORDER_NO) VALUES (1);',
                'UPDATE',
                '  -- 受注',
                '    T_ODER SET ORDER_STATUS = ''90'';',
                'ALTER TABLE T_ORDER_DETAIL ADD FOREIGN KEY (ORDER_NO) REFERENCES T_ORDER (ORDER_NO) ON DELETE CASCADE ON UPDATE CASCADE;',
                'INSERT INTO T_ORDER (ORDER_NO) VALUES (1) ON DUPLICATE KEY UPDATE UPD_COUNT = UPD_COUNT + 1;',
                'CREATE TRIGGER TR_ORDER BEFORE INSERT OR UPDATE ON T_ORDER FOR EACH ROW BEGIN NULL; END;') -join "`n") + "`n"), $utf8)
    # ほかのファイルは行をまたがない（行の終わりの UPDATE と、次の行のキー）
    [System.IO.File]::WriteAllText((Join-Path $src 'messages.properties'), "confirm.title=Do you want to UPDATE`nnext.key=value`n", $utf8)
    [System.IO.File]::WriteAllText((Join-Path $src 'OrderList.java'), ((@(
                'class OrderList {',
                '    void f() { error("E0001"); } // E0002 は使わなくなった',
                '    void g() { forward("SCR001"); /* SCR002 へは行かない */ }',
                '}') -join "`n") + "`n"), $utf8)
    $issues = @(& $TraceTool -DesignPath $designDir -SourcePath $src 6>$null)
    Assert-Equal @('T_ORDR|3', 'T_ORDER_DETAL|4', 'T_ORDER_LOGG|6', 'T_ODER|9') @($issues | Where-Object Item -eq '設計書に無い' | ForEach-Object { '{0}|{1}' -f $_.Id, ([regex]::Match($_.Detail, '(\d+) 行目').Groups[1].Value) }) 'テーブル名のある行。CASCADE・UPD_COUNT・ON はテーブルにしない'
    $missing = @($issues | Where-Object Item -eq '実装に無い' | ForEach-Object { $_.Id })
    Assert-True ($missing -contains 'E0002' -and $missing -contains 'SCR002') '行の後ろの // と /* */ の中の ID は数えない'
    Assert-True ($missing -notcontains 'E0001' -and $missing -notcontains 'SCR001')
}

Test-Case 'Test-CodeTrace: - や _ で区切ったメッセージ ID（MSG-E001・MSG_E002）も突き合わせる' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    $src = Join-Path $dir 'src'
    $null = New-Item -ItemType Directory -Path $design, $src
    New-SheetsBook (Join-Path $design 'メッセージ一覧.xlsx') ([ordered]@{ 'メッセージ一覧' = @(@('メッセージID', 'メッセージ'), @('MSG-E001', '必須です。'), @('MSG_E002', '桁数が多すぎます。'), @('E0003', '形式が違います。')) })
    [System.IO.File]::WriteAllText((Join-Path $src 'A.java'), "class A {`n  void f() { error(`"MSG-E001`"); error(`"MSG_E002`"); error(`"MSG-E099`"); }`n}`n", (New-Object System.Text.UTF8Encoding($false)))
    $issues = @(& $TraceTool -DesignPath $design -SourcePath $src -Kind Message 6>$null)
    Assert-Equal @('実装に無い|E0003', '設計書に無い|MSG-E099') @($issues | ForEach-Object { '{0}|{1}' -f $_.Item, $_.Id })
}

Test-Case 'Test-CodeTrace: Get-Help でヘルプが表示される' {
    $h = Get-Help $TraceTool
    Assert-Match $h.Synopsis 'ソースコード'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
