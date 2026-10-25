#     ExcelTools.psm1 の共通の部品のテスト（ツールのテストで確かめにくいもの）

Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking

Test-Case 'ExcelTools: Group-InOrder は、まとまりを最初に出てきた順に返す（PowerShell の版によらない）' {
    $items = @(
        [pscustomobject]@{ Process = '結合テスト'; No = 1 },
        [pscustomobject]@{ Process = '単体テスト'; No = 2 },
        [pscustomobject]@{ Process = '結合テスト'; No = 3 },
        [pscustomobject]@{ Process = 'abc'; No = 4 },
        [pscustomobject]@{ Process = 'ABC'; No = 5 })
    $groups = @($items | Group-InOrder Process)
    Assert-Equal @('結合テスト|2|1,3|0', '単体テスト|1|2|1', 'abc|2|4,5|2') @($groups | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Name, $_.Count, (@($_.Group | ForEach-Object { $_.No }) -join ','), $_.Index }) 'Group-Object と同じく大文字小文字は区別しない'
    $byBlock = @($items | Group-InOrder { $_.No % 2 })
    Assert-Equal @('1|3', '0|2') @($byBlock | ForEach-Object { '{0}|{1}' -f $_.Name, $_.Count }) 'スクリプトブロックで値を決める'
    $rows = @(@{ Values = @('a', 'OK') }, @{ Values = @('b', 'NG') }, @{ Values = @('c', 'OK') })
    Assert-Equal @('OK 2', 'NG 1') @($rows | Group-InOrder { $_.Values[1] } | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) '配列を持つ要素も 1 つとして扱う'
    Assert-Equal 0 @(@() | Group-InOrder Process).Count
}

Test-Case 'ExcelTools: Get-TestResultKind は、結果の丸（○ 記号・〇 漢数字・◯ 大きな丸）をどれも OK にする' {
    Assert-Equal @('OK', 'OK', 'OK', 'OK', 'NG', '対象外', '保留', '未実施') @(@('OK', '○', '〇', '◯', '×', '-', '保留', '') | ForEach-Object { Get-TestResultKind $_ })
}

Test-Case 'ExcelTools: Get-TestResultKind は、チェックの印（✓・☑・◎）と済・完了を OK に、× に似た印（✖・✗・X）を NG にする' {
    Assert-Equal @('OK', 'OK', 'OK', 'OK', 'OK', 'OK', 'OK', 'OK') @(@('✓', '✔', '☑', '◎', '済', '実施済み', '確認済', '完了') | ForEach-Object { Get-TestResultKind $_ }) 'OK'
    Assert-Equal @('NG', 'NG', 'NG', 'NG', 'NG', 'NG') @(@('✖', '✗', '✘', '☓', 'X', 'ｘ') | ForEach-Object { Get-TestResultKind $_ }) 'NG'
    Assert-Equal @('未実施', '未実施', '未実施') @(@('未完了', '未済', 'XX') | ForEach-Object { Get-TestResultKind $_ }) '「未」で始まるもの・X だけでないもの'
}

Test-Case 'ExcelTools: Test-ClosedStatus は、完了・済を含んでも「未完了」「未済」「未解決」「完了予定」「完了待ち」「差し戻し」「not done」は終わったものにしない' {
    $pattern = '完了|済|クローズ|close|done|解決|対応不要'
    Assert-Equal @($true, $true, $true, $true, $true, $true) @(@('完了', '済', '対応済み', '解決', 'Closed', 'done') | ForEach-Object { Test-ClosedStatus $_ $pattern }) '終わったもの'
    Assert-Equal @($false, $false, $false, $false, $false, $false, $false, $false, $false, $false) @(@('未完了', '未済', '未解決', '完了予定', '完了待ち', '対応済（差し戻し）', 'Not done', 'reopen', '対応中', '') | ForEach-Object { Test-ClosedStatus $_ $pattern }) '終わっていないもの'
}

Test-Case 'ExcelTools: Test-ClosedStatus は、「終了」を終わったものにし、「未終了」「終了予定」「終了待ち」は終わったものにしない' {
    $pattern = '完了|済|クローズ|close|done|解決|対応不要|終了'
    Assert-Equal @($true, $true, $true) @(@('終了', '対応終了', 'Closed') | ForEach-Object { Test-ClosedStatus $_ $pattern }) '終わったもの'
    Assert-Equal @($false, $false, $false) @(@('未終了', '終了予定', '終了待ち') | ForEach-Object { Test-ClosedStatus $_ $pattern }) '終わっていないもの'
}

Test-Case 'ExcelTools: Read-CsvFile は、見出しが「#」で始まる CSV（Redmine の書き出し）も読み、「#」の列を「＃」にする（「#TYPE」の行は読み飛ばす）' {
    $dir = New-TestDir
    # 「#」で始まる行は、ConvertFrom-Csv が注釈として読み飛ばす（引用符で囲んでも同じ）
    @('#,トラッカー,ステータス,題名', '101,バグ,終了,一覧の並び順が逆', '102,バグ,新規,登録できない') | Set-Content -LiteralPath (Join-Path $dir 'redmine.csv') -Encoding UTF8
    @('"#","題名"', '"101","一覧の並び順が逆"') | Set-Content -LiteralPath (Join-Path $dir 'quoted.csv') -Encoding UTF8
    $rows = Read-CsvFile (Join-Path $dir 'redmine.csv') @('題名') '課題管理表'
    Assert-Equal 2 $rows.Count '行の数'
    Assert-Equal '＃|トラッカー|ステータス|題名' ((@($rows[0].PSObject.Properties | ForEach-Object { $_.Name })) -join '|')
    Assert-Equal '101|102' (($rows | ForEach-Object { $_.'＃' }) -join '|')
    $rows = Read-CsvFile (Join-Path $dir 'quoted.csv') @('題名') '課題管理表'
    Assert-Equal '101|一覧の並び順が逆' ('{0}|{1}' -f $rows[0].'＃', $rows[0].'題名') '引用符で囲んだ「#」'
    # Windows PowerShell の Export-Csv が書く型の行は、これまでどおり読み飛ばす
    @('#TYPE System.Management.Automation.PSCustomObject', '"No","題名"', '"1","一覧の並び順が逆"') | Set-Content -LiteralPath (Join-Path $dir 'typed.csv') -Encoding UTF8
    $rows = Read-CsvFile (Join-Path $dir 'typed.csv') @('No', '題名') '課題管理表'
    Assert-Equal '1|1|一覧の並び順が逆' ('{0}|{1}|{2}' -f $rows.Count, $rows[0].No, $rows[0].'題名') '#TYPE の行'
}

Test-Case 'ExcelTools: Get-LeafName は、\ と（拡張子で終わる名前の）/ でフォルダを除き、「外部I/F設計書」の / は残す（PowerShell 5.1 と 7 で同じ）' {
    Assert-Equal @('基本設計書.xlsx', '基本設計書.xlsx', '基本設計書', '外部I/F設計書', '外部I/F設計書 v1.10', '') @(@('D:\納品物\基本設計書.xlsx', '設計書/基本設計書.xlsx', ' sub\基本設計書 ', '外部I/F設計書', '外部I/F設計書 v1.10', '') | ForEach-Object { Get-LeafName $_ })
    Assert-Equal (ConvertTo-NameKey '基本設計書') (ConvertTo-NameKey 'D:\納品物\基本設計書.xlsx') 'フォルダと拡張子を除く'
    Assert-True ((ConvertTo-NameKey '外部I/F設計書') -ne (ConvertTo-NameKey '内部I/F設計書')) '「I/F」の前が違う名前は別のもの'
}

Test-Case 'ExcelTools: Find-DesignTables は「■」の行で表を終える。-List（一覧の表）では終えずに、値がすべて「■」で始まる行とキーの列が見出しと同じ行を飛ばす' {
    # 2 行目は縦に結合した見出し（No）の下の行、4 行目はグループの見出し、5 行目は「■」で始まる指摘
    $hits = foreach ($c in @('A1=No', 'B1=指摘内容', 'C1=対応状況', 'A2=No', 'B2=内容', 'C2=状態', 'A3=1', 'B3=型が違う', 'C3=対応中', 'A4=■索引', 'A5=2', 'B5=■索引: 一意索引が足りない', 'C5=対応中', 'A6=3', 'B6=桁数が違う', 'C6=完了')) {
        $p = $c.Split('=')
        [pscustomobject]@{ File = '指摘.xlsx'; Path = '指摘.xlsx'; Sheet = '指摘'; Cell = $p[0]; Text = $p[1] }
    }
    $grid = (ConvertTo-DesignBooks $hits)['指摘.xlsx'].Sheets['指摘']
    $rows = { param($tables) @($tables | ForEach-Object { (@($_.Rows | ForEach-Object { $_.Row })) -join ',' }) }
    Assert-Equal @('2,3') (& $rows (Find-DesignTables $grid @('No', '指摘内容', '対応状況') 'No')) '設計書の表は、これまでどおり「■」の行（次の節）まで'
    Assert-Equal @('3,5,6') (& $rows (Find-DesignTables $grid @('No', '指摘内容', '対応状況') 'No' -List))
}

Test-Case 'ExcelTools: Read-BookDrawings は、位置で置いた図形・画像（absoluteAnchor）も、その位置のセルとセルの左上からのずれで読む（シートの上の位置は置いた位置のまま）' {
    $path = Join-Path (New-TestDir) '画面.xlsx'
    New-SheetsBook $path ([ordered]@{ '画面' = @(, @('項目', '値', '備考')) })
    Add-TestDrawing $path @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Screen'; From = @(1, 2); To = @(8, 20) })
    # A〜C 列は 70 ピクセル、行は 24 ピクセル。左から 100・上から 60 ピクセル（B3 の中）に、400 × 300 ピクセルで置く
    Edit-ZipText $path 'xl/drawings/drawing1.xml' { param($t) $t -replace '<xdr:twoCellAnchor>.*?</xdr:to>', '<xdr:absoluteAnchor><xdr:pos x="952500" y="571500"/><xdr:ext cx="3810000" cy="2857500"/>' -replace '</xdr:twoCellAnchor>', '</xdr:absoluteAnchor>' }
    $items = (Read-BookDrawings $path -Bounds)['画面'].ToArray()
    Assert-Equal @('図 1|pic|B3|3|2|30|12|400|300') @($items | ForEach-Object { $a = $_.Anchor; '{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $_.Name, $_.Kind, $_.Cell, $a.R1, $a.C1, $a.OX, $a.OY, $a.W, $a.H })
    Assert-Equal '100|60|400|300' ('{0}|{1}|{2}|{3}' -f $items[0].Bounds.X, $items[0].Bounds.Y, $items[0].Bounds.W, $items[0].Bounds.H)
}

Test-Case 'ExcelTools: Read-BookDrawings は、図形の段落の中の改行（<a:br>。Shift+Enter の改行）も、段落の区切りと同じく改行にして読む（Search-Excel と同じ）' {
    $path = Join-Path (New-TestDir) '図形.xlsx'
    New-SheetsBook $path ([ordered]@{ 'S' = @(@('項目'), @('x')) })
    Add-TestDrawing $path @(@{ Kind = 'sp'; Name = '図1'; Text = 'X'; From = @(1, 1); To = @(4, 4) })
    Edit-ZipText $path 'xl/drawings/drawing1.xml' { param($x) $x.Replace('<a:p><a:r><a:t>X</a:t></a:r></a:p>', '<a:p><a:r><a:t>受注</a:t></a:r><a:br/><a:r><a:t>管理システム</a:t></a:r></a:p><a:p><a:r><a:t>概要</a:t></a:r></a:p>') }
    $items = (Read-BookDrawings $path)['S'].ToArray()
    # 改行は ↵ にして比べる
    Assert-Equal @('図1|受注↵管理システム↵概要') @($items | ForEach-Object { '{0}|{1}' -f $_.Name, $_.Text.Replace("`n", '↵') })
}

Test-Case 'ExcelTools: Export-Workbook は日付（datetime）を日付のセルとして書き、Search-Excel では yyyy/MM/dd で読める' {
    $path = Join-Path (New-TestDir) '日付.xlsx'
    Export-Workbook -FilePath $path -Sheets @(@{ Name = '予定'; Columns = @(@{ Header = '開始'; Width = 12 }, @{ Header = '工数'; Width = 6 }, @{ Header = 'メモ'; Width = 10 }); Rows = @(@{ Values = @([datetime]'2026-11-02', 1.5, '2026/11/03') }) })
    $hits = @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern '\S' -Path $path -Target Value 6>$null)
    Assert-Equal 'A2=2026/11/02|B2=1.5|C2=2026/11/03' ((@($hits | Where-Object { $_.Cell -match '2$' } | ForEach-Object { '{0}={1}' -f $_.Cell, $_.Text })) -join '|') '日付のセル・数値のセル・文字のセル'
    $sheet = Read-XlsxSheet -Path $path
    Assert-Equal ([string](([datetime]'2026-11-02') - [datetime]'1899-12-30').TotalDays) ([string]$sheet.Rows[2]['A']) '日付はシリアル値（文字ではない）'
}

Test-Case 'ExcelTools: Measure-SourceLines の CodeText は、行の後ろのコメントと行の中のブロックコメントを除く（文字列の中の // や # は残す。行の数え方は変わらない）' {
    $dir = New-TestDir
    # コメントの行数 / ステップの行の行番号 / ステップの行の CodeText（空白は 1 つにそろえる）
    $measure = {
        param([string] $Name, [string[]] $Lines)
        $path = Join-Path $dir $Name
        [System.IO.File]::WriteAllText($path, (($Lines -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
        $m = Measure-SourceLines $path (Get-SourceLanguage ([System.IO.Path]::GetExtension($Name))) -CodeText
        '{0} / {1} / {2}' -f $m.Comment, (@($m.CodeLines | ForEach-Object { $_.No }) -join ','), (@($m.CodeLines | ForEach-Object { $_.CodeText -replace '\s+', ' ' }) -join ' | ')
    }
    Assert-Equal '1 / 1,2,3 / error("E0001"); } | String u = "http://x/#a"; int y = 2; | forward("SCR001");' (& $measure 'A.java' @(
            'error("E0001"); } // E0002',
            'String u = "http://x/#a"; /* E0003 */ int y = 2;',
            '/* E0004 */ forward("SCR001"); /* E0005',
            'E0006 */')) 'Java'
    Assert-Equal '1 / 1,2 / SELECT ''--'' AS A FROM T_ORDER | SELECT 1 FROM DUAL;' (& $measure 'a.sql' @(
            "SELECT '--' AS A FROM T_ORDER -- T_ORDR",
            'SELECT 1 /* T_X */ FROM DUAL;',
            '-- T_Y')) 'SQL'
    Assert-Equal '1 / 1,2 / x = "#E0001" | y = ''a#b''' (& $measure 'a.py' @(
            'x = "#E0001"  # E0002',
            "y = 'a#b'",
            '# E0003')) 'Python'
    Assert-Equal '2 / 1,2,3 / MsgBox "E0001" | x = 1 : | REMARK = "it''s"' (& $measure 'A.vb' @(
            'MsgBox "E0001" '' E0002',
            'x = 1 : REM E0003',
            'REMARK = "it''s"',
            ''' E0004',
            'REM E0005')) 'VB（REM は単語の始まりだけ）'
    Assert-Equal '1 / 2 / echo "E0001 # x"' (& $measure 'a.sh' @(
            '#!/bin/sh',
            'echo "E0001 # x" # E0002')) 'シェル'
}

Test-Case 'ExcelTools: Measure-SourceLines は、文字列と行コメントの中の /* をコメントの始まりにせず、1 行の中のブロックコメントをすべて見る' {
    $dir = New-TestDir
    # コメントの行数と、ステップの行の行番号
    $measure = {
        param([string] $Name, [string[]] $Lines)
        $path = Join-Path $dir $Name
        [System.IO.File]::WriteAllText($path, (($Lines -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
        $m = Measure-SourceLines $path (Get-SourceLanguage ([System.IO.Path]::GetExtension($Name)))
        '{0}|{1}' -f $m.Comment, (@($m.CodeLines | ForEach-Object { $_.No }) -join ',')
    }
    Assert-Equal '3|1,2,3,4,5,7,10,11' (& $measure 'WebConfig.java' @(
            'class WebConfig {',
            '    void a() { registry.addMapping("/api/**"); }',
            '    String glob = "**/*.java"; char q = ''"''; String s = "\"/*";',
            '    void b() { route(path); } // e.g. /api/*',
            '    int x = 1; /* a */ int y = 2; /* old:',
            '    error("E0003");',
            '    */ int z = 3; /* 1 */ /* 2 */',
            '    /* a */ /* b',
            '    c */ // d',
            '    void f() { error("E0001"); }',
            '}')) '文字列・文字・行コメントの中の /* は見ない。コードの後ろの 2 つ目のブロックコメントは次の行まで続く'
    Assert-Equal '1|1,2,3' (& $measure 'order.sql' @(
            "SELECT '/*' AS A, '--' AS B FROM DUAL; -- /* old",
            'SELECT 1 FROM DUAL;',
            '/* a */ SELECT 2 FROM DUAL; /* b',
            'c */')) 'SQL の文字列と -- の後ろ'
    Assert-Equal '2|1,4' (& $measure 'index.html' @(
            '<p>Don''t stop</p> <!-- a --> <p>5" screen</p> <!-- old:',
            '<p>old</p>',
            '-->',
            '<p>new</p>')) 'HTML の文の '' と " は文字列の囲みではない'
    Assert-Equal '0|1,2,3' (& $measure 'Open.ps1' @(
            '$open = ''<#''',
            'Write-Host "x" # <# old',
            'Write-Host "y"')) 'PowerShell の文字列と # の後ろの <#'
}

Test-Case 'ExcelTools: 版数は「第1.1版」を読み「第2回」は読まない。日付は和暦（令和・平成・R・H・元年）も読む' {
    Assert-Equal '1.1.0.0||2.0.0.0' ('{0}|{1}|{2}' -f (Get-NameVersion '設計書_第1.1版.xlsx'), (Get-NameVersion '定例会議事録_第2回.docx'), (Get-NameVersion 'Ver2.0'))
    Assert-Equal @('議事録_第1回.docx', '議事録_第2回.docx', '設計書_{v}.xlsx', '設計書_{v}.xlsx') @(@('議事録_第1回.docx', '議事録_第2回.docx', '設計書_第1.0版.xlsx', '設計書_v1.1.xlsx') | ForEach-Object { Get-VersionlessName $_ }) '「第N回」は版数として除かない'
    Assert-Equal '第1.2版|v1.0|議事録_第2回' ('{0}|{1}|{2}' -f (Get-NameVersionText '設計書_第1.2版.xlsx'), (Get-NameVersionText '設計書_v1.0.xlsx'), (Get-NameVersionText '議事録_第2回.docx'))
    $dates = @('令和8年10月15日', '令和元年5月1日', 'R8.10.15', 'H30/4/1', '平成31年4月30日', '㋿8年1月2日', '㍻30年4月1日', '2026/10/15') | ForEach-Object { $d = ConvertTo-DateValue $_; if ($d) { $d.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { "（読めない: $_）" } }
    Assert-Equal @('2026/10/15', '2019/05/01', '2026/10/15', '2018/04/01', '2019/04/30', '2026/01/02', '2018/04/01', '2026/10/15') @($dates) '元号の合字（㋿・㍻）も、Windows PowerShell 5.1 で読む'
    Assert-True ($null -eq (ConvertTo-DateValue 'VER8.1.2')) '英字の後ろの R・H・S は和暦にしない'
}

Test-Case 'ExcelTools: ConvertTo-DateValue は、年の無い日付を AsOf の前後半年の年にし（年明けの 12/28 は前の年）、後ろの曜日（(木)・（木）・(Thu)）も読む' {
    $read = {
        param([string] $Text, [string] $AsOf)
        $d = ConvertTo-DateValue $Text ([datetime]$AsOf)
        if ($d) { $d.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture) } else { "（読めない: $Text）" }
    }
    Assert-Equal @('2026/12/28', '2027/01/07', '2027/01/07', '2027/01/07', '2027/06/30', '2026/07/15') @(@('12/28', '1/7(木)', '1月7日（木）', '1/7 (Thu)', '6/30', '7/15') | ForEach-Object { & $read $_ '2027-01-08' }) '年明け（基準日 2027/01/08）'
    Assert-Equal @('2027/01/05', '2026/10/01') @(@('1/5(火)', '10/1') | ForEach-Object { & $read $_ '2026-12-20' }) '年末（基準日 2026/12/20）'
}

Test-Case 'ExcelTools: Resolve-PhaseName は、「コードレビュー」のように別名に書いた名前はそのまま読み、「詳細設計レビュー」は「レビュー」を除いて読む' {
    $phases = Read-PhaseList (Join-Path $script:ToolsDir 'test-quality/工程.csv')
    $names = @(@('コードレビュー', '詳細設計レビュー', 'DDレビュー', '製造工程', 'UT', '設計', '') | ForEach-Object { $p = Resolve-PhaseName $phases $_; if ($p) { $p.Name } else { '（読めない）' } })
    Assert-Equal @('製造', '詳細設計', '詳細設計', '製造', '単体テスト', '（読めない）', '（読めない）') $names
}

function ConvertTo-WordBlockTexts {
    #     Word の本文（w:body の中の XML）を Get-WordBlocks で読み、「場所|文字」の一覧にする
    param([string] $Body)
    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $Body + '</w:body></w:document>')
    $nsm = New-WordNamespaceManager $doc
    return @(Get-WordBlocks $doc $nsm | ForEach-Object { '{0}|{1}' -f $_.Cell, $_.Text })
}

Test-Case 'ExcelTools: Get-WordBlocks は、変更履歴で移動した文字の移動元（w:moveFrom）とルビ（ふりがな。w:rt）を読まない（移動先とルビを振った文字は読む）' {
    $p = { param($t) '<w:r><w:t>' + $t + '</w:t></w:r>' }
    $body = '<w:p>' + (& $p '担当:') + '<w:r><w:ruby><w:rubyPr><w:hps w:val="10"/></w:rubyPr><w:rt>' + (& $p 'やまだ') + '</w:rt><w:rubyBase>' + (& $p '山田') + '</w:rubyBase></w:ruby></w:r></w:p>' +
        '<w:p><w:moveFrom w:id="1" w:author="佐藤">' + (& $p '移動した文') + '</w:moveFrom>' + (& $p '本文') + '</w:p>' +
        '<w:p><w:moveTo w:id="2" w:author="佐藤">' + (& $p '移動した文') + '</w:moveTo></w:p>' +
        '<w:tbl><w:tr><w:tc><w:p>' + (& $p 'セル1') + '</w:p></w:tc><w:tc><w:p><w:moveFrom w:id="3" w:author="佐藤">' + (& $p '旧') + '</w:moveFrom>' + (& $p 'セル2') + '</w:p></w:tc></w:tr></w:tbl>'
    Assert-Equal @('段落 1|担当:山田', '段落 2|本文', '段落 3|移動した文', '表 1 の 1 行目|セル1 | セル2') (ConvertTo-WordBlockTexts $body)
}

Test-Case 'ExcelTools: Get-WordBlocks は、表の中の内容コントロール（繰り返しセクション）の行も文書の順に数えて読み、セルの中の表の行は表の行として読まない' {
    $row = { param($cells) '<w:tr>' + ((@($cells | ForEach-Object { '<w:tc><w:p><w:r><w:t>' + $_ + '</w:t></w:r></w:p></w:tc>' })) -join '') + '</w:tr>' }
    $item = { param($cells) '<w:sdt><w:sdtPr/><w:sdtContent>' + (& $row $cells) + '</w:sdtContent></w:sdt>' }
    $inner = '<w:tc><w:tbl>' + (& $row @('内側')) + '</w:tbl><w:p/></w:tc>'
    $body = '<w:tbl>' + (& $row @('項目', '桁数')) + '<w:sdt><w:sdtPr/><w:sdtContent>' + (& $item @('受注番号', '10')) + (& $item @('受注日', '8')) + '</w:sdtContent></w:sdt>' +
        '<w:tr><w:tc><w:p><w:r><w:t>備考</w:t></w:r></w:p></w:tc>' + $inner + '</w:tr></w:tbl><w:p><w:r><w:t>後</w:t></w:r></w:p><w:tbl>' + (& $row @('次の表')) + '</w:tbl>'
    Assert-Equal @('表 1 の 1 行目|項目 | 桁数', '表 1 の 2 行目|受注番号 | 10', '表 1 の 3 行目|受注日 | 8', '表 1 の 4 行目|備考 | 内側', '段落 1|後', '表 2 の 1 行目|次の表') (ConvertTo-WordBlockTexts $body)
}

Test-Case 'ExcelTools: Get-EvidenceFilePaths は、フォルダかファイルの名前がほかのファイルのテストID のブックをエビデンスとし、分けたテスト仕様書（数字の名前のフォルダ・受注一覧_2.xlsx）は除かない' {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) 'テスト仕様書'
    $path = { param($rel) [System.IO.Path]::Combine($root, $rel) }
    $case = { param($id, $rel) [pscustomobject]@{ Values = @{ 'テストID' = $id }; Place = [pscustomobject]@{ Path = (& $path $rel); File = [System.IO.Path]::GetFileName($rel) } } }
    $cases = @(
        (& $case 'UT-001' 'UT_受注一覧.xlsx'), (& $case 'UT-002' 'UT_受注一覧.xlsx'), (& $case 'UT-003' 'UT_受注一覧.xlsx'),
        (& $case '1' 'UT-001/エビデンス.xlsx'), (& $case '1' 'UT-002_2.xlsx'), (& $case 'UT-003' 'UT-003.xlsx'), (& $case 'UT-009' 'UT-009.xlsx'),
        (& $case '1' '1/UT.xlsx'), (& $case '2' '2/UT.xlsx'), (& $case '3' '受注一覧_2.xlsx'))
    $set = Get-EvidenceFilePaths $cases
    Assert-Equal @('UT-001/エビデンス.xlsx', 'UT-002_2.xlsx', 'UT-003.xlsx') @(@('UT_受注一覧.xlsx', 'UT-001/エビデンス.xlsx', 'UT-002_2.xlsx', 'UT-003.xlsx', 'UT-009.xlsx', '1/UT.xlsx', '2/UT.xlsx', '受注一覧_2.xlsx') | Where-Object { $set.Contains((& $path $_)) }) '自分のテストID だけの名前のブック（UT-009.xlsx）はテスト仕様書'
    # 表の無いブックも、置き場所の名前で確かめる（Test-Traceability.ps1 はシートのすべてのセルを読むため）
    $set = Get-EvidenceFilePaths $cases ([string[]]@((& $path 'UT-002/画面.xlsx'), (& $path 'メモ.xlsx')))
    Assert-Equal @('UT-002/画面.xlsx') @(@('UT-002/画面.xlsx', 'メモ.xlsx') | Where-Object { $set.Contains((& $path $_)) })
}

Test-Case 'ExcelTools: 障害として数えない状態（取下げ・誤指摘・仕様どおりなど）は、障害管理表を読むツールで同じにする（1 つの品質評価の中でバグの数が食い違わない）' {
    $pattern = Get-ExcludedBugPattern
    Assert-Equal @($true, $true, $true, $true, $true, $true, $true) @(@('取下げ', '誤指摘', '仕様どおり', '仕様通り', '重複', 'キャンセル', 'Cancel') | ForEach-Object { $_ -match $pattern })
    Assert-Equal @($false, $false, $false) @(@('完了', '対応中', '') | ForEach-Object { $_ -match $pattern })
    $dir = New-TestDir
    $bugs = Join-Path $dir 'bugs.csv'
    @('障害番号,機能,障害内容,起票日,状態,混入工程,摘出工程,横展開',
        'B001,SCR001,一覧の件数が違う,2026/10/20,完了,詳細設計,単体テスト,要',
        'B002,SCR002,登録できない,2026/10/20,対応中,詳細設計,単体テスト,要',
        'B003,SCR002,操作の勘違い,2026/10/20,誤指摘,詳細設計,単体テスト,要',
        'B004,SCR003,表示の順番,2026/10/21,仕様どおり,詳細設計,単体テスト,要',
        'B005,SCR003,同じ障害の起票,2026/10/21,取下げ,詳細設計,単体テスト,要') | Set-Content -LiteralPath $bugs -Encoding UTF8
    $spec = Join-Path $dir 'UT.csv'
    @('No,機能,確認内容,結果', '1,SCR001,一覧の件数,OK') | Set-Content -LiteralPath $spec -Encoding UTF8
    $findings = Join-Path $dir 'findings.csv'
    @('設計書,指摘内容,分類,状態', 'DD02_画面設計書_SCR001_受注一覧,桁数が違う,誤り,完了') | Set-Content -LiteralPath $findings -Encoding UTF8
    $tool = { param($name) Join-Path $script:ToolsDir "$name.ps1" }
    Assert-Match (Get-HostText { & (& $tool 'Measure-TestQuality') -Path $spec -BugList $bugs -AsOf '2026-10-23' 3>$null | Out-Null }) 'バグ 2 件（未完了 1 件。取下げなどで数えなかったもの 3 件）'
    Assert-Match (Get-HostText { & (& $tool 'Measure-DefectRemoval') -BugList $bugs 3>$null | Out-Null }) '欠陥 2 件（バグ 2・レビューの指摘 0。取下げなどで数えなかったもの 3 件）'
    Assert-Match (Get-HostText { & (& $tool 'Find-SimilarDefects') -BugList $bugs -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') 3>$null | Out-Null }) '障害 5 件のうち、横展開を探した障害 2 件（横展開が不要・済みのもの、障害ではないもの 3 件を除く）'
    Assert-Match (Get-HostText { & (& $tool 'Measure-ReviewEffectiveness') -Findings $findings -BugList $bugs 3>$null | Out-Null }) '設計で入ったバグ 2 件'
    Assert-Match (Get-HostText { & (& $tool 'New-WeeklyReport') -BugList $bugs -AsOf '2026-10-23' 3>$null | Out-Null }) '・今週: 起票 2 件'
}

Test-Case 'ExcelTools: Select-DesignFiles は、同じ文書の古い版・ツールの出力・書き換える前のファイルのフォルダを読まず、種類ごとに 1 回知らせる' {
    $dir = New-TestDir
    $sheet = [ordered]@{ 'S' = @(, @('画面ID', 'SCR001')) }
    foreach ($f in @('A_v1.0.xlsx', 'A_v1.2.xlsx', 'A_v1.10.xlsx', 'B.xlsx', 'v1.0/C.xlsx', 'v1.1/C.xlsx', 'sub/D_v1.0.xlsx', '置換前/A_v1.10.xlsx', '置換前/深い/B.xlsx')) {
        $p = Join-Path $dir $f
        New-Item -ItemType Directory -Force -Path (Split-Path $p -Parent) | Out-Null
        New-SheetsBook $p $sheet
    }
    Add-BackupMarker (Join-Path $dir '置換前') 'Replace-Excel.ps1'
    # ツールが書き出したブックには印がある（テストで作ったブックには無い）
    $null = & (Join-Path $script:ToolsDir 'Search-Excel.ps1') 'SCR001' (Join-Path $dir 'B.xlsx') -OutXlsx (Join-Path $dir '検索結果.xlsx') 6>$null
    Assert-Equal 'Search-Excel.ps1 の出力' (Get-ToolOutputKind (Join-Path $dir '検索結果.xlsx'))
    Assert-Equal $null (Get-ToolOutputKind (Join-Path $dir 'B.xlsx'))

    $sel = Select-DesignFiles -Path $dir -Exclude (Join-Path $dir 'sub')
    Assert-Equal @('A_v1.10.xlsx', 'B.xlsx', 'v1.1/C.xlsx') @($sel.Files | ForEach-Object { ConvertTo-Slash $_.Relative } | Sort-Object) '版数は数として比べる（1.10 は 1.2 より新しい）。フォルダ名の版数も見る'
    Assert-Equal @($sel.Files | ForEach-Object { $_.Path }) $sel.Paths
    $text = Get-HostText { Write-DesignSkipWarning $sel }
    Assert-Match $text '古い版 3 件は読みませんでした（'
    Assert-Match $text 'A_v1\.0\.xlsx → A_v1\.10\.xlsx'
    Assert-Match $text 'A_v1\.2\.xlsx → A_v1\.10\.xlsx'
    Assert-Match $text 'v1\.0[\\/]C\.xlsx → v1\.1[\\/]C\.xlsx'
    Assert-Match $text 'ツールが書き出したファイル 1 件は、設計書ではないので読みませんでした（検索結果\.xlsx（Search-Excel\.ps1 の出力）。.*［ユーザー設定］で「ExcelTools」を削除してください）'
    Assert-Match $text '書き換える前のファイルを残したフォルダの 2 件は読みませんでした（置換前。'
    Assert-NotMatch $text 'D_v1\.0' '-Exclude のフォルダは知らせない'
    Assert-Equal '' (Get-HostText { Write-DesignSkipWarning $sel }) '知らせるのは 1 回だけ'

    # 書き換える前のファイルのフォルダも、そのフォルダを指定すれば読む
    $sel = Select-DesignFiles -Path (Join-Path $dir '置換前')
    Assert-Equal @('A_v1.10.xlsx', '深い/B.xlsx') @($sel.Files | ForEach-Object { ConvertTo-Slash $_.Relative } | Sort-Object)
}

Test-Case 'ExcelTools: Select-DesignHits は、ID の欄が差し込み欄のままのシート（雛形）を除き、ほかのシートに ID の無いブックはブックごと除く' {
    $dir = New-TestDir
    New-SheetsBook (Join-Path $dir 'テーブル定義書.xlsx') ([ordered]@{
            'T_ORDER' = @(@('テーブルID', 'T_ORDER'), @('物理名', 'データ型'), @('ORDER_NO', 'CHAR'))
            '雛形'    = @(@('テーブルID', '{{テーブルID}}'), @('物理名', 'データ型'), @('{{主キー}}', 'CHAR'))
            # ID の欄でない差し込み欄（メールの文面など）のシートは除かない
            'メール'  = @(@('件名', '{{顧客名}} 様'), @('本文', 'ご注文ありがとうございます'))
        })
    New-Item -ItemType Directory -Path (Join-Path $dir 'sub') | Out-Null
    New-SheetsBook (Join-Path $dir 'sub/画面設計書_雛形.xlsx') ([ordered]@{
            '表紙'         = @(, @('画面設計書（{{画面ID}}）'))
            '画面項目定義' = @(@('画面ID', '{{画面ID}}', '画面名', '{{画面名}}'), @('項目名', '項目ID'))
            'イベント定義' = @(, @('イベントID', 'イベント名'))
        })
    $sel = Select-DesignFiles -Path $dir
    $hits = Select-DesignHits @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern '\S' -Path $sel.Paths -Target Value 6>$null) $sel -NoWarning
    Assert-Equal @('テーブル定義書.xlsx|T_ORDER', 'テーブル定義書.xlsx|メール') @($hits | ForEach-Object { '{0}|{1}' -f $_.File, $_.Sheet } | Select-Object -Unique)
    $text = Get-HostText { Write-DesignSkipWarning $sel }
    Assert-Match $text '雛形（差し込み欄のまま）のシート・ブック 2 件は読みませんでした（'
    Assert-Match $text 'テーブル定義書\.xlsx 雛形（テーブルID が \{\{テーブルID\}\}）'
    Assert-Match $text 'sub[\\/]画面設計書_雛形\.xlsx（ブックのすべてのシート。画面ID が \{\{画面ID\}\}）' 'ファイルの名前は指定したフォルダからの相対パス'
}
