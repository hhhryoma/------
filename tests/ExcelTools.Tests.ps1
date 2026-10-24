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

Test-Case 'ExcelTools: Get-LeafName は、\ と（拡張子で終わる名前の）/ でフォルダを除き、「外部I/F設計書」の / は残す（PowerShell 5.1 と 7 で同じ）' {
    Assert-Equal @('基本設計書.xlsx', '基本設計書.xlsx', '基本設計書', '外部I/F設計書', '外部I/F設計書 v1.10', '') @(@('D:\納品物\基本設計書.xlsx', '設計書/基本設計書.xlsx', ' sub\基本設計書 ', '外部I/F設計書', '外部I/F設計書 v1.10', '') | ForEach-Object { Get-LeafName $_ })
    Assert-Equal (ConvertTo-NameKey '基本設計書') (ConvertTo-NameKey 'D:\納品物\基本設計書.xlsx') 'フォルダと拡張子を除く'
    Assert-True ((ConvertTo-NameKey '外部I/F設計書') -ne (ConvertTo-NameKey '内部I/F設計書')) '「I/F」の前が違う名前は別のもの'
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
