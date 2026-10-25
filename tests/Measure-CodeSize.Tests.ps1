#     Measure-CodeSize.ps1 のテスト。ソースコードはテストの中で作る

$CodeSizeTool = Join-Path $script:ToolsDir 'Measure-CodeSize.ps1'

function Write-Source {
    param([string] $Root, [string] $Relative, [string[]] $Lines)
    $path = Join-Path $Root $Relative
    New-Item -ItemType Directory -Force -Path (Split-Path $path -Parent) | Out-Null
    [System.IO.File]::WriteAllText($path, (($Lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($false)))
}

function New-SourceTree {
    param([string] $Root, [switch] $Old)
    if ($Old) {
        Write-Source $Root 'order/OrderService.java' @('package order;', '', 'public class OrderService {', '    public int count() { return 0; }', '}')
        Write-Source $Root 'old/Legacy.java' @('class Legacy {}')
    }
    else {
        Write-Source $Root 'order/OrderService.java' @(
            'package order;', '', '/*', ' * 受注のサービス', ' */', 'public class OrderService {', '    // 件数を数える',
            '    public int count() { return 1; } // 末尾のコメント', '    /* 1 行のブロック */', '    public int total() {',
            '        return 2; /* 途中から', '                     次の行まで */', '    }', '}')
        Write-Source $Root 'customer/Customer.cs' @('// 顧客', 'class Customer', '{', '}')
        Write-Source $Root 'node_modules/lib/x.js' @('var x = 1;')
        Write-Source $Root 'readme.txt' @('説明')
    }
    Write-Source $Root 'order/order.sql' @('-- 受注の検索', 'SELECT *', '  FROM T_ORDER /* 受注 */', ' WHERE ORDER_NO = :no;')
    Write-Source $Root 'main.py' @('# main', 'print(1)')
}

Test-Case 'Measure-CodeSize: 言語ごとのコメントの書き方で、行数・コメント・空行・ステップ数をフォルダごとに数える' {
    $root = Join-Path (New-TestDir) 'src'
    New-SourceTree $root
    $csv = Join-Path (New-TestDir) 'size.csv'
    $r = @(& $CodeSizeTool $root -OutCsv $csv 6>$null)
    $byGroup = @{}
    foreach ($g in $r) { $byGroup[$g.Group] = '{0}|{1}|{2}|{3}|{4}' -f $g.Files, $g.Lines, $g.Comment, $g.Blank, $g.Steps }
    Assert-Equal '2|18|7|1|10' $byGroup['order'] 'ブロックコメント・行の途中から始まるコメント・行末のコメント'
    Assert-Equal '1|4|1|0|3' $byGroup['customer']
    Assert-Equal '1|2|1|0|1' $byGroup['（直下）']
    Assert-Equal 3 $r.Count 'node_modules と言語でないファイルは数えない'
    $rows = @(Import-Csv -LiteralPath $csv -Encoding UTF8)
    Assert-Equal @('機能', 'KLOC', 'ステップ数', 'ファイル数', '行数', 'コメント', '空行') @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    Assert-Equal '0.01' @($rows | Where-Object { $_.'機能' -eq 'order' })[0].KLOC

    $text = Get-HostText { & $CodeSizeTool $root | Out-Null }
    Assert-Match $text '4 ファイル: ステップ数 14（0\.014 KLOC）・コメント 9・空行 1'
}

Test-Case 'Measure-CodeSize: -Baseline の前の版と比べて、新規・修正・削除のステップ数を出す' {
    $dir = New-TestDir
    $new = Join-Path $dir 'v2'; $old = Join-Path $dir 'v1'
    New-SourceTree $new
    New-SourceTree $old -Old
    $out = Join-Path $dir 'size.xlsx'
    $r = @(& $CodeSizeTool $new -Baseline $old -OutXlsx $out 6>$null)
    $byGroup = @{}
    foreach ($g in $r) { $byGroup[$g.Group] = '{0}|{1}|{2}' -f $g.Added, $g.Changed, $g.Deleted }
    Assert-Equal '3|1|0' $byGroup['order'] 'count() の書き換えは修正、total() は新規'
    Assert-Equal '3|0|0' $byGroup['customer'] '新しいファイル'
    Assert-Equal '0|0|1' $byGroup['old'] '前の版にだけあるファイル'
    $files = Read-XlsxSheet -Path $out -Index 2
    $status = @{}
    foreach ($row in $files.Rows.Values) { if ($row['A'] -and $row['A'] -ne 'ファイル') { $status[$row['A']] = $row['D'] } }
    Assert-Equal '修正|変更なし|新規|削除' ('{0}|{1}|{2}|{3}' -f $status['order/OrderService.java'], $status['order/order.sql'], $status['customer/Customer.cs'], $status['old/Legacy.java'])
    $text = Get-HostText { & $CodeSizeTool $new -Baseline $old | Out-Null }
    Assert-Match $text '前の版との差: 新規 6・修正 1・削除 1 ステップ'
}

Test-Case 'Measure-CodeSize: 大文字・小文字だけを変えた行も修正に数える' {
    $dir = New-TestDir
    $new = Join-Path $dir 'v2'; $old = Join-Path $dir 'v1'
    Write-Source $old 'app/A.java' @('class A {', '    int userId = getUserId();', '}')
    Write-Source $new 'app/A.java' @('class A {', '    int userID = getUserID();', '}')
    $r = @(& $CodeSizeTool $new -Baseline $old 6>$null)
    Assert-Equal '0|1|0' ('{0}|{1}|{2}' -f $r[0].Added, $r[0].Changed, $r[0].Deleted)
}

Test-Case 'Measure-CodeSize: 空のファイル（Python の __init__.py など）は 0 行に数える' {
    $root = Join-Path (New-TestDir) 'src'
    Write-Source $root 'pkg/main.py' @('print(1)')
    [System.IO.File]::WriteAllText((Join-Path $root 'pkg/__init__.py'), '')
    $r = @(& $CodeSizeTool $root 6>$null)
    Assert-Equal 'pkg|2|1|0|1' ('{0}|{1}|{2}|{3}|{4}' -f $r[0].Group, $r[0].Files, $r[0].Lines, $r[0].Blank, $r[0].Steps)
}

Test-Case 'Measure-CodeSize: 固定形式の COBOL の一連番号（1〜6 桁目）と識別領域（73 桁目から）は、空行の判定と前の版との比較に使わない' {
    $dir = New-TestDir
    $new = Join-Path $dir 'v2'; $old = Join-Path $dir 'v1'
    Write-Source $old 'batch/PGM01.cbl' @('000100 IDENTIFICATION DIVISION.', '000200 PROGRAM-ID. PGM01.', '000300', '000400*COMMENT', '000500 PROCEDURE DIVISION.', '000600     DISPLAY "A".', ('000700     STOP RUN.'.PadRight(72) + 'PGM01V10'))
    # 1 行足して、一連番号を振り直し、識別領域を書き換えた版
    Write-Source $new 'batch/PGM01.cbl' @('000100 IDENTIFICATION DIVISION.', '000200 PROGRAM-ID. PGM01.', '000300', '000400*COMMENT', '000500 PROCEDURE DIVISION.', '000600     DISPLAY "NEW".', '000700     DISPLAY "A".', ('000800     STOP RUN.'.PadRight(72) + 'PGM01V11'))
    $r = @(& $CodeSizeTool $old 6>$null)
    Assert-Equal '7|1|1|5' ('{0}|{1}|{2}|{3}' -f $r[0].Lines, $r[0].Blank, $r[0].Comment, $r[0].Steps) '一連番号だけの行は空行'
    $r = @(& $CodeSizeTool $new -Baseline $old 6>$null)
    Assert-Equal '1|0|0' ('{0}|{1}|{2}' -f $r[0].Added, $r[0].Changed, $r[0].Deleted) '一連番号の振り直しと識別領域は修正にしない'
}

Test-Case 'Measure-CodeSize: 既定で数えないフォルダ（out・vendor・node_modules など）は Test-CodeTrace と同じにし、除いたフォルダを表示する' {
    $root = Join-Path (New-TestDir) 'src'
    Write-Source $root 'scr001/OrderList.java' @('class OrderList {', '    String id = "E0001";', '}')
    Write-Source $root 'out/orderList.js' @('var a = 1;', 'var b = 2;')
    Write-Source $root 'vendor/acme/Lib.php' @('<?php', 'echo 1;')
    Write-Source $root 'web/node_modules/lib/x.js' @('var x = 1;')
    $text = Get-HostText { $script:size = @(& $CodeSizeTool $root) }
    Assert-Equal @('scr001|1|3') @($script:size | ForEach-Object { '{0}|{1}|{2}' -f $_.Group, $_.Files, $_.Steps }) 'ビルドの出力（out）・外部のライブラリ（vendor）は KLOC に入れない'
    Assert-Match $text '数えなかったフォルダ（-Exclude）: out（1 ファイル）・vendor（1 ファイル）・web/node_modules（1 ファイル）'
    $trace = Join-Path $script:ToolsDir 'Test-CodeTrace.ps1'
    $text = Get-HostText { & $trace -DesignPath (Join-Path $script:RepoRoot 'docs/design-sample') -SourcePath $root | Out-Null }
    Assert-Match $text 'ソース 1 ファイルで' 'Measure-CodeSize と同じファイルを読む'
    Assert-Match $text '見なかったフォルダ（-Exclude）: out（1 ファイル）・vendor（1 ファイル）・web/node_modules（1 ファイル）'
    # -Exclude を空にすると、すべてのフォルダを数える
    $text = Get-HostText { & $CodeSizeTool $root -Exclude @() | Out-Null }
    Assert-Match $text '^\s*4 ファイル: ステップ数 8'
    Assert-NotMatch $text '数えなかったフォルダ'
}

Test-Case 'Measure-CodeSize: Get-Help でヘルプが表示される' {
    $h = Get-Help $CodeSizeTool
    Assert-Match $h.Synopsis 'ステップ数'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
