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

Test-Case 'Measure-CodeSize: Get-Help でヘルプが表示される' {
    $h = Get-Help $CodeSizeTool
    Assert-Match $h.Synopsis 'ステップ数'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
