#     Compare-Excel.ps1 のテスト。フィクスチャは tests\fixtures\compare（make_fixtures.py で生成）

$CompareExcel = Join-Path $script:ToolsDir 'Compare-Excel.ps1'
$cmpDir = Join-Path $script:Fixtures 'compare'

function Invoke-Compare {
    param([hashtable] $Options)
    return , @(& $CompareExcel @Options 6>$null)
}

function Get-DiffKeys {
    param($Diffs)
    return @($Diffs | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.File, $_.Location, $_.Kind, $_.Change } | Sort-Object)
}

function New-TableBook {
    #     ExcelTools.psm1 の Export-Workbook で、1 シートの表のブックを作る（テスト用）
    param([string] $Path, [string[]] $Header, $Rows)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $Path -Sheets @(@{
        Name = 'Data'
        Columns = @($Header | ForEach-Object { @{ Header = $_; Width = 10 } })
        Rows = @($Rows | ForEach-Object { @{ Values = $_ } })
    })
}

Test-Case 'Compare-Excel: 値・数式・行・コメント・シートの差分を出す' {
    $diffs = Invoke-Compare @{ Old = (Join-Path $cmpDir 'v1/項目定義.xlsx'); New = (Join-Path $cmpDir 'v2/項目定義.xlsx') }
    $expected = @(
        '項目定義.xlsx|項目定義!C3|値|変更',
        '項目定義.xlsx|項目定義!C6|数式|変更',
        '項目定義.xlsx|項目定義!A4|行|追加',
        '項目定義.xlsx|項目定義!A5|行|削除',
        '項目定義.xlsx|項目定義!B2|コメント|追加',
        '項目定義.xlsx|変更履歴|シート|追加',
        '項目定義.xlsx|旧メモ|シート|削除'
    ) | Sort-Object
    Assert-Equal $expected (Get-DiffKeys $diffs)
    $c3 = @($diffs | Where-Object Location -eq '項目定義!C3')[0]
    Assert-Equal '8' $c3.Old
    Assert-Equal '10' $c3.New
    Assert-Match (@($diffs | Where-Object Kind -eq '数式')[0].Detail) '=SUM\(C2:C4\) → =SUM\(C2:C5\)'
    Assert-Match (@($diffs | Where-Object { $_.Kind -eq '行' -and $_.Change -eq '追加' })[0].Detail) '顧客区分'
}

Test-Case 'Compare-Excel: 挿入した行より下の行は「変更」にならない' {
    # 顧客名の行は 4 行目 → 5 行目に移動しているが、内容は同じ
    $diffs = Invoke-Compare @{ Old = (Join-Path $cmpDir 'v1/項目定義.xlsx'); New = (Join-Path $cmpDir 'v2/項目定義.xlsx') }
    Assert-Equal 0 @($diffs | Where-Object { $_.Old -eq '顧客名' -or $_.New -eq '顧客名' }).Count
}

Test-Case 'Compare-Excel: フォルダどうしで、同じ・差分あり・追加・削除を判定する' {
    $out = Join-Path (New-TestDir) 'diff.xlsx'
    $diffs = Invoke-Compare @{ Old = (Join-Path $cmpDir 'v1'); New = (Join-Path $cmpDir 'v2'); OutXlsx = $out }
    Assert-Equal 0 @($diffs | Where-Object File -eq '同じ.xlsx').Count '同じファイル'
    Assert-Equal 'ファイル|追加' (@($diffs | Where-Object File -eq '追加された.xlsx' | ForEach-Object { '{0}|{1}' -f $_.Kind, $_.Change }) -join ',')
    Assert-Equal 'ファイル|削除' (@($diffs | Where-Object File -eq '削除された.xlsx' | ForEach-Object { '{0}|{1}' -f $_.Kind, $_.Change }) -join ',')

    $files = Read-XlsxSheet -Path $out -Index 2
    $state = @{}
    foreach ($r in 2..$files.Rows.Count) { $state[$files.Rows[$r]['A']] = $files.Rows[$r]['B'] }
    Assert-Equal '同じ' $state['同じ.xlsx']
    Assert-Equal '差分あり' $state['項目定義.xlsx']
    Assert-Equal '追加' $state['追加された.xlsx']
    Assert-Equal '削除' $state['削除された.xlsx']
}

Test-Case 'Compare-Excel: -OutXlsx のリンクは、削除は旧のファイル、それ以外は新のファイルを開く' {
    $out = Join-Path (New-TestDir) 'diff.xlsx'
    $null = Invoke-Compare @{ Old = (Join-Path $cmpDir 'v1/項目定義.xlsx'); New = (Join-Path $cmpDir 'v2/項目定義.xlsx'); OutXlsx = $out }
    $sheet = Read-XlsxSheet -Path $out -Index 1
    $byRef = @{}
    foreach ($r in 2..$sheet.Rows.Count) { $byRef[$sheet.Rows[$r]['C']] = $r }
    $linkOf = { param($loc) @($sheet.Links | Where-Object { $_.Ref -eq "C$($byRef[$loc])" })[0] }
    $changed = & $linkOf '項目定義!C3'
    Assert-Equal "'項目定義'!C3" $changed.Location
    Assert-Match $changed.Target 'v2'
    $deleted = & $linkOf '項目定義!A5'
    Assert-Equal "'項目定義'!A5" $deleted.Location
    Assert-Match $deleted.Target 'v1'
}

Test-Case 'Compare-Excel: 大きな表の先頭に 1 行挿入しても、差分は挿入と本当の変更だけ' {
    $dir = New-TestDir
    $oldRows = @(1..300 | ForEach-Object { , @("ID$_", "名前$_", $_) })
    $newRows = New-Object System.Collections.Generic.List[object]
    $newRows.Add(@('ID0', '先頭に追加', 0))
    foreach ($r in $oldRows) { $newRows.Add($r) }
    $newRows[251] = @('ID251', '名前を変更', 251)
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('ID', '名前', '数') -Rows $oldRows
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('ID', '名前', '数') -Rows $newRows

    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('new.xlsx|Data!A2|行|追加', 'new.xlsx|Data!B253|値|変更') (Get-DiffKeys $diffs)
    Assert-Equal 'Data!B252' (@($diffs | Where-Object Kind -eq '値')[0].OldLocation) '旧の場所'

    # 行を対応付けなければ、挿入した行より下がすべて変更になる
    $naive = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx'); NoRowAlign = $true }
    Assert-True ($naive.Count -gt 800) "行を対応付けない場合の差分: $($naive.Count)"
}

Test-Case 'Compare-Excel: 厳密な LCS の上限を超える大きさでも、挿入した行を正しく対応付ける' {
    $dir = New-TestDir
    $oldRows = @(1..1200 | ForEach-Object { , @("ID$_", "名前$_") })
    $newRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $oldRows) { $newRows.Add($r) }
    $newRows.Insert(0, @('ID0', '先頭に追加'))
    $newRows.RemoveAt(1001)                       # 旧の 1001 行目（ID1001）を削除
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('ID', '名前') -Rows $oldRows
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('ID', '名前') -Rows $newRows
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('new.xlsx|Data!A1002|行|削除', 'new.xlsx|Data!A2|行|追加') (Get-DiffKeys $diffs)
}

Test-Case 'Compare-Excel: ファイルとフォルダを混ぜて指定するとエラー' {
    Assert-Throws { & $CompareExcel -Old (Join-Path $cmpDir 'v1') -New (Join-Path $cmpDir 'v2/同じ.xlsx') 6>$null } '両方とも'
}

Test-Case 'Compare-Excel: Get-Help でヘルプが表示される' {
    $h = Get-Help $CompareExcel
    Assert-Match $h.Synopsis '差分'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
