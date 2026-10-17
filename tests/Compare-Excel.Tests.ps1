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

function Read-ZipEntryText {
    #     xlsx の中の 1 つのパーツを文字列で読む
    param([string] $Path, [string] $Name)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $sr = New-Object System.IO.StreamReader($zip.GetEntry($Name).Open())
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    }
    finally { $fs.Dispose() }
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
    Assert-True ($naive.Count -gt 500) "行を対応付けない場合の差分: $($naive.Count)"
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

function New-SheetsBook {
    #     複数シートのブックを作る（テスト用）。Sheets: @{ 名前 = @(見出しの配列, 行の配列...) } の順序付き辞書
    param([string] $Path, $Sheets)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $list = @()
    foreach ($name in $Sheets.Keys) {
        $header = $Sheets[$name][0]
        $rows = @($Sheets[$name] | Select-Object -Skip 1)
        $list += @{ Name = $name; Columns = @($header | ForEach-Object { @{ Header = $_; Width = 10 } }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) }
    }
    Export-Workbook -FilePath $Path -Sheets $list
}

Test-Case 'Compare-Excel: -KeyColumn でキーの値どうしを対応付け、並べ替えは「移動」だけになる' {
    $dir = New-TestDir
    # 1 つのシートに 2 つの表（カラム定義とインデックス定義）。空行で区切る
    New-SheetsBook (Join-Path $dir 'old.xlsx') ([ordered]@{ 'T_ORDER' = @(
        @('No', '物理名', '桁数'),
        @('1', 'ORDER_NO', '12'), @('2', 'CUSTOMER_CD', '8'), @('3', 'ORDER_DATE', ''), @('4', 'REMARKS', '500'),
        @(),
        @('No', 'インデックス名', 'カラム'),
        @('1', 'IX_01', 'CUSTOMER_CD'), @('2', 'IX_02', 'ORDER_DATE')) })
    New-SheetsBook (Join-Path $dir 'new.xlsx') ([ordered]@{ 'T_ORDER' = @(
        @('No', '物理名', '桁数'),
        @('1', 'ORDER_DATE', ''), @('2', 'ORDER_NO', '12'), @('3', 'CUSTOMER_CD', '10'), @('4', 'STATUS', '2'),
        @(),
        @('No', 'インデックス名', 'カラム'),
        @('1', 'IX_02', 'ORDER_DATE'), @('2', 'IX_01', 'CUSTOMER_CD')) })

    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx'); KeyColumn = @('物理名', 'インデックス名') }
    $keys = Get-DiffKeys $diffs
    Assert-True ($keys -contains 'new.xlsx|T_ORDER!C4|値|変更') 'CUSTOMER_CD の桁数 8 → 10'
    Assert-True ($keys -contains 'new.xlsx|T_ORDER!A5|行|追加') 'STATUS の行を追加'
    Assert-True ($keys -contains 'new.xlsx|T_ORDER!A5|行|削除') 'REMARKS の行を削除（旧 5 行目）'
    Assert-Equal 'T_ORDER!A5' (@($diffs | Where-Object { $_.Kind -eq '行' -and $_.Change -eq '削除' })[0].OldLocation) '削除は旧の場所'
    Assert-Equal 2 @($diffs | Where-Object Change -eq '移動').Count '順番が変わった行（ORDER_DATE と IX_02）'
    # No 列（連番）の振り直しは差分になるが、キー以外の列の値が並べ替えで「変更」にならない
    Assert-Equal 0 @($diffs | Where-Object { $_.Kind -eq '値' -and ($_.Old -match '^(ORDER_|IX_|CUSTOMER)' -or $_.New -match '^(ORDER_|IX_|CUSTOMER)') }).Count 'キーの値は変更にならない'
}

Test-Case 'Compare-Excel: 内容が同じまま場所が変わった行は「移動」として出す' {
    $dir = New-TestDir
    $base = @(1..10 | ForEach-Object { , @("項目$_", "説明$_") })
    $moved = New-Object System.Collections.Generic.List[object]
    foreach ($r in $base) { $moved.Add($r) }
    $block = @($moved[1], $moved[2])
    $moved.RemoveRange(1, 2)
    $moved.InsertRange(6, [object[]]$block)       # 項目2・項目3 を後ろへ移動
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('項目', '説明') -Rows $base
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('項目', '説明') -Rows $moved
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('移動', '移動') @($diffs | ForEach-Object { $_.Change }) ((Get-DiffKeys $diffs) -join ', ')
    Assert-Match ($diffs[0].Detail + $diffs[1].Detail) '項目[23]'
}

Test-Case 'Compare-Excel: 中身がよく似たシートの名前変更を見つけ、中身の差分も出す' {
    $dir = New-TestDir
    $rows = @(1..8 | ForEach-Object { , @("項目$_", "値$_") })
    $changed = @($rows | ForEach-Object { , @($_[0], $_[1]) })
    $changed[3] = @('項目4', '値4（改）')
    New-SheetsBook (Join-Path $dir 'old.xlsx') ([ordered]@{ 'メモ' = @(, @('項目', '値')) + $rows; '別' = @(@('A'), @('x')) })
    New-SheetsBook (Join-Path $dir 'new.xlsx') ([ordered]@{ 'メモ（改）' = @(, @('項目', '値')) + $changed; '別' = @(@('A'), @('x')) })
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('new.xlsx|メモ（改）!B5|値|変更', 'new.xlsx|メモ（改）|シート|名前変更') (Get-DiffKeys $diffs)
    Assert-Match (@($diffs | Where-Object Change -eq '名前変更')[0].Detail) 'メモ → メモ（改）'
    Assert-Equal 'メモ!B5' (@($diffs | Where-Object Kind -eq '値')[0].OldLocation) '旧の場所は旧のシート名'
}

Test-Case 'Compare-Excel: 列を挿入しても、右側の列が「変更」にならない' {
    $dir = New-TestDir
    $old = @(1..12 | ForEach-Object { , @("ID$_", "名前$_", "$_") })
    $new = @(1..12 | ForEach-Object { , @("ID$_", "区分$($_ % 3)", "名前$_", "$_") })
    $new[4] = @('ID5', '区分2', '名前5（改）', '5')
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('ID', '名前', '数') -Rows $old
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('ID', '区分', '名前', '数') -Rows $new
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('new.xlsx|Data!B1|列|追加', 'new.xlsx|Data!C6|値|変更') (Get-DiffKeys $diffs)
    Assert-Equal 'Data!B6' (@($diffs | Where-Object Kind -eq '値')[0].OldLocation) '旧の場所は旧の列'

    $naive = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx'); NoColumnAlign = $true }
    Assert-True ($naive.Count -gt 20) "列を対応付けない場合の差分: $($naive.Count)"
}

Test-Case 'Compare-Excel: 行の挿入で振り直された No 列は「連番の振り直し」1 件にまとめる' {
    $dir = New-TestDir
    $old = @(1..10 | ForEach-Object { , @("$_", "項目$_", "$($_ * 100)") })
    $new = New-Object System.Collections.Generic.List[object]
    foreach ($r in $old) { $new.Add(@($r[0], $r[1], $r[2])) }
    $new.Insert(4, @('5', '追加項目', '999'))     # 5 行目に挿入し、以降の No を 1 ずつ振り直す
    for ($i = 5; $i -lt $new.Count; $i++) { $new[$i][0] = "$($i + 1)" }
    # 金額がたまたま同じだけ増えた行（番号が連続しない）はまとめない
    $new[1][2] = '201'; $new[2][2] = '301'
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('No', '項目', '金額') -Rows $old
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('No', '項目', '金額') -Rows $new
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx') }
    Assert-Equal @('new.xlsx|Data!A6|行|追加', 'new.xlsx|Data!A7|連番|変更', 'new.xlsx|Data!C3|値|変更', 'new.xlsx|Data!C4|値|変更') (Get-DiffKeys $diffs)
    $renum = @($diffs | Where-Object Kind -eq '連番')[0]
    Assert-Equal '連番の振り直し: A7～A12（5～10 → 6～11、6 個）' $renum.Detail
}

Test-Case 'Compare-Excel: 長い文字列の変更は、変わった部分だけを [-削除-]{+追加+} で示す' {
    $dir = New-TestDir
    New-TableBook -Path (Join-Path $dir 'old.xlsx') -Header @('No', '説明') -Rows @(
        @('1', '受注情報を受注テーブルに登録する。登録時に受注番号を採番し、ステータスを「受付」に設定する。'),
        @('2', "SELECT ORDER_NO, CUSTOMER_CD FROM T_ORDER WHERE STATUS = '01'"),
        @('3', '短い'),
        @('4', 'まったく別の文章に書き換えた場合はどうなるかを確認する'))
    New-TableBook -Path (Join-Path $dir 'new.xlsx') -Header @('No', '説明') -Rows @(
        @('1', '受注情報を受注テーブルに更新する。登録時に受注番号を採番し、ステータスを「受付済」に設定する。'),
        @('2', "SELECT ORDER_NO, CUSTOMER_CD, ORDER_DATE FROM T_ORDER WHERE STATUS = '02'"),
        @('3', '長い'),
        @('4', '内容を全面的に見直して、新しい仕様として記述しなおしたもの'))
    $out = Join-Path $dir 'diff.xlsx'
    $diffs = Invoke-Compare @{ Old = (Join-Path $dir 'old.xlsx'); New = (Join-Path $dir 'new.xlsx'); OutXlsx = $out }
    $detail = @{}
    foreach ($d in $diffs) { $detail[$d.Cell] = $d.Detail }
    Assert-Equal '値: …情報を受注テーブルに[-登録-]{+更新+}する。登録時に受注番…、ステータスを「受付{+済+}」に設定する。' $detail['B2']
    # 英数字は語の単位で比べ、空白や区切りは変更の外に出す
    Assert-Match $detail['B3'] '_CD\{\+, ORDER_DATE\+\} FROM'
    Assert-Match $detail['B3'] "'\[-01-\]\{\+02\+\}'"
    Assert-Equal '値: 短い → 長い' $detail['B4'] '短い文字列は全体を示す'
    Assert-Match $detail['B5'] '^値: まったく別の.* → 内容を全面的に' 'ほとんど書き換えた場合は全体を示す'
    Assert-Equal '受注情報を受注テーブルに登録する。登録時に受注番号を採番し、ステータスを「受付」に設定する。' (@($diffs | Where-Object Cell -eq 'B2')[0].Old) 'Old は元の文字列のまま'

    # -OutXlsx では、変更前の削除した部分を取り消し線、変更後の追加した部分を太字にする
    $xml = Read-ZipEntryText $out 'xl/worksheets/sheet1.xml'
    Assert-Match $xml '<c r="G2"[^>]*><is>(?:(?!</c>).)*<rPr><strike/>(?:(?!</c>).)*<t xml:space="preserve">登録</t>'
    Assert-Match $xml '<c r="H2"[^>]*><is>(?:(?!</c>).)*<rPr><b/><u/>(?:(?!</c>).)*<t xml:space="preserve">更新</t>'
    Assert-Equal '受注情報を受注テーブルに更新する。登録時に受注番号を採番し、ステータスを「受付済」に設定する。' (Read-XlsxSheet -Path $out -Index 1).Rows[2]['H'] '文字列としては変更後の全体'
}

foreach ($variant in @('', 'sst')) {
    $label = '文字列をセルに直接持つブック'
    if ($variant) { $label = '共有文字列のブック' }
    $fmtDir = Join-Path $cmpDir 'format'
    if ($variant) { $fmtDir = Join-Path $fmtDir $variant }
    Test-Case "Compare-Excel: -CompareFormat で太字・塗りつぶし・部分書式・取り消し線の変更を、範囲にまとめて出す（$label）" {
        $diffs = Invoke-Compare @{ Old = (Join-Path $fmtDir 'old.xlsx'); New = (Join-Path $fmtDir 'new.xlsx'); CompareFormat = $true }
        $got = @($diffs | ForEach-Object { '{0}|{1}|{2}' -f $_.Location, $_.Kind, $_.Detail })
        Assert-Equal @(
            '画面項目!A2|行|行を追加: A:0 | B:受注ID | C:新規に追加した項目',
            '画面項目!B3|書式|書式: 太字 なし → あり（旧 B2）',
            '画面項目!A4:C4|書式|書式: 塗りつぶし 黄 → なし（旧 A3:C3）',
            '画面項目!C5|書式|書式: 部分書式 「正式名称」赤 → なし（旧 C4）',
            '画面項目!C6|値|値: 個人・法人の区分 → 個人・法人・団体の区分（旧 C5）',
            '画面項目!C8:C9|書式|書式: 取り消し線 なし → あり（旧 C7:C8）') $got
        $fill = @($diffs | Where-Object Location -eq '画面項目!A4:C4')[0]
        Assert-Equal '塗りつぶし 黄' $fill.Old
        Assert-Equal '標準' $fill.New
    }
}

Test-Case 'Compare-Excel: -CompareFormat を付けなければ書式は比べない' {
    $fmtDir = Join-Path $cmpDir 'format'
    $diffs = Invoke-Compare @{ Old = (Join-Path $fmtDir 'old.xlsx'); New = (Join-Path $fmtDir 'new.xlsx') }
    Assert-Equal @('A2|行', 'C6|値') @($diffs | ForEach-Object { '{0}|{1}' -f $_.Cell, $_.Kind })
}

Test-Case 'Compare-Excel: ファイルとフォルダを混ぜて指定するとエラー' {
    Assert-Throws { & $CompareExcel -Old (Join-Path $cmpDir 'v1') -New (Join-Path $cmpDir 'v2/同じ.xlsx') 6>$null } '両方とも'
}

Test-Case 'Compare-Excel: Get-Help でヘルプが表示される' {
    $h = Get-Help $CompareExcel
    Assert-Match $h.Synopsis '差分'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
