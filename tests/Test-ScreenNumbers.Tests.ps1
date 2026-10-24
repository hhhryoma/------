#     Test-ScreenNumbers.ps1 のテスト。画面設計書はテストの中で作る（New-SheetsBook と Add-TestDrawing は Invoke-Tests.ps1）

$ScreenNumbersTool = Join-Path $script:ToolsDir 'Test-ScreenNumbers.ps1'

function Invoke-ScreenNumbers {
    param([hashtable] $Options)
    return , @(& $ScreenNumbersTool @Options 6>$null 3>$null)
}

function Get-ScreenNumberKeys {
    param($Issues)
    return @($Issues | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Level, $_.Item, $_.File, $_.Location })
}

function New-ScreenBook {
    #     画面レイアウト（スクリーンショットと重ねた番号）と画面項目定義のシートを持つ画面設計書（テスト用）。Items: @(@(No, 項目名), ...)
    param([string] $Path, $Items, $Drawing, [string] $Id = 'SCR002', [string] $Name = '受注登録')
    $header = @('画面ID', $Id, '画面名', $Name)
    $table = @(, @('No', '項目名', '桁数')) + @($Items | ForEach-Object { , @([string]$_[0], $_[1], '10') })
    New-SheetsBook $Path ([ordered]@{ '画面レイアウト' = @(, $header); '画面項目定義' = @(, $header) + @(, @()) + $table })
    Add-TestDrawing $Path $Drawing
}

function New-Marker {
    #     重ねた番号の図形（列・行は 0 から。1 マスの大きさ）
    param([string] $Name, [string] $Text, [int] $Col, [int] $Row, [switch] $Hidden)
    return @{ Kind = 'sp'; Name = $Name; Text = $Text; From = @($Col, $Row); To = @(($Col + 1), ($Row + 1)); Hidden = [bool]$Hidden }
}

Test-Case 'Test-ScreenNumbers: 画面項目定義に無い番号と、同じ番号の重複を見つける（丸数字・No.3・(5) など。吹き出し・画像の外・非表示は読まない）' {
    $dir = New-TestDir
    $book = Join-Path $dir '画面設計書_SCR002.xlsx'
    $items = @(@(1, '受注日'), @(2, '顧客コード'), @(3, '顧客名'), @(4, '支払方法'), @(5, '希望納品日'), @(6, '配送時間帯'), @(7, '備考'), @(8, '登録'), @(21, '合計金額'))
    New-ScreenBook $book $items @(
        @{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 2); To = @(12, 30) },
        (New-Marker '楕円 1' '①' 2 4), (New-Marker '楕円 2' '①' 6 4),
        (New-Marker 'テキスト 3' 'No.3' 2 7), (New-Marker '楕円 4' '（５）' 2 9), (New-Marker '楕円 5' '㉑' 2 11),
        (New-Marker '楕円 6' '❾' 2 13), (New-Marker '楕円 7' '⑦⑧' 2 15),
        @{ Kind = 'sp'; Name = '吹き出し 8'; Text = '⑤ 銀行振込のときだけ表示'; From = @(4, 15); To = @(8, 17) },
        (New-Marker '楕円 9' '⑳' 13 4), (New-Marker '楕円 10' '⑲' 2 17 -Hidden),
        @{ Kind = 'grpSp'; Name = 'グループ化 11'; From = @(8, 20); To = @(11, 26); Children = @(
                @{ Kind = 'sp'; Name = '楕円 12'; Text = '㉒'; Off = @(0, 0); Ext = @(500, 500) },
                @{ Kind = 'sp'; Name = '矢印 13'; Off = @(500, 500); Ext = @(500, 500) }) })
    $text = Get-HostText { & $ScreenNumbersTool -Path $dir }
    Assert-Match $text '番号を重ねたスクリーンショットのあるシート 1・重ねた番号 8 個' '吹き出し・画像の外・非表示の図形は数えない'
    $issues = Invoke-ScreenNumbers @{ Path = $dir }
    $file = '画面設計書_SCR002.xlsx'
    Assert-Equal @("警告|番号の重複|$file|画面レイアウト!G5", "エラー|定義に無い番号|$file|画面レイアウト!C14", "エラー|定義に無い番号|$file|画面レイアウト!I22") (Get-ScreenNumberKeys $issues)
    Assert-Equal 'No 1 の番号が 2 か所にあります（楕円 1: C5、楕円 2: G5）' $issues[0].Detail
    Assert-Equal '重ねた番号 ❾（楕円 6）は、画面項目定義に無い番号です（画面項目定義の No: 1〜8・21）' $issues[1].Detail
    Assert-Match $issues[2].Detail '^重ねた番号 ㉒（楕円 12）は' 'グループの中の番号も読む'
}

Test-Case 'Test-ScreenNumbers: 画面項目定義が別のブックでも、シート名と画面名で対応付ける。見つからない画面は「定義なし」' {
    $dir = New-TestDir
    $layout = Join-Path $dir 'レイアウト.xlsx'
    New-SheetsBook $layout ([ordered]@{ '受注登録' = @(, @('メモ')); '受注一覧' = @(, @('メモ')) })
    $screen = @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 2); To = @(12, 30) }, (New-Marker '楕円 2' '②' 2 4), (New-Marker '楕円 3' '⑤' 2 6))
    Add-TestDrawing $layout $screen
    Add-TestDrawing $layout $screen -Sheet 2
    $header = { param($id, $name) @('画面ID', $id, '画面名', $name) }
    $table = @(@(), @('No', '項目名', '桁数'), @('1', '受注日', '8'), @('2', '顧客コード', '8'))
    New-SheetsBook (Join-Path $dir '画面項目定義.xlsx') ([ordered]@{
            'SCR002' = @(, (& $header 'SCR002' '受注登録')) + $table
            'SCR003' = @(, (& $header 'SCR003' '受注詳細')) + $table + @(, @('5', '受注金額', '9')) })
    $issues = Invoke-ScreenNumbers @{ Path = $dir }
    Assert-Equal @('エラー|定義に無い番号|レイアウト.xlsx|受注登録!C7', '情報|定義なし|レイアウト.xlsx|受注一覧!C5') (Get-ScreenNumberKeys $issues) '受注登録は SCR002 の表（No 5 が無い）。受注一覧の表は無い'
    Assert-Match $issues[1].Detail '画面項目定義（No・項目名・桁数の表）が見つからない.*重ねた番号 2 個'
}

function New-BaselineBooks {
    #     前の版（v1.0）と今の版（v1.1）。v1.1 は No 3 に「担当者」を足して、4 以降が 1 つずつずれた
    param([string] $Dir)
    $oldItems = @(@(1, '受注日'), @(2, '顧客コード'), @(3, '顧客名'), @(4, '支払方法'), @(5, '希望納品日'))
    $newItems = @(@(1, '受注日'), @(2, '顧客コード'), @(3, '担当者'), @(4, '顧客名'), @(5, '支払方法'), @(6, '希望納品日'))
    $pic = @{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 2); To = @(12, 30) }
    New-Item -ItemType Directory -Path (Join-Path $Dir 'v1.0'), (Join-Path $Dir 'v1.1') | Out-Null
    # ① はそのまま。③ は別の場所に動かした（付け直した）。④ は古いまま。⑤ は ⑥ に直した
    New-ScreenBook (Join-Path $Dir 'v1.0/画面設計書_SCR002_v1.0.xlsx') $oldItems @($pic, (New-Marker '楕円 1' '①' 2 4), (New-Marker '楕円 3' '③' 2 8), (New-Marker '楕円 4' '④' 2 12), (New-Marker '楕円 5' '⑤' 6 12))
    New-ScreenBook (Join-Path $Dir 'v1.1/画面設計書_SCR002_v1.1.xlsx') $newItems @($pic, (New-Marker '楕円 1' '①' 2 4), (New-Marker '楕円 3' '③' 6 8), (New-Marker '楕円 4' '④' 2 12), (New-Marker '楕円 5' '⑥' 6 12))
}

Test-Case 'Test-ScreenNumbers: -Baseline の前の版と比べて、同じ場所の番号が指す項目が変わったもの（No がずれたのに直していない番号）を見つける' {
    $dir = New-TestDir
    New-BaselineBooks $dir
    $issues = Invoke-ScreenNumbers @{ Path = (Join-Path $dir 'v1.1'); Baseline = (Join-Path $dir 'v1.0') }
    Assert-Equal @('警告|指す項目の変化|画面設計書_SCR002_v1.1.xlsx|画面レイアウト!C13') (Get-ScreenNumberKeys $issues) '直した ⑥・動かした ③ は出さない'
    Assert-Equal '重ねた番号 ④（楕円 4）が指す項目が、前の版の「支払方法」から「顧客名」に変わりました。「支払方法」は、いまは No 5 です。番号が古いままでないか確かめてください' $issues[0].Detail
    $none = Invoke-ScreenNumbers @{ Path = (Join-Path $dir 'v1.1') }
    Assert-Equal 0 $none.Count '前の版が無ければ、番号はどれも画面項目定義にある'
}

Test-Case 'Test-ScreenNumbers: -OutXlsx に指摘一覧と、重ねた番号ごとに指す項目を並べた番号の一覧を保存する' {
    $dir = New-TestDir
    New-BaselineBooks $dir
    $out = Join-Path $dir '画面の番号.xlsx'
    $null = Invoke-ScreenNumbers @{ Path = (Join-Path $dir 'v1.1'); Baseline = (Join-Path $dir 'v1.0'); OutXlsx = $out }
    $issues = Read-XlsxSheet $out 1
    Assert-Equal '画面設計書_SCR002_v1.1.xlsx|警告|指す項目の変化|画面レイアウト!C13' ('{0}|{1}|{2}|{3}' -f $issues.Rows[2]['A'], $issues.Rows[2]['B'], $issues.Rows[2]['C'], $issues.Rows[2]['D'])
    Assert-Equal "'画面レイアウト'!C13" $issues.Links[0].Location
    $list = Read-XlsxSheet $out 2
    Assert-Equal @('指摘一覧', '番号の一覧') $list.Sheets
    Assert-Equal 'ファイル|場所|番号|図形|項目名（画面項目定義）|前の版の項目名|結果' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G') | ForEach-Object { $list.Rows[1][$_] }) -join '|')
    $rows = @(2..5 | ForEach-Object { $r = $list.Rows[$_]; '{0}|{1}|{2}|{3}|{4}' -f $r['B'], $r['C'], $r['E'], $r['F'], $r['G'] })
    Assert-Equal @('画面レイアウト!C5|①|受注日|受注日|OK', '画面レイアウト!G9|③|担当者||OK', '画面レイアウト!C13|④|顧客名|支払方法|指す項目の変化', '画面レイアウト!G13|⑥|希望納品日|希望納品日|OK') $rows
}

Test-Case 'Test-ScreenNumbers: スクリーンショットの無い詳細設計サンプルでは、何も知らせない' {
    $sample = Join-Path $script:RepoRoot 'docs/design-sample'
    $text = Get-HostText { & $ScreenNumbersTool -Path $sample }
    Assert-Match $text 'スクリーンショット（画像）の上に重ねた番号が見つかりませんでした'
    Assert-Equal 0 (Invoke-ScreenNumbers @{ Path = $sample }).Count
}

Test-Case 'Test-ScreenNumbers: レイアウトと画面項目定義のブックの両方に画面ID の名前のシート（SCR002）があっても、その画面の表と対応付ける' {
    $dir = New-TestDir
    $layout = Join-Path $dir 'レイアウト.xlsx'
    New-SheetsBook $layout ([ordered]@{ 'SCR002' = @(, @('メモ')); 'SCR003' = @(, @('メモ')) })
    $screen = @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 2); To = @(12, 30) }, (New-Marker '楕円 2' '②' 2 4), (New-Marker '楕円 3' '⑤' 2 6))
    Add-TestDrawing $layout $screen
    Add-TestDrawing $layout $screen -Sheet 2
    $header = { param($id, $name) @('画面ID', $id, '画面名', $name) }
    $table = @(@(), @('No', '項目名', '桁数'), @('1', '受注日', '8'), @('2', '顧客コード', '8'))
    New-SheetsBook (Join-Path $dir '画面項目定義.xlsx') ([ordered]@{
            'SCR002' = @(, (& $header 'SCR002' '受注登録')) + $table
            'SCR003' = @(, (& $header 'SCR003' '受注詳細')) + $table + @(, @('5', '受注金額', '9')) })
    $issues = Invoke-ScreenNumbers @{ Path = $dir }
    Assert-Equal @('エラー|定義に無い番号|レイアウト.xlsx|SCR002!C7') (Get-ScreenNumberKeys $issues) 'SCR002 の表に No 5 は無い。SCR003 の表には ②・⑤ がある（「定義なし」にしない）'
}

Test-Case 'Test-ScreenNumbers: Get-Help でヘルプが表示される' {
    $h = Get-Help $ScreenNumbersTool
    Assert-Match $h.Synopsis '画面項目定義'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
