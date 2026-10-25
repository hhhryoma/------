#     Test-Screenshots.ps1 のテスト。元のワイヤーフレームと設計書はテストの中で作る（New-SheetsBook・Add-TestDrawing・Edit-ZipText は Invoke-Tests.ps1）

$ScreenshotsTool = Join-Path $script:ToolsDir 'Test-Screenshots.ps1'

function Invoke-Screenshots {
    param([hashtable] $Options)
    return , @(& $ScreenshotsTool @Options 6>$null 3>$null)
}

function Get-ScreenStates {
    param($Results)
    return @($Results | ForEach-Object { '{0}|{1}' -f $_.Screen, $_.Status })
}

function New-ScreenshotSet {
    #     元のワイヤーフレーム（表紙・受注一覧・受注登録。画面のシートは A1:F20 が印刷範囲で、L 列は範囲の外のメモ）と、
    #     設計書（受注一覧・受注登録・受注詳細。受注登録は「画面名」の欄で対応付け、スクリーンショットに番号を 2 つ重ねる）
    param([string] $Dir)
    $wire = Join-Path $Dir 'ワイヤーフレーム'
    $design = Join-Path $Dir '設計書'
    New-Item -ItemType Directory -Path $wire, $design | Out-Null
    $book = Join-Path $wire '受注管理_ワイヤーフレーム.xlsx'
    # 見出しの行の列の数だけ書かれるので、見出しを L 列まで並べる
    $screen = { param($title, $button) @(@($title, '', '', '', '', '', '', '', '', '', '', ''), @('検索条件', '受注日'), @('', $button), @('', '', '', '', '', '', '', '', '', '', '', 'メモ: 文言は仮')) }
    New-SheetsBook $book ([ordered]@{ '表紙' = @(, @('受注管理 ワイヤーフレーム')); '受注一覧' = (& $screen '受注一覧' '検索ボタン'); '受注登録' = (& $screen '受注登録' '登録ボタン') })
    Edit-ZipText $book 'xl/workbook.xml' { param($t) $t.Replace('<definedNames>', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="1">''受注一覧''!$A$1:$F$20</definedName><definedName name="_xlnm.Print_Area" localSheetId="2">''受注登録''!$A$1:$F$20</definedName>') }
    Add-TestDrawing $book @(@{ Kind = 'sp'; Name = '四角形 1'; Text = '一覧'; From = @(1, 5); To = @(5, 12) }) -Sheet 2
    Add-TestDrawing $book @(@{ Kind = 'sp'; Name = '四角形 1'; Text = '明細'; From = @(1, 5); To = @(5, 12) }) -Sheet 3
    New-DesignScreen (Join-Path $design 'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx') '受注一覧' $null 'Screen'
    New-DesignScreen (Join-Path $design 'DD02_画面設計書_SCR002_受注登録_v1.0.xlsx') '画面レイアウト' '受注登録' 'Screen'
    New-DesignScreen (Join-Path $design 'DD02_画面設計書_SCR003_受注詳細_v1.0.xlsx') '画面レイアウト' '受注詳細' 'Screen'
    return [pscustomobject]@{ Wire = $wire; Book = $book; Design = $design }
}

function New-DesignScreen {
    #     スクリーンショット（Png）を貼った設計書。Label があれば「画面名」の欄に書き、番号を 2 つ重ねる。表紙には小さいロゴ
    param([string] $Path, [string] $Sheet, [string] $Label, [string] $Png)
    $header = @('メモ')
    if ($Label) { $header = @('画面名', $Label) }
    New-SheetsBook $Path ([ordered]@{ '表紙' = @(, @('画面設計書')); $Sheet = @(, $header) })
    Add-TestDrawing $Path @(@{ Kind = 'pic'; Name = 'ロゴ'; Png = 'Logo'; From = @(1, 1); To = @(2, 3) })
    $items = @(@{ Kind = 'pic'; Name = '図 1'; Png = $Png; From = @(1, 2); To = @(8, 20) })
    if ($Label) { $items += @(@{ Kind = 'sp'; Name = '楕円 1'; Text = '①'; From = @(2, 4); To = @(3, 5) }, @{ Kind = 'sp'; Name = '楕円 2'; Text = '②'; From = @(2, 8); To = @(3, 9) }) }
    Add-TestDrawing $Path $items -Sheet 2
}

Test-Case 'Test-Screenshots: 記録してから、元の印刷範囲の中が変わると「元が更新された」、貼り替えると「貼り替えた」と出す（範囲の外のメモは見ない）' {
    $set = New-ScreenshotSet (New-TestDir)
    $opts = @{ Path = $set.Design; Wireframe = $set.Wire }
    Assert-Equal @('受注一覧|記録なし', '受注登録|記録なし', '受注詳細|元のシートなし') (Get-ScreenStates (Invoke-Screenshots $opts)) '表紙のロゴはスクリーンショットとみなさない'
    $recordFile = Join-Path $set.Wire 'スクリーンショットの記録.csv'
    Assert-True (-not (Test-Path -LiteralPath $recordFile)) '-Record を付けなければ記録しない'

    $text = Get-HostText { & $ScreenshotsTool @opts -Record | Out-Null }
    Assert-Match $text '2 画面を記録しました'
    Assert-Equal @('受注一覧|最新', '受注登録|最新', '受注詳細|元のシートなし') (Get-ScreenStates (Invoke-Screenshots $opts))
    Assert-Equal '受注一覧|受注管理_ワイヤーフレーム.xlsx|受注一覧|DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx|受注一覧' (@((Read-TextFile $recordFile | ConvertFrom-Csv) | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.'画面', $_.'元のファイル', $_.'元のシート', $_.'設計書のファイル', $_.'設計書のシート' })[0])

    # 範囲の外のメモ（受注一覧）と、範囲の中のボタンの文字（受注登録）を直す
    Edit-ZipText $set.Book 'xl/worksheets/sheet2.xml' { param($t) $t.Replace('メモ: 文言は仮', 'メモ: 文言は確定') }
    Edit-ZipText $set.Book 'xl/worksheets/sheet3.xml' { param($t) $t.Replace('登録ボタン', '登録する') }
    $results = Invoke-Screenshots $opts
    Assert-Equal @('受注一覧|最新', '受注登録|元が更新された', '受注詳細|元のシートなし') (Get-ScreenStates $results)
    Assert-Match $results[1].Detail '^元のワイヤーフレームが、記録（.+）のあとで変わりました。スクリーンショットを撮り直して貼り替え、記録してください$'

    $text = Get-HostText { & $ScreenshotsTool @opts -Record | Out-Null }
    Assert-Match $text '記録する画面はありませんでした'
    Assert-Match $text '元が更新されたままの画面は記録していません.*受注登録'

    # 撮り直して貼り替える
    New-DesignScreen (Join-Path $set.Design 'DD02_画面設計書_SCR002_受注登録_v1.0.xlsx') '画面レイアウト' '受注登録' 'Retake'
    $results = Invoke-Screenshots $opts
    Assert-Equal '受注登録|貼り替えた' (Get-ScreenStates $results)[1]
    Assert-Match $results[1].Detail '元のワイヤーフレームが変わり、スクリーンショットも貼り替えました.*画像の上に重ねた図形が 2 個あります。位置がずれていないか確かめてください$'
    $null = Invoke-Screenshots ($opts + @{ Record = $true })
    Assert-Equal @('受注一覧|最新', '受注登録|最新', '受注詳細|元のシートなし') (Get-ScreenStates (Invoke-Screenshots $opts))

    # 貼り替えの要らない変更は、-Screen で選んで記録する
    Edit-ZipText $set.Book 'xl/worksheets/sheet2.xml' { param($t) $t.Replace('検索ボタン', '検索') }
    Assert-Equal '受注一覧|元が更新された' (Get-ScreenStates (Invoke-Screenshots $opts))[0]
    $null = Invoke-Screenshots ($opts + @{ Record = $true; Screen = '受注一*' })
    Assert-Equal '受注一覧|最新' (Get-ScreenStates (Invoke-Screenshots $opts))[0]
}

Test-Case 'Test-Screenshots: 図形の位置・文字が変わっても「元が更新された」になる。-WhatIf では記録しない' {
    $set = New-ScreenshotSet (New-TestDir)
    $opts = @{ Path = $set.Design; Wireframe = $set.Wire }
    Assert-Equal @('受注一覧|記録なし', '受注登録|記録なし', '受注詳細|元のシートなし') (Get-ScreenStates (Invoke-Screenshots ($opts + @{ Record = $true; WhatIf = $true }))) '-WhatIf では記録の前の状態'
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $set.Wire 'スクリーンショットの記録.csv'))) '-WhatIf では書かない'
    $null = Invoke-Screenshots ($opts + @{ Record = $true })
    Edit-ZipText $set.Book 'xl/drawings/drawing2.xml' { param($t) $t.Replace('<xdr:col>5</xdr:col>', '<xdr:col>4</xdr:col>') }
    Edit-ZipText $set.Book 'xl/drawings/drawing3.xml' { param($t) $t.Replace('明細', '明細（2 行）') }
    Assert-Equal @('受注一覧|元が更新された', '受注登録|元が更新された', '受注詳細|元のシートなし') (Get-ScreenStates (Invoke-Screenshots $opts))
}

Test-Case 'Test-Screenshots: -OutXlsx に画面ごとの状態と、元のシートの一覧（印刷範囲・設計書で使っているか）を保存する' {
    $set = New-ScreenshotSet (New-TestDir)
    $out = Join-Path $set.Wire '状態.xlsx'
    $null = Invoke-Screenshots @{ Path = $set.Design; Wireframe = $set.Wire; OutXlsx = $out }
    $states = Read-XlsxSheet $out 1
    Assert-Equal '画面|状態|設計書|元のワイヤーフレーム|記録日時|内容' ((@('A', 'B', 'C', 'D', 'E', 'F') | ForEach-Object { $states.Rows[1][$_] }) -join '|')
    Assert-Equal @('受注一覧|記録なし|受注管理_ワイヤーフレーム.xlsx 受注一覧', '受注登録|記録なし|受注管理_ワイヤーフレーム.xlsx 受注登録', '受注詳細|元のシートなし|') @(2..4 | ForEach-Object { '{0}|{1}|{2}' -f $states.Rows[$_]['A'], $states.Rows[$_]['B'], $states.Rows[$_]['D'] })
    Assert-Equal "'画面レイアウト'!B3" @($states.Links | Where-Object { $_.Target -like '*SCR002*' })[0].Location
    $sheets = Read-XlsxSheet $out 2
    Assert-Equal @('表紙|（印刷範囲なし）|設計書に無い', '受注一覧|A1:F20|使っている', '受注登録|A1:F20|使っている') @(2..4 | ForEach-Object { '{0}|{1}|{2}' -f $sheets.Rows[$_]['B'], $sheets.Rows[$_]['C'], $sheets.Rows[$_]['D'] })
}

Test-Case 'Test-Screenshots: 位置で置いた（absoluteAnchor）スクリーンショットも読み、その位置のセルを場所にする' {
    $set = New-ScreenshotSet (New-TestDir)
    # 受注一覧のスクリーンショットを、セルではなく位置（左から 100・上から 60 ピクセル。B3 の中）で、400 × 300 ピクセルに置く
    $pos = '<xdr:absoluteAnchor><xdr:pos x="952500" y="571500"/><xdr:ext cx="3810000" cy="2857500"/>'
    Edit-ZipText (Join-Path $set.Design 'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx') 'xl/drawings/drawing2.xml' { param($t) $t -replace '<xdr:twoCellAnchor>.*?</xdr:to>', $pos -replace '</xdr:twoCellAnchor>', '</xdr:absoluteAnchor>' }
    $results = Invoke-Screenshots @{ Path = $set.Design; Wireframe = $set.Wire }
    Assert-Equal @('受注一覧|記録なし', '受注登録|記録なし', '受注詳細|元のシートなし') (Get-ScreenStates $results)
    Assert-Equal 'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx 受注一覧!B3' $results[0].Location
}

Test-Case 'Test-Screenshots: -Screen は -Record と一緒に使う。Get-Help でヘルプが表示される' {
    Assert-Throws { & $ScreenshotsTool -Path '.' -Wireframe '.' -Screen '受注登録' } '-Record と一緒に'
    $h = Get-Help $ScreenshotsTool
    Assert-Match $h.Synopsis 'スクリーンショット'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
