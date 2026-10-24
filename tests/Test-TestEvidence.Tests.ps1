#     Test-TestEvidence.ps1 のテスト。テスト仕様書とエビデンス（ブック・画像）はテストの中で作る

$EvidenceTool = Join-Path $script:ToolsDir 'Test-TestEvidence.ps1'

function New-EvidencePng {
    #     テスト用の画像のファイルを作る（Png は $script:TestPng のキー）
    param([string] $FilePath, [string] $Png)
    $dir = Split-Path -Parent $FilePath
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    [System.IO.File]::WriteAllBytes($FilePath, [Convert]::FromBase64String($script:TestPng[$Png]))
}

function Get-IssueKeys {
    param($Issues)
    return @($Issues | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Level, $_.Item, $_.Id, $_.Location })
}

function New-EvidenceSet {
    #     テスト仕様書（UT_受注一覧.xlsx）と、エビデンスのフォルダ（ブック・画像・ケースごとのフォルダ）を作る
    param([string] $Dir)
    $spec = Join-Path $Dir 'UT_受注一覧.xlsx'
    New-SheetsBook $spec ([ordered]@{ 'テストケース' = @(
            @('テストID', '確認内容', '結果', 'エビデンス'),
            @('UT-001', '一覧が表示される', 'OK', ''),        # シート UT-001
            @('UT-002', '検索できる', 'OK', ''),              # UT-002.png
            @('UT-003', 'クリアできる', 'OK', ''),            # UT-003\1.png（ケースごとのフォルダ）
            @('UT-004', '並べ替えできる', 'OK', ''),          # まとめのシートの目印
            @('UT-005', '削除できる', 'NG', ''),              # 無い
            @('UT-006', '印刷できる', 'OK', '-'),             # 要らない
            @('UT-007', '出力できる', '', ''),                # 結果が空欄なのにある
            @('UT-008', '戻れる', 'OK', 'あり'),              # あると書いたのに無い
            @('UT-009', '件数が出る', 'OK', '画面確認'),      # エビデンスの列のシート
            @('UT-010', '閉じられる', 'OK', 'UT-999.png'),    # エビデンスの列のファイルが無い
            @('UT-011', '開ける', 'OK', '')) })               # 空のシート
    $ev = Join-Path $Dir 'エビデンス'
    New-Item -ItemType Directory -Path $ev | Out-Null
    $book = Join-Path $ev 'UT_受注一覧_エビデンス.xlsx'
    New-SheetsBook $book ([ordered]@{
            'UT-001' = @(, @('一覧の画面'))
            'まとめ' = @(@('まとめ'), @('説明'), @('■UT-004 並べ替え'), @('受注日の順'))
            '画面確認' = @(, @('件数: 3 件'))
            '画面一覧' = @(, @('画面'))
            'UT-011' = @(, @('')) })
    Add-TestDrawing -Path $book -Sheet 1 -Items @(@{ Kind = 'pic'; Name = '画面'; Png = 'Screen'; From = @(0, 1); To = @(6, 12) })
    # まとめ: 目印（3 行目）より上の画像はどのケースか分からない。下の画像は UT-004
    Add-TestDrawing -Path $book -Sheet 2 -Items @(@{ Kind = 'pic'; Name = '上'; Png = 'New'; From = @(2, 0); To = @(4, 1) }, @{ Kind = 'pic'; Name = '並べ替え'; Png = 'Old'; From = @(0, 4); To = @(4, 8) })
    Add-TestDrawing -Path $book -Sheet 4 -Items @(@{ Kind = 'pic'; Name = '一覧'; Png = 'Logo'; From = @(0, 1); To = @(4, 8) })
    New-EvidencePng (Join-Path $ev 'UT-002.png') 'New'
    New-EvidencePng (Join-Path $ev 'UT-003/1.png') 'Retake'
    New-EvidencePng (Join-Path $ev 'UT-007.png') 'Logo'
    return [pscustomobject]@{ Spec = $spec; Evidence = $ev; Book = $book }
}

Test-Case 'Test-TestEvidence: シート名・画像の名前・フォルダ名・目印・エビデンスの列でケースと結び付け、エビデンスの無いケースを見つける' {
    $set = New-EvidenceSet (New-TestDir)
    $issues = @(& $EvidenceTool -Path $set.Spec -EvidencePath $set.Evidence 6>$null)
    $keys = Get-IssueKeys $issues
    Assert-Equal @(
        'エラー|エビデンスが無い|UT-005|テストケース!D6',
        'エラー|エビデンスが無い|UT-008|テストケース!D9',
        'エラー|エビデンスが見つからない|UT-010|テストケース!D11',
        '警告|実施していないケースのエビデンス|UT-007|テストケース!D8',
        '警告|エビデンスが空|UT-011|テストケース!D12') @($keys | Select-Object -First 5)
    Assert-Equal 7 $keys.Count ($keys -join "`n")
    Assert-True ($keys -contains '情報|どのケースか分からないエビデンス||まとめ!A1') '目印より上の画像'
    Assert-True ($keys -contains '情報|どのケースか分からないエビデンス||画面一覧!A1') 'ケースに合わないシートの画像'
    $byId = @{}
    foreach ($i in $issues) { if ($i.Id) { $byId[$i.Id] = $i.Detail } }
    Assert-Equal '結果が NG ですが、エビデンスが見つかりません（シート名・ファイル名・目印をテストID にしてください）' $byId['UT-005']
    Assert-Match $byId['UT-008'] '^エビデンスの列は「あり」ですが、エビデンスが見つかりません'
    Assert-Match $byId['UT-010'] '^エビデンスの列の「UT-999\.png」が見つかりません'
    Assert-Equal '結果が空欄ですが、エビデンスがあります（UT-007.png）。結果の書き忘れではありませんか' $byId['UT-007']
    Assert-Equal 'エビデンス（UT_受注一覧_エビデンス.xlsx UT-011!A1）に、画像も文字もありません' $byId['UT-011']

    $text = Get-HostText { & $EvidenceTool -Path $set.Spec -EvidencePath $set.Evidence | Out-Null }
    Assert-Match $text 'テストケース 11 件（実施済み 10 件）: エビデンスあり 6 件・無し 3 件・不要 1 件'
    Assert-Match $text 'エビデンス: シート 5・目印 1・画像 3（画像 7 枚）。どのケースか分からないもの 2・決められないもの 0'
    Assert-Match $text 'エラー 3 件 / 警告 2 件 / 情報 2 件'
}

Test-Case 'Test-TestEvidence: -OutXlsx にケース別・エビデンス別のシートを出し、エビデンスの場所へリンクする' {
    $set = New-EvidenceSet (New-TestDir)
    $out = Join-Path (Split-Path -Parent $set.Spec) 'エビデンスのチェック.xlsx'
    $null = & $EvidenceTool -Path $set.Spec -EvidencePath $set.Evidence -OutXlsx $out 6>$null
    $issues = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('指摘', 'ケース別', 'エビデンス別', '読み取り結果') $issues.Sheets
    Assert-True (@($issues.Links | Where-Object { $_.Location -eq "'テストケース'!D6" }).Count -eq 1) 'テスト仕様書のエビデンスの列へのリンク'
    $cases = Read-XlsxSheet -Path $out -Index 2
    $rows = @(2..($cases.Rows.Count) | ForEach-Object { $r = $cases.Rows[$_]; '{0}|{1}|{2}|{3}|{4}' -f $r['D'], $r['E'], $r['G'], $r['H'], $r['I'] })
    Assert-Equal @(
        'UT-001|OK|あり|1|UT_受注一覧_エビデンス.xlsx UT-001!A1',
        'UT-002|OK|あり|1|UT-002.png',
        'UT-003|OK|あり|1|1.png',
        'UT-004|OK|あり|1|UT_受注一覧_エビデンス.xlsx まとめ!A3',
        'UT-005|NG|無し|0|',
        'UT-006|OK|不要|0|',
        'UT-007||あり|1|UT-007.png',
        'UT-008|OK|無し|0|',
        'UT-009|OK|あり|0|UT_受注一覧_エビデンス.xlsx 画面確認!A1',
        'UT-010|OK|無し|0|',
        'UT-011|OK|あり|0|UT_受注一覧_エビデンス.xlsx UT-011!A1') $rows
    Assert-True (@($cases.Links | Where-Object { $_.Location -eq "'まとめ'!A3" }).Count -eq 1) '目印のセルへのリンク'
    # 結果のブックをテスト仕様書と同じフォルダに置いても、テスト仕様書として読まない
    $again = @(& $EvidenceTool -Path (Split-Path -Parent $set.Spec) -EvidencePath $set.Evidence 6>$null)
    Assert-Equal 7 $again.Count (($again | ForEach-Object { '{0}|{1}|{2}' -f $_.Item, $_.Id, $_.File }) -join "`n")
    $units = Read-XlsxSheet -Path $out -Index 3
    $rows = @(2..($units.Rows.Count) | ForEach-Object { $r = $units.Rows[$_]; '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $r['A'], $r['B'], $r['C'], $r['D'], $r['E'], $r['F'], $r['G'] })
    Assert-True ($rows -contains 'UT_受注一覧_エビデンス.xlsx|まとめ!A3|目印|1|1|UT-004|') ($rows -join "`n")
    Assert-True ($rows -contains 'UT_受注一覧_エビデンス.xlsx|まとめ!A1|シート|1|2||どのケースか分からない') ($rows -join "`n")
    Assert-True ($rows -contains '1.png||画像|1|0|UT-003|') ($rows -join "`n")
}

Test-Case 'Test-TestEvidence: 違うケースのエビデンスに同じ画像があれば使い回しを知らせ、多くのケースにある画像（ロゴなど）は除く' {
    $dir = New-TestDir
    $spec = Join-Path $dir 'UT_受注登録.xlsx'
    $rows = @(, @('No', '確認内容', '結果'))
    for ($i = 1; $i -le 8; $i++) { $rows += , @([string]$i, "確認 $i", 'OK') }
    New-SheetsBook $spec ([ordered]@{ 'ケース' = $rows })
    $ev = Join-Path $dir 'エビデンス'
    New-Item -ItemType Directory -Path $ev | Out-Null
    $book = Join-Path $ev '受注登録_エビデンス.xlsx'
    $sheets = [ordered]@{}
    for ($i = 1; $i -le 6; $i++) { $sheets["No.$i"] = @(, @("ケース $i")) }
    New-SheetsBook $book $sheets
    # どのシートにもロゴ（共通の画像）。No.1 と No.2 は同じ画面の画像（使い回し）
    $screens = @('Screen', 'Screen', 'Old', 'New', 'Retake', '')
    for ($i = 1; $i -le 6; $i++) {
        $items = @(@{ Kind = 'pic'; Name = 'ロゴ'; Png = 'Logo'; From = @(5, 0); To = @(6, 1) })
        if ($screens[$i - 1]) { $items += @{ Kind = 'pic'; Name = '画面'; Png = $screens[$i - 1]; From = @(0, 2); To = @(5, 10) } }
        Add-TestDrawing -Path $book -Sheet $i -Items $items
    }
    # ブックの名前がケース（7）なら、ブック全体がそのケースのエビデンス
    $whole = Join-Path $ev '7.xlsx'
    New-SheetsBook $whole ([ordered]@{ '画面' = @(, @('登録後')); 'ログ' = @(, @('登録しました')) })
    Add-TestDrawing -Path $whole -Sheet 1 -Items @(@{ Kind = 'pic'; Name = 'ロゴ'; Png = 'Logo'; From = @(5, 0); To = @(6, 1) })

    $issues = @(& $EvidenceTool -Path $spec -EvidencePath $ev 6>$null)
    Assert-Equal @(
        'エラー|エビデンスが無い|8|ケース!A9',
        '警告|同じ画像の使い回し|1|ケース!A2',
        '警告|同じ画像の使い回し|2|ケース!A3',
        '情報|多くのケースにある画像||7.xlsx') (Get-IssueKeys $issues)
    Assert-Equal 'エビデンスの画像が、2（受注登録_エビデンス.xlsx No.2!A1）のエビデンスの画像と同じです（撮り直し忘れ・貼り間違いではありませんか）' $issues[1].Detail
    Assert-Equal '同じ画像が 7 件のケースのエビデンスにあります（1・2・3 ほか）。ロゴなどの共通の画像なら問題ありません' $issues[3].Detail
    $text = Get-HostText { & $EvidenceTool -Path $spec -EvidencePath $ev | Out-Null }
    Assert-Match $text 'エビデンス: ブック 1・シート 6（画像 12 枚）'
}

Test-Case 'Test-TestEvidence: 機能ごとに同じテストID でも、置き場所の名前で機能を見分け、テスト仕様書の表の中の画像も読む' {
    $dir = New-TestDir
    $spec = Join-Path $dir 'IT_受注.xlsx'
    New-SheetsBook $spec ([ordered]@{ 'ケース' = @(
            @('機能', 'No', '確認内容', '結果'),
            @('受注一覧', '1', '一覧', 'OK'),
            @('', '2', '検索', 'OK'),
            @('受注登録', '1', '登録', 'OK'),
            @('', '2', '取消', 'OK'),
            @('', '3', '確認', 'OK')) })
    # 表の中に貼った画像は、その行のケース（受注登録の 3）。表より上の画像（ロゴ）は見ない
    Add-TestDrawing -Path $spec -Items @(@{ Kind = 'pic'; Name = 'ロゴ'; Png = 'Logo'; From = @(6, 0); To = @(8, 1) }, @{ Kind = 'pic'; Name = '確認'; Png = 'Screen'; From = @(6, 5); To = @(10, 9) })
    New-EvidencePng (Join-Path $dir '受注一覧/1.png') 'Old'
    New-EvidencePng (Join-Path $dir '受注一覧_No2.png') 'New'
    New-EvidencePng (Join-Path $dir 'その他/1.png') 'Retake'
    $book = Join-Path $dir '受注登録_エビデンス.xlsx'
    New-SheetsBook $book ([ordered]@{ '1' = @(, @('登録の画面')); '2' = @(, @('取消しました')) })
    Add-TestDrawing -Path $book -Sheet 1 -Items @(@{ Kind = 'pic'; Name = '画面'; Png = 'Retake'; From = @(0, 1); To = @(5, 9) })
    $out = Join-Path $dir 'check.xlsx'

    # エビデンスの置き場所を指定しなければ、テスト仕様書と同じフォルダ
    $issues = @(& $EvidenceTool -Path $dir -OutXlsx $out 6>$null)
    Assert-Equal @('情報|どのケースか決められない||1.png') (Get-IssueKeys $issues)
    Assert-Equal '「1」に合うケースが 2 件あります（1（受注一覧）・1（受注登録））。ブック・フォルダの名前に機能の名前を入れるか、エビデンスの列にシート名・ファイル名を書いてください' $issues[0].Detail
    $cases = Read-XlsxSheet -Path $out -Index 2
    $rows = @(2..($cases.Rows.Count) | ForEach-Object { $r = $cases.Rows[$_]; '{0}|{1}|{2}|{3}' -f $r['C'], $r['D'], $r['G'], $r['I'] })
    Assert-Equal @(
        '受注一覧|1|あり|1.png',
        '受注一覧|2|あり|受注一覧_No2.png',
        '受注登録|1|あり|受注登録_エビデンス.xlsx 1!A1',
        '受注登録|2|あり|受注登録_エビデンス.xlsx 2!A1',
        '受注登録|3|あり|IT_受注.xlsx ケース!G6') $rows
}

Test-Case 'Test-TestEvidence: エビデンスは、サブフォルダのものも含めてパスの順に並べる' {
    $dir = New-TestDir
    $spec = Join-Path $dir 'UT.xlsx'
    New-SheetsBook $spec ([ordered]@{ 'ケース' = @(@('No', '結果'), @('1', 'OK'), @('2', 'OK')) })
    $ev = Join-Path $dir 'エビデンス'
    # Get-ChildItem は、フォルダの直下のファイルを先に返す
    New-EvidencePng (Join-Path $ev 'UT-2.png') 'Old'
    New-EvidencePng (Join-Path $ev 'A/1.png') 'New'
    $out = Join-Path $dir 'check.xlsx'
    $null = & $EvidenceTool -Path $spec -EvidencePath $ev -OutXlsx $out 6>$null
    $units = Read-XlsxSheet -Path $out -Index 3
    Assert-Equal @('1.png', 'UT-2.png') @(2..($units.Rows.Count) | ForEach-Object { $units.Rows[$_]['A'] })
}

Test-Case 'Test-TestEvidence: Get-Help でヘルプが表示される' {
    $h = Get-Help $EvidenceTool
    Assert-Match $h.Synopsis 'エビデンス'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
