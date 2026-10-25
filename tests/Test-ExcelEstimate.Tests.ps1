#     Test-ExcelEstimate.ps1 のテスト。フィクスチャは tests\fixtures\estimate（make_fixtures.py で生成）。
#     数式のある小さな見積書はテストの中で作る（New-SheetsBook・Edit-ZipText は Invoke-Tests.ps1）

$EstimateTool = Join-Path $script:ToolsDir 'Test-ExcelEstimate.ps1'
$estDir = Join-Path $script:Fixtures 'estimate'

function Invoke-Estimate {
    param([hashtable] $Options)
    return , @(& $EstimateTool @Options 6>$null 3>$null)
}

function Get-EstimateKeys {
    param($Issues)
    return @($Issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Level, $_.Item, $_.Location } | Sort-Object)
}

function New-EstimateBook {
    #     数式のある見積書を一時フォルダに作る（New-SheetsBook の表で、= で始まる文字のセルを数式にする）。返り値: ファイルのパス
    param($Sheets)
    $path = Join-Path (New-TestDir) '見積書.xlsx'
    New-SheetsBook $path $Sheets
    for ($i = 1; $i -le $Sheets.Count; $i++) {
        Edit-ZipText $path "xl/worksheets/sheet$i.xml" { param($t) $t -replace '<c r="([A-Z]+\d+)" s="2" t="inlineStr"><is><t xml:space="preserve">=([^<]*)</t></is></c>', '<c r="$1"><f>$2</f></c>' }
    }
    return $path
}

$expectedEstimate = @(
    '警告|換算の日数|内訳!C2',
    '警告|文字列の数値|見積!D7',
    '警告|合計の範囲|見積!D9',
    '警告|単価|見積!E5',
    '警告|手入力の合計|見積!F10',
    '警告|周りと違う数式|見積!F4',
    '警告|数式の列の数値|見積!F6',
    '情報|端数処理|内訳!D2') | Sort-Object

foreach ($variant in @('', 'lo')) {
    $label = 'openpyxl で作った見積書（セル内の文字列）'
    $file = Join-Path $estDir '見積書.xlsx'
    if ($variant) { $label = 'LibreOffice で保存し直した見積書（共有文字列・計算結果あり）'; $file = Join-Path (Join-Path $estDir $variant) '見積書.xlsx' }
    Test-Case "Test-ExcelEstimate: 合計の範囲漏れ・手入力の合計・違う数式・文字列の数値・換算・単価・端数処理を見つける（$label）" {
        $issues = Invoke-Estimate @{ Path = $file }
        Assert-Equal $expectedEstimate (Get-EstimateKeys $issues)
        $byItem = @{}
        foreach ($i in $issues) { $byItem[$i.Item] = $i.Detail }
        Assert-Match $byItem['合計の範囲'] '=SUM\(D2:D7\).*D8 が入っていません'
        Assert-Match $byItem['手入力の合計'] '「合計」の行の F10 が数式ではなく数値（4,000,000）'
        Assert-Match $byItem['周りと違う数式'] '上下のセル（=D3\*E3）と違う計算です（=D4\*E3）'
        Assert-Match $byItem['換算の日数'] '20: 見積!D11.* / 21: 内訳!C2'
        Assert-Match $byItem['単価'] '「SE」の単価が複数あります（50,000: 見積!E2 ほか 1 か所 / 45,000: 見積!E5）'
    }
}

Test-Case 'Test-ExcelEstimate: 共有数式（コピーした数式）の写しも、元の数式から計算を求めて比べる' {
    $issues = Invoke-Estimate @{ Path = (Join-Path $estDir '共有数式.xlsx') }
    Assert-Equal (@('警告|周りと違う数式|明細!C4', '警告|合計の範囲|明細!C6') | Sort-Object) (Get-EstimateKeys $issues)
    Assert-Match (@($issues | Where-Object Item -eq '周りと違う数式')[0].Detail) '上下のセル（=A3\*B3）と違う計算です（=A4\*B3）'
    Assert-Match (@($issues | Where-Object Item -eq '合計の範囲')[0].Detail) 'C5 が入っていません'
}

Test-Case 'Test-ExcelEstimate: 見積書ではない設計書では誤って指摘しない。-Check で項目を絞れる' {
    $issues = Invoke-Estimate @{ Path = (Join-Path $script:RepoRoot 'docs/design-sample') }
    Assert-Equal 0 $issues.Count ((Get-EstimateKeys $issues) -join ', ')
    $only = Invoke-Estimate @{ Path = (Join-Path $estDir '見積書.xlsx'); Check = @('Sum', 'UnitPrice') }
    Assert-Equal (@('警告|合計の範囲|見積!D9', '警告|単価|見積!E5') | Sort-Object) (Get-EstimateKeys $only)
}

Test-Case 'Test-ExcelEstimate: エラー値を指摘し、パスワード付きのファイルは情報として知らせる' {
    $dir = Copy-Fixture 'deliverable'
    $issues = Invoke-Estimate @{ Path = $dir }
    Assert-True (@($issues | Where-Object { $_.Item -eq 'エラー値' -and $_.Location -eq '表紙!B2' }).Count -eq 1) 'bad.xlsx の #REF!'
    Assert-Equal '情報' (@($issues | Where-Object File -eq 'encrypted.xlsx')[0].Level)
}

Test-Case 'Test-ExcelEstimate: -OutXlsx で該当セルへのリンク付きの一覧を保存し、ファイルは変更しない' {
    $dir = Copy-Fixture 'estimate'
    $file = Join-Path $dir '見積書.xlsx'
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $out = Join-Path (New-TestDir) 'check.xlsx'
    $null = Invoke-Estimate @{ Path = $file; OutXlsx = $out }
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal 9 $sheet.Rows.Count '見出しと 8 件'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'見積'!D9" }).Count -eq 1) '合計のセルへのリンク'
}

Test-Case 'Test-ExcelEstimate: Excel で開いている（書き込みでロックされた）見積書もチェックする' -WindowsOnly {
    $dir = Copy-Fixture 'estimate'
    $file = Join-Path $dir '見積書.xlsx'
    # Excel で開いたときと同じく、書き込みのために開き、ほかには読み取りだけを許す
    $lock = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
    try { $issues = Invoke-Estimate @{ Path = $file } }
    finally { $lock.Dispose() }
    Assert-Equal $expectedEstimate (Get-EstimateKeys $issues)
}

Test-Case 'Test-ExcelEstimate: マクロシート（Excel 4.0 マクロ。リレーションの種類は xlMacrosheet）の数式もチェックする' {
    $file = New-EstimateBook ([ordered]@{
        '表紙' = @(@('見積書', ''), @('件名', '受注管理システム'))
        '見積' = @(@('工程', '金額'), @('設計', 100000), @('製造', 200000), @('試験', 300000), @('合計', '=SUM(B2:B3)')) })
    Edit-ZipText $file 'xl/_rels/workbook.xml.rels' { param($t) $t.Replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"', 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="worksheets/sheet2.xml"') }
    Edit-ZipText $file '[Content_Types].xml' { param($t) $t.Replace('/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"', '/xl/worksheets/sheet2.xml" ContentType="application/vnd.ms-excel.macrosheet+xml"') }
    Edit-ZipText $file 'xl/worksheets/sheet2.xml' { param($t) $t.Replace('<worksheet xmlns=', '<xm:macrosheet xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main" xmlns=').Replace('</worksheet>', '</xm:macrosheet>') }
    $issues = Invoke-Estimate @{ Path = $file }
    Assert-Equal @('警告|合計の範囲|見積!B5') (Get-EstimateKeys $issues)
    Assert-Match $issues[0].Detail 'B4 が入っていません'
}

Test-Case 'Test-ExcelEstimate: 率を掛ける数式（*20%・*8%）は、換算の日数・時間とみなさない' {
    # 人月の換算は 21 日、人時の換算は 7.5 時間にそろっている。PM 管理の 20%・8% は換算ではない
    $file = New-EstimateBook ([ordered]@{ '見積' = @(
        @('工程', '人日', '人月', '金額', '人時'),
        @('設計', 42, '=B2/21', '=C2*800000', '=B2*7.5'),
        @('製造', 63, '=B3/21', '=C3*800000', '=B3*7.5'),
        @('PM 管理', '', '', '=SUM(D2:D3)*20%', '=SUM(E2:E3)*8%'),
        @('合計', '', '', '=SUM(D2:D4)', '=SUM(E2:E4)')) })
    $issues = Invoke-Estimate @{ Path = $file }
    Assert-Equal 0 $issues.Count ((Get-EstimateKeys $issues) -join ', ')
}

Test-Case 'Test-ExcelEstimate: 合計の範囲の外でも、合計の数式が使っているセル（値引きを引く =SUM(B2:B3)-B4 など）は入れ忘れとみなさない' {
    # 案A・案B は値引きの 2 行を数式で引いている。案C は 4 行目の値引きを引き忘れている（条件!D4 はほかのシートのセル）
    $file = New-EstimateBook ([ordered]@{ '見積' = @(
        @('項目', '案A', '案B', '案C'),
        @('設計', 100, 100, 100),
        @('製造', 200, 200, 200),
        @('値引き', 10, 10, 10),
        @('端数値引き', 1, 1, 1),
        @('合計', '=SUM(B2:B3)-B4-$B$5', '=SUM(C2:C3)-SUM(C4:C5)', '=SUM(D2:D3)*(1-条件!D4)-D5')) })
    $issues = Invoke-Estimate @{ Path = $file }
    Assert-Equal '警告|合計の範囲|見積!D6' ((Get-EstimateKeys $issues) -join ', ')
    Assert-Match $issues[0].Detail '範囲に D4 が入っていません'
}

Test-Case 'Test-ExcelEstimate: 合計の範囲の上にある表の見出し（年度の数値・月の日付）は、入れ忘れとみなさない' {
    # 年度別: 1 行目は表題で、2 行目の見出しに年度（数値）。月別: 「2026年度」の下の 2 行目に月（日付）。明細: 2 行目の設計を合計に入れ忘れている
    $file = New-EstimateBook ([ordered]@{
        '年度別' = @(@('要員計画', '', ''), @('要員', 2026, 2027), @('PM', 1, 1), @('SE', 2, 3), @('PG', 3, 4), @('合計', '=SUM(B3:B5)', '=SUM(C3:C5)'))
        '月別' = @(@('要員', '2026年度', ''), @('', [datetime]'2026-04-01', [datetime]'2026-05-01'), @('PM', 1, 1), @('SE', 2, 3), @('PG', 3, 4), @('合計', '=SUM(B3:B5)', '=SUM(C3:C5)'))
        '明細' = @(@('項目', '金額'), @('設計', 100), @('製造', 200), @('テスト', 300), @('合計', '=SUM(B3:B4)')) })
    $issues = Invoke-Estimate @{ Path = $file }
    Assert-Equal '警告|合計の範囲|明細!B5' ((Get-EstimateKeys $issues) -join ', ') '見出しではない、範囲のすぐ上の数値は指摘する'
    Assert-Match $issues[0].Detail '範囲に B2 が入っていません'
}

Test-Case 'Test-ExcelEstimate: 合計の範囲の上の、元号（[$-411]ggge"年度"）・曜日（aaa）の表示形式の日付の見出しは入れ忘れとみなさない（G/標準"人"・経過時間 [h]:mm:ss の数値は指摘する）' {
    # 年度別: 2 行目の見出しは年度の初日の日付を「令和8年度」と表示。曜日別: 2 行目の見出しは日付を「木」と表示。
    # 人数・工数: 2 行目の設計の人数（G/標準"人"）・時間（組み込みの [h]:mm:ss）を合計に入れ忘れている
    $file = New-EstimateBook ([ordered]@{
        '年度別' = @(@('要員', '当年度', '翌年度'), @('', 46113, 46478), @('PM', 1, 1), @('SE', 2, 3), @('PG', 3, 4), @('合計', '=SUM(B3:B5)', '=SUM(C3:C5)'))
        '曜日別' = @(@('作業', '初日', '翌日'), @('', 46296, 46297), @('設計', 1, 2), @('製造', 3, 4), @('合計', '=SUM(B3:B4)', '=SUM(C3:C4)'))
        '人数' = @(@('項目', '人数'), @('設計', 3), @('製造', 5), @('テスト', 4), @('合計', '=SUM(B3:B4)'))
        '工数' = @(@('項目', '時間'), @('設計', 0.5), @('製造', 1.25), @('テスト', 0.75), @('合計', '=SUM(B3:B4)')) })
    # 2 行目の値に、上の表示形式（177〜179 と組み込みの 46）を付ける（シートの順に 5〜8 番の書式）
    Edit-ZipText $file 'xl/styles.xml' {
        param($x)
        $x = $x.Replace('</numFmts>', '<numFmt numFmtId="177" formatCode="[$-411]ggge&quot;年度&quot;"/><numFmt numFmtId="178" formatCode="aaa"/><numFmt numFmtId="179" formatCode="G/標準&quot;人&quot;"/></numFmts>')
        $x.Replace('<cellXfs count="5">', '<cellXfs count="9">').Replace('</cellXfs>', ((@(177, 178, 179, 46) | ForEach-Object { '<xf numFmtId="{0}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' -f $_ }) -join '') + '</cellXfs>')
    }
    for ($i = 1; $i -le 4; $i++) {
        Edit-ZipText $file "xl/worksheets/sheet$i.xml" { param($x) $x.Replace('<c r="B2" s="2">', ('<c r="B2" s="{0}">' -f ($i + 4))).Replace('<c r="C2" s="2">', ('<c r="C2" s="{0}">' -f ($i + 4))) }
    }
    $issues = Invoke-Estimate @{ Path = $file }
    $keys = @(Get-EstimateKeys $issues)
    Assert-Equal 2 $keys.Count ($keys -join ', ')
    Assert-True ($keys -contains '警告|合計の範囲|人数!B5') ($keys -join ', ')
    Assert-True ($keys -contains '警告|合計の範囲|工数!B5') ($keys -join ', ')
}

Test-Case 'Test-ExcelEstimate: Get-Help でヘルプが表示される' {
    $h = Get-Help $EstimateTool
    Assert-Match $h.Synopsis '見積書'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
