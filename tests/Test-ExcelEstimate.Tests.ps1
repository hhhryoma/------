#     Test-ExcelEstimate.ps1 のテスト。フィクスチャは tests\fixtures\estimate（make_fixtures.py で生成）

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

Test-Case 'Test-ExcelEstimate: Get-Help でヘルプが表示される' {
    $h = Get-Help $EstimateTool
    Assert-Match $h.Synopsis '見積書'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
