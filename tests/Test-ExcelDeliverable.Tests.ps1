#     Test-ExcelDeliverable.ps1 のテスト。フィクスチャは tests\fixtures\deliverable（make_fixtures.py で生成）

$Deliverable = Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1'
$delivDir = Join-Path $script:Fixtures 'deliverable'

function Invoke-Check {
    param([hashtable] $Options)
    return , @(& $Deliverable @Options 6>$null)
}

function Get-IssueKeys {
    param($Issues)
    return @($Issues | ForEach-Object { '{0}|{1}|{2}' -f $_.Level, $_.Item, $_.Location } | Sort-Object)
}

Test-Case 'Test-ExcelDeliverable: 問題のないブックは指摘なし（見出し固定のシートを含む）' {
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'good.xlsx') }
    Assert-Equal 0 $issues.Count ((Get-IssueKeys $issues) -join ', ')
}

Test-Case 'Test-ExcelDeliverable: 体裁・消し忘れ・情報漏れ・壊れをすべて検出する' {
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'bad.xlsx') }
    $expected = @(
        'エラー|エラー値|表紙!B2',
        'エラー|名前の定義|',
        '警告|非表示シート|作業用',
        '警告|非表示シート|隠し',
        '警告|残っている文言|表紙!C5',
        '警告|プロパティ|',
        '警告|プロパティ|',
        '警告|プロパティ|',
        '警告|外部リンク|',
        '警告|作業グループ|',
        '警告|先頭シート|詳細',
        '警告|表示位置|表紙',
        '警告|選択セル|表紙!C5',
        '警告|倍率|表紙',
        '警告|コメント|表紙!D1',
        '警告|ファイルへのリンク|表紙!A10',
        '情報|非表示の行・列|表紙',
        '情報|表示モード|詳細'
    ) | Sort-Object
    Assert-Equal $expected (Get-IssueKeys $issues)
    $detail = ($issues | ForEach-Object { $_.Detail }) -join "`n"
    Assert-Match $detail '作成者「山田 太郎」'
    Assert-Match $detail '最終更新者「佐藤 花子」'
    Assert-Match $detail '会社名「株式会社サンプル」'
    Assert-Match $detail '\\\\fileserver\\share\\原価表\.xlsx'
    Assert-Match $detail '作業グループ）: 表紙、詳細'
    Assert-Match $detail 'veryHidden'
}

Test-Case 'Test-ExcelDeliverable: エラーを先に、同じファイルの指摘はまとめて並べる' {
    $issues = Invoke-Check @{ Path = $delivDir }
    $levels = @($issues | Where-Object { $_.File -eq 'bad.xlsx' } | ForEach-Object { $_.Level })
    Assert-Equal @('エラー', 'エラー') @($levels[0..1]) '先頭はエラー'
    Assert-Equal '情報' $levels[-1] '最後は情報'
}

Test-Case 'Test-ExcelDeliverable: マクロ・パスワード付き・xls は情報として知らせる' {
    $issues = Invoke-Check @{ Path = $delivDir }
    $info = @($issues | Where-Object { $_.File -ne 'bad.xlsx' } | ForEach-Object { '{0}|{1}' -f $_.File, $_.Item } | Sort-Object)
    Assert-Equal @('encrypted.xlsx|パスワード', 'legacy.xls|形式', 'macro.xlsm|マクロ') $info
}

Test-Case 'Test-ExcelDeliverable: ファイルを変更しない' {
    $dir = Copy-Fixture 'deliverable'
    $before = @{}
    foreach ($f in Get-ChildItem -LiteralPath $dir -File) { $before[$f.Name] = @((Get-FileHash -LiteralPath $f.FullName).Hash, $f.LastWriteTimeUtc) }
    $null = Invoke-Check @{ Path = $dir }
    foreach ($f in Get-ChildItem -LiteralPath $dir -File) {
        Assert-Equal $before[$f.Name][0] (Get-FileHash -LiteralPath $f.FullName).Hash "$($f.Name) の中身"
        Assert-Equal $before[$f.Name][1] $f.LastWriteTimeUtc "$($f.Name) の更新日時"
    }
}

Test-Case 'Test-ExcelDeliverable: -AllowedNames に一致するプロパティは指摘しない' {
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'bad.xlsx'); Check = 'Property'; AllowedNames = '株式会社サンプル|山田' }
    Assert-Equal 1 $issues.Count
    Assert-Match $issues[0].Detail '佐藤 花子'
}

Test-Case 'Test-ExcelDeliverable: -Check で項目を絞り込める' {
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'bad.xlsx'); Check = 'View' }
    Assert-Equal @('体裁') @($issues | ForEach-Object { $_.Category } | Sort-Object -Unique)
    Assert-Equal 6 $issues.Count
}

Test-Case 'Test-ExcelDeliverable: -Zoom で期待する倍率を変えられる' {
    # 表紙は 85%、詳細は 100%
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'bad.xlsx'); Check = 'View'; Zoom = 85 }
    $zoom = @($issues | Where-Object Item -eq '倍率')
    Assert-Equal 1 $zoom.Count
    Assert-Equal '詳細' $zoom[0].Location
    Assert-Match $zoom[0].Detail '100% です（85% にしてください）'
}

Test-Case 'Test-ExcelDeliverable: -Words で探す文言を変えられる' {
    $issues = Invoke-Check @{ Path = (Join-Path $delivDir 'bad.xlsx'); Check = 'Word'; Words = 'メモ付き' }
    Assert-Equal 1 $issues.Count
    Assert-Equal '表紙!D1' $issues[0].Location
}

Test-Case 'Test-ExcelDeliverable: -FileNamePattern に合わないファイル名を警告する' {
    $issues = Invoke-Check @{ Path = $delivDir; Check = 'Macro'; FileNamePattern = '^(good|bad)\.xlsx$' }
    $named = @($issues | Where-Object Item -eq 'ファイル名' | ForEach-Object { $_.File } | Sort-Object)
    Assert-Equal @('encrypted.xlsx', 'legacy.xls', 'macro.xlsm') $named
}

Test-Case 'Test-ExcelDeliverable: -OutXlsx で該当セルへのリンク付きの一覧を保存する' {
    $out = Join-Path (New-TestDir) 'check.xlsx'
    $issues = Invoke-Check @{ Path = $delivDir; OutXlsx = $out }
    $list = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('指摘一覧', 'ファイル別') $list.Sheets
    Assert-Equal ($issues.Count + 1) $list.Rows.Count '見出し＋指摘'
    Assert-True (@($list.Links | Where-Object { $_.Location -eq "'表紙'!C5" }).Count -ge 1) 'セルへのリンク'
    Assert-True (@($list.Links | Where-Object { -not $_.Location -and $_.Target -like 'file:///*bad.xlsx' }).Count -ge 1) '場所の無い指摘はファイルを開くリンク'
    $files = Read-XlsxSheet -Path $out -Index 2
    $result = @{}
    foreach ($r in 2..$files.Rows.Count) { $result[$files.Rows[$r]['A']] = $files.Rows[$r]['B'] }
    Assert-Equal 'OK' $result['good.xlsx']
    Assert-Equal '要対応' $result['bad.xlsx']
    Assert-Equal 'OK' $result['macro.xlsm'] '情報だけなら OK'
}

Test-Case 'Test-ExcelDeliverable: Get-Help でヘルプが表示される' {
    $h = Get-Help $Deliverable
    Assert-Match $h.Synopsis '提出'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
