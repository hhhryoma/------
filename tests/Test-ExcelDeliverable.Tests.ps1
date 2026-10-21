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

$printDir = Join-Path $script:Fixtures 'print'

Test-Case 'Test-ExcelDeliverable: 印刷で横がはみ出すシート・印刷範囲の外のデータ・ヘッダー／フッターの問題を見つける' {
    $issues = Invoke-Check @{ Path = (Join-Path $printDir 'wide.xlsx'); Check = @('Print', 'Word') }
    Assert-Equal @(
        '情報|用紙の向き|',
        '警告|ヘッダー／フッター|一覧',
        '警告|印刷範囲の外|範囲!E20',
        '警告|横のはみ出し|一覧',
        '警告|残っている文言|一覧') (Get-IssueKeys $issues)
    Assert-Match (@($issues | Where-Object Item -eq '横のはみ出し')[0].Detail) '約 5 ページ.*A4・縦向き'
    Assert-Match (@($issues | Where-Object Item -eq 'ヘッダー／フッター')[0].Detail) '&Z'
    Assert-Match (@($issues | Where-Object Item -eq '用紙の向き')[0].Detail) '縦: 一覧、収める / 横: 範囲'
}

Test-Case 'Test-ExcelDeliverable: -RequirePageNumber と -HeaderFooterText で、ページ番号や「社外秘」の無いシートを知らせる' {
    $issues = Invoke-Check @{ Path = (Join-Path $printDir 'wide.xlsx'); Check = 'Print'; RequirePageNumber = $true; HeaderFooterText = '社外秘' }
    Assert-Equal @('一覧', '範囲') @($issues | Where-Object Item -eq 'ページ番号' | ForEach-Object { $_.Location } | Sort-Object) '「収める」はフッターにページ番号あり'
    Assert-Equal @('一覧', '収める', '範囲') @($issues | Where-Object { $_.Detail -match '社外秘' } | ForEach-Object { $_.Location } | Sort-Object)
}

Test-Case 'Test-ExcelDeliverable: 表紙の版数・更新日と、改訂履歴の最後の行・ファイル名の版数の食い違いを見つける' {
    $issues = Invoke-Check @{ Path = (Join-Path $printDir '仕様書_v1.2.xlsx'); Check = 'Version' }
    Assert-Equal @(
        '警告|ファイル名の版数|',
        '警告|改訂履歴の版数|改訂履歴!A6',
        '警告|表紙の更新日|表紙!D6',
        '警告|表紙の版数|表紙!D5') (Get-IssueKeys $issues)
    Assert-Match (@($issues | Where-Object Item -eq '表紙の版数')[0].Detail) '表紙の版数（1\.0）と、改訂履歴の最後の版数（1\.1）'
    Assert-Match (@($issues | Where-Object Item -eq '表紙の更新日')[0].Detail) '2026/10/01.*2026/10/12'
    Assert-Match (@($issues | Where-Object Item -eq 'ファイル名の版数')[0].Detail) 'ファイル名の版数（1\.2）'

    # 「第1.1版」「Ver.1.1」「2026年10月12日」のような書き方でも、そろっていれば指摘なし
    $ok = Invoke-Check @{ Path = (Join-Path $printDir '仕様書_第1.1版.xlsx'); Check = 'Version' }
    Assert-Equal 0 $ok.Count ((Get-IssueKeys $ok) -join ', ')
    # 改訂履歴のシートの上部に「版数」「更新日」が並んでいても、改訂履歴の表を読む
    $ok = Invoke-Check @{ Path = (Join-Path $printDir '帳票ヘッダ付き.xlsx'); Check = 'Version' }
    Assert-Equal 0 $ok.Count ((Get-IssueKeys $ok) -join ', ')
}

Test-Case 'Test-ExcelDeliverable: 詳細設計サンプルは印刷・版数を含めて指摘なし' {
    $issues = Invoke-Check @{ Path = (Join-Path $script:RepoRoot 'docs/design-sample') }
    Assert-Equal 0 $issues.Count ((@($issues | ForEach-Object { '{0} {1}' -f $_.File, $_.Detail })) -join ' / ')
}

Test-Case 'Test-ExcelDeliverable: LibreOffice で保存したブック（pane の無い選択範囲）も表示の状態を確かめられる' {
    # LibreOffice は <pane> が無くても <selection pane="topLeft"> を書く
    $dir = New-TestDir
    $path = Join-Path $dir '見積書.xlsx'
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'estimate/lo/見積書.xlsx') -Destination $path
    $issues = Invoke-Check @{ Path = $path; Check = 'View' }
    Assert-Equal 0 $issues.Count ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')

    # 選択セルを C5 にすると、読み取りのエラーではなく選択セルの警告になる
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
        $sr = New-Object System.IO.StreamReader($entry.Open())
        $xml = $sr.ReadToEnd(); $sr.Dispose()
        Assert-Match $xml '<selection pane="topLeft" activeCell="A1"'
        $xml = $xml.Replace('<selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>', '<selection pane="topLeft" activeCell="C5" activeCellId="0" sqref="C5"/>')
        $entry.Delete()
        $sw = New-Object System.IO.StreamWriter($zip.CreateEntry('xl/worksheets/sheet1.xml').Open())
        $sw.Write($xml); $sw.Dispose()
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
    $issues = Invoke-Check @{ Path = $path; Check = 'View' }
    Assert-Equal @('警告|選択セル|見積!C5') (Get-IssueKeys $issues) ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')
}

function Read-ZipText {
    param([string] $Path, [string] $Name)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $sr = New-Object System.IO.StreamReader($zip.GetEntry($Name).Open())
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    }
    finally { $fs.Dispose() }
}

Test-Case 'Test-ExcelDeliverable: -Fix でプロパティの個人名・ヘッダーのファイルの場所・表示の状態を直し、残っている指摘を出す' {
    $dir = New-TestDir
    Copy-Item -LiteralPath (Join-Path $delivDir 'bad.xlsx') -Destination $dir
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'print/wide.xlsx') -Destination $dir
    $bad = Join-Path $dir 'bad.xlsx'; $wide = Join-Path $dir 'wide.xlsx'
    $hash = (Get-FileHash -LiteralPath $bad).Hash
    # -WhatIf では書き換えない
    $issues = Invoke-Check @{ Path = $dir; Fix = $true; WhatIf = $true }
    Assert-Equal $hash (Get-FileHash -LiteralPath $bad).Hash '-WhatIf'
    Assert-True (@($issues | Where-Object Item -eq 'プロパティ').Count -eq 3)

    $backup = Join-Path (New-TestDir) 'backup'
    $issues = Invoke-Check @{ Path = $dir; Fix = $true; AllowedNames = '株式会社サンプル'; BackupDir = $backup }
    Assert-Equal 0 @($issues | Where-Object { $_.Item -eq 'プロパティ' -or ($_.Item -eq 'ヘッダー／フッター' -and $_.Detail -like '*&Z*') }).Count ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')
    Assert-Equal @('表示モード') @($issues | Where-Object { $_.Category -eq '体裁' } | ForEach-Object { $_.Item }) '表示モード（改ページプレビュー）は意図どおりのことがあるので直さない'
    Assert-True (@($issues | Where-Object Item -eq '非表示シート').Count -eq 2) '内容に関わる指摘は残す'
    $core = Read-ZipText $bad 'docProps/core.xml'
    Assert-NotMatch $core '山田 太郎|佐藤 花子'
    Assert-Match (Read-ZipText $bad 'docProps/app.xml') '<Company>株式会社サンプル</Company>' '-AllowedNames に一致する会社名は残す'
    Assert-Match (Read-ZipText $wide 'xl/worksheets/sheet1.xml') '<oddFooter>&amp;L&amp;F</oddFooter>' 'ファイル名（&F）は残す'
    Assert-Equal $hash (Get-FileHash -LiteralPath (Join-Path $backup 'bad.xlsx')).Hash '直す前のファイル'
}

Test-Case 'Test-ExcelDeliverable: Get-Help でヘルプが表示される' {
    $h = Get-Help $Deliverable
    Assert-Match $h.Synopsis '提出'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
