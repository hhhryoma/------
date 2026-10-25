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

Test-Case 'Test-ExcelDeliverable: ファイル名の「第2期」「第3章」は版数とみなさない' {
    $dir = New-TestDir
    $path = Join-Path $dir '受注システム第2期_第3章_仕様書_第1.1版.xlsx'
    Copy-Item -LiteralPath (Join-Path $printDir '仕様書_第1.1版.xlsx') -Destination $path
    $ok = Invoke-Check @{ Path = $path; Check = 'Version' }
    Assert-Equal 0 $ok.Count ((@($ok | ForEach-Object { $_.Detail })) -join ' / ')
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

function New-DeliverableDocx {
    #     変更履歴・コメント・隠し文字・蛍光ペン・TODO・社内パスへのリンク・テンプレートの場所・作成者の残った Word 文書を作る
    param([string] $Path, [switch] $Clean)
    $w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    $body = '<w:p><w:r><w:t>受注一覧画面の仕様</w:t></w:r></w:p>'
    $parts = [ordered]@{}
    if (-not $Clean) {
        $body += '<w:p><w:ins w:id="1" w:author="山田"><w:r><w:t>追加した文</w:t></w:r></w:ins><w:del w:id="2" w:author="山田"><w:r><w:delText>消した文</w:delText></w:r></w:del></w:p>' +
            '<w:p><w:r><w:t>検索条件は TODO で決める</w:t></w:r></w:p>' +
            '<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>社内メモ</w:t></w:r><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>要確認</w:t></w:r></w:p>' +
            '<w:p><w:hyperlink r:id="rId9"><w:r><w:t>共有フォルダ</w:t></w:r></w:hyperlink></w:p>'
        $parts['word/comments.xml'] = '<?xml version="1.0" encoding="UTF-8"?><w:comments ' + $w + '><w:comment w:id="0" w:author="佐藤"><w:p><w:r><w:t>ここは仮です</w:t></w:r></w:p></w:comment></w:comments>'
        $parts['word/settings.xml'] = '<?xml version="1.0" encoding="UTF-8"?><w:settings ' + $w + '><w:trackRevisions/></w:settings>'
        $parts['word/_rels/document.xml.rels'] = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="\\fs01\share\受注" TargetMode="External"/></Relationships>'
        $parts['word/_rels/settings.xml.rels'] = '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///C:\Users\yamada\テンプレート\設計書.dotx" TargetMode="External"/></Relationships>'
        $parts['word/footer1.xml'] = '<?xml version="1.0" encoding="UTF-8"?><w:ftr ' + $w + '><w:p><w:r><w:t>（仮）社外秘</w:t></w:r></w:p></w:ftr>'
    }
    $parts['docProps/core.xml'] = '<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>' + $(if ($Clean) { '' } else { '山田 太郎' }) + '</dc:creator></cp:coreProperties>'
    $parts['word/document.xml'] = '<?xml version="1.0" encoding="UTF-8"?><w:document ' + $w + '><w:body>' + $body + '</w:body></w:document>'
    $parts['[Content_Types].xml'] = '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    $fs = [System.IO.File]::Create($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($k in $parts.Keys) {
            $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($k).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($parts[$k]); $sw.Dispose()
        }
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

Test-Case 'Test-ExcelDeliverable: Word の変更履歴・コメント・隠し文字・蛍光ペン・TODO・社内パスへのリンク・テンプレートの場所・作成者を見つける' {
    $dir = New-TestDir
    New-DeliverableDocx (Join-Path $dir '画面仕様書.docx')
    New-DeliverableDocx (Join-Path $dir '問題なし.docx') -Clean
    $issues = Invoke-Check @{ Path = $dir }
    Assert-Equal @(@(
        '警告|コメント|',
        '警告|ファイルへのリンク|本文',
        '警告|プロパティ|',
        '警告|変更履歴|',
        '警告|残っている文言|フッター!段落 1',
        '警告|残っている文言|本文!段落 3',
        '警告|残っている文言|本文!段落 4',
        '警告|隠し文字|本文',
        '情報|テンプレートの場所|',
        '情報|変更履歴|',
        '情報|蛍光ペン|') | Sort-Object) (Get-IssueKeys $issues) '問題なしの文書は指摘なし'
    Assert-True (@($issues | Where-Object { $_.Detail -match '変更履歴が残っています（挿入 1・削除 1）' }).Count -eq 1) '変更履歴の数'
    Assert-True (@($issues | Where-Object { $_.Detail -match 'コメントが残っています（佐藤）: ここは仮です' }).Count -eq 1) 'コメントの作成者と文'
    Assert-True (@($issues | Where-Object { $_.Detail -match '隠し文字が 1 か所あります: 社内メモ' }).Count -eq 1) '隠し文字'

    # -Fix はプロパティの作成者だけを直す（変更履歴・コメントは手で直す）
    $fixed = Invoke-Check @{ Path = (Join-Path $dir '画面仕様書.docx'); Fix = $true; Confirm = $false }
    Assert-Equal 0 @($fixed | Where-Object Item -eq 'プロパティ').Count 'プロパティを消した'
    Assert-True (@($fixed | Where-Object Item -eq '変更履歴').Count -gt 0) '変更履歴は残る'
}

Test-Case 'Test-ExcelDeliverable: Word の残っている文言の場所（表の行・テキストボックス）は、検索（Search-Excel -IncludeWord）と同じ' {
    $dir = New-TestDir
    $path = Join-Path $dir '仕様書.docx'
    $ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml"'
    $p = { param($t) '<w:p><w:r><w:t>' + $t + '</w:t></w:r></w:p>' }
    $box = '<w:txbxContent>' + (& $p '要確認の図') + '</w:txbxContent>'
    $body = (& $p '概要') + '<w:tbl><w:tr><w:tc>' + (& $p '項目') + '</w:tc><w:tc>' + (& $p '内容') + '</w:tc></w:tr><w:tr><w:tc>' + (& $p '締め日') + '</w:tc><w:tc>' + (& $p 'TBD') + '</w:tc></w:tr></w:tbl>' +
        (& $p 'TODO の段落') + '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' + $box + '</w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox>' + $box + '</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>'
    $fs = [System.IO.File]::Create($path)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    $sw = New-Object System.IO.StreamWriter($zip.CreateEntry('word/document.xml').Open(), (New-Object System.Text.UTF8Encoding($false)))
    $sw.Write('<w:document ' + $ns + '><w:body>' + $body + '</w:body></w:document>')
    $sw.Dispose(); $zip.Dispose(); $fs.Dispose()

    $issues = Invoke-Check @{ Path = $path; Check = 'Word' }
    $expected = @(@('本文!段落 2', '本文!テキストボックス 1', '本文!表 1 の 2 行目') | Sort-Object)
    Assert-Equal $expected @($issues | ForEach-Object { $_.Location } | Sort-Object) '表の中の段落は段落の番号に数えない。テキストボックスは互換用の写しを数えない'
    $hits = @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') 'TODO|TBD|要確認' $path -IncludeWord 6>$null)
    Assert-Equal $expected @($hits | ForEach-Object { $_.Location } | Sort-Object)
}

Test-Case 'Test-ExcelDeliverable: 画像の縦横比の崩れ・拡大しすぎ・印刷範囲からのはみ出しを見つける（非表示・トリミングした画像は正しく扱う）' {
    $book = Join-Path (New-TestDir) '画面設計書.xlsx'
    New-SheetsBook $book ([ordered]@{ '画面' = @(@('画面名', '受注登録')) })
    Edit-ZipText $book 'xl/workbook.xml' { param($t) $t.Replace('<definedNames>', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">''画面''!$A$1:$H$30</definedName>') }
    Add-TestDrawing $book @(
        @{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 2); To = @(3, 4); Ext = @(40, 20) },
        @{ Kind = 'pic'; Name = '図 2'; Png = 'Old'; From = @(1, 6); To = @(3, 9); Ext = @(80, 60) },
        @{ Kind = 'pic'; Name = '図 3'; Png = 'Old'; From = @(6, 12); To = @(10, 15); Ext = @(40, 30) },
        @{ Kind = 'pic'; Name = '図 4'; Png = 'Old'; From = @(12, 2); To = @(14, 4); Ext = @(40, 30) },
        @{ Kind = 'pic'; Name = '図 5'; Png = 'Old'; From = @(1, 20); To = @(3, 21); Ext = @(80, 20); Hidden = $true },
        @{ Kind = 'pic'; Name = '図 6'; Png = 'Old'; From = @(4, 20); To = @(5, 22); Ext = @(20, 30); Crop = @(50000, 0, 0, 0) })
    $issues = Invoke-Check @{ Path = $book; Check = 'Picture' }
    Assert-Equal @(@('警告|縦横比|画面!B3', '警告|拡大|画面!B7', '警告|印刷範囲の外|画面!G13', '警告|印刷範囲の外|画面!M3') | Sort-Object) (Get-IssueKeys $issues)
    $detail = @{}
    foreach ($i in $issues) { $detail[$i.Cell] = $i.Detail }
    Assert-Equal '画像「図 1」の縦横比が、元の画像と違います（横に 50% 伸びています）。縦横比を元に戻してください' $detail['B3']
    Assert-Equal '画像「図 2」を元の大きさの 200% に拡大して貼っています。印刷・PDF で文字が粗くなるので、元の大きさ（100%）以下にしてください' $detail['B7']
    Assert-Equal '画像「図 3」の一部が、印刷範囲（A1:H30）からはみ出しています。はみ出した部分は印刷されません' $detail['G13']
    Assert-Equal '画像「図 4」が印刷範囲（A1:H30）の外にあり、印刷されません' $detail['M3']
}

Test-Case 'Test-ExcelDeliverable: スクリーンショットの縮尺がシートによって違うことを知らせる（小さい画像は比べない）' {
    $book = Join-Path (New-TestDir) '画面設計書.xlsx'
    New-SheetsBook $book ([ordered]@{ '受注一覧' = @(, @('画面名')); '受注登録' = @(, @('画面名')); 'ロゴ' = @(, @('会社')) })
    Add-TestDrawing $book @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Screen'; From = @(1, 2); To = @(8, 20); Ext = @(400, 300) })
    Add-TestDrawing $book @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Screen'; From = @(1, 2); To = @(6, 14); Ext = @(280, 210) }) -Sheet 2
    Add-TestDrawing $book @(@{ Kind = 'pic'; Name = 'ロゴ'; Png = 'Logo'; From = @(1, 2); To = @(3, 4); Ext = @(20, 15) }) -Sheet 3
    $issues = Invoke-Check @{ Path = $book; Check = 'Picture' }
    Assert-Equal @('情報|縮尺|') (Get-IssueKeys $issues)
    Assert-Equal 'スクリーンショットの縮尺がシートによって違います（100%: 受注一覧 / 70%: 受注登録）。並べたときに文字の大きさがそろいません（撮った画面の拡大率が違うときは、そのままでかまいません）' $issues[0].Detail
}

Test-Case 'Test-ExcelDeliverable: -Fix の -BackupDir には、別々のフォルダの同じ名前のファイルをフォルダごとに分けて残し、前に残したファイルは上書きしない' {
    $dir = New-TestDir
    $hashes = New-Object System.Collections.Generic.List[string]
    foreach ($round in @('第1回', '第2回')) {
        $file = Join-Path $dir "$round/見積書.xlsx"
        New-Item -ItemType Directory -Force -Path (Split-Path $file -Parent) | Out-Null
        Copy-Item -LiteralPath (Join-Path $delivDir 'bad.xlsx') -Destination $file
        Set-CellValues $file '詳細' @{ A1 = $round }
        $hashes.Add((Get-FileHash -LiteralPath $file).Hash)
    }
    $roots = @((Join-Path $dir '第1回'), (Join-Path $dir '第2回'))
    $backup = Join-Path $dir 'bk'
    # 1 回目はプロパティだけを直す
    $null = Invoke-Check @{ Path = $roots; Check = 'Property'; Fix = $true; BackupDir = $backup }
    Assert-Equal $hashes.ToArray() @(@('第1回', '第2回') | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $backup "$_/見積書.xlsx")).Hash }) '指定したフォルダの名前のフォルダに、元のファイルを残す'

    # 2 回目（表示の状態を直す）でも、前に残したファイルは上書きせず、番号を付けて残す
    $fixed = (Get-FileHash -LiteralPath (Join-Path $dir '第1回/見積書.xlsx')).Hash
    $text = Get-HostText { $null = & $Deliverable -Path $roots -Check View -Fix -BackupDir $backup }
    Assert-Equal $hashes[0] (Get-FileHash -LiteralPath (Join-Path $backup '第1回/見積書.xlsx')).Hash '1 回目に残した元のファイル'
    Assert-Equal $fixed (Get-FileHash -LiteralPath (Join-Path $backup '第1回/見積書 (2).xlsx')).Hash '2 回目に直す前のファイル'
    Assert-Match $text '番号を付けて残し.*見積書 \(2\)\.xlsx'
}

Test-Case 'Test-ExcelDeliverable: -BackupDir の中のファイルはチェックせず、前に残したファイルを直さない' {
    $dir = New-TestDir
    Copy-Item -LiteralPath (Join-Path $delivDir 'bad.xlsx') -Destination $dir
    $hash = (Get-FileHash -LiteralPath (Join-Path $dir 'bad.xlsx')).Hash
    $backup = Join-Path $dir '直す前'
    $null = Invoke-Check @{ Path = $dir; Fix = $true; BackupDir = $backup }
    $issues = Invoke-Check @{ Path = $dir; Fix = $true; BackupDir = $backup }
    Assert-Equal @('bad.xlsx') @($issues | ForEach-Object { $_.File } | Sort-Object -Unique) '-BackupDir の中のファイル'
    Assert-Equal $hash (Get-FileHash -LiteralPath (Join-Path $backup 'bad.xlsx')).Hash '前に残した直す前のファイル'
    # 設計書を読むツールが、このフォルダを読まないための印のファイル
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $backup '書き換える前のファイル.txt'))) 'Test-ExcelDeliverable\.ps1 が書き換える前のファイル'
}

function New-ViewBook {
    #     シートの表示（<sheetView>）を書き換えたブックを作る（テスト用）。Views: シート名 → <sheetView> の XML の順序付き辞書
    param([string] $Path, $Views)
    $sheets = [ordered]@{}
    foreach ($name in $Views.Keys) { $sheets[$name] = @(@('No', '項目'), @('1', 'x')) }
    New-SheetsBook $Path $sheets
    $i = 0
    foreach ($name in $Views.Keys) {
        $i++
        $view = $Views[$name]
        Edit-ZipText $Path "xl/worksheets/sheet$i.xml" { param($t) [regex]::Replace($t, '<sheetViews>.*?</sheetViews>', '<sheetViews>' + $view + '</sheetViews>') }
    }
}

Test-Case 'Test-ExcelDeliverable: スクロールした状態でウィンドウ枠を固定したシートは、固定した行・列の直後をスクロール領域の左上とみなす' {
    # スクロールした状態で固定すると、Excel は固定した部分の左上を sheetView の topLeftCell に、スクロール領域の左上を pane の topLeftCell に保存する
    $book = Join-Path (New-TestDir) '画面一覧.xlsx'
    New-ViewBook $book ([ordered]@{
            # 1〜9 行目（タイトル）を上に隠して、10〜11 行目（表の見出し）を固定した
            '行'     = '<sheetView tabSelected="1" topLeftCell="A10" workbookViewId="0"><pane ySplit="2" topLeftCell="A12" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A12" sqref="A12"/></sheetView>'
            # 右へスクロールして、J〜K 列を固定した
            '列'     = '<sheetView topLeftCell="J1" workbookViewId="0"><pane xSplit="2" topLeftCell="L1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="L1" sqref="L1"/></sheetView>'
            # C 列と 10〜11 行目を固定した
            '行と列' = '<sheetView topLeftCell="C10" workbookViewId="0"><pane xSplit="1" ySplit="2" topLeftCell="D12" activePane="bottomRight" state="frozen"/><selection pane="topRight" activeCell="D10" sqref="D10"/><selection pane="bottomLeft" activeCell="C12" sqref="C12"/><selection pane="bottomRight" activeCell="D12" sqref="D12"/></sheetView>'
        })
    $issues = Invoke-Check @{ Path = $book; Check = 'View' }
    Assert-Equal 0 $issues.Count ((@($issues | ForEach-Object { '{0} {1}' -f $_.Location, $_.Detail })) -join ' / ')

    # 固定した行・列より上・左の行・列は、固定を解除しないと表示されないので、非表示の行・列と同じく情報として知らせる
    $hidden = Invoke-Check @{ Path = $book; Check = 'Hidden' }
    Assert-Equal (@(
            '情報|行|固定した行（10〜11 行目）より上の 1〜9 行目は、ウィンドウ枠の固定を解除しないと表示されません',
            '情報|列|固定した列（J〜K 列）より左の A〜I 列は、ウィンドウ枠の固定を解除しないと表示されません',
            '情報|行と列|固定した行（10〜11 行目）より上の 1〜9 行目は、ウィンドウ枠の固定を解除しないと表示されません',
            '情報|行と列|固定した列（C 列）より左の A〜B 列は、ウィンドウ枠の固定を解除しないと表示されません') | Sort-Object) @($hidden | ForEach-Object { '{0}|{1}|{2}' -f $_.Level, $_.Location, $_.Detail } | Sort-Object)
}

Test-Case 'Test-ExcelDeliverable: ウィンドウ枠を固定したシートでも、スクロールしたまま保存したシートと、表示されない A1 を選択したシートは知らせる' {
    $book = Join-Path (New-TestDir) '画面一覧.xlsx'
    New-ViewBook $book ([ordered]@{
            # 10〜11 行目を固定し、スクロール領域を 30 行目までスクロールした
            'スクロール' = '<sheetView tabSelected="1" topLeftCell="A10" workbookViewId="0"><pane ySplit="2" topLeftCell="A30" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A12" sqref="A12"/></sheetView>'
            # A〜B 列を固定して下へスクロールした（列だけを固定したときの sheetView の topLeftCell の行は、スクロールの位置）
            '列・下へ'   = '<sheetView topLeftCell="A50" workbookViewId="0"><pane xSplit="2" topLeftCell="C50" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="C1" sqref="C1"/></sheetView>'
            # 10〜11 行目を固定して右へスクロールした
            '行・右へ'   = '<sheetView topLeftCell="D10" workbookViewId="0"><pane ySplit="2" topLeftCell="D12" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A12" sqref="A12"/></sheetView>'
            # 選択した A1 は、固定した行より上にあって表示されない
            '上の行'     = '<sheetView topLeftCell="A10" workbookViewId="0"><pane ySplit="2" topLeftCell="A12" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A1" sqref="A1"/></sheetView>'
        })
    $issues = Invoke-Check @{ Path = $book; Check = 'View' }
    Assert-Equal (@('警告|表示位置|スクロール', '警告|表示位置|列・下へ', '警告|表示位置|行・右へ', '警告|選択セル|上の行!A1') | Sort-Object) (Get-IssueKeys $issues) ((@($issues | ForEach-Object { '{0} {1}' -f $_.Location, $_.Detail })) -join ' / ')
    Assert-Equal '選択セルが A1 です（A12 にしてください）' (@($issues | Where-Object Item -eq '選択セル')[0].Detail)
}

Test-Case 'Test-ExcelDeliverable: シート全体・列全体を選択したまま保存したシートを知らせる（結合したセルの選択は知らせない）' {
    $book = Join-Path (New-TestDir) '一覧.xlsx'
    New-ViewBook $book ([ordered]@{
            '全体' = '<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1:XFD1048576"/></sheetView>'
            '列'   = '<sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1:A1048576"/></sheetView>'
            '結合' = '<sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1:C2"/></sheetView>'
        })
    $issues = Invoke-Check @{ Path = $book; Check = 'View' }
    Assert-Equal (@('警告|選択セル|全体!A1', '警告|選択セル|列!A1') | Sort-Object) (Get-IssueKeys $issues)
    Assert-Equal '選択範囲が A1:XFD1048576 です（A1 だけを選択してください）' (@($issues | Where-Object Location -eq '全体!A1')[0].Detail)
    # Reset-ExcelView.ps1（-Fix）で A1 だけの選択に直る
    $null = & (Join-Path $script:ToolsDir 'Reset-ExcelView.ps1') -Path $book 6>$null 3>$null
    Assert-Equal 0 (Invoke-Check @{ Path = $book; Check = 'View' }).Count
}

Test-Case 'Test-ExcelDeliverable: ウィンドウに合わせて拡大縮小するグラフシートの倍率は確かめない' {
    $book = Join-Path (New-TestDir) '集計.xlsx'
    New-SheetsBook $book ([ordered]@{ '一覧' = @(@('No', '件数'), @('1', '3')); 'グラフ' = @(, @('No')); '固定倍率' = @(, @('No')) })
    # 2・3 つ目をグラフシートにする（グラフの中身は省く）。ウィンドウに合わせる（zoomToFit）と、zoomScale は保存したときの画面の大きさで変わる
    Edit-ZipText $book 'xl/_rels/workbook.xml.rels' { param($t) $t.Replace('relationships/worksheet" Target="worksheets/sheet2.xml"', 'relationships/chartsheet" Target="worksheets/sheet2.xml"').Replace('relationships/worksheet" Target="worksheets/sheet3.xml"', 'relationships/chartsheet" Target="worksheets/sheet3.xml"') }
    $chart = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews>{0}</sheetViews><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></chartsheet>'
    Edit-ZipText $book 'xl/worksheets/sheet2.xml' { param($t) $chart -f '<sheetView zoomScale="88" workbookViewId="0" zoomToFit="1"/>' }
    Edit-ZipText $book 'xl/worksheets/sheet3.xml' { param($t) $chart -f '<sheetView zoomScale="88" workbookViewId="0"/>' }
    $issues = Invoke-Check @{ Path = $book; Check = 'View' }
    Assert-Equal @('警告|倍率|固定倍率') (Get-IssueKeys $issues)
}

Test-Case 'Test-ExcelDeliverable: マクロシート（Excel 4.0 マクロ）も、ワークシートと同じように選択セルを確かめる' {
    $book = Join-Path (New-TestDir) '集計.xlsm'
    New-ViewBook $book ([ordered]@{
            '一覧'   = '<sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView>'
            'マクロ' = '<sheetView workbookViewId="0"><selection activeCell="C5" sqref="C5"/></sheetView>'
        })
    # 2 つ目をマクロシートにする（リレーションの種類は .../xlMacrosheet）
    Edit-ZipText $book 'xl/_rels/workbook.xml.rels' { param($t) $t.Replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"', 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet" Target="worksheets/sheet2.xml"') }
    $issues = Invoke-Check @{ Path = $book; Check = 'View' }
    Assert-Equal @('警告|選択セル|マクロ!C5') (Get-IssueKeys $issues)
}

Test-Case 'Test-ExcelDeliverable: 改訂履歴に同じ版数の行が続いても（変更点ごとの行）指摘せず、版数が前の行より小さい行だけを知らせる' {
    $book = Join-Path (New-TestDir) '受注仕様書.xlsx'
    New-SheetsBook $book ([ordered]@{
            '表紙'     = @(@('版数', '1.1'), @('更新日', '2026/10/12'))
            '改訂履歴' = @(@('版数', '改訂日', '改訂内容'), @('1.0', '2026/10/01', '初版'), @('1.1', '2026/10/12', '画面項目を追加'), @('1.1', '2026/10/12', 'メッセージを修正'))
        })
    $issues = Invoke-Check @{ Path = $book; Check = 'Version' }
    Assert-Equal 0 $issues.Count ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')
    # 1.0 → 1.2 → 1.1
    $issues = Invoke-Check @{ Path = (Join-Path $printDir '仕様書_v1.2.xlsx'); Check = 'Version' }
    Assert-Equal '改訂履歴の版数（1.1）が前の行（1.2）より小さくなっています' (@($issues | Where-Object Item -eq '改訂履歴の版数')[0].Detail)
}

Test-Case 'Test-ExcelDeliverable: 表紙の更新日と改訂履歴の改訂日は、和暦の文字（令和8年10月1日・R8.10.1）でも比べる' {
    $dir = New-TestDir
    # 表紙の更新日と、改訂履歴の最後の改訂日（4 つ目だけそろっている）
    $pairs = @(@('令和8年10月1日', '2026/10/12'), @('R8.10.1', 'R8.10.12'), @('2026/10/1', '令和8年10月12日'), @('R8.10.12', '令和8年10月12日'))
    for ($i = 0; $i -lt $pairs.Count; $i++) {
        New-SheetsBook (Join-Path $dir "仕様書$($i + 1).xlsx") ([ordered]@{
                '表紙'     = @(@('版数', '1.1'), @('更新日', $pairs[$i][0]))
                '改訂履歴' = @(@('版数', '改訂日', '改訂内容'), @('1.0', '2026/09/01', '初版'), @('1.1', $pairs[$i][1], '画面項目を追加'))
            })
    }
    # 地域の設定が和暦でも、日付は西暦で示す
    $issues = Invoke-InWareki { Invoke-Check @{ Path = $dir; Check = 'Version' } }
    Assert-Equal @('仕様書1.xlsx', '仕様書2.xlsx', '仕様書3.xlsx') @($issues | ForEach-Object { $_.File } | Sort-Object)
    foreach ($i in $issues) { Assert-Equal '表紙の更新日（2026/10/01）と、改訂履歴の最後の改訂日（2026/10/12）が違います' $i.Detail $i.File }
}

Test-Case 'Test-ExcelDeliverable: 相対パスで保存された社内のファイルへのリンク（../../社内/見積根拠.xlsx）も知らせる（チェックするフォルダの中のファイルへのリンクと URL は知らせない）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $root = Join-Path (New-TestDir) '提出物'
    $book = Join-Path $root '01_要件定義/要件定義書.xlsx'
    New-Item -ItemType Directory -Force -Path (Split-Path $book -Parent) | Out-Null
    # Excel・Word は、同じドライブ・共有フォルダのファイルへのリンクを、ファイルの場所からの相対パスで保存する
    $targets = @(
        '../../%E7%A4%BE%E5%86%85/%E8%A6%8B%E7%A9%8D%E6%A0%B9%E6%8B%A0.xlsx',
        '..\..\社内\メモ.xlsx',
        '../02_基本設計/基本設計書.xlsx',
        '別紙/画面一覧.xlsx',
        'https://example.com/manual',
        'mailto:support@example.com')
    $rows = @(for ($i = 0; $i -lt $targets.Count; $i++) { @{ Values = @("リンク$i"); Link = @{ Column = 0; Path = "link$i.xlsx"; Location = '' } } })
    Export-Workbook -FilePath $book -Sheets @(@{ Name = '参照'; Columns = @(@{ Header = 'リンク'; Width = 20 }); Rows = $rows })
    Edit-ZipText $book 'xl/worksheets/_rels/sheet1.xml.rels' { param($t) for ($i = 0; $i -lt $targets.Count; $i++) { $t = $t.Replace("file:///link$i.xlsx", $targets[$i]) }; $t }
    $docx = Join-Path $root '01_要件定義/画面仕様書.docx'
    New-DeliverableDocx $docx
    Edit-ZipText $docx 'word/_rels/document.xml.rels' { param($t) $t.Replace('\\fs01\share\受注', '../../社内/受注') }

    $issues = Invoke-Check @{ Path = $root; Check = 'Link' }
    $issues = @($issues | Where-Object Item -eq 'ファイルへのリンク')
    Assert-Equal (@(
            '参照!A2|社内のファイルやフォルダへのリンクがあります（../../社内/見積根拠.xlsx）',
            '参照!A3|社内のファイルやフォルダへのリンクがあります（..\..\社内\メモ.xlsx）',
            '本文|社内のファイルやフォルダへのリンクがあります（../../社内/受注）') | Sort-Object) @($issues | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Detail } | Sort-Object) 'チェックするフォルダの中の 02_基本設計・別紙へのリンクは、一緒に提出するファイル'
}

Test-Case 'Test-ExcelDeliverable: 列全体・行全体の印刷範囲（$A:$E・$1:$40）も印刷範囲として確かめる' {
    $dir = New-TestDir
    Copy-Item -LiteralPath (Join-Path $printDir 'wide.xlsx') -Destination $dir
    $wide = Join-Path $dir 'wide.xlsx'
    # 一覧（A1〜T1 に値）は A〜E 列、範囲（A1〜C10 と E20 に値）は 1〜5 行目と 7〜10 行目
    Edit-ZipText $wide 'xl/workbook.xml' { param($t) $t.Replace('<definedNames>', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">''一覧''!$A:$E</definedName>').Replace('''範囲''!$A$1:$C$10', '''範囲''!$1:$5,''範囲''!$7:$10') }
    $issues = Invoke-Check @{ Path = $wide; Check = 'Print' }
    Assert-Equal (@('情報|用紙の向き|', '警告|ヘッダー／フッター|一覧', '警告|印刷範囲の外|一覧!F1', '警告|印刷範囲の外|範囲!A6') | Sort-Object) (Get-IssueKeys $issues) '横のはみ出しは、印刷範囲の列（A〜E 列）と値のある列（行全体のとき）で見積もる'
    Assert-Match (@($issues | Where-Object Location -eq '範囲!A6')[0].Detail) '4 個あります（A6、B6、C6、E20）'

    # 画像の印刷範囲からのはみ出しも、列全体の印刷範囲で確かめる
    $book = Join-Path $dir '画面設計書.xlsx'
    New-SheetsBook $book ([ordered]@{ '画面' = @(@('画面名', '受注登録')) })
    Edit-ZipText $book 'xl/workbook.xml' { param($t) $t.Replace('<definedNames>', '<definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">''画面''!$A:$H</definedName>') }
    Add-TestDrawing $book @(
        @{ Kind = 'pic'; Name = '図 1'; Png = 'Old'; From = @(1, 40); To = @(3, 43); Ext = @(40, 30) },
        @{ Kind = 'pic'; Name = '図 2'; Png = 'Old'; From = @(12, 2); To = @(14, 4); Ext = @(40, 30) })
    $issues = Invoke-Check @{ Path = $book; Check = 'Picture' }
    Assert-Equal @('警告|印刷範囲の外|画面!M3') (Get-IssueKeys $issues) '41 行目の図 1 は、列全体の印刷範囲の中'
    Assert-Equal '画像「図 2」が印刷範囲（A:H）の外にあり、印刷されません' $issues[0].Detail
}

Test-Case 'Test-ExcelDeliverable: ヘッダー／フッターの文字の「&Z」（&&Z と書く）はファイルの場所とみなさない' {
    $dir = New-TestDir
    Copy-Item -LiteralPath (Join-Path $printDir 'wide.xlsx') -Destination $dir
    $wide = Join-Path $dir 'wide.xlsx'
    # 一覧のフッターは文字の「A&Z 社」とファイル名、収めるのフッターは文字の「&」とファイルの場所
    Edit-ZipText $wide 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('<oddFooter>&amp;L&amp;Z&amp;F</oddFooter>', '<oddFooter>&amp;LA&amp;&amp;Z 社&amp;F</oddFooter>') }
    Edit-ZipText $wide 'xl/worksheets/sheet2.xml' { param($t) $t.Replace('<oddFooter>&amp;C&amp;P / &amp;N</oddFooter>', '<oddFooter>&amp;C&amp;&amp;&amp;Z</oddFooter>') }
    $issues = Invoke-Check @{ Path = $wide; Check = 'Print' }
    Assert-Equal @('収める') @($issues | Where-Object Item -eq 'ヘッダー／フッター' | ForEach-Object { $_.Location })
}

Test-Case 'Test-ExcelDeliverable: Get-Help でヘルプが表示される' {
    $h = Get-Help $Deliverable
    Assert-Match $h.Synopsis '提出'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
