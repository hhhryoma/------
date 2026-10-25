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

Test-Case 'Test-ExcelDeliverable: Get-Help でヘルプが表示される' {
    $h = Get-Help $Deliverable
    Assert-Match $h.Synopsis '提出'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
