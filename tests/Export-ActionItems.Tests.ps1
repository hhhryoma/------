#     Export-ActionItems.ps1 のテスト。議事録（テキスト・Word・Excel）はテストの中で作る

$ActionItems = Join-Path $script:ToolsDir 'Export-ActionItems.ps1'

function New-MinutesDocx {
    #     段落と表だけの最小の Word 文書（word/document.xml）を作る
    param([string] $Path, [string[]] $Before, [string[][]] $Table, [string[]] $After)
    $esc = { param($t) [System.Security.SecurityElement]::Escape($t) }
    $p = { param($t) '<w:p><w:r><w:t xml:space="preserve">' + (& $esc $t) + '</w:t></w:r></w:p>' }
    $body = New-Object System.Text.StringBuilder
    foreach ($t in $Before) { [void]$body.Append((& $p $t)) }
    [void]$body.Append('<w:tbl>')
    foreach ($row in $Table) {
        [void]$body.Append('<w:tr>')
        foreach ($c in $row) { [void]$body.Append('<w:tc>' + (& $p $c) + '</w:tc>') }
        [void]$body.Append('</w:tr>')
    }
    [void]$body.Append('</w:tbl>')
    foreach ($t in $After) { [void]$body.Append((& $p $t)) }
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $body.ToString() + '</w:body></w:document>'
    $fs = [System.IO.File]::Create($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($entry in @(@('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'), @('word/document.xml', $xml))) {
            $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($entry[0]).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($entry[1]); $sw.Dispose()
        }
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

Test-Case 'Export-ActionItems: テキストの議事録の宿題・決定事項の節と【TODO】から、担当・期限を取り出して一覧にする' {
    $dir = New-TestDir
    @('定例会議 議事録', '日時：2026年10月20日（火）10:00～11:00', '会議名：受注システム 定例', '',
        '■ 議題', '1. 進捗の確認', '・画面設計は予定どおり。山田さんが説明した。', '',
        '■ 決定事項', '・受注日の初期値は当日とする。', '',
        '■ 宿題',
        '・画面設計書の受注日の桁数を直す（山田、10/23まで）',
        '・テーブル定義書のレビュー日程を決める　担当：佐藤　期限：10/27',
        '　※ 鈴木さんの予定も確認',
        '・お客様に締め日の扱いを確認する（来週中）',
        '・性能の目標値を確認する',
        '・なし', '',
        '■ その他', '・次回は10/27。【TODO】議事録を共有する @鈴木 今週中') | Set-Content -LiteralPath (Join-Path $dir '議事録.txt') -Encoding UTF8
    $text = Get-HostText { $script:r = @(& $ActionItems (Join-Path $dir '議事録.txt') -OutCsv (Join-Path $dir '宿題.csv')) }
    Assert-Equal @(
        '決定事項|受注日の初期値は当日とする||||受注システム 定例|2026/10/20',
        '宿題|画面設計書の受注日の桁数を直す|山田|2026/10/23||受注システム 定例|2026/10/20',
        '宿題|テーブル定義書のレビュー日程を決める|佐藤|2026/10/27|鈴木さんの予定も確認|受注システム 定例|2026/10/20',
        '宿題|お客様に締め日の扱いを確認する||2026/10/30|担当なし|受注システム 定例|2026/10/20',
        '宿題|性能の目標値を確認する|||担当なし・期限なし|受注システム 定例|2026/10/20',
        '宿題|議事録を共有する|鈴木|2026/10/23||受注システム 定例|2026/10/20') @($script:r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}' -f $_.Kind, $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }), $_.Note, $_.Meeting, $_.MeetingDate.ToString('yyyy/MM/dd') })
    Assert-Match $text '議事録 1 件: 宿題 5 件（担当なし 2・期限なし 1）・決定事項 1 件'

    # CSV は課題管理表の形で、週次報告の -IssueList にそのまま渡せる
    $csv = @(Import-Csv -LiteralPath (Join-Path $dir '宿題.csv') -Encoding UTF8)
    Assert-Equal 'No|区分|件名|担当|期限|状態|起票日|会議|メモ|出典' ((@($csv[0].PSObject.Properties | ForEach-Object { $_.Name })) -join '|')
    Assert-Equal '1|宿題|画面設計書の受注日の桁数を直す|山田|2026/10/23|未対応|2026/10/20' ((@('No', '区分', '件名', '担当', '期限', '状態', '起票日') | ForEach-Object { $csv[0].$_ }) -join '|')
    Assert-Equal '議事録.txt 13 行目' $csv[0].'出典'
    $weekly = Get-HostText { $null = & (Join-Path $script:ToolsDir 'New-WeeklyReport.ps1') -IssueList (Join-Path $dir '宿題.csv') -AsOf '2026-10-28' }
    Assert-Match $weekly '・未完了: 5 件・うち期限切れ 3 件'
}

Test-Case 'Export-ActionItems: Word の議事録の表（アクション・担当・期限）と【決定】、Excel の議事録の表を読む' {
    $dir = New-TestDir
    $docx = Join-Path $dir '20261021_定例.docx'
    New-MinutesDocx $docx @('定例会議 議事録') @(@('No', 'アクション', '担当', '期限'), @('1', 'IF 定義書を更新する', '佐藤さん', '2026/10/28'), @('2', '受入テストの範囲を決める', '鈴木', '次回')) @('【決定】結合テストは11/9から始める')
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $xlsx = Join-Path $dir '定例_2026-10-22.xlsx'
    $rows = @(@('画面設計書を直す', '山田', '10/30'), @('帳票の様式を確認する', '', '11/2'), @('', '', ''), @('■ 決定事項', '', ''), @('・帳票は PDF で出す', '', ''))
    Export-Workbook -FilePath $xlsx -Sheets @(@{ Name = '議事録'; Columns = @(@('項目', '担当', '期限') | ForEach-Object { @{ Header = $_; Width = 20 } }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    $out = Join-Path $dir '宿題.xlsx'
    $r = @(& $ActionItems $dir -OutXlsx $out 6>$null)
    Assert-Equal @(
        '宿題|画面設計書を直す|山田|2026/10/30|定例_2026-10-22',
        '宿題|帳票の様式を確認する||2026/11/02|定例_2026-10-22',
        '決定事項|帳票は PDF で出す|||定例_2026-10-22',
        '宿題|IF 定義書を更新する|佐藤|2026/10/28|20261021_定例',
        '宿題|受入テストの範囲を決める|鈴木||20261021_定例',
        '決定事項|結合テストは11/9から始める|||20261021_定例') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Kind, $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }), $_.Meeting }) 'Excel は行が飛んだところで表が終わる'
    Assert-Equal '次回まで' $r[4].Note
    Assert-Equal '2026/10/21' $r[3].MeetingDate.ToString('yyyy/MM/dd') 'ファイル名の日付'
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('宿題', '決定事項') $sheet.Sheets
    Assert-Equal 5 $sheet.Rows.Count '見出しと 4 件'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'議事録'!A2" }).Count -eq 1) 'Excel の議事録の該当セルへのリンク'
}

Test-Case 'Export-ActionItems: Markdown の表と、宿題の無い議事録' {
    $dir = New-TestDir
    @('# 定例 2026/10/23', '', '## 宿題', '', '| 内容 | 担当 | 期限 |', '|---|---|---|', '| 手順書を書く | 田中 | 10/30 |', '', '## メモ', '- 特になし') | Set-Content -LiteralPath (Join-Path $dir 'minutes.md') -Encoding UTF8
    $r = @(& $ActionItems (Join-Path $dir 'minutes.md') 6>$null)
    Assert-Equal @('宿題|手順書を書く|田中|2026/10/30') @($r | ForEach-Object { '{0}|{1}|{2}|{3:yyyy/MM/dd}' -f $_.Kind, $_.Title, $_.Owner, $_.Due })
    @('打ち合わせメモ', '・特に問題なし') | Set-Content -LiteralPath (Join-Path $dir 'memo.txt') -Encoding UTF8
    $text = Get-HostText { $null = & $ActionItems (Join-Path $dir 'memo.txt') -OutCsv (Join-Path $dir 'empty.csv') }
    Assert-Match $text '宿題 0 件'
    Assert-Equal 1 @(Get-Content -LiteralPath (Join-Path $dir 'empty.csv')).Count '見出しだけ'
}

Test-Case 'Export-ActionItems: Excel の議事録のセルの中の改行（Alt+Enter）を行の区切りとして読み、表のセルの改行は空白にする' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $xlsx = Join-Path $dir '20261020_定例.xlsx'
    $body = "■ 宿題`n・画面設計書を直す（山田、10/23まで）`n・テスト仕様書を書く（佐藤、10/30まで）`n■ 決定事項`n・受注日の初期値は当日とする"
    Export-Workbook -FilePath $xlsx -Sheets @(
        @{ Name = '議事録'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '内容'; Width = 60 }); Rows = @(
                @{ Values = @('議事内容', $body) },
                @{ Values = @('宿題', "・帳票の様式を確認する（鈴木、11/2まで）`n・IF 定義書を直す（田中、11/6まで）") }) },
        @{ Name = '宿題一覧'; Columns = @(@('内容', '担当', "対応`n期限") | ForEach-Object { @{ Header = $_; Width = 20 } }); Rows = @(@{ Values = @("DDL を直す`n（受注日の桁数）", '山田', '10/30') }) })
    $r = @(& $ActionItems $xlsx 6>$null)
    Assert-Equal @(
        '宿題|画面設計書を直す|山田|2026/10/23|議事録!B2',
        '宿題|テスト仕様書を書く|佐藤|2026/10/30|議事録!B2',
        '決定事項|受注日の初期値は当日とする|||議事録!B2',
        '宿題|帳票の様式を確認する|鈴木|2026/11/02|議事録!B3',
        '宿題|IF 定義書を直す|田中|2026/11/06|議事録!B3',
        '宿題|DDL を直す (受注日の桁数)|山田|2026/10/30|宿題一覧!A2') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}!{5}' -f $_.Kind, $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }), $_.Place.Sheet, $_.Place.Cell }) 'セルの中の 1 行ずつ（場所はそのセル）。表の見出し「対応↵期限」も期限の列'
    Assert-Equal '・画面設計書を直す（山田、10/23まで）' $r[0].Original
}

Test-Case 'Export-ActionItems: 行の中の担当・期限は、「説明日程」の「明日」を読まず、（山田、来週中）・和暦や - . 区切りの期限・2 人の担当・10月末・来週金曜も読む' {
    $dir = New-TestDir
    @('日時：2026年10月20日（火）', '■ 宿題',
        '・お客様への説明日程を調整する（山田さん）',
        '・証明日の扱いを確認する 担当：佐藤',
        '・帳票を直す（山田、来週中）',
        '・手順書を直す（佐藤、今週中）',
        '・説明会の資料を作る（佐藤、次回まで）',
        '・帳票定義書を直す 担当：山田 期限：2026-10-30',
        '・IF 定義書を直す 担当：佐藤 期限：令和8年11月6日',
        '・DDL を直す 担当：鈴木 期限：R8.11.13',
        '・旧帳票を削除する 担当：田中 期限：令和8年9月1日',
        '・テストデータを作る 担当：山田、佐藤 期限：10/30',
        '・環境を作る（山田さん、佐藤さん）',
        '・移行手順を書く 担当：鈴木 期限：10月末',
        '・IF を直す @田中 来週金曜まで') | Set-Content -LiteralPath (Join-Path $dir '議事録.txt') -Encoding UTF8
    $r = @(& $ActionItems (Join-Path $dir '議事録.txt') 6>$null)
    Assert-Equal @(
        'お客様への説明日程を調整する|山田||期限なし',
        '証明日の扱いを確認する|佐藤||期限なし',
        '帳票を直す|山田|2026/10/30|',
        '手順書を直す|佐藤|2026/10/23|',
        '説明会の資料を作る|佐藤||次回まで',
        '帳票定義書を直す|山田|2026/10/30|',
        'IF 定義書を直す|佐藤|2026/11/06|',
        'DDL を直す|鈴木|2026/11/13|',
        '旧帳票を削除する|田中|2026/09/01|',
        'テストデータを作る|山田、佐藤|2026/10/30|',
        '環境を作る|山田、佐藤||期限なし',
        '移行手順を書く|鈴木|2026/10/31|',
        'IF を直す|田中|2026/10/30|') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }), $_.Note }) '年のある期限は翌年にしない（令和8年9月1日は 2026/09/01）'
}

Test-Case 'Export-ActionItems: 日時・開催日の行の和暦（令和8年10月20日・R8.10.13）を会議の日にする（ファイルの更新日にしない）' {
    $dir = New-TestDir
    $a = Join-Path $dir 'teirei_a.txt'
    @('定例会議 議事録', '日時：令和8年10月20日（火）10:00～11:00', '■ 宿題', '・画面設計書を直す（山田、11/2まで）') | Set-Content -LiteralPath $a -Encoding UTF8
    $b = Join-Path $dir 'teirei_b.txt'
    @('定例会議 議事録', '開催日：R8.10.13', '■ 宿題', '・帳票を直す（佐藤、10/16まで）') | Set-Content -LiteralPath $b -Encoding UTF8
    foreach ($f in @($a, $b)) { (Get-Item -LiteralPath $f).LastWriteTime = [datetime]'2026-12-25 09:00' }
    $r = @(& $ActionItems @($a, $b) 6>$null)
    Assert-Equal @('画面設計書を直す|2026/11/02|2026/10/20', '帳票を直す|2026/10/16|2026/10/13') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Title, $_.Due.ToString('yyyy/MM/dd'), $_.MeetingDate.ToString('yyyy/MM/dd') }) '年の無い期限も会議の日の年'
}

Test-Case 'Export-ActionItems: 宿題の節の □・☐・▶・◇・▽ の行は箇条書きにし、◆・▼ の見出しと、節の見出しと同じ印の行で節を終える' {
    $dir = New-TestDir
    @('日時：2026年10月20日（火）', '■ 宿題',
        '□ 画面設計書を直す（山田、10/23まで）', '☐ テスト仕様書を書く（佐藤、10/30まで）', '▶ 帳票を直す（鈴木、10/30まで）', '◇ IF を直す（田中、10/30まで）', '▽ DDL を直す（田中、11/6まで）',
        '◆ その他', '・次回は10/27',
        '□宿題', '・手順書を書く（佐藤、11/6まで）', '□次回予定', '・10/27 10:00～ 会議室A') | Set-Content -LiteralPath (Join-Path $dir '議事録.txt') -Encoding UTF8
    $r = @(& $ActionItems (Join-Path $dir '議事録.txt') 6>$null)
    Assert-Equal @('画面設計書を直す|山田', 'テスト仕様書を書く|佐藤', '帳票を直す|鈴木', 'IF を直す|田中', 'DDL を直す|田中', '手順書を書く|佐藤') @($r | ForEach-Object { '{0}|{1}' -f $_.Title, $_.Owner }) '「□宿題」の節は「□次回予定」で終わる'
}

Test-Case 'Export-ActionItems: 宿題の次の担当・期限だけの行（「　担当：山田　期限：10/23」）は、前の宿題の担当・期限にする（前の宿題に無いものだけ）' {
    $dir = New-TestDir
    @('日時：2026年10月20日（火）', '■ 宿題',
        '・画面設計書の受注日の桁数を直す', '　担当：山田　期限：10/23',
        '・テスト仕様書を書く', '　（佐藤、10/30まで）',
        '・帳票を直す 担当：鈴木', '　期限：11/2',
        '・IF 定義書を直す 担当：田中', '　担当：山田　期限：11/6',
        '・説明会の資料を作る', '　※ 先方の都合を確認', '　担当：佐藤　次回まで') | Set-Content -LiteralPath (Join-Path $dir '議事録.txt') -Encoding UTF8
    $r = @(& $ActionItems (Join-Path $dir '議事録.txt') 6>$null)
    Assert-Equal @(
        '画面設計書の受注日の桁数を直す|山田|2026/10/23|',
        'テスト仕様書を書く|佐藤|2026/10/30|',
        '帳票を直す|鈴木|2026/11/02|',
        'IF 定義書を直す|田中|2026/11/06|',
        '説明会の資料を作る|佐藤||次回まで・先方の都合を確認') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }), $_.Note })
}

Test-Case 'Export-ActionItems: 宿題の表の状況の列を状態にし、前の議事録から写した同じ宿題は会議の日の新しいものだけを残す（週次報告で完了を期限切れにしない）' {
    $dir = New-TestDir
    $old = Join-Path $dir '20261013_定例.md'
    @('# 定例 2026/10/13', '', '## 宿題', '', '| 内容 | 担当 | 期限 | 状況 |', '|---|---|---|---|', '| 画面設計書を直す | 山田 | 10/16 | 対応中 |', '| 手順書を書く | 鈴木 | 10/16 | 未着手 |') | Set-Content -LiteralPath $old -Encoding UTF8
    $new = Join-Path $dir '20261020_定例.md'
    @('# 定例 2026/10/20', '', '## 宿題', '', '| 内容 | 担当 | 期限 | 状況 |', '|---|---|---|---|', '| 画面設計書を直す | 山田 | 10/16 | 完了 |', '| 帳票を直す | 佐藤 | 10/30 | 未着手 |') | Set-Content -LiteralPath $new -Encoding UTF8
    $csv = Join-Path $dir '宿題.csv'
    # 新しい議事録を先に読んでも、会議の日の新しいものを残す
    $null = & $ActionItems @($new, $old) -OutCsv $csv 6>$null
    $rows = @(Import-Csv -LiteralPath $csv -Encoding UTF8)
    Assert-Equal @(
        '1|画面設計書を直す|山田|2026/10/16|完了|2026/10/13|20261020_定例',
        '2|帳票を直す|佐藤|2026/10/30|未着手|2026/10/20|20261020_定例',
        '3|手順書を書く|鈴木|2026/10/16|未着手|2026/10/13|20261013_定例') @($rows | ForEach-Object { $row = $_; (@('No', '件名', '担当', '期限', '状態', '起票日', '会議') | ForEach-Object { $row.$_ }) -join '|' }) '起票日は最初に出てきた会議の日'
    $weekly = Get-HostText { $null = & (Join-Path $script:ToolsDir 'New-WeeklyReport.ps1') -IssueList $csv -AsOf '2026-10-21' }
    Assert-Match $weekly '・未完了: 2 件・うち期限切れ 1 件'
    Assert-NotMatch $weekly '【山田】'
}

Test-Case 'Export-ActionItems: Word の内容コントロール（w:sdt）の中の段落・表・セルも読む' {
    $dir = New-TestDir
    $docx = Join-Path $dir '20261020_定例.docx'
    New-MinutesDocx $docx @('■ 宿題', '・画面設計書を直す（山田、10/23まで）') @(@('No', 'アクション', '担当', '期限'), @('1', 'IF 定義書を更新する', '佐藤', '10/28')) @('【決定】結合テストは11/9から始める')
    Edit-ZipText $docx 'word/document.xml' {
        param($t)
        $cell = '<w:tc><w:p><w:r><w:t xml:space="preserve">佐藤</w:t></w:r></w:p></w:tc>'
        $t.Replace('<w:body>', '<w:body><w:sdt><w:sdtPr><w:alias w:val="議事内容"/></w:sdtPr><w:sdtContent>').Replace('</w:body>', '</w:sdtContent></w:sdt></w:body>').Replace($cell, '<w:sdt><w:sdtContent>' + $cell + '</w:sdtContent></w:sdt>')
    }
    $r = @(& $ActionItems $docx 6>$null)
    Assert-Equal @('宿題|画面設計書を直す|山田|2026/10/23', '宿題|IF 定義書を更新する|佐藤|2026/10/28', '決定事項|結合テストは11/9から始める||') @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Kind, $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }) }) '本文の内容コントロールの中の段落・表と、セルの内容コントロールの中のセル'
}

Test-Case 'Export-ActionItems: -OutXlsx の期限・起票日と決定事項の日付は日付のセルにする（CSV は yyyy/MM/dd の文字のまま）' {
    $dir = New-TestDir
    @('日時：2026年10月20日（火）', '■ 宿題', '・画面設計書を直す（山田、10/23まで）', '・性能の目標値を確認する', '■ 決定事項', '・受注日の初期値は当日とする') | Set-Content -LiteralPath (Join-Path $dir '議事録.txt') -Encoding UTF8
    $out = Join-Path $dir '宿題.xlsx'
    $null = & $ActionItems (Join-Path $dir '議事録.txt') -OutXlsx $out -OutCsv (Join-Path $dir '宿題.csv') 6>$null
    $serial = { param($Day) [string](([datetime]$Day) - [datetime]'1899-12-30').TotalDays }
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal ((@('2026-10-23', '2026-10-20') | ForEach-Object { & $serial $_ }) -join '|') ((@('E', 'G') | ForEach-Object { $sheet.Rows[2][$_] }) -join '|') '期限・起票日は日付のセル（シリアル値）'
    Assert-True (-not $sheet.Rows[3].ContainsKey('E')) '期限の無い宿題は空のセル'
    Assert-Equal (& $serial '2026-10-20') (Read-XlsxSheet -Path $out -Index 2).Rows[2]['C'] '決定事項の日付も日付のセル'
    $csv = @(Import-Csv -LiteralPath (Join-Path $dir '宿題.csv') -Encoding UTF8)
    Assert-Equal '2026/10/23|2026/10/20' ('{0}|{1}' -f $csv[0].'期限', $csv[0].'起票日')
}

Test-Case 'Export-ActionItems: 宿題の表で縦に結合した担当・期限は結合した各行の値として読み、結合したセルに書いた本文はくり返さない' {
    $dir = New-TestDir
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $xlsx = Join-Path $dir '20261020_定例.xlsx'
    Export-Workbook -FilePath $xlsx -Sheets @(
        @{ Name = '宿題'; Columns = @(@('No', '内容', '担当', '期限') | ForEach-Object { @{ Header = $_; Width = 20 } }); Rows = @(
                @{ Values = @('1', '画面設計書を直す', '山田', '10/30') }, @{ Values = @('2', 'テーブル定義書を直す', '', '') }) },
        @{ Name = '議事録'; Columns = @(@{ Header = '定例会議 議事録'; Width = 60 }); Rows = @(@{ Values = @("■ 宿題`n・IF 定義書を更新する（佐藤、10/28まで）") }) })
    # 宿題の表の担当・期限を 2 行ずつ結合し、議事録のシートの本文のセルを 5 行に結合する
    Edit-ZipText $xlsx 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('<pageMargins', '<mergeCells count="2"><mergeCell ref="C2:C3"/><mergeCell ref="D2:D3"/></mergeCells><pageMargins') }
    Edit-ZipText $xlsx 'xl/worksheets/sheet2.xml' { param($t) $t.Replace('<pageMargins', '<mergeCells count="1"><mergeCell ref="A2:A6"/></mergeCells><pageMargins') }
    $r = @(& $ActionItems $xlsx 6>$null)
    Assert-Equal @('画面設計書を直す|山田|2026/10/30', 'テーブル定義書を直す|山田|2026/10/30', 'IF 定義書を更新する|佐藤|2026/10/28') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Title, $_.Owner, $(if ($_.Due) { $_.Due.ToString('yyyy/MM/dd') } else { '' }) })
}

Test-Case 'Export-ActionItems: Get-Help でヘルプが表示される' {
    $h = Get-Help $ActionItems
    Assert-Match $h.Synopsis '議事録'
    Assert-Equal 5 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
