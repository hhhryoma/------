#     Search-Excel.ps1 のテスト。フィクスチャは tests\fixtures\grep（make_fixtures.py で生成）

$SearchExcel = Join-Path $script:ToolsDir 'Search-Excel.ps1'
$grep = Join-Path $script:Fixtures 'grep'

function Invoke-Search {
    # 画面出力と警告を抑えて結果だけ受け取る。警告は $script:LastWarnings に入れる
    param([Parameter(ValueFromRemainingArguments = $true)] $Rest)

    # 「-名前 値」「-スイッチ」の並びをハッシュテーブルにして渡す（配列のままだと名前が値として渡る）
    $named = @{}
    for ($i = 0; $i -lt $Rest.Count; $i++) {
        $name = ([string]$Rest[$i]).TrimStart('-')
        if ($i + 1 -lt $Rest.Count -and -not ([string]$Rest[$i + 1]).StartsWith('-')) { $named[$name] = $Rest[$i + 1]; $i++ }
        else { $named[$name] = $true }
    }
    $w = $null
    $hits = @(& $SearchExcel @named -WarningVariable w -WarningAction SilentlyContinue 6>$null)
    $script:LastWarnings = @($w | ForEach-Object { $_.Message })
    return , $hits
}

function Get-HitKeys {
    param($Hits)
    return @($Hits | ForEach-Object { '{0}|{1}|{2}' -f (ConvertTo-Slash $_.File), $_.Location, $_.Kind } | Sort-Object)
}

Test-Case 'Search-Excel: フォルダを再帰検索し、値・図形・非表示シートがヒットする' {
    $hits = Invoke-Search -Pattern 'TODO' -Path $grep -NoExcel
    $expected = @(
        'sample.xlsx|Sheet3!A2000|値',
        'sample.xlsx|作業用!A1|値',
        'sample.xlsx|顧客一覧!B10|図形',
        'sample.xlsx|顧客一覧!B4|値',
        'sample.xlsx|顧客一覧!G3|図形',
        'sub/disguised.xls|Sheet3!A2000|値',
        'sub/disguised.xls|作業用!A1|値',
        'sub/disguised.xls|顧客一覧!B10|図形',
        'sub/disguised.xls|顧客一覧!B4|値',
        'sub/disguised.xls|顧客一覧!G3|図形',
        'sub/inline.xlsx|inline!A1|値',
        'sub/inline.xlsx|inline!B2|値',
        "sub/quote.xlsx|O'Brien 表 1!C5|値"
    ) | Sort-Object
    Assert-Equal $expected (Get-HitKeys $hits)
}

Test-Case 'Search-Excel: $global:ExcelToolsSearchCache があれば、同じファイル・同じ条件の検索は前の結果を使う（更新したファイルは読み直す）' {
    $dir = Copy-Fixture 'grep'
    $global:ExcelToolsSearchCache = @{ '__Reused' = 0 }
    try {
        $first = Invoke-Search -Pattern 'TODO' -Path $dir -NoExcel
        Assert-Equal 0 $global:ExcelToolsSearchCache['__Reused']
        $first[0].Text = '書き換えた'   # 呼び出し側が結果を書き換えても、次の呼び出しには響かない
        # 別のフォルダから検索しても、ファイルの名前（File）はその検索の場所からのものになる
        $sub = Invoke-Search -Pattern 'TODO' -Path (Join-Path $dir 'sub') -NoExcel
        Assert-True ($global:ExcelToolsSearchCache['__Reused'] -ge 3) ([string]$global:ExcelToolsSearchCache['__Reused'])
        Assert-Equal @($first | Where-Object { $_.File -like 'sub*' } | ForEach-Object { '{0}|{1}' -f (ConvertTo-Slash $_.File).Substring(4), $_.Location } | Sort-Object) @($sub | ForEach-Object { '{0}|{1}' -f (ConvertTo-Slash $_.File), $_.Location } | Sort-Object)
        $again = Invoke-Search -Pattern 'TODO' -Path $dir -NoExcel
        Assert-Equal (Get-HitKeys $first) (Get-HitKeys $again)
        Assert-True (@($again | Where-Object { $_.Text -eq '書き換えた' }).Count -eq 0) '前の呼び出しの書き換えが残っている'
        # 条件が違えば読み直す。更新したファイルも読み直す
        $before = $global:ExcelToolsSearchCache['__Reused']
        $null = Invoke-Search -Pattern 'TODO' -Path (Join-Path $dir 'sample.xlsx') -NoExcel -CaseSensitive
        (Get-Item -LiteralPath (Join-Path $dir 'sample.xlsx')).LastWriteTime = (Get-Date).AddMinutes(-5)
        $null = Invoke-Search -Pattern 'TODO' -Path (Join-Path $dir 'sample.xlsx') -NoExcel
        Assert-Equal $before $global:ExcelToolsSearchCache['__Reused']
    }
    finally { Remove-Variable -Name ExcelToolsSearchCache -Scope Global -ErrorAction SilentlyContinue }
}

Test-Case 'Search-Excel: 暗号化・破損ファイルは警告して処理を続ける' {
    $null = Invoke-Search -Pattern 'TODO' -Path $grep -NoExcel
    $w = $script:LastWarnings -join "`n"
    Assert-Match $w 'encrypted\.xlsx.*パスワード'
    Assert-Match $w 'broken\.xlsx.*壊れている'
    Assert-Equal 2 $script:LastWarnings.Count '警告の数'
}

Test-Case 'Search-Excel: Excel の一時ファイル（~$）は検索しない' {
    $hits = Invoke-Search -Pattern 'TODO' -Path $grep -NoExcel
    Assert-True (-not @($hits | Where-Object { $_.File -like '*~$*' }).Count) '一時ファイルがヒットしています'
}

Test-Case 'Search-Excel: 本物の .xls は Excel で読む（Excel が無ければ理由を表示）' {
    $hits = Invoke-Search -Pattern '旧形式TODO' -Path (Join-Path $grep 'sub/legacy.xls')
    if ($hits.Count -eq 1) {
        Assert-Equal '値' $hits[0].Kind   # Excel がある環境
    }
    else {
        Assert-Match ($script:LastWarnings -join "`n") 'legacy\.xls.*Excel を起動できない'
    }
}

Test-Case 'Search-Excel: -NoExcel では .xls をスキップし、失敗にしない' {
    $null = Invoke-Search -Pattern 'x' -Path (Join-Path $grep 'sub/legacy.xls') -NoExcel
    Assert-Equal 0 $script:LastWarnings.Count
}

Test-Case 'Search-Excel: 値は表示に近い形（日付・時刻・数値・論理値）で検索される' {
    $hits = Invoke-Search -Pattern '.' -Path (Join-Path $grep 'sample.xlsx') -Target Value -SheetNamePattern '^顧客一覧$'
    $byCell = @{}
    foreach ($h in $hits) { $byCell[$h.Cell] = $h.Text }
    Assert-Equal '2026/10/16' $byCell['C2'] '日付'
    Assert-Equal '13:00:00' $byCell['D2'] '時刻'
    Assert-Equal '2026/01/02 09:30:00' $byCell['C3'] '和暦表示の日時'
    Assert-Equal '0.3' $byCell['E2'] '0.1+0.2'
    Assert-Equal '1.3' $byCell['E3'] '数式の結果'
    Assert-Equal '12345678901' $byCell['F1'] '大きな整数'
    Assert-Equal 'TRUE' $byCell['B5'] '論理値'
    Assert-Equal '複数行↵の文字列 TODO' $byCell['B4'] 'セル内改行'
}

Test-Case 'Search-Excel: ふりがなは検索対象にしない' {
    Assert-Equal 0 (Invoke-Search -Pattern 'カイシャ' -Path (Join-Path $grep 'sample.xlsx')).Count
    $hits = Invoke-Search -Pattern '^会社名$' -Path (Join-Path $grep 'sample.xlsx') -Target Value
    Assert-Equal '顧客一覧!B1' $hits[0].Location
}

Test-Case 'Search-Excel: -SimpleMatch で記号を含む文字列をそのまま探せる' {
    $hits = Invoke-Search -Pattern '(株)' -Path (Join-Path $grep 'sample.xlsx') -SimpleMatch -Target Value
    Assert-Equal @('顧客一覧!B2', '顧客一覧!E5') @($hits | ForEach-Object { $_.Location } | Sort-Object)
}

Test-Case 'Search-Excel: 結果は File・Location・Kind・Text を表示し、写し（PSObject.Copy）は元と別のものになる' {
    $hits = Invoke-Search -Pattern '(株)' -Path (Join-Path $grep 'sample.xlsx') -SimpleMatch -Target Value
    $shown = ($hits[0] | Format-Table | Out-String)
    Assert-Match $shown 'File\s+Location\s+Kind\s+Text'
    Assert-True (-not $shown.Contains('Path')) "表示する列: $shown"
    Assert-Equal 'sample.xlsx|顧客一覧|B2|値' ('{0}|{1}|{2}|{3}' -f $hits[0].File, $hits[0].Sheet, $hits[0].Cell, $hits[0].Kind)
    $copy = $hits[0].PSObject.Copy()
    $copy.Text = '書き換え'
    Assert-True ($hits[0].Text -ne '書き換え') '読み取り結果の使い回しは写しを渡す'
}

Test-Case 'Search-Excel: -IgnoreWidth で全角半角を区別しない' {
    Assert-Equal 0 (Invoke-Search -Pattern 'abc' -Path (Join-Path $grep 'sample.xlsx') -SimpleMatch).Count
    $hits = Invoke-Search -Pattern 'abc' -Path (Join-Path $grep 'sample.xlsx') -SimpleMatch -IgnoreWidth
    Assert-Equal '顧客一覧!B3' $hits[0].Location
}

Test-Case 'Search-Excel: 既定は大文字小文字を区別せず、-CaseSensitive で区別する' {
    Assert-True ((Invoke-Search -Pattern 'todo' -Path (Join-Path $grep 'sample.xlsx')).Count -gt 0)
    Assert-Equal 0 (Invoke-Search -Pattern 'todo' -Path (Join-Path $grep 'sample.xlsx') -CaseSensitive).Count
}

Test-Case 'Search-Excel: 数式を検索できる' {
    $hits = Invoke-Search -Pattern 'VLOOKUP' -Path (Join-Path $grep 'sample.xlsx') -Target Formula
    Assert-Equal 1 $hits.Count
    Assert-Equal '数式' $hits[0].Kind
    Assert-Match $hits[0].Text '^=VLOOKUP\(A2,'
}

Test-Case 'Search-Excel: 共有数式（コピーした数式）は、写しのセルも位置をずらした数式で検索される' {
    $hits = Invoke-Search -Pattern '.' -Path (Join-Path $script:Fixtures 'formula/明細.xlsx') -Target Formula
    $byCell = @{}
    foreach ($h in $hits) { $byCell[$h.Cell] = $h.Text }
    Assert-Equal 19 $hits.Count
    Assert-Equal '=ROUND(A4*B4*(1+税率),0)' $byCell['C4'] '縦にコピー'
    Assert-Equal '=SUM($C$2:C5)' $byCell['D5'] '$ の付いた参照はずらさない'
    Assert-Equal '=IF(A6>3,"A2以上",''集計 A1''!$B$1)' $byCell['E6'] '文字列やシート名の中の A2・A1 はずらさない'
    Assert-Equal '=COUNTA(C:C)' $byCell['C8'] '横にコピーした列の参照'
    Assert-Equal '=SUM(3:3)' $byCell['A11'] '行の参照'
    $hits = Invoke-Search -Pattern 'A5\*B5' -Path (Join-Path $script:Fixtures 'formula/明細.xlsx')
    Assert-Equal '明細!C5|数式' ('{0}|{1}' -f $hits[0].Location, $hits[0].Kind) '写しのセルだけにある文字も見つかる'
}

Test-Case 'Search-Excel: ヘッダー・フッター、名前の定義、入力規則を検索できる' {
    $file = Join-Path $script:Fixtures 'formula/明細.xlsx'
    $hits = Invoke-Search -Pattern 'システム|ページ|A&B' -Path $file
    Assert-Equal @('明細|ヘッダー・フッター|ヘッダー（左）|受注管理システム', '明細|ヘッダー・フッター|フッター（中央）|&[ページ番号] / &[総ページ数]', '明細|ヘッダー・フッター|フッター（右）|株式会社A&B') `
        @($hits | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Location, $_.Kind, $_.Name, $_.Text }) '書式（フォント）で分かれた文字もつなげて探す'
    $hits = Invoke-Search -Pattern '明細|税率' -Path $file -Target DefinedName
    Assert-Equal @('税率|=明細!$J$1', '明細!対象|=明細!$A$2:$A$6', '明細!Print_Area|=明細!$A$1:$G$11') @($hits | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Text }) 'Excel が内部で使う非表示の名前は除く'
    $hits = Invoke-Search -Pattern '対応中|集計|リストから' -Path $file -Target Validation
    Assert-Equal @('明細!F2|リスト（F2:F6）|未着手,対応中,完了', '明細!F2|入力時メッセージ（F2:F6）|未着手・対応中・完了から選んでください', '明細!F2|エラーメッセージ（F2:F6）|リストから選んでください', "明細!G2|リスト（G2:G6）|='集計 A1'!`$A`$1:`$A`$3") `
        @($hits | ForEach-Object { '{0}|{1}|{2}' -f $_.Location, $_.Name, $_.Text })
    $out = Join-Path (New-TestDir) 'hits.xlsx'
    $null = Invoke-Search -Pattern '税率' -Path $file -Target DefinedName -OutXlsx $out
    Assert-Equal 0 @((Read-XlsxSheet -Path $out).Links | Where-Object { $_.Location }).Count 'ブック全体の名前はファイルを開くだけのリンク'
}

Test-Case 'Search-Excel: コメントとシート名を検索できる' {
    $c = Invoke-Search -Pattern '住所' -Path (Join-Path $grep 'sample.xlsx') -Target Comment
    Assert-Equal 'コメント' $c[0].Kind
    Assert-Equal 'A2' $c[0].Cell
    $s = Invoke-Search -Pattern '作業' -Path (Join-Path $grep 'sample.xlsx') -Target SheetName
    Assert-Equal '作業用' $s[0].Location
}

Test-Case 'Search-Excel: 図形は互換用の重複（mc:Fallback）を数えない' {
    $hits = Invoke-Search -Pattern 'TODO実行' -Path (Join-Path $grep 'sample.xlsx') -Target Shape
    Assert-Equal 1 $hits.Count
    Assert-Equal 'Button 1' $hits[0].Name
    Assert-Equal 'B10' $hits[0].Cell
}

Test-Case 'Search-Excel: -SkipHiddenSheets と -SheetNameExclude でシートを除外できる' {
    $hits = Invoke-Search -Pattern 'TODO' -Path (Join-Path $grep 'sample.xlsx') -SkipHiddenSheets -SheetNameExclude 'Sheet3'
    Assert-Equal @('顧客一覧!B10', '顧客一覧!B4', '顧客一覧!G3') @($hits | ForEach-Object { $_.Location } | Sort-Object)
}

Test-Case 'Search-Excel: -Path にワイルドカードを使える。[ ] を含む名前は、その名前のファイルがあればそのまま探す' {
    $hits = Invoke-Search -Pattern 'TODO' -Path (Join-Path $grep 'sub/*.xlsx')
    Assert-Equal 3 $hits.Count

    $book = Join-Path (New-TestDir) '設計書[旧].xlsx'
    Copy-Item -LiteralPath (Join-Path $grep 'sample.xlsx') -Destination $book
    $hits = Invoke-Search -Pattern 'TODO' -Path $book -NoExcel
    Assert-Equal 5 $hits.Count ($script:LastWarnings -join ' / ')
    $escaped = [System.Management.Automation.WildcardPattern]::Escape($book)
    Assert-Equal 5 (Invoke-Search -Pattern 'TODO' -Path $escaped -NoExcel).Count 'ワイルドカードの記号を逃がしたパス（画面から）でも探す'
}

Test-Case 'Search-Excel: -OutCsv は BOM 付き UTF-8 で日本語の列名' {
    $csv = Join-Path (New-TestDir) 'out.csv'
    $hits = Invoke-Search -Pattern 'TODO' -Path $grep -NoExcel -OutCsv $csv
    $bytes = [System.IO.File]::ReadAllBytes($csv)
    Assert-Equal '239 187 191' (($bytes[0..2] | ForEach-Object { [int]$_ }) -join ' ') 'BOM'
    $rows = @(Import-Csv -LiteralPath $csv -Encoding UTF8)
    Assert-Equal $hits.Count $rows.Count '行数'
    Assert-Equal 'ファイル,シート,セル,種類,内容,名前,フルパス' (@($rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join ',')
}

Test-Case 'Search-Excel: -OutXlsx は該当セルへのリンク付きで、出力自体は次回の検索対象外' {
    $dir = Copy-Fixture 'grep'
    $xlsx = Join-Path $dir 'result.xlsx'
    $hits = Invoke-Search -Pattern 'TODO' -Path $dir -NoExcel -OutXlsx $xlsx
    $sheet = Read-XlsxSheet -Path $xlsx
    Assert-Equal $hits.Count $sheet.Links.Count 'リンクの数'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'O''Brien 表 1'!C5" }).Count -eq 1) "' を含むシート名のリンク"
    Assert-True (@($sheet.Links | Where-Object { $_.Target -like 'file:///*sample.xlsx' }).Count -gt 0) 'リンク先'
    Assert-Equal 'ファイル' $sheet.Rows[1]['A'] '見出し'
    # 制御文字は取り除かれて、XML として読める
    Assert-True (@($sheet.Rows.Values | Where-Object { $_['F'] -like 'TODO <&> "quote"*ctrl' }).Count -eq 1) '特殊文字のセル'

    $again = Invoke-Search -Pattern 'TODO' -Path $dir -NoExcel -OutXlsx $xlsx
    Assert-Equal $hits.Count $again.Count '再実行で出力ファイルがヒットしない'
}

Test-Case 'Search-Excel: -OutXlsx の拡張子が .xlsx 以外ならエラー' {
    Assert-Throws { & $SearchExcel -Pattern 'x' -Path $grep -OutXlsx (Join-Path (New-TestDir) 'r.csv') 6>$null } '\.xlsx'
}

Test-Case 'Search-Excel: Out-GridView が無い環境で -Open はエラー' -NonWindowsOnly {
    Assert-Throws { & $SearchExcel -Pattern 'x' -Path $grep -Open 6>$null } 'Out-GridView'
}

Test-Case 'Search-Excel: 不正な正規表現は分かりやすいエラー' {
    Assert-Throws { & $SearchExcel -Pattern '(abc' -Path $grep 6>$null } 'SimpleMatch'
}

function New-SearchDocx {
    #     段落・タブ・削除した変更履歴・表・テキストボックス（互換用の写し付き）・内容コントロール・コメント・ヘッダー・脚注のある Word 文書
    param([string] $Path)
    $ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml"'
    $p = { param($t) '<w:p><w:r><w:t>' + $t + '</w:t></w:r></w:p>' }
    $box = '<w:txbxContent><w:p><w:r><w:t>受注番号は自動採番</w:t></w:r></w:p></w:txbxContent>'
    $row = { param($cells) '<w:tr>' + ((@($cells | ForEach-Object { '<w:tc>' + (& $p $_) + '</w:tc>' })) -join '') + '</w:tr>' }
    $body = '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>受注番号の採番ルール</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' +
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="840"/></w:tabs></w:pPr><w:r><w:t>受注日</w:t></w:r><w:r><w:tab/><w:t>必須</w:t></w:r></w:p>' +
        '<w:p><w:r><w:t>残す文</w:t></w:r><w:del w:id="5" w:author="山田"><w:r><w:delText>受注番号（旧）</w:delText></w:r></w:del></w:p>' +
        '<w:tbl>' + (& $row @('項目', '桁数')) + (& $row @('受注番号', '10')) + (& $row @('受注日', '8')) + '</w:tbl>' +
        '<w:p><w:r><w:t>図の説明</w:t></w:r><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wps:wsp><wps:txbx>' + $box + '</wps:txbx></wps:wsp></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox>' + $box + '</v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>' +
        '<w:sdt><w:sdtContent>' + (& $p '受注番号の桁数は10桁') + '</w:sdtContent></w:sdt>' + (& $p '以上')
    $parts = [ordered]@{
        'word/document.xml' = '<?xml version="1.0" encoding="UTF-8"?><w:document ' + $ns + '><w:body>' + $body + '</w:body></w:document>'
        'word/header1.xml' = '<?xml version="1.0" encoding="UTF-8"?><w:hdr ' + $ns + '>' + (& $p '受注システム 基本設計書') + '</w:hdr>'
        'word/footnotes.xml' = '<?xml version="1.0" encoding="UTF-8"?><w:footnotes ' + $ns + '><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1">' + (& $p '受注番号は年度で切り替える') + '</w:footnote></w:footnotes>'
        'word/comments.xml' = '<?xml version="1.0" encoding="UTF-8"?><w:comments ' + $ns + '><w:comment w:id="0" w:author="佐藤">' + (& $p '受注番号の桁数を確認') + '</w:comment></w:comments>'
    }
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

Test-Case 'Search-Excel: -IncludeWord で Word の段落・表の行・テキストボックス・内容コントロール・脚注・コメントを探す（削除した変更履歴は探さない）' {
    $dir = New-TestDir
    New-SearchDocx (Join-Path $dir '基本設計書.docx')
    $hits = Invoke-Search -Pattern '受注番号' -Path $dir -IncludeWord
    Assert-Equal @(
        '本文!段落 1|段落|受注番号の採番ルール|',
        '本文!表 1 の 2 行目|表|受注番号 | 10|',
        '本文!段落 5|段落|受注番号の桁数は10桁|',
        '本文!テキストボックス 1|図形|受注番号は自動採番|テキストボックス 1',
        '脚注!段落 1|段落|受注番号は年度で切り替える|',
        '本文!段落 1|コメント|受注番号の桁数を確認|佐藤') @($hits | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Location, $_.Kind, $_.Text, $_.Name }) 'テキストボックスは互換用の写しを数えない。コメントは付いている段落'
    Assert-Equal "受注日`t必須" (Invoke-Search -Pattern '受注日' -Path $dir -IncludeWord -Target Value)[0].Text 'タブ位置の設定はタブにしない'
    Assert-Equal 0 (Invoke-Search -Pattern '旧' -Path $dir -IncludeWord).Count '削除した文字'
    Assert-Equal @('ヘッダー!段落 1|ヘッダー・フッター') @(Invoke-Search -Pattern 'システム' -Path $dir -IncludeWord -Target HeaderFooter | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Kind })
    Assert-Equal 0 (Invoke-Search -Pattern '受注番号' -Path $dir).Count '-IncludeWord が無ければ Word は探さない（他のツールから呼ぶとき）'

    # 結果の Excel のリンクは文書を開くだけ（Excel のセルの場所を付けない）
    $out = Join-Path $dir '結果.xlsx'
    $null = Invoke-Search -Pattern '受注番号' -Path (Join-Path $dir '基本設計書.docx') -IncludeWord -OutXlsx $out
    $links = @((Read-XlsxSheet -Path $out).Links)
    Assert-Equal 6 $links.Count
    Assert-True (@($links | Where-Object { $_.Location -or $_.Target -notlike '*基本設計書.docx' }).Count -eq 0) ($links | Out-String)
}

Test-Case 'Search-Excel: 経過時間の表示形式（[h]:mm・[h]:mm:ss など）のセルは、Excel の表示と同じ 7:30・31:15 の形で検索される（時刻の書式は 07:30:00 のまま）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $path = Join-Path (New-TestDir) '工数.xlsx'
    $rows = @(@('[h]:mm', 0.3125), @('[h]:mm', (31.25 / 24)), @('[h]:mm:ss（組み込み）', (31.25 / 24)), @('[mm]:ss', 0.3125), @('[h]"時間"mm"分"', 0.3125), @('h:mm（時刻）', 0.3125), @('数値', 0.3125))
    Export-Workbook -FilePath $path -Sheets @(@{ Name = '工数'; Columns = @(@{ Header = '書式'; Width = 16 }, @{ Header = '値'; Width = 8 }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    # B2・B3 は [h]:mm、B4 は組み込みの [h]:mm:ss（46）、B5 は [mm]:ss、B6 は [h]"時間"mm"分"、B7 は時刻の h:mm（組み込みの 20）にする
    Edit-ZipText $path 'xl/styles.xml' {
        param($x)
        $x = $x.Replace('</numFmts>', '<numFmt numFmtId="177" formatCode="[h]:mm"/><numFmt numFmtId="178" formatCode="[mm]:ss"/><numFmt numFmtId="179" formatCode="[h]&quot;時間&quot;mm&quot;分&quot;"/></numFmts>')
        $x.Replace('<cellXfs count="5">', '<cellXfs count="10">').Replace('</cellXfs>', ((@(177, 46, 178, 179, 20) | ForEach-Object { '<xf numFmtId="{0}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' -f $_ }) -join '') + '</cellXfs>')
    }
    Edit-ZipText $path 'xl/worksheets/sheet1.xml' {
        param($x)
        foreach ($p in @(@('B2', 5), @('B3', 5), @('B4', 6), @('B5', 7), @('B6', 8), @('B7', 9))) { $x = $x.Replace(('<c r="{0}" s="2">' -f $p[0]), ('<c r="{0}" s="{1}">' -f $p[0], $p[1])) }
        $x
    }
    $hits = Invoke-Search -Pattern '\S' -Path $path -Target Value
    Assert-Equal @('7:30', '31:15', '31:15:00', '450:00', '7時間30分', '07:30:00', '0.3125') @(2..8 | ForEach-Object { $c = 'B' + $_; @($hits | Where-Object { $_.Cell -eq $c })[0].Text }) '経過時間は 24 時間を超えても時間で数える。時刻・数値の書式はこれまでどおり'
}

Test-Case 'Search-Excel: -FillMerged は、結合したセルの値を範囲の最初の列の下の行にも返す（右の列には返さない。付けなければ左上のセルだけ）' {
    $path = Join-Path (New-TestDir) '結合.xlsx'
    New-SheetsBook $path ([ordered]@{ 'ケース' = @(@('No', '手順', '期待結果', '備考', ''), @('1', '検索する', '一覧に表示される', '横に結合', ''), @('2', '', '', '', ''), @('3', '', '', '縦横に結合', '')) })
    # B2:B4・C2:C3 は縦に、D2:E2 は横に、D4:E5 は縦横に結合する
    Edit-ZipText $path 'xl/worksheets/sheet1.xml' { param($t) $t.Replace('<pageMargins', '<mergeCells count="4"><mergeCell ref="B2:B4"/><mergeCell ref="C2:C3"/><mergeCell ref="D2:E2"/><mergeCell ref="D4:E5"/></mergeCells><pageMargins') }
    $cells = { param($hits) @($hits | ForEach-Object { '{0}|{1}' -f $_.Cell, $_.Text }) }
    $global:ExcelToolsSearchCache = @{ '__Reused' = 0 }
    try {
        $plain = Invoke-Search -Pattern '\S' -Path $path -Target Value
        $filled = Invoke-Search -Pattern '\S' -Path $path -Target Value -FillMerged
        Assert-Equal 0 $global:ExcelToolsSearchCache['__Reused'] '-FillMerged の有無は、読み取り結果の使い回しでは別の条件'
    }
    finally { Remove-Variable -Name ExcelToolsSearchCache -Scope Global -ErrorAction SilentlyContinue }
    $topLeft = @('A1|No', 'B1|手順', 'C1|期待結果', 'D1|備考', 'A2|1', 'B2|検索する', 'C2|一覧に表示される', 'D2|横に結合', 'A3|2', 'A4|3', 'D4|縦横に結合')
    Assert-Equal $topLeft (& $cells $plain) '付けなければ、これまでどおり左上のセルだけ'
    Assert-Equal ($topLeft + @('B3|検索する', 'B4|検索する', 'C3|一覧に表示される', 'D5|縦横に結合')) (& $cells $filled) '右の列（E2・E4・E5）には返さない'
}

Test-Case 'Search-Excel: 図形の段落の中の改行（<a:br>。Shift+Enter の改行）も、段落の区切りと同じく ↵ にする（改行をまたいでは一致しない）' {
    $path = Join-Path (New-TestDir) '図形.xlsx'
    New-SheetsBook $path ([ordered]@{ 'S' = @(@('項目'), @('x')) })
    Add-TestDrawing $path @(@{ Kind = 'sp'; Name = '図1'; Text = 'X'; From = @(1, 1); To = @(4, 4) })
    Edit-ZipText $path 'xl/drawings/drawing1.xml' { param($x) $x.Replace('<a:p><a:r><a:t>X</a:t></a:r></a:p>', '<a:p><a:r><a:t>受注</a:t></a:r><a:br/><a:r><a:t>管理システム</a:t></a:r></a:p><a:p><a:r><a:t>概要</a:t></a:r></a:p>') }
    $hits = Invoke-Search -Pattern '管理' -Path $path -Target Shape
    Assert-Equal @('S!B2|受注↵管理システム↵概要|図1') @($hits | ForEach-Object { '{0}|{1}|{2}' -f $_.Location, $_.Text, $_.Name })
    Assert-Equal 0 (Invoke-Search -Pattern '受注管理' -Path $path -Target Shape).Count '改行をまたいでは一致しない'
}

Test-Case 'Search-Excel: 曜日（aaa・aaaa）・元号（g）・和暦の年（e）だけを表示する表示形式のセルも、日付（yyyy/MM/dd）で検索される（G/標準 の表示形式は数値のまま）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $path = Join-Path (New-TestDir) 'WBS.xlsx'
    $rows = @(@('aaa', 46296), @('aaaa', 46297), @('[$-411]ggge"年"', 46298), @('[$-411]e"年"', 46299), @('G/標準"件"', 46300))
    Export-Workbook -FilePath $path -Sheets @(@{ Name = 'WBS'; Columns = @(@{ Header = '書式'; Width = 16 }, @{ Header = '値'; Width = 8 }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    # B2〜B6 に、上の表示形式（177〜181）を順に付ける
    Edit-ZipText $path 'xl/styles.xml' {
        param($x)
        $x = $x.Replace('</numFmts>', '<numFmt numFmtId="177" formatCode="aaa"/><numFmt numFmtId="178" formatCode="aaaa"/><numFmt numFmtId="179" formatCode="[$-411]ggge&quot;年&quot;"/><numFmt numFmtId="180" formatCode="[$-411]e&quot;年&quot;"/><numFmt numFmtId="181" formatCode="G/標準&quot;件&quot;"/></numFmts>')
        $x.Replace('<cellXfs count="5">', '<cellXfs count="10">').Replace('</cellXfs>', ((@(177..181) | ForEach-Object { '<xf numFmtId="{0}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' -f $_ }) -join '') + '</cellXfs>')
    }
    Edit-ZipText $path 'xl/worksheets/sheet1.xml' {
        param($x)
        foreach ($row in 2..6) { $x = $x.Replace(('<c r="B{0}" s="2">' -f $row), ('<c r="B{0}" s="{1}">' -f $row, ($row + 3))) }
        $x
    }
    $hits = Invoke-Search -Pattern '\S' -Path $path -Target Value
    Assert-Equal @('2026/10/01', '2026/10/02', '2026/10/03', '2026/10/04', '46300') @(2..6 | ForEach-Object { $c = 'B' + $_; @($hits | Where-Object { $_.Cell -eq $c })[0].Text })
}

Test-Case 'Search-Excel: 日付のセル（t="d"。値は 2026-10-01T09:30:00 の形）はシリアル値の日付と同じ形で、1904 年からの日付のブックの時刻だけの値も時刻だけで検索される' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $path = Join-Path (New-TestDir) '日付.xlsx'
    Export-Workbook -FilePath $path -Sheets @(@{ Name = '予定'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '値'; Width = 20 }); Rows = @(@{ Values = @('日付', 1) }, @{ Values = @('日時', 2) }, @{ Values = @('時刻', 3) }) })
    Edit-ZipText $path 'xl/worksheets/sheet1.xml' {
        param($x)
        $x.Replace('<c r="B2" s="2"><v>1</v>', '<c r="B2" s="4" t="d"><v>2026-10-01T00:00:00</v>').Replace('<c r="B3" s="2"><v>2</v>', '<c r="B3" s="4" t="d"><v>2026-10-01T09:30:00</v>').Replace('<c r="B4" s="2"><v>3</v>', '<c r="B4" s="4" t="d"><v>09:30:00</v>')
    }
    $hits = Invoke-Search -Pattern '\S' -Path $path -Target Value
    Assert-Equal @('2026/10/01', '2026/10/01 09:30:00', '09:30:00') @(2..4 | ForEach-Object { $c = 'B' + $_; @($hits | Where-Object { $_.Cell -eq $c })[0].Text }) '日付のセル'

    # 1904 年からの日付のブック（Mac の Excel で作ったものなど）: B2 は日付（yyyy/mm/dd）、B3 は時刻（h:mm）の表示形式
    $path = Join-Path (New-TestDir) '1904.xlsx'
    Export-Workbook -FilePath $path -Sheets @(@{ Name = '予定'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '値'; Width = 20 }); Rows = @(@{ Values = @('日付', 44834) }, @{ Values = @('時刻', 0.375) }) })
    Edit-ZipText $path 'xl/workbook.xml' { param($x) $x.Replace('<bookViews>', '<workbookPr date1904="1"/><bookViews>') }
    Edit-ZipText $path 'xl/styles.xml' { param($x) $x.Replace('<cellXfs count="5">', '<cellXfs count="6">').Replace('</cellXfs>', '<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>') }
    Edit-ZipText $path 'xl/worksheets/sheet1.xml' { param($x) $x.Replace('<c r="B2" s="2">', '<c r="B2" s="4">').Replace('<c r="B3" s="2">', '<c r="B3" s="5">') }
    $hits = Invoke-Search -Pattern '\S' -Path $path -Target Value
    Assert-Equal @('2026/10/01', '09:00:00') @(2..3 | ForEach-Object { $c = 'B' + $_; @($hits | Where-Object { $_.Cell -eq $c })[0].Text }) '1904 年からの日付のブック'
}

Test-Case 'Search-Excel: Get-Help でヘルプが表示される' {
    $h = Get-Help $SearchExcel
    Assert-Match $h.Synopsis 'grep'
    Assert-True (@($h.examples.example).Count -ge 4) '使用例'
}
