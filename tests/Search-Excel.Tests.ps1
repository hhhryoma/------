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

Test-Case 'Search-Excel: -Path にワイルドカードを使える' {
    $hits = Invoke-Search -Pattern 'TODO' -Path (Join-Path $grep 'sub/*.xlsx')
    Assert-Equal 3 $hits.Count
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

Test-Case 'Search-Excel: Get-Help でヘルプが表示される' {
    $h = Get-Help $SearchExcel
    Assert-Match $h.Synopsis 'grep'
    Assert-True (@($h.examples.example).Count -ge 4) '使用例'
}
