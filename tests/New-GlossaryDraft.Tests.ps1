#     New-GlossaryDraft.ps1 のテスト。設計書はテストの中で作る

$GlossaryTool = Join-Path $script:ToolsDir 'New-GlossaryDraft.ps1'

function New-TermBook {
    #     カタカナ語・略語・用語と、その表記ゆれのある設計書
    param([string] $Dir)
    New-SheetsBook (Join-Path $Dir '基本設計書.xlsx') ([ordered]@{
            '概要' = @(
                , @('説明', 'メモ')
                , @('ユーザーを登録する', 'ID を採番する')
                , @('ユーザ一覧を表示する', '受注番号を入力する')
                , @('ユーザー名を表示する', '受注番号')
                , @('ﾕｰｻﾞｰの権限', 'ＩＤ は 10 桁')
                , @('データベースに登録する', '受注番号の採番')
                , @('データ・ベースの接続', 'SCR001 で ID を表示する')
                , @('データベースを更新する', 'E0001 を表示して確認する')
                , @('確認', '確認'))
        })
}

Test-Case 'New-GlossaryDraft: カタカナ語・略語・用語を数え、全角半角・語末の長音・中黒の違いを表記ゆれとしてまとめる' {
    $dir = New-TestDir
    New-TermBook $dir
    $terms = @(& $GlossaryTool $dir 6>$null)
    Assert-Equal @(
        'ユーザー|カタカナ語|4|ユーザ（1）・ﾕｰｻﾞｰ（1）',
        'データベース|カタカナ語|3|データ・ベース（1）',
        'ID|略語|3|ＩＤ（1）',
        '受注番号|用語|3|') @($terms | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Term, $_.Kind, $_.Count, $_.Variants }) '3 回以上のもの。SCR001・E0001 は略語にしない。「確認」はどこにでも出てくる言葉なので用語にしない'
    Assert-Equal '基本設計書.xlsx|概要!A2' ('{0}|{1}' -f $terms[0].File, $terms[0].Location) '最初の場所'

    $only = @(& $GlossaryTool $dir -Kind Katakana -MinCount 4 -Exclude 'データベース' 6>$null)
    Assert-Equal @('ユーザー') @($only | ForEach-Object { $_.Term }) '-Kind・-MinCount・-Exclude'

    $text = Get-HostText { & $GlossaryTool $dir | Out-Null }
    Assert-Match $text '1 ファイルから用語 4（カタカナ語 2・略語 1・用語 1）。3 回以上出てくるもの'
    Assert-Match $text '表記ゆれ 3 組: ユーザー／ユーザ／ﾕｰｻﾞｰ・データベース／データ・ベース・ID／ＩＤ'
}

Test-Case 'New-GlossaryDraft: -OutCsv の用語集は Test-ExcelTerms.ps1 の -Glossary にそのまま渡せ、-OutXlsx に用語集の下書きを保存する' {
    $dir = New-TestDir
    $docs = Join-Path $dir '設計書'
    New-Item -ItemType Directory -Path $docs | Out-Null
    New-TermBook $docs
    $csv = Join-Path $dir '用語集.csv'
    $out = Join-Path $dir '用語集の下書き.xlsx'
    $null = & $GlossaryTool $docs -OutCsv $csv -OutXlsx $out 6>$null
    $bytes = [System.IO.File]::ReadAllBytes($csv)
    Assert-Equal '239|187|191' (@($bytes[0..2]) -join '|') 'CSV は BOM 付き'
    $lines = @([System.IO.File]::ReadAllLines($csv))
    Assert-Equal '"正しい表記","揺れ","備考"' $lines[0]
    Assert-Equal '"ユーザー","ユーザ/ﾕｰｻﾞｰ","ユーザー（2）・ユーザ（1）・ﾕｰｻﾞｰ（1）"' $lines[1]

    $found = @(& (Join-Path $script:ToolsDir 'Test-ExcelTerms.ps1') -Path $docs -Glossary $csv -NoAuto 6>$null 3>$null)
    Assert-Equal (@('データ・ベース→データベース', 'ユーザ→ユーザー', 'ﾕｰｻﾞｰ→ユーザー', 'ＩＤ→ID') | Sort-Object) (@($found | ForEach-Object { '{0}→{1}' -f $_.Variant, $_.Preferred }) | Sort-Object) 'Test-ExcelTerms が用語集として読める'

    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('用語集', '表記ゆれ') $book.Sheets
    Assert-Equal '用語|意味|種類|出現数|設計書の数|表記ゆれ|最初の場所' ((@('A', 'B', 'C', 'D', 'E', 'F', 'G') | ForEach-Object { $book.Rows[1][$_] }) -join '|')
    Assert-Equal 'ユーザー||カタカナ語|4|1' ('{0}|{1}|{2}|{3}|{4}' -f $book.Rows[2]['A'], $book.Rows[2]['B'], $book.Rows[2]['C'], $book.Rows[2]['D'], $book.Rows[2]['E']) '意味は空欄（人が書く）'
    Assert-True (@($book.Links | Where-Object { $_.Target -like '*基本設計書.xlsx' -and $_.Location -eq "'概要'!A2" }).Count -gt 0) '最初の場所へのリンク'
}

Test-Case 'New-GlossaryDraft: -IncludeWord で Word の文書の言葉も数える' {
    $dir = New-TestDir
    New-WordDoc (Join-Path $dir '要件定義書.docx') @('ログインユーザーを確認する', 'ログインユーザの権限', 'ログインユーザーの一覧')
    $terms = @(& $GlossaryTool $dir -IncludeWord -Kind Katakana 6>$null)
    Assert-Equal @('ログインユーザー|3|ログインユーザ（1）') @($terms | ForEach-Object { '{0}|{1}|{2}' -f $_.Term, $_.Count, $_.Variants })
    Assert-Equal 0 @(& $GlossaryTool $dir -Kind Katakana 6>$null).Count '-IncludeWord が無ければ Word は読まない'
}

Test-Case 'New-GlossaryDraft: Get-Help でヘルプが表示される' {
    $h = Get-Help $GlossaryTool
    Assert-Match $h.Synopsis '用語集'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
