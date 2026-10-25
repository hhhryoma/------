#     Replace-Excel.ps1 のテスト。フィクスチャは tests\fixtures\replace と tests\fixtures\grep（make_fixtures.py で生成）

$ReplaceExcel = Join-Path $script:ToolsDir 'Replace-Excel.ps1'
$SearchExcel = Join-Path $script:ToolsDir 'Search-Excel.ps1'

function Invoke-Replace {
    param([hashtable] $Options)
    return , @(& $ReplaceExcel @Options 6>$null 3>$null)
}

function Get-Texts {
    #     Search-Excel で、ファイルの中の文字を「場所=文字」の一覧にする
    param([string] $Path, [string] $Pattern = '[\s\S]')
    return @(& $SearchExcel -Pattern $Pattern -Path ([System.Management.Automation.WildcardPattern]::Escape($Path)) -Target Value, Comment, Shape 6>$null 3>$null |
        ForEach-Object { '{0}={1}' -f $_.Location, $_.Text } | Sort-Object)
}

function Read-PartText {
    param([string] $Path, [string] $Name)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $sr = New-Object System.IO.StreamReader($zip.GetEntry($Name).Open())
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    }
    finally { $fs.Dispose() }
}

function Assert-WellFormed {
    #     xlsx の中の XML がすべて読めることを確かめる
    param([string] $Path)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        foreach ($e in $zip.Entries) {
            if ($e.FullName -notmatch '\.(xml|rels)$') { continue }
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { $null = [xml]$sr.ReadToEnd() } catch { throw "$($e.FullName) が XML として読めません: $($_.Exception.Message)" } finally { $sr.Dispose() }
        }
    }
    finally { $fs.Dispose() }
}

Test-Case 'Replace-Excel: -WhatIf では置換する箇所を示すだけで、ファイルを変えない' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $r = Invoke-Replace @{ Pattern = '受注管理システム'; Replacement = '販売管理システム'; Path = $file; SimpleMatch = $true; WhatIf = $true }
    Assert-Equal @('本文!A1', '表紙!A1', '表紙!A2') @($r | ForEach-Object { $_.Location } | Sort-Object)
    Assert-Equal @('置換予定') @($r | ForEach-Object { $_.Result } | Sort-Object -Unique)
    Assert-Equal '受注管理システム 基本設計書 → 販売管理システム 基本設計書' (@($r | Where-Object Location -eq '表紙!A1')[0].Change)
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash 'ファイルが変わっています'
}

Test-Case 'Replace-Excel: 共有文字列を置換し、ふりがなは消し、ほかの文字列は変えない' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $before = Get-Texts $file
    $r = Invoke-Replace @{ Pattern = '受注管理システム'; Replacement = '販売管理システム'; Path = $file; SimpleMatch = $true; Force = $true }
    Assert-Equal @('置換') @($r | ForEach-Object { $_.Result } | Sort-Object -Unique)
    Assert-WellFormed $file
    $expected = @($before | ForEach-Object { $_.Replace('受注管理システム', '販売管理システム') } | Sort-Object)
    Assert-Equal $expected (Get-Texts $file)
    $sst = Read-PartText $file 'xl/sharedStrings.xml'
    Assert-NotMatch $sst 'ジュチュウ' '置換した文字列のふりがなが残っています'
    Assert-Match $sst '<si><t xml:space="preserve">販売管理システム</t><phoneticPr fontId="1" ?/></si>' 'ふりがなの設定（phoneticPr）は残す'
}

Test-Case 'Replace-Excel: 対象外のシートと共有している文字列は、元の文字列を残して対象のセルだけ付け替える' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $r = Invoke-Replace @{ Pattern = '受注管理システム'; Replacement = '販売管理システム'; Path = $file; SimpleMatch = $true; Force = $true; SheetNamePattern = '^表紙$' }
    Assert-Equal @('表紙!A1', '表紙!A2') @($r | ForEach-Object { $_.Location } | Sort-Object)
    $texts = Get-Texts $file
    Assert-True ($texts -contains '本文!A1=受注管理システム') '対象外のシートの文字は元のまま'
    Assert-True ($texts -contains '表紙!A2=販売管理システム') '対象のシートの同じ文字列は置換'
    $sst = Read-PartText $file 'xl/sharedStrings.xml'
    Assert-Match $sst 'uniqueCount="6"' '置換後の文字列を 1 つ足す'
}

Test-Case 'Replace-Excel: 正規表現の $1 を使え、書式の違う部分にまたがる文字も置換できる' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    # 「第1.0版」は LibreOffice がフォントごとに 3 つの部分（第 / 1.0 / 版）に分けて保存している
    Assert-Match (Read-PartText $file 'xl/sharedStrings.xml') '<t xml:space="preserve">第</t></r><r>.*?<t xml:space="preserve">1.0</t>' '前提: 部分に分かれている'
    $null = Invoke-Replace @{ Pattern = '第(\d+)\.0版'; Replacement = '第$1.1版'; Path = $file; Force = $true }
    Assert-True ((Get-Texts $file) -contains '表紙!A3=第1.1版')
}

Test-Case 'Replace-Excel: 文字列の一部だけの書式（赤字）は、置換していない文字ではそのまま' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $null = Invoke-Replace @{ Pattern = '客の正式'; Replacement = '客X'; Path = $file; Force = $true }
    Assert-True ((Get-Texts $file) -contains '本文!B1=顧客X名称')
    # 一致の先頭がある部分（黒の「顧客の」）に置換後の文字が入り、赤の部分には「名称」が残る
    Assert-Match (Read-PartText $file 'xl/sharedStrings.xml') '<t xml:space="preserve">顧客X</t></r><r><rPr>(?:(?!</rPr>).)*FFFF0000(?:(?!</rPr>).)*</rPr><t xml:space="preserve">名称</t>'
}

Test-Case 'Replace-Excel: テーブルの見出しは置換せず知らせる。セル内の文字列・コメントは置換する' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'table.xlsx'
    $table = Read-PartText $file 'xl/tables/table1.xml'
    $w = $null
    $r = @(& $ReplaceExcel -Pattern 'TODO' -Replacement '済' -Path $file -Force -WarningVariable w -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @('一覧!A2|コメント', '一覧!A2|値') @($r | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Kind } | Sort-Object)
    Assert-Match (($w | ForEach-Object { $_.Message }) -join "`n") 'テーブルの見出し.*一覧!A1'
    $texts = Get-Texts $file
    Assert-True ($texts -contains '一覧!A1=TODO列') '見出しは元のまま'
    Assert-True ($texts -contains '一覧!A2=済値')
    Assert-True ($texts -contains '一覧!A2=済: 確認') 'コメント'
    Assert-Equal $table (Read-PartText $file 'xl/tables/table1.xml') 'テーブルの定義は変えない'
}

Test-Case 'Replace-Excel: -Target Formula で数式を置換し、開いたときに再計算させる' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'table.xlsx'
    $r = Invoke-Replace @{ Pattern = 'B3\)'; Replacement = 'B4)'; Path = $file; Target = 'Formula'; Force = $true }
    Assert-Equal '一覧!C2|数式|=SUM(B2:B3) → =SUM(B2:B4)' ('{0}|{1}|{2}' -f $r[0].Location, $r[0].Kind, $r[0].Change)
    Assert-Match (Read-PartText $file 'xl/worksheets/sheet1.xml') '<f>SUM\(B2:B4\)</f>'
    Assert-Match (Read-PartText $file 'xl/workbook.xml') '<calcPr [^>]*fullCalcOnLoad="1"'
    Assert-WellFormed $file
}

Test-Case 'Replace-Excel: 共有数式（コピーした数式）は写しのセルも置換し、同じ形なら共有のまま、一部だけ変わるなら共有をやめる' {
    $dir = Copy-Fixture 'formula'
    $file = Join-Path $dir '明細.xlsx'
    $r = Invoke-Replace @{ Pattern = 'ROUND('; Replacement = 'ROUNDDOWN('; SimpleMatch = $true; Path = $file; Target = 'Formula'; Force = $true }
    Assert-Equal @('C2', 'C3', 'C4', 'C5', 'C6') @($r | ForEach-Object { $_.Cell }) '写しのセルも記録する'
    Assert-Equal '=ROUND(A6*B6*(1+税率),0) → =ROUNDDOWN(A6*B6*(1+税率),0)' $r[4].Change
    $sheet = Read-PartText $file 'xl/worksheets/sheet1.xml'
    Assert-Match $sheet '<f t="shared" ref="C2:C6" si="0">ROUNDDOWN\(A2\*B2' '先頭の数式だけ書き換える'
    Assert-Equal 4 ([regex]::Matches($sheet, '<f t="shared" si="0" ?/>').Count) '写しのセルは共有のまま'

    $r = Invoke-Replace @{ Pattern = 'A4*B4'; Replacement = 'A4*B3'; SimpleMatch = $true; Path = $file; Target = 'Formula'; Force = $true }
    Assert-Equal '明細!C4|=ROUNDDOWN(A4*B4*(1+税率),0) → =ROUNDDOWN(A4*B3*(1+税率),0)' ('{0}|{1}' -f $r[0].Location, $r[0].Change)
    Assert-Equal 1 $r.Count
    $sheet = Read-PartText $file 'xl/worksheets/sheet1.xml'
    Assert-Match $sheet '<c r="C3"><f>ROUNDDOWN\(A3\*B3\*\(1\+税率\),0\)</f>' 'ほかのセルにも数式を書く'
    Assert-Match $sheet '<c r="C4"><f>ROUNDDOWN\(A4\*B3\*'
    Assert-Equal 0 ([regex]::Matches($sheet, 'si="0"').Count) 'C 列の共有をやめる'
    Assert-Equal 4 ([regex]::Matches($sheet, '<f t="shared" si="1" ?/>').Count) 'ほかの共有数式はそのまま'
    $texts = @(& $SearchExcel -Pattern '.' -Path ([System.Management.Automation.WildcardPattern]::Escape($file)) -Target Formula 6>$null | Where-Object Cell -eq 'C6')
    Assert-Equal '=ROUNDDOWN(A6*B6*(1+税率),0)' $texts[0].Text
    Assert-WellFormed $file
}

Test-Case 'Replace-Excel: ヘッダー・フッターと入力規則のリスト・メッセージも置換する（書式・ページ番号などはそのまま）' {
    $dir = Copy-Fixture 'formula'
    $file = Join-Path $dir '明細.xlsx'
    $list = Join-Path (New-TestDir) 'list.csv'
    @('置換前,置換後', '受注管理システム,販売管理システム', 'A&B,A&C', '対応中,作業中', 'リストから,一覧から', '集計,合計') | Set-Content -LiteralPath $list -Encoding UTF8
    $r = Invoke-Replace @{ ListPath = $list; Path = $file; Force = $true }
    Assert-Equal @(
        '明細|ヘッダー・フッター|ヘッダー（左）|受注管理システム → 販売管理システム',
        '明細|ヘッダー・フッター|フッター（右）|株式会社A&B → 株式会社A&C',
        '明細!F2|入力規則|リスト（F2:F6）|未着手,対応中,完了 → 未着手,作業中,完了',
        '明細!F2|入力規則|入力時メッセージ（F2:F6）|未着手・対応中・完了から選んでください → 未着手・作業中・完了から選んでください',
        '明細!F2|入力規則|エラーメッセージ（F2:F6）|リストから選んでください → 一覧から選んでください') `
        @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Location, $_.Kind, $_.Name, $_.Change }) 'ほかのシートを参照するリスト（数式）は -Target Formula のときだけ'
    $sheet = Read-PartText $file 'xl/worksheets/sheet1.xml'
    Assert-Match $sheet '<oddHeader>&amp;L&amp;"ＭＳ ゴシック,太字"販売管理システム&amp;"ＭＳ ゴシック,標準"&amp;R&amp;D</oddHeader>'
    Assert-Match $sheet '<oddFooter>&amp;C&amp;P / &amp;N&amp;R株式会社A&amp;&amp;C</oddFooter>' '文字の & は && のまま'
    Assert-Match $sheet '<formula1>"未着手,作業中,完了"</formula1>'
    Assert-Match $sheet "<xm:f>'集計 A1'!\`$A\`$1:\`$A\`$3</xm:f>"
    Assert-WellFormed $file

    $r = Invoke-Replace @{ Pattern = "'集計 A1'"; Replacement = "'担当者'"; SimpleMatch = $true; Path = $file; Target = 'Formula'; Force = $true }
    Assert-Equal @('明細!E2', '明細!E3', '明細!E4', '明細!E5', '明細!E6', '明細!G2') @($r | ForEach-Object { $_.Location })
    Assert-Equal "リスト（G2:G6）|='集計 A1'!`$A`$1:`$A`$3 → ='担当者'!`$A`$1:`$A`$3" ('{0}|{1}' -f $r[5].Name, $r[5].Change)
}

Test-Case 'Replace-Excel: 図形・隠しシート・セル番地を省略したセルも置換し、-BackupDir には変えたファイルだけ残す' {
    $dir = Copy-Fixture 'grep'
    $backup = Join-Path (New-TestDir) 'backup'
    $r = Invoke-Replace @{ Pattern = 'TODO'; Replacement = 'DONE'; Path = $dir; Force = $true; BackupDir = $backup }
    $byFile = @{}
    foreach ($x in $r) { $byFile[(ConvertTo-Slash $x.File)] += @($x) }
    Assert-Equal @('sample.xlsx|Sheet3!A2000|値', 'sample.xlsx|作業用!A1|値', 'sample.xlsx|顧客一覧!B10|図形', 'sample.xlsx|顧客一覧!B4|値', 'sample.xlsx|顧客一覧!G3|図形', 'sub/inline.xlsx|inline!A1|値', 'sub/inline.xlsx|inline!B2|値', "sub/quote.xlsx|O'Brien 表 1!C5|値") `
        @($r | Where-Object Result -eq '置換' | ForEach-Object { '{0}|{1}|{2}' -f (ConvertTo-Slash $_.File), $_.Location, $_.Kind } | Sort-Object)
    Assert-Equal 'スキップ' $byFile['sub/encrypted.xlsx'][0].Result 'パスワード付き'
    Assert-Equal '失敗' $byFile['sub/broken.xlsx'][0].Result '壊れたファイル'
    Assert-Equal 0 @(Get-Texts (Join-Path $dir 'sample.xlsx') 'TODO').Count 'sample.xlsx に TODO が残っています'
    Assert-Equal 'インライン↵DONE' (@($r | Where-Object Location -eq 'inline!A1')[0].After) '改行（_x000D_）は表示では ↵'
    Assert-WellFormed (Join-Path $dir 'sample.xlsx')
    Assert-Equal @('sample.xlsx', 'sub/inline.xlsx', 'sub/quote.xlsx') @(Get-ChildItem -LiteralPath $backup -Recurse -File | ForEach-Object { ConvertTo-Slash $_.FullName.Substring($backup.Length + 1) } | Sort-Object)
    Assert-Equal (Get-FileHash -LiteralPath (Join-Path $script:Fixtures 'grep/sample.xlsx')).Hash (Get-FileHash -LiteralPath (Join-Path $backup 'sample.xlsx')).Hash 'バックアップは元のまま'
}

Test-Case 'Replace-Excel: 読み取り専用のファイルは書き換えない' {
    $dir = Copy-Fixture 'replace'
    $file = Get-Item -LiteralPath (Join-Path $dir 'shared.xlsx')
    $file.IsReadOnly = $true
    try {
        $hash = (Get-FileHash -LiteralPath $file.FullName).Hash
        $r = Invoke-Replace @{ Pattern = '受注'; Replacement = '販売'; Path = $file.FullName; Force = $true }
        Assert-Equal 'スキップ' $r[0].Result
        Assert-Match $r[0].Detail '読み取り専用'
        Assert-Equal $hash (Get-FileHash -LiteralPath $file.FullName).Hash
    }
    finally { (Get-Item -LiteralPath $file.FullName).IsReadOnly = $false }
}

Test-Case 'Replace-Excel: 開いている（ロックされた）ファイルは失敗として知らせる' -WindowsOnly {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $lock = [System.IO.File]::Open($file, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
    try {
        $r = Invoke-Replace @{ Pattern = '受注'; Replacement = '販売'; Path = $file; Force = $true }
        Assert-Equal @('失敗') @($r | ForEach-Object { $_.Result } | Sort-Object -Unique)
        Assert-Match $r[0].Detail '開いている'
    }
    finally { $lock.Dispose() }
}

Test-Case 'Replace-Excel: -OutXlsx で置換した箇所をリンク付きで保存する' {
    $dir = Copy-Fixture 'replace'
    $out = Join-Path (New-TestDir) 'result.xlsx'
    $null = Invoke-Replace @{ Pattern = '受注管理システム'; Replacement = '販売管理システム'; Path = $dir; SimpleMatch = $true; WhatIf = $true; OutXlsx = $out }
    $sheet = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal '置換前' $sheet.Rows[1]['E']
    Assert-Equal 4 $sheet.Rows.Count '見出しと 3 か所'
    Assert-True (@($sheet.Links | Where-Object { $_.Location -eq "'表紙'!A2" }).Count -eq 1) '該当セルへのリンク'
}

Test-Case 'Replace-Excel: 一致しなければ何も変えず、不正な正規表現はエラー' {
    $dir = Copy-Fixture 'replace'
    $file = Join-Path $dir 'shared.xlsx'
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $r = Invoke-Replace @{ Pattern = '存在しない文字'; Replacement = 'x'; Path = $file; Force = $true }
    Assert-Equal 0 $r.Count
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash
    Assert-Throws { & $ReplaceExcel -Pattern '(' -Replacement 'x' -Path $file -Force 6>$null } '正規表現として不正'
}

Test-Case 'Replace-Excel: 置換後の文字と大文字・小文字だけが違う語も置換する（置換後の文字の一部は、大文字・小文字が違っても置換しない）' {
    $file = Join-Path (New-TestDir) 'case.xlsx'
    New-SheetsBook $file ([ordered]@{ 'S' = @(@('項目'), @('Javascriptで実装'), @('JAVASCRIPT'), @('JavaScript'), @('DOCUMENT')) })
    $opt = @{ Pattern = 'Javascript'; Replacement = 'JavaScript'; Path = $file; SimpleMatch = $true; Force = $true }
    $r = Invoke-Replace $opt
    Assert-Equal @('S!A2|Javascriptで実装 → JavaScriptで実装', 'S!A3|JAVASCRIPT → JavaScript') @($r | ForEach-Object { '{0}|{1}' -f $_.Location, $_.Change } | Sort-Object)
    Assert-Equal 0 (Invoke-Replace $opt).Count '2 回目は置換する箇所がない'
    $r = Invoke-Replace @{ Pattern = 'Doc'; Replacement = 'Document'; Path = $file; SimpleMatch = $true; Force = $true }
    Assert-Equal 0 $r.Count '「DOCUMENT」の「DOC」は置換しない（「DocumentUMENT」にしない）'
}

Test-Case 'Replace-Excel: -BackupDir には、別々のフォルダの同じ名前のファイルをフォルダごとに分けて残し、前に残したファイルは上書きしない' {
    $dir = New-TestDir
    $hashes = New-Object System.Collections.Generic.List[string]
    foreach ($round in @('第1回', '第2回')) {
        $file = Join-Path $dir "$round/仕様書.xlsx"
        New-Item -ItemType Directory -Force -Path (Split-Path $file -Parent) | Out-Null
        New-SheetsBook $file ([ordered]@{ 'S' = @(@('項目'), @("受注管理システム $round"), @('TODO')) })
        $hashes.Add((Get-FileHash -LiteralPath $file).Hash)
    }
    $roots = @((Join-Path $dir '第1回'), (Join-Path $dir '第2回'))
    $backup = Join-Path $dir 'bk'
    $null = Invoke-Replace @{ Pattern = '受注'; Replacement = '販売'; Path = $roots; Force = $true; BackupDir = $backup }
    Assert-Equal $hashes.ToArray() @(@('第1回', '第2回') | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $backup "$_/仕様書.xlsx")).Hash }) '指定したフォルダの名前のフォルダに、元のファイルを残す'

    # 2 回目の置換でも、前に残したファイルは上書きせず、番号を付けて残す
    $replaced = (Get-FileHash -LiteralPath (Join-Path $dir '第1回/仕様書.xlsx')).Hash
    $text = Get-HostText { $null = & $ReplaceExcel -Pattern 'TODO' -Replacement '済み' -Path $roots -Force -BackupDir $backup }
    Assert-Equal $hashes[0] (Get-FileHash -LiteralPath (Join-Path $backup '第1回/仕様書.xlsx')).Hash '1 回目に残した元のファイル'
    Assert-Equal $replaced (Get-FileHash -LiteralPath (Join-Path $backup '第1回/仕様書 (2).xlsx')).Hash '2 回目に書き換える前のファイル'
    Assert-Match $text '番号を付けて残し.*仕様書 \(2\)\.xlsx'
}

Test-Case 'Replace-Excel: Get-Help でヘルプが表示される' {
    $h = Get-Help $ReplaceExcel
    Assert-Match $h.Synopsis '置換'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
