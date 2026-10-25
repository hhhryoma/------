#     New-DesignDocument.ps1 のテスト。雛形・設計書は docs\design-sample から作る

$NewDocTool = Join-Path $script:ToolsDir 'New-DesignDocument.ps1'
$sampleDir = Join-Path $script:RepoRoot 'docs/design-sample'
$screenList = Join-Path $sampleDir 'DD01_画面一覧.xlsx'

function New-ScreenTemplate {
    #     詳細設計サンプルの画面設計書（SCR001）から、差し込み欄のある雛形を作る
    param([string] $Dir, [string[]] $ExtraRows = @())
    $tpl = Join-Path $Dir '画面設計書_雛形.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD02_画面設計書_SCR001_受注一覧.xlsx') -Destination $tpl
    $csv = Join-Path $Dir 'tpl.csv'
    (@('置換前,置換後,正規表現', 'SCR001,{{画面ID}},', '受注一覧,{{画面名}},', '山田 太郎,{{作成者}},') + $ExtraRows) | Set-Content -LiteralPath $csv -Encoding UTF8
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -ListPath $csv -Path $tpl -Force 6>$null
    return $tpl
}

function Get-CellTexts {
    param([string] $File, [string] $Sheet)
    $map = @{}
    foreach ($h in @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($File)) -Target Value -SheetNamePattern "^$Sheet$" 6>$null)) { $map[$h.Cell] = $h.Text }
    return $map
}

Test-Case 'New-DesignDocument: 画面一覧の行ごとに、雛形の差し込み欄を置き換えた画面設計書を作る（同じ名前のファイルは作らない）' {
    $dir = New-TestDir
    $tpl = New-ScreenTemplate $dir
    $out = Join-Path $dir 'out'
    $opts = @{ Template = $tpl; List = $screenList; ListSheet = '^画面一覧$'; FileName = 'DD02_画面設計書_{{画面ID}}_{{画面名}}.xlsx'; OutDir = $out; Author = '鈴木 次郎' }
    $r = @(& $NewDocTool @opts 6>$null)
    Assert-Equal @('DD02_画面設計書_SCR001_受注一覧.xlsx|作成', 'DD02_画面設計書_SCR002_受注登録.xlsx|作成', 'DD02_画面設計書_SCR003_受注詳細.xlsx|作成') @($r | ForEach-Object { '{0}|{1}' -f $_.File, $_.Status })
    $cells = Get-CellTexts (Join-Path $out 'DD02_画面設計書_SCR003_受注詳細.xlsx') '画面概要'
    Assert-Equal 'SCR003|受注詳細|鈴木 次郎' ('{0}|{1}|{2}' -f $cells['S2'], $cells['AC2'], $cells['AM1'])
    $r = @(& $NewDocTool @opts 6>$null)
    Assert-Equal @('既存', '既存', '既存') @($r | ForEach-Object { $_.Status })
}

Test-Case 'New-DesignDocument: CSV の一覧と {{作成日}} を使える。-WhatIf では作らない。ファイル名・一覧の誤りを知らせる' {
    $dir = New-TestDir
    $tpl = New-ScreenTemplate $dir @('ODR-DD-02-01,ODR-{{作成日}},')
    $list = Join-Path $dir 'list.csv'
    @('画面ID,画面名', 'S10,出荷登録', 'S11,') | Set-Content -LiteralPath $list -Encoding UTF8
    $out = Join-Path $dir 'out'
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}_{{画面名}}.xlsx' -OutDir $out -Date '2026-11-02' -WhatIf 6>$null)
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $out 'S10_出荷登録.xlsx'))) '-WhatIf'
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}_{{画面名}}.xlsx' -OutDir $out -Date '2026-11-02' 6>$null)
    Assert-Equal @('S10_出荷登録.xlsx|作成', '|スキップ') @($r | ForEach-Object { '{0}|{1}' -f $_.File, $_.Status })
    Assert-Match $r[1].Detail '画面名'
    $cells = Get-CellTexts (Join-Path $out 'S10_出荷登録.xlsx') '表紙'
    Assert-True (@($cells.Values) -contains 'ODR-2026/11/02') ($cells.Values -join ', ')
    Assert-Throws { & $NewDocTool -Template $tpl -List $list -FileName 'a.xlsx' -OutDir $out 6>$null } '差し込み欄'
    @('画面ID', 'S12') | Set-Content -LiteralPath $list -Encoding UTF8
    Assert-Throws { & $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}.xlsx' -OutDir $out 6>$null } '画面名'
}

Test-Case 'New-DesignDocument: ヘッダー・フッターの差し込み欄も置き換える（書式で分かれていてもよい）' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '明細_雛形.xlsx'
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'formula/明細.xlsx') -Destination $tpl
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') '受注管理' '{{画面名}}' -Path $tpl -SimpleMatch -Target HeaderFooter -Force 6>$null
    $list = Join-Path $dir 'list.csv'
    @('画面ID,画面名', 'S20,出荷管理') | Set-Content -LiteralPath $list -Encoding UTF8
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}.xlsx' -OutDir (Join-Path $dir 'out') 6>$null)
    Assert-Equal 'S20.xlsx|作成' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    $h = @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern 'システム' -Path ([System.Management.Automation.WildcardPattern]::Escape((Join-Path $dir 'out/S20.xlsx'))) -Target HeaderFooter 6>$null)
    Assert-Equal 'ヘッダー（左）|出荷管理システム' ('{0}|{1}' -f $h[0].Name, $h[0].Text)
}

Test-Case 'New-DesignDocument: -Bump で表紙の版数・更新日と改訂履歴の行を書き換え、版数のチェックが通る' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD05_メッセージ一覧.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    $r = @(& $NewDocTool -Bump -Path $file -Description 'E0017 を追加' -Section 'メッセージ一覧' -Author '佐藤 花子' -Date '2026-11-02' 6>$null)
    Assert-Equal '版上げ' $r[0].Status
    Assert-Match $r[0].Detail '版数 1\.0 → 1\.1（表紙の版数・表紙の更新日・改訂履歴の 7 行目）'
    $cover = Get-CellTexts $file '表紙'
    Assert-Equal '1.1|2026/11/02|2026/10/01' ('{0}|{1}|{2}' -f $cover['X23'], $cover['X25'], $cover['X24']) '版数・最終更新日（作成日はそのまま）'
    $hist = Get-CellTexts $file '改訂履歴'
    Assert-Equal '1.1|2026/11/02|メッセージ一覧|E0017 を追加|佐藤 花子' ('{0}|{1}|{2}|{3}|{4}' -f $hist['A7'], $hist['E7'], $hist['K7'], $hist['W7'], $hist['AY7'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $file -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    # もう一度上げると次の行に足す。-Version で版数を指定できる
    $r = @(& $NewDocTool -Bump -Path $file -Description '全面改訂' -Version '2.0' -Author '佐藤 花子' -Date '2026-12-01' 6>$null)
    $hist = Get-CellTexts $file '改訂履歴'
    Assert-Equal '2.0|2026/12/01|全体' ('{0}|{1}|{2}' -f $hist['A8'], $hist['E8'], $hist['K8'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $file -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
}

Test-Case 'New-DesignDocument: ファイル名に版数があれば新しい版数の名前で別に保存する。-WhatIf では書き換えない' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD03_テーブル定義書_v1.1.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination $file
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $null = & $NewDocTool -Bump -Path $dir -Description 'x' -WhatIf 6>$null
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $dir -Filter '*.xlsx').Count '-WhatIf'
    $r = @(& $NewDocTool -Bump -Path $dir -Description '受注のインデックスを追加' -Section 'T_ORDER' -Date '2026-11-02' 6>$null)
    Assert-Equal 'DD03_テーブル定義書_v1.2.xlsx' $r[0].File
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash '元の版は変えない'
    $new = Join-Path $dir 'DD03_テーブル定義書_v1.2.xlsx'
    $hist = Get-CellTexts $new '改訂履歴'
    Assert-Equal '1.2|T_ORDER' ('{0}|{1}' -f $hist['A8'], $hist['K8'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $new -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    # 新しい版のファイルがすでにあれば作らない
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' 6>$null)
    Assert-Equal 'スキップ' $r[0].Status
    # 版数の無いブック
    $r = @(& $NewDocTool -Bump -Path (Join-Path $script:Fixtures 'grep/sample.xlsx') -Description 'x' -WhatIf 6>$null)
    Assert-Match $r[0].Detail '版数が見つかりません'
}

Test-Case 'New-DesignDocument: -CompareWith で前の版との差分から改訂箇所・改訂内容の行を足す（版数などは最初の行だけ）' {
    $dir = New-TestDir
    $oldDir = Join-Path $dir '納品'
    $null = New-Item -ItemType Directory -Path $oldDir
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination (Join-Path $oldDir 'DD03_テーブル定義書_v1.1.xlsx')
    $file = Join-Path $dir 'DD03_テーブル定義書_v1.1.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination $file
    $replace = Join-Path $script:ToolsDir 'Replace-Excel.ps1'
    $null = & $replace '^配送時間帯$' '配送希望時間帯' -Path $file -SheetNamePattern '^T_ORDER$' -Force 6>$null
    $null = & $replace '^顧客名カナ$' '顧客名（カナ）' -Path $file -SheetNamePattern '^M_CUSTOMER$' -Force 6>$null
    $r = @(& $NewDocTool -Bump -Path $file -CompareWith $oldDir -Author '佐藤 花子' -Date '2026-11-02' 6>$null)
    Assert-Equal 'DD03_テーブル定義書_v1.2.xlsx|版上げ' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    Assert-Match $r[0].Detail '改訂履歴の 8〜9 行目'
    $new = Join-Path $dir 'DD03_テーブル定義書_v1.2.xlsx'
    $hist = Get-CellTexts $new '改訂履歴'
    Assert-Equal '1.2|2026/11/02|M_CUSTOMER|CUSTOMER_NAME_KANA: 論理名を 顧客名カナ → 顧客名（カナ） に変更|佐藤 花子' ('{0}|{1}|{2}|{3}|{4}' -f $hist['A8'], $hist['E8'], $hist['K8'], $hist['W8'], $hist['AY8'])
    Assert-Equal '|||T_ORDER|DELIVERY_TIME_KBN: 論理名を 配送時間帯 → 配送希望時間帯 に変更|' ('|{0}|{1}|{2}|{3}|{4}' -f $hist['A9'], $hist['E9'], $hist['K9'], $hist['W9'], $hist['AY9'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $new -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    # 次の版上げは、続きの行（9 行目）の下に足す
    $r = @(& $NewDocTool -Bump -Path $new -Description '索引を追加' -Date '2026-11-10' 6>$null)
    Assert-Equal '1.3|索引を追加' ('{0}|{1}' -f (Get-CellTexts (Join-Path $dir 'DD03_テーブル定義書_v1.3.xlsx') '改訂履歴')['A10'], (Get-CellTexts (Join-Path $dir 'DD03_テーブル定義書_v1.3.xlsx') '改訂履歴')['W10'])
    # 前の版と同じ内容なら版を上げない。改訂内容も前の版も無ければエラー
    $r = @(& $NewDocTool -Bump -Path (Join-Path $oldDir 'DD03_テーブル定義書_v1.1.xlsx') -CompareWith (Join-Path $oldDir 'DD03_テーブル定義書_v1.1.xlsx') -WhatIf 6>$null)
    Assert-Equal 'スキップ' $r[0].Status
    Assert-Match $r[0].Detail '内容が同じ'
    Assert-Throws { & $NewDocTool -Bump -Path $file 6>$null } 'CompareWith'
}

function Read-BookPart {
    param([string] $Path, [string] $Name)
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $sr = New-Object System.IO.StreamReader($zip.GetEntry($Name).Open())
        try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
    }
    finally { $fs.Dispose() }
}

Test-Case 'New-DesignDocument: -Book / -TemplateSheet で、一覧の行ごとに雛形のシートをコピーして差し込み欄を置き換える' {
    $dir = New-TestDir
    $book = Join-Path $dir 'DD03_テーブル定義書.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination $book
    # 受注明細のシートを雛形にする（テーブル ID・名前を差し込み欄にする）
    $csv = Join-Path $dir 'tpl.csv'
    @('置換前,置換後,正規表現', '^T_ORDER_DETAIL$,{{テーブルID}},○', '^受注明細$,{{テーブル名}},○') | Set-Content -LiteralPath $csv -Encoding UTF8
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -ListPath $csv -Path $book -SheetNamePattern '^T_ORDER_DETAIL$' -Force 6>$null
    $list = Join-Path $dir 'list.csv'
    @('テーブルID,テーブル名', 'T_SHIPMENT,出荷', 'T_ORDER,受注', ',名前だけ') | Set-Content -LiteralPath $list -Encoding UTF8
    $opts = @{ Book = $book; TemplateSheet = 'T_ORDER_DETAIL'; List = $list; SheetName = '{{テーブルID}}' }
    $hash = (Get-FileHash -LiteralPath $book).Hash
    $null = & $NewDocTool @opts -WhatIf 6>$null
    Assert-Equal $hash (Get-FileHash -LiteralPath $book).Hash '-WhatIf'
    $r = @(& $NewDocTool @opts 6>$null)
    Assert-Equal @('|スキップ', 'T_ORDER|既存', 'T_SHIPMENT|作成') @($r | ForEach-Object { '{0}|{1}' -f $_.File, $_.Status } | Sort-Object)
    $cells = Get-CellTexts $book 'T_SHIPMENT'
    Assert-Equal 'T_SHIPMENT|出荷|T_SHIPMENT|出荷' ('{0}|{1}|{2}|{3}' -f $cells['S2'], $cells['AC2'], $cells['I5'], $cells['I6'])
    Assert-Equal '{{テーブルID}}|ORDER_NO' ('{0}|{1}' -f (Get-CellTexts $book 'T_ORDER_DETAIL')['S2'], $cells['J13']) '雛形はそのまま。ほかのセルは写す'
    Assert-Equal 'T_ORDER' (Get-CellTexts $book 'T_ORDER')['S2'] 'ほかのシートは変えない'
    Assert-Match (Read-BookPart $book 'xl/workbook.xml') '<sheet name="T_SHIPMENT" sheetId="8"[^>]*/></sheets>' '末尾に足す'
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $book -Check Error, Link, View 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
}

Test-Case 'New-DesignDocument: 雛形のシートの名前（印刷範囲など）も写し、図形などのあるシートは雛形にしない' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $book = Join-Path $dir 'book.xlsx'
    # Export-Workbook のシートには、オートフィルターの範囲の名前（シートだけの名前）が付く
    Export-Workbook -FilePath $book -Sheets @(
        @{ Name = '一覧'; Columns = @(@{ Header = 'ID'; Width = 8 }, @{ Header = '名前'; Width = 10 }); Rows = @(@{ Values = @('A01', '受付') }, @{ Values = @('A02', '出荷') }) },
        @{ Name = '雛形'; Columns = @(@{ Header = '項目'; Width = 8 }, @{ Header = '値'; Width = 10 }); Rows = @(@{ Values = @('ID', '{{ID}}') }, @{ Values = @('名前', '{{名前}}') }) })
    $r = @(& $NewDocTool -Book $book -TemplateSheet '雛形' -List $book -ListSheet '^一覧$' -SheetName '{{ID}}_{{名前}}' 6>$null)
    Assert-Equal @('A01_受付|作成', 'A02_出荷|作成') @($r | ForEach-Object { '{0}|{1}' -f $_.File, $_.Status })
    Assert-Equal 'A02|出荷' ('{0}|{1}' -f (Get-CellTexts $book 'A02_出荷')['B2'], (Get-CellTexts $book 'A02_出荷')['B3'])
    $wb = Read-BookPart $book 'xl/workbook.xml'
    Assert-Match $wb 'localSheetId="3" hidden="1">&apos;A02_出荷&apos;!\$A\$1:\$B\$3</definedName>' 'シートだけの名前を新しいシートにも付ける'

    $grep = Join-Path $dir 'grep.xlsx'
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'grep/sample.xlsx') -Destination $grep
    $list = Join-Path $dir 'list.csv'
    @('ID', 'X1') | Set-Content -LiteralPath $list -Encoding UTF8
    Assert-Throws { & $NewDocTool -Book $grep -TemplateSheet '顧客一覧' -List $list -SheetName '{{ID}}' 6>$null } 'コメントがあるため、コピーできません'
    Assert-Throws { & $NewDocTool -Book $grep -TemplateSheet '無いシート' -List $list -SheetName '{{ID}}' 6>$null } '見つかりません'
}

Test-Case 'New-DesignDocument: -TemplateSheet の大文字・小文字がシート名と違っても、雛形のシートの名前（オートフィルターの範囲など）を新しいシートに付け替える' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    $book = Join-Path $dir 'book.xlsx'
    Export-Workbook -FilePath $book -Sheets @(
        @{ Name = 'List'; Columns = @(@{ Header = 'ID'; Width = 8 }); Rows = @(@{ Values = @('A01') }) },
        @{ Name = 'Template'; Columns = @(@{ Header = '項目'; Width = 8 }, @{ Header = '値'; Width = 10 }); Rows = @(@{ Values = @('ID', '{{ID}}') }) })
    $r = @(& $NewDocTool -Book $book -TemplateSheet 'template' -List $book -ListSheet '^List$' -SheetName '{{ID}}' 6>$null)
    Assert-Equal 'A01|作成|雛形のシート「Template」から作りました' ('{0}|{1}|{2}' -f $r[0].File, $r[0].Status, $r[0].Detail)
    Assert-Equal 'A01' (Get-CellTexts $book 'A01')['B2']
    Assert-Match (Read-BookPart $book 'xl/workbook.xml') 'localSheetId="2" hidden="1">&apos;A01&apos;!\$A\$1:\$B\$2</definedName>'
}

Test-Case 'New-DesignDocument: 読み取り専用の雛形から作った設計書も、差し込み欄を置き換える（作った設計書は読み取り専用にしない）' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', '{{画面ID}}'), @('画面名', '{{画面名}}')) })
    (Get-Item -LiteralPath $tpl).IsReadOnly = $true
    $list = Join-Path $dir 'list.csv'
    @('画面ID,画面名', 'SCR001,受注一覧') | Set-Content -LiteralPath $list -Encoding UTF8
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}.xlsx' -OutDir (Join-Path $dir 'out') 6>$null)
    Assert-Equal 'SCR001.xlsx|作成' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    $cells = Get-CellTexts $r[0].Path '画面概要'
    Assert-Equal 'SCR001|受注一覧' ('{0}|{1}' -f $cells['B2'], $cells['B3'])
    Assert-True (-not (Get-Item -LiteralPath $r[0].Path).IsReadOnly) '作った設計書が読み取り専用'
}

Test-Case 'New-DesignDocument: Excel の一覧のセル内の改行は、設計書には改行のまま書き、ファイル名では空白にする' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面名', '{{画面名}}'), @('概要', '{{概要}}')) })
    $list = Join-Path $dir 'DD01_画面一覧.xlsx'
    New-SheetsBook $list ([ordered]@{ '画面一覧' = @(@('画面ID', '画面名', '概要'), @('SCR001', "受注`n一覧", "受注を検索する`n一覧から詳細へ遷移する")) })
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}_{{画面名}}.xlsx' -OutDir (Join-Path $dir 'out') 6>$null)
    Assert-Equal 'SCR001_受注 一覧.xlsx|作成' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    $rows = (Read-XlsxSheet $r[0].Path).Rows
    Assert-Equal "受注`n一覧|受注を検索する`n一覧から詳細へ遷移する" ('{0}|{1}' -f $rows[2]['B'], $rows[3]['B'])
}

Test-Case 'New-DesignDocument: 括弧の内側に空白のある差し込み欄（{{ 画面名 }}）も置き換える（値の $ はそのまま書く）' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', '{{画面ID}}'), @('画面名', '{{ 画面名 }}')) })
    $list = Join-Path $dir 'list.csv'
    @('画面ID,画面名', 'SCR001,受注一覧($1)') | Set-Content -LiteralPath $list -Encoding UTF8
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{ 画面ID }}.xlsx' -OutDir (Join-Path $dir 'out') 6>$null)
    Assert-Equal 'SCR001.xlsx|作成' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    $cells = Get-CellTexts $r[0].Path '画面概要'
    Assert-Equal 'SCR001|受注一覧($1)' ('{0}|{1}' -f $cells['B2'], $cells['B3'])
}

Test-Case 'New-DesignDocument: -Bump で版数の桁（1.01 → 1.02）と表紙の版数の書き方（全角・「（案）」など）を保つ。-Version の 0 も落とさない' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD05_メッセージ一覧_v1.01.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    Set-CellValues $file '表紙' @{ X23 = '第1.01版' }
    Set-CellValues $file '改訂履歴' @{ A6 = '1.01' }
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Equal 'DD05_メッセージ一覧_v1.02.xlsx' $r[0].File
    Assert-Match $r[0].Detail '版数 1\.01 → 1\.02'
    Assert-Equal '第1.02版|1.02' ('{0}|{1}' -f (Get-CellTexts $r[0].Path '表紙')['X23'], (Get-CellTexts $r[0].Path '改訂履歴')['A7'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $r[0].Path -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    # 全角の版数は数だけを変え（全角のまま）、ほかの文字はそのまま
    $file = Join-Path $dir 'DD05.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    Set-CellValues $file '表紙' @{ X23 = '第１．０版（案）' }
    $null = & $NewDocTool -Bump -Path $file -Description 'x' -Version '1.02' -Date '2026-11-02' 6>$null
    Assert-Equal '第１．０２版（案）|1.02' ('{0}|{1}' -f (Get-CellTexts $file '表紙')['X23'], (Get-CellTexts $file '改訂履歴')['A7'])
}

Test-Case 'New-DesignDocument: -Bump で数値の版数のセルは表示形式（0.00）の桁で読み、表示が変わる版数（1.9 → 1.10）は文字で書く' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD05.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    Set-CellValues $file '表紙' @{ X23 = [double]1.9 }
    Set-CellValues $file '改訂履歴' @{ A6 = '1.9' }
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Match $r[0].Detail '版数 1\.9 → 1\.10'
    Assert-Equal '1.10|1.10' ('{0}|{1}' -f (Get-CellTexts $file '表紙')['X23'], (Get-CellTexts $file '改訂履歴')['A7'])
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $file -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
    # 表示形式が 0.00 の数値のセル（1.1 を 1.10 と表示する）は 1.10 と読む。表示が同じになる版数（1.11）は数値のまま書く
    $file = Join-Path $dir 'DD05_2.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    Edit-ZipText $file 'xl/styles.xml' { param($t) $t.Replace('<xf numFmtId="0" fontId="5"', '<xf numFmtId="2" fontId="5"') }
    Set-CellValues $file '表紙' @{ X23 = [double]1.1 }
    Set-CellValues $file '改訂履歴' @{ A6 = [double]1.1 }
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Match $r[0].Detail '版数 1\.10 → 1\.11'
    Assert-Match (Read-BookPart $file 'xl/worksheets/sheet1.xml') '<c r="X23" s="6"><v>1\.11</v></c>'
    Assert-Equal '1.11' (Get-CellTexts $file '改訂履歴')['A7']
}

Test-Case 'New-DesignDocument: -Bump で、ファイル名の「第2期」「第3章」は版数にしない（「第1.0版」は版数）' {
    $dir = New-TestDir
    foreach ($name in @('受注システム第2期_DD05_v1.0.xlsx', 'DD05_第3章.xlsx', 'DD05_第1.0版.xlsx')) {
        Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination (Join-Path $dir $name)
    }
    $r = @(& $NewDocTool -Bump -Path $dir -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Equal @('DD05_第1.1版.xlsx', 'DD05_第3章.xlsx', '受注システム第2期_DD05_v1.1.xlsx') @($r | ForEach-Object { $_.File } | Sort-Object)
    Assert-Equal 5 @(Get-ChildItem -LiteralPath $dir -Filter '*.xlsx').Count '版数の無いファイルは上書きする'
}

Test-Case 'New-DesignDocument: -Bump で足す改訂履歴の行に、Excel が行に書く属性（x14ac:dyDescent）を名前空間ごと写す' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD05.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    # Excel で保存した改訂履歴のように、行に x14ac:dyDescent を付ける
    Edit-ZipText $file 'xl/worksheets/sheet2.xml' {
        param($t)
        $t = $t.Replace('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"')
        return ($t -replace '<row r="(\d+)"', '<row r="$1" x14ac:dyDescent="0.4"')
    }
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Match $r[0].Detail '改訂履歴の 7 行目'
    $row = [regex]::Match((Read-BookPart $file 'xl/worksheets/sheet2.xml'), '<row [^>]*\br="7"[^>]*>').Value
    Assert-Match $row '\sx14ac:dyDescent="0\.4"'
    Assert-NotMatch $row '\sdyDescent='
}

Test-Case 'New-DesignDocument: -Bump で表紙の版数・更新日が数式のセルなら、数式は残して計算結果だけを書き換える' {
    $dir = New-TestDir
    $file = Join-Path $dir 'DD05.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD05_メッセージ一覧.xlsx') -Destination $file
    Edit-ZipText $file 'xl/worksheets/sheet1.xml' {
        param($t)
        $t = $t.Replace('<c r="X23" s="6" t="inlineStr"><is><t>1.0</t></is></c>', '<c r="X23" s="6" t="str"><f>LOOKUP(2,1/(改訂履歴!A6:A100&lt;&gt;""),改訂履歴!A6:A100)</f><v>1.0</v></c>')
        return $t.Replace('<c r="X25" s="8" t="n"><v>46296</v></c>', '<c r="X25" s="8"><f>LOOKUP(2,1/(改訂履歴!E6:E100&lt;&gt;""),改訂履歴!E6:E100)</f><v>46296</v></c>')
    }
    $r = @(& $NewDocTool -Bump -Path $file -Description 'x' -Date '2026-11-02' 6>$null)
    Assert-Match $r[0].Detail '版数 1\.0 → 1\.1（表紙の版数（数式の結果）・表紙の更新日（数式の結果）・改訂履歴の 7 行目）'
    $xml = Read-BookPart $file 'xl/worksheets/sheet1.xml'
    Assert-Match $xml '<c r="X23" s="6" t="str"><f>LOOKUP\(2,1/\(改訂履歴!A6:A100&lt;&gt;""\),改訂履歴!A6:A100\)</f><v>1\.1</v></c>'
    Assert-Match $xml '<c r="X25" s="8"><f>LOOKUP\([^<]*\)</f><v>46328</v></c>'
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $file -Check Version 6>$null)
    Assert-Equal 0 $issues.Count (@($issues | ForEach-Object { $_.Detail }) -join ' / ')
}

Test-Case 'New-DesignDocument: -FileName に拡張子が無ければ雛形の拡張子を付け、違う拡張子は受け付けない' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', '{{画面ID}}'), @('画面名', '{{画面名}}')) })
    $list = Join-Path $dir '画面一覧.csv'
    @('画面ID,画面名', 'SCR001,受注一覧') | Set-Content -LiteralPath $list -Encoding UTF8
    $out = Join-Path $dir 'out'
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}_{{画面名}}_v1.0' -OutDir $out 6>$null)
    Assert-Equal 'SCR001_受注一覧_v1.0.xlsx|作成' ('{0}|{1}' -f $r[0].File, $r[0].Status)
    $sheet = Read-XlsxSheet -Path (Join-Path $out 'SCR001_受注一覧_v1.0.xlsx') -Index 1
    Assert-Equal 'SCR001|受注一覧' ('{0}|{1}' -f $sheet.Rows[2]['B'], $sheet.Rows[3]['B']) '差し込み欄を置き換える'
    Assert-Throws { & $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}.xlsm' -OutDir $out 6>$null } '拡張子は雛形と同じ'
}

Test-Case 'New-DesignDocument: -WhatIf では、作るファイル・足すシート・版数と書き換える場所を「予定」の行で知らせる（画面の「試しに実行」の記録の欄に出る）' {
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = New-TestDir
    # 雛形から作る
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', '{{画面ID}}'), @('画面名', '{{画面名}}')) })
    $list = Join-Path $dir '画面一覧.csv'
    @('画面ID,画面名', 'SCR001,受注一覧', 'SCR002,受注登録') | Set-Content -LiteralPath $list -Encoding UTF8
    $out = Join-Path $dir 'out'
    $r = @(& $NewDocTool -Template $tpl -List $list -FileName '{{画面ID}}_{{画面名}}.xlsx' -OutDir $out -WhatIf -InformationVariable info 6>$null)
    Assert-Equal @('SCR001_受注一覧.xlsx|作成予定', 'SCR002_受注登録.xlsx|作成予定') @($r | ForEach-Object { '{0}|{1}' -f $_.File, $_.Status })
    Assert-Match (@($info | ForEach-Object { [string]$_.MessageData }) -join ' / ') '2 ファイルを作る予定です（一覧の 2 行'
    Assert-True (-not (Test-Path -LiteralPath $out)) '作らない'

    # 雛形のシートからシートを足す
    $book = Join-Path $dir 'book.xlsx'
    Export-Workbook -FilePath $book -Sheets @(
        @{ Name = '一覧'; Columns = @(@{ Header = 'ID'; Width = 8 }); Rows = @(@{ Values = @('A01') }, @{ Values = @('A02') }) },
        @{ Name = '雛形'; Columns = @(@{ Header = '項目'; Width = 8 }, @{ Header = '値'; Width = 10 }); Rows = @(@{ Values = @('ID', '{{ID}}') }) })
    $hash = (Get-FileHash -LiteralPath $book).Hash
    $r = @(& $NewDocTool -Book $book -TemplateSheet '雛形' -List $book -ListSheet '^一覧$' -SheetName '{{ID}}' -WhatIf 6>$null)
    Assert-Equal @('A01|作成予定|雛形のシート「雛形」から作ります', 'A02|作成予定|雛形のシート「雛形」から作ります') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.File, $_.Status, $_.Detail })
    Assert-Equal $hash (Get-FileHash -LiteralPath $book).Hash '書き換えない'

    # 版を上げる（書き換える場所は、実際に版を上げたときと同じ）
    $file = Join-Path $dir 'DD03_テーブル定義書_v1.1.xlsx'
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination $file
    $hash = (Get-FileHash -LiteralPath $file).Hash
    $r = @(& $NewDocTool -Bump -Path $file -Description '索引を追加' -WhatIf 6>$null)
    Assert-Equal 'DD03_テーブル定義書_v1.1.xlsx → DD03_テーブル定義書_v1.2.xlsx|版上げ予定|版数 1.1 → 1.2（表紙の版数・表紙の更新日・改訂履歴の 8 行目）' ('{0}|{1}|{2}' -f $r[0].File, $r[0].Status, $r[0].Detail)
    Assert-Equal $hash (Get-FileHash -LiteralPath $file).Hash '書き換えない'
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $dir -Filter 'DD03_*.xlsx').Count '新しい版のファイルも作らない'

    # 前の版と比べるときも、-WhatIf で改訂履歴に足す行を数える（前の版と同じ内容とみなさない）
    $oldDir = Join-Path $dir '納品'
    New-Item -ItemType Directory -Path $oldDir | Out-Null
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination (Join-Path $oldDir 'DD03_テーブル定義書_v1.1.xlsx')
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') '^配送時間帯$' '配送希望時間帯' -Path $file -SheetNamePattern '^T_ORDER$' -Force 6>$null
    $r = @(& $NewDocTool -Bump -Path $file -CompareWith $oldDir -WhatIf 6>$null)
    Assert-Equal '版上げ予定' $r[0].Status $r[0].Detail
    Assert-Match $r[0].Detail '改訂履歴の 8 行目'
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $dir -Filter 'DD03_*.xlsx').Count '新しい版のファイルを作らない'
}

Test-Case 'New-DesignDocument: Get-Help でヘルプが表示される' {
    $h = Get-Help $NewDocTool
    Assert-Match $h.Synopsis '雛形'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'New-DesignDocument: -Bump で新しい版数の名前で保存したときは、版を上げる前のファイルが残っていることを知らせる（設計書を読むツールは新しい版だけを読む）' {
    $dir = New-TestDir
    Copy-Item -LiteralPath (Join-Path $sampleDir 'DD03_テーブル定義書.xlsx') -Destination (Join-Path $dir 'DD03_テーブル定義書_v1.1.xlsx')
    $w = $null
    $null = & $NewDocTool -Bump -Path $dir -Description 'x' -WhatIf -WarningVariable w -WarningAction SilentlyContinue 6>$null
    Assert-Equal 0 @($w).Count '-WhatIf では知らせない'
    $null = & $NewDocTool -Bump -Path $dir -Description '受注のインデックスを追加' -WarningVariable w -WarningAction SilentlyContinue 6>$null
    Assert-Match (($w | ForEach-Object { $_.Message }) -join "`n") '版を上げる前のファイル 1 件を、同じフォルダに残しています（DD03_テーブル定義書_v1\.1\.xlsx）。版を上げる前に直したときは'
    $tables = @(& (Join-Path $script:ToolsDir 'Measure-DesignSize.ps1') -Path $dir 6>$null 3>$null | Where-Object Kind -eq 'テーブル')
    Assert-Equal 4 $tables.Count '古い版のテーブルを二重に数えない'
}
