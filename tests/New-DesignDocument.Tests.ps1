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

Test-Case 'New-DesignDocument: Get-Help でヘルプが表示される' {
    $h = Get-Help $NewDocTool
    Assert-Match $h.Synopsis '雛形'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
