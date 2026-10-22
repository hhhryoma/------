#     Measure-ReviewEffectiveness.ps1 のテスト。指摘管理表・障害管理表・ページ数の CSV はテストの中で作る

$EffectTool = Join-Path $script:ToolsDir 'Measure-ReviewEffectiveness.ps1'

function New-EffectData {
    param([string] $Dir)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $finding = { param($doc, $text) , @($doc, $text, '誤り', '完了') }
    $rows = @(
        (& $finding 'DD02_画面設計書_SCR001_受注一覧' '桁数が違う'), (& $finding 'DD02_画面設計書_SCR001_受注一覧' '型が違う'),
        (& $finding 'DD02_画面設計書_SCR001_受注一覧' 'イベントの抜け'), (& $finding 'DD02_画面設計書_SCR001_受注一覧' '遷移先の誤り'),
        (& $finding 'DD02_画面設計書_SCR002_受注登録' '必須の誤り'),
        (& $finding 'DD02_画面設計書_SCR003_受注詳細' '表記'), (& $finding 'DD02_画面設計書_SCR003_受注詳細' '桁数'), (& $finding 'DD02_画面設計書_SCR003_受注詳細' '型'), (& $finding 'DD02_画面設計書_SCR003_受注詳細' '抜け'),
        (& $finding 'DD06_IF定義書_IF001_出荷指示データ' '項目の抜け'), (& $finding 'DD06_IF定義書_IF001_出荷指示データ' '桁数'))
    Export-Workbook -FilePath (Join-Path $Dir '指摘管理表.xlsx') -Sheets @(@{ Name = '指摘'; Columns = @(@('設計書', '指摘内容', '分類', '状態') | ForEach-Object { @{ Header = $_; Width = 12 } }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) })
    $pages = Join-Path $Dir 'pages.csv'
    @('"ファイル","ページ数"', '"DD02_画面設計書_SCR001_受注一覧.pdf","4"', '"DD02_画面設計書_SCR002_受注登録.pdf","10"', '"DD02_画面設計書_SCR003_受注詳細.pdf","4"', '"DD06_IF定義書_IF001_出荷指示データ.pdf","2"') | Set-Content -LiteralPath $pages -Encoding UTF8
    # 設計で入ったバグ: SCR002（指摘が少ない）・SCR003（指摘は基準内）・受注一覧（機能の名前で SCR001）。製造で入ったバグ・取下げは数えない
    @('障害番号,障害内容,機能,混入工程,状態',
        'B001,登録の必須チェックが無い,SCR002,詳細設計,完了',
        'B002,登録の桁あふれ,SCR002,DD,完了',
        'B003,詳細の表示項目の抜け,SCR003,基本設計レビュー,完了',
        'B004,一覧のソート順,受注一覧,詳細設計,完了',
        'B005,コーディングミス,SCR001,製造,完了',
        'B006,操作ミス,SCR003,詳細設計,取下げ',
        'B007,どこかの設計誤り,SCR999,詳細設計,完了') | Set-Content -LiteralPath (Join-Path $Dir 'bugs.csv') -Encoding UTF8
    return $pages
}

Test-Case 'Measure-ReviewEffectiveness: 設計書ごとに指摘密度と設計で入ったバグを並べ、レビューの見逃しを見つける' {
    $dir = New-TestDir
    $pages = New-EffectData $dir
    $out = Join-Path $dir 'effect.xlsx'
    $options = @{ Findings = (Join-Path $dir '指摘管理表.xlsx'); BugList = (Join-Path $dir 'bugs.csv'); SizeCsv = $pages; AsOf = '2026-10-20' }
    $r = @(& $EffectTool @options -OutXlsx $out 6>$null)
    $got = @($r | Sort-Object Document | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}' -f $_.Document, $_.Findings, $_.DesignBugs, $_.BugDensity, $_.RemovalRate, $_.Zone })
    Assert-Equal @(
        'DD02_画面設計書_SCR001_受注一覧|4|1|0.25|0.8|レビューの見逃しの疑い',
        'DD02_画面設計書_SCR002_受注登録|1|2|0.2|0.333|レビューの見逃し',
        'DD02_画面設計書_SCR003_受注詳細|4|1|0.25|0.8|レビューの見逃しの疑い',
        'DD06_IF定義書_IF001_出荷指示データ|2|0|0|1|良好') $got '指摘密度の判定（基準値 0.5〜2.0）× 設計で入ったバグの有無'
    Assert-True (@($r | Where-Object NeedsAction).Count -eq 3)

    $text = Get-HostText { & $EffectTool @options 3>$null | Out-Null }
    Assert-Match $text '設計書 4 件・設計で入ったバグ 5 件（設計書に対応付けたもの 4 件）'
    Assert-Match $text '設計工程の摘出率（レビューの指摘 ÷ \(指摘 \+ 設計で入ったバグ\)）: 73\.3%（指摘 11 件・バグ 4 件）'
    Assert-Match $text '設計書に対応付けられなかったバグ: B007（SCR999）'

    $book = Read-XlsxSheet -Path $out
    Assert-Equal @('設計書別', '設計で入ったバグ', '読み取り結果') $book.Sheets
    $charts = @(Read-XlsxCharts $out)
    Assert-Equal '指摘密度 × 設計で入ったバグの密度' $charts[0].Title

    # 機能の対応表で、障害管理表の SCR999 を IF001 の設計書に対応付ける
    @('機能,別名,設計書', 'IF001,SCR999,DD06_IF定義書_IF001_出荷指示データ') | Set-Content -LiteralPath (Join-Path $dir 'map.csv') -Encoding UTF8
    $r = @(& $EffectTool @options -FunctionMap (Join-Path $dir 'map.csv') 6>$null)
    Assert-Equal '1|レビューの見逃しの疑い' ('{0}|{1}' -f @($r | Where-Object { $_.Document -like 'DD06*' })[0].DesignBugs, @($r | Where-Object { $_.Document -like 'DD06*' })[0].Zone)
}

Test-Case 'Measure-ReviewEffectiveness: Get-Help でヘルプが表示される' {
    $h = Get-Help $EffectTool -Full
    Assert-Match $h.Synopsis 'レビューで見逃した'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
