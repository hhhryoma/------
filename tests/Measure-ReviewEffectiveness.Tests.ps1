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

Test-Case 'Measure-ReviewEffectiveness: 設計書の欄が設計書の種類（画面設計書）や「－」「なし」のバグも機能で対応付け、ほかの機能の設計書には対応付けない' {
    $dir = New-TestDir
    @('設計書,指摘内容,分類,状態',
        'DD02_画面設計書_SCR001_受注一覧,桁数が違う,誤り,完了',
        'DD02_画面設計書_SCR003_受注詳細,型が違う,誤り,完了',
        'DD06_IF定義書_IF001_出荷指示データ,項目の抜け,漏れ,完了') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    # B1 の「画面設計書」は 2 つの設計書に当たるので、機能（SCR001）で決める。B2 の IF002 はレビューした設計書が無い（「IF定義書」だけで IF001 にしない）
    @('障害番号,障害内容,機能,設計書,混入工程,状態',
        'B1,一覧のソート順,SCR001,画面設計書,詳細設計,完了',
        'B2,出荷日の桁,IF002,IF定義書,詳細設計,完了',
        'B3,出荷指示の項目の抜け,IF001,－,詳細設計,完了',
        'B4,詳細の表示項目の抜け,SCR003,,詳細設計,完了',
        'B5,詳細の初期表示,SCR003,なし,詳細設計,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    $text = Get-HostText { $script:effect = @(& $EffectTool -Findings (Join-Path $dir 'findings.csv') -BugList (Join-Path $dir 'bugs.csv') -AsOf '2026-10-20') }
    Assert-Equal @('DD02_画面設計書_SCR001_受注一覧|B1', 'DD02_画面設計書_SCR003_受注詳細|B4、B5', 'DD06_IF定義書_IF001_出荷指示データ|B3') @($script:effect | Sort-Object Document | ForEach-Object { '{0}|{1}' -f $_.Document, $_.BugIds })
    Assert-Match $text '設計書に対応付けられなかったバグ: B2（IF002・IF定義書）'
}

Test-Case 'Measure-ReviewEffectiveness: 全体の摘出率に未レビューの設計書のバグを入れず、混入工程を工程として読めないバグは知らせる' {
    $dir = New-TestDir
    @('設計書,指摘内容,分類,状態', 'SCR001_受注一覧,桁数が違う,誤り,完了', 'SCR001_受注一覧,型が違う,誤り,完了', 'SCR001_受注一覧,イベントの抜け,漏れ,完了', 'SCR001_受注一覧,遷移先の誤り,誤り,完了') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    @('ファイル,ページ数', 'SCR001_受注一覧.pdf,4', 'SCR002_受注登録.pdf,4') | Set-Content -LiteralPath (Join-Path $dir 'pages.csv') -Encoding UTF8
    @('設計書,工数', 'SCR001_受注一覧,2') | Set-Content -LiteralPath (Join-Path $dir 'reviewlog.csv') -Encoding UTF8
    # B2 はレビューしていない SCR002 のバグ。B3・B4 の混入工程（設計・UI設計）は工程として読めない
    @('障害番号,障害内容,設計書,混入工程,状態',
        'B1,一覧のソート順,SCR001_受注一覧,詳細設計,完了',
        'B2,登録の桁あふれ,SCR002_受注登録,詳細設計,完了',
        'B3,一覧の表示項目の抜け,SCR001_受注一覧,設計,完了',
        'B4,一覧の配色,SCR001_受注一覧,UI設計,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    $options = @{ Findings = (Join-Path $dir 'findings.csv'); BugList = (Join-Path $dir 'bugs.csv'); SizeCsv = (Join-Path $dir 'pages.csv'); ReviewLog = (Join-Path $dir 'reviewlog.csv'); AsOf = '2026-10-20' }
    $text = Get-HostText { $script:effect = @(& $EffectTool @options -WarningVariable script:effectWarnings -WarningAction SilentlyContinue) }
    Assert-Equal @('SCR001_受注一覧|1|レビューの見逃しの疑い', 'SCR002_受注登録.pdf|1|未レビュー（設計で入ったバグあり）') @($script:effect | Sort-Object Document | ForEach-Object { '{0}|{1}|{2}' -f $_.Document, $_.DesignBugs, $_.Zone })
    Assert-Match $text '設計工程の摘出率（レビューの指摘 ÷ \(指摘 \+ 設計で入ったバグ\)）: 80\.0%（指摘 4 件・バグ 1 件）' '未レビューの設計書（SCR002）のバグ B2 は全体に入れない'
    $unread = @($script:effectWarnings | Where-Object { "$_" -match '混入工程を工程として読めないバグ 2 件' })
    Assert-True ($unread.Count -eq 1 -and "$($unread[0])" -match ': (設計、UI設計|UI設計、設計)$') ('読めない混入工程を知らせる: ' + (@($script:effectWarnings) -join ' | '))
}

Test-Case 'Measure-ReviewEffectiveness: 指摘管理表に設計書の列が無い（設計書ごとのレビュー記録）ときは、設計書の無い指摘の件数とファイルを知らせ、その指摘にゾーンを付けない' {
    $dir = New-TestDir
    $review = Join-Path $dir 'レビュー'
    New-Item -ItemType Directory -Path $review | Out-Null
    foreach ($name in @('SCR001_受注一覧', 'SCR002_受注登録')) {
        New-SheetsBook (Join-Path $review "レビュー記録_$name.xlsx") ([ordered]@{ '指摘' = @(@('指摘内容', '分類', '状態'), @('桁数が違う', '誤り', '完了'), @('型が違う', '誤り', '完了')) })
    }
    @('ファイル,ページ数', 'SCR001_受注一覧.pdf,4', 'SCR002_受注登録.pdf,4') | Set-Content -LiteralPath (Join-Path $dir 'pages.csv') -Encoding UTF8
    # B2 の「設計書なし」は、設計書の無い指摘のまとまり（（設計書なし））に対応付けない
    @('障害番号,障害内容,機能,設計書,混入工程,状態', 'B1,一覧のソート順,SCR001,,詳細設計,完了', 'B2,どこかの設計誤り,,設計書なし,詳細設計,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    $text = Get-HostText { $script:effect = @(& $EffectTool -Findings $review -BugList (Join-Path $dir 'bugs.csv') -SizeCsv (Join-Path $dir 'pages.csv') -AsOf '2026-10-20' -WarningVariable script:effectWarnings -WarningAction SilentlyContinue) }
    $noDoc = @($script:effect | Where-Object Document -eq '（設計書なし）')
    Assert-Equal '4|0||False' ('{0}|{1}|{2}|{3}' -f $noDoc[0].Findings, $noDoc[0].DesignBugs, $noDoc[0].Zone, $noDoc[0].NeedsAction) '設計書の無い指摘にはバグを対応付けず、ゾーンを付けない'
    Assert-Match $text '設計書に対応付けられなかったバグ: B2（設計書なし）'
    $warning = @($script:effectWarnings | Where-Object { "$_" -match '設計書の無い指摘が 4 件' })
    Assert-True ($warning.Count -eq 1) ('設計書の無い指摘を知らせる: ' + (@($script:effectWarnings) -join ' | '))
    Assert-Match "$($warning[0])" '設計書の列を足して'
    Assert-Match "$($warning[0])" 'レビュー記録_SCR001_受注一覧\.xlsx'
    Assert-Match "$($warning[0])" 'レビュー記録_SCR002_受注登録\.xlsx'
}

Test-Case 'Measure-ReviewEffectiveness: 版を上げて設計書の名前の版数が変わっても、指摘・規模・バグ・機能の対応表を同じ設計書に対応付ける' {
    $dir = New-TestDir
    # 指摘はレビューした v1.0 の名前、規模（ページ数）は版を上げたあとの v1.1（v1.0 の行も残っている）、バグは v1.1 の名前と機能
    @('設計書,指摘内容,分類,状態',
        'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx,桁数が違う,誤り,完了',
        'DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx,型が違う,誤り,完了',
        'DD02_画面設計書_SCR002_受注登録_v1.0.xlsx,必須の誤り,誤り,完了') | Set-Content -LiteralPath (Join-Path $dir 'findings.csv') -Encoding UTF8
    $pages = Join-Path $dir 'pages.csv'
    @('"ファイル","ページ数"', '"DD02_画面設計書_SCR001_受注一覧_v1.0.pdf","3"', '"DD02_画面設計書_SCR001_受注一覧_v1.1.pdf","4"', '"DD02_画面設計書_SCR002_受注登録_v1.1.pdf","10"') | Set-Content -LiteralPath $pages -Encoding UTF8
    @('障害番号,障害内容,機能,設計書,混入工程,状態',
        'B1,登録の必須チェックが無い,,DD02_画面設計書_SCR002_受注登録_v1.1.xlsx,詳細設計,完了',
        'B2,一覧のソート順,一覧画面,,詳細設計,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    # 機能の対応表の設計書は版数を付けずに書く
    @('機能,別名,設計書', 'SCR001,一覧画面,DD02_画面設計書_SCR001_受注一覧') | Set-Content -LiteralPath (Join-Path $dir 'map.csv') -Encoding UTF8
    $text = Get-HostText { $script:effect = @(& $EffectTool -Findings (Join-Path $dir 'findings.csv') -BugList (Join-Path $dir 'bugs.csv') -SizeCsv $pages -FunctionMap (Join-Path $dir 'map.csv') -AsOf '2026-10-20' 3>$null) }
    Assert-Equal @('DD02_画面設計書_SCR001_受注一覧_v1.0.xlsx|2|1|4|B2', 'DD02_画面設計書_SCR002_受注登録_v1.0.xlsx|1|1|10|B1') @($script:effect | Sort-Object Document | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Document, $_.Findings, $_.DesignBugs, $_.Size, $_.BugIds })
    Assert-Match $text '設計書 2 件・設計で入ったバグ 2 件（設計書に対応付けたもの 2 件）'
    Assert-NotMatch $text '対応付けられなかったバグ'
}

Test-Case 'Measure-ReviewEffectiveness: Get-Help でヘルプが表示される' {
    $h = Get-Help $EffectTool -Full
    Assert-Match $h.Synopsis 'レビューで見逃した'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
