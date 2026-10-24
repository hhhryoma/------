#     Measure-TestQuality.ps1 のテスト。テスト仕様書・障害管理表はテストの中で作る（Export-Workbook）

$TestQualityTool = Join-Path $script:ToolsDir 'Measure-TestQuality.ps1'

function New-SheetBook {
    param([string] $FilePath, $Sheets)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $FilePath -Sheets @($Sheets | ForEach-Object {
            @{ Name = $_.Name; Columns = @($_.Header | ForEach-Object { @{ Header = $_; Width = 12 } }); Rows = @($_.Rows | ForEach-Object { @{ Values = $_ } }) }
        })
}

function New-TestData {
    #     テスト仕様書（200 ケース、10/01〜10/20 に 1 日 10 件）と、48 件のバグが S 字に収束する障害管理表、機能ごとの KLOC
    param([string] $Dir)
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le 200; $i++) {
        $function = ''
        if ($i -eq 1) { $function = 'SCR001' } elseif ($i -eq 121) { $function = 'SCR002' } elseif ($i -eq 181) { $function = 'SCR003' }
        $planned = (Get-Date '2026-10-01').AddDays([math]::Floor(($i - 1) / 10)).ToString('yyyy/MM/dd')
        $done = $planned; $result = 'OK'; $bug = ''
        if ($i -gt 190) { $done = ''; $result = $(if ($i -le 195) { '保留' } else { '' }) }
        switch ($i) {
            5 { $result = 'NG'; $bug = 'B001' }     # バグは完了しているのに NG のまま
            15 { $result = 'NG'; $bug = 'B047' }
            25 { $result = 'NG' }                   # バグ番号が無い
            35 { $result = 'NG'; $bug = 'B999' }    # 障害管理表に無い
            50 { $result = '対象外' }
            60 { $done = '' }                       # 実施日が無い
        }
        $rows.Add(@([string]$i, $function, "確認 $i", $planned, $done, $result, $bug))
    }
    New-SheetBook (Join-Path $Dir 'UT_受注.xlsx') @(@{ Name = 'テストケース'; Header = @('No', '機能', '確認内容', '予定日', '実施日', '結果', '障害番号'); Rows = $rows })

    # バグの累計（日ごと）: 1,3,6,11,17,24,30,35,39,42,44,45,46,46,47,47,47,48,48,48
    $daily = @(1, 2, 3, 5, 6, 7, 6, 5, 4, 3, 2, 1, 1, 0, 1, 0, 0, 1, 0, 0)
    $bugRows = New-Object System.Collections.Generic.List[object]
    $n = 0
    for ($d = 0; $d -lt $daily.Count; $d++) {
        for ($j = 0; $j -lt $daily[$d]; $j++) {
            $n++
            $function = $(if ($n -le 30) { 'SCR001' } elseif ($n -le 45) { 'SCR002' } else { 'SCR003' })
            $found = (Get-Date '2026-10-01').AddDays($d)
            $closed = ''; $status = '対応中'
            if ($n -le 45) { $closed = $found.AddDays(2).ToString('yyyy/MM/dd'); $status = '完了' }
            $bugRows.Add(@(('B{0:000}' -f $n), $function, "障害 $n", @('高', '中', '低')[$n % 3], @('設計誤り', 'コーディングミス')[$n % 2], @('詳細設計', '製造')[$n % 2], '単体テスト', $found.ToString('yyyy/MM/dd'), $closed, $status))
        }
    }
    $bugRows.Add(@('B049', 'SCR001', '仕様の誤解', '低', '', '', '単体テスト', '2026/10/05', '', '取下げ'))
    New-SheetBook (Join-Path $Dir 'バグ管理表.xlsx') @(@{ Name = '障害'; Header = @('障害番号', '発生機能', '障害内容', '重要度', '原因分類', '混入工程', '摘出工程', '起票日', '完了日', '状態'); Rows = $bugRows })

    $size = Join-Path $Dir 'kloc.csv'
    @('機能,KLOC', 'SCR001,2.0', 'SCR002,1.5', 'SCR003,1.0') | Set-Content -LiteralPath $size -Encoding UTF8
    return $size
}

Test-Case 'Measure-TestQuality: 機能ごとのテスト密度・バグ密度とゾーン、信頼度成長曲線で収束を判定し、食い違いを見つける' {
    $dir = New-TestDir
    $size = New-TestData $dir
    $out = Join-Path $dir 'test-quality.xlsx'
    $options = @{ Path = (Join-Path $dir 'UT_受注.xlsx'); BugList = (Join-Path $dir 'バグ管理表.xlsx'); SizeCsv = $size; AsOf = '2026-10-25' }
    $r = @(& $TestQualityTool @options -OutXlsx $out 6>$null)
    Assert-Equal @(
        'SCR001|119|119|4|30|0|59.5|15|良好',
        'SCR002|60|60|0|15|0|40|10|良好',
        'SCR003|20|10|0|3|3|20|3|テスト不足') @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $_.Function, $_.Cases, $_.Executed, $_.NG, $_.Bugs, $_.OpenBugs, $_.TestDensity, $_.BugDensity, $_.Zone }) '機能は上の行から引き継ぐ。対象外は数えない'
    Assert-True $r[2].NeedsAction
    Assert-Equal '5|5|0.5' ('{0}|{1}|{2}' -f $r[2].Hold, $r[2].NotRun, $r[2].Progress)

    $text = Get-HostText { & $TestQualityTool @options | Out-Null }
    Assert-Match $text 'テストケース 199 件（実施 189・OK 185・NG 4・保留 5・未実施 5。対象外 1）消化率 95\.0%（2026/10/25 時点）'
    Assert-Match $text 'バグ 48 件（未完了 3 件。取下げなどで数えなかったもの 1 件）'
    Assert-Match $text '信頼度成長曲線（横軸: 実施したケースの数）: 潜在バグ数 4[89](\.\d)? 件（推定）・収束率 9\d\.\d% → 収束している（基準 90%）'
    Assert-Match $text 'ゾーン: 良好 2・テスト不足 1'
    Assert-Match $text '要対応: SCR003（テスト不足）テストの観点を見直し、テストケースを追加する'
    Assert-Match $text '食い違い: 14 件（予定日を過ぎても未実施 10・NG なのにバグ番号が無い 1・再テストの結果が無い 1・実施日が無い 1・障害管理表に無いバグ番号 1）'

    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('サマリー', '機能別', 'ゾーン分析', '推移', 'バグの内訳', '食い違い', '読み取り結果') $book.Sheets
    $summary = @{}
    foreach ($row in $book.Rows.Values) { $summary[$row['A']] = $row['B'] }
    Assert-Equal '0.95' ([string]$summary['消化率'])
    Assert-Match $summary['収束の判定'] '^収束している（基準: 収束率 0\.9）$'
    $zone = Read-XlsxSheet -Path $out -Index 3
    Assert-Match $zone.Rows[3]['C'] '^【良好】\nSCR001（テスト 59\.5 / バグ 15）\nSCR002（テスト 40 / バグ 10）$'
    Assert-Match $zone.Rows[4]['B'] '^【テスト不足】\nSCR003（テスト 20 / バグ 3）$'
    $trend = Read-XlsxSheet -Path $out -Index 4
    Assert-Equal '2026/10/01|10|10|9|1|1' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $trend.Rows[2]['A'], $trend.Rows[2]['B'], $trend.Rows[2]['C'], $trend.Rows[2]['D'], $trend.Rows[2]['E'], $trend.Rows[2]['G'])
    Assert-True ([double]$trend.Rows[21]['J'] -gt 46) '信頼度成長曲線の推定値'
    $issueSheet = Read-XlsxSheet -Path $out -Index 6
    Assert-Equal 15 $issueSheet.Rows.Count '見出しと食い違い 14 件'
    Assert-True (@($issueSheet.Links | Where-Object { $_.Location -eq "'テストケース'!A26" }).Count -eq 1) 'NG なのにバグ番号が無いケース（No 25）へのリンク'
    $charts = @(Read-XlsxCharts $out)
    Assert-Equal @('テスト密度 × バグ密度', 'テストの消化（累計）', 'バグの累計と信頼度成長曲線') @($charts | ForEach-Object { $_.Title })
    Assert-Equal @("'機能別'!`$M`$2", "'機能別'!`$O`$2", "'機能別'!`$M`$3", "'機能別'!`$O`$3", "'機能別'!`$M`$4", "'機能別'!`$O`$4") $charts[0].Refs '機能ごとに 1 点（系列の名前を点に表示）'
    Assert-True ($charts[2].Refs -contains "'推移'!`$J`$2:`$J`$21") '信頼度成長曲線の推定値'
    $readSheet = Read-XlsxSheet -Path $out -Index 7
    $caseRow = @($readSheet.Rows.Values | Where-Object { $_['A'] -eq 'テスト仕様書' })[0]
    Assert-Equal 'UT_受注.xlsx|テストケース|1|200' ('{0}|{1}|{2}|{3}' -f $caseRow['B'], $caseRow['C'], $caseRow['D'], $caseRow['E']) '読み取り結果: ファイル・シート・見出しの行・読んだ行数'
    Assert-Match $caseRow['F'] 'テスト内容（確認内容）'
    Assert-Match $caseRow['G'] '工程'
    Assert-Match $text '読み取り（テスト仕様書）: 1 ファイル・200 行。見出しが無い項目: .*工程（1/1 表）'
    $breakdown = Read-XlsxSheet -Path $out -Index 5
    Assert-True (@($breakdown.Rows.Values | Where-Object { $_['A'] -eq '混入工程' -and $_['B'] -eq '詳細設計' -and $_['C'] -eq '24' }).Count -eq 1) '混入工程の内訳'
}

Test-Case 'Measure-TestQuality: シートごとの機能・同じフォルダの障害管理表を読み分け、増え続けるバグは収束していないと判定する' {
    $dir = New-TestDir
    $sheets = @()
    foreach ($name in @('SCR101', 'SCR102')) {
        $rows = @(1..50 | ForEach-Object { , @([string]$_, "確認 $_", (Get-Date '2026-11-01').AddDays([math]::Floor(($_ - 1) / 5)).ToString('yyyy/MM/dd'), 'OK') })
        $sheets += @{ Name = $name; Header = @('項番', 'テスト項目', '実施日', '判定'); Rows = $rows }
    }
    New-SheetBook (Join-Path $dir 'IT.xlsx') $sheets
    # 機能の列が無く、1 日 2 件ずつ増え続ける障害管理表（「No」「確認結果」があってもテスト仕様書としては読まない）
    $lines = @('No,件名,起票日,状態,確認結果')
    for ($i = 1; $i -le 20; $i++) { $lines += ('{0},障害 {0},{1},対応中,未確認' -f $i, (Get-Date '2026-11-01').AddDays([math]::Floor(($i - 1) / 2)).ToString('yyyy/MM/dd')) }
    $lines | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8

    $warnings = $null
    $r = @(& $TestQualityTool -Path $dir -BugList (Join-Path $dir 'bugs.csv') -AsOf '2026-11-20' -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @('SCR101|50|50', 'SCR102|50|50') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Function, $_.Cases, $_.Executed })
    Assert-True (@($warnings | Where-Object { "$_" -match '機能の列が無い' }).Count -eq 1) '機能の列が無い障害管理表を知らせる'
    $text = Get-HostText { & $TestQualityTool -Path $dir -BugList (Join-Path $dir 'bugs.csv') -AsOf '2026-11-20' 3>$null | Out-Null }
    Assert-Match $text 'バグ 20 件（未完了 20 件'
    Assert-Match $text '収束していない'
}

Test-Case 'Measure-TestQuality: バグが少ないうち・消化が進まないうちは収束を判定せず、直近の新規バグを並べる' {
    $dir = New-TestDir
    $rows = @(1..40 | ForEach-Object {
            $done = ''; $result = ''
            if ($_ -le 30) { $done = (Get-Date '2026-11-01').AddDays([math]::Floor(($_ - 1) / 5)).ToString('yyyy/MM/dd'); $result = 'OK' }
            , @([string]$_, "確認 $_", $done, $result) })
    New-SheetBook (Join-Path $dir 'UT.xlsx') @(@{ Name = 'SCR201'; Header = @('項番', 'テスト項目', '実施日', '判定'); Rows = $rows })
    $lines = @('No,件名,起票日,状態')
    foreach ($d in @('2026-11-01', '2026-11-02', '2026-11-03', '2026-11-04', '2026-11-05', '2026-11-06')) { $lines += ('{0},障害 {0},{0},完了' -f $d) }
    $lines | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    $options = @{ Path = (Join-Path $dir 'UT.xlsx'); BugList = (Join-Path $dir 'bugs.csv'); AsOf = '2026-11-06' }

    $text = Get-HostText { & $TestQualityTool @options 3>$null | Out-Null }
    Assert-Match $text 'まだ判定できない（バグ 6 件・消化率 75%。バグ 10 件以上・消化率 50% 以上で判定します）'
    # 直近 3 日（11/04〜11/06）と、その前の 3 日（11/01〜11/03）
    $text = Get-HostText { & $TestQualityTool @options -RecentDays 3 3>$null | Out-Null }
    Assert-Match $text '直近 3 日の新規バグ: 3 件（その前の 3 日: 3 件）'
    $text = Get-HostText { & $TestQualityTool @options -MinBugs 5 -MinProgress 0.9 3>$null | Out-Null }
    Assert-Match $text 'まだ判定できない（バグ 6 件・消化率 75%。バグ 5 件以上・消化率 90% 以上で判定します）'
    $text = Get-HostText { & $TestQualityTool @options -MinBugs 5 3>$null | Out-Null }
    Assert-NotMatch $text 'まだ判定できない'
}

Test-Case 'Measure-TestQuality: -DesignPath では画面・IF などだけを機能とし、FP を分母にする。機能の対応表で書き方をそろえる' {
    $dir = New-TestDir
    $design = Join-Path $script:RepoRoot 'docs/design-sample'
    New-SheetBook (Join-Path $dir 'UT.xlsx') @(@{ Name = 'ケース'; Header = @('No', '機能', '確認内容', '結果'); Rows = @(
                , @('1', 'SCR001', '初期表示', 'OK')
                , @('2', '受注登録', '登録する', 'OK')
                , @('3', '受注登録画面', '登録の確認', 'OK')) })
    @('No,件名,機能,起票日,状態', '1,登録できない,受注登録,2026/10/01,完了') | Set-Content -LiteralPath (Join-Path $dir 'bugs.csv') -Encoding UTF8
    @('機能,別名,設計書', 'SCR002,受注登録|受注登録画面,DD02_画面設計書_SCR002_受注登録') | Set-Content -LiteralPath (Join-Path $dir 'map.csv') -Encoding UTF8
    $r = @(& $TestQualityTool -Path (Join-Path $dir 'UT.xlsx') -BugList (Join-Path $dir 'bugs.csv') -DesignPath $design -FunctionMap (Join-Path $dir 'map.csv') -AsOf '2026-10-20' 3>$null 6>$null)
    $got = @($r | ForEach-Object { '{0}|{1}|{2}|{3}|{4}' -f $_.Function, $_.Cases, $_.Bugs, $_.Size, $_.Unit })
    # テーブル定義書・メッセージ一覧・コード定義書は機能にしない。SCR002 は対応表で「受注登録」「受注登録画面」と設計書をそろえる
    # FP = 4 × EI + 5 × EO + 4 × EQ（SCR001: EQ 2 → 8、SCR002: EI 1・EQ 2 → 12、IF001: EO 1 → 5）
    Assert-Equal @('SCR001|1|0|8|FP', 'SCR002|2|1|12|FP', 'SCR003 受注詳細|0|0|8|FP', 'IF001 出荷指示データ|0|0|5|FP') $got
    $r = @(& $TestQualityTool -Path (Join-Path $dir 'UT.xlsx') -DesignPath $design -DesignUnit '規模' -AsOf '2026-10-20' 3>$null 6>$null)
    Assert-Equal '17|規模' ('{0}|{1}' -f @($r | Where-Object Function -eq 'SCR001')[0].Size, @($r | Where-Object Function -eq 'SCR001')[0].Unit)
}

Test-Case 'Measure-TestQuality: ケースの数が同じ工程は、先に出てきた工程の基準値を使う（PowerShell の版によらない）' {
    $dir = New-TestDir
    # 先に出てくる「結合テスト」は、名前の順では「単体テスト」のあと
    New-SheetBook (Join-Path $dir 'IT_受注.xlsx') @(@{ Name = 'テストケース'; Header = @('No', '機能', '確認内容', '工程', '結果'); Rows = @(
                @('1', 'SCR001', '確認 1', '結合テスト', 'OK'),
                @('2', 'SCR001', '確認 2', '単体テスト', 'OK')) })
    $size = Join-Path $dir 'kloc.csv'
    @('機能,KLOC', 'SCR001,1.0') | Set-Content -LiteralPath $size -Encoding UTF8
    $limits = Join-Path $dir 'limits.csv'
    @('指標,分母,工程,下限,上限', 'テスト密度,KLOC,結合テスト,1,3', 'テスト密度,KLOC,単体テスト,100,200') | Set-Content -LiteralPath $limits -Encoding UTF8
    $r = @(& $TestQualityTool -Path (Join-Path $dir 'IT_受注.xlsx') -SizeCsv $size -Limits $limits -AsOf '2026-10-25' 6>$null 3>$null | Where-Object { $_.PSObject.TypeNames -contains 'TestQuality.Function' })
    Assert-Equal 'SCR001|結合テスト|2|基準内' ('{0}|{1}|{2}|{3}' -f $r[0].Function, $r[0].Process, $r[0].TestDensity, $r[0].TestJudgement) '結合テストの基準値（1〜3）と比べる'
}

Test-Case 'Measure-TestQuality: Get-Help でヘルプが表示される' {
    $h = Get-Help $TestQualityTool
    Assert-Match $h.Synopsis 'テスト'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
