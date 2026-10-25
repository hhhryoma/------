#     New-QualityLimits.ps1 のテスト。実績の CSV はテストの中で作る

$LimitsTool = Join-Path $script:ToolsDir 'New-QualityLimits.ps1'

function New-ActualsCsv {
    param([string] $Dir)
    $csv = Join-Path $Dir '実績.csv'
    @('プロジェクト,指標,分母,工程,値',
        'A,指摘密度,ページ数,詳細設計,0.5',
        'B,指摘密度,ページ数,詳細設計,1.0',
        'C,指摘密度,ページ数,詳細設計,1.5',
        'D,指摘密度,ページ数,詳細設計,2.0',
        'A,バグ密度,KLOC,,10',
        'B,バグ密度,KLOC,,12',
        'C,バグ密度,KLOC,,-') | Set-Content -LiteralPath $csv -Encoding UTF8
    return $csv
}

Test-Case 'New-QualityLimits: 指標 × 分母 × 工程ごとに、平均 ± 標準偏差（またはパーセント点）で基準値を求める' {
    $dir = New-TestDir
    $csv = New-ActualsCsv $dir
    $out = Join-Path $dir '基準値.csv'
    $r = @(& $LimitsTool $csv -OutCsv $out 6>$null)
    $review = @($r | Where-Object Metric -eq '指摘密度')[0]
    Assert-Equal '詳細設計|4|1.25|0.605|1.895' ('{0}|{1}|{2}|{3}|{4}' -f $review.Process, $review.Count, $review.Mean, $review.Low, $review.High)
    $bug = @($r | Where-Object Metric -eq 'バグ密度')[0]
    Assert-Equal '2|11|9.586|12.414|True' ('{0}|{1}|{2}|{3}|{4}' -f $bug.Count, $bug.Mean, $bug.Low, $bug.High, $bug.Few) '数値でない値は数えない'
    $rows = @(Import-Csv -LiteralPath $out -Encoding UTF8)
    Assert-Equal @('指標', '分母', '工程', '下限', '上限', '備考') @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    Assert-Match $bug.Note '^平均 ± 1σ（2 件。.*）。実績が少ないため参考値$'
    Assert-Equal @('指摘密度') @($rows | ForEach-Object { $_.'指標' }) '実績が少ない（参考値の）バグ密度は基準値の CSV に書かない'

    $p = @(& $LimitsTool $csv -Method Percentile 6>$null | Where-Object Metric -eq '指摘密度')[0]
    Assert-Equal '0.875|1.625' ('{0}|{1}' -f $p.Low, $p.High)
    $text = Get-HostText { & $LimitsTool $csv | Out-Null }
    Assert-Match $text '実績 6 件（プロジェクト 4）から、基準値 2 組を求めました（平均 ± 1σ）'
    Assert-Match $text '参考値の組（実績が 3 件未満か、下限と上限が同じ。基準値の CSV には書きません）: バグ密度/KLOC 9\.586〜12\.414'
}

Test-Case 'New-QualityLimits: 作った基準値の CSV を Measure-ReviewQuality の -Limits にそのまま使える' {
    $dir = New-TestDir
    $out = Join-Path $dir '基準値.csv'
    $null = & $LimitsTool (New-ActualsCsv $dir) -OutCsv $out 6>$null
    $findings = Join-Path $dir 'findings.csv'
    @('設計書,指摘内容,工程', 'DD02_画面設計書_SCR001_受注一覧,誤り1,詳細設計') | Set-Content -LiteralPath $findings -Encoding UTF8
    $pages = Join-Path $dir 'pages.csv'
    @('ファイル,ページ数', 'DD02_画面設計書_SCR001_受注一覧.pdf,4') | Set-Content -LiteralPath $pages -Encoding UTF8
    $r = @(& (Join-Path $script:ToolsDir 'Measure-ReviewQuality.ps1') -Path $findings -SizeCsv $pages -Limits $out 6>$null)
    Assert-Equal '0.25|少ない（レビュー不足の疑い）' ('{0}|{1}' -f $r[0].Density, $r[0].Judgement)
}

Test-Case 'New-QualityLimits: Excel で開いている（書き込みでロックされた）CSV も読める' -WindowsOnly {
    $dir = New-TestDir
    $csv = New-ActualsCsv $dir
    # Excel で開いたときと同じく、書き込みのために開き、ほかには読み取りだけを許す
    $lock = [System.IO.File]::Open($csv, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
    try {
        $r = @(& $LimitsTool $csv 6>$null)
        Assert-Equal @('バグ密度', '指摘密度') @($r | ForEach-Object { $_.Metric } | Sort-Object)
    }
    finally { $lock.Dispose() }
}

Test-Case 'New-QualityLimits: 時点を日付として比べ、Excel で保存した 2026/9/30 の形でも、プロジェクトの最後の時点の値を使う' {
    $dir = New-TestDir
    $csv = Join-Path $dir '実績.csv'
    @('プロジェクト,時点,指標,分母,工程,値', 'A,2026/9/30,指摘密度,ページ数,,0.5', 'A,2026/10/15,指摘密度,ページ数,,1.5', 'B,2026/10/15,指摘密度,ページ数,,1.0') | Set-Content -LiteralPath $csv -Encoding UTF8
    $r = @(& $LimitsTool $csv 6>$null)
    Assert-Equal '2|1.25' ('{0}|{1}' -f $r[0].Count, $r[0].Mean) 'A は最後の時点（2026/10/15）の 1.5、B は 1.0'
}

Test-Case 'New-QualityLimits: 実績が 1 件・値がすべて同じ組は、下限 = 上限の基準値を CSV に書かず（-Limits で判定しない）、参考値として表示する' {
    $dir = New-TestDir
    $csv = Join-Path $dir '実績.csv'
    # 1 つめのプロジェクトが終わったときの実績（New-QualityReport の -HistoryCsv の形）と、3 プロジェクトとも同じ値の指標
    @('プロジェクト,時点,区分,指標,分母,工程,値', 'A,2026/09/30,レビュー,指摘密度,ページ数,,0.5', 'A,2026/09/30,テスト,テスト密度,KLOC,,40', 'B,2026/09/30,テスト,テスト密度,KLOC,,40', 'C,2026/09/30,テスト,テスト密度,KLOC,,40') | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Join-Path $dir '基準値.csv'
    $text = Get-HostText { & $LimitsTool $csv -OutCsv $out | Out-Null }
    Assert-Equal 0 @(Import-Csv -LiteralPath $out -Encoding UTF8).Count '下限 = 上限の基準値を書いている'
    Assert-Match $text '参考値の組.*指摘密度/ページ数 0\.5〜0\.5'
    Assert-Match $text '参考値の組.*テスト密度/KLOC 40〜40'
    $r = @(& $LimitsTool $csv 6>$null | Where-Object Metric -eq 'テスト密度')
    Assert-Equal '3|True' ('{0}|{1}' -f $r[0].Count, $r[0].Few)
    Assert-Match $r[0].Note '。下限と上限が同じため参考値$'

    # 基準値の CSV を -Limits に渡しても、実績とわずかに違う設計書を基準外にしない
    $findings = Join-Path $dir 'findings.csv'
    @('設計書,指摘内容,工程', 'DD02_画面設計書_SCR001_受注一覧,誤り1,詳細設計') | Set-Content -LiteralPath $findings -Encoding UTF8
    $pages = Join-Path $dir 'pages.csv'
    @('ファイル,ページ数', 'DD02_画面設計書_SCR001_受注一覧.pdf,4') | Set-Content -LiteralPath $pages -Encoding UTF8
    $d = @(& (Join-Path $script:ToolsDir 'Measure-ReviewQuality.ps1') -Path $findings -SizeCsv $pages -Limits $out 3>$null 6>$null)
    Assert-Equal '0.25|' ('{0}|{1}' -f $d[0].Density, $d[0].Judgement)
}

Test-Case 'New-QualityLimits: 平均 − σ が 0 以下の組は下限を空にして知らせ（下限 0 にしない）、Measure-TestQuality はテストケースの無い機能を「少ない」にする' {
    $dir = New-TestDir
    $csv = Join-Path $dir '実績.csv'
    # ばらつきが大きい（平均 73.75・標準偏差 84.8）
    @('プロジェクト,時点,区分,指標,分母,工程,値', 'A,2026/03/31,テスト,テスト密度,KLOC,,20', 'B,2026/03/31,テスト,テスト密度,KLOC,,30', 'C,2026/03/31,テスト,テスト密度,KLOC,,45', 'D,2026/03/31,テスト,テスト密度,KLOC,,200') | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Join-Path $dir '基準値.csv'
    $text = Get-HostText { $script:noLowLimits = @(& $LimitsTool $csv -OutCsv $out) }
    Assert-Equal '|158.541' ('{0}|{1}' -f $script:noLowLimits[0].Low, $script:noLowLimits[0].High)
    Assert-Match $text '平均 − 1σ が 0 以下になる組は、下限を空にしました.*-Method Percentile も試してください: テスト密度/KLOC'
    $row = @(Import-Csv -LiteralPath $out -Encoding UTF8)[0]
    Assert-Equal '|158.541' ('{0}|{1}' -f $row.'下限', $row.'上限') '下限を 0 と書かない'
    Assert-Match $row.'備考' '平均 − 1σ が 0 以下のため、下限なし'

    # 規模があるのにテストケースの無い機能（SCR002）は、下限が空でも「少ない」。ケースのある機能は基準内
    @('No,機能,確認内容,結果', '1,SCR001,確認 1,OK', '2,SCR001,確認 2,OK') | Set-Content -LiteralPath (Join-Path $dir 'cases.csv') -Encoding UTF8
    @('機能,KLOC', 'SCR001,0.05', 'SCR002,0.05') | Set-Content -LiteralPath (Join-Path $dir 'kloc.csv') -Encoding UTF8
    $r = @(& (Join-Path $script:ToolsDir 'Measure-TestQuality.ps1') -Path (Join-Path $dir 'cases.csv') -SizeCsv (Join-Path $dir 'kloc.csv') -Limits $out -AsOf '2026-10-20' 3>$null 6>$null | Where-Object { $_.PSObject.TypeNames -contains 'TestQuality.Function' })
    Assert-Equal @('SCR001|40|基準内', 'SCR002|0|少ない') @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.Function, $_.TestDensity, $_.TestJudgement })
}

Test-Case 'New-QualityLimits: 同じプロジェクトの名前がいくつもの CSV にあれば知らせ、出力先のフォルダの名前（品質報告）は CSV ごとに別のプロジェクトとして数える' {
    $dir = New-TestDir
    # 以前の New-QualityReport.ps1 で -Project を省いて作った、2 つの案件の実績（どちらもプロジェクトが「品質報告」）
    @('プロジェクト,時点,区分,指標,分母,工程,値', '品質報告,2026/09/30,レビュー,指摘密度,ページ数,,0.5', '品質報告,2026/10/15,レビュー,指摘密度,ページ数,,1.0') | Set-Content -LiteralPath (Join-Path $dir '実績_案件A.csv') -Encoding UTF8
    @('プロジェクト,時点,区分,指標,分母,工程,値', '品質報告,2026/10/10,レビュー,指摘密度,ページ数,,2.0') | Set-Content -LiteralPath (Join-Path $dir '実績_案件B.csv') -Encoding UTF8
    $r = @(& $LimitsTool $dir -MinSamples 1 -WarningVariable script:limitWarnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal '2|1.5' ('{0}|{1}' -f $r[0].Count, $r[0].Mean) '案件A は最後の時点の 1.0、案件B は 2.0'
    Assert-Match ($script:limitWarnings -join ' ') 'プロジェクト「品質報告」が、いくつもの実績の CSV（実績_案件A\.csv・実績_案件B\.csv）にあります。出力先のフォルダの名前なので、ファイルごとに別のプロジェクトとして数えます'

    # ほかの名前は同じプロジェクトとして数え、知らせる
    foreach ($f in @('実績_案件A.csv', '実績_案件B.csv')) { $p = Join-Path $dir $f; [System.IO.File]::WriteAllText($p, ([System.IO.File]::ReadAllText($p) -replace '品質報告', '受注'), (New-Object System.Text.UTF8Encoding($true))) }
    $r = @(& $LimitsTool $dir -MinSamples 1 -WarningVariable script:limitWarnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal '1|1' ('{0}|{1}' -f $r[0].Count, $r[0].Mean) '受注の最後の時点（10/15）の 1.0'
    Assert-Match ($script:limitWarnings -join ' ') 'プロジェクト「受注」が、いくつもの実績の CSV（実績_案件A\.csv・実績_案件B\.csv）にあります。同じプロジェクトとして'
}

Test-Case 'New-QualityLimits: Get-Help でヘルプが表示される' {
    $h = Get-Help $LimitsTool
    Assert-Match $h.Synopsis '基準値'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
