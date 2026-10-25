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

Test-Case 'New-QualityLimits: Get-Help でヘルプが表示される' {
    $h = Get-Help $LimitsTool
    Assert-Match $h.Synopsis '基準値'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
