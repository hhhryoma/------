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
    Assert-Match @($rows | Where-Object { $_.'指標' -eq 'バグ密度' })[0].'備考' '^平均 ± 1σ（2 件。.*）。実績が少ないため参考値$'

    $p = @(& $LimitsTool $csv -Method Percentile 6>$null | Where-Object Metric -eq '指摘密度')[0]
    Assert-Equal '0.875|1.625' ('{0}|{1}' -f $p.Low, $p.High)
    $text = Get-HostText { & $LimitsTool $csv | Out-Null }
    Assert-Match $text '実績 6 件（プロジェクト 4）から、基準値 2 組を求めました（平均 ± 1σ）'
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

Test-Case 'New-QualityLimits: Get-Help でヘルプが表示される' {
    $h = Get-Help $LimitsTool
    Assert-Match $h.Synopsis '基準値'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
