[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateSet('Sigma', 'Percentile')]
    [string] $Method = 'Sigma',

    [ValidateRange(0.1, 5.0)]
    [double] $Sigma = 1.0,

    [ValidateCount(2, 2)]
    [double[]] $Percentiles = @(25, 75),

    [ValidateRange(1, 1000)]
    [int] $MinSamples = 3,

    [string] $OutCsv
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if ($OutCsv) { $OutCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutCsv) }
if ($Percentiles[0] -ge $Percentiles[1] -or $Percentiles[0] -lt 0 -or $Percentiles[1] -gt 100) { throw '-Percentiles は 0〜100 の小さい方・大きい方の順に指定してください（例: 25, 75）' }

Update-TypeData -TypeName 'QualityLimits.Limit' -DefaultDisplayPropertySet Metric, Unit, Process, Count, Mean, Low, High -Force -WhatIf:$false

$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$headerOf = @{
    '指標' = @('指標', '指標名', 'メトリクス')
    '分母' = @('分母', '単位', '規模の単位')
    '工程' = @('工程', 'フェーズ', 'テスト工程', 'レビュー工程')
    '値'   = @('値', '実績', '実績値', '測定値')
    'プロジェクト' = @('プロジェクト', '案件', 'プロジェクト名', '案件名')
    '時点' = @('時点', '集計日', '日付')
}

# ---------------------------------------------------------------- 実績を読む


$samples = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    foreach ($item in @(Get-PathItem $p)) {
        $files = @($item)
        if ($item.PSIsContainer) { $files = @(Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.csv' -Recurse) }
        foreach ($f in $files) {
            $rows = Read-CsvFile $f.FullName @() '実績の CSV'
            if ($rows.Count -eq 0) { continue }
            $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
            $col = @{}
            foreach ($k in $headerOf.Keys) {
                $labels = @($headerOf[$k] | ForEach-Object { ConvertTo-Label $_ })
                $col[$k] = @($names | Where-Object { $labels -contains (ConvertTo-Label $_) })[0]
            }
            if (-not ($col['指標'] -and $col['分母'] -and $col['値'])) { Write-Warning ("{0}: 指標・分母・値の列が見つかりません" -f $f.Name); continue }
            $line = 1
            foreach ($r in $rows) {
                $line++
                $value = 0.0
                if (-not [double]::TryParse(([string]$r.($col['値'])).Trim(), [System.Globalization.NumberStyles]::Float, $invariant, [ref]$value)) { continue }
                $process = ''
                if ($col['工程']) { $process = ([string]$r.($col['工程'])).Trim() }
                $project = ''
                if ($col['プロジェクト']) { $project = ([string]$r.($col['プロジェクト'])).Trim() }
                $date = ''
                if ($col['時点']) { $date = ([string]$r.($col['時点'])).Trim() }
                $samples.Add([pscustomobject]@{ Metric = ([string]$r.($col['指標'])).Trim(); Unit = ([string]$r.($col['分母'])).Trim(); Process = $process; Value = $value; Project = $project; Date = $date })
            }
        }
    }
}
if ($samples.Count -eq 0) { throw '実績が見つかりません（指標・分母・値の列がある CSV を指定してください）' }

# New-QualityReport.ps1 の実績の CSV のように、同じプロジェクトの値が時点ごとにあるときは、最後の時点の値だけを使う
$latest = @{}
foreach ($x in $samples) {
    if (-not ($x.Project -and $x.Date)) { continue }
    $k = '{0}|{1}|{2}|{3}' -f $x.Project, $x.Metric, $x.Unit, $x.Process
    if (-not $latest.ContainsKey($k) -or [string]::CompareOrdinal($x.Date, $latest[$k]) -gt 0) { $latest[$k] = $x.Date }
}
if ($latest.Count -gt 0) {
    $kept = New-Object System.Collections.Generic.List[object]
    foreach ($x in $samples) {
        if ($x.Project -and $x.Date -and $x.Date -ne $latest[('{0}|{1}|{2}|{3}' -f $x.Project, $x.Metric, $x.Unit, $x.Process)]) { continue }
        $kept.Add($x)
    }
    $samples = $kept
}

# ---------------------------------------------------------------- 基準値を求める


function Get-Percentile {
    #     並べた値の p パーセント点（線形補間）
    param([double[]] $Sorted, [double] $P)
    if ($Sorted.Count -eq 1) { return $Sorted[0] }
    $rank = ($P / 100.0) * ($Sorted.Count - 1)
    $lo = [math]::Floor($rank); $hi = [math]::Ceiling($rank)
    return $Sorted[$lo] + ($Sorted[$hi] - $Sorted[$lo]) * ($rank - $lo)
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($g in @($samples | Group-Object Metric, Unit, Process | Sort-Object Name)) {
    $first = $g.Group[0]
    $values = [double[]]@($g.Group | ForEach-Object { $_.Value })
    [Array]::Sort($values)
    $n = $values.Count
    $mean = ($values | Measure-Object -Average).Average
    $sd = 0.0
    if ($n -gt 1) { $sd = [math]::Sqrt((($values | ForEach-Object { ($_ - $mean) * ($_ - $mean) }) | Measure-Object -Sum).Sum / ($n - 1)) }
    if ($Method -eq 'Sigma') {
        $low = [math]::Max(0.0, $mean - $Sigma * $sd); $high = $mean + $Sigma * $sd
        $basis = '平均 ± {0}σ' -f $Sigma
    }
    else {
        $low = Get-Percentile $values $Percentiles[0]; $high = Get-Percentile $values $Percentiles[1]
        $basis = '{0}〜{1} パーセント点' -f $Percentiles[0], $Percentiles[1]
    }
    $note = '{0}（{1} 件。平均 {2}・標準偏差 {3}・最小 {4}・最大 {5}）' -f $basis, $n, [math]::Round($mean, 3), [math]::Round($sd, 3), $values[0], $values[$n - 1]
    if ($n -lt $MinSamples) { $note += '。実績が少ないため参考値' }
    $r = [pscustomobject]@{
        Metric = $first.Metric; Unit = $first.Unit; Process = $first.Process; Count = $n
        Mean = [math]::Round($mean, 3); StdDev = [math]::Round($sd, 3); Low = [math]::Round($low, 3); High = [math]::Round($high, 3); Note = $note; Few = ($n -lt $MinSamples)
    }
    $r.PSObject.TypeNames.Insert(0, 'QualityLimits.Limit')
    $results.Add($r)
}
$results

if ($OutCsv) {
    # Measure-ReviewQuality.ps1・Measure-TestQuality.ps1 の -Limits にそのまま渡せる形
    $rows = foreach ($r in $results) { [pscustomobject][ordered]@{ '指標' = $r.Metric; '分母' = $r.Unit; '工程' = $r.Process; '下限' = $r.Low.ToString($invariant); '上限' = $r.High.ToString($invariant); '備考' = $r.Note } }
    [System.IO.File]::WriteAllText($OutCsv, (((@($rows) | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Write-Host ''
Write-Host ('実績 {0} 件（プロジェクト {1}）から、基準値 {2} 組を求めました（{3}）' -f $samples.Count, @($samples | Where-Object { $_.Project } | ForEach-Object { $_.Project } | Sort-Object -Unique).Count, $results.Count, $(if ($Method -eq 'Sigma') { "平均 ± ${Sigma}σ" } else { '{0}〜{1} パーセント点' -f $Percentiles[0], $Percentiles[1] }))
$few = @($results | Where-Object Few)
if ($few.Count -gt 0) { Write-Host ('実績が {0} 件未満の組: {1}（参考値です）' -f $MinSamples, ((@($few | ForEach-Object { '{0}/{1}{2}' -f $_.Metric, $_.Unit, $(if ($_.Process) { '/' + $_.Process } else { '' }) })) -join '・')) }
if ($OutCsv) { Write-Host "基準値の CSV を保存しました: $OutCsv" }

# .SYNOPSIS
#     過去のプロジェクトの品質の実績（指摘密度・バグ密度など）から、社内の基準値（下限・上限）の CSV を作ります（Excel 不要）。
#
# .DESCRIPTION
#     実績の CSV（指標・分母・工程・値の列。プロジェクト・時点の列は任意）を読み、指標 × 分母 × 工程の組ごとに下限・上限を求めます。
#     同じプロジェクトの値が時点ごとにあるとき（New-QualityReport.ps1 -HistoryCsv の CSV）は、最後の時点の値だけを使います。
#
#       Sigma（既定）   平均 ± -Sigma × 標準偏差（下限は 0 未満にしない）
#       Percentile      -Percentiles のパーセント点（既定 25〜75。外れ値に引きずられにくい）
#
#     -OutCsv の CSV は、Measure-ReviewQuality.ps1（指摘密度・レビュー工数密度）と Measure-TestQuality.ps1（テスト密度・バグ密度）の
#     -Limits にそのまま渡せます。実績が -MinSamples 件未満の組は、備考に「参考値」と書きます。
#
# .PARAMETER Path
#     実績の CSV のファイルまたはフォルダ（複数可）。列は 指標・分母・工程（任意）・値（実績）・プロジェクト（任意）。
#
# .PARAMETER Method
#     基準値の求め方。Sigma（平均 ± 標準偏差の倍数、既定）か Percentile（パーセント点）。
#
# .PARAMETER Sigma
#     Sigma のときの標準偏差の倍数（既定 1）。
#
# .PARAMETER Percentiles
#     Percentile のときの下限・上限のパーセント点（既定 25, 75）。
#
# .PARAMETER MinSamples
#     これより実績が少ない組は参考値とします（既定 3）。
#
# .PARAMETER OutCsv
#     基準値の CSV（指標・分母・工程・下限・上限・備考。UTF-8・BOM 付き）を保存します。
#
# .EXAMPLE
#     .\New-QualityLimits.ps1 D:\品質実績\実績.csv -OutCsv D:\品質実績\基準値.csv
#     .\Measure-TestQuality.ps1 D:\テスト -BugList D:\障害管理表.xlsx -SizeCsv D:\規模.csv -Limits D:\品質実績\基準値.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1 が必要です。
