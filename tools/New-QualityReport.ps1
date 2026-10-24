[CmdletBinding()]
param(
    [string] $Config,

    [string] $OutDir,

    [string[]] $DesignPath,

    [string[]] $Findings,

    [string[]] $AiReview,

    [string[]] $ReviewLog,

    [string] $PageCsv,

    [string[]] $TestPath,

    [string[]] $BugList,

    [string[]] $Requirement,

    [string] $SizeCsv,

    [string] $FunctionMap,

    [string] $ReviewLimits,

    [string] $TestLimits,

    [string] $Criteria,

    [string] $HistoryCsv,

    [string] $Project,

    [string] $Title = '品質評価',

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
# 日付は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

# 設定ファイル（-Config）の値は、コマンドで指定しなかったパラメーターにだけ使う
if ($Config) {
    $settings = Read-QualityReportConfig $Config
    foreach ($k in $settings.Keys) { if (-not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value $settings[$k] } }
}
if (-not $OutDir) { throw '-OutDir（品質評価の保存先）か、-Config（品質評価の設定ファイル）を指定してください' }
if (-not $Findings -and -not $TestPath -and -not $DesignPath) { throw '-Findings（レビューの指摘管理表）・-TestPath（テスト仕様書）・-DesignPath（設計書）のどれかを指定してください' }
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$detailDir = Join-Path $OutDir '詳細'
New-Item -ItemType Directory -Force -Path $detailDir | Out-Null
$reportPath = Join-Path $OutDir ('{0}_{1}.xlsx' -f $Title, $AsOf.ToString('yyyyMMdd', $invariant))
if (-not $Criteria) { $Criteria = Join-Path $PSScriptRoot 'quality-report/判定基準.csv' }
$Criteria = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Criteria)
if ($HistoryCsv) { $HistoryCsv = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($HistoryCsv) }
if (-not $Project) { $Project = Split-Path -Path $OutDir -Leaf }

# ---------------------------------------------------------------- ヘルパー


$rows = New-Object System.Collections.Generic.List[object]      # 品質評価の行
$actions = New-Object System.Collections.Generic.List[object]   # 要対応の行
$ran = New-Object System.Collections.Generic.List[string]
$failed = New-Object System.Collections.Generic.List[string]

# 判定基準: 指標（* で複数の指標にも書ける）ごとに、向き（以上・以下）と ○・△ のしきい値。△ が空なら × にはしない
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$criteriaRows = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Criteria @('指標', '向き', '○', '△') '判定基準の CSV')) {
    $pattern = ([string]$r.'指標').Trim()
    if (-not $pattern) { continue }
    $direction = ([string]$r.'向き').Trim()
    if (@('以上', '以下') -notcontains $direction) { throw ("判定基準の CSV の向きは「以上」か「以下」にしてください（{0}: {1}）: {2}" -f $pattern, $direction, $Criteria) }
    $limits = @{}
    $percent = $false
    foreach ($k in @('○', '△')) {
        $t = ([string]$r.$k).Trim()
        $limits[$k] = $null
        if (-not $t) { continue }
        if ($t.EndsWith('%')) { $percent = $true }
        $v = 0.0
        if (-not [double]::TryParse($t.TrimEnd('%').Trim(), [System.Globalization.NumberStyles]::Float, $invariant, [ref]$v)) { throw ("判定基準の CSV の {0} が数値ではありません（{1}: {2}）: {3}" -f $k, $pattern, $t, $Criteria) }
        if ($t.EndsWith('%')) { $v = $v / 100.0 }
        $limits[$k] = $v
    }
    if ($null -eq $limits['○']) { throw ("判定基準の CSV の ○ が空です（{0}）: {1}" -f $pattern, $Criteria) }
    $format = { param($v) if ($percent) { '{0}% ' -f [math]::Round($v * 100, 1) } else { '{0} 件' -f $v } }.GetNewClosure()
    if ($direction -eq '以上') {
        $text = '{0}以上なら ○' -f (& $format $limits['○'])
        if ($null -ne $limits['△']) { $text += '、{0}以上なら △、それ未満は ×' -f (& $format $limits['△']) } else { $text += '、それ未満は △' }
    }
    else {
        $text = '{0}以下なら ○' -f (& $format $limits['○'])
        if ($null -ne $limits['△'] -and $limits['△'] -gt $limits['○']) { $text += '、{0}以下なら △、それより多ければ ×' -f (& $format $limits['△']) }
        elseif ($null -ne $limits['△']) { $text += '、それより多ければ ×' }
        else { $text += '、それより多ければ △' }
    }
    $criteriaRows.Add([pscustomobject]@{ Pattern = $pattern; Higher = ($direction -eq '以上'); Good = $limits['○']; Fair = $limits['△']; Text = $text })
}

function Get-Criterion {
    #     指標の名前に合う判定基準（最初に合った行）。無ければ $null
    param([string] $Name)
    foreach ($c in $criteriaRows) { if ($Name -like $c.Pattern) { return $c } }
    return $null
}

function Get-MetricJudgement {
    param($Criterion, [double] $Value)
    $ok = { param($limit) if ($Criterion.Higher) { $Value -ge $limit } else { $Value -le $limit } }
    if (& $ok $Criterion.Good) { return '○' }
    if ($null -eq $Criterion.Fair -or (& $ok $Criterion.Fair)) { return '△' }
    return '×'
}

function Add-Metric {
    #     品質評価の 1 行（区分・指標・実績・判定・判定の考え方・詳細のファイル）。判定は判定基準の CSV で決める
    #       Key・Unit: 実績の CSV（-HistoryCsv）に書く指標の名前と分母（密度は New-QualityLimits.ps1 で基準値にできる名前にする）
    param([string] $Area, [string] $Name, $Value, [string] $Note, [string] $Detail, [string] $Key, [string] $Unit)
    $judgement = ''
    $basis = $Note
    $c = Get-Criterion $Name
    if ($c -and $Value -is [double]) {
        $judgement = Get-MetricJudgement $c $Value
        $basis = (@($Note, $c.Text) | Where-Object { $_ }) -join '。'
    }
    if (-not $Key) { $Key = $Name }
    $rows.Add([pscustomobject]@{ Area = $Area; Name = $Name; Value = $Value; Judgement = $judgement; Basis = $basis; Detail = $Detail; Key = $Key; Unit = $Unit })
}

function Add-FixedMetric {
    #     判定を決めて渡す指標（収束の判定など、数値でないもの）
    param([string] $Area, [string] $Name, $Value, [string] $Judgement, [string] $Basis, [string] $Detail)
    $rows.Add([pscustomobject]@{ Area = $Area; Name = $Name; Value = $Value; Judgement = $Judgement; Basis = $Basis; Detail = $Detail; Key = $Name; Unit = '' })
}

function Add-Action {
    param([string] $Area, [string] $Target, [string] $Content, [string] $Action)
    $actions.Add([pscustomobject]@{ Area = $Area; Target = $Target; Content = $Content; Action = $Action })
}

function Read-SheetPairs {
    #     xlsx の 1 シートの A 列 → B 列の値（ツールのサマリーのシートを読む）
    param([string] $FilePath, [string] $Sheet)
    $pairs = @{}
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value -SheetNamePattern ('^{0}$' -f [regex]::Escape($Sheet)) 6>$null)
    $books = ConvertTo-DesignBooks $hits
    foreach ($book in $books.Values) {
        $grid = $book.Sheets[$Sheet]
        if (-not $grid) { continue }
        foreach ($row in $grid.Rows) { if ($grid.Cells[$row].ContainsKey(1) -and $grid.Cells[$row].ContainsKey(2)) { $pairs[$grid.Cells[$row][1].Trim()] = $grid.Cells[$row][2].Trim() } }
    }
    return $pairs
}

function Read-SheetTable {
    #     xlsx の 1 シートの表（1 行目が見出し）を、見出し → 値 の行の一覧にする
    param([string] $FilePath, [string] $Sheet)
    $list = New-Object System.Collections.Generic.List[object]
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value -SheetNamePattern ('^{0}$' -f [regex]::Escape($Sheet)) 6>$null)
    $books = ConvertTo-DesignBooks $hits
    foreach ($book in $books.Values) {
        $grid = $book.Sheets[$Sheet]
        if (-not $grid -or $grid.Rows.Count -eq 0) { continue }
        $header = $grid.Cells[$grid.Rows[0]]
        foreach ($row in @($grid.Rows | Select-Object -Skip 1)) {
            $r = @{}
            foreach ($col in $header.Keys) { $r[$header[$col].Trim()] = [string]$grid.Cells[$row][$col] }
            $list.Add($r)
        }
    }
    return , $list
}

function Invoke-Tool {
    #     ツールを実行し、パイプラインの結果を返す（ツールの表示は出さない）。失敗したら $null を返し、ほかのツールは続ける
    param([string] $Name, [hashtable] $Arguments)
    Write-Host ("  {0} を実行しています…" -f $Name)
    try {
        $out = @(& (Join-Path $PSScriptRoot "$Name.ps1") @Arguments 6>$null)
        $ran.Add($Name)
        return , $out
    }
    catch {
        Write-Warning ("{0} が失敗しました: {1}" -f $Name, $_.Exception.Message)
        $failed.Add(('{0}（{1}）' -f $Name, $_.Exception.Message))
        return $null
    }
}

# 各ツールが同じ Excel（指摘管理表・テスト仕様書・設計書）を読み直さないよう、Search-Excel.ps1 の読み取り結果を使い回す
$global:ExcelToolsSearchCache = @{ '__Reused' = 0 }
# 途中でエラー・中断（Ctrl+C）になっても、使い回しの結果は残さない
try {
    # ツールの結果（実行しなかった・失敗したものは $null）
    $docs = $null; $functions = $null; $specIssues = $null; $phases = $null; $gaps = $null; $issues = $null; $effect = $null; $pairs = $null; $pointRates = @()
    $common = @{}
    if ($NoExcel) { $common['NoExcel'] = $true }

    # ---------------------------------------------------------------- レビュー


    if ($Findings) {
        $file = Join-Path $detailDir 'レビュー.xlsx'
        $args1 = @{ Path = $Findings; AsOf = $AsOf; OutXlsx = $file } + $common
        if ($PageCsv) { $args1['SizeCsv'] = $PageCsv } elseif ($DesignPath) { $args1['DesignPath'] = $DesignPath }
        if ($ReviewLog) { $args1['ReviewLog'] = $ReviewLog }
        if ($ReviewLimits) { $args1['Limits'] = $ReviewLimits }
        $docs = Invoke-Tool 'Measure-ReviewQuality' $args1
    }
    if ($null -ne $docs) {
        $total = ($docs | Measure-Object -Property Findings -Sum).Sum
        $open = ($docs | Measure-Object -Property Open -Sum).Sum
        $overdue = ($docs | Measure-Object -Property Overdue -Sum).Sum
        $judgedDocs = @($docs | Where-Object { $_.Judgement })
        $outside = @($judgedDocs | Where-Object { $_.Judgement -match '^(少ない|多い)' })
        Add-Metric 'レビュー' '指摘件数' ([double]$total) '' $file
        Add-Metric 'レビュー' '未完了の指摘' ([double]$open) '' $file
        Add-Metric 'レビュー' '期限切れの指摘' ([double]$overdue) '' $file
        if ($judgedDocs.Count -gt 0) {
            Add-Metric 'レビュー' '指摘密度が基準外の設計書' ([double]$outside.Count) ('基準値と比べた設計書 {0} 件のうち' -f $judgedDocs.Count) $file
        }
        foreach ($d in $outside) {
            $todo = 'レビューの観点・範囲が足りているか見直す'
            if ($d.Judgement -like '多い*') { $todo = '作成者・工程の問題がないか原因を分析する' }
            Add-Action 'レビュー' $d.Document ('指摘密度 {0}（{1}）' -f $d.Density, $d.Judgement) $todo
        }
        # 全体の指摘密度・レビュー工数密度（未レビューの設計書は除く）。実績の CSV では New-QualityLimits.ps1 で基準値にできる名前で残す
        $sized = @($docs | Where-Object { $null -ne $_.Size -and $_.Size -gt 0 -and $_.Reviewed -ne $false })
        if ($sized.Count -gt 0) {
            $unit = $sized[0].Unit
            $sizeSum = ($sized | Measure-Object -Property Size -Sum).Sum
            $findingSum = ($sized | Measure-Object -Property Findings -Sum).Sum
            Add-Metric 'レビュー' '指摘密度（全体）' ([double][math]::Round($findingSum / $sizeSum, 3)) ('指摘 {0} 件 ÷ {1} {2}' -f $findingSum, [math]::Round($sizeSum, 2), $unit) $file '指摘密度' $unit
            $withEffort = @($sized | Where-Object { $null -ne $_.Effort })
            if ($withEffort.Count -gt 0) {
                $effortSum = ($withEffort | Measure-Object -Property Effort -Sum).Sum
                $effortSize = ($withEffort | Measure-Object -Property Size -Sum).Sum
                Add-Metric 'レビュー' 'レビュー工数密度（全体）' ([double][math]::Round($effortSum / $effortSize, 3)) ('{0} 人時 ÷ {1} {2}' -f [math]::Round($effortSum, 2), [math]::Round($effortSize, 2), $unit) $file 'レビュー工数密度' $unit
            }
        }
        # レビュー実施率は、レビュー記録と設計書の一覧（規模）があるときだけ（未レビューを決められるのはこのときだけ）
        if ($ReviewLog -and ($PageCsv -or $DesignPath) -and @($docs).Count -gt 0) {
            $unreviewed = @($docs | Where-Object { $_.Reviewed -eq $false })
            $rate = [math]::Round((@($docs).Count - $unreviewed.Count) / @($docs).Count, 3)
            Add-Metric 'レビュー' 'レビュー実施率' ([double]$rate) ('設計書 {0} 件のうち {1} 件' -f @($docs).Count, (@($docs).Count - $unreviewed.Count)) $file
            foreach ($d in $unreviewed) { Add-Action 'レビュー' $d.Document '未レビュー（レビュー記録も指摘も無い）' 'レビューを実施する' }
        }
        if ($ReviewLog) {
            $zoneActions = @($docs | Where-Object { $_.PSObject.Properties['NeedsAction'] -and $_.NeedsAction })
            Add-Metric 'レビュー' '要対応のゾーンの設計書' ([double]$zoneActions.Count) 'レビュー工数密度 × 指摘密度' $file
            foreach ($d in $zoneActions) { Add-Action 'レビュー' $d.Document ('ゾーン: {0}（{1}）' -f $d.Zone, $d.ZoneView) $d.ZoneAction }
        }
    }

    # ---------------------------------------------------------------- テスト


    if ($TestPath) {
        $file = Join-Path $detailDir 'テスト.xlsx'
        $args2 = @{ Path = $TestPath; AsOf = $AsOf; OutXlsx = $file } + $common
        if ($BugList) { $args2['BugList'] = $BugList }
        if ($SizeCsv) { $args2['SizeCsv'] = $SizeCsv } elseif ($DesignPath) { $args2['DesignPath'] = $DesignPath }
        if ($TestLimits) { $args2['Limits'] = $TestLimits }
        if ($FunctionMap) { $args2['FunctionMap'] = $FunctionMap }
        $functions = Invoke-Tool 'Measure-TestQuality' $args2
    }
    if ($null -ne $functions) {
        $summary = Read-SheetPairs $file 'サマリー'
        if ($summary.ContainsKey('消化率')) {
            $rate = [double]::Parse($summary['消化率'], [System.Globalization.CultureInfo]::InvariantCulture)
            Add-Metric 'テスト' '消化率' $rate '' $file
        }
        if ($BugList) {
            Add-Metric 'テスト' 'バグ' ([double]$summary['バグ']) '' $file
            $openBugs = [int]$summary['未完了のバグ']
            Add-Metric 'テスト' '未完了のバグ' ([double]$openBugs) '' $file
            $converge = [string]$summary['収束の判定']
            $j = '△'
            if ($converge -like '収束している*') { $j = '○' } elseif ($converge -like '収束していない*') { $j = '×' } elseif ($converge -like 'まだ判定できない*') { $j = '' }
            Add-FixedMetric 'テスト' 'バグの収束（信頼度成長曲線）' $converge $j '収束していれば ○（バグが少ないうち・テストの前半は判定しない）' $file
            foreach ($key in @($summary.Keys | Where-Object { $_ -match '^直近 \d+ 日の新規バグ$' })) {
                $beforeKey = $key -replace '^直近', 'その前の'
                Add-Metric 'テスト' $key ([double]$summary[$key]) ('その前の同じ日数: {0} 件' -f $summary[$beforeKey]) $file
            }
            if ($j -eq '×') { Add-Action 'テスト' '全体' $converge 'テストを続け、バグの収束を確かめる' }
        }
        $judged = @($functions | Where-Object { $_.TestJudgement -or $_.BugJudgement })
        if ($judged.Count -gt 0) {
            $outsideFunctions = @($judged | Where-Object { $_.TestJudgement -match '^(少ない|多い)' -or $_.BugJudgement -match '^(少ない|多い)' })
            Add-Metric 'テスト' 'テスト密度・バグ密度が基準外の機能' ([double]$outsideFunctions.Count) ('基準値と比べた機能 {0} 件のうち' -f $judged.Count) $file
        }
        foreach ($f in @($functions | Where-Object NeedsAction)) { Add-Action 'テスト' $f.Function ('ゾーン: {0}（{1}）' -f $f.Zone, $f.ZoneView) $f.ZoneAction }
        # 全体のテスト密度・バグ密度（規模のある機能の合計）
        $sizedFunctions = @($functions | Where-Object { $null -ne $_.Size -and $_.Size -gt 0 })
        if ($sizedFunctions.Count -gt 0) {
            $unit = $sizedFunctions[0].Unit
            $sizeSum = ($sizedFunctions | Measure-Object -Property Size -Sum).Sum
            $caseSum = ($sizedFunctions | Measure-Object -Property Cases -Sum).Sum
            Add-Metric 'テスト' 'テスト密度（全体）' ([double][math]::Round($caseSum / $sizeSum, 3)) ('ケース {0} 件 ÷ {1} {2}' -f $caseSum, [math]::Round($sizeSum, 2), $unit) $file 'テスト密度' $unit
            if ($BugList) {
                $bugSum = ($sizedFunctions | Measure-Object -Property Bugs -Sum).Sum
                Add-Metric 'テスト' 'バグ密度（全体）' ([double][math]::Round($bugSum / $sizeSum, 3)) ('バグ {0} 件 ÷ {1} {2}' -f $bugSum, [math]::Round($sizeSum, 2), $unit) $file 'バグ密度' $unit
            }
        }
        if ($summary.ContainsKey('食い違い')) {
            $mismatch = [int]$summary['食い違い']
            Add-Metric 'テスト' 'テスト仕様書と障害管理表の食い違い' ([double]$mismatch) 'NG なのにバグ番号が無い・再テストの結果が無い など' $file
        }
    }
    if ($TestPath) {
        $file = Join-Path $detailDir 'テスト仕様書のチェック.xlsx'
        $specArgs = @{ Path = $TestPath; OutXlsx = $file } + $common
        if ($DesignPath) { $specArgs['DesignPath'] = $DesignPath }
        $specIssues = Invoke-Tool 'Test-TestSpec' $specArgs
    }
    if ($null -ne $specIssues) {
        $specErrors = @($specIssues | Where-Object Level -eq 'エラー').Count
        $specWarnings = @($specIssues | Where-Object Level -eq '警告').Count
        Add-Metric 'テスト仕様書' 'エラー（期待結果が空など）' ([double]$specErrors) '' $file
        Add-Metric 'テスト仕様書' '警告（あいまいな期待結果・異常系や境界値の不足など）' ([double]$specWarnings) '' $file
    }

    # ---------------------------------------------------------------- 工程（混入工程・摘出工程）


    if ($BugList -or $Findings) {
        $file = Join-Path $detailDir '工程別の欠陥.xlsx'
        $args3 = @{ OutXlsx = $file } + $common
        if ($BugList) { $args3['BugList'] = $BugList }
        if ($Findings) { $args3['Findings'] = $Findings }
        $phases = Invoke-Tool 'Measure-DefectRemoval' $args3
    }
    if ($null -ne $phases) {
        $inTime = ($phases | Measure-Object -Property FoundInTime -Sum).Sum
        $leaked = ($phases | Measure-Object -Property Leaked -Sum).Sum
        if (($inTime + $leaked) -gt 0) {
            Add-Metric '工程' '摘出すべき工程までに見つかった割合' ([double][math]::Round($inTime / ($inTime + $leaked), 3)) ('{0} / {1} 件' -f $inTime, ($inTime + $leaked)) $file
            Add-Metric '工程' '後の工程への流出' ([double]$leaked) '' $file
            foreach ($p in @($phases | Where-Object { $_.Leaked -gt 0 })) { Add-Action '工程' $p.Phase ('{0} で入った欠陥のうち {1} 件が後の工程で見つかった（摘出率 {2}）' -f $p.Phase, $p.Leaked, $p.RemovalRate) ('{0}のレビューの観点を見直す' -f $p.Phase) }
        }
    }

    # ---------------------------------------------------------------- レビューの効き目（設計で入ったのにテストで見つかったバグ）


    if ($Findings -and $BugList) {
        $file = Join-Path $detailDir 'レビューの効き目.xlsx'
        $args5 = @{ Findings = $Findings; BugList = $BugList; AsOf = $AsOf; OutXlsx = $file } + $common
        if ($PageCsv) { $args5['SizeCsv'] = $PageCsv } elseif ($DesignPath) { $args5['DesignPath'] = $DesignPath }
        if ($ReviewLog) { $args5['ReviewLog'] = $ReviewLog }
        if ($ReviewLimits) { $args5['Limits'] = $ReviewLimits }
        if ($FunctionMap) { $args5['FunctionMap'] = $FunctionMap }
        $effect = Invoke-Tool 'Measure-ReviewEffectiveness' $args5
    }
    if ($null -ne $effect) {
        $reviewedDocs = @($effect | Where-Object { $_.Reviewed -ne $false })
        $findingSum = ($reviewedDocs | Measure-Object -Property Findings -Sum).Sum
        $bugSum = ($effect | Measure-Object -Property DesignBugs -Sum).Sum
        if ($findingSum + $bugSum -gt 0) {
            Add-Metric 'レビュー' '設計工程の摘出率' ([double][math]::Round($findingSum / ($findingSum + $bugSum), 3)) ('レビューの指摘 {0} 件 ÷ (指摘 + 設計で入ったバグ {1} 件)' -f $findingSum, $bugSum) $file
        }
        $missed = @($effect | Where-Object NeedsAction)
        Add-Metric 'レビュー' 'レビューで見逃した設計書' ([double]$missed.Count) '設計で入ったバグがテストで見つかった設計書（指摘密度が少ない・基準内）など' $file
        foreach ($d in $missed) { Add-Action 'レビュー' $d.Document ('{0}（設計で入ったバグ {1} 件: {2}）' -f $d.Zone, $d.DesignBugs, $d.BugIds) $d.Action }
    }

    # ---------------------------------------------------------------- AI レビュー（AI と人の指摘の突き合わせ）


    if ($AiReview -and $Findings) {
        $file = Join-Path $detailDir 'AIと人の指摘.xlsx'
        $pairs = Invoke-Tool 'Compare-ReviewFindings' (@{ AiPath = $AiReview; HumanPath = $Findings; OutXlsx = $file } + $common)
    }
    elseif ($AiReview) { Write-Warning '-AiReview（AI のレビュー結果）は、-Findings（人の指摘管理表）と一緒に指定してください' }
    if ($null -ne $pairs) {
        $matched = @($pairs | Where-Object Status -eq '一致').Count
        $humanCount = $matched + @($pairs | Where-Object Status -eq '人だけ').Count
        $aiCount = $matched + @($pairs | Where-Object Status -eq 'AI だけ').Count
        if ($humanCount -gt 0) { Add-Metric 'AI レビュー' '人の指摘のうち AI も見つけた割合' ([double][math]::Round($matched / $humanCount, 3)) ('{0} / {1} 件' -f $matched, $humanCount) $file }
        if ($aiCount -gt 0) { Add-Metric 'AI レビュー' 'AI の指摘のうち人と一致した割合' ([double][math]::Round($matched / $aiCount, 3)) ('{0} / {1} 件。AI だけの指摘には、人が見落とした正しい指摘も入る' -f $matched, $aiCount) $file }
        # 観点ごとの、AI の指摘のうち人と一致した割合（実績の CSV に残し、「観点別の一致率」のシートで推移を見る）
        $pointRates = @($pairs | Where-Object { $_.PointId -and $_.Status -ne '人だけ' } | Group-Object PointId | Sort-Object Name | ForEach-Object {
                $hit = @($_.Group | Where-Object Status -eq '一致').Count
                [pscustomobject]@{ Point = $_.Name; Ai = $_.Count; Matched = $hit; Rate = [double][math]::Round($hit / $_.Count, 3) }
            })
        # 観点表への反映案（Compare-ReviewFindings の結果のシート）は、要対応にも並べる
        $proposals = @()
        if (Test-Path -LiteralPath $file) { $proposals = Read-SheetTable $file '観点表への反映案' }
        Add-Metric 'AI レビュー' '観点表への反映案' ([double]$proposals.Count) '人だけの指摘から観点を足す・判定基準を具体的にする案、AI だけの指摘が多い観点を見直す案' $file
        foreach ($p in $proposals) {
            $target = $p['観点ID']
            if (-not $target) { $target = $(if ($p['分類']) { '観点表（分類 ' + $p['分類'] + '）' } else { '観点表' }) }
            Add-Action 'AI レビュー' $target ('{0}（{1} 件）: {2}' -f $p['提案'], $p['件数'], $p['指摘の例']) '観点表を直し、Invoke-DesignReview.ps1 -CheckOnly で書き方を確かめる'
        }
    }

    # ---------------------------------------------------------------- 網羅（要件・設計・テストのつながり）


    if ($DesignPath -and $TestPath) {
        $file = Join-Path $detailDir '網羅.xlsx'
        $args4 = @{ DesignPath = $DesignPath; TestPath = $TestPath; OutXlsx = $file } + $common
        if ($Requirement) { $args4['Requirement'] = $Requirement }
        $gaps = Invoke-Tool 'Test-Traceability' $args4
    }
    if ($null -ne $gaps) {
        if ($Requirement) {
            $reqs = Read-SheetTable $file '要件'
            if ($reqs.Count -gt 0) {
                $covered = @($reqs | Where-Object { $_['判定'] -eq '○' }).Count
                $rate = [math]::Round($covered / $reqs.Count, 3)
                Add-Metric '網羅' '設計とテストのある要件の割合' ([double]$rate) ('{0} / {1} 件' -f $covered, $reqs.Count) $file
            }
        }
        # 種類（画面・イベント・メッセージ・IF）ごとに、網羅のシートに出てきた順に並べる
        $coverage = Read-SheetTable $file '網羅'
        foreach ($kind in @($coverage | ForEach-Object { $_['種類'] } | Select-Object -Unique)) {
            $group = @($coverage | Where-Object { $_['種類'] -eq $kind })
            $covered = @($group | Where-Object { $_['判定'] -eq '○' }).Count
            $rate = [math]::Round($covered / $group.Count, 3)
            Add-Metric '網羅' ('テストのある{0}の割合' -f $kind) ([double]$rate) ('{0} / {1} 件' -f $covered, $group.Count) $file
        }
        if ($gaps.Count -gt 0) { Add-Action '網羅' ('テストの無い要素 {0} 件' -f $gaps.Count) ((@($gaps | Group-InOrder Kind | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・') 'テストケースを追加する（詳細\網羅.xlsx の抜けのシート）' }
    }

    # ---------------------------------------------------------------- 設計書の整合性


    if ($DesignPath) {
        $file = Join-Path $detailDir '設計書の整合性.xlsx'
        $issues = Invoke-Tool 'Test-DesignConsistency' (@{ Path = $DesignPath; OutXlsx = $file } + $common)
    }
    if ($null -ne $issues) {
        $errors = @($issues | Where-Object Level -eq 'エラー').Count
        $warnings = @($issues | Where-Object Level -eq '警告').Count
        Add-Metric '設計書' '参照の食い違い（エラー）' ([double]$errors) '定義されていない参照・重複した定義など' $file
        Add-Metric '設計書' '参照の食い違い（警告）' ([double]$warnings) '桁数・必須の違い、一覧との違いなど' $file
    }
}
finally {
    $reused = [int]$global:ExcelToolsSearchCache['__Reused']
    Remove-Variable -Name ExcelToolsSearchCache -Scope Global -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- 品質評価のブック


# ---------------------------------------------------------------- 実績の CSV（前回との比較と、推移）


$dateText = $AsOf.ToString('yyyy/MM/dd', $invariant)
$history = New-Object System.Collections.Generic.List[object]
$previous = @{}; $previousDate = $null
$historyDates = @()
if ($HistoryCsv -and (Test-Path -LiteralPath $HistoryCsv)) {
    foreach ($h in (Read-CsvFile $HistoryCsv @('プロジェクト', '時点', '区分', '指標', '分母', '値') '実績の CSV')) { $history.Add($h) }
    $mine = @($history | Where-Object { ([string]$_.'プロジェクト') -eq $Project -and ([string]$_.'時点') -ne $dateText })
    # 時点は yyyy/MM/dd なので、文字の順で日付の順になる
    $historyDates = @($mine | ForEach-Object { [string]$_.'時点' } | Where-Object { $_ -lt $dateText } | Sort-Object -Unique)
    if ($historyDates.Count -gt 0) {
        $previousDate = $historyDates[-1]
        foreach ($h in @($mine | Where-Object { ([string]$_.'時点') -eq $previousDate })) { $previous[('{0}|{1}|{2}' -f $h.'区分', $h.'指標', $h.'分母')] = [string]$h.'値' }
    }
}
$invariantNumber = { param($t) $v = 0.0; if ([double]::TryParse([string]$t, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$v)) { $v } else { $null } }

$metricRows = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
    $values = @($r.Area, $r.Name, $r.Value)
    $detailColumn = 7
    if ($HistoryCsv) {
        $before = $null
        $k = '{0}|{1}|{2}' -f $r.Area, $r.Key, $r.Unit
        if ($previous.ContainsKey($k)) { $before = & $invariantNumber $previous[$k] }
        $change = ''
        if ($null -ne $before -and $r.Value -is [double]) { $change = [double][math]::Round($r.Value - $before, 3) }
        $values += @($(if ($null -ne $before) { [double]$before } else { '' }), $change)
        $detailColumn = 9
    }
    $values += @($r.Judgement, $r.Basis, '', '', [System.IO.Path]::GetFileName($r.Detail))
    $row = @{ Values = $values }
    if ($r.Detail) { $row.Link = @{ Column = $detailColumn; Path = $r.Detail; Location = '' } }
    $metricRows.Add($row)
}
$metricColumns = @(@{ Header = '区分'; Width = 12 }, @{ Header = '指標'; Width = 36 }, @{ Header = '実績'; Width = 14 })
if ($HistoryCsv) { $metricColumns += @(@{ Header = $(if ($previousDate) { '前回（' + $previousDate + '）' } else { '前回' }); Width = 14 }, @{ Header = '増減'; Width = 9 }) }
$metricColumns += @(@{ Header = '判定'; Width = 6 }, @{ Header = '判定の考え方'; Width = 40 }, @{ Header = '定性分析（原因）'; Width = 40 }, @{ Header = '対策'; Width = 40 }, @{ Header = '詳細'; Width = 24 })
$actionRows = New-Object System.Collections.Generic.List[object]
foreach ($a in $actions) { $actionRows.Add(@{ Values = @($a.Area, $a.Target, $a.Content, $a.Action, '', '') }) }
$scope = New-Object System.Collections.Generic.List[object]
foreach ($pair in @(@('時点', $AsOf.ToString('yyyy/MM/dd', $invariant)), @('設計書', ($DesignPath -join '、')), @('レビューの指摘管理表', ($Findings -join '、')), @('レビュー記録', ($ReviewLog -join '、')), @('テスト仕様書', ($TestPath -join '、')), @('障害管理表', ($BugList -join '、')), @('要件一覧', ($Requirement -join '、')), @('実行したツール', ($ran -join '、')), @('失敗したツール', ($failed -join '、')))) {
    if ($pair[1]) { $scope.Add(@{ Values = $pair }) }
}
$reportSheets = @(
    @{ Name = '品質評価'; Columns = $metricColumns; Rows = $metricRows },
    @{ Name = '要対応'; Columns = @(@{ Header = '区分'; Width = 12 }, @{ Header = '対象'; Width = 30 }, @{ Header = '内容'; Width = 60 }, @{ Header = '対応の例'; Width = 40 }, @{ Header = '担当'; Width = 12 }, @{ Header = '期限'; Width = 12 }); Rows = $actionRows },
    @{ Name = '対象'; Columns = @(@{ Header = '項目'; Width = 20 }, @{ Header = '内容'; Width = 80 }); Rows = $scope })

# 推移: 今回の指標ごとに、このプロジェクトの直近の時点（最大 12 回）の値を並べる
$numeric = @($rows | Where-Object { $_.Value -is [double] })
if ($HistoryCsv) {
    $dates = @(@($historyDates | Select-Object -Last 11) + $dateText)
    $byDate = @{}
    foreach ($h in @($history | Where-Object { ([string]$_.'プロジェクト') -eq $Project })) { $byDate[('{0}|{1}|{2}|{3}' -f $h.'時点', $h.'区分', $h.'指標', $h.'分母')] = [string]$h.'値' }
    $trendRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $numeric) {
        $values = @($r.Area, $r.Name)
        foreach ($d in $dates) {
            if ($d -eq $dateText) { $values += $r.Value; continue }
            $v = $byDate[('{0}|{1}|{2}|{3}' -f $d, $r.Area, $r.Key, $r.Unit)]
            $n = & $invariantNumber $v
            $values += $(if ($null -ne $n) { [double]$n } else { '' })
        }
        $trendRows.Add(@{ Values = $values })
    }
    $reportSheets += @{ Name = '推移'; Columns = @(@(@{ Header = '区分'; Width = 12 }, @{ Header = '指標'; Width = 36 }) + @($dates | ForEach-Object { @{ Header = $_; Width = 11 } })); Rows = $trendRows }
}
# 観点別の一致率: AI レビューの観点ごとに、今回と、このプロジェクトの直近の時点の一致率を並べる（観点表を直した効果を見る）
$pointArea = 'AI レビュー（観点別）'
if ($pointRates.Count -gt 0) {
    $pointDates = @($dateText)
    if ($HistoryCsv) { $pointDates = $dates }
    $pointRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in $pointRates) {
        $values = @($p.Point, [double]$p.Ai, [double]$p.Matched)
        foreach ($d in $pointDates) {
            if ($d -eq $dateText) { $values += $p.Rate; continue }
            $n = & $invariantNumber $byDate[('{0}|{1}|{2}|{3}' -f $d, $pointArea, $p.Point, '一致率')]
            $values += $(if ($null -ne $n) { [double]$n } else { '' })
        }
        $pointRows.Add(@{ Values = $values })
    }
    $reportSheets += @{ Name = '観点別の一致率'; Columns = @(@(@{ Header = '観点ID'; Width = 10 }, @{ Header = 'AI の指摘'; Width = 9 }, @{ Header = '人と一致'; Width = 9 }) + @($pointDates | ForEach-Object { @{ Header = $_; Width = 11 } })); Rows = $pointRows }
}
Export-Workbook -FilePath $reportPath -Sheets $reportSheets

# 実績の CSV に今回の値を足す（同じプロジェクト・同じ時点の行は置き換える）。New-QualityLimits.ps1 にそのまま渡せる
if ($HistoryCsv) {
    $kept = @($history | Where-Object { -not (([string]$_.'プロジェクト') -eq $Project -and ([string]$_.'時点') -eq $dateText) } | ForEach-Object {
            [pscustomobject][ordered]@{ 'プロジェクト' = $_.'プロジェクト'; '時点' = $_.'時点'; '区分' = $_.'区分'; '指標' = $_.'指標'; '分母' = $_.'分母'; '工程' = $(if ($_.PSObject.Properties['工程']) { $_.'工程' } else { '' }); '値' = $_.'値' } })
    $added = @($numeric | ForEach-Object { [pscustomobject][ordered]@{ 'プロジェクト' = $Project; '時点' = $dateText; '区分' = $_.Area; '指標' = $_.Key; '分母' = $_.Unit; '工程' = ''; '値' = ([double]$_.Value).ToString($invariant) } })
    $added += @($pointRates | ForEach-Object { [pscustomobject][ordered]@{ 'プロジェクト' = $Project; '時点' = $dateText; '区分' = $pointArea; '指標' = $_.Point; '分母' = '一致率'; '工程' = ''; '値' = $_.Rate.ToString($invariant) } })
    $all = @($kept + $added)
    [System.IO.File]::WriteAllText($HistoryCsv, (((@($all) | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

$rows | ForEach-Object { [pscustomobject]@{ Area = $_.Area; Name = $_.Name; Value = $_.Value; Judgement = $_.Judgement } }

Write-Host ''
$counts = @($rows | Where-Object { $_.Judgement } | Group-Object Judgement | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
Write-Host ('指標 {0} 件（{1}）・要対応 {2} 件' -f $rows.Count, ($counts -join '・'), $actions.Count)
if ($failed.Count -gt 0) { Write-Host ('失敗したツール: {0}' -f ($failed -join '、')) }
if ($reused -gt 0) { Write-Host ('同じ Excel の読み直し {0} 回を省きました（前のツールの読み取り結果を使いました）' -f $reused) }
if ($HistoryCsv) {
    $prevText = '前回の実績はありません'
    if ($previousDate) { $prevText = '前回 ' + $previousDate + ' と比べました' }
    Write-Host ('実績の CSV に {0} 件を足しました（プロジェクト「{1}」・{2}。{3}）: {4}' -f $numeric.Count, $Project, $dateText, $prevText, $HistoryCsv)
}
Write-Host "品質評価を保存しました: $reportPath"
Write-Host "各ツールの詳細: $detailDir"
if (@($rows | Where-Object { @('×', '△') -contains $_.Judgement }).Count -gt 0) { Write-Host ('定性分析（原因）・対策の下書き: .\Invoke-QualityAnalysis.ps1 "{0}"' -f $reportPath) }

# .SYNOPSIS
#     レビュー・テスト・工程・網羅・設計書の整合性の各ツールを実行し、工程完了判定（品質会議）に使う品質評価のブックにまとめます（Excel 不要）。
#
# .DESCRIPTION
#     渡した資料に応じて次のツールを実行し、指標ごとの実績・判定（○△×）・判定の考え方を 1 枚にまとめます。
#     定性分析（原因）・対策の列は空欄です。会議の前に記入してください（Invoke-QualityAnalysis.ps1 で AI に下書きを書かせられます）。
#     各ツールの結果は「詳細」フォルダに保存し、リンクから開けます。
#
#       -Findings        Measure-ReviewQuality.ps1（指摘件数・未完了・期限切れ・指摘密度。-ReviewLog でゾーン分析）
#       -TestPath        Measure-TestQuality.ps1（消化率・未完了のバグ・収束・テスト密度とバグ密度・食い違い）と Test-TestSpec.ps1
#       -BugList / -Findings  Measure-DefectRemoval.ps1（摘出すべき工程までに見つかった割合・流出）
#       -Findings と -BugList  Measure-ReviewEffectiveness.ps1（設計工程の摘出率・レビューで見逃した設計書）
#       -AiReview と -Findings  Compare-ReviewFindings.ps1（AI と人の指摘の一致率・観点表への反映案。観点別の一致率のシート）
#       -DesignPath と -TestPath  Test-Traceability.ps1（テストのある画面・イベント・メッセージ・IF の割合。-Requirement で要件も）
#       -DesignPath      Test-DesignConsistency.ps1（設計書どうしの参照の食い違い）
#
#     「要対応」のシートには、基準から外れた設計書・機能、要対応のゾーン、流出のあった工程、テストの抜けを書きます（担当・期限は空欄）。
#
# .PARAMETER Config
#     品質評価の設定ファイル（CSV。項目・値）か、それを置いたフォルダ（「品質評価の設定.csv」を読む）。
#     資料の場所などを毎回指定しなくて済むようにします。ひな形は tools\quality-report\品質評価の設定_ひな形.csv。
#     コマンドで指定したパラメーターは、設定ファイルより優先します。
#
# .PARAMETER OutDir
#     品質評価のブック（品質評価_日付.xlsx）と、各ツールの詳細（詳細フォルダ）を保存するフォルダ。
#
# .PARAMETER DesignPath
#     設計書のフォルダ（整合性・網羅のチェックと、規模の分母に使う）。
#
# .PARAMETER Findings
#     レビューの指摘管理表のファイル・フォルダ。
#
# .PARAMETER AiReview
#     AI のレビュー結果（Invoke-DesignReview.ps1 のレビュー結果.xlsx など）。-Findings の人の指摘と突き合わせます。
#
# .PARAMETER ReviewLog
#     レビュー記録のファイル・フォルダ（レビュー工数のゾーン分析に使う）。
#
# .PARAMETER PageCsv
#     設計書のページ数の CSV（指摘密度の分母。指定しなければ -DesignPath の規模）。
#
# .PARAMETER TestPath
#     テスト仕様書のファイル・フォルダ。
#
# .PARAMETER BugList
#     障害管理表のファイル・フォルダ。
#
# .PARAMETER Requirement
#     要件一覧のファイル（設計とテストのある要件の割合に使う。-DesignPath と -TestPath も必要）。
#
# .PARAMETER SizeCsv
#     機能ごとの規模の CSV（テスト密度・バグ密度の分母。Measure-CodeSize.ps1 -OutCsv の KLOC など）。
#
# .PARAMETER FunctionMap
#     機能の対応表の CSV（機能・別名・設計書。Measure-TestQuality.ps1 の -FunctionMap）。
#
# .PARAMETER ReviewLimits
#     レビューの基準値の CSV（Measure-ReviewQuality.ps1 の -Limits）。
#
# .PARAMETER TestLimits
#     テストの基準値の CSV（Measure-TestQuality.ps1 の -Limits）。
#
# .PARAMETER Criteria
#     判定（○△×）の基準の CSV（指標・向き・○・△）。既定は tools\quality-report\判定基準.csv。
#
# .PARAMETER HistoryCsv
#     実績の CSV（プロジェクト・時点・区分・指標・分母・工程・値）。指定すると、今回の数値の指標を足し（同じ時点は置き換える）、
#     品質評価に前回の値と増減の列、推移のシート（直近 12 回）を加えます。New-QualityLimits.ps1 にそのまま渡せます
#     （密度は「指摘密度」「テスト密度」などの名前と分母で書く）。
#
# .PARAMETER Project
#     実績の CSV に書くプロジェクトの名前。既定は -OutDir のフォルダの名前。
#
# .PARAMETER Title
#     品質評価のブックの名前（既定は「品質評価」）。
#
# .PARAMETER AsOf
#     評価の時点（期限切れ・予定日を過ぎても未実施の判定に使う）。既定は今日。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\New-QualityReport.ps1 -OutDir D:\品質報告 -DesignPath D:\設計書 -Findings D:\レビュー -ReviewLog D:\レビュー記録 `
#         -TestPath D:\テスト仕様書 -BugList D:\障害管理表.xlsx -SizeCsv D:\規模.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに各ツールと ExcelTools.psm1 が必要です。
#     判定（○△×）は機械的な目安です。基準は -Criteria の CSV で、密度の基準値は Measure-ReviewQuality.ps1・Measure-TestQuality.ps1 の基準値の CSV で決めます。
