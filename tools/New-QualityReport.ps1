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
# プロジェクトの既定は出力先のフォルダの名前。「品質報告」など出力先によく付ける名前なら、その上のフォルダ（案件のフォルダ）の名前
# （-Config のときの既定（設定ファイルのフォルダの名前）とそろえ、別の案件を同じプロジェクトにしない）
if (-not $Project) { $Project = Get-ProjectName $OutDir }

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
        # 99.9% を 0.9990000000000001 にしない（ちょうど 999 / 1000 件の割合を、基準に届かないと判定しないため）
        if ($t.EndsWith('%')) { $v = [math]::Round($v / 100.0, 10) }
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
    #       Rate: 件数から求めた割合。丸める前の値で判定し、実績には小数 3 桁に切り捨てた値を書く（2999 / 3000 件を 1 と書かない）
    param([string] $Area, [string] $Name, $Value, [string] $Note, [string] $Detail, [string] $Key, [string] $Unit, [switch] $Rate)
    $judgement = ''
    $basis = $Note
    $c = Get-Criterion $Name
    if ($c -and $Value -is [double]) {
        $judgement = Get-MetricJudgement $c $Value
        $basis = (@($Note, $c.Text) | Where-Object { $_ }) -join '。'
    }
    # 浮動小数の誤差（0.29 × 1000 = 289.99…）で 1 つ下に切り捨てないよう、わずかに足してから切り捨てる
    if ($Rate) { $Value = [math]::Floor($Value * 1000 + 1e-9) / 1000 }
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

function Add-Failure {
    #     ツールを失敗したツールにする（実行したツールからは除く）。呼び出し元は、そのツールの指標を出さない
    param([string] $Name, [string] $Reason)
    Write-Warning ("{0} が失敗しました: {1}" -f $Name, $Reason)
    [void]$ran.Remove($Name)
    $failed.Add(('{0}（{1}）' -f $Name, $Reason))
}

function Invoke-Tool {
    #     ツールを実行し、パイプラインの結果を返す（ツールの表示は出さない）。失敗したら $null を返し、ほかのツールは続ける。
    #     資料のパスが見つからない（書き間違いなど）ときは実行せず、詳細（-OutXlsx）を保存できなかった（前回の詳細を Excel で
    #     開いているなど）ときも失敗にする（読めなかった資料を 0 件として ○ にしない・前回の詳細を今回の結果として読まないため）
    param([string] $Name, [hashtable] $Arguments)
    Write-Host ("  {0} を実行しています…" -f $Name)
    try {
        $missing = @(foreach ($k in @($Arguments.Keys | Sort-Object)) {
                if ($k -ne 'OutXlsx') { @($Arguments[$k]) | Where-Object { $_ -is [string] -and $_ -and @(Get-PathItem $_).Count -eq 0 } }
            })
        if ($missing.Count -gt 0) { throw ('見つかりません: {0}' -f ($missing -join '、')) }
        # 詳細の更新日時が変わらなければ、保存できなかった（実行の開始の時刻とは比べない。共有フォルダの時計がずれていても見分けられる）
        $detail = $Arguments['OutXlsx']
        $before = $null
        if (Test-Path -LiteralPath $detail) { $before = (Get-Item -LiteralPath $detail).LastWriteTimeUtc }
        $out = @(& (Join-Path $PSScriptRoot "$Name.ps1") @Arguments 6>$null)
        if (-not (Test-Path -LiteralPath $detail) -or (Get-Item -LiteralPath $detail).LastWriteTimeUtc -eq $before) { throw ('詳細を保存できませんでした。Excel で開いていれば閉じて、実行し直してください: {0}' -f $detail) }
        $ran.Add($Name)
        return , $out
    }
    catch {
        Add-Failure $Name $_.Exception.Message
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
    # 指摘管理表の表を 1 つも読めなかった（見出しが読めないなど）ときは、未完了・期限切れを 0 件として ○ にせず、失敗にする
    # （表を読んで指摘が 0 件のときは、0 件として判定する。読んだ表は詳細の読み取り結果のシートに見出しの行がある）
    if ($null -ne $docs -and -not ($docs | Measure-Object -Property Findings -Sum).Sum) {
        $reviewTables = Read-SheetTable $file '読み取り結果'
        if (@($reviewTables | Where-Object { $_['種類'] -eq '指摘管理表' -and $_['見出しの行'] }).Count -eq 0) {
            Add-Failure 'Measure-ReviewQuality' '指摘を読めませんでした。指摘管理表の見出しを確かめてください'
            $docs = $null
        }
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
            $rate = (@($docs).Count - $unreviewed.Count) / @($docs).Count
            Add-Metric 'レビュー' 'レビュー実施率' ([double]$rate) ('設計書 {0} 件のうち {1} 件' -f @($docs).Count, (@($docs).Count - $unreviewed.Count)) $file -Rate
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
        $caseCount = [double]$summary['テストケース（対象外を除く）']
        # テストケースを 1 件も読めなかった（見出しが読めないなど）ときは、食い違いなどを 0 件として ○ にせず、失敗にする
        if ($caseCount -eq 0) { Add-Failure 'Measure-TestQuality' 'テストケースを読めませんでした。テスト仕様書の見出しを確かめてください'; $functions = $null }
    }
    if ($null -ne $functions) {
        # 消化率は、サマリーの丸めた値（1999 / 2000 件で 1）ではなく、件数から求めて判定する
        Add-Metric 'テスト' '消化率' ([double]$summary['実施（OK・NG）'] / $caseCount) '' $file -Rate
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
        # 全体のテスト密度・バグ密度（規模のある機能の合計）。規模の無い機能のケース・バグは入れられないので、
        # 除いた件数を判定の考え方に書き、その機能を要対応に並べる（黙って除くと、密度が実際より小さく見える）
        $sizedFunctions = @($functions | Where-Object { $null -ne $_.Size -and $_.Size -gt 0 })
        if ($sizedFunctions.Count -gt 0) {
            $unit = $sizedFunctions[0].Unit
            $sizeSum = ($sizedFunctions | Measure-Object -Property Size -Sum).Sum
            $caseSum = ($sizedFunctions | Measure-Object -Property Cases -Sum).Sum
            # 障害管理表に機能の列が無いと、機能ごとのバグ（バグ密度）が無い。バグがあるのに全体のバグ密度を 0 とは書かない
            $bugByFunction = $BugList -and (@($sizedFunctions | Where-Object { $null -ne $_.BugDensity }).Count -gt 0 -or [double]$summary['バグ'] -eq 0)
            if ($BugList -and -not $bugByFunction) { Add-Action 'テスト' '障害管理表' '機能の列が無いため、バグ密度を出せない' '障害管理表に機能の列を足す' }
            $bugSum = 0
            if ($bugByFunction) { $bugSum = ($sizedFunctions | Measure-Object -Property Bugs -Sum).Sum }
            $leftCases = [int]($caseCount - $caseSum)
            $leftBugs = 0
            if ($bugByFunction) { $leftBugs = [int]([double]$summary['バグ'] - $bugSum) }
            $caseNote = 'ケース {0} 件 ÷ {1} {2}' -f $caseSum, [math]::Round($sizeSum, 2), $unit
            if ($leftCases -gt 0) { $caseNote += '。規模の無い機能のケース {0} 件は除く' -f $leftCases }
            Add-Metric 'テスト' 'テスト密度（全体）' ([double][math]::Round($caseSum / $sizeSum, 3)) $caseNote $file 'テスト密度' $unit
            if ($bugByFunction) {
                $bugNote = 'バグ {0} 件 ÷ {1} {2}' -f $bugSum, [math]::Round($sizeSum, 2), $unit
                if ($leftBugs -gt 0) { $bugNote += '。規模の無い機能のバグ {0} 件は除く' -f $leftBugs }
                Add-Metric 'テスト' 'バグ密度（全体）' ([double][math]::Round($bugSum / $sizeSum, 3)) $bugNote $file 'バグ密度' $unit
            }
            foreach ($f in @($functions | Where-Object { -not ($null -ne $_.Size -and $_.Size -gt 0) -and ($_.Cases -gt 0 -or ($bugByFunction -and $_.Bugs -gt 0)) })) {
                $counts = 'ケース {0} 件' -f $f.Cases
                if ($bugByFunction) { $counts += '・バグ {0} 件' -f $f.Bugs }
                $todo = '規模の CSV に足すか、機能の対応表で規模のある機能に対応付ける'
                if ($f.Function -eq '（機能なし）') { $todo = '障害管理表の機能の列を書く' }
                Add-Action 'テスト' $f.Function ('規模が無いため、全体のテスト密度・バグ密度に入れていない（{0}）' -f $counts) $todo
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
    # テストケースを 1 件も読めなかった（機能別のシートが空）ときは、エラー・警告を 0 件として ○ にせず、失敗にする
    if ($null -ne $specIssues -and (Read-SheetTable $file '機能別').Count -eq 0) { Add-Failure 'Test-TestSpec' 'テストケースを読めませんでした。テスト仕様書の見出しを確かめてください'; $specIssues = $null }
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
            Add-Metric '工程' '摘出すべき工程までに見つかった割合' ([double]($inTime / ($inTime + $leaked))) ('{0} / {1} 件' -f $inTime, ($inTime + $leaked)) $file -Rate
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
        # 未レビューの設計書は、指摘もバグも入れない（Measure-ReviewEffectiveness.ps1 の全体の摘出率と同じ）
        $bugSum = ($reviewedDocs | Measure-Object -Property DesignBugs -Sum).Sum
        if ($findingSum + $bugSum -gt 0) {
            Add-Metric 'レビュー' '設計工程の摘出率' ([double]($findingSum / ($findingSum + $bugSum))) ('レビューの指摘 {0} 件 ÷ (指摘 + 設計で入ったバグ {1} 件)' -f $findingSum, $bugSum) $file -Rate
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
        if ($humanCount -gt 0) { Add-Metric 'AI レビュー' '人の指摘のうち AI も見つけた割合' ([double]($matched / $humanCount)) ('{0} / {1} 件' -f $matched, $humanCount) $file -Rate }
        if ($aiCount -gt 0) { Add-Metric 'AI レビュー' 'AI の指摘のうち人と一致した割合' ([double]($matched / $aiCount)) ('{0} / {1} 件。AI だけの指摘には、人が見落とした正しい指摘も入る' -f $matched, $aiCount) $file -Rate }
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
                Add-Metric '網羅' '設計とテストのある要件の割合' ([double]($covered / $reqs.Count)) ('{0} / {1} 件' -f $covered, $reqs.Count) $file -Rate
            }
        }
        # 種類（画面・イベント・メッセージ・IF）ごとに、網羅のシートに出てきた順に並べる
        $coverage = Read-SheetTable $file '網羅'
        foreach ($kind in @($coverage | ForEach-Object { $_['種類'] } | Select-Object -Unique)) {
            $group = @($coverage | Where-Object { $_['種類'] -eq $kind })
            $covered = @($group | Where-Object { $_['判定'] -eq '○' }).Count
            Add-Metric '網羅' ('テストのある{0}の割合' -f $kind) ([double]($covered / $group.Count)) ('{0} / {1} 件' -f $covered, $group.Count) $file -Rate
        }
        if ($gaps.Count -gt 0) { Add-Action '網羅' ('テストの無い要素 {0} 件' -f $gaps.Count) ((@($gaps | Group-InOrder Kind | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・') 'テストケースを追加する（詳細\網羅.xlsx の抜けのシート）' }
    }

    # ---------------------------------------------------------------- 設計書の整合性


    if ($DesignPath) {
        $file = Join-Path $detailDir '設計書の整合性.xlsx'
        $issues = Invoke-Tool 'Test-DesignConsistency' (@{ Path = $DesignPath; OutXlsx = $file } + $common)
    }
    # 設計書から定義も参照も読めなかった（定義と参照のシートが空）ときは、食い違いを 0 件として ○ にせず、失敗にする
    if ($null -ne $issues -and (Read-SheetTable $file '定義と参照').Count -eq 0) { Add-Failure 'Test-DesignConsistency' '設計書から定義も参照も読めませんでした。参照のルールのファイル名・見出しを確かめてください'; $issues = $null }
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
    foreach ($h in (Read-CsvFile $HistoryCsv @('プロジェクト', '時点', '区分', '指標', '分母', '値') '実績の CSV')) {
        # Excel で開いて保存すると、時点が 2026/10/6 の形になる。日付として読み、yyyy/MM/dd の形にそろえる（そのまま書き戻す）
        $day = ConvertTo-DateValue ([string]$h.'時点')
        if ($day) { $h.'時点' = $day.ToString('yyyy/MM/dd', $invariant) }
        $history.Add($h)
    }
    $mine = @($history | Where-Object { ([string]$_.'プロジェクト') -eq $Project -and ([string]$_.'時点') -ne $dateText })
    # 実績の CSV に別の 1 つのプロジェクトの行だけがあるときは、プロジェクトの名前が変わった（以前の既定の「品質報告」など）かもしれないので知らせる
    $otherProjects = @($history | ForEach-Object { ([string]$_.'プロジェクト').Trim() } | Where-Object { $_ -and $_ -ne $Project } | Sort-Object -Unique)
    if ($otherProjects.Count -eq 1 -and @($history | Where-Object { ([string]$_.'プロジェクト') -eq $Project }).Count -eq 0) {
        Write-Warning ("実績の CSV には、プロジェクト「{0}」の行だけがあり、今回のプロジェクト「{1}」の行がありません。同じプロジェクトなら、CSV のプロジェクトの列を「{1}」に直すか、-Project（設定ファイルの「プロジェクト」）に「{0}」を指定してください: {2}" -f $otherProjects[0], $Project, $HistoryCsv)
    }
    # 時点は yyyy/MM/dd なので、文字の順で日付の順になる
    $historyDates = @($mine | ForEach-Object { [string]$_.'時点' } | Where-Object { $_ -lt $dateText } | Sort-Object -Unique)
    if ($historyDates.Count -gt 0) {
        $previousDate = $historyDates[-1]
        foreach ($h in @($mine | Where-Object { ([string]$_.'時点') -eq $previousDate })) { $previous[('{0}|{1}|{2}' -f $h.'区分', $h.'指標', $h.'分母')] = [string]$h.'値' }
    }
}
$invariantNumber = { param($t) $v = 0.0; if ([double]::TryParse([string]$t, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$v)) { $v } else { $null } }
# 分母の以前の名前。Measure-TestQuality.ps1 の設計書の FP（4 × EI + 5 × EO + 4 × EQ）は、以前は「FP」と書いた（Measure-DesignSize.ps1 の
# NESMA の FP と見分けるため「FP（トランザクション）」にした）。以前の実績の行も、前回・推移に使う
$oldUnitNames = @{ 'FP（トランザクション）' = 'FP' }
function Get-HistoryValue {
    #     実績の値（Map: キー → 値。キーは Prefix|分母）。今の分母の名前の行が無ければ、以前の分母の名前の行
    param([hashtable] $Map, [string] $Prefix, [string] $Unit)
    $k = '{0}|{1}' -f $Prefix, $Unit
    if ($Map.ContainsKey($k)) { return $Map[$k] }
    if ($Unit -and $oldUnitNames.ContainsKey($Unit)) {
        $k = '{0}|{1}' -f $Prefix, $oldUnitNames[$Unit]
        if ($Map.ContainsKey($k)) { return $Map[$k] }
    }
    return $null
}

# 同じ時点の品質評価のブックを作り直すとき（資料を直して同じ日にもう一度作るなど）は、前のブックに記入した
# 定性分析（原因）・対策（区分と指標で当てる）と、要対応の担当・期限（区分・対象・内容で当て、内容の数値が変わった行は
# 区分と対象が前と今回でそれぞれ 1 つだけのときに当てる）を引き継ぐ。引き継げなかった記入があれば、前のブックを残す
$oneLine = { param($t) ([string]$t -replace '\r\n|\n|\r', '↵').Trim() }
$restoreText = { param($t) ([string]$t).Replace('↵', "`n").Trim() }
$oldNotes = [ordered]@{}
$oldOwners = New-Object System.Collections.Generic.List[object]
if (Test-Path -LiteralPath $reportPath) {
    foreach ($r in (Read-SheetTable $reportPath '品質評価')) {
        $cause = & $restoreText $r['定性分析（原因）']; $measure = & $restoreText $r['対策']
        if ($cause -or $measure) { $oldNotes[('{0}|{1}' -f (& $oneLine $r['区分']), (& $oneLine $r['指標']))] = [pscustomobject]@{ Name = (& $oneLine $r['指標']); Cause = $cause; Measure = $measure; Used = $false } }
    }
    foreach ($r in (Read-SheetTable $reportPath '要対応')) {
        $owner = & $restoreText $r['担当']; $due = & $restoreText $r['期限']
        if (-not ($owner -or $due)) { continue }
        # 期限の日付のセル（Search-Excel は yyyy/MM/dd で読む）は、日付のセルとして書き戻す
        $dueDay = [datetime]::MinValue
        if ([datetime]::TryParseExact($due, 'yyyy/MM/dd', $invariant, [System.Globalization.DateTimeStyles]::None, [ref]$dueDay)) { $due = $dueDay }
        $oldOwners.Add([pscustomobject]@{ Area = (& $oneLine $r['区分']); Target = (& $oneLine $r['対象']); Content = (& $oneLine $r['内容']); Owner = $owner; Due = $due; Used = $false })
    }
}

$metricRows = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
    $values = @($r.Area, $r.Name, $r.Value)
    $detailColumn = 7
    if ($HistoryCsv) {
        $before = $null
        $v = Get-HistoryValue $previous ('{0}|{1}' -f $r.Area, $r.Key) $r.Unit
        if ($null -ne $v) { $before = & $invariantNumber $v }
        $change = ''
        if ($null -ne $before -and $r.Value -is [double]) { $change = [double][math]::Round($r.Value - $before, 3) }
        $values += @($(if ($null -ne $before) { [double]$before } else { '' }), $change)
        $detailColumn = 9
    }
    $cause = ''; $measure = ''
    $note = $oldNotes[('{0}|{1}' -f (& $oneLine $r.Area), (& $oneLine $r.Name))]
    if ($note -and -not $note.Used) { $cause = $note.Cause; $measure = $note.Measure; $note.Used = $true }
    $values += @($r.Judgement, $r.Basis, $cause, $measure, [System.IO.Path]::GetFileName($r.Detail))
    $row = @{ Values = $values }
    if ($r.Detail) { $row.Link = @{ Column = $detailColumn; Path = $r.Detail; Location = '' } }
    $metricRows.Add($row)
}
$metricColumns = @(@{ Header = '区分'; Width = 12 }, @{ Header = '指標'; Width = 36 }, @{ Header = '実績'; Width = 14 })
if ($HistoryCsv) { $metricColumns += @(@{ Header = $(if ($previousDate) { '前回（' + $previousDate + '）' } else { '前回' }); Width = 14 }, @{ Header = '増減'; Width = 9 }) }
$metricColumns += @(@{ Header = '判定'; Width = 6 }, @{ Header = '判定の考え方'; Width = 40 }, @{ Header = '定性分析（原因）'; Width = 40 }, @{ Header = '対策'; Width = 40 }, @{ Header = '詳細'; Width = 24 })
# 要対応の担当・期限: 区分・対象・内容が同じ行に当て、残りは区分・対象が前と今回でそれぞれ 1 つだけの行に当てる
$owners = @{}
foreach ($pass in 1, 2) {
    for ($i = 0; $i -lt $actions.Count; $i++) {
        if ($owners.ContainsKey($i)) { continue }
        $a = $actions[$i]
        $area = & $oneLine $a.Area; $target = & $oneLine $a.Target; $content = & $oneLine $a.Content
        $hit = @($oldOwners | Where-Object { -not $_.Used -and $_.Area -eq $area -and $_.Target -eq $target -and ($pass -eq 2 -or $_.Content -eq $content) })
        if ($pass -eq 2) {
            $same = @(for ($j = 0; $j -lt $actions.Count; $j++) { if (-not $owners.ContainsKey($j) -and (& $oneLine $actions[$j].Area) -eq $area -and (& $oneLine $actions[$j].Target) -eq $target) { $j } })
            if ($hit.Count -ne 1 -or $same.Count -ne 1) { continue }
        }
        if ($hit.Count -eq 0) { continue }
        $hit[0].Used = $true
        $owners[$i] = $hit[0]
    }
}
$actionRows = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $actions.Count; $i++) {
    $a = $actions[$i]
    $owner = ''; $due = ''
    if ($owners.ContainsKey($i)) { $owner = $owners[$i].Owner; $due = $owners[$i].Due }
    $actionRows.Add(@{ Values = @($a.Area, $a.Target, $a.Content, $a.Action, $owner, $due) })
}
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
            $v = Get-HistoryValue $byDate ('{0}|{1}|{2}' -f $d, $r.Area, $r.Key) $r.Unit
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
# 引き継げなかった記入（今回の品質評価に無い指標・要対応の行）があれば、上書きする前のブックを残す
$carriedNotes = @($oldNotes.Values | Where-Object Used)
$carriedOwners = @($oldOwners | Where-Object Used)
$lostNotes = @($oldNotes.Values | Where-Object { -not $_.Used })
$lostOwners = @($oldOwners | Where-Object { -not $_.Used })
$keptBook = $null
if ($lostNotes.Count + $lostOwners.Count -gt 0) {
    $keptBook = Get-BackupPath (Join-Path $OutDir ('{0}_{1}_前の版.xlsx' -f $Title, $AsOf.ToString('yyyyMMdd', $invariant)))
    Copy-Item -LiteralPath $reportPath -Destination $keptBook
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
if ($carriedNotes.Count + $carriedOwners.Count -gt 0) {
    $carriedText = '同じ時点の前の品質評価から、定性分析（原因）・対策 {0} 件' -f $carriedNotes.Count
    if ($carriedNotes.Count -gt 0) { $carriedText += '（{0}）' -f (@($carriedNotes | ForEach-Object { $_.Name }) -join '・') }
    Write-Host ('{0}と、要対応の担当・期限 {1} 件を引き継ぎました' -f $carriedText, $carriedOwners.Count)
}
if ($keptBook) {
    $lostText = @($lostNotes | ForEach-Object { '定性分析・対策（{0}）' -f $_.Name }) + @($lostOwners | ForEach-Object { '要対応の担当・期限（{0} {1}）' -f $_.Target, $_.Content })
    Write-Warning ('今回の品質評価に無い指標・要対応の行に記入したものは、引き継げませんでした: {0}。前のブックを残しました: {1}' -f ($lostText -join '・'), $keptBook)
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
#     同じ時点のブックを作り直すときは、前のブックに記入した定性分析（原因）・対策と、要対応の担当・期限を引き継ぎます
#     （今回の品質評価に無い指標・要対応の行に記入したものがあれば、前のブックを「<ブック名>_前の版.xlsx」として残します）。
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
#     「要対応」のシートには、基準から外れた設計書・機能、要対応のゾーン、流出のあった工程、テストの抜け、
#     規模が無いため全体のテスト密度・バグ密度に入れていない機能を書きます（担当・期限は空欄）。
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
#     実績の CSV に書くプロジェクトの名前。既定は -OutDir のフォルダの名前（「品質報告」など出力先によく付ける名前なら、
#     その上のフォルダ（案件のフォルダ）の名前。-Config のときは設定ファイルのフォルダの名前）。
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
