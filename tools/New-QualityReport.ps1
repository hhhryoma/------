[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $OutDir,

    [string[]] $DesignPath,

    [string[]] $Findings,

    [string[]] $ReviewLog,

    [string] $PageCsv,

    [string[]] $TestPath,

    [string[]] $BugList,

    [string[]] $Requirement,

    [string] $SizeCsv,

    [string] $ReviewLimits,

    [string] $TestLimits,

    [string] $Title = '品質評価',

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Findings -and -not $TestPath -and -not $DesignPath) { throw '-Findings（レビューの指摘管理表）・-TestPath（テスト仕様書）・-DesignPath（設計書）のどれかを指定してください' }
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$detailDir = Join-Path $OutDir '詳細'
New-Item -ItemType Directory -Force -Path $detailDir | Out-Null
$reportPath = Join-Path $OutDir ('{0}_{1}.xlsx' -f $Title, $AsOf.ToString('yyyyMMdd'))

# ---------------------------------------------------------------- ヘルパー


$rows = New-Object System.Collections.Generic.List[object]      # 品質評価の行
$actions = New-Object System.Collections.Generic.List[object]   # 要対応の行
$ran = New-Object System.Collections.Generic.List[string]
$failed = New-Object System.Collections.Generic.List[string]

function Add-Metric {
    #     品質評価の 1 行（区分・指標・実績・判定・判定の考え方・詳細のファイル）
    param([string] $Area, [string] $Name, $Value, [string] $Judgement, [string] $Basis, [string] $Detail)
    $rows.Add([pscustomobject]@{ Area = $Area; Name = $Name; Value = $Value; Judgement = $Judgement; Basis = $Basis; Detail = $Detail })
}

function Add-Action {
    param([string] $Area, [string] $Target, [string] $Content, [string] $Action)
    $actions.Add([pscustomobject]@{ Area = $Area; Target = $Target; Content = $Content; Action = $Action })
}

function Get-CountJudgement {
    #     問題の数の判定（0 なら ○、それ以外は △。Severe なら ×）
    param([int] $Count, [switch] $Severe)
    if ($Count -eq 0) { return '○' }
    if ($Severe) { return '×' }
    return '△'
}

function Get-RateJudgement {
    #     割合の判定（1 なら ○、Warn 以上なら △、それ未満は ×）
    param([double] $Rate, [double] $Warn = 0.8)
    if ($Rate -ge 1.0) { return '○' }
    if ($Rate -ge $Warn) { return '△' }
    return '×'
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

# ツールの結果（実行しなかった・失敗したものは $null）
$docs = $null; $functions = $null; $specIssues = $null; $phases = $null; $gaps = $null; $issues = $null
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
    Add-Metric 'レビュー' '指摘件数' ([double]$total) '' '' $file
    Add-Metric 'レビュー' '未完了の指摘' ([double]$open) (Get-CountJudgement $open) '0 件なら ○' $file
    Add-Metric 'レビュー' '期限切れの指摘' ([double]$overdue) (Get-CountJudgement $overdue -Severe) '0 件なら ○、残っていれば ×' $file
    if ($judgedDocs.Count -gt 0) {
        Add-Metric 'レビュー' '指摘密度が基準外の設計書' ([double]$outside.Count) (Get-CountJudgement $outside.Count) ('基準値と比べた設計書 {0} 件のうち。0 件なら ○' -f $judgedDocs.Count) $file
    }
    foreach ($d in $outside) {
        $todo = 'レビューの観点・範囲が足りているか見直す'
        if ($d.Judgement -like '多い*') { $todo = '作成者・工程の問題がないか原因を分析する' }
        Add-Action 'レビュー' $d.Document ('指摘密度 {0}（{1}）' -f $d.Density, $d.Judgement) $todo
    }
    if ($ReviewLog) {
        $zoneActions = @($docs | Where-Object { $_.PSObject.Properties['NeedsAction'] -and $_.NeedsAction })
        Add-Metric 'レビュー' '要対応のゾーンの設計書' ([double]$zoneActions.Count) (Get-CountJudgement $zoneActions.Count) 'レビュー工数密度 × 指摘密度。0 件なら ○' $file
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
    $functions = Invoke-Tool 'Measure-TestQuality' $args2
}
if ($null -ne $functions) {
    $summary = Read-SheetPairs $file 'サマリー'
    if ($summary.ContainsKey('消化率')) {
        $rate = [double]::Parse($summary['消化率'], [System.Globalization.CultureInfo]::InvariantCulture)
        Add-Metric 'テスト' '消化率' $rate (Get-RateJudgement $rate 0.9) '100% なら ○、90% 以上なら △' $file
    }
    if ($BugList) {
        Add-Metric 'テスト' 'バグ' ([double]$summary['バグ']) '' '' $file
        $openBugs = [int]$summary['未完了のバグ']
        Add-Metric 'テスト' '未完了のバグ' ([double]$openBugs) (Get-CountJudgement $openBugs) '0 件なら ○' $file
        $converge = [string]$summary['収束の判定']
        $j = '△'
        if ($converge -like '収束している*') { $j = '○' } elseif ($converge -like '収束していない*') { $j = '×' } elseif ($converge -like 'まだ判定できない*') { $j = '' }
        Add-Metric 'テスト' 'バグの収束（信頼度成長曲線）' $converge $j '収束していれば ○（バグが少ないうち・テストの前半は判定しない）' $file
        foreach ($key in @($summary.Keys | Where-Object { $_ -match '^直近 \d+ 日の新規バグ$' })) {
            $beforeKey = $key -replace '^直近', 'その前の'
            Add-Metric 'テスト' $key ([double]$summary[$key]) '' ('その前の同じ日数: {0} 件' -f $summary[$beforeKey]) $file
        }
        if ($j -eq '×') { Add-Action 'テスト' '全体' $converge 'テストを続け、バグの収束を確かめる' }
    }
    $judged = @($functions | Where-Object { $_.TestJudgement -or $_.BugJudgement })
    if ($judged.Count -gt 0) {
        $outsideFunctions = @($judged | Where-Object { $_.TestJudgement -match '^(少ない|多い)' -or $_.BugJudgement -match '^(少ない|多い)' })
        Add-Metric 'テスト' 'テスト密度・バグ密度が基準外の機能' ([double]$outsideFunctions.Count) (Get-CountJudgement $outsideFunctions.Count) ('基準値と比べた機能 {0} 件のうち。0 件なら ○' -f $judged.Count) $file
    }
    foreach ($f in @($functions | Where-Object NeedsAction)) { Add-Action 'テスト' $f.Function ('ゾーン: {0}（{1}）' -f $f.Zone, $f.ZoneView) $f.ZoneAction }
    if ($summary.ContainsKey('食い違い')) {
        $mismatch = [int]$summary['食い違い']
        Add-Metric 'テスト' 'テスト仕様書と障害管理表の食い違い' ([double]$mismatch) (Get-CountJudgement $mismatch) 'NG なのにバグ番号が無い・再テストの結果が無い など。0 件なら ○' $file
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
    Add-Metric 'テスト仕様書' 'エラー（期待結果が空など）' ([double]$specErrors) (Get-CountJudgement $specErrors -Severe) '0 件なら ○' $file
    Add-Metric 'テスト仕様書' '警告（あいまいな期待結果・異常系や境界値の不足など）' ([double]$specWarnings) (Get-CountJudgement $specWarnings) '0 件なら ○' $file
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
        Add-Metric '工程' '摘出すべき工程までに見つかった割合' ([double][math]::Round($inTime / ($inTime + $leaked), 3)) '' '社内の基準と比べて判定してください' $file
        Add-Metric '工程' '後の工程への流出' ([double]$leaked) (Get-CountJudgement $leaked) '0 件なら ○' $file
        foreach ($p in @($phases | Where-Object { $_.Leaked -gt 0 })) { Add-Action '工程' $p.Phase ('{0} で入った欠陥のうち {1} 件が後の工程で見つかった（摘出率 {2}）' -f $p.Phase, $p.Leaked, $p.RemovalRate) ('{0}のレビューの観点を見直す' -f $p.Phase) }
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
            Add-Metric '網羅' '設計とテストのある要件の割合' ([double]$rate) (Get-RateJudgement $rate) ('{0} / {1} 件。100% なら ○、80% 以上なら △' -f $covered, $reqs.Count) $file
        }
    }
    # 種類（画面・イベント・メッセージ・IF）ごとに、網羅のシートに出てきた順に並べる
    $coverage = Read-SheetTable $file '網羅'
    foreach ($kind in @($coverage | ForEach-Object { $_['種類'] } | Select-Object -Unique)) {
        $group = @($coverage | Where-Object { $_['種類'] -eq $kind })
        $covered = @($group | Where-Object { $_['判定'] -eq '○' }).Count
        $rate = [math]::Round($covered / $group.Count, 3)
        Add-Metric '網羅' ('テストのある{0}の割合' -f $kind) ([double]$rate) (Get-RateJudgement $rate) ('{0} / {1} 件。100% なら ○、80% 以上なら △' -f $covered, $group.Count) $file
    }
    if ($gaps.Count -gt 0) { Add-Action '網羅' ('テストの無い要素 {0} 件' -f $gaps.Count) ((@($gaps | Group-Object Kind | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・') 'テストケースを追加する（詳細\網羅.xlsx の抜けのシート）' }
}

# ---------------------------------------------------------------- 設計書の整合性


if ($DesignPath) {
    $file = Join-Path $detailDir '設計書の整合性.xlsx'
    $issues = Invoke-Tool 'Test-DesignConsistency' (@{ Path = $DesignPath; OutXlsx = $file } + $common)
}
if ($null -ne $issues) {
    $errors = @($issues | Where-Object Level -eq 'エラー').Count
    $warnings = @($issues | Where-Object Level -eq '警告').Count
    Add-Metric '設計書' '参照の食い違い（エラー）' ([double]$errors) (Get-CountJudgement $errors -Severe) '定義されていない参照・重複した定義など。0 件なら ○' $file
    Add-Metric '設計書' '参照の食い違い（警告）' ([double]$warnings) (Get-CountJudgement $warnings) '桁数・必須の違い、一覧との違いなど。0 件なら ○' $file
}

# ---------------------------------------------------------------- 品質評価のブック


$metricRows = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
    $row = @{ Values = @($r.Area, $r.Name, $r.Value, $r.Judgement, $r.Basis, '', '', [System.IO.Path]::GetFileName($r.Detail)) }
    if ($r.Detail) { $row.Link = @{ Column = 7; Path = $r.Detail; Location = '' } }
    $metricRows.Add($row)
}
$actionRows = New-Object System.Collections.Generic.List[object]
foreach ($a in $actions) { $actionRows.Add(@{ Values = @($a.Area, $a.Target, $a.Content, $a.Action, '', '') }) }
$scope = New-Object System.Collections.Generic.List[object]
foreach ($pair in @(@('時点', $AsOf.ToString('yyyy/MM/dd')), @('設計書', ($DesignPath -join '、')), @('レビューの指摘管理表', ($Findings -join '、')), @('レビュー記録', ($ReviewLog -join '、')), @('テスト仕様書', ($TestPath -join '、')), @('障害管理表', ($BugList -join '、')), @('要件一覧', ($Requirement -join '、')), @('実行したツール', ($ran -join '、')), @('失敗したツール', ($failed -join '、')))) {
    if ($pair[1]) { $scope.Add(@{ Values = $pair }) }
}
Export-Workbook -FilePath $reportPath -Sheets @(
    @{ Name = '品質評価'; Columns = @(@{ Header = '区分'; Width = 12 }, @{ Header = '指標'; Width = 36 }, @{ Header = '実績'; Width = 14 }, @{ Header = '判定'; Width = 6 }, @{ Header = '判定の考え方'; Width = 40 }, @{ Header = '定性分析（原因）'; Width = 40 }, @{ Header = '対策'; Width = 40 }, @{ Header = '詳細'; Width = 24 }); Rows = $metricRows },
    @{ Name = '要対応'; Columns = @(@{ Header = '区分'; Width = 12 }, @{ Header = '対象'; Width = 30 }, @{ Header = '内容'; Width = 60 }, @{ Header = '対応の例'; Width = 40 }, @{ Header = '担当'; Width = 12 }, @{ Header = '期限'; Width = 12 }); Rows = $actionRows },
    @{ Name = '対象'; Columns = @(@{ Header = '項目'; Width = 20 }, @{ Header = '内容'; Width = 80 }); Rows = $scope })

$rows | ForEach-Object { [pscustomobject]@{ Area = $_.Area; Name = $_.Name; Value = $_.Value; Judgement = $_.Judgement } }

Write-Host ''
$counts = @($rows | Where-Object { $_.Judgement } | Group-Object Judgement | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
Write-Host ('指標 {0} 件（{1}）・要対応 {2} 件' -f $rows.Count, ($counts -join '・'), $actions.Count)
if ($failed.Count -gt 0) { Write-Host ('失敗したツール: {0}' -f ($failed -join '、')) }
Write-Host "品質評価を保存しました: $reportPath"
Write-Host "各ツールの詳細: $detailDir"

# .SYNOPSIS
#     レビュー・テスト・工程・網羅・設計書の整合性の各ツールを実行し、工程完了判定（品質会議）に使う品質評価のブックにまとめます（Excel 不要）。
#
# .DESCRIPTION
#     渡した資料に応じて次のツールを実行し、指標ごとの実績・判定（○△×）・判定の考え方を 1 枚にまとめます。
#     定性分析（原因）・対策の列は空欄です。会議の前に記入してください。各ツールの結果は「詳細」フォルダに保存し、リンクから開けます。
#
#       -Findings        Measure-ReviewQuality.ps1（指摘件数・未完了・期限切れ・指摘密度。-ReviewLog でゾーン分析）
#       -TestPath        Measure-TestQuality.ps1（消化率・未完了のバグ・収束・テスト密度とバグ密度・食い違い）と Test-TestSpec.ps1
#       -BugList / -Findings  Measure-DefectRemoval.ps1（摘出すべき工程までに見つかった割合・流出）
#       -DesignPath と -TestPath  Test-Traceability.ps1（テストのある画面・イベント・メッセージ・IF の割合。-Requirement で要件も）
#       -DesignPath      Test-DesignConsistency.ps1（設計書どうしの参照の食い違い）
#
#     「要対応」のシートには、基準から外れた設計書・機能、要対応のゾーン、流出のあった工程、テストの抜けを書きます（担当・期限は空欄）。
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
# .PARAMETER ReviewLimits
#     レビューの基準値の CSV（Measure-ReviewQuality.ps1 の -Limits）。
#
# .PARAMETER TestLimits
#     テストの基準値の CSV（Measure-TestQuality.ps1 の -Limits）。
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
#     判定（○△×）は機械的な目安です。基準値は Measure-ReviewQuality.ps1・Measure-TestQuality.ps1 の基準値の CSV で決めます。
