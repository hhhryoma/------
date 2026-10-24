[CmdletBinding()]
param(
    [string[]] $BugList,

    [string[]] $Findings,

    [string] $Phases,

    [string] $Headers,

    [string] $FindingHeaders,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $BugList -and -not $Findings) { throw '-BugList（障害管理表）か -Findings（レビューの指摘管理表）を指定してください' }
if (-not $Phases) { $Phases = Join-Path $PSScriptRoot 'test-quality/工程.csv' }
if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
if (-not $FindingHeaders) { $FindingHeaders = Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv' }
foreach ($name in @('Phases', 'Headers', 'FindingHeaders', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'DefectRemoval.Phase' -DefaultDisplayPropertySet Phase, Injected, FoundInTime, Leaked, RemovalRate, Detected, FromEarlier -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 数えない状態（取下げ・重複など）
$excludedPattern = '取下げ|取り下げ|取消|取り消|却下|対象外|重複|仕様どおり|仕様通り|障害ではない|非障害|バグではない|誤指摘'
$unknown = '（不明）'

# ---------------------------------------------------------------- 工程


$phaseList = Read-PhaseList $Phases

function Resolve-Phase {
    param([string] $Text)
    return (Resolve-PhaseName $phaseList $Text)
}

# 工程 → 摘出すべき工程（UT のような別名でも書ける。空か、工程として読めなければ同じ工程）
$expectedOf = @{}
$unreadExpected = New-Object System.Collections.Generic.List[string]
foreach ($p in $phaseList) {
    $expected = $null
    if ($p.Expected) {
        $expected = Resolve-Phase $p.Expected
        if (-not $expected) { $unreadExpected.Add(('{0}の「{1}」' -f $p.Name, $p.Expected)) }
    }
    if (-not $expected) { $expected = $p }
    $expectedOf[$p.Name] = $expected
}
if ($unreadExpected.Count -gt 0) { Write-Warning ('工程の CSV の摘出すべき工程に、工程として読めない値があります（その工程で入った欠陥は、同じ工程までに見つけるものとして数えます）: {0}' -f ($unreadExpected -join '、')) }

# ---------------------------------------------------------------- 欠陥を読む


$defects = New-Object System.Collections.Generic.List[object]
$excludedCount = 0
$unreadPhases = @{}

function Add-Defect {
    param([string] $Source, [string] $Content, [string] $InjectedText, [string] $DetectedText, [string] $Cause, [string] $Severity, [string] $Status, $Record)
    if ($Status -match $excludedPattern) { $script:excludedCount++; return }
    $injected = Resolve-Phase $InjectedText
    $detected = Resolve-Phase $DetectedText
    foreach ($pair in @(@($InjectedText, $injected), @($DetectedText, $detected))) {
        if ($pair[0] -and -not $pair[1]) { $unreadPhases[$pair[0].Trim()] = $true }
    }
    $defects.Add([pscustomobject]@{
        Source = $Source; Content = $Content.Trim(); Cause = $Cause.Trim(); Severity = $Severity.Trim()
        Injected = $(if ($injected) { $injected.Name } else { $unknown }); Detected = $(if ($detected) { $detected.Name } else { $unknown })
        InjectedPhase = $injected; DetectedPhase = $detected; Place = $Record; Seq = $defects.Count
    })
}

$bugFields = @('障害内容', '起票日', '状態', '重要度', '障害ID', '原因', '混入工程', '摘出工程')
# 読み取り結果（どのファイル・シートをどの見出しで読んだか）
$readLog = New-Object System.Collections.Generic.List[object]
if ($BugList) {
    $candidates = Read-HeaderCandidates $Headers $bugFields
    $records = Read-HeaderRecords -Inputs (Get-InputFiles $BugList '障害管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $bugFields `
        -Need @('障害内容', @('混入工程', '摘出工程')) -Key @('障害内容') -What '障害管理表' -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
    foreach ($rec in $records) {
        $v = $rec.Values
        Add-Defect 'バグ' ([string]$v['障害内容']) ([string]$v['混入工程']) ([string]$v['摘出工程']) ([string]$v['原因']) ([string]$v['重要度']) ([string]$v['状態']) $rec
    }
}
if ($Findings) {
    # レビューの指摘: 摘出工程は「工程」（レビュー工程）、混入工程は「混入工程」の列（無ければ、レビューした工程で入ったものとみなす）
    $findingFields = @('指摘内容', '設計書', '状態', '重要度', '分類', '工程', '混入工程')
    $candidates = Read-HeaderCandidates $FindingHeaders $findingFields
    if ($candidates['混入工程'].Count -eq 0) { $candidates['混入工程'] = (Read-HeaderCandidates $Headers @('混入工程'))['混入工程'] }
    $records = Read-HeaderRecords -Inputs (Get-InputFiles $Findings '指摘管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $findingFields `
        -Need @('指摘内容', '工程') -Key @('指摘内容') -What '指摘管理表' -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
    foreach ($rec in $records) {
        $v = $rec.Values
        $injected = [string]$v['混入工程']
        if (-not $injected.Trim()) { $injected = [string]$v['工程'] }
        Add-Defect '指摘' ([string]$v['指摘内容']) $injected ([string]$v['工程']) ([string]$v['分類']) ([string]$v['重要度']) ([string]$v['状態']) $rec
    }
}
if ($defects.Count -eq 0) { Write-Warning '欠陥が見つかりません（障害管理表は障害内容と混入工程か摘出工程、指摘管理表は指摘内容と工程の見出しがある表を探します）' }
if ($unreadPhases.Count -gt 0) { Write-Warning ('工程として読めない値があります（工程の CSV の別名に足してください）: {0}' -f ((@($unreadPhases.Keys) | Sort-Object) -join '、')) }

# ---------------------------------------------------------------- 工程ごとの集計


# 摘出すべき工程までに見つかったか（摘出工程の順番が、摘出すべき工程の順番以下）
function Test-InTime {
    param($Defect)
    if (-not $Defect.InjectedPhase -or -not $Defect.DetectedPhase) { return $null }
    return $Defect.DetectedPhase.Index -le $expectedOf[$Defect.InjectedPhase.Name].Index
}

foreach ($d in $defects) { $d | Add-Member -NotePropertyName InTime -NotePropertyValue (Test-InTime $d) }

$results = New-Object System.Collections.Generic.List[object]
foreach ($p in $phaseList) {
    $injected = @($defects | Where-Object { $_.Injected -eq $p.Name })
    $detected = @($defects | Where-Object { $_.Detected -eq $p.Name })
    $inTime = @($injected | Where-Object { $_.InTime -eq $true }).Count
    $leaked = @($injected | Where-Object { $_.InTime -eq $false }).Count
    $rate = $null
    if (($inTime + $leaked) -gt 0) { $rate = [math]::Round($inTime / ($inTime + $leaked), 3) }
    # この工程で見つかったもののうち、前の工程で入ったもの（前の工程から流出してきたもの）
    $fromEarlier = @($detected | Where-Object { $_.InjectedPhase -and $_.InjectedPhase.Index -lt $p.Index }).Count
    if ($injected.Count -eq 0 -and $detected.Count -eq 0) { continue }
    $r = [pscustomobject]@{
        Phase = $p.Name; Expected = $p.Expected; Injected = $injected.Count; FoundInTime = $inTime; Leaked = $leaked; RemovalRate = $rate
        Detected = $detected.Count; FromEarlier = $fromEarlier
    }
    $r.PSObject.TypeNames.Insert(0, 'DefectRemoval.Phase')
    $results.Add($r)
}
$results

# ---------------------------------------------------------------- 出力


# 工程の並び（工程の CSV の順。読めない工程は最後）
$phaseNames = @($phaseList | ForEach-Object { $_.Name })
if (@($defects | Where-Object { $_.Injected -eq $unknown -or $_.Detected -eq $unknown }).Count -gt 0) { $phaseNames += $unknown }
$usedInjected = @($phaseNames | Where-Object { $n = $_; @($defects | Where-Object { $_.Injected -eq $n }).Count -gt 0 })
$usedDetected = @($phaseNames | Where-Object { $n = $_; @($defects | Where-Object { $_.Detected -eq $n }).Count -gt 0 })

# 原因のパレート（多い順と累積の割合）
$pareto = New-Object System.Collections.Generic.List[object]
$withCause = @($defects | Where-Object { $_.Cause })
$cumulative = 0
foreach ($g in @($withCause | Group-Object Cause | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)) {
    $cumulative += $g.Count
    $pareto.Add([pscustomobject]@{ Cause = $g.Name; Count = $g.Count; Ratio = [math]::Round($g.Count / $withCause.Count, 3); Cumulative = [math]::Round($cumulative / $withCause.Count, 3) })
}

if ($OutXlsx) {
    $matrixRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $usedInjected) {
        $values = @($i)
        foreach ($d in $usedDetected) { $values += [double]@($defects | Where-Object { $_.Injected -eq $i -and $_.Detected -eq $d }).Count }
        $values += [double]@($defects | Where-Object { $_.Injected -eq $i }).Count
        $matrixRows.Add(@{ Values = $values })
    }
    $totals = @('合計')
    foreach ($d in $usedDetected) { $totals += [double]@($defects | Where-Object { $_.Detected -eq $d }).Count }
    $totals += [double]$defects.Count
    $matrixRows.Add(@{ Values = $totals })
    $matrixColumns = @(@{ Header = '混入工程 ＼ 摘出工程'; Width = 18 }) + @($usedDetected | ForEach-Object { @{ Header = $_; Width = 10 } }) + @(@{ Header = '合計'; Width = 8 })

    $phaseRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $rate = ''
        if ($null -ne $r.RemovalRate) { $rate = [double]$r.RemovalRate }
        $phaseRows.Add(@{ Values = @($r.Phase, $r.Expected, [double]$r.Injected, [double]$r.FoundInTime, [double]$r.Leaked, $rate, [double]$r.Detected, [double]$r.FromEarlier) })
    }
    $paretoRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in $pareto) { $paretoRows.Add(@{ Values = @($p.Cause, [double]$p.Count, [double]$p.Ratio, [double]$p.Cumulative) }) }
    $leakRows = New-Object System.Collections.Generic.List[object]
    foreach ($d in @($defects | Where-Object { $_.InTime -eq $false } | Sort-Object @{ e = { $_.InjectedPhase.Index } }, @{ e = { $_.DetectedPhase.Index } }, Seq)) {
        $where = '{0} {1} 行目' -f $d.Place.File, $d.Place.Row
        if ($d.Place.Sheet) { $where = '{0} {1}!{2}{3}' -f $d.Place.File, $d.Place.Sheet, (ConvertTo-ColumnName $d.Place.Column), $d.Place.Row }
        $row = @{ Values = @($d.Source, $d.Injected, $d.Detected, $d.Content, $d.Cause, $d.Severity, $where) }
        if ($d.Place.Sheet) { $row.Link = @{ Column = 6; Path = $d.Place.Path; Location = ("'{0}'!{1}{2}" -f $d.Place.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $d.Place.Column), $d.Place.Row) } }
        $leakRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '工程別'; Columns = @(@{ Header = '工程'; Width = 12 }, @{ Header = '摘出すべき工程'; Width = 12 }, @{ Header = '混入'; Width = 7 }, @{ Header = 'それまでに摘出'; Width = 10 }, @{ Header = '流出'; Width = 7 }, @{ Header = '摘出率'; Width = 8 }, @{ Header = 'この工程で摘出'; Width = 10 }, @{ Header = 'うち前の工程から流出'; Width = 12 }); Rows = $phaseRows },
            @{ Name = '工程マトリクス'; Columns = $matrixColumns; Rows = $matrixRows },
            @{ Name = '原因'; Columns = @(@{ Header = '原因'; Width = 24 }, @{ Header = '件数'; Width = 7 }, @{ Header = '割合'; Width = 8 }, @{ Header = '累積の割合'; Width = 10 }); Rows = $paretoRows; Charts = @(@{ Type = 'Column'; Title = '原因（パレート図）'; Category = 0; Series = @(@{ Column = 1 }, @{ Column = 3; Type = 'Line'; Secondary = $true }); YTitle = '件数'; SecondaryTitle = '累積の割合'; SecondaryPercent = $true }) },
            @{ Name = '流出'; Columns = @(@{ Header = '種類'; Width = 6 }, @{ Header = '混入工程'; Width = 10 }, @{ Header = '摘出工程'; Width = 10 }, @{ Header = '内容'; Width = 50 }, @{ Header = '原因'; Width = 14 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '場所'; Width = 40 }); Rows = $leakRows }, (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$bugCount = @($defects | Where-Object Source -eq 'バグ').Count
$findingCount = @($defects | Where-Object Source -eq '指摘').Count
Write-Host ('欠陥 {0} 件（バグ {1}・レビューの指摘 {2}。取下げなどで数えなかったもの {3} 件）' -f $defects.Count, $bugCount, $findingCount, $excludedCount)
$judged = @($defects | Where-Object { $null -ne $_.InTime })
if ($judged.Count -gt 0) {
    $leakedAll = @($judged | Where-Object { $_.InTime -eq $false })
    Write-Host ('摘出すべき工程までに見つかった欠陥: {0} / {1} 件（{2:0.0}%）。流出 {3} 件' -f ($judged.Count - $leakedAll.Count), $judged.Count, (($judged.Count - $leakedAll.Count) * 100.0 / $judged.Count), $leakedAll.Count)
    $designPhases = @($phaseList | Where-Object { $expectedOf[$_.Name].Name -eq $_.Name -and $_.Name -match '設計|要件' } | ForEach-Object { $_.Name })
    # 設計で入り、テストの工程（設計の工程でない、名前に「テスト」「試験」を含む工程）で見つかったもの。後の工程のレビューや製造で見つかったものは数えない
    $designLeak = @($leakedAll | Where-Object { $designPhases -contains $_.Injected -and $designPhases -notcontains $_.Detected -and $_.Detected -match 'テスト|試験' })
    if ($bugCount -gt 0 -and $designLeak.Count -gt 0) { Write-Host ('  設計で入り、テストで見つかった欠陥: {0} 件（{1}）' -f $designLeak.Count, ((@($designLeak | Group-Object Injected | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・')) }
}
if ($pareto.Count -gt 0) { Write-Host ('原因の上位: {0}' -f ((@($pareto | Select-Object -First 3 | ForEach-Object { '{0} {1}（{2:0}%）' -f $_.Cause, $_.Count, ($_.Ratio * 100) })) -join '・')) }
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     障害管理表とレビューの指摘管理表から、欠陥の混入工程 × 摘出工程と、工程ごとの摘出率・流出を出します（Excel 不要）。
#
# .DESCRIPTION
#     欠陥（テストのバグとレビューの指摘）を、どの工程で入り（混入工程）、どの工程で見つかったか（摘出工程）で数えます。
#
#       工程別       工程ごとに、入った欠陥・摘出すべき工程までに見つかったもの・それより後に見つかったもの（流出）・摘出率、
#                    その工程で見つかった欠陥のうち前の工程から流出してきたもの
#       工程マトリクス  混入工程 × 摘出工程の件数
#       原因         原因の分類の多い順と累積の割合（パレート図に使う）
#       流出         摘出すべき工程より後で見つかった欠陥の一覧（障害管理表・指摘管理表の該当行へのリンク付き）
#
#     工程の名前・並び・別名（UT・詳細設計レビュー など）・摘出すべき工程は tools\test-quality\工程.csv で決めます。
#     例: 製造で入った欠陥は単体テストまで、詳細設計で入った欠陥は詳細設計（のレビュー）までに見つけるもの。
#
# .PARAMETER BugList
#     障害管理表のファイルまたはフォルダ（Excel・CSV。複数可）。混入工程・摘出工程の列を読みます。
#
# .PARAMETER Findings
#     レビューの指摘管理表のファイルまたはフォルダ（Excel・CSV。複数可）。工程（レビュー工程）を摘出工程とし、
#     混入工程の列が無ければ、レビューした工程で入った欠陥とみなします。
#
# .PARAMETER Phases
#     工程の CSV（工程・別名・摘出すべき工程）。既定は tools\test-quality\工程.csv。
#
# .PARAMETER Headers
#     障害管理表の見出しの候補の CSV。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER FindingHeaders
#     指摘管理表の見出しの候補の CSV。既定は tools\review-quality\指摘管理表の見出し.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（工程別・工程マトリクス・原因・流出）。
#
# .EXAMPLE
#     .\Measure-DefectRemoval.ps1 -BugList D:\テスト\障害管理表.xlsx -Findings D:\レビュー -OutXlsx D:\工程別の欠陥.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
