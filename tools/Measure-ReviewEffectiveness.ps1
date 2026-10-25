[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string[]] $Findings,

    [Parameter(Mandatory = $true)]
    [string[]] $BugList,

    [string[]] $DesignPath,

    [string] $SizeCsv,

    [string[]] $ReviewLog,

    [string] $Limits,

    [string] $FunctionMap,

    [string] $Phases,

    [string[]] $DesignPhases = @('要件定義', '基本設計', '詳細設計'),

    [string] $Headers,

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Phases) { $Phases = Join-Path $PSScriptRoot 'test-quality/工程.csv' }
if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
foreach ($name in @('Phases', 'Headers', 'OutXlsx', 'FunctionMap', 'SizeCsv', 'Limits')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($DesignPath -and $SizeCsv) { throw '-DesignPath と -SizeCsv は、どちらか一方を指定してください' }

Update-TypeData -TypeName 'ReviewEffectiveness.Document' -DefaultDisplayPropertySet Document, Findings, DesignBugs, FindingDensity, BugDensity, RemovalRate, Zone -Force -WhatIf:$false

# 数えない状態（取下げ・重複など）
$excludedPattern = '取下げ|取り下げ|取消|取り消|却下|対象外|重複|仕様どおり|仕様通り|障害ではない|非障害|バグではない|キャンセル|cancel'

# 指摘密度の判定 × 設計で入ったバグの有無 → ゾーン・要対応・対応の例
$zoneTable = @{
    '少ない|あり' = @('レビューの見逃し', $true, '観点を見直して再レビューし、同じ種類の誤りがほかの設計書に無いか確かめる')
    '基準内|あり' = @('レビューの見逃しの疑い', $true, '見逃したバグの原因を分析し、レビューの観点表に足す')
    '多い|あり'   = @('品質に問題', $true, '設計書の作り直しや、観点を決めた追加のレビューを検討する')
    '|あり'       = @('レビューの見逃しの疑い', $true, '見逃したバグの原因を分析し、レビューの観点表に足す')
    '少ない|なし' = @('テストの確認', $false, 'テストで見つかっていないだけの可能性がある。テストの網羅を確かめる')
    '基準内|なし' = @('良好', $false, '')
    '多い|なし'   = @('良好（レビューで摘出）', $false, '')
    '|なし'       = @('', $false, '')
}

# ---------------------------------------------------------------- レビュー（設計書ごとの指摘・規模）


$reviewArgs = @{ Path = $Findings; AsOf = $AsOf }
if ($DesignPath) { $reviewArgs['DesignPath'] = $DesignPath }
if ($SizeCsv) { $reviewArgs['SizeCsv'] = $SizeCsv }
if ($ReviewLog) { $reviewArgs['ReviewLog'] = $ReviewLog }
if ($Limits) { $reviewArgs['Limits'] = $Limits }
if ($NoRecurse) { $reviewArgs['NoRecurse'] = $true }
if ($NoExcel) { $reviewArgs['NoExcel'] = $true }
$docs = @(& (Join-Path $PSScriptRoot 'Measure-ReviewQuality.ps1') @reviewArgs 6>$null 3>$null)
if ($docs.Count -eq 0) { throw 'レビューの指摘管理表から設計書を読めませんでした' }

# 設計書の無い指摘（設計書の列が無いか空。Measure-ReviewQuality.ps1 はこの名前にまとめる）は、どの設計書の指摘にも数えられないので、
# 件数と、どのファイルの指摘か（指摘管理表を見出しで読み直し、最大 3 つ）を知らせる。バグは対応付けず、ゾーンも付けない
$noDocument = '（設計書なし）'
$noDocFindings = ($docs | Where-Object { $_.Document -eq $noDocument } | Measure-Object -Property Findings -Sum).Sum + 0
if ($noDocFindings -gt 0) {
    $findingFields = @('設計書', '指摘内容', '状態', '重要度', '分類')
    $findingCandidates = Read-HeaderCandidates (Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv') $findingFields
    $findingInputs = Get-InputFiles $Findings '指摘管理表' -NoRecurse:$NoRecurse 3>$null
    $findingRecords = Read-HeaderRecords -Inputs $findingInputs -Candidates $findingCandidates -Fields $findingFields -Need @('指摘内容', @('設計書', '状態', '重要度', '分類')) `
        -Key @('指摘内容') -What '指摘管理表' -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel 3>$null
    $noDocFiles = @($findingRecords | Where-Object { -not ([string]$_.Values['設計書']).Trim() } | ForEach-Object { $_.File } | Select-Object -Unique)
    $examples = ''
    if ($noDocFiles.Count -gt 0) { $examples = ': ' + ((@($noDocFiles | Select-Object -First 3)) -join '・') + $(if ($noDocFiles.Count -gt 3) { ' ほか' } else { '' }) }
    Write-Warning ('設計書の無い指摘が {0} 件あります。どの設計書の指摘にも数えないため、設計書の指摘密度が低く出ます（指摘管理表に設計書の列を足してください）{1}' -f $noDocFindings, $examples)
}

# ---------------------------------------------------------------- 設計で入ったバグ


$phaseList = Read-PhaseList $Phases
$designPhaseNames = @($DesignPhases | ForEach-Object { $p = Resolve-PhaseName $phaseList $_; if ($p) { $p.Name } })
$functionNames = Read-FunctionMap $FunctionMap
$fields = @('障害ID', '障害内容', '機能', '設計書', '混入工程', '状態')
$candidates = Read-HeaderCandidates $Headers $fields
$readLog = New-Object System.Collections.Generic.List[object]
$records = Read-HeaderRecords -Inputs (Get-InputFiles $BugList '障害管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
    -Need @('障害内容', '混入工程', @('機能', '設計書')) -Key @('障害内容') -What '障害管理表' -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog

function Get-TargetName {
    #     障害管理表の設計書・機能の欄の名前。空と同じ意味の書き方（－・なし・空白 など）は空にする
    param([string] $Text)
    $t = ([string]$Text).Trim()
    if ((ConvertTo-Label $t) -match '^([-ー―‐−–—]+|なし|無し|無|空白|空欄|不明|N/?A)$') { return '' }
    return $t
}

$bugs = New-Object System.Collections.Generic.List[object]
# 混入工程を工程として読めないバグ（数えずに知らせる）
$unreadPhases = @{}
$unreadCount = 0
foreach ($rec in $records) {
    $v = $rec.Values
    if (([string]$v['状態']) -match $excludedPattern) { continue }
    $injected = ([string]$v['混入工程']).Trim()
    $phase = Resolve-PhaseName $phaseList $injected
    if ($injected -and -not $phase) { $unreadPhases[$injected] = $true; $unreadCount++ }
    if (-not $phase -or $designPhaseNames -notcontains $phase.Name) { continue }
    $docName = Get-TargetName $v['設計書']
    $funcName = Get-TargetName $v['機能']
    # Target は表示用（機能・設計書）
    $target = (@(@($funcName, $docName) | Where-Object { $_ } | Select-Object -Unique)) -join '・'
    $bugs.Add([pscustomobject]@{ Id = @(ConvertTo-BugIds $v['障害ID'])[0]; Content = ([string]$v['障害内容']).Trim(); DocName = $docName; Function = $funcName; Target = $target; Phase = $phase.Name; Document = $null; Place = $rec })
}
if ($unreadCount -gt 0) { Write-Warning ('混入工程を工程として読めないバグ {0} 件は数えていません（工程の CSV の別名に足してください）: {1}' -f $unreadCount, ((@($unreadPhases.Keys) | Sort-Object) -join '、')) }

# バグの設計書・機能を、レビューの設計書に対応付ける。設計書の欄、機能の欄の順に、対応表で同じ機能になるもの、同じ名前、
# どちらかがもう一方を含む名前の順に探し、設計書がただ 1 つに決まったものにする。設計書の欄を含む名前で探すときは、機能も含む
# 設計書に限る（「画面設計書」のような設計書の種類で、ほかの機能の設計書に対応付けないため）。設計書の無い指摘のまとまりには対応付けない。
# 名前は版数も除いて比べる（ConvertTo-DocNameKey。レビューした …_v1.0 の指摘と、版を上げたあとの …_v1.1 で書いたバグを対応付ける）
$docIndexes = @(for ($i = 0; $i -lt $docs.Count; $i++) { if ($docs[$i].Document -ne $noDocument) { $i } })
$docKeys = @($docs | ForEach-Object { ConvertTo-DocNameKey $_.Document })
$docCanonical = @($docs | ForEach-Object { Resolve-FunctionName $functionNames $_.Document })
foreach ($b in $bugs) {
    foreach ($pair in @(@($b.DocName, $b.Function), @($b.Function, ''))) {
        if (-not $pair[0] -or $b.Document) { continue }
        $canonical = Resolve-FunctionName $functionNames $pair[0]
        $key = ConvertTo-DocNameKey $pair[0]
        $also = ConvertTo-DocNameKey $pair[1]
        $hit = @()
        if ($canonical) { $hit = @(foreach ($i in $docIndexes) { if ($docCanonical[$i] -eq $canonical) { $i } }) }
        if ($hit.Count -ne 1) { $hit = @(foreach ($i in $docIndexes) { if ($docKeys[$i] -eq $key) { $i } }) }
        if ($hit.Count -ne 1) {
            $hit = @(foreach ($i in $docIndexes) {
                    if (($docKeys[$i].Contains($key) -or $key.Contains($docKeys[$i])) -and (-not $also -or $docKeys[$i].Contains($also) -or $also.Contains($docKeys[$i]))) { $i }
                })
        }
        if ($hit.Count -eq 1) { $b.Document = $docs[$hit[0]].Document }
    }
}

# ---------------------------------------------------------------- 設計書ごとの見逃し


$results = New-Object System.Collections.Generic.List[object]
foreach ($d in $docs) {
    $mine = @($bugs | Where-Object { $_.Document -eq $d.Document })
    $level = ''
    foreach ($l in @('少ない', '基準内', '多い')) { if (([string]$d.Judgement).StartsWith($l)) { $level = $l } }
    $bugLevel = $(if ($mine.Count -gt 0) { 'あり' } else { 'なし' })
    $zone = $zoneTable[('{0}|{1}' -f $level, $bugLevel)]
    if ($d.Reviewed -eq $false) {
        $zone = @('', $false, '')
        if ($mine.Count -gt 0) { $zone = @('未レビュー（設計で入ったバグあり）', $true, 'レビューを実施する') }
    }
    if ($d.Document -eq $noDocument) { $zone = @('', $false, '') }
    $bugDensity = $null
    if ($null -ne $d.Size -and $d.Size -gt 0) { $bugDensity = [math]::Round($mine.Count / $d.Size, 3) }
    $rate = $null
    if ($d.Findings + $mine.Count -gt 0) { $rate = [math]::Round($d.Findings / ($d.Findings + $mine.Count), 3) }
    $r = [pscustomobject]@{
        Document = $d.Document; Findings = $d.Findings; DesignBugs = $mine.Count; Size = $d.Size; Unit = $d.Unit
        FindingDensity = $d.Density; FindingJudgement = $d.Judgement; BugDensity = $bugDensity; RemovalRate = $rate
        Zone = $zone[0]; NeedsAction = [bool]$zone[1]; Action = $zone[2]; Reviewed = $d.Reviewed
        BugIds = ((@($mine | ForEach-Object { $_.Id } | Where-Object { $_ })) -join '、')
    }
    $r.PSObject.TypeNames.Insert(0, 'ReviewEffectiveness.Document')
    $results.Add($r)
}
$results

$matchedBugs = @($bugs | Where-Object { $_.Document })
$unmatched = @($bugs | Where-Object { -not $_.Document })
# 全体の摘出率は、未レビューの設計書（指摘もバグも）を除いて出す
$reviewedResults = @($results | Where-Object { $_.Reviewed -ne $false })
$reviewedFindings = ($reviewedResults | Measure-Object -Property Findings -Sum).Sum + 0
$reviewedBugs = ($reviewedResults | Measure-Object -Property DesignBugs -Sum).Sum + 0
$overallRate = $null
if ($reviewedFindings + $reviewedBugs -gt 0) { $overallRate = [math]::Round($reviewedFindings / ($reviewedFindings + $reviewedBugs), 3) }

if ($OutXlsx) {
    $docRows = New-Object System.Collections.Generic.List[object]
    $num = { param($v) if ($null -eq $v) { '' } else { [double]$v } }
    foreach ($r in $results) {
        $docRows.Add(@{ Values = @($r.Document, [double]$r.Findings, [double]$r.DesignBugs, (& $num $r.Size), $r.Unit, (& $num $r.FindingDensity), $r.FindingJudgement, (& $num $r.BugDensity), (& $num $r.RemovalRate), $r.Zone, $r.Action, $r.BugIds) })
    }
    $bugRows = New-Object System.Collections.Generic.List[object]
    foreach ($b in $bugs) {
        $where = '{0} {1} 行目' -f $b.Place.File, $b.Place.Row
        if ($b.Place.Sheet) { $where = '{0} {1}!{2}{3}' -f $b.Place.File, $b.Place.Sheet, (ConvertTo-ColumnName $b.Place.Column), $b.Place.Row }
        $row = @{ Values = @($b.Id, $b.Content, $b.Phase, $b.Target, $(if ($b.Document) { $b.Document } else { '（対応付けられない）' }), $where) }
        if ($b.Place.Sheet) { $row.Link = @{ Column = 5; Path = $b.Place.Path; Location = ("'{0}'!{1}{2}" -f $b.Place.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $b.Place.Column), $b.Place.Row) } }
        $bugRows.Add($row)
    }
    # グラフ: 指摘密度 × 設計で入ったバグの密度（両方ある設計書を最大 50。点に設計書の名前）
    $charts = @()
    $plotted = @(for ($i = 0; $i -lt $results.Count; $i++) { if ($null -ne $results[$i].FindingDensity -and $null -ne $results[$i].BugDensity) { $i } }) | Select-Object -First 50
    if (@($plotted).Count -gt 0) {
        $charts = @(@{ Type = 'Scatter'; Title = '指摘密度 × 設計で入ったバグの密度'; Category = 5; Labels = $true; XTitle = '指摘密度'; YTitle = '設計で入ったバグの密度'; Width = 10; Height = 20
                Series = @($plotted | ForEach-Object { @{ Column = 7; Rows = @($_); Name = $results[$_].Document } }) })
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '設計書別'; Columns = @(@{ Header = '設計書'; Width = 36 }, @{ Header = '指摘'; Width = 7 }, @{ Header = '設計で入ったバグ'; Width = 9 }, @{ Header = '規模'; Width = 8 }, @{ Header = '分母'; Width = 8 }, @{ Header = '指摘密度'; Width = 9 }, @{ Header = '指摘密度の判定'; Width = 16 }, @{ Header = 'バグの密度'; Width = 9 }, @{ Header = 'レビューの摘出率'; Width = 9 }, @{ Header = 'ゾーン'; Width = 18 }, @{ Header = '対応の例'; Width = 50 }, @{ Header = 'バグ'; Width = 20 }); Rows = $docRows; Charts = $charts },
            @{ Name = '設計で入ったバグ'; Columns = @(@{ Header = '障害ID'; Width = 10 }, @{ Header = '障害内容'; Width = 40 }, @{ Header = '混入工程'; Width = 10 }, @{ Header = '機能・設計書'; Width = 24 }, @{ Header = '対応付けた設計書'; Width = 36 }, @{ Header = '場所'; Width = 40 }); Rows = $bugRows },
            (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ('設計書 {0} 件・設計で入ったバグ {1} 件（設計書に対応付けたもの {2} 件）' -f $results.Count, $bugs.Count, $matchedBugs.Count)
if ($null -ne $overallRate) { Write-Host ('設計工程の摘出率（レビューの指摘 ÷ (指摘 + 設計で入ったバグ)）: {0:0.0}%（指摘 {1} 件・バグ {2} 件）' -f ($overallRate * 100), $reviewedFindings, $reviewedBugs) }
$zoned = @($results | Where-Object { $_.Zone })
if ($zoned.Count -gt 0) { Write-Host ('ゾーン: {0}' -f ((@($zoned | Group-InOrder Zone | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・')) }
foreach ($r in @($results | Where-Object NeedsAction)) { Write-Host ('  要対応: {0}（{1}。バグ {2}）{3}' -f $r.Document, $r.Zone, $r.BugIds, $r.Action) }
if ($unmatched.Count -gt 0) { Write-Host ('設計書に対応付けられなかったバグ: {0}（-FunctionMap の対応表に機能と設計書を書くと対応付きます）' -f ((@($unmatched | ForEach-Object { '{0}（{1}）' -f $_.Id, $_.Target })) -join '・')) }
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書ごとに、レビューの指摘と、設計で入ったのにテストで見つかったバグを並べ、レビューで見逃した設計書を見つけます（Excel 不要）。
#
# .DESCRIPTION
#     レビューの指摘管理表を Measure-ReviewQuality.ps1 で集計し、障害管理表のうち混入工程が設計（-DesignPhases）のバグを、
#     機能・設計書の列で設計書に対応付けます。設計書ごとに、指摘密度の判定（基準値と比べた 少ない・基準内・多い）と、
#     設計で入ったバグの有無でゾーンに分けます。
#
#       指摘が少なく、バグあり   レビューの見逃し（要対応）
#       指摘は基準内で、バグあり レビューの見逃しの疑い（要対応）
#       指摘が多く、バグもあり   品質に問題（要対応）
#       指摘が少なく、バグなし   テストの確認（テストで見つかっていないだけの可能性）
#
#     あわせて、設計書ごとと全体の「レビューの摘出率」（レビューの指摘 ÷ (指摘 + 設計で入ったバグ)）を出します。
#
# .PARAMETER Findings
#     レビューの指摘管理表のファイル・フォルダ。
#
# .PARAMETER BugList
#     障害管理表のファイル・フォルダ。「障害内容」「混入工程」と、「機能」か「設計書」の見出しがある表を読みます。
#
# .PARAMETER DesignPath
#     設計書のフォルダ（指摘密度・バグの密度の分母の規模と、設計書の一覧）。
#
# .PARAMETER SizeCsv
#     設計書のページ数などの CSV（-DesignPath の代わりの分母）。
#
# .PARAMETER ReviewLog
#     レビュー記録（渡すと、レビュー記録も指摘も無い設計書を未レビューとします）。
#
# .PARAMETER Limits
#     指摘密度の基準値の CSV（Measure-ReviewQuality.ps1 の -Limits）。
#
# .PARAMETER FunctionMap
#     機能の対応表の CSV（機能・別名・設計書）。障害管理表の機能を設計書に対応付けるのに使います。
#
# .PARAMETER Phases
#     工程の CSV（既定は tools\test-quality\工程.csv）。
#
# .PARAMETER DesignPhases
#     設計とみなす混入工程（既定は 要件定義・基本設計・詳細設計）。
#
# .PARAMETER Headers
#     障害管理表の見出しの候補の CSV（既定は tools\test-quality\見出し.csv）。
#
# .PARAMETER AsOf
#     集計の時点（期限切れの判定などに使う）。既定は今日。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     設計書別（散布図付き）・設計で入ったバグ・読み取り結果を xlsx に保存します。
#
# .EXAMPLE
#     .\Measure-ReviewEffectiveness.ps1 -Findings D:\レビュー -BugList D:\障害管理表.xlsx -DesignPath D:\設計書 -OutXlsx D:\レビューの効き目.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Measure-ReviewQuality.ps1 が必要です。
