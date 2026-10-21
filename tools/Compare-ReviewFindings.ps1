[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $AiPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [string[]] $HumanPath,

    [ValidateRange(0, 100)]
    [int] $RowTolerance = 2,

    [ValidateRange(0.0, 1.0)]
    [double] $MinSimilarity = 0.35,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv' }
foreach ($name in @('Headers', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'ReviewCompare.Pair' -DefaultDisplayPropertySet Status, Document, AiLocation, HumanLocation, Basis, AiContent, HumanContent -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|誤指摘'

# 指摘管理表の見出しの候補に、場所・観点・判定の候補を足す（CSV に書けば CSV が優先）
$fields = @('設計書', '指摘内容', '状態', '重要度', '分類', '場所', 'シート', '観点ID', '判定')
$candidates = Read-HeaderCandidates $Headers $fields
$extra = @{
    '場所'   = @('場所', '箇所', '指摘箇所', '該当箇所', '位置', 'セル', '該当セル', 'セル番地')
    'シート' = @('シート', 'シート名', '該当シート')
    '観点ID' = @('観点ID', '観点No', '観点番号', 'チェック項目ID')
    '判定'   = @('判定', '結果')
}
foreach ($k in $extra.Keys) { if ($candidates[$k].Count -eq 0) { $candidates[$k] = @($extra[$k] | ForEach-Object { ConvertTo-Label $_ }) } }

# ---------------------------------------------------------------- 指摘を読む


function ConvertTo-Bigrams {
    #     文字の 2 文字ずつの組（似ている度合いを比べるため。全角・半角と空白はそろえる）
    param([string] $Text)
    $t = (ConvertTo-Label $Text).ToUpperInvariant()
    $set = New-Object 'System.Collections.Generic.HashSet[string]'
    if ($t.Length -eq 1) { [void]$set.Add($t) }
    for ($i = 0; $i -lt $t.Length - 1; $i++) { [void]$set.Add($t.Substring($i, 2)) }
    return , $set
}

function Get-Similarity {
    #     2 つの指摘内容の似ている度合い（2 文字ずつの組の重なり。0〜1）
    param($A, $B)
    if ($A.Count -eq 0 -or $B.Count -eq 0) { return 0.0 }
    $common = 0
    foreach ($x in $A) { if ($B.Contains($x)) { $common++ } }
    return $common / ($A.Count + $B.Count - $common)
}

# 読み取り結果（どのファイル・シートをどの見出しで読んだか）
$readLog = New-Object System.Collections.Generic.List[object]
function Read-Findings {
    param([string[]] $Paths, [string] $Source)
    $records = Read-HeaderRecords -Inputs (Get-InputFiles $Paths "${Source}の指摘" -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
        -Need @('指摘内容', '設計書') -Key @('指摘内容') -What "${Source}の指摘" -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($rec in $records) {
        $v = $rec.Values
        if (([string]$v['状態']) -match $excludedPattern) { continue }
        # AI のレビュー結果は、判定が OK の行を指摘として数えない
        if (([string]$v['判定']).Normalize($nfkc).Trim() -match '^(OK|○|〇)$') { continue }
        $place = ([string]$v['場所']).Normalize($nfkc).Trim()
        $sheet = ([string]$v['シート']).Trim()
        $row = 0; $col = 0
        $m = [regex]::Match($place, "^\s*'?(?<sheet>[^!']+?)'?\s*!\s*\`$?(?<col>[A-Za-z]{1,3})\`$?(?<row>\d+)")
        $spaced = $false
        if (-not $m.Success) { $m = [regex]::Match($place, '(?<![A-Za-z])\$?(?<col>[A-Za-z]{1,3})\$?(?<row>\d+)(?!\d)'); $spaced = $true }
        if ($m.Success) {
            if ($m.Groups['sheet'].Success) { $sheet = $m.Groups['sheet'].Value.Trim() }
            # 「画面項目定義 M11」のように、シート名とセルを空白で区切って書いた場所
            elseif ($spaced -and -not $sheet -and $m.Index -gt 0) { $sheet = $place.Substring(0, $m.Index).Trim() }
            $row = [int]$m.Groups['row'].Value
            $col = Get-ColumnIndex $m.Groups['col'].Value.ToUpperInvariant()
        }
        elseif (-not $sheet -and $place) { $sheet = $place }
        # 場所は「シート!セル」の形にそろえる（シートの列に書いた指摘も）
        $location = $place
        if ($place -and ([string]$v['シート']).Trim() -and $place -notmatch '!') { $location = '{0}!{1}' -f ([string]$v['シート']).Trim(), $place }
        if (-not $location -and $sheet) { $location = $sheet }
        $document = ([string]$v['設計書']).Trim()
        if ($document -match '\.(xlsx|xlsm|xls|xlsb|pdf)$') { $document = [System.IO.Path]::GetFileNameWithoutExtension($document) }
        $list.Add([pscustomobject]@{
            Source = $Source; Document = $document; DocKey = (ConvertTo-NameKey $v['設計書']); Content = ([string]$v['指摘内容']).Trim()
            Location = $location; SheetKey = (ConvertTo-Label $sheet).ToUpperInvariant(); Row = $row; Col = $col
            PointId = ([string]$v['観点ID']).Trim(); Bigrams = (ConvertTo-Bigrams $v['指摘内容']); Place = $rec; Matched = $null; Seq = $list.Count
        })
    }
    return , $list
}

$ai = Read-Findings $AiPath 'AI'
$human = Read-Findings $HumanPath '人'
if ($ai.Count -eq 0) { Write-Warning 'AI の指摘が見つかりません（指摘内容と設計書の見出しがある表を探します）' }
if ($human.Count -eq 0) { Write-Warning '人の指摘が見つかりません（指摘内容と設計書の見出しがある表を探します）' }

# ---------------------------------------------------------------- 突き合わせる


# 同じ設計書の組み合わせについて、場所が近い（同じシートで行が RowTolerance 以内）か、内容が似ている組を、良い順に 1 対 1 で決める
$candidatePairs = New-Object System.Collections.Generic.List[object]
foreach ($a in $ai) {
    foreach ($h in $human) {
        if (-not $a.DocKey -or $a.DocKey -ne $h.DocKey) { continue }
        $similarity = Get-Similarity $a.Bigrams $h.Bigrams
        $near = $a.SheetKey -and $a.SheetKey -eq $h.SheetKey -and $a.Row -gt 0 -and $h.Row -gt 0 -and [math]::Abs($a.Row - $h.Row) -le $RowTolerance
        if ($near) { $candidatePairs.Add([pscustomobject]@{ Ai = $a; Human = $h; Score = 1 + $similarity; Basis = '場所'; Similarity = $similarity }) }
        elseif ($similarity -ge $MinSimilarity) { $candidatePairs.Add([pscustomobject]@{ Ai = $a; Human = $h; Score = $similarity; Basis = '内容'; Similarity = $similarity }) }
    }
}
$pairs = New-Object System.Collections.Generic.List[object]
# 同じ良さの組は読んだ順にする（Windows PowerShell 5.1 の Sort-Object は同じ値の順番を保たない）
foreach ($p in @($candidatePairs | Sort-Object @{ e = 'Score'; Descending = $true }, @{ e = { $_.Ai.Seq } }, @{ e = { $_.Human.Seq } })) {
    if ($p.Ai.Matched -or $p.Human.Matched) { continue }
    $p.Ai.Matched = $p.Human
    $p.Human.Matched = $p.Ai
    $pairs.Add($p)
}

function New-Pair {
    param([string] $Status, $Ai, $Human, [string] $Basis, $Similarity, [int] $Seq)
    $doc = $(if ($Human) { $Human.Document } else { $Ai.Document })
    $similarityValue = $null
    if ($null -ne $Similarity) { $similarityValue = [math]::Round($Similarity, 2) }
    $o = [pscustomobject]@{
        Status = $Status; Document = $doc; Basis = $Basis; Similarity = $similarityValue
        PointId = $(if ($Ai) { $Ai.PointId } else { '' })
        AiLocation = $(if ($Ai) { $Ai.Location } else { '' }); AiContent = $(if ($Ai) { $Ai.Content } else { '' })
        HumanLocation = $(if ($Human) { $Human.Location } else { '' }); HumanContent = $(if ($Human) { $Human.Content } else { '' })
        AiPlace = $(if ($Ai) { $Ai.Place } else { $null }); HumanPlace = $(if ($Human) { $Human.Place } else { $null }); Seq = $Seq
    }
    $o.PSObject.TypeNames.Insert(0, 'ReviewCompare.Pair')
    return $o
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($p in $pairs) { $results.Add((New-Pair '一致' $p.Ai $p.Human $p.Basis $p.Similarity $results.Count)) }
foreach ($a in @($ai | Where-Object { -not $_.Matched })) { $results.Add((New-Pair 'AI だけ' $a $null '' $null $results.Count)) }
foreach ($h in @($human | Where-Object { -not $_.Matched })) { $results.Add((New-Pair '人だけ' $null $h '' $null $results.Count)) }
$sorted = @($results | Sort-Object @{ e = { ConvertTo-NameKey $_.Document } }, @{ e = { @{ '一致' = 0; 'AI だけ' = 1; '人だけ' = 2 }[$_.Status] } }, Seq)
$sorted

# ---------------------------------------------------------------- 集計


$matched = $pairs.Count
$byPoint = @($ai | Where-Object { $_.PointId } | Group-Object PointId | ForEach-Object {
        $hit = @($_.Group | Where-Object { $_.Matched }).Count
        [pscustomobject]@{ PointId = $_.Name; Ai = $_.Count; Matched = $hit; Rate = [math]::Round($hit / $_.Count, 3) }
    } | Sort-Object PointId)
$docKeys = @(($ai.ToArray() + $human.ToArray()) | ForEach-Object { $_.DocKey } | Sort-Object -Unique)
$byDocument = @($docKeys | ForEach-Object {
        $k = $_
        $a = @($ai | Where-Object { $_.DocKey -eq $k }); $h = @($human | Where-Object { $_.DocKey -eq $k })
        $name = @(@($h) + @($a) | ForEach-Object { $_.Document })[0]
        $hit = @($a | Where-Object { $_.Matched }).Count
        [pscustomobject]@{ Document = $name; Ai = $a.Count; Human = $h.Count; Matched = $hit; AiOnly = $a.Count - $hit; HumanOnly = $h.Count - $hit }
    })

if ($OutXlsx) {
    $pairRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $sorted) {
        $similarity = ''
        if ($null -ne $r.Similarity) { $similarity = [double]$r.Similarity }
        $row = @{ Values = @($r.Status, $r.Document, $r.PointId, $r.AiLocation, $r.AiContent, $r.HumanLocation, $r.HumanContent, $r.Basis, $similarity) }
        # 人の指摘があれば人の指摘管理表の行、無ければ AI のレビュー結果の行を開く
        $place = $r.HumanPlace
        $column = 5
        if (-not $place) { $place = $r.AiPlace; $column = 3 }
        if ($place -and $place.Sheet) { $row.Link = @{ Column = $column; Path = $place.Path; Location = ("'{0}'!{1}{2}" -f $place.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $place.Column), $place.Row) } }
        $pairRows.Add($row)
    }
    $pointRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in $byPoint) { $pointRows.Add(@{ Values = @($p.PointId, [double]$p.Ai, [double]$p.Matched, [double]$p.Rate) }) }
    $docRows = New-Object System.Collections.Generic.List[object]
    foreach ($d in $byDocument) { $docRows.Add(@{ Values = @($d.Document, [double]$d.Ai, [double]$d.Human, [double]$d.Matched, [double]$d.AiOnly, [double]$d.HumanOnly) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '突き合わせ'; Columns = @(@{ Header = '結果'; Width = 8 }, @{ Header = '設計書'; Width = 30 }, @{ Header = '観点ID'; Width = 8 }, @{ Header = 'AI の場所'; Width = 20 }, @{ Header = 'AI の指摘'; Width = 40 }, @{ Header = '人の場所'; Width = 20 }, @{ Header = '人の指摘'; Width = 40 }, @{ Header = '決め手'; Width = 7 }, @{ Header = '似ている度合い'; Width = 9 }); Rows = $pairRows },
            @{ Name = '観点別'; Columns = @(@{ Header = '観点ID'; Width = 10 }, @{ Header = 'AI の指摘'; Width = 9 }, @{ Header = '人と一致'; Width = 9 }, @{ Header = '一致率'; Width = 8 }); Rows = $pointRows },
            @{ Name = '設計書別'; Columns = @(@{ Header = '設計書'; Width = 36 }, @{ Header = 'AI の指摘'; Width = 9 }, @{ Header = '人の指摘'; Width = 9 }, @{ Header = '一致'; Width = 7 }, @{ Header = 'AI だけ'; Width = 8 }, @{ Header = '人だけ'; Width = 8 }); Rows = $docRows }, (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$aiRate = 0.0; $humanRate = 0.0
if ($ai.Count -gt 0) { $aiRate = $matched * 100.0 / $ai.Count }
if ($human.Count -gt 0) { $humanRate = $matched * 100.0 / $human.Count }
Write-Host ('AI の指摘 {0} 件: 人と一致 {1}（{2:0.0}%）・AI だけ {3}' -f $ai.Count, $matched, $aiRate, ($ai.Count - $matched))
Write-Host ('人の指摘 {0} 件: AI と一致 {1}（{2:0.0}%）・人だけ {3}' -f $human.Count, $matched, $humanRate, ($human.Count - $matched))
if ($byPoint.Count -gt 0) {
    $low = @($byPoint | Where-Object { $_.Ai -ge 2 } | Sort-Object Rate, PointId | Select-Object -First 3)
    if ($low.Count -gt 0) { Write-Host ('一致率の低い観点: {0}' -f ((@($low | ForEach-Object { '{0} {1}/{2}' -f $_.PointId, $_.Matched, $_.Ai })) -join '・')) }
}
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     AI のレビュー（Invoke-DesignReview.ps1 の指摘一覧）と人のレビューの指摘管理表を突き合わせ、一致・AI だけ・人だけに分けます（Excel 不要）。
#
# .DESCRIPTION
#     同じ設計書の指摘どうしを、場所が近い（同じシートで行の差が -RowTolerance 以内）か、指摘内容が似ている
#     （2 文字ずつの組の重なりが -MinSimilarity 以上）組から順に 1 対 1 で対応付けます。
#
#       突き合わせ  一致・AI だけ・人だけの一覧（指摘管理表・レビュー結果の該当行へのリンク付き）
#       観点別      AI の指摘のうち人と一致した割合（観点表の判定基準の見直しに使う）
#       設計書別    設計書ごとの件数
#
#     AI の指摘のうち人と一致した割合と、人の指摘のうち AI と一致した割合を出します。
#     Invoke-DesignReview.ps1 の -TextFormat（Rows / Markdown）や観点表を変えたときに、同じ人の指摘と比べて効果を確かめられます。
#
# .PARAMETER AiPath
#     AI のレビュー結果のファイル・フォルダ（Invoke-DesignReview.ps1 のレビュー結果.xlsx など）。判定が OK の行は数えません。
#
# .PARAMETER HumanPath
#     人のレビューの指摘管理表のファイル・フォルダ（Excel・CSV）。
#
# .PARAMETER RowTolerance
#     場所が近いとみなす行の差（既定 2）。
#
# .PARAMETER MinSimilarity
#     場所が違うときに、内容が似ているとみなす度合い（0〜1。既定 0.35）。
#
# .PARAMETER Headers
#     指摘管理表の見出しの候補の CSV。既定は tools\review-quality\指摘管理表の見出し.csv（場所・シート・観点ID・判定の候補は足して使います）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（突き合わせ・観点別・設計書別）。
#
# .EXAMPLE
#     .\Compare-ReviewFindings.ps1 D:\design-review\レビュー結果.xlsx D:\レビュー\指摘管理表.xlsx -OutXlsx D:\AIと人の比較.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     対応付けは機械的な目安です。「AI だけ」には、人が見落とした正しい指摘と、AI の誤った指摘の両方が入ります。
