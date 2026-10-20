[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $DesignPath,

    [string] $SizeCsv,

    [string] $Limits,

    [string] $Headers,

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$sizeScript = Join-Path $PSScriptRoot 'Measure-DesignSize.ps1'

if (-not $Limits) { $Limits = Join-Path $PSScriptRoot 'review-quality/基準値_サンプル.csv' }
if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv' }
foreach ($name in @('Limits', 'Headers', 'SizeCsv', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($DesignPath -and $SizeCsv) { throw '-DesignPath と -SizeCsv は、どちらか一方を指定してください' }

Update-TypeData -TypeName 'ReviewQuality.Document' -DefaultDisplayPropertySet Document, Findings, Open, Overdue, Density, Judgement -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# 対応が終わった・指摘として数えない状態
$closedPattern = '完了|済|クローズ|close|done|対応不要|不要|対処なし'
$excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|誤指摘'

# ---------------------------------------------------------------- 設定


$candidates = @{}
foreach ($r in (Read-CsvFile $Headers @('項目', '見出し') '指摘管理表の見出し')) {
    $field = ([string]$r.'項目').Trim()
    if ($field) { $candidates[$field] = @(([string]$r.'見出し') -split '\|' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ }) }
}
$fields = @('設計書', '指摘内容', '分類', '重要度', '状態', '対応要否', '起票日', '期限', '完了日', '対応者', '工程')
foreach ($f in $fields) { if (-not $candidates.ContainsKey($f)) { $candidates[$f] = @() } }
if ($candidates['指摘内容'].Count -eq 0) { throw "指摘管理表の見出しに「指摘内容」の行がありません: $Headers" }

$limitRows = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Limits @('分母', '下限', '上限') '基準値')) {
    $unit = ([string]$r.'分母').Trim()
    if (-not $unit) { continue }
    $process = ''
    if ($r.PSObject.Properties['工程']) { $process = ([string]$r.'工程').Trim() }
    $low = $null; $high = $null
    if (([string]$r.'下限').Trim()) { $low = [double]$r.'下限' }
    if (([string]$r.'上限').Trim()) { $high = [double]$r.'上限' }
    $limitRows.Add([pscustomobject]@{ Unit = $unit; Process = $process; Low = $low; High = $high })
}

function ConvertTo-DocKey {
    #     設計書の名前を比べる形（フォルダと拡張子を除き、全角・半角と空白をそろえる）
    param([string] $Name)
    $n = ([string]$Name).Trim()
    $n = $n.Split(@('\', '/'))[-1]
    if ($n -match '\.(xlsx|xlsm|xls|xlsb|pdf|docx?|pptx?)$') { $n = [System.IO.Path]::GetFileNameWithoutExtension($n) }
    return (ConvertTo-Label $n).ToUpperInvariant()
}

function ConvertTo-Date {
    #     日付の文字（2026/10/1・2026-10-01・2026年10月1日・10/1）を日付にする（無ければ $null）
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return $null }
    $m = [regex]::Match($t, '(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})')
    try {
        if ($m.Success) { return New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value) }
        $m = [regex]::Match($t, '^(\d{1,2})\s*[/月]\s*(\d{1,2})日?$')
        if ($m.Success) { return New-Object DateTime $AsOf.Year, ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value) }
    }
    catch { }
    return $null
}

# ---------------------------------------------------------------- 指摘管理表を読む


$findings = New-Object System.Collections.Generic.List[object]
$excluded = 0

function Add-Finding {
    param([hashtable] $Values, $Place)
    if (-not ([string]$Values['指摘内容']).Trim()) { return }
    $status = ([string]$Values['状態']).Trim()
    $required = ([string]$Values['対応要否']).Trim()
    if ($status -match $excludedPattern) { $script:excluded++; return }
    $closedOn = ConvertTo-Date $Values['完了日']
    $closed = ($status -match $closedPattern) -or ($required -match '^(不要|否|×|-)$') -or ((-not $status) -and $closedOn)
    $due = ConvertTo-Date $Values['期限']
    $doc = ([string]$Values['設計書']).Trim()
    if (-not $doc) { $doc = '（設計書なし）' }
    $findings.Add([pscustomobject]@{
        Document = $doc; Key = (ConvertTo-DocKey $doc); Content = ([string]$Values['指摘内容']).Trim()
        Category = ([string]$Values['分類']).Trim(); Severity = ([string]$Values['重要度']).Trim(); Status = $status
        Owner = ([string]$Values['対応者']).Trim(); Process = ([string]$Values['工程']).Trim()
        Opened = (ConvertTo-Date $Values['起票日']); Due = $due; ClosedOn = $closedOn; Closed = [bool]$closed
        Overdue = ((-not $closed) -and $due -and $due -lt $AsOf); Place = $Place
    })
}

function Find-FirstLabel {
    param($Labels, [string[]] $Names)
    foreach ($n in $Names) { if ($Labels -contains $n) { return $n } }
    return $null
}

$excelFiles = New-Object System.Collections.Generic.List[string]
$csvFiles = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            if (@(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Where-Object { $_.Extension -match '^\.xls[xmb]?$' -and -not $_.Name.StartsWith('~$') }).Count -gt 0) { $excelFiles.Add($item.FullName) }
            foreach ($c in @(Get-ChildItem -LiteralPath $item.FullName -File -Filter '*.csv' -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { $csvFiles.Add($c) }
        }
        elseif ($item.Extension -eq '.csv') { $csvFiles.Add($item) }
        else { $excelFiles.Add([System.Management.Automation.WildcardPattern]::Escape($item.FullName)) }
    }
}

if ($excelFiles.Count -gt 0) {
    $searchArgs = @{ Pattern = '\S'; Path = $excelFiles.ToArray(); Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
    $books = ConvertTo-DesignBooks $hits
    foreach ($book in $books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # 見出しの行: 指摘内容と、設計書・状態・重要度・分類のどれかの見出しがある行
            $map = $null
            foreach ($row in $grid.Rows) {
                $labels = @($grid.Labels[$row].Values)
                $m = @{}
                foreach ($f in $fields) { $l = Find-FirstLabel $labels $candidates[$f]; if ($l) { $m[$f] = $l } }
                if ($m['指摘内容'] -and @(@('設計書', '状態', '重要度', '分類') | Where-Object { $m[$_] }).Count -gt 0) { $map = $m; break }
            }
            if (-not $map) { continue }
            foreach ($t in (Find-DesignTables $grid @($map.Values) $map['指摘内容'])) {
                foreach ($dr in $t.Rows) {
                    $v = @{}
                    foreach ($f in $map.Keys) { $v[$f] = [string]$dr.Cells[$t.ColumnOf[$map[$f]]] }
                    $col = $t.ColumnOf[$map['指摘内容']]
                    Add-Finding $v ([pscustomobject]@{ File = $book.File; Path = $book.Path; Sheet = $sheetName; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $col), $dr.Row) })
                }
            }
        }
    }
}
foreach ($c in $csvFiles) {
    $rows = Read-CsvFile $c.FullName @() '指摘管理表'
    if ($rows.Count -eq 0) { continue }
    $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $map = @{}
    foreach ($f in $fields) { foreach ($n in $names) { if (-not $map[$f] -and $candidates[$f] -contains (ConvertTo-Label $n)) { $map[$f] = $n } } }
    if (-not $map['指摘内容']) { Write-Warning ("{0}: 指摘内容の列が見つかりません" -f $c.Name); continue }
    $line = 1
    foreach ($r in $rows) {
        $line++
        $v = @{}
        foreach ($f in $map.Keys) { $v[$f] = [string]$r.($map[$f]) }
        Add-Finding $v ([pscustomobject]@{ File = $c.Name; Path = $c.FullName; Sheet = ''; Cell = ('{0} 行目' -f $line) })
    }
}

# ---------------------------------------------------------------- 規模（分母）


$sizes = [ordered]@{}   # 設計書の比較用の名前 → @{ Name; Size; Unit }
if ($SizeCsv) {
    $rows = Read-CsvFile $SizeCsv @() '規模の CSV'
    if ($rows.Count -gt 0) {
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $nameCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(ファイル|設計書|成果物|ファイル名|名前|対象)$' })[0]
        $sizeCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(ページ数|規模|サイズ|件数|ページ)$' })[0]
        if (-not $nameCol -or -not $sizeCol) { throw "規模の CSV には、ファイル（設計書）の列と、ページ数（規模）の列が必要です: $SizeCsv" }
        foreach ($r in $rows) {
            if ($r.PSObject.Properties['結果'] -and ([string]$r.'結果').Trim() -and ([string]$r.'結果').Trim() -ne '成功') { continue }
            $n = 0.0
            if (-not [double]::TryParse(([string]$r.$sizeCol).Trim(), [ref]$n)) { continue }
            $k = ConvertTo-DocKey $r.$nameCol
            if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = ([string]$r.$nameCol).Split(@('\', '/'))[-1]; Size = 0.0; Unit = $sizeCol } }
            $sizes[$k].Size += $n
        }
    }
}
elseif ($DesignPath) {
    $sizeArgs = @{ Path = $DesignPath }
    if ($NoRecurse) { $sizeArgs['NoRecurse'] = $true }
    if ($NoExcel) { $sizeArgs['NoExcel'] = $true }
    foreach ($item in @(& $sizeScript @sizeArgs 6>$null 3>$null)) {
        $k = ConvertTo-DocKey $item.File
        if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = $item.File; Size = 0.0; Unit = '規模' } }
        $sizes[$k].Size += [double]$item.Size
    }
}

function Find-SizeKey {
    #     指摘の設計書の名前に対応する規模（同じ名前、またはどちらかがもう一方を含む名前がただ 1 つ）
    param([string] $Key)
    if ($sizes.Contains($Key)) { return $Key }
    $hit = @($sizes.Keys | Where-Object { $_.Contains($Key) -or $Key.Contains($_) })
    if ($hit.Count -eq 1) { return $hit[0] }
    return $null
}

# ---------------------------------------------------------------- 集計


function Format-Breakdown {
    #     「高 3・中 5」のような内訳（多い順）
    param($Items, [string] $Property)
    $groups = @($Items | Where-Object { $_.$Property } | Group-Object $Property | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)
    return (@($groups | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '・')
}

$docs = [ordered]@{}
foreach ($f in $findings) {
    $sk = Find-SizeKey $f.Key
    $k = $f.Key
    if ($sk) { $k = $sk }
    if (-not $docs.Contains($k)) { $docs[$k] = [pscustomobject]@{ Name = $f.Document; SizeKey = $sk; Items = (New-Object System.Collections.Generic.List[object]) } }
    $docs[$k].Items.Add($f)
}
foreach ($k in $sizes.Keys) { if (-not $docs.Contains($k)) { $docs[$k] = [pscustomobject]@{ Name = $sizes[$k].Name; SizeKey = $k; Items = (New-Object System.Collections.Generic.List[object]) } } }

$results = New-Object System.Collections.Generic.List[object]
foreach ($k in $docs.Keys) {
    $d = $docs[$k]
    $items = $d.Items
    $size = $null; $unit = ''; $density = $null; $judgement = ''
    if ($d.SizeKey) { $size = $sizes[$d.SizeKey].Size; $unit = $sizes[$d.SizeKey].Unit }
    $process = @($items | Where-Object { $_.Process } | Group-Object Process | Sort-Object Count -Descending | ForEach-Object { $_.Name })[0]
    if ($null -ne $size -and $size -gt 0) {
        $density = [math]::Round($items.Count / $size, 3)
        $limit = @($limitRows | Where-Object { $_.Unit -eq $unit -and $process -and $_.Process -eq $process })[0]
        if (-not $limit) { $limit = @($limitRows | Where-Object { $_.Unit -eq $unit -and -not $_.Process })[0] }
        if ($limit) {
            if ($null -ne $limit.Low -and $density -lt $limit.Low) { $judgement = '少ない（レビュー不足の疑い）' }
            elseif ($null -ne $limit.High -and $density -gt $limit.High) { $judgement = '多い（品質に問題の疑い）' }
            else { $judgement = '基準内' }
        }
    }
    elseif ($sizes.Count -gt 0) { $judgement = '規模なし' }
    $r = [pscustomobject]@{
        Document = $d.Name; Findings = $items.Count; Open = @($items | Where-Object { -not $_.Closed }).Count; Overdue = @($items | Where-Object Overdue).Count
        Size = $size; Unit = $unit; Density = $density; Judgement = $judgement; Process = $process
        BySeverity = (Format-Breakdown $items 'Severity'); ByCategory = (Format-Breakdown $items 'Category')
    }
    $r.PSObject.TypeNames.Insert(0, 'ReviewQuality.Document')
    $results.Add($r)
}
$results = @($results | Sort-Object @{ e = { $_.Findings }; Descending = $true }, Document)
$results

# 日ごとの起票・完了（収束の確認）
$trend = New-Object System.Collections.Generic.List[object]
$dates = @($findings | ForEach-Object { $_.Opened; $_.ClosedOn } | Where-Object { $_ } | Sort-Object -Unique)
$cumOpen = 0; $cumClosed = 0
foreach ($day in $dates) {
    $o = @($findings | Where-Object { $_.Opened -eq $day }).Count
    $c = @($findings | Where-Object { $_.ClosedOn -eq $day }).Count
    $cumOpen += $o; $cumClosed += $c
    $trend.Add([pscustomobject]@{ Date = $day; Opened = $o; Closed = $c; TotalOpened = $cumOpen; TotalClosed = $cumClosed; Remaining = $cumOpen - $cumClosed })
}

if ($OutXlsx) {
    $docRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $size = ''; $density = ''
        if ($null -ne $r.Size) { $size = [double]$r.Size }
        if ($null -ne $r.Density) { $density = [double]$r.Density }
        $docRows.Add(@{ Values = @($r.Document, [double]$r.Findings, [double]$r.Open, [double]$r.Overdue, $size, $r.Unit, $density, $r.Judgement, $r.BySeverity, $r.ByCategory) })
    }
    $breakRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in @(@('分類', 'Category'), @('重要度', 'Severity'), @('対応者', 'Owner'), @('工程', 'Process'), @('状態', 'Status'))) {
        foreach ($g in @($findings | Where-Object { $_.($p[1]) } | Group-Object $p[1] | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)) {
            $breakRows.Add(@{ Values = @($p[0], $g.Name, [double]$g.Count, [double]@($g.Group | Where-Object { -not $_.Closed }).Count, [double]@($g.Group | Where-Object Overdue).Count) })
        }
    }
    $openRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in @($findings | Where-Object { -not $_.Closed } | Sort-Object @{ e = { -not $_.Overdue } }, @{ e = { if ($_.Due) { $_.Due } else { [datetime]::MaxValue } } })) {
        $due = ''
        if ($f.Due) { $due = $f.Due.ToString('yyyy/MM/dd') }
        $overdue = ''
        if ($f.Overdue) { $overdue = '期限切れ' }
        $row = @{ Values = @($f.Document, $f.Content, $f.Severity, $f.Category, $f.Owner, $due, $overdue, $f.Status, ('{0} {1}' -f $f.Place.File, $f.Place.Cell).Trim()) }
        if ($f.Place.Sheet) { $row.Link = @{ Column = 8; Path = $f.Place.Path; Location = ("'{0}'!{1}" -f $f.Place.Sheet.Replace("'", "''"), $f.Place.Cell) } }
        $openRows.Add($row)
    }
    $trendRows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $trend) { $trendRows.Add(@{ Values = @($t.Date.ToString('yyyy/MM/dd'), [double]$t.Opened, [double]$t.Closed, [double]$t.TotalOpened, [double]$t.TotalClosed, [double]$t.Remaining) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '設計書別'; Columns = @(@{ Header = '設計書'; Width = 40 }, @{ Header = '指摘'; Width = 7 }, @{ Header = '未完了'; Width = 7 }, @{ Header = '期限切れ'; Width = 8 }, @{ Header = '規模'; Width = 8 }, @{ Header = '分母'; Width = 9 }, @{ Header = '指摘密度'; Width = 9 }, @{ Header = '判定'; Width = 26 }, @{ Header = '重要度'; Width = 20 }, @{ Header = '分類'; Width = 30 }); Rows = $docRows },
            @{ Name = '内訳'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '値'; Width = 20 }, @{ Header = '指摘'; Width = 7 }, @{ Header = '未完了'; Width = 7 }, @{ Header = '期限切れ'; Width = 8 }); Rows = $breakRows },
            @{ Name = '未完了'; Columns = @(@{ Header = '設計書'; Width = 30 }, @{ Header = '指摘内容'; Width = 50 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '分類'; Width = 10 }, @{ Header = '対応者'; Width = 10 }, @{ Header = '期限'; Width = 11 }, @{ Header = '期限切れ'; Width = 8 }, @{ Header = '状態'; Width = 9 }, @{ Header = '指摘管理表'; Width = 36 }); Rows = $openRows },
            @{ Name = '推移'; Columns = @(@{ Header = '日付'; Width = 11 }, @{ Header = '起票'; Width = 7 }, @{ Header = '完了'; Width = 7 }, @{ Header = '起票の累計'; Width = 10 }, @{ Header = '完了の累計'; Width = 10 }, @{ Header = '未完了'; Width = 7 }); Rows = $trendRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$open = @($findings | Where-Object { -not $_.Closed }).Count
$overdue = @($findings | Where-Object Overdue).Count
Write-Host ('指摘 {0} 件（未完了 {1} 件・うち期限切れ {2} 件。取下げなどで数えなかったもの {3} 件）/ 設計書 {4} 件（{5} 時点）' -f $findings.Count, $open, $overdue, $excluded, $results.Count, $AsOf.ToString('yyyy/MM/dd'))
foreach ($p in @(@('重要度', 'Severity'), @('分類', 'Category'))) {
    $text = Format-Breakdown $findings $p[1]
    if ($text) { Write-Host ('{0}: {1}' -f $p[0], $text) }
}
$flagged = @($results | Where-Object { $_.Judgement -like '少ない*' -or $_.Judgement -like '多い*' })
if ($flagged.Count -gt 0) { Write-Host ('基準値から外れた設計書: {0} 件' -f $flagged.Count) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     レビューの指摘管理表（Excel・CSV）を集計し、設計書ごとの指摘件数・指摘密度・未完了・期限切れを出します（Excel 不要）。
#
# .DESCRIPTION
#     指摘管理表から、指摘内容・設計書・分類・重要度・対応状況・起票日・期限・完了日・対応者などの列を見出しの文字で探して読み、
#     次を集計します。Invoke-DesignReview.ps1 のレビュー結果（指摘一覧）もそのまま読めます。
#
#       設計書別   指摘件数・未完了・期限切れ、規模あたりの指摘件数（指摘密度）と基準値との比較、重要度・分類の内訳
#       内訳       分類・重要度・対応者・工程・状態ごとの件数
#       未完了     対応が終わっていない指摘の一覧（期限切れから順。指摘管理表の該当行へのリンク付き）
#       推移       日ごとの起票・完了の件数と、未完了の残り（収束の確認）
#
#     指摘密度の分母は、-SizeCsv（Export-ExcelToPdf.ps1 -LogPath のページ数など）か、-DesignPath（Measure-DesignSize.ps1 の規模）で渡します。
#     基準値（下限・上限）から外れた設計書は「少ない（レビュー不足の疑い）」「多い（品質に問題の疑い）」と判定します。
#
# .PARAMETER Path
#     指摘管理表のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER DesignPath
#     設計書のフォルダ・ファイル。Measure-DesignSize.ps1 で数えた規模（項目数・カラム数など）を指摘密度の分母にします。
#
# .PARAMETER SizeCsv
#     設計書ごとの規模の CSV（ファイル・ページ数などの列）。Export-ExcelToPdf.ps1 -CountOnly -LogPath の CSV をそのまま使えます。
#
# .PARAMETER Limits
#     指摘密度の基準値の CSV（分母・工程・下限・上限）。既定は tools\review-quality\基準値_サンプル.csv（数値は例です）。
#
# .PARAMETER Headers
#     指摘管理表の見出しの CSV（項目と、見出しの候補）。既定は tools\review-quality\指摘管理表の見出し.csv。
#
# .PARAMETER AsOf
#     期限切れを判定する日。既定は今日。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（設計書別・内訳・未完了・推移）。
#
# .EXAMPLE
#     .\Measure-ReviewQuality.ps1 D:\レビュー\指摘管理表.xlsx -DesignPath D:\設計書 -OutXlsx D:\品質分析.xlsx
#
# .EXAMPLE
#     # PDF のページ数を分母にする
#     .\Export-ExcelToPdf.ps1 -Path D:\設計書_PDF -CountOnly -LogPath D:\pages.csv
#     .\Measure-ReviewQuality.ps1 D:\レビュー -SizeCsv D:\pages.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1（-DesignPath では Measure-DesignSize.ps1 も）が必要です。
#     基準値のサンプルの数値は例です。社内の実績・標準値に合わせて直してください。
