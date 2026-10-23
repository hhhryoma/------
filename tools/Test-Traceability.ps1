[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $DesignPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [string[]] $TestPath,

    [string[]] $Requirement,

    [string] $Rules,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Rules) { $Rules = Join-Path $PSScriptRoot 'traceability/網羅のルール.csv' }
foreach ($name in @('Rules', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'Traceability.Gap' -DefaultDisplayPropertySet Kind, Scope, Id, Name, Judgement -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# テスト仕様書で、行の機能（画面）を表す列の見出し。空の行は上の行の値を引き継ぐ
$functionHeaders = @('機能', '機能名', '機能ID', '対象機能', 'テスト対象', '画面', '画面ID', '画面名', '対象画面', '対象画面ID')
# 要件一覧の見出し
$requirementIdHeaders = @('要件ID', '要件No', '要求ID', '要求No', '要件番号', '要求番号', 'ID', 'No')
$requirementNameHeaders = @('要件名', '要件', '要求名', '要求', '要件内容', '要求内容', '機能要件', '内容', '概要')

# ---------------------------------------------------------------- ルール


function ConvertTo-Field {
    #     ルールの列・名前・範囲（「@見出し」はシートの見出しの右の値、それ以外は表の列）
    param([string] $Text)
    $t = ([string]$Text).Trim()
    if (-not $t) { return $null }
    $isLabel = $t.StartsWith('@')
    return [pscustomobject]@{ Label = (ConvertTo-Label $t.TrimStart('@')); IsLabel = $isLabel; Text = $t }
}

$ruleList = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Rules @('種類', 'ファイル名', '探し方') '網羅のルール')) {
    $kind = ([string]$r.'種類').Trim()
    if (-not $kind) { continue }
    $get = { param($name) if ($r.PSObject.Properties[$name]) { ([string]$r.$name).Trim() } else { '' } }
    $how = & $get '探し方'
    $pattern = $null
    if ($how -ne 'ID' -and $how -ne '名前') {
        try { $pattern = New-Object System.Text.RegularExpressions.Regex($how, 'IgnoreCase') }
        catch { throw ("網羅のルールの「探し方」の正規表現が正しくありません（{0}）: {1}" -f $kind, $_.Exception.Message) }
    }
    $fileRx = $null; $sheetRx = $null
    if (& $get 'ファイル名') { $fileRx = New-Object System.Text.RegularExpressions.Regex((& $get 'ファイル名'), 'IgnoreCase') }
    if (& $get 'シート名') { $sheetRx = New-Object System.Text.RegularExpressions.Regex((& $get 'シート名'), 'IgnoreCase') }
    $ruleList.Add([pscustomobject]@{
        Kind = $kind; FileRx = $fileRx; SheetRx = $sheetRx
        Headers = @((& $get '見出し') -split '/' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ })
        Column = (ConvertTo-Field (& $get '列')); Name = (ConvertTo-Field (& $get '名前')); Scope = (ConvertTo-Field (& $get '範囲'))
        How = $how; Pattern = $pattern
    })
}

# ---------------------------------------------------------------- ヘルパー


function ConvertTo-Key {
    param([string] $Text)
    return (ConvertTo-Label $Text).ToUpperInvariant()
}

function New-TokenRegex {
    #     ID がほかの英数字に続かない形で出てくるか（SCR001 が SCR0011 に当たらないように）
    param([string] $Id)
    return New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9]){0}(?![A-Za-z0-9])' -f [regex]::Escape($Id.Normalize($nfkc).Trim())), 'IgnoreCase')
}

function Get-FilePaths {
    #     ファイル・フォルダの中のファイルのフルパス（大文字小文字を区別しない集合）
    param([string[]] $Paths)
    $set = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in $Paths) {
        foreach ($item in @(Get-PathItem $p)) {
            if ($item.PSIsContainer) { foreach ($f in @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)) { [void]$set.Add($f.FullName) } }
            else { [void]$set.Add($item.FullName) }
        }
    }
    return , $set
}

function Read-Books {
    #     Search-Excel で値を読み、ファイル → シート → 行 → 列 の形にする（Skip のファイルと -OutXlsx は除く）
    param([string[]] $Paths, $Skip)
    $searchArgs = @{ Pattern = '\S'; Path = $Paths; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) -and -not ($Skip -and $Skip.Contains($_.Path)) })
    return ConvertTo-DesignBooks $hits
}

function New-Place {
    param($Book, [string] $Sheet, [int] $Row, [int] $Col)
    return [pscustomobject]@{ File = $Book.File; Path = $Book.Path; Sheet = $Sheet; Row = $Row; Col = $Col; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $Col), $Row) }
}

function Format-Place {
    param($Place)
    if (-not $Place) { return '' }
    if (-not $Place.Sheet) { return $Place.File }
    return '{0} {1}!{2}' -f $Place.File, $Place.Sheet, $Place.Cell
}

# ---------------------------------------------------------------- 要件一覧を読む


$requirements = [ordered]@{}   # 要件 ID のキー → @{ Id; Name; Place }
$requirementFiles = $null
if ($Requirement) {
    $requirementFiles = Get-FilePaths $Requirement
    $idLabels = @($requirementIdHeaders | ForEach-Object { ConvertTo-Label $_ })
    $nameLabels = @($requirementNameHeaders | ForEach-Object { ConvertTo-Label $_ })
    $excelPaths = @($requirementFiles | Where-Object { $_ -match '\.xls[xmb]?$' -and -not ([System.IO.Path]::GetFileName($_)).StartsWith('~$') } | ForEach-Object { [System.Management.Automation.WildcardPattern]::Escape($_) })
    if ($excelPaths.Count -gt 0) {
        foreach ($book in (Read-Books $excelPaths $null).Values) {
            foreach ($sheetName in $book.Sheets.Keys) {
                $grid = $book.Sheets[$sheetName]
                $idLabel = $null; $nameLabel = $null
                foreach ($row in $grid.Rows) {
                    $labels = @($grid.Labels[$row].Values)
                    $idLabel = @($idLabels | Where-Object { $labels -contains $_ })[0]
                    $nameLabel = @($nameLabels | Where-Object { $labels -contains $_ })[0]
                    if ($idLabel -and $nameLabel) { break }
                }
                if (-not ($idLabel -and $nameLabel)) { continue }
                foreach ($t in (Find-DesignTables $grid @($idLabel, $nameLabel) $idLabel)) {
                    foreach ($dr in $t.Rows) {
                        $id = ([string]$dr.Cells[$t.ColumnOf[$idLabel]]).Trim()
                        $k = ConvertTo-Key $id
                        if ($k -and -not $requirements.Contains($k)) { $requirements[$k] = [pscustomobject]@{ Id = $id; Name = ([string]$dr.Cells[$t.ColumnOf[$nameLabel]]).Trim(); Place = (New-Place $book $sheetName $dr.Row $t.ColumnOf[$idLabel]) } }
                    }
                }
            }
        }
    }
    foreach ($csv in @($requirementFiles | Where-Object { $_ -match '\.csv$' })) {
        $rows = Read-CsvFile $csv @() '要件一覧'
        if ($rows.Count -eq 0) { continue }
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $idCol = @($names | Where-Object { $idLabels -contains (ConvertTo-Label $_) })[0]
        $nameCol = @($names | Where-Object { $nameLabels -contains (ConvertTo-Label $_) })[0]
        if (-not $idCol) { Write-Warning ("{0}: 要件 ID の列が見つかりません" -f [System.IO.Path]::GetFileName($csv)); continue }
        $line = 1
        foreach ($r in $rows) {
            $line++
            $id = ([string]$r.$idCol).Trim()
            $k = ConvertTo-Key $id
            $nameText = ''
            if ($nameCol) { $nameText = ([string]$r.$nameCol).Trim() }
            if ($k -and -not $requirements.Contains($k)) { $requirements[$k] = [pscustomobject]@{ Id = $id; Name = $nameText; Place = [pscustomobject]@{ File = [System.IO.Path]::GetFileName($csv); Path = $csv; Sheet = ''; Row = $line; Col = 0; Cell = '' } } }
        }
    }
    if ($requirements.Count -eq 0) { Write-Warning '要件一覧から要件を読めませんでした（要件ID と要件名の見出しがある表を探します）' }
}

# ---------------------------------------------------------------- 設計書から確かめる要素を読む


$designBooks = Read-Books $DesignPath $requirementFiles
$elements = [ordered]@{}   # 種類|範囲|ID → 要素

function Add-Element {
    param($Rule, [string] $Id, [string] $Name, [string] $Scope, $Place)
    $id = ([string]$Id).Trim()
    if (-not $id) { return }
    $scopeText = ([string]$Scope).Trim()
    $key = '{0}|{1}|{2}' -f $Rule.Kind, (ConvertTo-Key $scopeText), (ConvertTo-Key $id)
    if ($elements.Contains($key)) { return }
    $how = $Rule.How
    if ($how -ne '名前') { $how = 'ID' }
    $elements[$key] = [pscustomobject]@{ Kind = $Rule.Kind; Id = $id; Name = ([string]$Name).Trim(); Scope = $scopeText; How = $how; Place = $Place; Tests = 0; TestPlace = $null }
}

$bookPaths = [string[]]@($designBooks.Keys)
[Array]::Sort($bookPaths, [StringComparer]::Ordinal)
foreach ($bookPath in $bookPaths) {
    $book = $designBooks[$bookPath]
    $fileName = [System.IO.Path]::GetFileName($book.Path)
    foreach ($rule in $ruleList) {
        if ($rule.FileRx -and -not $rule.FileRx.IsMatch($fileName)) { continue }
        foreach ($sheetName in @($book.Sheets.Keys | Where-Object { -not $rule.SheetRx -or $rule.SheetRx.IsMatch($_) })) {
            $grid = $book.Sheets[$sheetName]
            $scope = ''
            if ($rule.Scope) { $scope = Find-LabelValue $grid $rule.Scope.Label; if (-not $scope) { continue } }
            if ($rule.Pattern) {
                # 正規表現: シートのすべてのセルから ID を集める
                foreach ($row in $grid.Rows) {
                    foreach ($col in $grid.Cols[$row]) {
                        foreach ($m in $rule.Pattern.Matches($grid.Cells[$row][$col].Normalize($nfkc))) {
                            $id = $m.Value
                            if ($m.Groups['id'].Success) { $id = $m.Groups['id'].Value }
                            Add-Element $rule $id '' $scope (New-Place $book $sheetName $row $col)
                        }
                    }
                }
            }
            elseif ($rule.Headers.Count -gt 0) {
                foreach ($t in (Find-DesignTables $grid $rule.Headers $rule.Column.Label)) {
                    if (-not $t.ColumnOf.ContainsKey($rule.Column.Label)) { continue }
                    foreach ($dr in $t.Rows) {
                        $nameText = ''
                        if ($rule.Name -and $t.ColumnOf.ContainsKey($rule.Name.Label)) { $nameText = [string]$dr.Cells[$t.ColumnOf[$rule.Name.Label]] }
                        Add-Element $rule ([string]$dr.Cells[$t.ColumnOf[$rule.Column.Label]]) $nameText $scope (New-Place $book $sheetName $dr.Row $t.ColumnOf[$rule.Column.Label])
                    }
                }
            }
            elseif ($rule.Column -and $rule.Column.IsLabel) {
                # 1 ファイル 1 つ（IF 定義書の「IF ID」など）
                $c = Find-LabelCell $grid $rule.Column.Label
                if (-not ($c -and $c.Value)) { continue }
                $nameText = ''
                if ($rule.Name -and $rule.Name.IsLabel) { $nameText = Find-LabelValue $grid $rule.Name.Label }
                Add-Element $rule $c.Value $nameText $scope (New-Place $book $sheetName $c.Row $c.Col)
                break
            }
        }
    }
}
if ($elements.Count -eq 0 -and $requirements.Count -eq 0) { Write-Warning '設計書から確かめる要素が見つかりません（網羅のルールのファイル名・見出しを確かめてください）' }

# ---------------------------------------------------------------- テスト仕様書を読む


# テストのセル: 文字と、属する範囲（ファイル名・シート名・行の機能・行のすべての文字）
$testBooks = Read-Books $TestPath $requirementFiles
$testCells = New-Object System.Collections.Generic.List[object]
$functionLabels = @($functionHeaders | ForEach-Object { ConvertTo-Label $_ })
foreach ($book in $testBooks.Values) {
    foreach ($sheetName in $book.Sheets.Keys) {
        $grid = $book.Sheets[$sheetName]
        $headerRow = 0; $functionCol = 0
        foreach ($row in $grid.Rows) {
            foreach ($col in $grid.Cols[$row]) { if ($functionLabels -contains $grid.Labels[$row][$col]) { $headerRow = $row; $functionCol = $col; break } }
            if ($headerRow) { break }
        }
        # シートの上の「画面ID: SCR001」のような見出しの右の値は、シートのすべての行の範囲にする
        $sheetScope = (@($functionLabels | ForEach-Object { Find-LabelValue $grid $_ } | Where-Object { $_ }) -join ' ')
        $current = ''
        foreach ($row in $grid.Rows) {
            $cells = $grid.Cells[$row]
            if ($headerRow -and $row -gt $headerRow -and $cells.ContainsKey($functionCol) -and $cells[$functionCol].Trim()) { $current = $cells[$functionCol].Trim() }
            $rowText = (@($grid.Cols[$row] | ForEach-Object { $cells[$_] }) -join ' ').Normalize($nfkc)
            $scopeText = ('{0} {1} {2} {3} {4}' -f $book.File, $sheetName, $sheetScope, $current, $rowText).Normalize($nfkc)
            foreach ($col in $grid.Cols[$row]) {
                $text = $cells[$col].Normalize($nfkc)
                $testCells.Add([pscustomobject]@{ Text = $text; Label = (ConvertTo-Label $text).ToUpperInvariant(); Scope = $scopeText; Place = (New-Place $book $sheetName $row $col) })
            }
        }
    }
}
if ($testCells.Count -eq 0) { Write-Warning 'テスト仕様書が見つかりません' }

# ---------------------------------------------------------------- 突き合わせる


# 範囲ごとのテストのセル（同じ画面の要素が多いので、範囲で先に絞る）
$cellsByScope = @{}
function Get-ScopeCells {
    param([string] $Scope)
    if (-not $Scope) { return , $testCells }
    $k = ConvertTo-Key $Scope
    if (-not $cellsByScope.ContainsKey($k)) {
        $rx = New-TokenRegex $Scope
        $cellsByScope[$k] = @($testCells | Where-Object { $rx.IsMatch($_.Scope) })
    }
    return , $cellsByScope[$k]
}

foreach ($e in $elements.Values) {
    $cells = Get-ScopeCells $e.Scope
    if ($e.How -eq '名前' -and $e.Name) {
        $label = (ConvertTo-Label $e.Name).ToUpperInvariant()
        $found = @($cells | Where-Object { $_.Label.Contains($label) })
    }
    else {
        $rx = New-TokenRegex $e.Id
        $found = @($cells | Where-Object { $rx.IsMatch($_.Text) })
    }
    $e.Tests = $found.Count
    if ($found.Count -gt 0) { $e.TestPlace = $found[0].Place }
}

# 要件: 設計書とテスト仕様書のどこに出てくるか
$designCells = New-Object System.Collections.Generic.List[object]
foreach ($book in $designBooks.Values) {
    foreach ($sheetName in $book.Sheets.Keys) {
        $grid = $book.Sheets[$sheetName]
        foreach ($row in $grid.Rows) { foreach ($col in $grid.Cols[$row]) { $designCells.Add([pscustomobject]@{ Text = $grid.Cells[$row][$col].Normalize($nfkc); Place = (New-Place $book $sheetName $row $col) }) } }
    }
}
$requirementResults = New-Object System.Collections.Generic.List[object]
foreach ($k in $requirements.Keys) {
    $req = $requirements[$k]
    $rx = New-TokenRegex $req.Id
    $inDesign = @($designCells | Where-Object { $rx.IsMatch($_.Text) })
    $inTest = @($testCells | Where-Object { $rx.IsMatch($_.Text) })
    $judgement = '○'
    if ($inDesign.Count -eq 0 -and $inTest.Count -eq 0) { $judgement = '設計・テストなし' }
    elseif ($inDesign.Count -eq 0) { $judgement = '設計なし' }
    elseif ($inTest.Count -eq 0) { $judgement = 'テストなし' }
    $requirementResults.Add([pscustomobject]@{
        Id = $req.Id; Name = $req.Name; Place = $req.Place; Judgement = $judgement
        DesignFiles = (@($inDesign | ForEach-Object { $_.Place.File } | Select-Object -Unique) -join '、'); DesignPlace = @($inDesign | ForEach-Object { $_.Place })[0]
        TestFiles = (@($inTest | ForEach-Object { $_.Place.File } | Select-Object -Unique) -join '、'); Tests = $inTest.Count
    })
}

# 要件一覧に無い要件 ID: 一覧の ID の形（REQ-001 → REQ-\d{3}）で、設計書とテスト仕様書を探す
$unknownRequirements = New-Object System.Collections.Generic.List[object]
if ($requirements.Count -gt 0) {
    $shapes = @($requirements.Values | ForEach-Object {
            $id = $_.Id.Normalize($nfkc)
            if ($id -match '[A-Za-z]' -and $id -match '\d') { [regex]::Replace([regex]::Escape($id), '\d+', [System.Text.RegularExpressions.MatchEvaluator] { param($m) '\d{' + $m.Value.Length + '}' }) }
        } | Sort-Object -Unique)
    if ($shapes.Count -gt 0) {
        $shapeRx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9])(?:{0})(?![A-Za-z0-9])' -f ($shapes -join '|')), 'IgnoreCase')
        $seen = @{}
        foreach ($c in ($designCells.ToArray() + $testCells.ToArray())) {
            foreach ($m in $shapeRx.Matches($c.Text)) {
                $k = ConvertTo-Key $m.Value
                if ($requirements.Contains($k) -or $seen.ContainsKey($k)) { continue }
                $seen[$k] = $true
                $unknownRequirements.Add([pscustomobject]@{ Id = $m.Value; Place = $c.Place })
            }
        }
    }
}

# ---------------------------------------------------------------- 結果


$gaps = New-Object System.Collections.Generic.List[object]
function Add-Gap {
    param([string] $Kind, [string] $Scope, [string] $Id, [string] $Name, [string] $Judgement, $Place)
    $g = [pscustomobject]@{ Kind = $Kind; Scope = $Scope; Id = $Id; Name = $Name; Judgement = $Judgement; Place = (Format-Place $Place); Path = $Place.Path; Sheet = $Place.Sheet; Cell = $Place.Cell }
    $g.PSObject.TypeNames.Insert(0, 'Traceability.Gap')
    $gaps.Add($g)
}
foreach ($r in $requirementResults) { if ($r.Judgement -ne '○') { Add-Gap '要件' '' $r.Id $r.Name $r.Judgement $r.Place } }
foreach ($u in $unknownRequirements) { Add-Gap '要件' '' $u.Id '' '要件一覧に無い要件 ID' $u.Place }
foreach ($e in $elements.Values) {
    if ($e.Tests -gt 0) { continue }
    $judgement = 'テストなし'
    if ($e.How -eq '名前') { $judgement = 'テストなし（名前が出てこない）' }
    Add-Gap $e.Kind $e.Scope $e.Id $e.Name $judgement $e.Place
}
$gaps

if ($OutXlsx) {
    $coverRows = New-Object System.Collections.Generic.List[object]
    foreach ($e in $elements.Values) {
        $judgement = '○'
        if ($e.Tests -eq 0) { $judgement = 'テストなし' }
        $row = @{ Values = @($e.Kind, $e.Scope, $e.Id, $e.Name, $(if ($e.How -eq '名前') { '名前' } else { 'ID' }), [double]$e.Tests, (Format-Place $e.TestPlace), $judgement, (Format-Place $e.Place)) }
        if ($e.TestPlace -and $e.TestPlace.Sheet) { $row.Link = @{ Column = 6; Path = $e.TestPlace.Path; Location = ("'{0}'!{1}" -f $e.TestPlace.Sheet.Replace("'", "''"), $e.TestPlace.Cell) } }
        elseif ($e.Place.Sheet) { $row.Link = @{ Column = 8; Path = $e.Place.Path; Location = ("'{0}'!{1}" -f $e.Place.Sheet.Replace("'", "''"), $e.Place.Cell) } }
        $coverRows.Add($row)
    }
    $reqRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $requirementResults) {
        $row = @{ Values = @($r.Id, $r.Name, $r.DesignFiles, $r.TestFiles, [double]$r.Tests, $r.Judgement, (Format-Place $r.Place)) }
        if ($r.Place.Sheet) { $row.Link = @{ Column = 6; Path = $r.Place.Path; Location = ("'{0}'!{1}" -f $r.Place.Sheet.Replace("'", "''"), $r.Place.Cell) } }
        $reqRows.Add($row)
    }
    $gapRows = New-Object System.Collections.Generic.List[object]
    foreach ($g in $gaps) {
        $row = @{ Values = @($g.Kind, $g.Scope, $g.Id, $g.Name, $g.Judgement, $g.Place) }
        if ($g.Sheet) { $row.Link = @{ Column = 5; Path = $g.Path; Location = ("'{0}'!{1}" -f $g.Sheet.Replace("'", "''"), $g.Cell) } }
        $gapRows.Add($row)
    }
    $sheets = @(
        @{ Name = '抜け'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '範囲'; Width = 10 }, @{ Header = 'ID'; Width = 14 }, @{ Header = '名前'; Width = 24 }, @{ Header = '判定'; Width = 26 }, @{ Header = '場所'; Width = 50 }); Rows = $gapRows },
        @{ Name = '網羅'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = '範囲'; Width = 10 }, @{ Header = 'ID'; Width = 14 }, @{ Header = '名前'; Width = 24 }, @{ Header = '探し方'; Width = 7 }, @{ Header = 'テストの数'; Width = 9 }, @{ Header = '最初のテストの場所'; Width = 44 }, @{ Header = '判定'; Width = 10 }, @{ Header = '設計書の場所'; Width = 50 }); Rows = $coverRows })
    if ($requirements.Count -gt 0) {
        $sheets += @{ Name = '要件'; Columns = @(@{ Header = '要件ID'; Width = 12 }, @{ Header = '要件名'; Width = 30 }, @{ Header = '設計書'; Width = 40 }, @{ Header = 'テスト仕様書'; Width = 40 }, @{ Header = 'テストの数'; Width = 9 }, @{ Header = '判定'; Width = 14 }, @{ Header = '要件一覧の場所'; Width = 40 }); Rows = $reqRows }
    }
    try { Export-Workbook -FilePath $OutXlsx -Sheets $sheets }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
# 網羅のルールの順に、種類ごとのテストのある要素の割合
$coverage = @(@($ruleList | ForEach-Object { $_.Kind } | Select-Object -Unique) | ForEach-Object {
        $kind = $_
        $group = @($elements.Values | Where-Object { $_.Kind -eq $kind })
        if ($group.Count -gt 0) {
            $covered = @($group | Where-Object { $_.Tests -gt 0 }).Count
            '{0} {1}/{2}（{3:0}%）' -f $kind, $covered, $group.Count, ($covered * 100.0 / $group.Count)
        }
    })
if ($coverage.Count -gt 0) { Write-Host ('テストの網羅: {0}' -f ($coverage -join '・')) }
if ($requirements.Count -gt 0) {
    $ok = @($requirementResults | Where-Object Judgement -eq '○').Count
    Write-Host ('要件 {0} 件: 設計とテストあり {1}・設計なし {2}・テストなし {3}・設計・テストなし {4}。要件一覧に無い要件 ID {5} 件' -f $requirements.Count, $ok, @($requirementResults | Where-Object Judgement -eq '設計なし').Count, @($requirementResults | Where-Object Judgement -eq 'テストなし').Count, @($requirementResults | Where-Object Judgement -eq '設計・テストなし').Count, $unknownRequirements.Count)
}
Write-Host ('抜け: {0} 件' -f $gaps.Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     要件一覧 → 設計書 → テスト仕様書のつながりと、設計の要素（画面・イベント・メッセージ・IF）にテストがあるかを確かめます（Excel 不要）。
#
# .DESCRIPTION
#     設計書から、網羅のルール（tools\traceability\網羅のルール.csv）に沿って確かめる要素を集め、テスト仕様書に出てくるかを調べます。
#
#       画面       画面一覧の画面 ID が、テスト仕様書に出てくるか
#       イベント   画面設計書のイベントの名前（検索・クリアなど）が、その画面のテストに出てくるか
#       メッセージ 画面設計書で使うメッセージ ID（E0003 など）が、その画面のテストに出てくるか
#       IF         IF 定義書の IF ID が、テスト仕様書に出てくるか
#
#     「その画面のテスト」は、ファイル名・シート名・行（機能・画面の列。空なら上の行の値）に画面 ID があるテスト仕様書のセルです。
#     -Requirement で要件一覧を渡すと、要件 ID ごとに設計書とテスト仕様書に出てくるかを確かめ、一覧に無い要件 ID も知らせます。
#
# .PARAMETER DesignPath
#     設計書のフォルダまたはファイル（複数可・ワイルドカード可）。
#
# .PARAMETER TestPath
#     テスト仕様書のフォルダまたはファイル（複数可・ワイルドカード可）。
#
# .PARAMETER Requirement
#     要件一覧のファイル（Excel・CSV。要件ID と要件名の見出しがある表）。設計書・テスト仕様書のフォルダの中にあっても、そちらとしては読みません。
#
# .PARAMETER Rules
#     網羅のルールの CSV（種類・ファイル名・シート名・見出し・列・名前・範囲・探し方）。既定は tools\traceability\網羅のルール.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（抜け・網羅・要件）。
#
# .EXAMPLE
#     .\Test-Traceability.ps1 D:\設計書 D:\テスト仕様書 -OutXlsx D:\網羅.xlsx
#
# .EXAMPLE
#     # 要件一覧からのつながりも確かめる
#     .\Test-Traceability.ps1 D:\設計書 D:\テスト仕様書 -Requirement D:\要件一覧.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     テストに「出てくる」かは文字で確かめます。テストの内容が正しいか、十分かまでは確かめません。
