[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $Rules,

    [string] $Exclude,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Rules) { $Rules = Join-Path $PSScriptRoot 'consistency/参照のルール.csv' }
foreach ($name in @('Rules', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

function New-RuleRegex {
    param([string] $Pattern, [string] $What)
    try { return [regex]::new($Pattern, [System.Text.RegularExpressions.RegexOptions]::CultureInvariant) }
    catch { throw ("{0}の正規表現が不正です: {1} 詳細: {2}" -f $What, $Pattern, $_.Exception.InnerException.Message) }
}

$excludeRx = $null
if ($Exclude) { $excludeRx = New-RuleRegex $Exclude '-Exclude ' }

Update-TypeData -TypeName 'DesignConsistency.Issue' -DefaultDisplayPropertySet File, Location, Level, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

function ConvertTo-Key {
    #     ID を比べるための形（全角・半角と大文字・小文字をそろえる）
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    return $Text.Normalize($nfkc).Trim().ToUpperInvariant()
}

function ConvertTo-Field {
    #     ルールの「列」「名前」「親」「元」。「@画面ID」はシートの見出しの右の値、それ以外は表の列の見出し
    param([string] $Text)
    $t = ([string]$Text).Trim()
    if (-not $t) { return $null }
    $isLabel = $t.StartsWith('@')
    $t = $t.TrimStart('@').Trim()
    return [pscustomobject]@{ IsLabel = $isLabel; Label = (ConvertTo-Label $t); Text = $t }
}

# ---------------------------------------------------------------- 参照のルール


$ruleList = New-Object System.Collections.Generic.List[object]
$line = 1
foreach ($r in (Read-CsvFile $Rules @('区分', '種類', 'ファイル名', 'シート名', '見出し', '列', '名前', '親', '元', '書き方', '重複') '参照のルール')) {
    $line++
    $section = ([string]$r.'区分').Trim()
    if (-not $section) { continue }
    $where = '参照のルールの {0} 行目' -f $line
    if (@('定義', '一覧', '参照') -notcontains $section) { throw ("{0}: 「区分」は 定義 / 一覧 / 参照 のどれかにしてください（{1}）" -f $where, $section) }
    $kind = ([string]$r.'種類').Trim()
    if (-not $kind) { throw "${where}: 「種類」が空です" }
    $rule = [pscustomobject]@{
        Line = $line; Section = $section; Kind = $kind; FileRx = $null; SheetRx = $null
        Headers = @(([string]$r.'見出し') -split '/' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ })
        Column = (ConvertTo-Field $r.'列'); Name = (ConvertTo-Field $r.'名前'); Parent = (ConvertTo-Field $r.'親'); From = (ConvertTo-Field $r.'元')
        Pattern = $null; AllowDuplicate = (([string]$r.'重複').Trim() -eq '可')
        Found = $false; Global = $false; Place = $null; Coverage = @{}; Pairs = @{}
    }
    if (([string]$r.'ファイル名').Trim()) { $rule.FileRx = New-RuleRegex ([string]$r.'ファイル名').Trim() "${where}の「ファイル名」" }
    if (([string]$r.'シート名').Trim()) { $rule.SheetRx = New-RuleRegex ([string]$r.'シート名').Trim() "${where}の「シート名」" }
    if ($section -eq '参照') {
        $p = ([string]$r.'書き方').Trim()
        if (-not $p) { throw "${where}: 参照の「書き方」（正規表現）が空です" }
        $rule.Pattern = New-RuleRegex $p "${where}の「書き方」"
        if (@($rule.Pattern.GetGroupNames()) -notcontains 'id') { throw "${where}: 「書き方」に ID を表す (?<id>...) がありません" }
        if ($rule.Column -and ($rule.Column.IsLabel -or $rule.Headers.Count -eq 0)) { throw "${where}: 参照の「列」は表の列の見出しで書き、「見出し」も指定してください" }
        if ($rule.From -and -not $rule.From.IsLabel -and $rule.Headers.Count -eq 0) { throw "${where}: 「元」に表の列を指定するときは「見出し」も指定してください" }
    }
    else {
        if (-not $rule.Column) { throw "${where}: 「列」（ID の列）が空です" }
        if (-not $rule.Column.IsLabel -and $rule.Headers.Count -eq 0) { throw "${where}: 「列」に表の列を指定するときは「見出し」も指定してください（シートの見出しの右の値は「@画面ID」のように書きます）" }
        foreach ($f in @($rule.Name, $rule.Parent)) {
            if ($f -and -not $f.IsLabel -and $rule.Column.IsLabel) { throw "${where}: 「列」が「@」で始まるときは、「名前」「親」も「@」で始まる見出しで書いてください" }
        }
    }
    $ruleList.Add($rule)
}
if ($ruleList.Count -eq 0) { throw "参照のルールが空です: $Rules" }

# ---------------------------------------------------------------- 設計書を読む


$searchArgs = @{ Pattern = '\S'; Path = $Path; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$books = ConvertTo-DesignBooks $hits

$defs = New-Object System.Collections.Generic.List[object]   # 定義と一覧
$refs = New-Object System.Collections.Generic.List[object]   # 参照
$refSeen = New-Object 'System.Collections.Generic.HashSet[string]'

function New-Place {
    param($Book, [string] $Sheet, [int] $Row, [int] $Col)
    return [pscustomobject]@{
        File = $Book.File; Path = $Book.Path; Sheet = $Sheet; SheetIndex = @($Book.Sheets.Keys).IndexOf($Sheet); Row = $Row; Col = $Col
        Cell = ('{0}{1}' -f (ConvertTo-ColumnName $Col), $Row)
    }
}

function Add-Definition {
    param($Rule, [string] $Id, [string] $Name, [string] $Parent, $Place, [string] $Unit, $Values)
    $idKey = ConvertTo-Key $Id
    if (-not $idKey) { return }
    $parentKey = ConvertTo-Key $Parent
    $key = $idKey
    $shown = $Id.Trim()
    if ($Rule.Parent) { $key = '{0}.{1}' -f $parentKey, $idKey; $shown = '{0}.{1}' -f ([string]$Parent).Trim(), $Id.Trim() }
    $defs.Add([pscustomobject]@{
        Rule = $Rule; Section = $Rule.Section; Kind = $Rule.Kind; Id = $Id.Trim(); Shown = $shown; Name = ([string]$Name).Trim()
        ParentKey = $parentKey; Key = $key; Place = $Place; Unit = $Unit; Values = $Values
    })
}

function Select-Majority {
    #     シートごとの値のうち、いちばん多いもの（同数なら先に出たもの）
    param($Items, [scriptblock] $KeyOf)
    $counts = @{}
    foreach ($i in $Items) { $k = & $KeyOf $i.Value; if ($counts.ContainsKey($k)) { $counts[$k]++ } else { $counts[$k] = 1 } }
    $best = $null; $bestCount = 0
    foreach ($i in $Items) { $k = & $KeyOf $i.Value; if ($counts[$k] -gt $bestCount) { $best = $i; $bestCount = $counts[$k] } }
    return $best
}

function Read-FileDefinition {
    #     表を持たない定義（1 ファイル = 1 つ）。各シートのヘッダの「画面ID」などの値を集め、いちばん多い値をそのファイルの ID とする
    param($Rule, $Book, [string[]] $SheetNames)
    $ids = New-Object System.Collections.Generic.List[object]
    $names = New-Object System.Collections.Generic.List[object]
    $parent = $null
    foreach ($s in $SheetNames) {
        $grid = $Book.Sheets[$s]
        $c = Find-LabelCell $grid $Rule.Column.Label
        if ($c -and $c.Value) { $ids.Add([pscustomobject]@{ Value = $c.Value; Place = (New-Place $Book $s $c.Row $c.Col) }) }
        if ($Rule.Name) {
            $n = Find-LabelCell $grid $Rule.Name.Label
            if ($n -and $n.Value) { $names.Add([pscustomobject]@{ Value = $n.Value; Place = (New-Place $Book $s $n.Row $n.Col) }) }
        }
        if ($Rule.Parent -and -not $parent) { $parent = Find-LabelValue $grid $Rule.Parent.Label }
    }
    if ($ids.Count -eq 0) { return $false }
    $id = Select-Majority $ids { param($v) ConvertTo-Key $v }
    $name = $null
    if ($names.Count -gt 0) { $name = (Select-Majority $names { param($v) ConvertTo-Label $v }).Value }
    Add-Definition $Rule $id.Value $name $parent $id.Place 'ファイル' ([pscustomobject]@{ Ids = $ids; Names = $names })
    return $true
}

function Read-TableDefinitions {
    #     表から読む定義。「列」が「@」で始まるときは表のあるシートごとに 1 つ、それ以外は表の行ごとに 1 つ
    param($Rule, $Book, [string[]] $SheetNames)
    $found = $false
    $keyHeader = $null
    if (-not $Rule.Column.IsLabel) { $keyHeader = $Rule.Column.Label }
    foreach ($s in $SheetNames) {
        $grid = $Book.Sheets[$s]
        $sheetName = $null; $sheetParent = $null
        if ($Rule.Name -and $Rule.Name.IsLabel) { $sheetName = Find-LabelValue $grid $Rule.Name.Label }
        if ($Rule.Parent -and $Rule.Parent.IsLabel) { $sheetParent = Find-LabelValue $grid $Rule.Parent.Label }
        foreach ($t in (Find-DesignTables $grid $Rule.Headers $keyHeader)) {
            if ($Rule.Column.IsLabel) {
                $c = Find-LabelCell $grid $Rule.Column.Label
                if (-not $c -or -not $c.Value) { continue }
                $found = $true
                Add-Definition $Rule $c.Value $sheetName $sheetParent (New-Place $Book $s $c.Row $c.Col) 'シート' $null
                break
            }
            if (-not $t.ColumnOf.ContainsKey($Rule.Column.Label)) { continue }
            $found = $true
            $idCol = $t.ColumnOf[$Rule.Column.Label]
            $nameCol = $null; $parentCol = $null
            if ($Rule.Name -and -not $Rule.Name.IsLabel -and $t.ColumnOf.ContainsKey($Rule.Name.Label)) { $nameCol = $t.ColumnOf[$Rule.Name.Label] }
            if ($Rule.Parent -and -not $Rule.Parent.IsLabel -and $t.ColumnOf.ContainsKey($Rule.Parent.Label)) { $parentCol = $t.ColumnOf[$Rule.Parent.Label] }
            $lastParent = $null
            foreach ($dr in $t.Rows) {
                $id = [string]$dr.Cells[$idCol]
                if (-not $id.Trim()) { continue }
                $name = $sheetName
                if ($nameCol) { $name = [string]$dr.Cells[$nameCol] }
                $parent = $sheetParent
                if ($parentCol) {
                    # 区分の ID を最初の行にだけ書く表（2 行目以降が空）は、上の行の値を使う
                    $parent = [string]$dr.Cells[$parentCol]
                    if ($parent.Trim()) { $lastParent = $parent } else { $parent = $lastParent }
                }
                Add-Definition $Rule $id $name $parent (New-Place $Book $s $dr.Row $idCol) '行' $null
            }
        }
    }
    return $found
}

function Add-Reference {
    param($Rule, [string] $Text, [string] $From, $Book, [string] $Sheet, [int] $Row, [int] $Col)
    if (-not $Text) { return }
    $matches2 = $Rule.Pattern.Matches($Text.Normalize($nfkc))
    if ($matches2.Count -eq 0) { return }
    $fromKey = ConvertTo-Key $From
    $Place = New-Place $Book $Sheet $Row $Col
    foreach ($m in $matches2) {
        $id = $m.Groups['id'].Value
        if (-not $id) { continue }
        $parent = ''
        if ($m.Groups['parent'].Success) { $parent = $m.Groups['parent'].Value }
        $shown = $id
        $key = ConvertTo-Key $id
        if ($parent) { $shown = '{0}.{1}' -f $parent, $id; $key = '{0}.{1}' -f (ConvertTo-Key $parent), $key }
        if ($excludeRx -and $excludeRx.IsMatch($shown)) { continue }
        if ($Rule.From -and $fromKey) {
            $pairKey = '{0}>{1}' -f $fromKey, $key
            if (-not $Rule.Pairs.ContainsKey($pairKey)) { $Rule.Pairs[$pairKey] = [pscustomobject]@{ FromKey = $fromKey; From = $From.Trim(); To = $shown; Place = $Place } }
        }
        if (-not $refSeen.Add(('{0}|{1}|{2}|{3}|{4}' -f $Rule.Kind, $key, $Place.Path, $Place.Sheet, $Place.Cell))) { continue }
        $refs.Add([pscustomobject]@{ Rule = $Rule; Kind = $Rule.Kind; Shown = $shown; Key = $key; Parent = $parent; ParentKey = (ConvertTo-Key $parent); Raw = $Text; Place = $Place })
    }
}

function Add-Coverage {
    #     「元」が「@画面ID」の参照: この画面の参照を読んだことを記録する（組み合わせを比べる範囲）
    param($Rule, [string] $From, $Book, [string] $Sheet)
    $k = ConvertTo-Key $From
    if ($k -and -not $Rule.Coverage.ContainsKey($k)) { $Rule.Coverage[$k] = '{0}（{1}）' -f $Book.File, $Sheet }
}

function Read-References {
    param($Rule, $Book, [string[]] $SheetNames)
    foreach ($s in $SheetNames) {
        $grid = $Book.Sheets[$s]
        $fromLabel = $null
        if ($Rule.From -and $Rule.From.IsLabel) { $fromLabel = Find-LabelValue $grid $Rule.From.Label }
        if ($Rule.Headers.Count -eq 0) {
            # 見出しの指定がなければ、シートのすべてのセルから探す
            $Rule.Found = $true
            if ($Rule.From) { Add-Coverage $Rule $fromLabel $Book $s }
            foreach ($row in $grid.Rows) {
                $cells = $grid.Cells[$row]
                foreach ($col in $grid.Cols[$row]) {
                    # 一致しないセルが大半なので、先に確かめてから記録する
                    if ($Rule.Pattern.IsMatch($cells[$col].Normalize($nfkc))) { Add-Reference $Rule $cells[$col] $fromLabel $Book $s $row $col }
                }
            }
            continue
        }
        $keyHeader = $null
        if ($Rule.Column) { $keyHeader = $Rule.Column.Label }
        foreach ($t in (Find-DesignTables $grid $Rule.Headers $keyHeader)) {
            if ($Rule.Column -and -not $t.ColumnOf.ContainsKey($Rule.Column.Label)) { continue }
            $Rule.Found = $true
            $fromCol = $null
            if ($Rule.From) {
                if ($Rule.From.IsLabel) { Add-Coverage $Rule $fromLabel $Book $s }
                else {
                    # 「元」が表の列（画面遷移の一覧など）: すべての画面の組み合わせが書かれている前提で比べる
                    $Rule.Global = $true
                    if (-not $Rule.Place) { $Rule.Place = '{0}（{1}）' -f $Book.File, $s }
                    if ($t.ColumnOf.ContainsKey($Rule.From.Label)) { $fromCol = $t.ColumnOf[$Rule.From.Label] }
                }
            }
            foreach ($dr in $t.Rows) {
                $from = $fromLabel
                if ($fromCol) { $from = [string]$dr.Cells[$fromCol] }
                if ($Rule.Column) { $cols = @($t.ColumnOf[$Rule.Column.Label]) } else { $cols = @($dr.Cells.Keys | Sort-Object) }
                foreach ($col in $cols) {
                    if ($dr.Cells.ContainsKey($col)) { Add-Reference $Rule ([string]$dr.Cells[$col]) $from $Book $s $dr.Row $col }
                }
            }
        }
    }
}

# 重複の 2 つ目などが OS によって変わらないよう、パスの順に読む
$bookPaths = [string[]]@($books.Keys)
[Array]::Sort($bookPaths, [StringComparer]::Ordinal)
foreach ($bookPath in $bookPaths) {
    $book = $books[$bookPath]
    $fileName = [System.IO.Path]::GetFileName($book.Path)
    foreach ($rule in $ruleList) {
        if ($rule.FileRx -and -not $rule.FileRx.IsMatch($fileName)) { continue }
        $sheetNames = @($book.Sheets.Keys | Where-Object { -not $rule.SheetRx -or $rule.SheetRx.IsMatch($_) })
        if ($sheetNames.Count -eq 0) { continue }
        if ($rule.Section -eq '参照') { Read-References $rule $book $sheetNames; continue }
        if ($rule.Headers.Count -eq 0) { $ok = Read-FileDefinition $rule $book $sheetNames } else { $ok = Read-TableDefinitions $rule $book $sheetNames }
        if ($ok) { $rule.Found = $true }
        elseif ($rule.Section -eq '定義' -and $rule.FileRx) {
            $what = $rule.Column.Text
            if ($rule.Headers.Count -gt 0) { $what = '見出し「{0}」の表' -f ($rule.Headers -join ' / ') }
            Write-Warning ("{0}: {1}の定義（{2}）が見つかりません" -f $book.File, $rule.Kind, $what)
        }
    }
}

# ---------------------------------------------------------------- 突き合わせる


$script:Issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    param([ValidateSet('エラー', '警告', '情報')] [string] $Level, [string] $Item, [string] $Detail, $Place)
    $issue = [pscustomobject]@{ File = ''; Location = ''; Level = $Level; Item = $Item; Detail = $Detail; Sheet = ''; Cell = ''; Path = ''; SheetIndex = 0; Row = 0; Col = 0 }
    if ($Place) {
        $issue.File = $Place.File; $issue.Path = $Place.Path; $issue.Sheet = $Place.Sheet; $issue.Cell = $Place.Cell
        $issue.SheetIndex = $Place.SheetIndex; $issue.Row = $Place.Row; $issue.Col = $Place.Col
        $issue.Location = '{0}!{1}' -f $Place.Sheet, $Place.Cell
    }
    $issue.PSObject.TypeNames.Insert(0, 'DesignConsistency.Issue')
    $script:Issues.Add($issue)
}

function Format-Places {
    #     「シート!セル」の一覧（ほかのファイルはファイル名も付ける）。多いときは先頭の 5 か所
    param($Places, [string] $BasePath)
    $list = @($Places | Select-Object -First 5 | ForEach-Object {
        if ($_.Path -eq $BasePath) { '{0}!{1}' -f $_.Sheet, $_.Cell } else { '{0} {1}!{2}' -f $_.File, $_.Sheet, $_.Cell }
    })
    $text = $list -join '、'
    if (@($Places).Count -gt 5) { $text += ' ほか {0} か所' -f (@($Places).Count - 5) }
    return $text
}

function Find-Original {
    #     全角・半角をそろえる前の、セルに書かれたとおりの文字
    param([string] $Raw, [string] $Shown)
    if ($Raw.IndexOf($Shown, [System.StringComparison]::Ordinal) -ge 0) { return $Shown }
    for ($i = 0; $i -lt $Raw.Length; $i++) {
        for ($len = 1; $i + $len -le $Raw.Length -and $len -le $Shown.Length * 2; $len++) {
            $sub = $Raw.Substring($i, $len)
            if ([string]::Equals($sub.Normalize($nfkc), $Shown, [System.StringComparison]::Ordinal)) { return $sub }
        }
    }
    return $Shown
}

# 種類ごとの定義（定義と一覧。参照先があるかどうかは両方で確かめる）
$index = @{}
foreach ($d in $defs) {
    if (-not $index.ContainsKey($d.Kind)) { $index[$d.Kind] = @{ ById = @{}; Parents = @{} } }
    $e = $index[$d.Kind]
    if (-not $e.ById.ContainsKey($d.Key)) { $e.ById[$d.Key] = $d }
    if ($d.ParentKey) { $e.Parents[$d.ParentKey] = $true }
}
$kinds = @($ruleList | ForEach-Object { $_.Kind } | Select-Object -Unique)

# 重複定義: 同じ ID が 2 か所以上で定義されている（一覧は一覧の中で）
foreach ($section in @('定義', '一覧')) {
    $groups = [ordered]@{}
    foreach ($d in $defs) {
        if ($d.Section -ne $section -or $d.Rule.AllowDuplicate) { continue }
        $gk = '{0}|{1}' -f $d.Kind, $d.Key
        if (-not $groups.Contains($gk)) { $groups[$gk] = New-Object System.Collections.Generic.List[object] }
        $groups[$gk].Add($d)
    }
    foreach ($g in $groups.Values) {
        if ($g.Count -lt 2) { continue }
        # 2 つ目の場所に、ほかの場所を添えて知らせる
        $others = @($g | Where-Object { $_ -ne $g[1] } | ForEach-Object { $_.Place })
        if ($section -eq '定義') { $fmt = '{0}「{1}」は、ほかの場所でも定義されています（{2}）' } else { $fmt = '{0}「{1}」が、一覧のほかの行にも書かれています（{2}）' }
        Add-Issue 'エラー' '重複定義' ($fmt -f $g[0].Kind, $g[0].Shown, (Format-Places $others $g[1].Place.Path)) $g[1].Place
    }
}

# ヘッダの違い: 1 つのファイル（シート）の中で ID・名前が食い違う、ファイル名（シート名）に別の ID が書かれている
foreach ($d in $defs) {
    if ($d.Section -ne '定義') { continue }
    if ($d.Unit -eq 'ファイル') {
        foreach ($v in $d.Values.Ids) {
            if ((ConvertTo-Key $v.Value) -ne (ConvertTo-Key $d.Id)) {
                Add-Issue '警告' 'ヘッダの違い' ('{0}が「{1}」です（このファイルのほかのシートは「{2}」）' -f $d.Rule.Column.Text, $v.Value, $d.Id) $v.Place
            }
        }
        foreach ($v in $d.Values.Names) {
            if ((ConvertTo-Label $v.Value) -ne (ConvertTo-Label $d.Name)) {
                Add-Issue '警告' 'ヘッダの違い' ('{0}が「{1}」です（このファイルのほかのシートは「{2}」）' -f $d.Rule.Name.Text, $v.Value, $d.Name) $v.Place
            }
        }
    }
    if (@('ファイル', 'シート') -notcontains $d.Unit -or $d.Rule.Parent) { continue }
    # ファイル名・シート名に書かれた同じ種類の ID のうち、いちばん長いもの（T_ORDER_DETAIL と T_ORDER なら T_ORDER_DETAIL）
    if ($d.Unit -eq 'ファイル') { $target = [System.IO.Path]::GetFileNameWithoutExtension($d.Place.Path) } else { $target = $d.Place.Sheet }
    $target = ConvertTo-Key $target
    $best = $null
    foreach ($k in $index[$d.Kind].ById.Keys) {
        if ($best -and $k.Length -le $best.Length) { continue }
        if ([regex]::IsMatch($target, '(?<![A-Z0-9])' + [regex]::Escape($k) + '(?![A-Z0-9])')) { $best = $k }
    }
    if ($best -and $best -ne $d.Key) {
        $whereName = 'ファイル名'
        if ($d.Unit -eq 'シート') { $whereName = 'シート名' }
        Add-Issue '警告' 'ヘッダの違い' ('{0}は「{1}」ですが、{2}は「{3}」です' -f $whereName, $index[$d.Kind].ById[$best].Id, $d.Rule.Column.Text, $d.Id) $d.Place
    }
}

# 一覧との違い: 一覧（画面一覧・テーブル一覧）と、設計書の定義の過不足・名前
foreach ($kind in $kinds) {
    $main = [ordered]@{}; $list = [ordered]@{}
    foreach ($d in $defs) {
        if ($d.Kind -ne $kind) { continue }
        if ($d.Section -eq '定義') { if (-not $main.Contains($d.Key)) { $main[$d.Key] = $d } }
        elseif (-not $list.Contains($d.Key)) { $list[$d.Key] = $d }
    }
    if ($main.Count -eq 0 -or $list.Count -eq 0) { continue }
    $first = @($list.Values)[0]
    $listName = '{0}（{1}）' -f $first.Place.File, $first.Place.Sheet
    foreach ($l in $list.Values) {
        if (-not $main.Contains($l.Key)) {
            Add-Issue '警告' '一覧との違い' ('一覧にある{0}「{1} {2}」の定義が見つかりません（設計書が無いか、ID が違います）' -f $kind, $l.Shown, $l.Name).Replace(' 」', '」') $l.Place
            continue
        }
        $m = $main[$l.Key]
        if ($l.Name -and $m.Name -and (ConvertTo-Label $l.Name) -ne (ConvertTo-Label $m.Name)) {
            Add-Issue '警告' '一覧との違い' ('{0}「{1}」の名前が違います（一覧: {2} / {3}: {4}）' -f $kind, $l.Shown, $l.Name, $m.Place.File, $m.Name) $l.Place
        }
    }
    foreach ($m in $main.Values) {
        if (-not $list.Contains($m.Key)) {
            Add-Issue '警告' '一覧との違い' ('{0}「{1} {2}」が一覧（{3}）にありません' -f $kind, $m.Shown, $m.Name, $listName).Replace(' 」', '」') $m.Place
        }
    }
}

# 参照先なし: 参照している ID が、どの設計書にも定義されていない
$unchecked = [ordered]@{}
$missingCount = @{}
foreach ($ref in $refs) {
    if (-not $index.ContainsKey($ref.Kind)) {
        if (-not $unchecked.Contains($ref.Kind)) { $unchecked[$ref.Kind] = 0 }
        $unchecked[$ref.Kind]++
        continue
    }
    $e = $index[$ref.Kind]
    $context = ''
    if ($ref.Rule.Column) { $context = '（{0}）' -f $ref.Rule.Column.Text }
    if ($ref.Parent -and -not $e.Parents.ContainsKey($ref.ParentKey)) {
        $parentLabel = @($ruleList | Where-Object { $_.Kind -eq $ref.Kind -and $_.Section -ne '参照' -and $_.Parent } | ForEach-Object { $_.Parent.Text })[0]
        Add-Issue 'エラー' '参照先なし' ('{0}「{1}」が定義されていません（{2}）' -f $parentLabel, $ref.Parent, $ref.Shown) $ref.Place
        $missingCount[$ref.Kind]++
        continue
    }
    if (-not $e.ById.ContainsKey($ref.Key)) {
        Add-Issue 'エラー' '参照先なし' ('{0}「{1}」が定義されていません{2}' -f $ref.Kind, $ref.Shown, $context) $ref.Place
        $missingCount[$ref.Kind]++
        continue
    }
    $d = $e.ById[$ref.Key]
    $original = Find-Original $ref.Raw $ref.Shown
    if (-not [string]::Equals($original, $d.Shown, [System.StringComparison]::Ordinal)) {
        Add-Issue '情報' '表記の違い' ('「{0}」は、定義では「{1}」と書かれています（全角・半角や大文字・小文字）' -f $original, $d.Shown) $ref.Place
    }
}
foreach ($k in $unchecked.Keys) {
    Add-Issue '情報' '未確認' ('{0}の定義が見つからないため、{0}の参照（{1} か所）は確かめていません' -f $k, $unchecked[$k]) $null
}

# 組み合わせの違い: 「元」のある参照（画面の遷移先など）を、書かれている設計書どうしで比べる
foreach ($g in @($ruleList | Where-Object { $_.Section -eq '参照' -and $_.From -and $_.Found } | Group-Object Kind)) {
    $rs = @($g.Group)
    if ($rs.Count -lt 2) { continue }
    $reported = @{}
    foreach ($r1 in $rs) {
        foreach ($pk in @($r1.Pairs.Keys | Sort-Object)) {
            $p = $r1.Pairs[$pk]
            foreach ($r2 in $rs) {
                if ($r2.Line -eq $r1.Line -or $r2.Pairs.ContainsKey($pk)) { continue }
                if ($r2.Global) { $where2 = $r2.Place }
                elseif ($r2.Coverage.ContainsKey($p.FromKey)) { $where2 = $r2.Coverage[$p.FromKey] }
                else { continue }
                $rk = '{0}|{1}' -f $pk, $r2.Line
                if ($reported.ContainsKey($rk)) { continue }
                $reported[$rk] = $true
                $label1 = $r1.Kind; if ($r1.Column) { $label1 = $r1.Column.Text }
                $label2 = $r2.Kind; if ($r2.Column) { $label2 = $r2.Column.Text }
                Add-Issue '警告' '組み合わせの違い' ('{0} の{1}「{2}」が、{3}の{4}にありません' -f $p.From, $label1, $p.To, $where2, $label2) $p.Place
            }
        }
    }
}

# ---------------------------------------------------------------- 結果


$levelOrder = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
$sorted = @($script:Issues | Sort-Object @{ e = { if ($_.File) { 0 } else { 1 } } }, @{ e = { $_.File } }, @{ e = { $levelOrder[$_.Level] } }, @{ e = { $_.SheetIndex } }, @{ e = { $_.Row } }, @{ e = { $_.Col } })
$sorted

# 種類ごとの定義と参照の数
$stats = New-Object System.Collections.Generic.List[object]
foreach ($kind in $kinds) {
    $defKeys = @($defs | Where-Object { $_.Kind -eq $kind -and $_.Section -eq '定義' } | ForEach-Object { $_.Key } | Select-Object -Unique).Count
    $listKeys = @($defs | Where-Object { $_.Kind -eq $kind -and $_.Section -eq '一覧' } | ForEach-Object { $_.Key } | Select-Object -Unique).Count
    $refCount = @($refs | Where-Object Kind -eq $kind).Count
    if ($defKeys + $listKeys + $refCount -eq 0) { continue }
    $missing = 0
    if ($missingCount.ContainsKey($kind)) { $missing = $missingCount[$kind] }
    $stats.Add([pscustomobject]@{ Kind = $kind; Defined = $defKeys; Listed = $listKeys; References = $refCount; Missing = $missing; Checked = $index.ContainsKey($kind) })
}

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.File, $i.Level, $i.Item, $i.Location, $i.Detail) }
        if ($i.Sheet) { $row.Link = @{ Column = 3; Path = $i.Path; Location = ("'{0}'!{1}" -f $i.Sheet.Replace("'", "''"), $i.Cell) } }
        $rows.Add($row)
    }
    $statRows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $stats) {
        $missing = [double]$s.Missing
        if (-not $s.Checked) { $missing = '（定義なし）' }
        $statRows.Add(@{ Values = @($s.Kind, [double]$s.Defined, [double]$s.Listed, [double]$s.References, $missing) })
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '指摘一覧'; Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 16 }, @{ Header = '場所'; Width = 22 }, @{ Header = '内容'; Width = 90 }); Rows = $rows },
            @{ Name = '定義と参照'; Columns = @(@{ Header = '種類'; Width = 14 }, @{ Header = '定義'; Width = 8 }, @{ Header = '一覧'; Width = 8 }, @{ Header = '参照'; Width = 8 }, @{ Header = '参照先なし'; Width = 12 }); Rows = $statRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$defText = @($stats | Where-Object { $_.Defined -gt 0 } | ForEach-Object { '{0} {1}' -f $_.Kind, $_.Defined }) -join '・'
if (-not $defText) { $defText = 'なし' }
Write-Host ('{0} ファイルを読みました（定義: {1} / 参照: {2} か所）' -f $books.Count, $defText, $refs.Count)
Write-Host ('エラー {0} 件 / 警告 {1} 件 / 情報 {2} 件' -f @($script:Issues | Where-Object Level -eq 'エラー').Count, @($script:Issues | Where-Object Level -eq '警告').Count, @($script:Issues | Where-Object Level -eq '情報').Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書どうしの参照の食い違い（定義されていない ID、重複した定義、一覧との過不足など）を見つけます（Excel 不要。ファイルは変更しません）。
#
# .DESCRIPTION
#     設計書のフォルダを読み、参照のルール（CSV）に沿って「定義」（テーブル・カラム・画面・メッセージ・コードなど）と
#     「参照」（画面項目の対応項目、イベントの使用テーブル・メッセージ・遷移先など）を集めて突き合わせます。
#
#       参照先なし       参照している ID が、どの設計書にも定義されていない（エラー）
#       重複定義         同じ ID が 2 か所以上で定義されている（エラー）
#       一覧との違い     画面一覧・テーブル一覧と設計書の過不足、名前の違い（警告）
#       ヘッダの違い     1 つのファイルの中でシートによって画面ID・画面名が違う、ファイル名・シート名の ID と違う（警告）
#       組み合わせの違い 画面遷移の一覧とイベント定義の遷移先など、同じ組み合わせを書く場所どうしの過不足（警告）
#       表記の違い       参照の ID の全角・半角、大文字・小文字が定義と違う（情報）
#       未確認           定義が見つからない種類の参照（設計書がフォルダに無いなど。情報）
#
#     結果は 1 指摘 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#     参照先の設計書（テーブル定義書・メッセージ一覧など）も含めて指定してください。
#
# .PARAMETER Rules
#     参照のルールの CSV（区分・種類・ファイル名・シート名・見出し・列・名前・親・元・書き方・重複）。
#     既定は tools\consistency\参照のルール.csv。社内の設計書の書式に合わせて直して使えます。
#
# .PARAMETER Exclude
#     確かめない参照の ID の正規表現。ほかのシステムのテーブルなど、手元に定義の無いものを除きます。例: '^EXT_'
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘一覧と、種類ごとの定義・参照の数。場所をクリックすると該当セルが開きます）。
#
# .EXAMPLE
#     .\Test-DesignConsistency.ps1 D:\設計書
#
# .EXAMPLE
#     .\Test-DesignConsistency.ps1 D:\設計書 -Exclude '^EXT_' -OutXlsx D:\整合性チェック.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     見つけるのは機械的に確かめられる食い違いです。内容の妥当性はレビューで確かめてください。
