[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $BasicDesign,

    [string] $Rules,

    [string] $Exclude,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string[]] $Where,

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
Update-TypeData -TypeName 'DesignConsistency.Usage' -DefaultDisplayPropertySet File, Location, Usage, Id, Text -Force -WhatIf:$false

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


# 「属性」に書ける名前（桁数・型などを、定義と参照で比べる）
$attributeNames = @('桁数', '小数', '型', '必須', '既定値', '入出力')

function ConvertTo-Attributes {
    #     ルールの「属性」（桁数=桁数;必須=必須 のように、属性の名前=表の列の見出し）
    param([string] $Text, [string] $RuleWhere)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($part in ([string]$Text -split '[;；]')) {
        if (-not $part.Trim()) { continue }
        $kv = $part -split '[=＝]', 2
        $name = $kv[0].Trim()
        if ($attributeNames -notcontains $name -or $kv.Count -lt 2 -or -not $kv[1].Trim()) { throw ("{0}: 「属性」は「桁数=桁数;必須=必須」のように、{1} のどれかと列の見出しを = でつないで書いてください（{2}）" -f $RuleWhere, ($attributeNames -join '・'), $part) }
        $list.Add([pscustomobject]@{ Name = $name; Label = (ConvertTo-Label $kv[1]) })
    }
    return , $list
}

$ruleList = New-Object System.Collections.Generic.List[object]
$line = 1
foreach ($r in (Read-CsvFile $Rules @('区分', '種類', 'ファイル名', 'シート名', '見出し', '列', '名前', '親', '元', '書き方', '重複') '参照のルール')) {
    $line++
    $section = ([string]$r.'区分').Trim()
    if (-not $section) { continue }
    $ruleWhere = '参照のルールの {0} 行目' -f $line
    if (@('定義', '一覧', '参照') -notcontains $section) { throw ("{0}: 「区分」は 定義 / 一覧 / 参照 のどれかにしてください（{1}）" -f $ruleWhere, $section) }
    $kind = ([string]$r.'種類').Trim()
    if (-not $kind) { throw "${ruleWhere}: 「種類」が空です" }
    $rule = [pscustomobject]@{
        Line = $line; Section = $section; Kind = $kind; FileRx = $null; SheetRx = $null
        Headers = @(([string]$r.'見出し') -split '/' | ForEach-Object { ConvertTo-Label $_ } | Where-Object { $_ })
        Column = (ConvertTo-Field $r.'列'); Name = (ConvertTo-Field $r.'名前'); Parent = (ConvertTo-Field $r.'親'); From = (ConvertTo-Field $r.'元')
        Pattern = $null; AllowDuplicate = (([string]$r.'重複').Trim() -eq '可')
        Found = $false
        # 2 段階目で足した列（無い CSV も読める）
        Attributes = (ConvertTo-Attributes ([string]$r.'属性') $ruleWhere)
        CheckUnused = (@('確認', '○', '〇', '1') -contains ([string]$r.'未使用').Trim()); Crud = ([string]$r.'CRUD').Trim().ToUpperInvariant()
    }
    if ($rule.Crud -and $rule.Crud -notmatch '^[CRUD]+$') { throw ("{0}: 「CRUD」は C・R・U・D の組み合わせで書いてください（{1}）" -f $ruleWhere, $rule.Crud) }
    if ($rule.Attributes.Count -gt 0 -and $rule.Headers.Count -eq 0) { throw "${ruleWhere}: 「属性」を使うときは「見出し」も指定してください（表の列から読むため）" }
    if (([string]$r.'ファイル名').Trim()) { $rule.FileRx = New-RuleRegex ([string]$r.'ファイル名').Trim() "${ruleWhere}の「ファイル名」" }
    if (([string]$r.'シート名').Trim()) { $rule.SheetRx = New-RuleRegex ([string]$r.'シート名').Trim() "${ruleWhere}の「シート名」" }
    if ($section -eq '参照') {
        $p = ([string]$r.'書き方').Trim()
        if (-not $p) { throw "${ruleWhere}: 参照の「書き方」（正規表現）が空です" }
        $rule.Pattern = New-RuleRegex $p "${ruleWhere}の「書き方」"
        if (@($rule.Pattern.GetGroupNames()) -notcontains 'id') { throw "${ruleWhere}: 「書き方」に ID を表す (?<id>...) がありません" }
        if ((@($rule.Pattern.GetGroupNames()) -contains 'crud' -or $rule.Crud) -and -not $rule.From) { throw "${ruleWhere}: CRUD 図に使う参照（「書き方」の (?<crud>...) か「CRUD」）には、「元」（@画面ID など）も指定してください" }
        if ($rule.Column -and ($rule.Column.IsLabel -or $rule.Headers.Count -eq 0)) { throw "${ruleWhere}: 参照の「列」は表の列の見出しで書き、「見出し」も指定してください" }
        if ($rule.From -and -not $rule.From.IsLabel -and $rule.Headers.Count -eq 0) { throw "${ruleWhere}: 「元」に表の列を指定するときは「見出し」も指定してください" }
    }
    else {
        if (-not $rule.Column) { throw "${ruleWhere}: 「列」（ID の列）が空です" }
        if (-not $rule.Column.IsLabel -and $rule.Headers.Count -eq 0) { throw "${ruleWhere}: 「列」に表の列を指定するときは「見出し」も指定してください（シートの見出しの右の値は「@画面ID」のように書きます）" }
        foreach ($f in @($rule.Name, $rule.Parent)) {
            if ($f -and -not $f.IsLabel -and $rule.Column.IsLabel) { throw "${ruleWhere}: 「列」が「@」で始まるときは、「名前」「親」も「@」で始まる見出しで書いてください" }
        }
    }
    $ruleList.Add($rule)
}
if ($ruleList.Count -eq 0) { throw "参照のルールが空です: $Rules" }

# ---------------------------------------------------------------- 設計書を読む


# 同じ文書の古い版・ツールの出力・書き換える前のファイルと、非表示のシート・雛形のシート（差し込み欄のまま）は読まない（知らせる）
$designFiles = Select-DesignFiles -Path @(@($Path) + @($BasicDesign) | Where-Object { $_ }) -NoRecurse:$NoRecurse -Exclude @($OutXlsx)
$searchArgs = @{ Pattern = '\S'; Path = $designFiles.Paths; Target = 'Value'; SkipHiddenSheets = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = Select-DesignHits @(& $searchScript @searchArgs 6>$null) $designFiles
$books = ConvertTo-DesignBooks $hits

# 基本設計書: -BasicDesign で指定したファイル・フォルダの中のもの、または名前で見分けたもの。
# 基本設計書の一覧・定義は、詳細設計書の定義と別に比べる（同じ画面・テーブルが両方にあっても重複にしない）
$basicRoots = @(foreach ($p in @($BasicDesign | Where-Object { $_ })) { foreach ($item in @(Get-PathItem $p)) { $item.FullName.TrimEnd('\', '/') } })
function Get-DesignSource {
    #     設計書の出どころ: '基本設計' か ''（詳細設計など）。-BasicDesign の中のファイルと、名前で見分けたもの（ExcelTools.psm1 の Get-DesignPhase）
    param([string] $FilePath)
    return Get-DesignPhase $FilePath $basicRoots
}

$script:SourceOf = @{}
foreach ($bookPath in @($books.Keys)) { $script:SourceOf[$bookPath] = Get-DesignSource $bookPath }
# 詳細設計書があれば、CRUD 図は詳細設計書から作る（基本設計書の使用テーブルを混ぜない）
$hasDetail = @($script:SourceOf.Values | Where-Object { -not $_ }).Count -gt 0

$defs = New-Object System.Collections.Generic.List[object]   # 定義と一覧
$defSeen = New-Object 'System.Collections.Generic.HashSet[string]'
$refs = New-Object System.Collections.Generic.List[object]   # 参照
$refSeen = New-Object 'System.Collections.Generic.HashSet[string]'

function New-Place {
    param($Book, [string] $Sheet, [int] $Row, [int] $Col)
    return [pscustomobject]@{
        File = $Book.File; Path = $Book.Path; Sheet = $Sheet; SheetIndex = @($Book.Sheets.Keys).IndexOf($Sheet); Row = $Row; Col = $Col
        Cell = ('{0}{1}' -f (ConvertTo-ColumnName $Col), $Row)
    }
}

function Get-RowAttributes {
    #     表の行から、ルールの「属性」の値を読む（属性の名前 → 値）
    param($Rule, $Table, $DataRow)
    $attrs = @{}
    foreach ($a in $Rule.Attributes) {
        if ($Table.ColumnOf.ContainsKey($a.Label)) { $attrs[$a.Name] = ([string]$DataRow.Cells[$Table.ColumnOf[$a.Label]]).Trim() }
    }
    return $attrs
}

function Add-Definition {
    param($Rule, [string] $Id, [string] $Name, [string] $Parent, $Place, [string] $Unit, $Values, $Attrs)
    $idKey = ConvertTo-Key $Id
    if (-not $idKey) { return }
    $parentKey = ConvertTo-Key $Parent
    $key = $idKey
    $shown = $Id.Trim()
    if ($Rule.Parent) { $key = '{0}.{1}' -f $parentKey, $idKey; $shown = '{0}.{1}' -f ([string]$Parent).Trim(), $Id.Trim() }
    # 同じセルを 2 つのルールで読んだとき（ファイル名が 2 つの一覧のルールに当たるときなど）は 1 つにする
    if (-not $defSeen.Add(('{0}|{1}|{2}|{3}|{4}|{5}' -f $Rule.Section, $Rule.Kind, $key, $Place.Path, $Place.Sheet, $Place.Cell))) { return }
    $defs.Add([pscustomobject]@{
        Rule = $Rule; Section = $Rule.Section; Kind = $Rule.Kind; Id = $Id.Trim(); Shown = $shown; Name = ([string]$Name).Trim()
        ParentKey = $parentKey; Key = $key; Place = $Place; Unit = $Unit; Values = $Values; Attrs = $Attrs; Source = [string]$script:SourceOf[$Place.Path]
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
                Add-Definition $Rule $id $name $parent (New-Place $Book $s $dr.Row $idCol) '行' $null (Get-RowAttributes $Rule $t $dr)
            }
        }
    }
    return $found
}

# CRUD 図: 元（画面ID など）のキー → @{ Shown; Targets = 参照先のキー → @{ Kind; Shown; Letters; Place }; Place }
$script:Crud = [ordered]@{}
# 機能一覧の機能ID → 画面・帳票などの ID の一覧（基本設計の CRUD 図の行が機能 ID のとき）
$script:FunctionMap = @{}

function Add-Crud {
    #     Place: 使用テーブルなどを書いた場所（基本設計の CRUD 図と比べた結果を知らせる場所。最初の場所）
    param([string] $From, [string] $Kind, [string] $TargetShown, [string] $TargetKey, [string] $Letters, $Place)
    $fk = ConvertTo-Key $From
    if (-not $fk -or -not $TargetKey) { return }
    if (-not $script:Crud.Contains($fk)) { $script:Crud[$fk] = [pscustomobject]@{ Shown = $From.Trim(); Targets = @{}; Place = $Place } }
    $targets = $script:Crud[$fk].Targets
    if (-not $targets.ContainsKey($TargetKey)) { $targets[$TargetKey] = [pscustomobject]@{ Kind = $Kind; Shown = $TargetShown; Letters = ''; Place = $Place } }
    foreach ($ch in ($Letters.ToUpperInvariant() -replace '[^CRUD]', '').ToCharArray()) {
        if ($targets[$TargetKey].Letters.IndexOf($ch) -lt 0) { $targets[$TargetKey].Letters += $ch }
    }
}

# 組み合わせを比べる単位（参照のルール × 出どころ）。基本設計書と詳細設計書の画面遷移は、同じルールで読んでも別々に比べる
#     Pairs: 元 > 参照先 → 場所、Coverage: 元（@画面ID）→ 読んだ場所、Global: 「元」が表の列（すべての組み合わせが書かれている）、Place: 表の場所
$refGroups = [ordered]@{}

function Get-RefGroup {
    param($Rule, $Book)
    $source = [string]$script:SourceOf[$Book.Path]
    $k = '{0}|{1}' -f $Rule.Line, $source
    if (-not $refGroups.Contains($k)) {
        $refGroups[$k] = [pscustomobject]@{ Key = $k; Rule = $Rule; Kind = $Rule.Kind; Source = $source; Pairs = @{}; Coverage = @{}; Global = $false; Place = $null }
    }
    return $refGroups[$k]
}

function Add-Reference {
    #     From: 参照元（ルールの「元」）の ID の一覧。1 つのセルに複数の ID が書かれていれば、ID ごとに組み合わせ・CRUD 図に入れる
    param($Rule, [string] $Text, [string[]] $From, $Book, [string] $Sheet, [int] $Row, [int] $Col, $Attrs)
    if (-not $Text) { return }
    $matches2 = $Rule.Pattern.Matches($Text.Normalize($nfkc))
    if ($matches2.Count -eq 0) { return }
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
        # CRUD 図: 「書き方」の (?<crud>...) か、ルールの「CRUD」。カラムの参照（T_ORDER.ORDER_NO）はテーブルへの操作として数える
        $letters = $Rule.Crud
        if ($m.Groups['crud'].Success) { $letters = $m.Groups['crud'].Value }
        $group = Get-RefGroup $Rule $Book
        if ($hasDetail -and $group.Source) { $letters = '' }
        foreach ($f in $From) {
            $fromKey = ConvertTo-Key $f
            if (-not $fromKey) { continue }
            $pairKey = '{0}>{1}' -f $fromKey, $key
            if (-not $group.Pairs.ContainsKey($pairKey)) { $group.Pairs[$pairKey] = [pscustomobject]@{ FromKey = $fromKey; From = $f.Trim(); To = $shown; Place = $Place } }
            if (-not $letters) { continue }
            if ($parent) { Add-Crud $f $Rule.Kind $parent (ConvertTo-Key $parent) $letters $Place } else { Add-Crud $f $Rule.Kind $id (ConvertTo-Key $id) $letters $Place }
        }
        if (-not $refSeen.Add(('{0}|{1}|{2}|{3}|{4}' -f $Rule.Kind, $key, $Place.Path, $Place.Sheet, $Place.Cell))) { continue }
        $refs.Add([pscustomobject]@{ Rule = $Rule; Kind = $Rule.Kind; Shown = $shown; Key = $key; Parent = $parent; ParentKey = (ConvertTo-Key $parent); Raw = $Text; Place = $Place; Attrs = $Attrs })
    }
}

function Add-Coverage {
    #     「元」が「@画面ID」の参照: この画面の参照を読んだことを記録する（組み合わせを比べる範囲）
    param($Rule, [string] $From, $Book, [string] $Sheet)
    $k = ConvertTo-Key $From
    $coverage = (Get-RefGroup $Rule $Book).Coverage
    if ($k -and -not $coverage.ContainsKey($k)) { $coverage[$k] = '{0}（{1}）' -f $Book.File, $Sheet }
}

function Split-Ids {
    #     1 つのセルに書かれた複数の ID（「SCR002、SCR003」）を分ける（「、」「,」「・」「/」とセル内の改行（Search-Excel の ↵）で区切る）
    param([string] $Text)
    return @($Text.Normalize($nfkc) -split '[、,・/↵]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
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
                    $group = Get-RefGroup $Rule $Book
                    $group.Global = $true
                    if (-not $group.Place) { $group.Place = '{0}（{1}）' -f $Book.File, $s }
                    if ($t.ColumnOf.ContainsKey($Rule.From.Label)) { $fromCol = $t.ColumnOf[$Rule.From.Label] }
                }
            }
            $lastFrom = ''
            foreach ($dr in $t.Rows) {
                $from = $fromLabel
                if ($fromCol) {
                    # 元の ID を最初の行にだけ書く表（縦に結合したセルなど。2 行目以降が空）は、上の行の値を使う
                    $fromText = [string]$dr.Cells[$fromCol]
                    if ($fromText.Trim()) { $lastFrom = $fromText } else { $fromText = $lastFrom }
                    $from = Split-Ids $fromText
                }
                if ($Rule.Column) { $cols = @($t.ColumnOf[$Rule.Column.Label]) } else { $cols = @($dr.Cells.Keys | Sort-Object) }
                $attrs = $null
                if ($Rule.Attributes.Count -gt 0) { $attrs = Get-RowAttributes $Rule $t $dr }
                foreach ($col in $cols) {
                    if ($dr.Cells.ContainsKey($col)) { Add-Reference $Rule ([string]$dr.Cells[$col]) $from $Book $s $dr.Row $col $attrs }
                }
                # 機能一覧（機能ID の列のある表）: 機能ID → その行の画面・帳票などの ID（基本設計の CRUD 図の行が機能 ID のとき、詳細設計の行に読み替える）
                if ($Rule.Column -and $t.ColumnOf.ContainsKey('機能ID')) {
                    $functionId = ConvertTo-Key ([string]$dr.Cells[$t.ColumnOf['機能ID']])
                    if ($functionId) {
                        if (-not $script:FunctionMap.ContainsKey($functionId)) { $script:FunctionMap[$functionId] = New-Object System.Collections.Generic.List[string] }
                        foreach ($m in $Rule.Pattern.Matches(([string]$dr.Cells[$cols[0]]).Normalize($nfkc))) {
                            $mk = ConvertTo-Key $m.Groups['id'].Value
                            if ($mk -and -not $script:FunctionMap[$functionId].Contains($mk)) { $script:FunctionMap[$functionId].Add($mk) }
                        }
                    }
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
    # 基本設計書（-BasicDesign の中のファイル・フォルダ名で見分けたもの）は、名前に「基本設計」があるものとしてもルールに当てはめる
    # （「基本設計」の一覧のルールで、名前に基本設計と書いていないファイル（上流\一覧.xlsx など）も読む）
    $basicName = $null
    if ($script:SourceOf[$bookPath]) { $basicName = '基本設計_' + $fileName }
    foreach ($rule in $ruleList) {
        if ($rule.FileRx -and -not $rule.FileRx.IsMatch($fileName) -and -not ($basicName -and $rule.FileRx.IsMatch($basicName))) { continue }
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

# ---------------------------------------------------------------- 影響調査（-Where）


if ($Where) {
    #     指定した ID（カラム・テーブル・メッセージなど）の定義と、参照しているセルの一覧を出す（食い違いは確かめない）
    $usages = New-Object System.Collections.Generic.List[object]
    foreach ($w in $Where) {
        $pattern = ConvertTo-Key $w
        # 「T_ORDER」のように親（テーブル）だけを指定したときは、そのカラムの参照（T_ORDER.ORDER_NO など）も含める
        $isParent = $pattern.IndexOf('.') -lt 0
        $before = $usages.Count
        $items = New-Object System.Collections.Generic.List[object]
        # 同じ定義が何行にも書かれているもの（コード区分など）は最初の行だけ
        $seenDefs = @{}
        foreach ($d in $defs) {
            if ($d.Key -notlike $pattern) { continue }
            $dk = '{0}|{1}|{2}' -f $d.Section, $d.Kind, $d.Key
            if ($seenDefs.ContainsKey($dk)) { continue }
            $seenDefs[$dk] = $true
            $items.Add(@($d.Kind, $d.Shown, $d.Section, $d.Name, $d.Place))
        }
        foreach ($ref in $refs) {
            if ($ref.Key -like $pattern -or ($isParent -and $ref.ParentKey -and $ref.ParentKey -like $pattern)) { $items.Add(@($ref.Kind, $ref.Shown, '参照', $ref.Raw, $ref.Place)) }
        }
        foreach ($it in $items) {
            $pl = $it[4]
            $u = [pscustomobject]@{
                Target = $w; Kind = $it[0]; Id = $it[1]; Usage = $it[2]; Text = $it[3]
                File = $pl.File; Location = ('{0}!{1}' -f $pl.Sheet, $pl.Cell); Sheet = $pl.Sheet; Cell = $pl.Cell; Path = $pl.Path; SheetIndex = $pl.SheetIndex; Row = $pl.Row; Col = $pl.Col
            }
            $u.PSObject.TypeNames.Insert(0, 'DesignConsistency.Usage')
            $usages.Add($u)
        }
        if ($usages.Count -eq $before) { Write-Warning ("「{0}」の定義も参照も見つかりません" -f $w) }
    }
    $usageOrder = @{ '定義' = 0; '一覧' = 1; '参照' = 2 }
    $sortedUsages = @($usages | Sort-Object @{ e = { $_.Target } }, @{ e = { $usageOrder[$_.Usage] } }, @{ e = { $_.File } }, @{ e = { $_.SheetIndex } }, @{ e = { $_.Row } }, @{ e = { $_.Col } })
    $sortedUsages
    if ($OutXlsx) {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($u in $sortedUsages) {
            $rows.Add(@{ Values = @($u.Target, $u.File, $u.Usage, $u.Location, $u.Id, $u.Text); Link = @{ Column = 3; Path = $u.Path; Location = ("'{0}'!{1}" -f $u.Sheet.Replace("'", "''"), $u.Cell) } })
        }
        try {
            Export-Workbook -FilePath $OutXlsx -Sheets @(@{ Name = '影響範囲'; Columns = @(@{ Header = '調べた ID'; Width = 24 }, @{ Header = 'ファイル'; Width = 36 }, @{ Header = '種類'; Width = 7 }, @{ Header = '場所'; Width = 22 }, @{ Header = 'ID'; Width = 26 }, @{ Header = '内容'; Width = 60 }); Rows = $rows })
        }
        catch {
            Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
            $OutXlsx = $null
        }
    }
    Write-Host ''
    foreach ($g in @($sortedUsages | Group-InOrder Target)) {
        $count = { param($kind) @($g.Group | Where-Object Usage -eq $kind).Count }
        Write-Host ('{0}: 定義 {1} か所・一覧 {2} か所・参照 {3} か所（{4} ファイル）' -f $g.Name, (& $count '定義'), (& $count '一覧'), (& $count '参照'), @($g.Group | ForEach-Object { $_.Path } | Select-Object -Unique).Count)
    }
    if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
    return
}

# ---------------------------------------------------------------- 突き合わせる


$script:Issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    #     Kind: 定義・参照の種類（ルールの「種類」。テーブル・メッセージ・画面など）
    param([ValidateSet('エラー', '警告', '情報')] [string] $Level, [string] $Item, [string] $Detail, $Place, [string] $Kind)
    $issue = [pscustomobject]@{ File = ''; Location = ''; Level = $Level; Item = $Item; Detail = $Detail; Kind = $Kind; Sheet = ''; Cell = ''; Path = ''; SheetIndex = 0; Row = 0; Col = 0 }
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

# 重複定義: 同じ ID が 2 か所以上で定義されている（一覧は一覧の中で。基本設計書と詳細設計書は、それぞれの中で）。
# 親のある種類（カラム・画面項目）は、同じファイルの中だけで見る（別のファイルで同じ親の ID を使ったときは、親の重複として 1 回だけ知らせる）
foreach ($section in @('定義', '一覧')) {
    $groups = [ordered]@{}
    foreach ($d in $defs) {
        if ($d.Section -ne $section -or $d.Rule.AllowDuplicate) { continue }
        $gk = '{0}|{1}|{2}' -f $d.Kind, $d.Key, $d.Source
        if ($d.Rule.Parent) { $gk += '|' + $d.Place.Path }
        if (-not $groups.Contains($gk)) { $groups[$gk] = New-Object System.Collections.Generic.List[object] }
        $groups[$gk].Add($d)
    }
    foreach ($g in $groups.Values) {
        if ($g.Count -lt 2) { continue }
        # 2 つ目の場所に、ほかの場所を添えて知らせる
        $others = @($g | Where-Object { $_ -ne $g[1] } | ForEach-Object { $_.Place })
        if ($section -eq '定義') { $fmt = '{0}「{1}」は、ほかの場所でも定義されています（{2}）' } else { $fmt = '{0}「{1}」が、一覧のほかの行にも書かれています（{2}）' }
        $text = $fmt -f $g[0].Kind, $g[0].Shown, (Format-Places $others $g[1].Place.Path)
        # 別々のファイルの一覧（基本設計書と詳細設計書の画面一覧など）なら、基本設計書を見分けられなかったかもしれない
        if ($section -eq '一覧' -and @($g | ForEach-Object { $_.Place.Path } | Select-Object -Unique).Count -gt 1) { $text += '。基本設計書の一覧なら、-BasicDesign で基本設計書を指定してください' }
        Add-Issue 'エラー' '重複定義' $text $g[1].Place $g[0].Kind
    }
}

# ヘッダの違い: 1 つのファイル（シート）の中で ID・名前が食い違う、ファイル名（シート名）に別の ID が書かれている
foreach ($d in $defs) {
    if ($d.Section -ne '定義') { continue }
    if ($d.Unit -eq 'ファイル') {
        foreach ($v in $d.Values.Ids) {
            if ((ConvertTo-Key $v.Value) -ne (ConvertTo-Key $d.Id)) {
                Add-Issue '警告' 'ヘッダの違い' ('{0}が「{1}」です（このファイルのほかのシートは「{2}」）' -f $d.Rule.Column.Text, $v.Value, $d.Id) $v.Place $d.Kind
            }
        }
        foreach ($v in $d.Values.Names) {
            if ((ConvertTo-Label $v.Value) -ne (ConvertTo-Label $d.Name)) {
                Add-Issue '警告' 'ヘッダの違い' ('{0}が「{1}」です（このファイルのほかのシートは「{2}」）' -f $d.Rule.Name.Text, $v.Value, $d.Name) $v.Place $d.Kind
            }
        }
    }
    if (@('ファイル', 'シート') -notcontains $d.Unit -or $d.Rule.Parent) { continue }
    # ファイル名・シート名に書かれた同じ種類の ID のうち、いちばん長いもの（T_ORDER_DETAIL と T_ORDER なら T_ORDER_DETAIL）。
    # 同じ長さなら左に書かれたもの（ハッシュテーブルの順は PowerShell 7 では実行ごとに変わるので、順によらず決める）
    if ($d.Unit -eq 'ファイル') { $target = [System.IO.Path]::GetFileNameWithoutExtension($d.Place.Path) } else { $target = $d.Place.Sheet }
    $target = ConvertTo-Key $target
    $best = $null; $bestAt = 0
    foreach ($k in $index[$d.Kind].ById.Keys) {
        if ($best -and $k.Length -lt $best.Length) { continue }
        $hit = [regex]::Match($target, '(?<![A-Z0-9])' + [regex]::Escape($k) + '(?![A-Z0-9])')
        if (-not $hit.Success -or ($best -and $k.Length -eq $best.Length -and $hit.Index -gt $bestAt)) { continue }
        $best = $k; $bestAt = $hit.Index
    }
    if ($best -and $best -ne $d.Key) {
        $whereName = 'ファイル名'
        if ($d.Unit -eq 'シート') { $whereName = 'シート名' }
        Add-Issue '警告' 'ヘッダの違い' ('{0}は「{1}」ですが、{2}は「{3}」です' -f $whereName, $index[$d.Kind].ById[$best].Id, $d.Rule.Column.Text, $d.Id) $d.Place $d.Kind
    }
}

# 一覧との違い: 一覧（画面一覧・テーブル一覧）と、設計書の定義の過不足・名前。基本設計書の一覧は、詳細設計書の一覧と別に比べる。
# 基本設計との違い: 詳細設計書に定義がある種類は、基本設計書の一覧と定義（基本設計の画面設計書など）をまとめて基本設計の一覧として、
# 詳細設計書の定義と過不足・名前・属性（桁数・必須・入出力）を比べる（基本設計書だけを指定したときは、これまでどおり基本設計書の中で比べる）
#   ID で見つからなければ、同じ親（画面項目なら画面）の中で名前が同じものと比べる（ID が違うと知らせる）。ID の列が名前の列のルール
#   （ID の無い基本設計の画面項目一覧）は、名前で比べる。親のある種類は、両方に親（画面）があるものだけを比べる（親の過不足は親の種類で知らせる）
function Get-Digits {
    #     桁数の数（「12」「11,2」「10桁」など。無ければ $null）
    param([string] $Text)
    $m = [regex]::Match(([string]$Text).Normalize($nfkc), '^\s*(\d+)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return $null
}

function Test-NameOnlyRule {
    #     ID の列が名前の列のルール（ID の無い基本設計の画面項目一覧など）。名前で比べる
    param($Rule)
    return [bool]($Rule.Name -and -not $Rule.Name.IsLabel -and $Rule.Column.Label -eq $Rule.Name.Label)
}

function ConvertTo-IoKey {
    #     入出力の書き方（I・入力・O・出力・表示・IO・入出力）をそろえる
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim().ToUpperInvariant()
    if (-not $t) { return '' }
    $in = $t -match 'I|入'; $out = $t -match 'O|出|表示'
    if ($in -and $out) { return 'IO' }
    if ($in) { return 'I' }
    if ($out) { return 'O' }
    return $t
}

function Compare-DefinitionAttributes {
    #     一覧（基本設計）と定義の属性（桁数・必須・入出力）の違い。両方に列のある属性だけを比べる（型は書き方がまちまちなので比べない）
    param($List, $Main)
    $la = $List.Attrs; $ma = $Main.Attrs
    if (-not $la -or -not $ma) { return @() }
    $diffs = New-Object System.Collections.Generic.List[object]
    if ($la.ContainsKey('桁数') -and $ma.ContainsKey('桁数')) {
        $a = Get-Digits $la['桁数']; $b = Get-Digits $ma['桁数']
        if ($null -ne $a -and $null -ne $b -and $a -ne $b) { $diffs.Add([pscustomobject]@{ Name = '桁数'; List = $la['桁数']; Main = $ma['桁数'] }) }
    }
    if ($la.ContainsKey('必須') -and $ma.ContainsKey('必須') -and (Test-MarkedValue $la['必須']) -ne (Test-MarkedValue $ma['必須'])) {
        $diffs.Add([pscustomobject]@{ Name = '必須'; List = $(if ($la['必須']) { $la['必須'] } else { '（空）' }); Main = $(if ($ma['必須']) { $ma['必須'] } else { '（空）' }) })
    }
    if ($la['入出力'] -and $ma['入出力'] -and (ConvertTo-IoKey $la['入出力']) -ne (ConvertTo-IoKey $ma['入出力'])) {
        $diffs.Add([pscustomobject]@{ Name = '入出力'; List = $la['入出力']; Main = $ma['入出力'] })
    }
    return $diffs.ToArray()
}

$basicAsList = @{}   # 基本設計書の定義を、基本設計の一覧として比べた種類
foreach ($kind in $kinds) {
    $kindDefs = @($defs | Where-Object { $_.Kind -eq $kind })
    $byBasic = @($kindDefs | Where-Object { $_.Section -eq '定義' -and -not $_.Source }).Count -gt 0 -and @($kindDefs | Where-Object { $_.Source }).Count -gt 0
    if ($byBasic) { $basicAsList[$kind] = $true }
    $main = [ordered]@{}
    $lists = [ordered]@{}   # 出どころ → キー → 一覧の行
    foreach ($d in $kindDefs) {
        if ($d.Section -eq '定義' -and -not ($byBasic -and $d.Source)) { if (-not $main.Contains($d.Key)) { $main[$d.Key] = $d }; continue }
        if (-not $lists.Contains($d.Source)) { $lists[$d.Source] = [ordered]@{} }
        # 名前で比べる一覧（ID の無い画面項目一覧）は、同じ名前の行（検索条件と一覧の「受注番号」など）も別々に比べる
        $lk = $d.Key
        if (Test-NameOnlyRule $d.Rule) { $lk = '{0}|{1}|{2}' -f $d.Key, $d.Place.Path, $d.Place.Cell }
        if (-not $lists[$d.Source].Contains($lk)) { $lists[$d.Source][$lk] = $d }
    }
    if ($main.Count -eq 0) { continue }
    # 名前（同じ親の中）→ 定義（設計書の順）
    $byName = @{}
    $mainParents = @{}
    foreach ($m in $main.Values) {
        if ($m.ParentKey) { $mainParents[$m.ParentKey] = $true }
        if (-not $m.Name) { continue }
        $nk = '{0}|{1}' -f $m.ParentKey, (ConvertTo-Label $m.Name).ToUpperInvariant()
        if (-not $byName.ContainsKey($nk)) { $byName[$nk] = New-Object System.Collections.Generic.List[object] }
        $byName[$nk].Add($m)
    }
    foreach ($source in $lists.Keys) {
        $list = $lists[$source]
        $files = @($list.Values | ForEach-Object { $_.Place.File } | Select-Object -Unique)
        $first = @($list.Values)[0]
        $listName = '{0}（{1}）' -f $first.Place.File, $first.Place.Sheet
        if ($files.Count -gt 1) { $listName = '{0} ほか {1} ファイル' -f $first.Place.File, ($files.Count - 1) }
        if ($byBasic -and $source) {
            $item = '基本設計との違い'
            $listLabel = '基本設計'
            $missingFmt = '基本設計にある{0}「{1} {2}」が、詳細設計にありません（設計書が無いか、ID が違います）'
            $absentFmt = '{0}「{1} {2}」が基本設計（{3}）にありません'
        }
        else {
            $item = '一覧との違い'
            $listLabel = '一覧'
            $missingFmt = '一覧にある{0}「{1} {2}」の定義が見つかりません（設計書が無いか、ID が違います）'
            $absentFmt = '{0}「{1} {2}」が一覧（{3}）にありません'
        }
        $listParents = @{}
        foreach ($l in $list.Values) { if ($l.ParentKey) { $listParents[$l.ParentKey] = $true } }
        $matched = @{}   # 一覧の行と比べた定義のキー
        foreach ($l in $list.Values) {
            $byNameOnly = Test-NameOnlyRule $l.Rule
            $m = $null
            if (-not $byNameOnly -and $main.Contains($l.Key)) { $m = $main[$l.Key] }
            if (-not $m -and $l.Name) {
                $candidates = $byName[('{0}|{1}' -f $l.ParentKey, (ConvertTo-Label $l.Name).ToUpperInvariant())]
                if ($byNameOnly) {
                    # 同じ名前が 2 つ以上あれば、まだ比べていないものを設計書の順に
                    if ($candidates) { $m = @($candidates | Where-Object { -not $matched.ContainsKey($_.Key) })[0] }
                }
                elseif ($candidates -and $candidates.Count -eq 1 -and -not $matched.ContainsKey($candidates[0].Key) -and -not $list.Contains($candidates[0].Key)) {
                    # ID で見つからず、同じ名前の定義が 1 つだけ（ほかの行の ID でもない）なら、ID の違い
                    $m = $candidates[0]
                    Add-Issue '警告' $item ('{0}「{1}」の ID が違います（{2}: {3} / {4}: {5}）' -f $kind, $l.Name, $listLabel, $l.Shown, $m.Place.File, $m.Shown) $l.Place $kind
                }
            }
            if (-not $m) {
                # 親（画面）の設計書が無いものは、親の過不足として知らせる
                if ($l.Rule.Parent -and $l.ParentKey -and -not $mainParents.ContainsKey($l.ParentKey)) { continue }
                if ($byNameOnly) { $text = '{0}にある{1}「{2}」が、{3}にありません（名前が違うか、無くした項目です）' -f $listLabel, $kind, $l.Shown, $(if ($listLabel -eq '基本設計') { '詳細設計' } else { '設計書' }) }
                else { $text = ($missingFmt -f $kind, $l.Shown, $l.Name).Replace(' 」', '」') }
                Add-Issue '警告' $item $text $l.Place $kind
                continue
            }
            $matched[$m.Key] = $true
            if (-not $byNameOnly -and $l.Name -and $m.Name -and (ConvertTo-Label $l.Name) -ne (ConvertTo-Label $m.Name)) {
                Add-Issue '警告' $item ('{0}「{1}」の名前が違います（{2}: {3} / {4}: {5}）' -f $kind, $l.Shown, $listLabel, $l.Name, $m.Place.File, $m.Name) $l.Place $kind
            }
            $diffs = @(Compare-DefinitionAttributes $l $m)
            if ($diffs.Count -gt 0) {
                Add-Issue '警告' $item ('{0}「{1}」の{2}が違います（{3}: {4} / {5}: {6}）' -f $kind, $m.Shown, (@($diffs | ForEach-Object { $_.Name }) -join '・'),
                    $listLabel, (@($diffs | ForEach-Object { '{0} {1}' -f $_.Name, $_.List }) -join '・'), $m.Place.File, (@($diffs | ForEach-Object { '{0} {1}' -f $_.Name, $_.Main }) -join '・')) $l.Place $kind
            }
        }
        # 親のある種類（画面項目・カラム）は、一覧に親（画面）があるものだけを、親ごとにまとめて情報にする
        # （基本設計で画面項目を書いた画面だけ。詳細設計で足した項目（ボタンなど）はよくあるため）
        $absentByParent = [ordered]@{}
        foreach ($m in $main.Values) {
            if ($matched.ContainsKey($m.Key)) { continue }
            if ($m.Rule.Parent -and $m.ParentKey) {
                if (-not $listParents.ContainsKey($m.ParentKey)) { continue }
                if (-not $absentByParent.Contains($m.ParentKey)) { $absentByParent[$m.ParentKey] = New-Object System.Collections.Generic.List[object] }
                $absentByParent[$m.ParentKey].Add($m)
                continue
            }
            Add-Issue '警告' $item ($absentFmt -f $kind, $m.Shown, $m.Name, $listName).Replace(' 」', '」') $m.Place $kind
        }
        foreach ($group in $absentByParent.Values) {
            $names = @($group | ForEach-Object { if ($_.Name) { $_.Name } else { $_.Id } })
            $shownNames = @($names | Select-Object -First 10) -join '・'
            if ($names.Count -gt 10) { $shownNames += ' ほか {0} 件' -f ($names.Count - 10) }
            $parentShown = ($group[0].Shown -split '\.')[0]
            Add-Issue '情報' $item ('{0}（{1}）に無い{2}が {3} 件あります（{4}: {5}）' -f $listLabel, $listName, $kind, $group.Count, $parentShown, $shownNames) $group[0].Place $kind
        }
    }
}

function Test-Attributes {
    #     参照（画面項目・IF の項目・外部キーなど）と、参照先の定義（カラム）の属性（ルールの「属性」）を比べる
    param($Ref, $Def)
    $ra = $Ref.Attrs; $da = $Def.Attrs
    # 「入出力」のある参照（画面項目）は入力（I）の項目だけ、無い参照（IF の項目など）はカラムの値を読む側として確かめる
    $hasIo = $ra.ContainsKey('入出力')
    $isInput = $hasIo -and (([string]$ra['入出力']).Normalize($nfkc) -match 'I|入')
    if ($ra['型'] -and $da['型'] -and (ConvertTo-Key $ra['型']) -ne (ConvertTo-Key $da['型'])) {
        Add-Issue '警告' '属性の違い' ('型が「{0}」ですが、{1} は「{2}」です' -f $ra['型'], $Def.Shown, $da['型']) $Ref.Place $Ref.Kind
    }
    foreach ($n in @('桁数', '小数')) {
        if (-not $ra.ContainsKey($n) -or -not $da.ContainsKey($n)) { continue }
        $a = Get-Digits $ra[$n]; $b = Get-Digits $da[$n]
        if ($null -eq $a -or $null -eq $b -or $a -eq $b) { continue }
        # 表示だけの項目は、編集（カンマ・日付の区切りなど）で長くなることがあるので、長いほうは見ない
        if ($hasIo -and -not $isInput -and $a -gt $b) { continue }
        $tail = ''
        if ($isInput -and $a -gt $b) { $tail = '（入力した値がカラムに入りきりません）' }
        Add-Issue '警告' '属性の違い' ('{0}が {1} ですが、{2} は {3} です{4}' -f $n, $a, $Def.Shown, $b, $tail) $Ref.Place $Ref.Kind
    }
    if ($ra.ContainsKey('必須') -and $da.ContainsKey('必須')) {
        # 必須の印（○・要 など）と既定値の無い書き方（「－」「なし」など）は、テーブル定義書から DDL を作るとき（Read-TableDesign）と同じに読む
        $refRequired = Test-MarkedValue $ra['必須']; $defRequired = Test-MarkedValue $da['必須']
        # 入力が任意の項目は、その画面がテーブルを登録・更新する（使用テーブルの CRUD に C・U がある）ときだけ確かめる（検索条件などを除くため）
        $writes = $false
        $owner = $fileOwner[$Ref.Place.Path]
        if ($owner -and $Ref.ParentKey -and $script:Crud.Contains($owner) -and $script:Crud[$owner].Targets.ContainsKey($Ref.ParentKey)) {
            $writes = $script:Crud[$owner].Targets[$Ref.ParentKey].Letters -match '[CU]'
        }
        if ($isInput -and $writes -and -not $refRequired -and $defRequired -and (Test-BlankValue $da['既定値'])) {
            Add-Issue '警告' '属性の違い' ('入力が任意ですが、{0} は必須で既定値がありません（未入力では登録できません）' -f $Def.Shown) $Ref.Place $Ref.Kind
        }
        elseif (-not $hasIo -and $refRequired -and -not $defRequired) {
            Add-Issue '警告' '属性の違い' ('必須の項目ですが、{0} は必須ではありません（値が無いことがあります）' -f $Def.Shown) $Ref.Place $Ref.Kind
        }
    }
}

# ファイル → そのファイルで定義している画面・IF などの ID（CRUD 図の行）
$fileOwner = @{}
foreach ($d in $defs) { if ($d.Section -eq '定義' -and $d.Unit -eq 'ファイル' -and -not $fileOwner.ContainsKey($d.Place.Path)) { $fileOwner[$d.Place.Path] = $d.Key } }

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
        Add-Issue 'エラー' '参照先なし' ('{0}「{1}」が定義されていません（{2}）' -f $parentLabel, $ref.Parent, $ref.Shown) $ref.Place $ref.Kind
        $missingCount[$ref.Kind]++
        continue
    }
    if (-not $e.ById.ContainsKey($ref.Key)) {
        Add-Issue 'エラー' '参照先なし' ('{0}「{1}」が定義されていません{2}' -f $ref.Kind, $ref.Shown, $context) $ref.Place $ref.Kind
        $missingCount[$ref.Kind]++
        continue
    }
    $d = $e.ById[$ref.Key]
    $original = Find-Original $ref.Raw $ref.Shown
    if (-not [string]::Equals($original, $d.Shown, [System.StringComparison]::Ordinal)) {
        Add-Issue '情報' '表記の違い' ('「{0}」は、定義では「{1}」と書かれています（全角・半角や大文字・小文字）' -f $original, $d.Shown) $ref.Place $ref.Kind
    }
    if ($ref.Attrs -and $d.Attrs) { Test-Attributes $ref $d }
}

# 未使用: どこからも参照されていない定義（ルールの「未使用」が「確認」の種類だけ）。
# 子の種類（カラムの親はテーブル）への参照があれば、親の定義も使っているとみなす
$parentKindOf = @{}
foreach ($r in $ruleList) {
    if ($r.Section -ne '定義' -or -not $r.Parent) { continue }
    $p = @($ruleList | Where-Object { $_.Section -eq '定義' -and $_.Kind -ne $r.Kind -and $_.Column.Label -eq $r.Parent.Label })[0]
    if ($p) { $parentKindOf[$r.Kind] = $p.Kind }
}
$used = @{}
foreach ($ref in $refs) {
    $used['{0}|{1}' -f $ref.Kind, $ref.Key] = $true
    if ($ref.ParentKey -and $parentKindOf.ContainsKey($ref.Kind)) { $used['{0}|{1}' -f $parentKindOf[$ref.Kind], $ref.ParentKey] = $true }
}
# 参照を探した種類（参照のルールに当てはまる設計書があったもの）だけを確かめる
$searched = @{}
foreach ($r in $ruleList) {
    if ($r.Section -ne '参照' -or -not $r.Found) { continue }
    $searched[$r.Kind] = $true
    if ($parentKindOf.ContainsKey($r.Kind)) { $searched[$parentKindOf[$r.Kind]] = $true }
}
$reportedUnused = @{}
foreach ($d in $defs) {
    if ($d.Section -ne '定義' -or -not $d.Rule.CheckUnused -or -not $searched.ContainsKey($d.Kind)) { continue }
    # 詳細設計書と比べた基本設計書の定義は、一覧として比べたので確かめない
    if ($d.Source -and $basicAsList.ContainsKey($d.Kind)) { continue }
    $uk = '{0}|{1}' -f $d.Kind, $d.Key
    if ($used.ContainsKey($uk) -or $reportedUnused.ContainsKey($uk)) { continue }
    $reportedUnused[$uk] = $true
    Add-Issue '情報' '未使用' ('{0}「{1} {2}」は、どこからも参照されていません' -f $d.Kind, $d.Shown, $d.Name).Replace(' 」', '」') $d.Place $d.Kind
}
foreach ($k in $unchecked.Keys) {
    Add-Issue '情報' '未確認' ('{0}の定義が見つからないため、{0}の参照（{1} か所）は確かめていません' -f $k, $unchecked[$k]) $null $k
}

# 組み合わせの違い: 「元」のある参照（画面の遷移先など）を、書かれている設計書どうしで比べる（基本設計書と詳細設計書も）
foreach ($g in @($refGroups.Values | Where-Object { $_.Rule.From } | Group-InOrder Kind)) {
    $rs = @($g.Group)
    if ($rs.Count -lt 2) { continue }
    $reported = @{}
    foreach ($r1 in $rs) {
        foreach ($pk in @($r1.Pairs.Keys | Sort-Object)) {
            $p = $r1.Pairs[$pk]
            foreach ($r2 in $rs) {
                if ($r2.Key -eq $r1.Key -or $r2.Pairs.ContainsKey($pk)) { continue }
                if ($r2.Global) { $where2 = $r2.Place }
                elseif ($r2.Coverage.ContainsKey($p.FromKey)) { $where2 = $r2.Coverage[$p.FromKey] }
                else { continue }
                $rk = '{0}|{1}' -f $pk, $r2.Key
                if ($reported.ContainsKey($rk)) { continue }
                $reported[$rk] = $true
                $label1 = $r1.Kind; if ($r1.Rule.Column) { $label1 = $r1.Rule.Column.Text }
                $label2 = $r2.Kind; if ($r2.Rule.Column) { $label2 = $r2.Rule.Column.Text }
                Add-Issue '警告' '組み合わせの違い' ('{0} の{1}「{2}」が、{3}の{4}にありません' -f $p.From, $label1, $p.To, $where2, $label2) $p.Place $r1.Kind
            }
        }
    }
}

# CRUD の違い: 基本設計書の CRUD 図（名前に CRUD のあるシート。行は画面・IF などの ID か機能一覧の機能 ID、列はテーブルの ID か名前）と、
# 詳細設計書の使用テーブル・設定元から作った CRUD 図（-OutXlsx の CRUD シート）を比べる
$basicCrudSheets = New-Object System.Collections.Generic.List[string]
if ($hasDetail -and $script:Crud.Count -gt 0) {
    # テーブル（詳細設計の CRUD 図の列）の ID・名前 → キー
    $tableKinds = @{}
    foreach ($e in $script:Crud.Values) { foreach ($t in $e.Targets.Values) { $tableKinds[$t.Kind] = $true; if ($parentKindOf.ContainsKey($t.Kind)) { $tableKinds[$parentKindOf[$t.Kind]] = $true } } }
    $tableById = @{}; $tableByName = @{}
    foreach ($d in $defs) {
        if (-not $tableKinds.ContainsKey($d.Kind) -or $d.ParentKey) { continue }
        $tableById[$d.Key] = $d
        $nl = ConvertTo-Label $d.Name
        if (-not $nl) { continue }
        if ($tableByName.ContainsKey($nl) -and $tableByName[$nl].Key -ne $d.Key) { $tableByName[$nl] = $null } elseif (-not $tableByName.ContainsKey($nl)) { $tableByName[$nl] = $d }
    }
    $resolveTable = {
        #     列の見出し（T_ORDER・受注・T_ORDER（受注）・受注（T_ORDER）など）→ テーブルのキー（分からなければ $null）
        param([string] $Text)
        $parts = @([string]$Text) + @(([string]$Text).Normalize($nfkc) -split '[（）()\[\]【】↵\s/]' | Where-Object { $_ })
        foreach ($p in $parts) { $k = ConvertTo-Key $p; if ($tableById.ContainsKey($k)) { return $k } }
        foreach ($p in $parts) { $nl = ConvertTo-Label $p; if ($nl -and $tableByName.ContainsKey($nl) -and $tableByName[$nl]) { return $tableByName[$nl].Key } }
        return $null
    }
    # 行の ID になる種類（詳細設計の CRUD 図の行の画面・IF などの種類）
    $rowKinds = @{}
    foreach ($d in $defs) { if ($script:Crud.Contains($d.Key)) { $rowKinds[$d.Kind] = $true } }
    $rowKindOf = @{}
    foreach ($d in $defs) { if ($rowKinds.ContainsKey($d.Kind) -and -not $rowKindOf.ContainsKey($d.Key)) { $rowKindOf[$d.Key] = $d.Kind } }
    $detailDefined = @{}
    foreach ($d in $defs) { if ($d.Section -eq '定義' -and -not $d.Source) { $detailDefined[$d.Key] = $true } }
    $crudOrder = { param([string] $Letters) -join @('C', 'R', 'U', 'D' | Where-Object { $Letters.Contains($_) }) }

    $covered = @{}          # 基本設計の CRUD 図にあった詳細設計の行
    $basicRowKinds = @{}    # 基本設計の CRUD 図にあった行の種類
    $missingColumns = [ordered]@{}   # 基本設計の CRUD 図に列の無いテーブル → 使っている行
    foreach ($bookPath in $bookPaths) {
        if (-not $script:SourceOf[$bookPath]) { continue }
        $book = $books[$bookPath]
        foreach ($s in $book.Sheets.Keys) {
            if ($s -notmatch '(?i)CRUD') { continue }
            $grid = $book.Sheets[$s]
            # 見出しの行: テーブルと分かる見出しが 2 つ以上ある最初の行（列番号 → テーブルのキー。[ordered] は数のキーを位置とみなすので @{}）
            $headerRow = $null; $tableCols = @{}
            foreach ($row in $grid.Rows) {
                $found = @{}
                foreach ($col in $grid.Cols[$row]) { $k = & $resolveTable $grid.Cells[$row][$col]; if ($k) { $found[$col] = $k } }
                if ($found.Count -ge 2) { $headerRow = $row; $tableCols = $found; break }
            }
            if (-not $headerRow) { continue }
            $sheetName = '{0}（{1}）' -f $book.File, $s
            $basicCrudSheets.Add($sheetName)
            $tableColList = [int[]]@($tableCols.Keys)
            [Array]::Sort($tableColList)
            $firstCol = $tableColList[0]
            foreach ($col in $grid.Cols[$headerRow]) {
                if ($col -le $firstCol -or $tableCols.ContainsKey($col)) { continue }
                $label = ([string]$grid.Cells[$headerRow][$col]).Trim()
                if ($label -match '^(備考|名前|名称|No\.?|摘要)$') { continue }
                Add-Issue '情報' '未確認' ('基本設計の CRUD 図の列「{0}」は、テーブルの ID・名前と合わないため比べていません' -f $label) (New-Place $book $s $headerRow $col) 'テーブル'
            }
            $sheetTables = @{}
            foreach ($k in $tableCols.Values) { $sheetTables[$k] = $true }
            foreach ($row in @($grid.Rows | Where-Object { $_ -gt $headerRow })) {
                # 行の ID: 表の列より左のセルのうち、画面・IF などの ID か、機能一覧の機能 ID のもの
                $rowText = $null; $rowCol = 0; $keys = @()
                foreach ($col in $grid.Cols[$row]) {
                    if ($col -ge $firstCol) { break }
                    $k = ConvertTo-Key $grid.Cells[$row][$col]
                    if ($rowKindOf.ContainsKey($k) -or $script:Crud.Contains($k)) { $rowText = $grid.Cells[$row][$col].Trim(); $rowCol = $col; $keys = @($k); break }
                    if ($script:FunctionMap.ContainsKey($k) -and $script:FunctionMap[$k].Count -gt 0) {
                        $rowText = '{0}（{1}）' -f $grid.Cells[$row][$col].Trim(), ($script:FunctionMap[$k] -join '・'); $rowCol = $col; $keys = @($script:FunctionMap[$k]); break
                    }
                }
                if (-not $rowText) { continue }
                # 詳細設計の定義も CRUD も無い行（基本設計にしか無い画面など）は、一覧の違いで知らせるので比べない
                $keys = @($keys | Where-Object { $script:Crud.Contains($_) -or $detailDefined.ContainsKey($_) })
                if ($keys.Count -eq 0) { continue }
                foreach ($k in $keys) { $covered[$k] = $true; if ($rowKindOf.ContainsKey($k)) { $basicRowKinds[$rowKindOf[$k]] = $true } }
                # 詳細設計の CRUD（行が機能なら、その機能の画面などをまとめたもの）
                $detail = @{}
                foreach ($k in $keys) {
                    if (-not $script:Crud.Contains($k)) { continue }
                    foreach ($tk in $script:Crud[$k].Targets.Keys) {
                        $t = $script:Crud[$k].Targets[$tk]
                        if (-not $detail.ContainsKey($tk)) { $detail[$tk] = [pscustomobject]@{ Letters = ''; Shown = $t.Shown; Place = $t.Place } }
                        foreach ($ch in $t.Letters.ToCharArray()) { if ($detail[$tk].Letters.IndexOf($ch) -lt 0) { $detail[$tk].Letters += $ch } }
                    }
                }
                foreach ($col in $tableColList) {
                    $tk = $tableCols[$col]
                    $text = ''
                    if ($grid.Cells[$row].ContainsKey($col)) { $text = ([string]$grid.Cells[$row][$col]).Normalize($nfkc).Trim() }
                    $b = & $crudOrder ($text.ToUpperInvariant() -replace '[^CRUD]', '')
                    $d = ''
                    if ($detail.ContainsKey($tk)) { $d = & $crudOrder $detail[$tk].Letters }
                    $tableShown = $tableById[$tk].Shown
                    if (-not $text -and -not $d) { continue }
                    $place = New-Place $book $s $row $col
                    if ($text -and -not $d) {
                        $what = $b; if (-not $b) { $what = $text }
                        Add-Issue '警告' 'CRUD の違い' ('基本設計の CRUD 図には {0} → {1}（{2}）がありますが、詳細設計（使用テーブルなど）にはありません' -f $rowText, $tableShown, $what) $place 'テーブル'
                    }
                    elseif (-not $text -and $d) {
                        Add-Issue '警告' 'CRUD の違い' ('詳細設計には {0} → {1}（{2}）がありますが、基本設計の CRUD 図（{3}）は空欄です' -f $rowText, $tableShown, $d, $sheetName) $place 'テーブル'
                    }
                    elseif ($b -and $b -ne $d) {
                        Add-Issue '警告' 'CRUD の違い' ('{0} → {1} の CRUD が違います（基本設計: {2} / 詳細設計: {3}）' -f $rowText, $tableShown, $b, $d) $place 'テーブル'
                    }
                }
                # 基本設計の CRUD 図に列の無いテーブルを、詳細設計で使っている
                foreach ($tk in $detail.Keys) {
                    if ($sheetTables.ContainsKey($tk)) { continue }
                    $mk = '{0}|{1}' -f $sheetName, $tk
                    if (-not $missingColumns.Contains($mk)) { $missingColumns[$mk] = [pscustomobject]@{ Table = $detail[$tk].Shown; Sheet = $sheetName; Rows = (New-Object System.Collections.Generic.List[string]); Place = $detail[$tk].Place } }
                    $missingColumns[$mk].Rows.Add(('{0}: {1}' -f $rowText, (& $crudOrder $detail[$tk].Letters)))
                }
            }
        }
    }
    foreach ($m in $missingColumns.Values) {
        $rows = @($m.Rows | Select-Object -First 3) -join '・'
        if ($m.Rows.Count -gt 3) { $rows += ' ほか {0} 件' -f ($m.Rows.Count - 3) }
        Add-Issue '警告' 'CRUD の違い' ('{0} は詳細設計で使っています（{1}）が、基本設計の CRUD 図（{2}）に列がありません' -f $m.Table, $rows, $m.Sheet) $m.Place 'テーブル'
    }
    # 詳細設計の CRUD 図の行（画面・IF など）のうち、基本設計の CRUD 図に無いもの（基本設計の CRUD 図に同じ種類の行があるときだけ）
    if ($basicCrudSheets.Count -gt 0) {
        foreach ($k in $script:Crud.Keys) {
            if ($covered.ContainsKey($k) -or -not $rowKindOf.ContainsKey($k) -or -not $basicRowKinds.ContainsKey($rowKindOf[$k])) { continue }
            $e = $script:Crud[$k]
            $uses = @($e.Targets.Keys | Sort-Object | ForEach-Object { '{0}: {1}' -f $e.Targets[$_].Shown, (& $crudOrder $e.Targets[$_].Letters) }) -join '・'
            Add-Issue '警告' 'CRUD の違い' ('{0} の行が、基本設計の CRUD 図（{1}）にありません（詳細設計: {2}）' -f $e.Shown, ($basicCrudSheets -join '・'), $uses) $e.Place 'テーブル'
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

function Get-CrudSheet {
    #     CRUD 図のシート（行: 画面・IF・帳票・バッチ、列: テーブル。C 登録・R 参照・U 更新・D 削除）
    $names = @{}
    $kindOf = @{}
    foreach ($d in $defs) { if ($d.Section -eq '定義' -and -not $names.ContainsKey($d.Key)) { $names[$d.Key] = $d.Name; $kindOf[$d.Key] = $d.Kind } }
    # 1 列目の見出しは、行の種類（画面・IF・帳票・バッチなど。定義の無い行だけなら「画面・IF」）
    $rowKindNames = @($script:Crud.Keys | Where-Object { $kindOf.ContainsKey($_) } | ForEach-Object { $kindOf[$_] } | Select-Object -Unique)
    $order = @('画面', 'IF', '帳票', 'バッチ')
    $rowKindNames = @(@($order | Where-Object { $rowKindNames -contains $_ }) + @($rowKindNames | Where-Object { $order -notcontains $_ }))
    $firstHeader = '画面・IF'
    if ($rowKindNames.Count -gt 0) { $firstHeader = $rowKindNames -join '・' }
    $targets = [ordered]@{}
    foreach ($k in @($script:Crud.Values | ForEach-Object { $_.Targets.Keys } | Sort-Object -Unique)) {
        $t = @($script:Crud.Values | Where-Object { $_.Targets.ContainsKey($k) })[0].Targets[$k]
        $targets[$k] = $t.Shown
    }
    $columns = @(@{ Header = $firstHeader; Width = 12 }, @{ Header = '名前'; Width = 20 })
    foreach ($k in $targets.Keys) {
        $header = $targets[$k]
        if ($names[$k]) { $header = '{0}（{1}）' -f $targets[$k], $names[$k] }
        $columns += @{ Header = $header; Width = 10 }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($fk in @($script:Crud.Keys | Sort-Object)) {
        $entry = $script:Crud[$fk]
        $values = @($entry.Shown, [string]$names[$fk])
        foreach ($k in $targets.Keys) {
            $letters = ''
            if ($entry.Targets.ContainsKey($k)) { $l = $entry.Targets[$k].Letters; $letters = -join @('C', 'R', 'U', 'D' | Where-Object { $l.Contains($_) }) }
            $values += $letters
        }
        $rows.Add(@{ Values = $values })
    }
    return @{ Name = 'CRUD'; Columns = $columns; Rows = $rows }
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
    $sheets = @(
        @{ Name = '指摘一覧'; Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 16 }, @{ Header = '場所'; Width = 22 }, @{ Header = '内容'; Width = 90 }); Rows = $rows },
        @{ Name = '定義と参照'; Columns = @(@{ Header = '種類'; Width = 14 }, @{ Header = '定義'; Width = 8 }, @{ Header = '一覧'; Width = 8 }, @{ Header = '参照'; Width = 8 }, @{ Header = '参照先なし'; Width = 12 }); Rows = $statRows })
    if ($script:Crud.Count -gt 0) { $sheets += Get-CrudSheet }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets $sheets
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
$basicBooks = @($books.Keys | Where-Object { $script:SourceOf[$_] })
if ($basicBooks.Count -gt 0 -and $hasDetail) {
    $basicNames = @($basicBooks | Sort-Object | Select-Object -First 3 | ForEach-Object { $books[$_].File }) -join '・'
    if ($basicBooks.Count -gt 3) { $basicNames += ' ほか {0} ファイル' -f ($basicBooks.Count - 3) }
    Write-Host ('うち基本設計書 {0} ファイル（{1}）。基本設計書の一覧・定義は、詳細設計書と別に比べました' -f $basicBooks.Count, $basicNames)
    if ($basicCrudSheets.Count -gt 0) { Write-Host ('基本設計の CRUD 図（{0}）を、詳細設計の使用テーブルから作った CRUD 図と比べました' -f ($basicCrudSheets -join '・')) }
}
Write-Host ('エラー {0} 件 / 警告 {1} 件 / 情報 {2} 件' -f @($script:Issues | Where-Object Level -eq 'エラー').Count, @($script:Issues | Where-Object Level -eq '警告').Count, @($script:Issues | Where-Object Level -eq '情報').Count)
if ($OutXlsx) {
    Write-Host "結果を保存しました: $OutXlsx"
    if ($script:Crud.Count -gt 0) { Write-Host ('CRUD 図（CRUD シート）: 画面・IF {0} 件 × テーブル {1} 件' -f $script:Crud.Count, @($script:Crud.Values | ForEach-Object { $_.Targets.Keys } | Sort-Object -Unique).Count) }
}

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
#       基本設計との違い 基本設計書の一覧（画面一覧・テーブル一覧・IF 一覧・帳票一覧・バッチ一覧・画面項目一覧）・定義と、詳細設計書の定義の過不足、
#                        名前・ID の違い、画面項目の桁数・必須・入出力の違い（警告。基本設計に無い詳細設計の画面項目は、画面ごとにまとめて情報）
#       CRUD の違い      基本設計書の CRUD 図（名前に CRUD のあるシート）と、詳細設計書の使用テーブル・設定元から作った CRUD 図の違い（警告）
#       ヘッダの違い     1 つのファイルの中でシートによって画面ID・画面名が違う、ファイル名・シート名の ID と違う（警告）
#       組み合わせの違い 画面遷移の一覧とイベント定義の遷移先など、同じ組み合わせを書く場所どうしの過不足（警告）
#       属性の違い       画面項目・IF の項目・外部キーと、参照先のカラムの桁数・型が違う。入力が任意なのにカラムが
#                        必須で既定値が無い（登録・更新する画面だけ）。IF の必須の項目なのにカラムが必須でない（警告）
#       表記の違い       参照の ID の全角・半角、大文字・小文字が定義と違う（情報）
#       未使用           どこからも参照されていないメッセージ・コード区分・テーブル（ルールの「未使用」が「確認」の種類。情報）
#       未確認           定義が見つからない種類の参照（設計書がフォルダに無いなど。情報）
#
#     結果は 1 指摘 1 行のオブジェクト（Kind は定義・参照の種類）で出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#     -OutXlsx には、イベント定義の使用テーブル（T_ORDER(R,U) など）・IF の設定元・帳票設計書とバッチ設計書の使用テーブルから作った CRUD 図も入れます。
#     -Where を付けると、食い違いは確かめずに、指定した ID の定義と参照している場所（影響範囲）を一覧にします。
#
#     基本設計書も一緒に指定すると、基本設計書の一覧・定義を詳細設計書と別に比べます（同じ画面・テーブルが両方にあっても重複にしません）。
#     基本設計書は、ファイル名から上のフォルダへ順に見て、名前に「基本設計・外部設計・概要設計」か先頭の「BD」（BD01_画面一覧.xlsx）が
#     あるもの（「詳細設計・内部設計」か先頭の「DD」が先に見つかれば詳細設計）と、-BasicDesign で指定したものです。
#     画面遷移も、基本設計書と詳細設計書（画面一覧の画面遷移・イベント定義の遷移先）で比べます。
#     基本設計書の画面項目一覧（名前に「画面項目」のあるシートの、画面ID・項目名の表）は、詳細設計書の画面項目定義と、画面ごとに
#     項目名で組にして（同じ名前は上から順に）、過不足と桁数・必須・入出力の違いを比べます。
#     基本設計書の CRUD 図（行は画面・IF などの ID か機能一覧の機能 ID、列はテーブルの ID か名前）は、詳細設計書の使用テーブルから作った
#     CRUD 図と、行 × テーブルごとに C・R・U・D を比べます（○ など C・R・U・D でない印は「使う」として、使っているかだけを比べます）。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#     参照先の設計書（テーブル定義書・メッセージ一覧など）も含めて指定してください。
#     同じ文書の古い版・ツールが書き出したファイル・書き換える前のファイル（-BackupDir）と、非表示のシート・雛形のシート（ID が {{テーブルID}} などのまま）は読みません（読まなかったものは警告で知らせます）。
#
# .PARAMETER BasicDesign
#     基本設計書のフォルダまたはファイル（複数可）。名前で基本設計書と見分けられないとき（上流\一覧.xlsx など）に指定します。
#     -Path と一緒に読み、中のファイルは名前に「基本設計」があるものとして参照のルールに当てはめます。
#
# .PARAMETER Rules
#     参照のルールの CSV（区分・種類・ファイル名・シート名・見出し・列・名前・親・元・書き方・重複・属性・未使用・CRUD。
#     属性・未使用・CRUD の列は無くてもかまいません）。
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
# .PARAMETER Where
#     影響調査: 指定した ID（T_ORDER.DELIVERY_TIME_KBN・E0001・C005 など。複数可・ワイルドカード可）の定義と、
#     参照しているセルを一覧にします（食い違いは確かめません）。テーブルだけ（T_ORDER）を指定すると、
#     そのカラムの参照（T_ORDER.ORDER_NO など）も含めます。Compare-Excel.ps1 で変わったカラムを渡せば、直すべき画面・IF が分かります。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘一覧、種類ごとの定義・参照の数、CRUD 図。場所をクリックすると該当セルが開きます）。
#     -Where を付けたときは、影響範囲の一覧を保存します。
#
# .EXAMPLE
#     .\Test-DesignConsistency.ps1 D:\設計書
#
# .EXAMPLE
#     .\Test-DesignConsistency.ps1 D:\設計書 -Exclude '^EXT_' -OutXlsx D:\整合性チェック.xlsx
#
# .EXAMPLE
#     # 基本設計書と詳細設計書の過不足・名前の違い・画面遷移の違いも確かめる
#     .\Test-DesignConsistency.ps1 D:\設計書\詳細設計 -BasicDesign D:\設計書\基本設計
#
# .EXAMPLE
#     # カラムの桁数を変える前に、使っている画面・IF を調べる
#     .\Test-DesignConsistency.ps1 D:\設計書 -Where 'T_ORDER.DELIVERY_TIME_KBN'
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     見つけるのは機械的に確かめられる食い違いです。内容の妥当性はレビューで確かめてください。
