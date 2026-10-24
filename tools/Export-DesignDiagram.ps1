[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateSet('ScreenFlow', 'ER')]
    [string[]] $Kind = @('ScreenFlow', 'ER'),

    [ValidateSet('Key', 'All', 'None')]
    [string] $Columns = 'Key',

    [string] $OutMarkdown,

    [string] $OutXlsx,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'ddl/テーブル定義書の見出し.csv' }
foreach ($name in @('Headers', 'OutMarkdown', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'DesignDiagram.Edge' -DefaultDisplayPropertySet Kind, From, To, Label -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# ---------------------------------------------------------------- 設計書を読む


$searchArgs = @{ Pattern = '\S'; Path = $Path; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$books = ConvertTo-DesignBooks $hits

$edges = New-Object System.Collections.Generic.List[object]
$findings = New-Object System.Collections.Generic.List[string]
function Add-Edge {
    param([string] $EdgeKind, $From, $To, [string] $Label, [string] $Note, $Place)
    $e = [pscustomobject]@{ Kind = $EdgeKind; From = $From.Id; FromName = $From.Name; To = $To.Id; ToName = $To.Name; Label = $Label; Note = $Note; File = $Place.File; Path = $Place.Path; Sheet = $Place.Sheet; Row = $Place.Row }
    $e.PSObject.TypeNames.Insert(0, 'DesignDiagram.Edge')
    $edges.Add($e)
}

# 画面（画面一覧の表と、画面設計書のヘッダの画面 ID・画面名）。Id → @{ Id; Name; Missing }
$screens = [ordered]@{}
$screenFlow = $Kind -contains 'ScreenFlow'
if ($screenFlow) {
    $design = Read-ScreenDesign $books
    foreach ($book in $books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            foreach ($t in (Find-DesignTables $grid @('画面ID', '画面名') '画面ID')) {
                $labels = @($grid.Cols[$t.HeaderRow] | ForEach-Object { $grid.Labels[$t.HeaderRow][$_] })
                $at = [array]::IndexOf($labels, '画面ID')
                if ($at -lt 0 -or $at + 1 -ge $labels.Count -or $labels[$at + 1] -ne '画面名') { continue }
                foreach ($dr in $t.Rows) {
                    $id = ([string]$dr.Cells[$t.ColumnOf['画面ID']]).Normalize($nfkc).Trim()
                    if ($id -match '^[A-Za-z][A-Za-z0-9_\-]*\d' -and -not $screens.Contains($id.ToUpperInvariant())) {
                        $screens[$id.ToUpperInvariant()] = [pscustomobject]@{ Id = $id; Name = ([string]$dr.Cells[$t.ColumnOf['画面名']]).Trim(); Missing = $false }
                    }
                }
            }
        }
    }
    # 画面設計書のヘッダの画面 ID も、画面一覧と同じく全角・半角をそろえる
    foreach ($x in @($design.Items.ToArray()) + @($design.Events.ToArray())) {
        $sid = ([string]$x.Screen).Normalize($nfkc).Trim()
        if ($sid -and -not $screens.Contains($sid.ToUpperInvariant())) { $screens[$sid.ToUpperInvariant()] = [pscustomobject]@{ Id = $sid; Name = ([string]$x.ScreenName).Trim(); Missing = $false } }
    }
    # 画面名 → 画面 ID（遷移先に画面名だけを書いたとき）
    $byName = @{}
    foreach ($s in $screens.Values) { if ($s.Name) { $byName[(ConvertTo-Label $s.Name)] = $s } }
    foreach ($ev in $design.Events) {
        $next = ([string]$ev.Next).Normalize($nfkc).Trim()
        # 遷移しないことを表す書き方（-・なし・N/A など）
        if (-not $next -or $next -match '^(?i)(-|－|ー|―|‐|−|×|なし|無し|N/A)$' -or -not $ev.Screen) { continue }
        $from = $screens[([string]$ev.Screen).Normalize($nfkc).Trim().ToUpperInvariant()]
        $targets = New-Object System.Collections.Generic.List[object]
        foreach ($m in [regex]::Matches($next, '(?<![A-Za-z0-9])[A-Za-z][A-Za-z0-9_\-]*\d[A-Za-z0-9_\-]*')) {
            $id = $m.Value.ToUpperInvariant()
            if (-not $screens.Contains($id)) {
                $findings.Add(('{0} の{1}（{2}）の遷移先「{3}」が、画面一覧にありません' -f $ev.Screen, $ev.Name, $ev.Id, $m.Value))
                $screens[$id] = [pscustomobject]@{ Id = $m.Value; Name = ''; Missing = $true }
            }
            if (-not $targets.Contains($screens[$id])) { $targets.Add($screens[$id]) }
        }
        if ($targets.Count -eq 0) {
            $label = ConvertTo-Label $next
            $match = @($byName.Keys | Where-Object { $label.Contains($_) } | Sort-Object Length -Descending | Select-Object -First 1)
            if ($match.Count -gt 0) { $targets.Add($byName[$match[0]]) }
            else { $findings.Add(('{0} の{1}（{2}）の遷移先「{3}」の画面が分かりません（画面 ID を書いてください）' -f $ev.Screen, $ev.Name, $ev.Id, $next)) }
        }
        foreach ($to in $targets) { Add-Edge '画面遷移' $from $to $ev.Name ('{0} {1}' -f $ev.Id, $ev.Name).Trim() $ev }
    }
}

# テーブルと外部キー。Id → @{ Id; Name; Columns; Missing }
$tableMap = [ordered]@{}
if ($Kind -contains 'ER') {
    $tables = Read-TableDesign -Books $books -Candidates (Read-TableHeaderCandidates $Headers)
    foreach ($t in $tables) { if (-not $tableMap.Contains($t.Id.ToUpperInvariant())) { $tableMap[$t.Id.ToUpperInvariant()] = [pscustomobject]@{ Id = $t.Id; Name = $t.Name; Columns = $t.Columns; Missing = $false; Place = $t.Place } } }
    foreach ($t in @($tableMap.Values)) {
        foreach ($c in $t.Columns) {
            if (-not $c.FkTable) { continue }
            $key = $c.FkTable.ToUpperInvariant()
            if (-not $tableMap.Contains($key)) {
                $findings.Add(('{0}.{1} の外部キーの参照先「{2}」が、テーブル定義書にありません' -f $t.Id, $c.Name, $c.FkTable))
                $tableMap[$key] = [pscustomobject]@{ Id = $c.FkTable; Name = ''; Columns = @(); Missing = $true; Place = $null }
            }
            $place = [pscustomobject]@{ File = $c.Place.File; Path = $c.Place.Path; Sheet = $c.Place.Sheet; Row = [int]([regex]::Match($c.Place.Cell, '\d+$').Value) }
            Add-Edge '外部キー' $t $tableMap[$key] $c.Name ('{0}.{1} → {2}.{3}' -f $t.Id, $c.Name, $c.FkTable, $c.FkColumn).TrimEnd('.') $place
        }
    }
}

# ---------------------------------------------------------------- 並べ方（階層ごとに左から右へ）


function Get-Layers {
    #     ノードを、矢印の元から先へ階層に分ける（入ってくる矢印の無いノードから幅優先。届かないノードは新しい起点）。
    #     Nodes: ID（大文字）の順の一覧。Links: @(@(元, 先)...)。戻り値: ID → @{ Level; Index }
    param([string[]] $Nodes, $Links)
    $out = @{}; $in = @{}
    foreach ($n in $Nodes) { $out[$n] = New-Object System.Collections.Generic.List[string]; $in[$n] = 0 }
    foreach ($l in $Links) { if ($l[0] -ne $l[1] -and -not $out[$l[0]].Contains($l[1])) { $out[$l[0]].Add($l[1]); $in[$l[1]]++ } }
    $place = @{}
    $count = @{}
    $visit = {
        param([string] $Root)
        $queue = New-Object System.Collections.Generic.Queue[string]
        $place[$Root] = @{ Level = 0 }
        $queue.Enqueue($Root)
        while ($queue.Count -gt 0) {
            $n = $queue.Dequeue()
            foreach ($m in $out[$n]) {
                if ($place.ContainsKey($m)) { continue }
                $place[$m] = @{ Level = $place[$n].Level + 1 }
                $queue.Enqueue($m)
            }
        }
    }
    foreach ($n in $Nodes) { if ($in[$n] -eq 0 -and -not $place.ContainsKey($n)) { & $visit $n } }
    foreach ($n in $Nodes) { if (-not $place.ContainsKey($n)) { & $visit $n } }
    # 階層の中の順は、ノードの順
    foreach ($n in $Nodes) {
        $lv = $place[$n].Level
        if (-not $count.ContainsKey($lv)) { $count[$lv] = 0 }
        $place[$n].Index = $count[$lv]
        $count[$lv]++
    }
    return $place
}

function Get-DepthLayers {
    #     ER 図の階層: 参照されるテーブル（親）を左に、参照するテーブル（子）を右に置く（親のいちばん深い階層 + 1）。
    #     互いに参照するテーブル（循環）は、深さ優先でたどって戻る矢印を除いてから階層を決める
    param([string[]] $Nodes, $Links)
    $out = @{}; $in = @{}
    foreach ($n in $Nodes) { $out[$n] = New-Object System.Collections.Generic.List[string]; $in[$n] = 0 }
    foreach ($l in $Links) { if ($l[0] -ne $l[1] -and -not $out[$l[0]].Contains($l[1])) { $out[$l[0]].Add($l[1]); $in[$l[1]]++ } }
    $state = @{}
    $kept = New-Object System.Collections.Generic.List[object]
    $roots = @(@($Nodes | Where-Object { $in[$_] -eq 0 }) + @($Nodes))
    foreach ($root in $roots) {
        if ($state.ContainsKey($root)) { continue }
        # 1: たどっている途中・2: たどり終わった
        $state[$root] = 1
        $stack = New-Object System.Collections.Generic.Stack[object]
        $stack.Push(@($root, 0))
        while ($stack.Count -gt 0) {
            $top = $stack.Pop()
            $n = $top[0]; $i = $top[1]
            if ($i -ge $out[$n].Count) { $state[$n] = 2; continue }
            $stack.Push(@($n, ($i + 1)))
            $m = $out[$n][$i]
            if (-not $state.ContainsKey($m)) { $kept.Add(@($n, $m)); $state[$m] = 1; $stack.Push(@($m, 0)) }
            elseif ($state[$m] -eq 2) { $kept.Add(@($n, $m)) }
        }
    }
    $level = @{}
    foreach ($n in $Nodes) { $level[$n] = 0 }
    for ($i = 0; $i -lt $Nodes.Count; $i++) {
        $changed = $false
        foreach ($l in $kept) {
            if ($level[$l[1]] -lt $level[$l[0]] + 1) { $level[$l[1]] = $level[$l[0]] + 1; $changed = $true }
        }
        if (-not $changed) { break }
    }
    $place = @{}; $count = @{}
    foreach ($n in $Nodes) {
        $lv = $level[$n]
        if (-not $count.ContainsKey($lv)) { $count[$lv] = 0 }
        $place[$n] = @{ Level = $lv; Index = $count[$lv] }
        $count[$lv]++
    }
    return $place
}

function New-DiagramShapes {
    #     ノードと矢印を、階層ごとの列に並べた図形にする（Export-Workbook の Shapes）。
    #     Nodes: @{ Key; Text; Body; Fill; Line; Dash } の一覧。
    #     Pairs: @{ A; B; AtoB = @(ラベル...); BtoA = @(ラベル...); Forward（A → B がある）; Backward（B → A がある） } の一覧（A・B は Key）
    #     縦の位置: いちばん左の列は上から順に積み、ほかの列は、左の列でつながる四角の中ほどの高さの平均にそろえる（重なれば下へずらす）
    param($Nodes, $Place, $Pairs, [int] $Width, [int] $GapX, [int] $GapY, [int] $Top = 3)
    $shapes = New-Object System.Collections.Generic.List[object]
    $box = @{}
    $neighbors = @{}
    foreach ($n in $Nodes) { $neighbors[$n.Key] = New-Object System.Collections.Generic.List[string] }
    foreach ($pair in $Pairs) {
        if ($pair.A -eq $pair.B -or -not $neighbors.ContainsKey($pair.A) -or -not $neighbors.ContainsKey($pair.B)) { continue }
        $neighbors[$pair.A].Add($pair.B); $neighbors[$pair.B].Add($pair.A)
    }
    $id = 2
    $levels = @($Nodes | ForEach-Object { $Place[$_.Key].Level } | Sort-Object -Unique)
    foreach ($lv in $levels) {
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($n in @($Nodes | Where-Object { $Place[$_.Key].Level -eq $lv })) {
            $lines = @($n.Text).Count + @($n.Body | Where-Object { $null -ne $_ }).Count
            $height = [Math]::Max(6, 2 * $lines + 2)
            # 左の列でつながる四角の中ほどの高さの平均（無ければ、上から順）
            $centers = @($neighbors[$n.Key] | Where-Object { $box.ContainsKey($_) } | ForEach-Object { $box[$_].Row + $box[$_].Height / 2 })
            $want = [double]::MaxValue
            if ($centers.Count -gt 0) { $want = ($centers | Measure-Object -Average).Average - $height / 2 }
            $list.Add([pscustomobject]@{ Node = $n; Height = $height; Want = $want; Index = $Place[$n.Key].Index })
        }
        $list.Sort([System.Comparison[object]] { param($x, $y) if ($x.Want -ne $y.Want) { return $x.Want.CompareTo($y.Want) }; return $x.Index.CompareTo($y.Index) })
        $next = $Top
        foreach ($item in $list) {
            $row = $next
            if ($item.Want -ne [double]::MaxValue) { $row = [Math]::Max($next, [int][Math]::Round($item.Want)) }
            $n = $item.Node
            $b = @{ Type = 'Box'; Id = $id; Name = $n.Key; Col = 1 + $lv * ($Width + $GapX); Row = $row; Width = $Width; Height = $item.Height; Text = $n.Text; Body = $n.Body; Fill = $n.Fill; Line = $n.Line; Dash = $n.Dash; Level = $lv }
            $next = $row + $item.Height + $GapY
            $box[$n.Key] = $b
            $shapes.Add($b)
            $id++
        }
    }
    foreach ($pair in $Pairs) {
        $a = $box[$pair.A]; $b = $box[$pair.B]
        if (-not $a -or -not $b) { continue }
        if ($pair.A -eq $pair.B) {
            # 同じ画面への遷移（同じテーブルへの外部キー）は、四角のすぐ下に「↻ ラベル」と書く
            $self = ('↻ ' + ($pair.AtoB -join '・')).TrimEnd()
            $shapes.Add(@{ Type = 'Text'; Id = $id; Name = ('{0} のラベル' -f $a.Name); Col = $a.Col; Row = $a.Row + $a.Height; Width = $a.Width; Height = 2; Text = @($self); Align = 'l' })
            $id++
            continue
        }
        # 左右に並ぶなら右端の中ほどから左端の中ほどへ、同じ列なら下端から上端へ
        $vertical = $a.Level -eq $b.Level
        $forward = $a.Level -lt $b.Level -or ($vertical -and $a.Row -lt $b.Row)
        $first = $a; $second = $b; $firstLabels = @($pair.AtoB); $secondLabels = @($pair.BtoA); $firstOn = $pair.Forward; $secondOn = $pair.Backward
        if (-not $forward) { $first = $b; $second = $a; $firstLabels = @($pair.BtoA); $secondLabels = @($pair.AtoB); $firstOn = $pair.Backward; $secondOn = $pair.Forward }
        if ($vertical) {
            $from = @(($first.Col + $first.Width / 2), ($first.Row + $first.Height)); $to = @(($second.Col + $second.Width / 2), $second.Row)
            $sites = @(2, 0); $marks = @('↓', '↑')
        }
        else {
            $from = @(($first.Col + $first.Width), ($first.Row + $first.Height / 2)); $to = @($second.Col, ($second.Row + $second.Height / 2))
            $sites = @(3, 1); $marks = @('→', '←')
        }
        # 矢じりは向かう先に。両向きなら両端に（ラベルの無い遷移も向きは数える）
        $arrow = @{ Type = 'Arrow'; Id = $id; Name = ('{0}-{1}' -f $first.Name, $second.Name); From = $from; To = $to; StartShape = $first.Id; StartSite = $sites[0]; EndShape = $second.Id; EndSite = $sites[1]; Head = [bool]$secondOn }
        if (-not $firstOn) { $arrow = @{ Type = 'Arrow'; Id = $id; Name = ('{0}-{1}' -f $second.Name, $first.Name); From = $to; To = $from; StartShape = $second.Id; StartSite = $sites[1]; EndShape = $first.Id; EndSite = $sites[0]; Head = $false } }
        $shapes.Add($arrow)
        $id++
        $text = New-Object System.Collections.Generic.List[string]
        if ($firstOn -and $secondOn) {
            if ($firstLabels.Count -gt 0) { $text.Add(('{0} {1}' -f $marks[0], ($firstLabels -join '・'))) }
            if ($secondLabels.Count -gt 0) { $text.Add(('{0} {1}' -f $marks[1], ($secondLabels -join '・'))) }
        }
        else { $text.Add((@($firstLabels) + @($secondLabels) | Where-Object { $_ }) -join '・') }
        if (-not ($text -join '')) { continue }
        # ラベルは線と四角に重ならないところへ: 縦の線は右、横に近い線は上、右下がりの線は左下、右上がりの線は左上
        $midCol = [int][Math]::Floor(($from[0] + $to[0]) / 2); $midRow = [int][Math]::Floor(($from[1] + $to[1]) / 2)
        # 列をいくつもまたぐ線は、途中の列の四角に重ならないよう、左の四角のすぐ右の隙間に置く
        $spanX = $to[0] - $from[0]
        if (-not $vertical -and [Math]::Abs($spanX) -gt $GapX + 1) {
            $t = ($GapX / 2) / [Math]::Abs($spanX)
            $midCol = [int][Math]::Floor($from[0] + $spanX * $t); $midRow = [int][Math]::Floor($from[1] + ($to[1] - $from[1]) * $t)
        }
        $labelHeight = 2 * $text.Count
        $labelWidth = $GapX + 2
        $label = @{ Type = 'Text'; Id = $id; Name = ('{0} のラベル' -f $arrow.Name); Width = $labelWidth; Height = $labelHeight; Text = $text.ToArray(); Align = 'l' }
        $dy = [Math]::Abs($to[1] - $from[1])
        if ($vertical) { $label.Col = $midCol + 1; $label.Row = $midRow - $text.Count }
        elseif ($dy -le [Math]::Max($first.Height, $second.Height) / 2 + 2) {
            $label.Col = $midCol - [int][Math]::Floor($labelWidth / 2); $label.Row = [Math]::Min($midRow, [Math]::Min($from[1], $to[1])) - $labelHeight; $label.Align = 'ctr'
        }
        else {
            $leftY = $from[1]; $rightY = $to[1]
            if ($from[0] -gt $to[0]) { $leftY = $to[1]; $rightY = $from[1] }
            $label.Col = $midCol - 1 - $labelWidth; $label.Align = 'r'
            if ($rightY -gt $leftY) { $label.Row = $midRow + 1 } else { $label.Row = $midRow - $labelHeight - 1 }
        }
        $shapes.Add($label)
        $id++
    }
    return , $shapes
}

function Get-Pairs {
    #     矢印（元 → 先・ラベル）を、ノードの組ごとにまとめる（両向きの矢印は 1 本にする）
    param($Links)
    $pairs = [ordered]@{}
    foreach ($l in $Links) {
        $from = $l[0]; $to = $l[1]
        $key = @($from, $to) | Sort-Object
        $k = $key -join '|'
        if (-not $pairs.Contains($k)) { $pairs[$k] = @{ A = $from; B = $to; AtoB = (New-Object System.Collections.Generic.List[string]); BtoA = (New-Object System.Collections.Generic.List[string]); Forward = $false; Backward = $false } }
        $p = $pairs[$k]
        $list = $p.AtoB
        if ($p.A -eq $from) { $p.Forward = $true } else { $p.Backward = $true; $list = $p.BtoA }
        if ($l[2] -and -not $list.Contains($l[2])) { $list.Add($l[2]) }
    }
    return @($pairs.Values)
}

# ---------------------------------------------------------------- 図を作る


function ConvertTo-MermaidId {
    param([string] $Text)
    return ($Text -replace '[^A-Za-z0-9_]', '_')
}

function New-MermaidIds {
    #     画面・テーブルの Mermaid の ID（大文字の ID → 記号を _ にした ID）。英字で始まらないもの（日本語のテーブル名など）と、
    #     ほかと同じになったものは n1・n2…にする（図には元の ID を書く）
    param([string[]] $Ids)
    $map = @{}
    $used = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $n = 0
    foreach ($id in $Ids) {
        $m = ConvertTo-MermaidId $id
        if ($m -notmatch '^[A-Za-z]' -or $used.Contains($m)) { do { $n++; $m = 'n{0}' -f $n } while ($used.Contains($m)) }
        [void]$used.Add($m)
        $map[$id.ToUpperInvariant()] = $m
    }
    return $map
}

function ConvertTo-MermaidText {
    param([string] $Text)
    return ($Text -replace '"', '#quot;')
}

$screenEdges = @($edges | Where-Object Kind -eq '画面遷移')
$fkEdges = @($edges | Where-Object Kind -eq '外部キー')
$markdown = New-Object System.Collections.Generic.List[string]
$sheets = New-Object System.Collections.Generic.List[object]
$notes = New-Object System.Collections.Generic.List[string]

if ($screenFlow) {
    $keys = @($screens.Keys)
    $links = @($screenEdges | ForEach-Object { , @($_.From.ToUpperInvariant(), $_.To.ToUpperInvariant(), $_.Label) })
    $place = Get-Layers $keys $links
    $incoming = @{}
    foreach ($l in $links) { if ($l[0] -ne $l[1]) { $incoming[$l[1]] = $true } }
    $orphans = @($screens.Values | Where-Object { -not $_.Missing -and -not $incoming.ContainsKey($_.Id.ToUpperInvariant()) } | ForEach-Object { $_.Id })
    if ($orphans.Count -gt 0) { $notes.Add(('どこからも遷移してこない画面: {0}（最初に開く画面なら問題ありません）' -f ($orphans -join '・'))) }

    $markdown.Add('## 画面遷移図')
    $markdown.Add('')
    $markdown.Add('```mermaid')
    $markdown.Add('flowchart LR')
    $mid = New-MermaidIds @($screens.Values | ForEach-Object { $_.Id })
    foreach ($s in $screens.Values) {
        $text = $s.Id
        if ($s.Name) { $text = '{0}<br/>{1}' -f $s.Id, $s.Name }
        $markdown.Add(('    {0}["{1}"]' -f $mid[$s.Id.ToUpperInvariant()], (ConvertTo-MermaidText $text)))
    }
    foreach ($e in $screenEdges) {
        # 名前の無いイベントはラベル無しの矢印（-->|""| は Mermaid が読めない）
        $arrow = '-->'
        if ($e.Label) { $arrow = '-->|"{0}"|' -f (ConvertTo-MermaidText $e.Label) }
        $markdown.Add(('    {0} {1} {2}' -f $mid[$e.From.ToUpperInvariant()], $arrow, $mid[$e.To.ToUpperInvariant()]))
    }
    $missing = @($screens.Values | Where-Object Missing)
    if ($missing.Count -gt 0) {
        $markdown.Add('    classDef missing stroke-dasharray: 5 5,stroke:#c00000')
        $markdown.Add(('    class {0} missing' -f (@($missing | ForEach-Object { $mid[$_.Id.ToUpperInvariant()] }) -join ',')))
    }
    $markdown.Add('```')
    $markdown.Add('')

    $nodes = @($screens.Values | ForEach-Object {
            if ($_.Missing) { @{ Key = $_.Id.ToUpperInvariant(); Text = @($_.Id, '（画面一覧に無い）'); Fill = 'FFFFFF'; Line = 'C00000'; Dash = $true } }
            else { @{ Key = $_.Id.ToUpperInvariant(); Text = @($_.Id, $_.Name); Fill = 'DDEBF7'; Line = '2F5597'; Dash = $false } }
        })
    $shapes = New-DiagramShapes $nodes $place (Get-Pairs $links) 16 14 4
    $sheets.Add(@{ Name = '画面遷移図'; Columns = @(); Rows = @(); Title = ('画面遷移図（画面 {0}・遷移 {1}）' -f @($screens.Values | Where-Object { -not $_.Missing }).Count, $screenEdges.Count); Grid = @{ ColumnWidth = 1.43; RowHeight = 7.5 }; Shapes = $shapes })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($e in $screenEdges) {
        $row = @{ Values = @($e.From, $e.FromName, $e.Note, $e.To, $e.ToName, ('{0} {1} {2} 行目' -f $e.File, $e.Sheet, $e.Row)) }
        if ($e.Path) { $row.Link = @{ Column = 2; Path = $e.Path; Location = ("'{0}'!A{1}" -f $e.Sheet.Replace("'", "''"), $e.Row) } }
        $rows.Add($row)
    }
    $sheets.Add(@{ Name = '画面遷移'; Columns = @(@{ Header = '遷移元'; Width = 10 }, @{ Header = '遷移元の画面名'; Width = 18 }, @{ Header = 'イベント'; Width = 20 }, @{ Header = '遷移先'; Width = 10 }, @{ Header = '遷移先の画面名'; Width = 18 }, @{ Header = '設計書'; Width = 50 }); Rows = $rows })
}

if ($Kind -contains 'ER') {
    $keys = @($tableMap.Keys)
    # 親（参照されるテーブル）→ 子（参照するテーブル）
    $links = @($fkEdges | ForEach-Object { , @($_.To.ToUpperInvariant(), $_.From.ToUpperInvariant(), $_.Label) })
    $place = Get-DepthLayers $keys $links
    $related = @{}
    foreach ($l in $links) { $related[$l[0]] = $true; $related[$l[1]] = $true }
    $alone = @($tableMap.Values | Where-Object { -not $_.Missing -and -not $related.ContainsKey($_.Id.ToUpperInvariant()) } | ForEach-Object { $_.Id })
    if ($alone.Count -gt 0) { $notes.Add(('外部キーの関連の無いテーブル: {0}' -f ($alone -join '・'))) }

    $markdown.Add('## ER 図')
    $markdown.Add('')
    $markdown.Add('```mermaid')
    $markdown.Add('erDiagram')
    $mid = New-MermaidIds @($tableMap.Values | ForEach-Object { $_.Id })
    foreach ($e in $fkEdges) { $markdown.Add(('    {0} ||--o{{ {1} : "{2}"' -f $mid[$e.To.ToUpperInvariant()], $mid[$e.From.ToUpperInvariant()], (ConvertTo-MermaidText $e.Label))) }
    foreach ($t in $tableMap.Values) {
        # ID を変えたテーブル（日本語の名前など）は、別名（n1["顧客"]）で元の ID を書く
        $id = $mid[$t.Id.ToUpperInvariant()]
        $head = $id
        if ($id -cne $t.Id) { $head = '{0}["{1}"]' -f $id, (ConvertTo-MermaidText $t.Id) }
        $cols = @()
        if ($Columns -ne 'None') {
            $cols = @($t.Columns | Where-Object { $Columns -eq 'All' -or $null -ne $_.PkOrder -or $_.FkTable })
        }
        if ($cols.Count -eq 0) {
            if (-not $related.ContainsKey($t.Id.ToUpperInvariant()) -or $head -cne $id) { $markdown.Add(('    {0} {{' -f $head)); $markdown.Add('    }') }
            continue
        }
        $markdown.Add(('    {0} {{' -f $head))
        foreach ($c in $cols) {
            $keysText = @(); if ($null -ne $c.PkOrder) { $keysText += 'PK' }; if ($c.FkTable) { $keysText += 'FK' }
            $type = ([string]$c.BaseType -replace '[^A-Za-z0-9_]', '')
            if (-not $type) { $type = 'string' }
            $line = '        {0} {1}' -f $type, (ConvertTo-MermaidId $c.Name)
            if ($keysText.Count -gt 0) { $line += ' ' + ($keysText -join ', ') }
            if ($c.Logical) { $line += ' "{0}"' -f (ConvertTo-MermaidText $c.Logical) }
            $markdown.Add($line)
        }
        $markdown.Add('    }')
    }
    $markdown.Add('```')
    $markdown.Add('')
    $markdown.Add('| テーブル | 名前 | 参照するテーブル（外部キー） |')
    $markdown.Add('| --- | --- | --- |')
    foreach ($t in $tableMap.Values) {
        $refs = @($fkEdges | Where-Object { $_.From -eq $t.Id } | ForEach-Object { '{0}（{1}）' -f $_.To, $_.Label })
        $name = $t.Name
        if ($t.Missing) { $name = '（テーブル定義書に無い）' }
        $markdown.Add(('| {0} | {1} | {2} |' -f $t.Id, $name, ($refs -join '・')))
    }
    $markdown.Add('')

    $nodes = @($tableMap.Values | ForEach-Object {
            $t = $_
            $body = @()
            if ($Columns -ne 'None') {
                $body = @($t.Columns | Where-Object { $Columns -eq 'All' -or $null -ne $_.PkOrder -or $_.FkTable } | ForEach-Object {
                        $mark = @(); if ($null -ne $_.PkOrder) { $mark += 'PK' }; if ($_.FkTable) { $mark += 'FK' }
                        ('{0} {1}' -f ($mark -join ','), $_.Name).Trim()
                    })
            }
            if ($t.Missing) { @{ Key = $t.Id.ToUpperInvariant(); Text = @($t.Id, '（テーブル定義書に無い）'); Fill = 'FFFFFF'; Line = 'C00000'; Dash = $true } }
            else { @{ Key = $t.Id.ToUpperInvariant(); Text = @($t.Id, $t.Name); Body = $body; Fill = 'FFF2CC'; Line = 'BF9000'; Dash = $false } }
        })
    # 矢印は子から親へ（外部キーの向き）
    $childLinks = @($fkEdges | ForEach-Object { , @($_.From.ToUpperInvariant(), $_.To.ToUpperInvariant(), $_.Label) })
    $shapes = New-DiagramShapes $nodes $place (Get-Pairs $childLinks) 20 14 4
    $sheets.Add(@{ Name = 'ER図'; Columns = @(); Rows = @(); Title = ('ER 図（テーブル {0}・外部キー {1}。矢印は外部キーの参照先へ）' -f @($tableMap.Values | Where-Object { -not $_.Missing }).Count, $fkEdges.Count); Grid = @{ ColumnWidth = 1.43; RowHeight = 7.5 }; Shapes = $shapes })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($e in $fkEdges) {
        $row = @{ Values = @($e.From, $e.FromName, $e.Label, $e.To, $e.ToName, $e.Note, ('{0} {1} {2} 行目' -f $e.File, $e.Sheet, $e.Row)) }
        if ($e.Path) { $row.Link = @{ Column = 2; Path = $e.Path; Location = ("'{0}'!A{1}" -f $e.Sheet.Replace("'", "''"), $e.Row) } }
        $rows.Add($row)
    }
    $sheets.Add(@{ Name = '外部キー'; Columns = @(@{ Header = 'テーブル'; Width = 18 }, @{ Header = 'テーブル名'; Width = 16 }, @{ Header = 'カラム'; Width = 18 }, @{ Header = '参照先'; Width = 18 }, @{ Header = '参照先の名前'; Width = 16 }, @{ Header = '内容'; Width = 44 }, @{ Header = '設計書'; Width = 44 }); Rows = $rows })
}

# ---------------------------------------------------------------- 結果


$edges

if ($OutMarkdown) {
    $doc = New-Object System.Collections.Generic.List[string]
    $doc.Add('# 画面遷移図・ER 図')
    $doc.Add('')
    $doc.Add(('Export-DesignDiagram.ps1 で設計書（{0}）から作りました。Mermaid の図は、VS Code（Mermaid の拡張機能）・GitHub・Mermaid Live Editor などで見られます。' -f (@($Path | ForEach-Object { Split-Path -Leaf $_ }) -join '・')))
    $doc.Add('')
    if ($findings.Count -gt 0) {
        $doc.Add('> 設計書の確認が要るところ:')
        foreach ($f in $findings) { $doc.Add(('> - {0}' -f $f)) }
        $doc.Add('')
    }
    foreach ($line in $markdown) { $doc.Add($line) }
    $dir = Split-Path -Parent $OutMarkdown
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($OutMarkdown, (($doc -join "`n") + "`n"), (New-Object System.Text.UTF8Encoding($false)))
}
if ($OutXlsx -and $sheets.Count -gt 0) {
    try { Export-Workbook -FilePath $OutXlsx -Sheets $sheets.ToArray() }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$parts = New-Object System.Collections.Generic.List[string]
if ($screenFlow) { $parts.Add(('画面遷移図: 画面 {0}・遷移 {1}' -f @($screens.Values | Where-Object { -not $_.Missing }).Count, $screenEdges.Count)) }
if ($Kind -contains 'ER') { $parts.Add(('ER 図: テーブル {0}・外部キー {1}' -f @($tableMap.Values | Where-Object { -not $_.Missing }).Count, $fkEdges.Count)) }
Write-Host ($parts -join '、')
foreach ($f in $findings) { Write-Warning $f }
foreach ($n in $notes) { Write-Host $n }
if ($OutMarkdown) { Write-Host "Markdown（Mermaid）を保存しました: $OutMarkdown" }
if ($OutXlsx) { Write-Host "Excel の図を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書から画面遷移図（イベント定義の遷移先）と ER 図（テーブル定義書の外部キー）を作り、Mermaid（Markdown）と Excel の図形で保存します（Excel 不要）。
#
# .DESCRIPTION
#     画面遷移図: 画面一覧の画面と、画面設計書のイベント定義の「遷移先」（画面 ID か画面名）から、画面 → 画面の矢印（ラベルはイベント名）を作ります。
#     ER 図: テーブル定義書のテーブルと外部キー（参照先のテーブル・カラム）から、テーブルの関連を作ります（Export-TableDdl.ps1 と同じ見出しで読む）。
#
#     -OutMarkdown には Mermaid の図（flowchart・erDiagram）を、-OutXlsx には Excel の図形（四角と矢印。動かすと矢印が付いてくる）と一覧を保存します。
#     Excel の図は、矢印の元から先へ左から右に並べます（ER 図は、参照されるテーブルを左に）。両向きの遷移は 1 本の矢印にします。
#     遷移先が画面一覧に無い画面・参照先がテーブル定義書に無いテーブルは警告し、図には点線の四角で入れます。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（画面一覧・画面設計書・テーブル定義書。複数可）。
#
# .PARAMETER Kind
#     作る図（ScreenFlow: 画面遷移図 / ER: ER 図）。既定は両方。
#
# .PARAMETER Columns
#     ER 図に書くカラム（Key: 主キーと外部キー（既定）/ All: すべて / None: 書かない）。
#
# .PARAMETER OutMarkdown
#     Mermaid の図を Markdown（.md）に保存します。
#
# .PARAMETER OutXlsx
#     Excel の図形の図と一覧（画面遷移・外部キー）を xlsx に保存します。
#
# .PARAMETER Headers
#     テーブル定義書の見出しの候補の CSV。既定は tools\ddl\テーブル定義書の見出し.csv（Export-TableDdl.ps1 と同じ）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の設計書は読まなくなります）。
#
# .EXAMPLE
#     .\Export-DesignDiagram.ps1 D:\設計書 -OutMarkdown D:\図.md -OutXlsx D:\図.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     図の並びは目安です。Excel の図形は動かせるので、見やすく並べ直してください。
