[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Old,

    [Parameter(Mandatory = $true, Position = 1)]
    [string] $New,

    [switch] $NoRowAlign,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$maxAlignCells = 1000000   # 行の対応付けを厳密な LCS で行う「旧の行数 × 新の行数」の上限（超えたら patience diff）
$pairWindow = 50           # 一致しなかった行どうしを「変更された行」として対応付けるとき、先を探す行数

if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'CompareExcel.Diff' -DefaultDisplayPropertySet File, Location, Change, Detail -Force -WhatIf:$false

# ---------------------------------------------------------------- ブックの読み取り


function Read-Book {
    #     Search-Excel.ps1 でブックの中身（値・数式・コメント・図形・シート名）を読み、シートごとに整理する。
    #     Rows のキーは行番号の文字列（[ordered] に整数で引くと位置指定になるため）
    param([string] $FilePath)

    $searchArgs = @{
        Pattern = '[\s\S]'
        Path    = [System.Management.Automation.WildcardPattern]::Escape($FilePath)
        Target  = @('Value', 'Formula', 'Comment', 'Shape', 'SheetName')
    }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $w = $null
    $hits = @(& $searchScript @searchArgs -WarningVariable w -WarningAction SilentlyContinue 6>$null)
    if ($w) { throw (@($w | ForEach-Object { $_.Message }) -join ' / ') }

    $book = [ordered]@{}
    foreach ($h in $hits) {
        if (-not $book.Contains($h.Sheet)) {
            $book[$h.Sheet] = [pscustomobject]@{ Rows = @{}; Comments = @{}; Shapes = @{} }
        }
        $s = $book[$h.Sheet]
        switch ($h.Kind) {
            { $_ -eq '値' -or $_ -eq '数式' } {
                $m = [regex]::Match($h.Cell, '^([A-Z]+)(\d+)$')
                $row = $m.Groups[2].Value
                if (-not $s.Rows.ContainsKey($row)) { $s.Rows[$row] = @{ V = @{}; F = @{} } }
                if ($h.Kind -eq '値') { $s.Rows[$row].V[$m.Groups[1].Value] = $h.Text }
                else { $s.Rows[$row].F[$m.Groups[1].Value] = $h.Text }
            }
            'コメント' { $s.Comments[$h.Cell] = $h.Text }
            '図形' {
                $key = $h.Name
                if (-not $key) { $key = $h.Cell }
                $s.Shapes[$key] = [pscustomobject]@{ Text = $h.Text; Cell = $h.Cell }
            }
        }
    }
    return $book
}

function Get-RowSignature {
    # 行の内容を 1 つの文字列にする（同じ内容の行は同じ文字列になる）
    param($Row)
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($c in ($Row.V.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts.Add("$c`t" + $Row.V[$c]) }
    foreach ($c in ($Row.F.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts.Add("$c=`t" + $Row.F[$c]) }
    return ($parts -join "`n")
}

function Get-RowSimilarity {
    #     2 つの行がどれだけ似ているか（0〜1）。同じ列で同じ内容なら 1 点、同じ列に何か入っているだけなら 0.5 点
    param($A, $B)
    $cols = @{}
    foreach ($k in @($A.V.Keys) + @($A.F.Keys) + @($B.V.Keys) + @($B.F.Keys)) { $cols[$k] = $true }
    if ($cols.Count -eq 0) { return 1.0 }
    $score = 0.0
    foreach ($c in $cols.Keys) {
        $ia = $A.V.ContainsKey($c) -or $A.F.ContainsKey($c)
        $ib = $B.V.ContainsKey($c) -or $B.F.ContainsKey($c)
        if (-not ($ia -and $ib)) { continue }
        if ($A.V[$c] -ceq $B.V[$c] -and $A.F[$c] -ceq $B.F[$c]) { $score += 1.0 } else { $score += 0.5 }
    }
    return $score / $cols.Count
}

# ---------------------------------------------------------------- 行の対応付け


function Get-RowMapping {
    #     旧と新の行を対応付ける。返り値は Pairs（旧行, 新行）、Removed（旧だけの行）、Added（新だけの行）。
    #     1) 先頭と末尾の同じ行を対応付け  2) 残りを LCS（同じ内容の行の最長共通部分列）で対応付け
    #     3) 間に残った行は、似ている行どうしを順に「変更された行」として対応付ける
    param([string[]] $OldRows, [string[]] $NewRows, $OldBook, $NewBook)

    $pairs = New-Object System.Collections.Generic.List[object]
    $removed = New-Object System.Collections.Generic.List[string]
    $added = New-Object System.Collections.Generic.List[string]

    if ($NoRowAlign) {
        # 行番号どうしで比べる
        $all = @{}
        foreach ($r in $OldRows) { $all[$r] = $true }
        foreach ($r in $NewRows) { $all[$r] = $true }
        foreach ($r in $all.Keys) {
            $inOld = $OldBook.Rows.ContainsKey($r); $inNew = $NewBook.Rows.ContainsKey($r)
            if ($inOld -and $inNew) { $pairs.Add(@($r, $r)) } elseif ($inOld) { $removed.Add($r) } else { $added.Add($r) }
        }
        return [pscustomobject]@{ Pairs = $pairs; Removed = $removed; Added = $added }
    }

    $a = @($OldRows | ForEach-Object { Get-RowSignature $OldBook.Rows[$_] })
    $b = @($NewRows | ForEach-Object { Get-RowSignature $NewBook.Rows[$_] })
    $n = $a.Count; $m = $b.Count

    $head = 0
    while ($head -lt $n -and $head -lt $m -and $a[$head] -ceq $b[$head]) { $pairs.Add(@($OldRows[$head], $NewRows[$head])); $head++ }
    $tail = 0
    while ($tail -lt ($n - $head) -and $tail -lt ($m - $head) -and $a[$n - 1 - $tail] -ceq $b[$m - 1 - $tail]) { $tail++ }

    $oi = $head; $oEnd = $n - $tail
    $ni = $head; $nEnd = $m - $tail
    $matches = New-Object System.Collections.Generic.List[object]
    $on = $oEnd - $oi; $nn = $nEnd - $ni
    if ($on -gt 0 -and $nn -gt 0) {
        if ($on * $nn -le $maxAlignCells) {
            # LCS の表を作り、後ろからたどって一致する行を拾う
            $dp = New-Object 'int[,]' ($on + 1), ($nn + 1)
            for ($i = $on - 1; $i -ge 0; $i--) {
                for ($j = $nn - 1; $j -ge 0; $j--) {
                    if ($a[$oi + $i] -ceq $b[$ni + $j]) { $dp[$i, $j] = $dp[($i + 1), ($j + 1)] + 1 }
                    elseif ($dp[($i + 1), $j] -ge $dp[$i, ($j + 1)]) { $dp[$i, $j] = $dp[($i + 1), $j] }
                    else { $dp[$i, $j] = $dp[$i, ($j + 1)] }
                }
            }
            $i = 0; $j = 0
            while ($i -lt $on -and $j -lt $nn) {
                if ($a[$oi + $i] -ceq $b[$ni + $j]) { $matches.Add(@(($oi + $i), ($ni + $j))); $i++; $j++ }
                elseif ($dp[($i + 1), $j] -ge $dp[$i, ($j + 1)]) { $i++ }
                else { $j++ }
            }
        }
        else {
            # 行数が多いときは patience diff: 旧と新のどちらでも 1 回しか出てこない行を目印にし、
            # 順序が保たれる最長の組（最長増加部分列）を一致とする
            $countA = @{}; $countB = @{}; $posB = @{}
            for ($i = $oi; $i -lt $oEnd; $i++) { $countA[$a[$i]] = 1 + [int]$countA[$a[$i]] }
            for ($j = $ni; $j -lt $nEnd; $j++) { $countB[$b[$j]] = 1 + [int]$countB[$b[$j]]; $posB[$b[$j]] = $j }
            $cand = New-Object System.Collections.Generic.List[object]
            for ($i = $oi; $i -lt $oEnd; $i++) {
                if ($countA[$a[$i]] -eq 1 -and $countB[$a[$i]] -eq 1) { $cand.Add(@($i, $posB[$a[$i]])) }
            }
            # 最長増加部分列（新側の位置が増えていく最長の並び）
            $tails = New-Object System.Collections.Generic.List[int]
            $prev = New-Object int[] $cand.Count
            for ($k = 0; $k -lt $cand.Count; $k++) {
                $lo = 0; $hi = $tails.Count
                while ($lo -lt $hi) { $mid = [int][math]::Floor(($lo + $hi) / 2); if ($cand[$tails[$mid]][1] -lt $cand[$k][1]) { $lo = $mid + 1 } else { $hi = $mid } }
                if ($lo -gt 0) { $prev[$k] = $tails[$lo - 1] } else { $prev[$k] = -1 }
                if ($lo -eq $tails.Count) { $tails.Add($k) } else { $tails[$lo] = $k }
            }
            $chain = New-Object System.Collections.Generic.List[object]
            if ($tails.Count -gt 0) { $k = $tails[$tails.Count - 1]; while ($k -ge 0) { $chain.Insert(0, $cand[$k]); $k = $prev[$k] } }
            foreach ($c in $chain) { $matches.Add($c) }
            Write-Verbose ("行数が多いため patience diff で対応付けました（{0} 行 × {1} 行、目印 {2} 行）" -f $on, $nn, $chain.Count)
        }
    }
    # 番兵として末尾を一致扱いにし、一致と一致の間（ギャップ）を処理する
    $matches.Add(@($oEnd, $nEnd))

    $po = $oi; $pn = $ni
    foreach ($mt in $matches) {
        $gapOld = @(); if ($mt[0] -gt $po) { $gapOld = @($po..($mt[0] - 1)) }
        $gapNew = @(); if ($mt[1] -gt $pn) { $gapNew = @($pn..($mt[1] - 1)) }
        $usedNew = @{}
        $next = 0
        foreach ($x in $gapOld) {
            $best = -1; $bestScore = 0.5
            $limit = [math]::Min($gapNew.Count, $next + $pairWindow)
            for ($k = $next; $k -lt $limit; $k++) {
                $score = Get-RowSimilarity $OldBook.Rows[$OldRows[$x]] $NewBook.Rows[$NewRows[$gapNew[$k]]]
                if ($score -ge $bestScore -and ($best -lt 0 -or $score -gt $bestScore)) { $best = $k; $bestScore = $score }
            }
            if ($best -ge 0) {
                $pairs.Add(@($OldRows[$x], $NewRows[$gapNew[$best]]))
                for ($k = $next; $k -lt $best; $k++) { $added.Add($NewRows[$gapNew[$k]]); $usedNew[$k] = $true }
                $usedNew[$best] = $true
                $next = $best + 1
            }
            else { $removed.Add($OldRows[$x]) }
        }
        for ($k = $next; $k -lt $gapNew.Count; $k++) { if (-not $usedNew.ContainsKey($k)) { $added.Add($NewRows[$gapNew[$k]]) } }
        if ($mt[0] -lt $oEnd) { $pairs.Add(@($OldRows[$mt[0]], $NewRows[$mt[1]])) }
        $po = $mt[0] + 1; $pn = $mt[1] + 1
    }
    for ($t = 0; $t -lt $tail; $t++) { $pairs.Add(@($OldRows[$oEnd + $t], $NewRows[$nEnd + $t])) }

    return [pscustomobject]@{ Pairs = $pairs; Removed = $removed; Added = $added }
}

# ---------------------------------------------------------------- 差分の出力


$script:Diffs = New-Object System.Collections.Generic.List[object]

function Add-Diff {
    param([string] $Sheet, [string] $NewCell, [string] $OldCell, [string] $Kind, [string] $Change, [string] $OldText, [string] $NewText)

    $cell = $NewCell
    if (-not $cell) { $cell = $OldCell }
    $location = $Sheet
    if ($cell) { $location = '{0}!{1}' -f $Sheet, $cell }
    $oldLocation = ''
    if ($OldCell -and $OldCell -ne $NewCell) { $oldLocation = '{0}!{1}' -f $Sheet, $OldCell }

    switch ($Change) {
        '変更' { $detail = '{0}: {1} → {2}' -f $Kind, (Get-Short $OldText), (Get-Short $NewText) }
        '追加' { $detail = '{0}を追加: {1}' -f $Kind, (Get-Short $NewText) }
        '削除' { $detail = '{0}を削除: {1}' -f $Kind, (Get-Short $OldText) }
    }
    if ($oldLocation -and $NewCell) { $detail += "（旧 $OldCell）" }

    $d = [pscustomobject]@{
        File        = $script:CurrentFile
        Location    = $location
        Sheet       = $Sheet
        Cell        = $cell
        OldLocation = $oldLocation
        Kind        = $Kind
        Change      = $Change
        Old         = $OldText
        New         = $NewText
        Detail      = $detail
        Path        = $script:CurrentNewPath
        OldPath     = $script:CurrentOldPath
    }
    $d.PSObject.TypeNames.Insert(0, 'CompareExcel.Diff')
    $script:Diffs.Add($d)
}

function Get-Short {
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    if ($Text.Length -gt 60) { return $Text.Substring(0, 60) + '…' }
    return $Text
}

function Format-RowContent {
    param($Row)
    $parts = @()
    foreach ($c in ($Row.V.Keys | Sort-Object { (Get-ColumnIndex $_) })) { $parts += ('{0}:{1}' -f $c, $Row.V[$c]) }
    foreach ($c in ($Row.F.Keys | Sort-Object { (Get-ColumnIndex $_) })) { if (-not $Row.V.ContainsKey($c)) { $parts += ('{0}:{1}' -f $c, $Row.F[$c]) } }
    return ($parts -join ' | ')
}

function Compare-Sheet {
    param([string] $Sheet, $OldSheet, $NewSheet)

    $oldRows = @($OldSheet.Rows.Keys | Sort-Object { [int]$_ })
    $newRows = @($NewSheet.Rows.Keys | Sort-Object { [int]$_ })
    $map = Get-RowMapping -OldRows $oldRows -NewRows $newRows -OldBook $OldSheet -NewBook $NewSheet

    $rowMap = @{}
    foreach ($p in $map.Pairs) {
        $rowMap[$p[0]] = $p[1]
        $o = $OldSheet.Rows[$p[0]]; $n = $NewSheet.Rows[$p[1]]
        foreach ($kind in @(@('V', '値'), @('F', '数式'))) {
            $cols = @{}
            foreach ($c in @($o[$kind[0]].Keys) + @($n[$kind[0]].Keys)) { $cols[$c] = $true }
            foreach ($c in ($cols.Keys | Sort-Object { (Get-ColumnIndex $_) })) {
                $ov = $o[$kind[0]][$c]; $nv = $n[$kind[0]][$c]
                if ($ov -ceq $nv) { continue }
                $newCell = "$c$($p[1])"; $oldCell = "$c$($p[0])"
                if ($null -eq $ov) { Add-Diff $Sheet $newCell $oldCell $kind[1] '追加' $null $nv }
                elseif ($null -eq $nv) { Add-Diff $Sheet $newCell $oldCell $kind[1] '削除' $ov $null }
                else { Add-Diff $Sheet $newCell $oldCell $kind[1] '変更' $ov $nv }
            }
        }
    }
    foreach ($r in ($map.Added | Sort-Object { [int]$_ })) {
        Add-Diff $Sheet "A$r" $null '行' '追加' $null (Format-RowContent $NewSheet.Rows[$r])
    }
    foreach ($r in ($map.Removed | Sort-Object { [int]$_ })) {
        Add-Diff $Sheet $null "A$r" '行' '削除' (Format-RowContent $OldSheet.Rows[$r]) $null
    }

    # コメント: 旧の位置を行の対応で新の位置に読み替えて比べる
    $oldComments = @{}
    foreach ($cell in $OldSheet.Comments.Keys) {
        $m = [regex]::Match($cell, '^([A-Z]+)(\d+)$')
        $mapped = $cell
        if ($rowMap.ContainsKey($m.Groups[2].Value)) { $mapped = $m.Groups[1].Value + $rowMap[$m.Groups[2].Value] }
        $oldComments[$mapped] = @($cell, $OldSheet.Comments[$cell])
    }
    foreach ($cell in $NewSheet.Comments.Keys) {
        if ($oldComments.ContainsKey($cell)) {
            if ($oldComments[$cell][1] -cne $NewSheet.Comments[$cell]) { Add-Diff $Sheet $cell $oldComments[$cell][0] 'コメント' '変更' $oldComments[$cell][1] $NewSheet.Comments[$cell] }
            $oldComments.Remove($cell)
        }
        else { Add-Diff $Sheet $cell $null 'コメント' '追加' $null $NewSheet.Comments[$cell] }
    }
    foreach ($k in $oldComments.Keys) { Add-Diff $Sheet $null $oldComments[$k][0] 'コメント' '削除' $oldComments[$k][1] $null }

    # 図形: 名前で対応付ける
    foreach ($name in $NewSheet.Shapes.Keys) {
        $ns = $NewSheet.Shapes[$name]
        if ($OldSheet.Shapes.ContainsKey($name)) {
            $os = $OldSheet.Shapes[$name]
            if ($os.Text -cne $ns.Text) { Add-Diff $Sheet $ns.Cell $os.Cell '図形' '変更' $os.Text $ns.Text }
        }
        else { Add-Diff $Sheet $ns.Cell $null '図形' '追加' $null $ns.Text }
    }
    foreach ($name in $OldSheet.Shapes.Keys) {
        if (-not $NewSheet.Shapes.ContainsKey($name)) { Add-Diff $Sheet $null $OldSheet.Shapes[$name].Cell '図形' '削除' $OldSheet.Shapes[$name].Text $null }
    }
}

function Compare-Book {
    param([string] $OldPath, [string] $NewPath)

    $oldBook = Read-Book $OldPath
    $newBook = Read-Book $NewPath
    foreach ($sheet in $newBook.Keys) {
        if ($oldBook.Contains($sheet)) { Compare-Sheet -Sheet $sheet -OldSheet $oldBook[$sheet] -NewSheet $newBook[$sheet] }
        else { Add-Diff $sheet $null $null 'シート' '追加' $null ("{0} 行" -f $newBook[$sheet].Rows.Count) }
    }
    foreach ($sheet in $oldBook.Keys) {
        if (-not $newBook.Contains($sheet)) { Add-Diff $sheet $null $null 'シート' '削除' ("{0} 行" -f $oldBook[$sheet].Rows.Count) $null }
    }
}

# ---------------------------------------------------------------- 比べるファイルの組み合わせ


function Get-ExcelFiles {
    param([string] $Dir)
    $map = [ordered]@{}
    $root = (Resolve-Path -LiteralPath $Dir).ProviderPath.TrimEnd('\', '/')
    foreach ($f in (Get-ChildItem -LiteralPath $root -File -Recurse:(-not $NoRecurse) | Sort-Object FullName)) {
        if (@('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb') -notcontains $f.Extension.ToLowerInvariant()) { continue }
        if ($f.Name.StartsWith('~$')) { continue }
        if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
        $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
        $map[$rel.ToLowerInvariant()] = [pscustomobject]@{ Relative = $rel; Path = $f.FullName }
    }
    return $map
}

if (-not (Test-Path -LiteralPath $Old)) { throw "見つかりません: $Old" }
if (-not (Test-Path -LiteralPath $New)) { throw "見つかりません: $New" }
$oldItem = Get-Item -LiteralPath $Old
$newItem = Get-Item -LiteralPath $New
if ($oldItem.PSIsContainer -ne $newItem.PSIsContainer) { throw '-Old と -New は、両方ともファイルか、両方ともフォルダにしてください。' }

$pairs = New-Object System.Collections.Generic.List[object]
if ($oldItem.PSIsContainer) {
    $oldFiles = Get-ExcelFiles $oldItem.FullName
    $newFiles = Get-ExcelFiles $newItem.FullName
    foreach ($k in $newFiles.Keys) {
        $o = $null
        if ($oldFiles.Contains($k)) { $o = $oldFiles[$k].Path }
        $pairs.Add([pscustomobject]@{ Relative = $newFiles[$k].Relative; Old = $o; New = $newFiles[$k].Path })
    }
    foreach ($k in $oldFiles.Keys) {
        if (-not $newFiles.Contains($k)) { $pairs.Add([pscustomobject]@{ Relative = $oldFiles[$k].Relative; Old = $oldFiles[$k].Path; New = $null }) }
    }
}
else {
    $pairs.Add([pscustomobject]@{ Relative = $newItem.Name; Old = $oldItem.FullName; New = $newItem.FullName })
}

$summary = New-Object System.Collections.Generic.List[object]
$index = 0
foreach ($p in $pairs) {
    $index++
    Write-Progress -Activity 'Excel を比較中' -Status ("{0}/{1}  {2}" -f $index, $pairs.Count, $p.Relative) -PercentComplete ([int](100 * $index / $pairs.Count))
    $script:CurrentFile = $p.Relative
    $script:CurrentNewPath = $p.New
    $script:CurrentOldPath = $p.Old
    $before = $script:Diffs.Count
    $state = ''
    try {
        if (-not $p.Old) { Add-Diff '' $null $null 'ファイル' '追加' $null $p.Relative; $state = '追加' }
        elseif (-not $p.New) { Add-Diff '' $null $null 'ファイル' '削除' $p.Relative $null; $state = '削除' }
        else {
            Compare-Book -OldPath $p.Old -NewPath $p.New
            if ($script:Diffs.Count -eq $before) { $state = '同じ' } else { $state = '差分あり' }
        }
    }
    catch {
        $state = '失敗'
        Write-Warning ("{0}: 比較できませんでした: {1}" -f $p.Relative, $_.Exception.Message)
    }
    $summary.Add([pscustomobject]@{ File = $p.Relative; State = $state; Count = $script:Diffs.Count - $before; Old = $p.Old; New = $p.New })
}
Write-Progress -Activity 'Excel を比較中' -Completed

$script:Diffs

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($d in $script:Diffs) {
        $row = @{ Values = @($d.File, $d.Sheet, $d.Location, $d.OldLocation, $d.Kind, $d.Change, $d.Old, $d.New) }
        # 削除は旧のファイル、それ以外は新のファイルの該当セルを開く
        $target = $d.Path; $cell = $d.Cell
        if ($d.Change -eq '削除') {
            $target = $d.OldPath
            if ($d.OldLocation) { $cell = $d.OldLocation.Substring($d.OldLocation.LastIndexOf('!') + 1) }
        }
        if ($target) {
            $loc = ''
            if ($d.Sheet -and $d.Kind -ne 'シート' -and $cell) { $loc = "'{0}'!{1}" -f $d.Sheet.Replace("'", "''"), $cell }
            $row.Link = @{ Column = 2; Path = $target; Location = $loc }
        }
        $rows.Add($row)
    }
    $fileRows = New-Object System.Collections.Generic.List[object]
    foreach ($s in $summary) {
        $row = @{ Values = @($s.File, $s.State, $s.Count, $s.Old, $s.New) }
        $t = $s.New; if (-not $t) { $t = $s.Old }
        if ($t) { $row.Link = @{ Column = 0; Path = $t; Location = '' } }
        $fileRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{
                Name = '差分一覧'
                Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = 'シート'; Width = 16 }, @{ Header = '場所'; Width = 20 }, @{ Header = '旧の場所'; Width = 16 },
                    @{ Header = '種類'; Width = 9 }, @{ Header = '変更'; Width = 7 }, @{ Header = '変更前'; Width = 45 }, @{ Header = '変更後'; Width = 45 })
                Rows = $rows
            },
            @{
                Name = 'ファイル別'
                Columns = @(@{ Header = 'ファイル'; Width = 45 }, @{ Header = '結果'; Width = 10 }, @{ Header = '差分'; Width = 8 }, @{ Header = '旧'; Width = 60 }, @{ Header = '新'; Width = 60 })
                Rows = $fileRows
            })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$states = $summary | Group-Object State | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
Write-Host ("{0} ファイル: {1} / 差分 {2} 件" -f $summary.Count, ($states -join ' / '), $script:Diffs.Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     2 つの Excel ファイル（またはフォルダ）を比べ、セル単位の差分を一覧にします（Excel 不要）。
#
# .DESCRIPTION
#     値・数式・コメント・図形の文字・シートの追加削除を比べます。
#     行を挿入・削除した場合でも、それより下のセルがすべて「変更」にならないよう、
#     同じ内容の行どうしを対応付けてから比べます（-NoRowAlign で無効にできます）。
#
#     フォルダどうしを指定すると、同じ相対パスのファイルを比べ、片方にしかないファイルは追加・削除として出力します。
#     結果は 1 差分 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#
# .PARAMETER Old
#     比較元（変更前）のファイルまたはフォルダ。
#
# .PARAMETER New
#     比較先（変更後）のファイルまたはフォルダ。
#
# .PARAMETER NoRowAlign
#     行の対応付けを行わず、同じセル番地どうしを比べます。
#
# .PARAMETER NoRecurse
#     フォルダを比べるとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は比較できなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（差分一覧・ファイル別。場所をクリックすると該当セルが開きます）。
#
# .EXAMPLE
#     .\Compare-Excel.ps1 .\v1.0\テーブル定義書.xlsx .\v1.1\テーブル定義書.xlsx
#
# .EXAMPLE
#     # 版のフォルダどうしを比べて保存する
#     .\Compare-Excel.ps1 D:\納品\v1.0 D:\納品\v1.1 -OutXlsx D:\差分_v1.0-v1.1.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     列の挿入・削除は対応付けません（列がずれた場合はセルの変更として出力されます）。
