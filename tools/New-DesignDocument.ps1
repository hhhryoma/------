[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Create')]
param(
    [Parameter(ParameterSetName = 'Create', Mandatory = $true)]
    [string] $Template,

    [Parameter(ParameterSetName = 'Create', Mandatory = $true)]
    [string] $List,

    [Parameter(ParameterSetName = 'Create')]
    [string] $ListSheet,

    [Parameter(ParameterSetName = 'Create')]
    [string] $FileName,

    [Parameter(ParameterSetName = 'Create')]
    [string] $OutDir = '.',

    [Parameter(ParameterSetName = 'Bump', Mandatory = $true)]
    [switch] $Bump,

    [Parameter(ParameterSetName = 'Bump', Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Description,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $CompareWith,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Section = '全体',

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Version,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $HistorySheet = '改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録',

    [Parameter(ParameterSetName = 'Bump')]
    [switch] $NoRecurse,

    [string] $Author = [System.Environment]::UserName,

    [datetime] $Date = (Get-Date).Date
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$replaceScript = Join-Path $PSScriptRoot 'Replace-Excel.ps1'
$compareScript = Join-Path $PSScriptRoot 'Compare-Excel.ps1'

Update-TypeData -TypeName 'DesignDocument.Result' -DefaultDisplayPropertySet File, Status, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$dateText = $Date.ToString('yyyy/MM/dd', $invariant)

function New-Result {
    param([string] $File, [string] $Status, [string] $Detail, [string] $FilePath)
    $r = [pscustomobject]@{ File = $File; Status = $Status; Detail = $Detail; Path = $FilePath }
    $r.PSObject.TypeNames.Insert(0, 'DesignDocument.Result')
    return $r
}

# ================================================================ 雛形から作る（-Template / -List）


function Expand-Placeholder {
    #     {{名前}} を値に置き換える
    param([string] $Text, [hashtable] $Values)
    return [regex]::Replace($Text, '\{\{([^{}]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $k = $m.Groups[1].Value.Trim()
        if ($Values.ContainsKey($k)) { return [string]$Values[$k] }
        return $m.Value
    })
}

if ($PSCmdlet.ParameterSetName -eq 'Create') {
    $Template = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Template)
    $List = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($List)
    $OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
    if (-not (Test-Path -LiteralPath $Template)) { throw "雛形が見つかりません: $Template" }
    if (@('.xlsx', '.xlsm') -notcontains [System.IO.Path]::GetExtension($Template).ToLowerInvariant()) { throw "雛形は .xlsx / .xlsm にしてください: $Template" }
    if (-not (Test-Path -LiteralPath $List)) { throw "一覧が見つかりません: $List" }
    if (-not $FileName) { $FileName = [System.IO.Path]::GetFileName($Template) }

    # 雛形とファイル名の差し込み欄
    $tplHits = @(& $searchScript -Pattern '\{\{[^{}]+\}\}' -Path ([System.Management.Automation.WildcardPattern]::Escape($Template)) -Target Value, Comment, Shape -NoExcel 6>$null)
    $inTemplate = New-Object System.Collections.Generic.List[string]
    foreach ($h in $tplHits) {
        foreach ($m in [regex]::Matches([string]$h.Text, '\{\{([^{}]+)\}\}')) { $n = $m.Groups[1].Value.Trim(); if (-not $inTemplate.Contains($n)) { $inTemplate.Add($n) } }
    }
    $inName = @([regex]::Matches($FileName, '\{\{([^{}]+)\}\}') | ForEach-Object { $_.Groups[1].Value.Trim() } | Select-Object -Unique)
    if ($inName.Count -eq 0) { throw 'ファイル名に {{画面ID}} のような差し込み欄を入れてください（-FileName）。同じ名前のファイルになってしまうためです' }
    $builtin = @{ '作成日' = $dateText; '日付' = $dateText; '作成者' = $Author }
    $listNames = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($inName) + @($inTemplate)) { if (-not $builtin.ContainsKey($n) -and -not $listNames.Contains($n)) { $listNames.Add($n) } }

    # 一覧（CSV、または画面一覧などの Excel）
    $rows = New-Object System.Collections.Generic.List[object]
    $ext = [System.IO.Path]::GetExtension($List).ToLowerInvariant()
    if ($ext -eq '.csv') {
        foreach ($r in (Read-CsvFile $List $listNames.ToArray() '一覧')) {
            $v = @{}
            foreach ($n in $listNames) { $v[$n] = ([string]$r.$n).Trim() }
            $rows.Add([pscustomobject]@{ Values = $v; Where = '' })
        }
    }
    else {
        $searchArgs = @{ Pattern = '\S'; Path = [System.Management.Automation.WildcardPattern]::Escape($List); Target = 'Value' }
        if ($ListSheet) { $searchArgs['SheetNamePattern'] = $ListSheet }
        $books = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
        $labels = @($listNames | ForEach-Object { ConvertTo-Label $_ })
        foreach ($book in $books.Values) {
            foreach ($sheetName in $book.Sheets.Keys) {
                $grid = $book.Sheets[$sheetName]
                foreach ($t in (Find-DesignTables $grid $labels $labels[0])) {
                    foreach ($dr in $t.Rows) {
                        $v = @{}
                        for ($i = 0; $i -lt $listNames.Count; $i++) { $v[$listNames[$i]] = ([string]$dr.Cells[$t.ColumnOf[$labels[$i]]]).Trim() }
                        $rows.Add([pscustomobject]@{ Values = $v; Where = ('{0}!{1}{2}' -f $sheetName, (ConvertTo-ColumnName $t.ColumnOf[$labels[0]]), $dr.Row) })
                    }
                }
            }
        }
        if ($rows.Count -eq 0) { throw ("一覧に、見出し（{0}）がそろった表が見つかりません: {1}" -f ($listNames -join '・'), $List) }
    }

    if (-not $WhatIfPreference -and -not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
    $invalid = [regex]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())
    $created = 0
    foreach ($row in $rows) {
        $values = @{}
        foreach ($k in $builtin.Keys) { $values[$k] = $builtin[$k] }
        foreach ($k in $row.Values.Keys) { $values[$k] = $row.Values[$k] }
        $empty = @($inName | Where-Object { -not $values[$_] })
        if ($empty.Count -gt 0) {
            New-Result '' 'スキップ' ('ファイル名の差し込み欄（{0}）の値が空です（{1}）' -f ($empty -join '・'), $row.Where) ''
            continue
        }
        $outName = [regex]::Replace((Expand-Placeholder $FileName $values), "[$invalid]", '_')
        $target = Join-Path $OutDir $outName
        if (Test-Path -LiteralPath $target) { New-Result $outName '既存' '同じ名前のファイルがあるため作りません' $target; continue }
        if (-not $PSCmdlet.ShouldProcess($outName, '雛形から作成')) { continue }

        Copy-Item -LiteralPath $Template -Destination $target
        if ($inTemplate.Count -gt 0) {
            # 差し込み欄の置き換えは Replace-Excel.ps1 で行う（セル・図形・コメント。セルの書式や部分ごとの書式は保たれる）
            $csv = Join-Path ([System.IO.Path]::GetTempPath()) ('placeholders-' + [guid]::NewGuid().ToString('N') + '.csv')
            $lines = New-Object System.Collections.Generic.List[string]
            $lines.Add('置換前,置換後,正規表現')
            foreach ($n in $inTemplate) { $lines.Add(('"{0}","{1}",' -f ('{{' + $n + '}}').Replace('"', '""'), ([string]$values[$n]).Replace('"', '""'))) }
            [System.IO.File]::WriteAllLines($csv, $lines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
            try { $null = & $replaceScript -ListPath $csv -Path ([System.Management.Automation.WildcardPattern]::Escape($target)) -Force -Confirm:$false -WhatIf:$false 6>$null }
            finally { Remove-Item -LiteralPath $csv -Force -ErrorAction SilentlyContinue }
        }
        $created++
        New-Result $outName '作成' '' $target
    }
    Write-Host ''
    Write-Host ('{0} ファイルを作りました（一覧の {1} 行）: {2}' -f $created, $rows.Count, $OutDir)
    return
}

# ================================================================ 版を上げる（-Bump）


function ConvertTo-VersionText {
    #     版数の文字（「第1.1版」「Ver.1.10」など）を「1.1」の形にする（無ければ $null）
    param([string] $Text)
    if (-not $Text) { return $null }
    $m = [regex]::Match($Text.Normalize($nfkc), '(\d+(?:\.\d+)*)')
    if (-not $m.Success) { return $null }
    return (@($m.Groups[1].Value.Split('.') | ForEach-Object { [string][int64]$_ }) -join '.')
}

function Get-NextVersion {
    #     最後の数を 1 つ上げる（1.0 → 1.1、1.9 → 1.10）
    param([string] $Current)
    $parts = @($Current.Split('.'))
    if ($parts.Count -eq 1) { $parts += '0' }
    $parts[$parts.Count - 1] = [string]([int64]$parts[$parts.Count - 1] + 1)
    return ($parts -join '.')
}

function Set-VersionText {
    #     「第1.0版」の数の部分だけを新しい版数にする
    param([string] $Text, [string] $NewVersion)
    $t = $Text.Normalize($nfkc)
    $m = [regex]::Match($t, '\d+(?:\.\d+)*')
    if (-not $m.Success) { return $NewVersion }
    return $t.Substring(0, $m.Index) + $NewVersion + $t.Substring($m.Index + $m.Length)
}

$versionLabel = '^(版数|版|改訂版数|バージョン|Version|Ver\.?|Rev\.?|改版番号)$'
$coverDateLabel = '^(最終更新日|更新日|改訂日|改版日|最終改訂日)$'
$histDateLabel = '^(日付|改訂日|改版日|更新日|年月日|変更日)$'
$sectionLabel = '^(改訂箇所|改版箇所|変更箇所|改訂対象|対象|箇所|変更対象)$'
$descLabel = '^(改訂内容|改版内容|変更内容|内容|概要|改訂概要|変更概要)$'
$authorLabel = '^(改訂者|改版者|変更者|更新者|担当者|担当|作成者)$'

function Get-DateValue {
    #     日付のセルに書く値（もとが数値なら Excel のシリアル値、文字ならもとの書き方に合わせた文字）
    param($Cell, [string] $OldText)
    if (Test-NumericCellElement $Cell) { return ConvertTo-ExcelDate $Date }
    $t = ([string]$OldText).Normalize($nfkc)
    if ($t -match '\d{4}年') { return $Date.ToString('yyyy年M月d日', $invariant) }
    if ($t -match '\d{4}-\d') { return $Date.ToString('yyyy-MM-dd', $invariant) }
    return $dateText
}

function Find-LabelPairs {
    #     見出し（正規表現）に一致するセルと、その右の値のセルの組
    param($Grid, [string] $Pattern)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Grid.Rows) {
        $cols = $Grid.Cols[$row]
        for ($i = 0; $i -lt $cols.Count - 1; $i++) {
            if ($Grid.Labels[$row][$cols[$i]] -match $Pattern) { $list.Add([pscustomobject]@{ Row = $row; Col = $cols[$i + 1]; Text = $Grid.Cells[$row][$cols[$i + 1]] }) }
        }
    }
    return , $list
}

if (-not $Description -and -not $CompareWith) { throw '-Bump では、改訂内容（-Description）か、比べる前の版（-CompareWith）を指定してください。' }

# 改訂履歴に一度に足す行の上限（超えた分は「ほか n 件」の 1 行にまとめる）
$maxHistoryRows = 20

function Get-HistoryDrafts {
    #     前の版と比べ、改訂履歴に書く行（改訂箇所・改訂内容）の下書きを求める（Compare-Excel.ps1 -OutHistory）
    param([string] $OldPath, [string] $NewPath)
    $csv = Join-Path ([System.IO.Path]::GetTempPath()) ('history-' + [guid]::NewGuid().ToString('N') + '.csv')
    try {
        $null = & $compareScript -Old $OldPath -New $NewPath -OutHistory $csv -HistorySheet $HistorySheet -NoExcel -WarningAction SilentlyContinue 6>$null
        if (-not (Test-Path -LiteralPath $csv)) { return @() }
        return @(Read-TextFile $csv | ConvertFrom-Csv | ForEach-Object { [pscustomobject]@{ Section = $_.'改訂箇所'; Description = $_.'改訂内容' } })
    }
    finally { Remove-Item -LiteralPath $csv -Force -ErrorAction SilentlyContinue }
}

$files = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) { $found = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) } else { $found = @($item) }
        foreach ($f in $found) { if (@('.xlsx', '.xlsm') -contains $f.Extension.ToLowerInvariant() -and -not $f.Name.StartsWith('~$')) { $files.Add($f) } }
    }
}

# 前の版: ファイルならそのファイル（版を上げるファイルが 1 つのとき）、フォルダなら版数を除いた名前が同じファイル
$oldByName = $null
if ($CompareWith) {
    $cw = Get-Item -LiteralPath $CompareWith -ErrorAction SilentlyContinue
    if (-not $cw) { throw "-CompareWith が見つかりません: $CompareWith" }
    if ($cw.PSIsContainer) {
        $oldByName = @{}
        foreach ($o in (Get-ChildItem -LiteralPath $cw.FullName -File -Recurse | Where-Object { @('.xlsx', '.xlsm') -contains $_.Extension.ToLowerInvariant() -and -not $_.Name.StartsWith('~$') })) {
            $k = Get-VersionlessName $o.Name
            if ($oldByName.ContainsKey($k)) { $oldByName[$k] = $null } else { $oldByName[$k] = $o.FullName }
        }
    }
    elseif ($files.Count -gt 1) { throw '-CompareWith にファイルを指定できるのは、版を上げるファイルが 1 つのときだけです（複数のときは前の版のフォルダを指定してください）。' }
}

foreach ($f in $files) {
    $oldPath = $null
    if ($CompareWith) {
        if ($oldByName) { $oldPath = $oldByName[(Get-VersionlessName $f.Name)] } else { $oldPath = $cw.FullName }
        if (-not $oldPath) { New-Result $f.Name 'スキップ' ('前の版のファイルが -CompareWith のフォルダに見つかりません（版数を除いた名前が同じファイルが 1 つだけあるものを使います）') $f.FullName; continue }
    }
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($f.FullName)) -Target Value -NoExcel 6>$null)
    $books = ConvertTo-DesignBooks $hits
    if ($books.Count -eq 0) { New-Result $f.Name 'スキップ' '読めませんでした' $f.FullName; continue }
    $book = @($books.Values)[0]

    $pkg = Open-ZipPackage $f.FullName
    try {
        $sheetParts = (Get-WorkbookSheetParts $pkg).Sheets
        $coverName = @($sheetParts.Keys | Where-Object { $_ -match '^\s*表紙\s*$' })[0]
        if (-not $coverName) { $coverName = @($sheetParts.Keys)[0] }
        $histName = @($sheetParts.Keys | Where-Object { $_ -match $HistorySheet })[0]

        # ---- 表紙の版数・更新日
        $coverVersion = $null; $coverDate = $null
        if ($book.Sheets.Contains($coverName)) {
            $grid = $book.Sheets[$coverName]
            $pairs = Find-LabelPairs $grid $versionLabel
            foreach ($pair in $pairs) { if (ConvertTo-VersionText $pair.Text) { $coverVersion = $pair; break } }
            if (-not $coverVersion) {
                foreach ($row in $grid.Rows) { foreach ($col in $grid.Cols[$row]) {
                    if (-not $coverVersion -and $grid.Labels[$row][$col] -match '^(第\d+(\.\d+)*版|(Ver|Version|Rev)\.?\d+(\.\d+)*)$') { $coverVersion = [pscustomobject]@{ Row = $row; Col = $col; Text = $grid.Cells[$row][$col] } }
                } }
            }
            $pairs = Find-LabelPairs $grid $coverDateLabel
            if ($pairs.Count -gt 0) { $coverDate = $pairs[0] }
        }

        # ---- 改訂履歴: 版数と日付の見出しがある行のうち、下に版数の行がいちばん多いもの（Test-ExcelDeliverable.ps1 と同じ）
        $hist = $null
        if ($histName -and $book.Sheets.Contains($histName)) {
            $grid = $book.Sheets[$histName]
            foreach ($row in $grid.Rows) {
                $cols = @{}
                foreach ($col in $grid.Cols[$row]) {
                    $l = $grid.Labels[$row][$col]
                    if (-not $cols['Version'] -and $l -match $versionLabel) { $cols['Version'] = $col }
                    elseif (-not $cols['Date'] -and $l -match $histDateLabel) { $cols['Date'] = $col }
                    elseif (-not $cols['Section'] -and $l -match $sectionLabel) { $cols['Section'] = $col }
                    elseif (-not $cols['Description'] -and $l -match $descLabel) { $cols['Description'] = $col }
                    elseif (-not $cols['Author'] -and $l -match $authorLabel) { $cols['Author'] = $col }
                }
                if (-not $cols['Version'] -or -not $cols['Date']) { continue }
                $entries = @($grid.Rows | Where-Object { $_ -gt $row -and $grid.Cells[$_].ContainsKey($cols['Version']) -and (ConvertTo-VersionText $grid.Cells[$_][$cols['Version']]) })
                if (-not $hist -or $entries.Count -gt $hist.Entries.Count) { $hist = [pscustomobject]@{ HeaderRow = $row; Cols = $cols; Entries = $entries } }
            }
        }

        $current = $null
        if ($hist -and $hist.Entries.Count -gt 0) { $current = ConvertTo-VersionText $book.Sheets[$histName].Cells[$hist.Entries[-1]][$hist.Cols['Version']] }
        elseif ($coverVersion) { $current = ConvertTo-VersionText $coverVersion.Text }
        if (-not $current -and -not $Version) { New-Result $f.Name 'スキップ' '表紙・改訂履歴に版数が見つかりません' $f.FullName; continue }
        $newVersion = $Version
        if (-not $newVersion) { $newVersion = Get-NextVersion $current }
        $newVersion = ConvertTo-VersionText $newVersion

        # 改訂履歴に書く行（改訂箇所・改訂内容）。-Description の行と、前の版と比べた下書き
        $entries = New-Object System.Collections.Generic.List[object]
        if ($Description) { $entries.Add([pscustomobject]@{ Section = $Section; Description = $Description }) }
        if ($oldPath) {
            $drafts = @(Get-HistoryDrafts $oldPath $f.FullName)
            if ($drafts.Count -gt $maxHistoryRows - $entries.Count) {
                $keep = $maxHistoryRows - $entries.Count - 1
                foreach ($d in $drafts[0..($keep - 1)]) { $entries.Add($d) }
                $entries.Add([pscustomobject]@{ Section = '全体'; Description = ('ほか {0} 件の変更（Compare-Excel.ps1 の差分を参照）' -f ($drafts.Count - $keep)) })
            }
            else { foreach ($d in $drafts) { $entries.Add($d) } }
            if ($entries.Count -eq 0) { New-Result $f.Name 'スキップ' ('前の版（{0}）と内容が同じです（改訂内容を -Description で指定してください）' -f [System.IO.Path]::GetFileName($oldPath)) $f.FullName; continue }
        }

        # 改訂履歴に足す行（最後に書いてある行の下。改訂履歴の列が空いていなければ足さない）
        $newRow = 0
        $last = 0
        if ($hist) {
            $grid = $book.Sheets[$histName]
            $last = $hist.HeaderRow
            foreach ($r in $grid.Rows) {
                if ($r -le $last) { continue }
                foreach ($col in $hist.Cols.Values) { if ($grid.Cells[$r].ContainsKey($col)) { $last = $r; break } }
            }
            $newRow = $last + 1
            $busy = @($newRow..($newRow + $entries.Count - 1) | Where-Object { $row = $_; $grid.Cells.ContainsKey($row) -and @($hist.Cols.Values | Where-Object { $grid.Cells[$row].ContainsKey($_) }).Count -gt 0 })
            if ($busy.Count -gt 0) { New-Result $f.Name 'スキップ' ('改訂履歴の {0} 行目が空いていないため、行を足せません' -f $busy[0]) $f.FullName; continue }
        }

        # 保存先: ファイル名に版数（_v1.0・第1.0版 など）があれば新しい版数の名前で別に保存し、無ければ上書きする
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $fm = [regex]::Match($baseName, '(?i)(?:(?<![a-z])v(?:er)?\.?\s*|第\s*)(\d+(?:\.\d+)*)')
        $outPath = $null
        if ($fm.Success) {
            $g = $fm.Groups[1]
            $outPath = Join-Path $f.DirectoryName ($baseName.Substring(0, $g.Index) + $newVersion + $baseName.Substring($g.Index + $g.Length) + $f.Extension)
            if (Test-Path -LiteralPath $outPath) { New-Result $f.Name 'スキップ' ('新しい版のファイルがすでにあります: {0}' -f [System.IO.Path]::GetFileName($outPath)) $f.FullName; continue }
        }
        $what = '版数 {0} → {1}' -f $current, $newVersion
        $shown = $f.Name
        if ($outPath) { $shown = '{0} → {1}' -f $f.Name, [System.IO.Path]::GetFileName($outPath) }
        if (-not $PSCmdlet.ShouldProcess($shown, $what)) { continue }

        $changes = @{}
        $docs = @{}
        foreach ($name in @($coverName, $histName)) {
            if ($name -and -not $docs.ContainsKey($name)) { $docs[$name] = Read-PartXmlForEdit $pkg $sheetParts[$name] }
        }
        $done = New-Object System.Collections.Generic.List[string]
        if ($coverVersion) {
            $doc = $docs[$coverName]
            $c = Get-SheetCellElement $doc (Get-SheetRowElement $doc $coverVersion.Row -Create) $coverVersion.Row $coverVersion.Col -Create
            $value = Set-VersionText $coverVersion.Text $newVersion
            $num = 0.0
            if ((Test-NumericCellElement $c) -and [double]::TryParse($value, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$num)) { $value = $num }
            Set-SheetCellValue $doc $c $value
            $done.Add('表紙の版数')
        }
        if ($coverDate) {
            $doc = $docs[$coverName]
            $c = Get-SheetCellElement $doc (Get-SheetRowElement $doc $coverDate.Row -Create) $coverDate.Row $coverDate.Col -Create
            Set-SheetCellValue $doc $c (Get-DateValue $c $coverDate.Text)
            $done.Add('表紙の更新日')
        }
        if ($hist) {
            $doc = $docs[$histName]
            $grid = $book.Sheets[$histName]
            # 版数・日付の書き方は最後の版の行に、セルの書式（罫線など）と結合は最後に書いてある行に合わせる
            $srcRow = $hist.HeaderRow
            if ($hist.Entries.Count -gt 0) { $srcRow = $hist.Entries[-1] }
            $srcEl = Get-SheetRowElement $doc $srcRow
            $styleEl = $null
            if ($last -gt $hist.HeaderRow) { $styleEl = Get-SheetRowElement $doc $last }
            $versionText = $newVersion
            if ($hist.Entries.Count -gt 0) { $versionText = Set-VersionText $grid.Cells[$srcRow][$hist.Cols['Version']] $newVersion }
            for ($i = 0; $i -lt $entries.Count; $i++) {
                $row = $newRow + $i
                $rowEl = Get-SheetRowElement $doc $row -Create -TemplateRow $styleEl
                if ($styleEl) {
                    foreach ($sc in (Get-XmlChildElements $styleEl 'c')) {
                        $col = Get-ColumnIndex $sc.GetAttribute('r')
                        $null = Get-SheetCellElement $doc $rowEl $row $col -Create -Style $sc.GetAttribute('s')
                    }
                    Copy-SheetRowMerges $doc $last $row
                }
                # 版数・日付・改訂者は最初の行だけに書く（続きの行は改訂箇所・改訂内容だけ）
                $set = [ordered]@{ Section = $entries[$i].Section; Description = $entries[$i].Description }
                if ($i -eq 0) { $set = [ordered]@{ Version = $versionText; Date = $null; Section = $entries[$i].Section; Description = $entries[$i].Description; Author = $Author } }
                foreach ($k in $set.Keys) {
                    if (-not $hist.Cols[$k]) { continue }
                    $col = $hist.Cols[$k]
                    $c = Get-SheetCellElement $doc $rowEl $row $col -Create
                    $value = $set[$k]
                    if ($k -eq 'Date') {
                        $old = $null
                        if ($srcEl) { $old = Get-SheetCellElement $doc $srcEl $srcRow $col }
                        $oldText = ''
                        if ($grid.Cells.ContainsKey($srcRow)) { $oldText = $grid.Cells[$srcRow][$col] }
                        if (Test-NumericCellElement $old) { $value = Get-DateValue $old $oldText } else { $value = Get-DateValue $null $oldText }
                    }
                    if ($null -ne $value) { Set-SheetCellValue $doc $c $value }
                }
            }
            $lastNew = $newRow + $entries.Count - 1
            Update-SheetDimension $doc $lastNew
            if ($lastNew -gt $newRow) { $done.Add(('改訂履歴の {0}〜{1} 行目' -f $newRow, $lastNew)) } else { $done.Add(('改訂履歴の {0} 行目' -f $newRow)) }
        }
        foreach ($name in $docs.Keys) { $changes[$sheetParts[$name].ToLowerInvariant()] = ConvertTo-XmlString $docs[$name] }
    }
    finally { $pkg.Zip.Dispose() }

    if ($outPath) { Save-ZipPackage -FilePath $f.FullName -Changes $changes -OutPath $outPath; $resultPath = $outPath }
    else { Save-ZipPackage -FilePath $f.FullName -Changes $changes; $resultPath = $f.FullName }
    New-Result ([System.IO.Path]::GetFileName($resultPath)) '版上げ' ('{0}（{1}）' -f $what, ($done -join '・')) $resultPath
}

# .SYNOPSIS
#     設計書を雛形から一括で作り（-Template / -List）、または設計書の版を上げます（-Bump）。Excel 不要です。
#
# .DESCRIPTION
#     雛形から作る: 雛形の xlsx に {{画面ID}} {{画面名}} のような差し込み欄を書いておくと、一覧（画面一覧の Excel や CSV）の
#     行ごとに雛形をコピーし、差し込み欄を一覧の値に置き換えます（セル・図形・コメント。Replace-Excel.ps1 を使います）。
#     {{作成日}}（{{日付}}）と {{作成者}} は -Date・-Author の値になります。同じ名前のファイルがあれば作りません。
#
#     版を上げる（-Bump）: 表紙の版数と更新日を新しい版数・日付にし、改訂履歴の最後の行の下に
#     版数・改訂日・改訂箇所・改訂内容・改訂者の行を足します（最後の行の書式と結合を写します）。
#     -CompareWith で前の版を渡すと、前の版との差分から改訂箇所・改訂内容の行を作ります。
#     ファイル名に版数（_v1.0・第1.0版 など）があれば新しい版数の名前で別のファイルに保存し、無ければ上書きします。
#     表紙と改訂履歴の見つけ方は Test-ExcelDeliverable.ps1 の版数のチェックと同じです。
#
# .PARAMETER Template
#     雛形の xlsx / xlsm。差し込み欄は {{名前}} の形で書きます。
#
# .PARAMETER List
#     一覧の Excel または CSV。差し込み欄と同じ名前の見出しの列を読みます（Excel は見出しがそろった表を探します）。
#
# .PARAMETER ListSheet
#     一覧の Excel のうち、読むシートの名前の正規表現。
#
# .PARAMETER FileName
#     作るファイルの名前。例: 'DD02_画面設計書_{{画面ID}}_{{画面名}}.xlsx'。既定は雛形のファイル名（差し込み欄が必要です）。
#
# .PARAMETER OutDir
#     作ったファイルを保存するフォルダ。既定はカレントフォルダ。
#
# .PARAMETER Bump
#     版を上げます（-Path の設計書の表紙と改訂履歴を書き換えます）。
#
# .PARAMETER Path
#     版を上げる設計書のファイルまたはフォルダ（複数可・ワイルドカード可）。
#
# .PARAMETER Description
#     改訂履歴の改訂内容。-CompareWith を指定しないときは必須です。
#
# .PARAMETER CompareWith
#     前の版のファイル、または前の版のフォルダ。前の版と比べ（Compare-Excel.ps1 -OutHistory）、変更した行ごとに
#     「合計金額（TOTAL_AMOUNT）: 桁数を 9 → 11 に変更」のような改訂箇所・改訂内容の行を改訂履歴に足します。
#     版数・改訂日・改訂者は最初の行だけに書きます。-Description も指定すると、その行を最初に書きます。
#     フォルダを指定すると、版数を除いた名前が同じファイル（DD03_テーブル定義書_v1.0.xlsx と _v1.1.xlsx など）と比べます。
#     一度に足す行は 20 行までで、超えた分は「ほか n 件の変更」の 1 行にまとめます。
#
# .PARAMETER Section
#     改訂履歴の改訂箇所。既定は「全体」。
#
# .PARAMETER Version
#     新しい版数（例: 2.0）。既定は最後の数を 1 つ上げた版数（1.0 → 1.1）。
#
# .PARAMETER HistorySheet
#     改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER Author
#     {{作成者}} と改訂者に書く名前。既定はログオンしているユーザー名。
#
# .PARAMETER Date
#     {{作成日}} と改訂日・表紙の更新日に書く日付。既定は今日。
#
# .EXAMPLE
#     # 画面一覧の行ごとに、雛形から画面設計書を作る
#     .\New-DesignDocument.ps1 -Template .\画面設計書_雛形.xlsx -List D:\設計書\DD01_画面一覧.xlsx -ListSheet '^画面一覧$' -FileName 'DD02_画面設計書_{{画面ID}}_{{画面名}}.xlsx' -OutDir D:\設計書 -Author '山田 太郎'
#
# .EXAMPLE
#     # 版を上げる（まず -WhatIf で確かめる）
#     .\New-DesignDocument.ps1 -Bump D:\設計書\DD03_テーブル定義書_v1.1.xlsx -Description '配送時間帯を追加' -Section 'T_ORDER' -Author '佐藤 花子' -WhatIf
#
# .EXAMPLE
#     # 前の版（納品フォルダ）との差分から、改訂履歴の改訂箇所・改訂内容を書く
#     .\New-DesignDocument.ps1 -Bump D:\設計書 -CompareWith D:\納品\v1.0 -Author '佐藤 花子'
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Replace-Excel.ps1 が必要です。
#     ヘッダー／フッター・シート名の差し込み欄は置き換えません。
