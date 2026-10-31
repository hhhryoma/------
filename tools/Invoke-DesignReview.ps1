[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path,

    [string] $Checklist,

    [string] $OutDir,

    [ValidateRange(2000, 200000)]
    [int] $MaxChars = 8000,

    [switch] $PromptOnly,
    [switch] $ImportOnly,
    [switch] $CheckOnly,

    [string] $SheetNameExclude,

    [string] $Glossary,

    [string] $Baseline,

    [string] $Recheck,

    [ValidateSet('Rows', 'Markdown')]
    [string] $TextFormat = 'Rows',

    [ValidateSet('Auto', 'Watch', 'Enter')]
    [string] $AnswerMode = 'Auto',

    [switch] $NoRecurse
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
if (-not (Test-Path -LiteralPath $searchScript)) {
    throw "Search-Excel.ps1 が見つかりません。Invoke-DesignReview.ps1 と同じフォルダに置いてください: $searchScript"
}
if (-not $Checklist) {
    $Checklist = Join-Path (Join-Path $PSScriptRoot 'review-checklists') '詳細設計レビュー観点.xlsx'
}
if (@($PromptOnly, $ImportOnly, $CheckOnly | Where-Object { $_ }).Count -gt 1) {
    throw '-PromptOnly・-ImportOnly・-CheckOnly は同時に指定できません。'
}
if (-not $Path -and -not $ImportOnly -and -not $CheckOnly) {
    throw 'レビューする設計書のフォルダまたはファイルを指定してください（例: .\Invoke-DesignReview.ps1 D:\Design\詳細設計）。'
}

# 保存先の既定は、1 つ目の設計書のフォルダと同じ場所の design-review（D:\設計書 なら D:\design-review）。
# 今いるフォルダによらない（設計書のフォルダの中で実行しても、レビュー結果を設計書のフォルダに置かない）。設計書を指定しないときは今いるフォルダ
if (-not $OutDir) {
    $OutDir = '.\design-review'
    $first = $null
    foreach ($p in @($Path)) { if (-not $first) { $first = @(Get-PathItem $p)[0] } }
    if ($first) {
        $folder = $first.FullName
        if (-not $first.PSIsContainer) { $folder = $first.DirectoryName }
        $dirInfo = New-Object System.IO.DirectoryInfo($folder)
        if ($dirInfo.Parent) { $dirInfo = $dirInfo.Parent }
        $OutDir = Join-Path $dirInfo.FullName 'design-review'
    }
}
$OutDir    = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$Checklist = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Checklist)
if ($Glossary) {
    $Glossary = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Glossary)
    if (-not (Test-Path -LiteralPath $Glossary)) { throw "用語集が見つかりません: $Glossary" }
}
if ($Baseline -and ($ImportOnly -or $CheckOnly)) { throw '-Baseline は、依頼を作るとき（-ImportOnly・-CheckOnly 以外）に指定してください。' }
if ($Recheck -and ($ImportOnly -or $CheckOnly)) { throw '-Recheck は、依頼を作るとき（-ImportOnly・-CheckOnly 以外）に指定してください。' }
if ($Recheck -and $Baseline) { throw '-Recheck と -Baseline は同時に指定できません（再確認では、指摘のあったシートの全体を渡します）。' }
if ($Recheck) {
    $Recheck = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Recheck)
    # フォルダを指定したら、その中のレビュー結果
    if (Test-Path -LiteralPath $Recheck -PathType Container) { $Recheck = Join-Path $Recheck 'レビュー結果.xlsx' }
    if (-not (Test-Path -LiteralPath $Recheck -PathType Leaf)) { throw "前のレビュー結果が見つかりません: $Recheck" }
    # 前のレビューのフォルダに出力すると、前の指摘一覧（人が書いた対応要否など）を再確認の結果で上書きしてしまう
    if ([string]::Equals([System.IO.Path]::GetFullPath($Recheck), [System.IO.Path]::GetFullPath((Join-Path $OutDir 'レビュー結果.xlsx')), [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "-Recheck の前のレビュー結果と同じフォルダには出力できません（前の指摘一覧を上書きしてしまいます）。-OutDir に別のフォルダを指定してください: $OutDir"
    }
}
if ($Baseline -and @(Get-PathItem $Baseline).Count -eq 0) { throw "前の版の設計書が見つかりません: $Baseline" }
$promptDir = Join-Path $OutDir 'prompts'
$answerDir = Join-Path $OutDir 'answers'
$manifestPath = Join-Path $OutDir 'requests.json'
$reportPath   = Join-Path $OutDir 'レビュー結果.xlsx'

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$nfkc = [System.Text.NormalizationForm]::FormKC

$severityOrder = @{ '高' = 0; '中' = 1; '低' = 2 }

# ---------------------------------------------------------------- 設計書のテキスト化


function Test-WordPath {
    param([string] $FilePath)
    return @('.docx', '.docm') -contains [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
}

function Get-ExcelCells {
    #     Search-Excel.ps1 で全セル（値・コメント・図形・シート名）を読み出す。
    #     ブックの並び順（シート順、行優先）のまま返る。Word の文書は、段落・表の行・テキストボックス・コメント（シートの列は本文・脚注などの部分）。
    #     設計書の非表示のシートと雛形のシート（ID の欄が差し込み欄のまま。New-DesignDocument.ps1 -TemplateSheet の雛形）は読まない（AllSheets なら読む）
    param([string] $FilePath, [switch] $AllSheets)

    $warnings = $null
    $searchArgs = @{
        Pattern = '\S'
        Path    = [System.Management.Automation.WildcardPattern]::Escape($FilePath)
        Target  = @('Value', 'Comment', 'Shape', 'SheetName')
        IncludeWord = $true
    }
    if (-not $AllSheets) { $searchArgs['SkipHiddenSheets'] = $true }
    if ($SheetNameExclude) { $searchArgs['SheetNameExclude'] = $SheetNameExclude }

    $hits = @(& $searchScript @searchArgs -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)
    if ($warnings) { foreach ($w in $warnings) { Write-Warning $w.Message } }
    if ($AllSheets) { return $hits }
    # 雛形のシートは、依頼を作り終えてから知らせる
    return (Select-DesignHits $hits $script:designSelection -NoWarning)
}

function ConvertTo-DesignText {
    #     設計書を AI が読める行形式のテキストにする。
    #       「行番号| 列:値 | 列:値 ...」の形で、セル番地を AI が引用できるようにする。
    #     方眼紙の帳票ヘッダ（各シート 1〜3 行目）が前のシートと同じなら省略して文字数を節約する。
    #     Format: Rows か Markdown（省略時は -TextFormat）。返り値: シート名 → テキスト の順序付き辞書
    param($Cells, [string] $Format = $TextFormat)

    if (@($Cells).Count -gt 0 -and (Test-WordPath @($Cells)[0].Path)) { return ConvertTo-WordText $Cells }
    if ($Format -eq 'Markdown') { return ConvertTo-DesignMarkdown $Cells }

    $sheets = [ordered]@{}
    foreach ($c in $Cells) {
        if (-not $sheets.Contains($c.Sheet)) { $sheets[$c.Sheet] = New-Object System.Collections.Generic.List[object] }
        if ($c.Kind -ne 'シート名') { $sheets[$c.Sheet].Add($c) }
    }

    $seenHeader = @{}
    $result = [ordered]@{}
    foreach ($name in $sheets.Keys) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("### シート: $name").Append("`n")

        $items = $sheets[$name]
        if ($items.Count -eq 0) {
            [void]$sb.Append('（空）').Append("`n")
            $result[$name] = $sb.ToString()
            continue
        }

        # 値は行ごとにまとめる（Search-Excel は行優先で返すので順序はそのまま）
        $rows = [ordered]@{}
        foreach ($c in $items) {
            if ($c.Kind -ne '値') { continue }
            $m = [regex]::Match($c.Cell, '^([A-Z]+)(\d+)$')
            if (-not $m.Success) { continue }
            # [ordered] に整数で引くと位置指定になるため、行番号は文字列のキーにする
            $r = $m.Groups[2].Value
            if (-not $rows.Contains($r)) { $rows[$r] = New-Object System.Collections.Generic.List[string] }
            $rows[$r].Add($m.Groups[1].Value + ':' + $c.Text)
        }

        $omitted = $false
        foreach ($r in $rows.Keys) {
            $line = $rows[$r] -join ' | '
            if ([int]$r -le 3) {
                $key = "$r|$line"
                if ($seenHeader.ContainsKey($key)) {
                    if (-not $omitted) { [void]$sb.Append('（1～3行目の帳票ヘッダは前のシートと同じため省略）').Append("`n") }
                    $omitted = $true
                    continue
                }
                $seenHeader[$key] = $true
            }
            [void]$sb.Append("$r| $line").Append("`n")
        }

        foreach ($c in $items) {
            if ($c.Kind -eq '図形') {
                [void]$sb.Append(('図形（{0}付近）「{1}」: {2}' -f $c.Cell, $c.Name, $c.Text)).Append("`n")
            }
            elseif ($c.Kind -eq 'コメント') {
                [void]$sb.Append(('コメント（{0}）: {1}' -f $c.Cell, $c.Text)).Append("`n")
            }
        }
        $result[$name] = $sb.ToString()
    }
    return $result
}

function ConvertTo-WordText {
    #     Word の設計書を、段落・表の行ごとに「段落 12| 文」「表 1 の 3 行目| セル | セル」の形のテキストにする。
    #     テキストボックスは「テキストボックス 1| 文」、コメントは「コメント（段落 3）: 文」。
    #     返り値: 部分（本文・脚注など。シートの代わり）→ テキスト の順序付き辞書
    param($Cells)
    $parts = [ordered]@{}
    foreach ($c in $Cells) {
        if ($c.Kind -eq 'シート名') { continue }
        if (-not $parts.Contains($c.Sheet)) {
            $parts[$c.Sheet] = New-Object System.Text.StringBuilder
            [void]$parts[$c.Sheet].Append("### シート: $($c.Sheet)").Append("`n")
        }
        if ($c.Kind -eq 'コメント') { [void]$parts[$c.Sheet].Append(('コメント（{0}）: {1}' -f $c.Cell, $c.Text)).Append("`n") }
        else { [void]$parts[$c.Sheet].Append(('{0}| {1}' -f $c.Cell, $c.Text)).Append("`n") }
    }
    $result = [ordered]@{}
    foreach ($k in $parts.Keys) { $result[$k] = $parts[$k].ToString() }
    return $result
}

function ConvertTo-DesignMarkdown {
    #     -TextFormat Markdown: Export-ExcelToMarkdown と同じ形（表・「■」の小見出し・箇条書き）にする。
    #     AI がセル番地を引用できるよう、表に行番号と列（K:桁数）、箇条書きにセル番地を付ける。
    #     行形式と同じく、帳票ヘッダ（各シート 1〜3 行目）が前のシートと同じなら省略する。
    #     返り値: シート名 → テキスト の順序付き辞書
    param($Cells)

    # 1〜3 行目の「列:値」の並びが前のシートと同じ行を、シート|行 で覚える
    $headerRows = [ordered]@{}
    foreach ($c in $Cells) {
        if ($c.Kind -ne '値') { continue }
        $m = [regex]::Match([string]$c.Cell, '^([A-Z]+)(\d+)$')
        if (-not $m.Success -or [int]$m.Groups[2].Value -gt 3) { continue }
        $key = $c.Sheet + '|' + $m.Groups[2].Value
        if (-not $headerRows.Contains($key)) { $headerRows[$key] = New-Object System.Collections.Generic.List[string] }
        $headerRows[$key].Add($m.Groups[1].Value + ':' + $c.Text)
    }
    $seen = @{}
    $omitRows = @{}
    $omitSheets = @{}
    foreach ($key in $headerRows.Keys) {
        $sheet = $key.Substring(0, $key.LastIndexOf('|'))
        $line = $key.Substring($key.LastIndexOf('|') + 1) + '|' + ($headerRows[$key] -join ' | ')
        if ($seen.ContainsKey($line)) { $omitRows[$key] = $true; $omitSheets[$sheet] = $true }
        else { $seen[$line] = $true }
    }
    $kept = @($Cells | Where-Object {
            if ($_.Kind -ne '値') { return $true }
            $m = [regex]::Match([string]$_.Cell, '^[A-Z]+(\d+)$')
            -not ($m.Success -and $omitRows.ContainsKey($_.Sheet + '|' + $m.Groups[1].Value))
        })

    $result = [ordered]@{}
    $sheets = ConvertTo-SheetMarkdownLines $kept -CellAddress -ColumnLetter
    foreach ($name in $sheets.Keys) {
        $lines = $sheets[$name]
        if ($omitSheets.ContainsKey($name)) {
            $lines.Insert(2, '（1～3行目の帳票ヘッダは前のシートと同じため省略）')
            $lines.Insert(3, '')
            if ($lines.Count -gt 4 -and $lines[4] -eq '（空）') { $lines.RemoveRange(4, 2) }
        }
        while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
        $result[$name] = ($lines -join "`n") + "`n`n"
    }
    return $result
}

function Get-VersionlessRelative {
    #     相対パスから版数を除いた形（区切りは / にそろえる）。サブフォルダの違う同じ名前の設計書を見分けるため
    param([string] $Relative)
    return Get-VersionlessName ([string]$Relative).Replace('\', '/')
}

function Find-BaselineFiles {
    #     前の版（-Baseline）の設計書を、今の設計書と対応付ける。版数を除いた相対パス（SCR001\画面設計書.xlsx など）が同じものを先に、
    #     無ければ版数を除いたファイル名（Get-VersionlessName）が同じものに当てる（名前の同じ前の版が複数あれば、どれとも対応付けない）。
    #     戻り値: 今の設計書のパス → 前の版のパス
    param($Files)
    $byRelative = @{}
    $byName = @{}
    foreach ($item in @(Get-PathItem $Baseline)) {
        $found = @($item)
        $root = [System.IO.Path]::GetDirectoryName($item.FullName)
        if ($item.PSIsContainer) {
            $found = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)
            $root = $item.FullName.TrimEnd('\', '/')
        }
        foreach ($o in $found) {
            if (@('.xlsx', '.xlsm', '.xls', '.docx', '.docm') -notcontains $o.Extension.ToLowerInvariant() -or $o.Name.StartsWith('~$')) { continue }
            $k = Get-VersionlessRelative $o.FullName.Substring($root.Length).TrimStart('\', '/')
            if (-not $byRelative.ContainsKey($k)) { $byRelative[$k] = $o.FullName }
            $k = Get-VersionlessName $o.Name
            if (-not $byName.ContainsKey($k)) { $byName[$k] = New-Object System.Collections.Generic.List[string] }
            if (-not $byName[$k].Contains($o.FullName)) { $byName[$k].Add($o.FullName) }
        }
    }
    $pairs = @{}
    $used = @{}
    foreach ($f in $Files) {
        $k = Get-VersionlessRelative $f.Relative
        if ($byRelative.ContainsKey($k) -and $byRelative[$k] -ne $f.Path) { $pairs[$f.Path] = $byRelative[$k]; $used[$byRelative[$k]] = $true }
    }
    foreach ($f in $Files) {
        if ($pairs.ContainsKey($f.Path)) { continue }
        $k = Get-VersionlessName $f.Name
        if (-not $byName.ContainsKey($k)) { continue }
        $candidates = @($byName[$k] | Where-Object { -not $used.ContainsKey($_) -and $_ -ne $f.Path })
        if ($candidates.Count -eq 1) { $pairs[$f.Path] = $candidates[0]; $used[$candidates[0]] = $true }
        elseif ($candidates.Count -gt 1) { Write-Warning ("前の版に、同じ名前の設計書が複数あるため、前の版と比べません（全体を渡します）: {0}" -f $f.Relative) }
    }
    return $pairs
}

function ConvertTo-DiffText {
    #     -Baseline: 設計書のうち、前の版から変わった行（★）と、その節の見出し（「■」の行と次の行）だけを、行の形のテキストにする（-TextFormat によらない）。
    #     変わったセルには前の版の値（旧: …）を添え、消えた行は「★削除（旧 n 行目）」で示す。新しいシートと、列を足した・消したシートは全部を入れ、
    #     どの行も判定するよう、すべての行・図形・コメントに ★ を付ける。
    #     Diffs は Compare-Excel.ps1 の結果。返り値: シート名 → テキスト の順序付き辞書（変わっていないシートは入れない）
    param($Cells, $Diffs)
    $rowOf = { param([string] $Cell) $m = [regex]::Match([string]$Cell, '^[A-Z]+(\d+)$'); if ($m.Success) { [int]$m.Groups[1].Value } else { 0 } }
    $changed = @{}; $old = @{}; $notes = @{}; $whole = @{}
    $note = {
        param([string] $Sheet, [string] $Text)
        if (-not $notes.ContainsKey($Sheet)) { $notes[$Sheet] = New-Object System.Collections.Generic.List[string] }
        $notes[$Sheet].Add($Text)
    }
    $removedSheets = New-Object System.Collections.Generic.List[string]
    foreach ($d in $Diffs) {
        $sheet = [string]$d.Sheet
        if (-not $changed.ContainsKey($sheet)) { $changed[$sheet] = @{}; $old[$sheet] = @{} }
        if (@('値', '数式') -contains $d.Kind) {
            $r = & $rowOf $d.NewCell
            if ($r -le 0) { continue }
            $changed[$sheet][$r] = $true
            if ($d.Change -ne '追加') { $old[$sheet][[string]$d.NewCell] = [string]$d.Old }
        }
        elseif ($d.Kind -eq '行') {
            if ($d.Change -eq '削除') { & $note $sheet ('★削除（旧 {0} 行目）| {1}' -f (& $rowOf $d.OldCell), $d.Old) }
            elseif ((& $rowOf $d.NewCell) -gt 0) { $changed[$sheet][(& $rowOf $d.NewCell)] = $true }
        }
        elseif (@('コメント', '図形') -contains $d.Kind) {
            if ($d.Change -eq '削除') { & $note $sheet ('★{0}を削除（旧 {1}）: {2}' -f $d.Kind, $d.OldCell, $d.Old) }
            elseif ($d.Change -eq '変更') { & $note $sheet ('★{0}（{1}付近）: {2}（旧: {3}）' -f $d.Kind, $d.NewCell, $d.New, $d.Old) }
            elseif ($d.Change -eq '追加') { & $note $sheet ('★{0}（{1}付近）: {2}' -f $d.Kind, $d.NewCell, $d.New) }
        }
        elseif ($d.Kind -eq '連番') { & $note $sheet ('★{0}' -f $d.Detail) }
        elseif ($d.Kind -eq '列' -or ($d.Kind -eq 'シート' -and $d.Change -eq '追加')) { $whole[$sheet] = $true }
        elseif ($d.Kind -eq 'シート' -and $d.Change -eq '名前変更') { & $note $sheet ('★シート名を「{0}」から変更' -f $d.Old) }
        elseif ($d.Kind -eq 'シート' -and $d.Change -eq '削除') { $removedSheets.Add($sheet) }
    }
    # シートごとのセル（値）と、シートの並び
    $sheets = [ordered]@{}
    foreach ($c in $Cells) {
        if (-not $sheets.Contains($c.Sheet)) { $sheets[$c.Sheet] = New-Object System.Collections.Generic.List[object] }
        $sheets[$c.Sheet].Add($c)
    }
    $result = [ordered]@{}
    if ($removedSheets.Count -gt 0) { $result['（削除したシート）'] = "### 前の版から削除したシート`n" + (($removedSheets | ForEach-Object { "★$_" }) -join "`n") + "`n" }
    foreach ($name in $sheets.Keys) {
        if ($whole.ContainsKey($name)) {
            # （List を @() に入れると PowerShell 7.4 では例外になるので、配列にして渡す）
            $t = ConvertTo-DesignText -Cells $sheets[$name].ToArray() -Format Rows
            foreach ($k in $t.Keys) {
                $lines = @(($t[$k].TrimEnd("`n") -split "`n") | Select-Object -Skip 1 | ForEach-Object { if ($_ -match '^(\d+\||図形（|コメント（)') { '★' + $_ } else { $_ } })
                $result[$k] = ((@("### シート: $k（前の版から大きく変わったため、全体）") + $lines) -join "`n") + "`n"
            }
            continue
        }
        $rows = @{}
        if ($changed.ContainsKey($name)) { $rows = $changed[$name] }
        if ($rows.Count -eq 0 -and -not $notes.ContainsKey($name)) { continue }
        # 行 → 列の順のセル
        $byRow = @{}
        foreach ($c in $sheets[$name]) {
            if ($c.Kind -ne '値') { continue }
            $m = [regex]::Match([string]$c.Cell, '^([A-Z]+)(\d+)$')
            if (-not $m.Success) { continue }
            $r = [int]$m.Groups[2].Value
            if (-not $byRow.ContainsKey($r)) { $byRow[$r] = New-Object System.Collections.Generic.List[object] }
            $byRow[$r].Add([pscustomobject]@{ Col = $m.Groups[1].Value; Text = $c.Text })
        }
        # 変わった行の節の見出し: 上にある「■」の行と、その次の行（表の見出し）
        $sections = @($byRow.Keys | Where-Object { $byRow[$_][0].Text.TrimStart().StartsWith('■') } | Sort-Object)
        $all = @($byRow.Keys | Sort-Object)
        $show = @{}
        foreach ($r in $rows.Keys) {
            $show[$r] = $true
            $sec = @($sections | Where-Object { $_ -le $r } | Select-Object -Last 1)
            if ($sec.Count -gt 0) {
                $show[$sec[0]] = $true
                $next = @($all | Where-Object { $_ -gt $sec[0] } | Select-Object -First 1)
                if ($next.Count -gt 0) { $show[$next[0]] = $true }
            }
        }
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("### シート: $name（前の版から変わった行だけ）").Append("`n")
        foreach ($r in @($show.Keys | Sort-Object)) {
            # 列の順に、セルの値（変わったセルは前の版の値も）と、値を消したセル
            $parts = New-Object System.Collections.Generic.List[object]
            $seen = @{}
            if ($byRow.ContainsKey($r)) {
                foreach ($c in $byRow[$r]) {
                    $cell = '{0}{1}' -f $c.Col, $r
                    $seen[$cell] = $true
                    $t = '{0}:{1}' -f $c.Col, $c.Text
                    if ($rows.ContainsKey($r) -and $old[$name].ContainsKey($cell)) { $t += '（旧: {0}）' -f $old[$name][$cell] }
                    $parts.Add([pscustomobject]@{ Col = (Get-ColumnIndex $c.Col); Text = $t })
                }
            }
            if ($rows.ContainsKey($r)) {
                foreach ($cell in @($old[$name].Keys | Where-Object { (& $rowOf $_) -eq $r -and -not $seen.ContainsKey($_) })) {
                    $col = $cell -replace '\d', ''
                    $parts.Add([pscustomobject]@{ Col = (Get-ColumnIndex $col); Text = ('{0}:（削除。旧: {1}）' -f $col, $old[$name][$cell]) })
                }
            }
            $parts = @($parts | Sort-Object Col | ForEach-Object { $_.Text })
            $mark = ''
            if ($rows.ContainsKey($r)) { $mark = '★' }
            [void]$sb.Append(('{0}{1}| {2}' -f $mark, $r, ($parts -join ' | '))).Append("`n")
        }
        if ($notes.ContainsKey($name)) { foreach ($n in $notes[$name]) { [void]$sb.Append($n).Append("`n") } }
        $result[$name] = $sb.ToString()
    }
    return $result
}

# ---------------------------------------------------------------- 観点表


function ConvertTo-MethodName {
    #     「確かめ方」の値（AI・ツール・両方）。それ以外は空
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim().ToUpperInvariant()
    if (@('AI', 'ツール', '両方') -contains $t) { return $t }
    return ''
}

function ConvertTo-RuleSpecs {
    #     「ツールの確認」の欄（「,」「、」・改行区切り）を、確認の名前と（ ）の中の種類に分ける。例: 参照先なし（メッセージ）
    param([string] $Text)
    $specs = New-Object System.Collections.Generic.List[object]
    foreach ($t in (([string]$Text) -split '[,、\r\n]+')) {
        $t = $t.Trim()
        if (-not $t) { continue }
        $m = [regex]::Match($t, '^(?<name>.+?)\s*[（(]\s*(?<kind>[^（()）]*?)\s*[）)]$')
        if ($m.Success) { $specs.Add([pscustomobject]@{ Text = $t; Name = $m.Groups['name'].Value; Kind = $m.Groups['kind'].Value }) }
        else { $specs.Add([pscustomobject]@{ Text = $t; Name = $t; Kind = '' }) }
    }
    return , $specs
}

function Read-Checklist {
    #     観点表（xlsx / csv）を読む。1 行目または「観点ID」を含む行を見出しとし、列は見出しの文字で特定する。
    #     確かめ方（AI・ツール・両方）が空欄なら、ツールの確認があれば両方、無ければ AI
    param([string] $FilePath)

    if (-not (Test-Path -LiteralPath $FilePath)) { throw "観点表が見つかりません: $FilePath" }

    $records = New-Object System.Collections.Generic.List[object]
    if ([System.IO.Path]::GetExtension($FilePath) -eq '.csv') {
        $headers = @()
        $rowNo = 1
        foreach ($r in (Import-Csv -LiteralPath $FilePath -Encoding UTF8)) {
            $rowNo++
            if ($headers.Count -eq 0) { $headers = @($r.PSObject.Properties | ForEach-Object { $_.Name }) }
            $r | Add-Member -NotePropertyName '__Row' -NotePropertyValue $rowNo
            $records.Add($r)
        }
    }
    else {
        $cells = Get-ExcelCells -FilePath $FilePath -AllSheets | Where-Object { $_.Kind -eq '値' }
        $firstSheet = ($cells | Select-Object -First 1).Sheet
        $grid = [ordered]@{}
        foreach ($c in ($cells | Where-Object { $_.Sheet -eq $firstSheet })) {
            $m = [regex]::Match($c.Cell, '^([A-Z]+)(\d+)$')
            $r = $m.Groups[2].Value   # 文字列のキー（[ordered] に整数で引くと位置指定になる）
            if (-not $grid.Contains($r)) { $grid[$r] = @{} }
            # 表示用の ↵ を改行に戻す
            $grid[$r][$m.Groups[1].Value] = $c.Text.Replace('↵', "`n")
        }

        $headerRow = $null; $columns = @{}
        foreach ($r in $grid.Keys) {
            if ($grid[$r].Values -contains '観点ID') {
                $headerRow = $r
                foreach ($k in $grid[$r].Keys) { $columns[$grid[$r][$k]] = $k }
                break
            }
        }
        if ($null -eq $headerRow) { throw "観点表の先頭シートに「観点ID」の見出しがありません: $FilePath" }
        $headers = @($columns.Keys)

        foreach ($r in $grid.Keys) {
            if ([int]$r -le [int]$headerRow) { continue }
            $o = [ordered]@{ '__Row' = [int]$r }
            foreach ($h in $columns.Keys) { $o[$h] = $grid[$r][$columns[$h]] }
            $records.Add([pscustomobject]$o)
        }
    }

    $all = New-Object System.Collections.Generic.List[object]
    foreach ($r in $records) {
        if (-not ([string]$r.'観点ID').Trim()) { continue }
        $enabledRaw = ''
        if ($r.PSObject.Properties['有効']) { $enabledRaw = ([string]$r.'有効').Trim() }
        $target = [string]$r.'対象設計書'
        if (-not $target.Trim()) { $target = '*' }
        $methodRaw = ''
        if ($r.PSObject.Properties['確かめ方']) { $methodRaw = ([string]$r.'確かめ方').Trim() }
        $rules = ConvertTo-RuleSpecs $(if ($r.PSObject.Properties['ツールの確認']) { [string]$r.'ツールの確認' } else { '' })
        $method = ConvertTo-MethodName $methodRaw
        if (-not $method) { $method = $(if ($rules.Count -gt 0) { '両方' } else { 'AI' }) }
        $all.Add([pscustomobject]@{
            Row        = $r.'__Row'
            Id         = ([string]$r.'観点ID').Trim()
            Category   = [string]$r.'分類'
            Point      = [string]$r.'観点'
            Criteria   = [string]$r.'判定基準'
            Severity   = ([string]$r.'重要度').Trim()
            Targets    = @($target -split '[,、\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            Refs       = @(([string]$r.'参照設計書') -split '[,、\r\n]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
            EnabledRaw = $enabledRaw
            Enabled    = (@('×', '✕', '-', '－', '0', 'FALSE', '無効') -notcontains $enabledRaw)
            MethodRaw  = $methodRaw
            Method     = $method
            Rules      = $rules
            # AI に依頼する観点・ツールで確かめる観点
            ByAi       = ($method -ne 'ツール')
            ByTool     = ($method -ne 'AI' -and $rules.Count -gt 0)
        })
    }

    return [pscustomobject]@{
        Path    = $FilePath
        Headers = $headers
        All     = $all
        Points  = @($all | Where-Object { $_.Enabled })
    }
}

function Test-Checklist {
    #     観点表の書き方の誤りを調べる。エラーはレビューが正しく動かないもの、警告は結果の質に関わるもの。
    #     Files を渡すと、対象設計書・参照設計書のパターンが実際の設計書に当たるかも調べる。
    param($List, $Files)

    $issues = New-Object System.Collections.Generic.List[object]
    $add = {
        param($Level, $Row, $Id, $Message)
        $issues.Add([pscustomobject]@{ Level = $Level; Row = $Row; Id = $Id; Message = $Message })
    }

    foreach ($h in @('観点', '判定基準')) {
        if ($List.Headers -notcontains $h) { & $add 'エラー' $null $null "見出し「$h」の列がありません" }
    }
    foreach ($h in @('対象設計書', '重要度')) {
        if ($List.Headers -notcontains $h) { & $add '警告' $null $null "見出し「$h」の列がありません（対象設計書が無い場合はすべての設計書が対象になります）" }
    }

    $points = @($List.Points)
    if ($points.Count -eq 0) { & $add 'エラー' $null $null '有効な観点が 1 つもありません' }

    # 回答の照合は大文字小文字を区別しないため、ID の重複も区別せずに調べる
    foreach ($g in ($points | Group-InOrder { $_.Id.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 })) {
        $rows = ($g.Group | ForEach-Object { "$($_.Row)行目" }) -join '、'
        foreach ($p in $g.Group) { & $add 'エラー' $p.Row $p.Id "観点IDが重複しています（$rows）" }
    }

    $vague = '適切|妥当|十分|適宜|必要に応じて|問題な[いく]'
    foreach ($p in $List.All) {
        if ($p.Id -match '[\s|｜]') {
            & $add 'エラー' $p.Row $p.Id '観点IDに空白や「|」が含まれています（回答を照合できません）'
        }
        if ($p.EnabledRaw -and @('○', '〇', '◯', '1', 'TRUE', '有効', '×', '✕', '-', '－', '0', 'FALSE', '無効') -notcontains $p.EnabledRaw) {
            & $add '警告' $p.Row $p.Id "「有効」の値「$($p.EnabledRaw)」は ○ か × にしてください（有効として扱います）"
        }
        if (-not $p.Enabled) { continue }

        if (-not $p.Point.Trim()) { & $add 'エラー' $p.Row $p.Id '観点が空欄です' }
        # 判定基準は AI に渡すときだけ調べる（ツールだけで確かめる観点は AI に渡さない）
        if ($p.ByAi -and -not $p.Criteria.Trim()) {
            & $add '警告' $p.Row $p.Id '判定基準が空欄です（AI の判定がぶれやすくなります）'
        }
        elseif ($p.ByAi) {
            # 「」や例示の中は判定基準そのものではないので除いて調べる
            $plain = [regex]::Replace($p.Criteria, '「[^」]*」|（例[^）]*）|\(例[^)]*\)', '')
            $found = @([regex]::Matches($plain, $vague) | ForEach-Object { $_.Value } | Sort-Object -Unique)
            if ($found.Count -gt 0) {
                & $add '警告' $p.Row $p.Id ("判定基準にあいまいな言葉があります（{0}）。何が書かれていれば OK かを具体的に書いてください" -f ($found -join '、'))
            }
        }
        if ($p.ByAi -and ($p.Point + $p.Criteria) -match '[|｜]') {
            & $add '警告' $p.Row $p.Id '観点・判定基準に「|」があります（プロンプトの区切りと紛れるため、別の記号にしてください）'
        }

        # 確かめ方・ツールの確認
        if ($p.MethodRaw -and -not (ConvertTo-MethodName $p.MethodRaw)) {
            & $add '警告' $p.Row $p.Id ("「確かめ方」の値「{0}」は AI・ツール・両方 のどれかにしてください（{1}として扱います）" -f $p.MethodRaw, $p.Method)
        }
        if ($p.Method -ne 'AI' -and $p.Rules.Count -eq 0) {
            & $add 'エラー' $p.Row $p.Id ("確かめ方が「{0}」ですが、「ツールの確認」が空欄です" -f $p.Method)
        }
        elseif ($p.Method -eq 'AI' -and $p.Rules.Count -gt 0) {
            & $add '警告' $p.Row $p.Id '確かめ方が「AI」のため、「ツールの確認」は使いません'
        }
        foreach ($r in $p.Rules) {
            if (-not $toolRules.Contains($r.Name)) {
                & $add 'エラー' $p.Row $p.Id ("ツールの確認「{0}」はありません（使える確認: {1}）" -f $r.Name, (@($toolRules.Keys) -join '、'))
            }
            elseif ($r.Kind -and $toolRules[$r.Name].Source -ne '整合性') {
                & $add 'エラー' $p.Row $p.Id ("ツールの確認「{0}」には、（ ）で種類を付けられません" -f $r.Text)
            }
            elseif ($r.Kind -and (Get-ConsistencyKinds) -notcontains $r.Kind) {
                & $add 'エラー' $p.Row $p.Id ("ツールの確認「{0}」の種類「{1}」は、整合性チェックの参照のルールにありません（種類: {2}）" -f $r.Text, $r.Kind, ((Get-ConsistencyKinds) -join '、'))
            }
        }
        if ($p.Severity -and @('高', '中', '低') -notcontains $p.Severity) {
            & $add '警告' $p.Row $p.Id "重要度「$($p.Severity)」は 高／中／低 のどれかにしてください（並び順が最後になります）"
        }
        elseif (-not $p.Severity) {
            & $add '警告' $p.Row $p.Id '重要度が空欄です'
        }

        if ($Files) {
            foreach ($t in $p.Targets) {
                if (-not @($Files | Where-Object { $_.Name -like $t }).Count) {
                    & $add '警告' $p.Row $p.Id "対象設計書「$t」に当たる設計書がありません（この観点は使われません）"
                }
            }
            foreach ($t in $p.Refs) {
                if (-not @($Files | Where-Object { $_.Name -like $t }).Count) {
                    & $add '警告' $p.Row $p.Id "参照設計書「$t」に当たる設計書がありません"
                }
            }
        }
    }

    if ($Files -and $points.Count -gt 0) {
        # 種類別の観点（対象設計書が * 以外）があるのに、共通の観点しか当たらない設計書も知らせる。
        # ファイル名の付け方がパターンと合っていない可能性がある
        $specific = @($points | Where-Object { @($_.Targets | Where-Object { $_ -ne '*' }).Count -gt 0 })
        foreach ($f in $Files) {
            if (-not @($points | Where-Object { Test-NameLike $f.Name $_.Targets }).Count) {
                & $add '警告' $null $null "設計書「$($f.Relative)」にはどの観点も当たりません（ファイル名と対象設計書のパターンを確認してください）"
            }
            elseif ($specific.Count -gt 0 -and -not @($specific | Where-Object { Test-NameLike $f.Name @($_.Targets | Where-Object { $_ -ne '*' }) }).Count) {
                & $add '情報' $null $null "設計書「$($f.Relative)」には共通の観点（対象設計書 *）しか当たりません"
            }
        }
    }
    return ,$issues
}

function Show-ChecklistIssues {
    #     チェック結果を表示する。エラーがあれば $true を返す。
    param($List, $Issues)

    $errors = @($Issues | Where-Object Level -eq 'エラー').Count
    $warns  = @($Issues | Where-Object Level -eq '警告').Count
    $infos  = @($Issues | Where-Object Level -eq '情報').Count
    $disabled = @($List.All | Where-Object { -not $_.Enabled }).Count

    Write-Host ("観点表: {0}（有効 {1} 観点、無効 {2} 観点）" -f $List.Path, $List.Points.Count, $disabled)
    $byMethod = @{ 'AI' = 0; 'ツール' = 0; '両方' = 0 }
    foreach ($p in $List.Points) { $byMethod[$p.Method]++ }
    if ($byMethod['ツール'] + $byMethod['両方'] -gt 0) {
        Write-Host ("  確かめ方: AI {0} 観点・ツール {1} 観点・両方 {2} 観点" -f $byMethod['AI'], $byMethod['ツール'], $byMethod['両方'])
    }
    if ($Issues.Count -eq 0) {
        Write-Host '  観点表のチェック: 問題ありません'
        return $false
    }
    Write-Host ("  観点表のチェック: エラー {0} 件 / 警告 {1} 件 / 情報 {2} 件" -f $errors, $warns, $infos)
    $order = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
    foreach ($i in ($Issues | Sort-Object @{ e = { $order[$_.Level] } }, @{ e = { if ($_.Row) { [int]$_.Row } else { 0 } } })) {
        $where = ''
        if ($i.Row) { $where = "$($i.Row)行目 " }
        if ($i.Id) { $where += "$($i.Id) " }
        $color = 'Yellow'
        if ($i.Level -eq 'エラー') { $color = 'Red' }
        elseif ($i.Level -eq '情報') { $color = 'Gray' }
        Write-Host ("  [{0}] {1}{2}" -f $i.Level, $where, $i.Message) -ForegroundColor $color
    }
    return ($errors -gt 0)
}

function Get-DesignFiles {
    #     レビュー対象の設計書を列挙する（出力フォルダと観点表は除く）。同じ文書の古い版・ツールの出力（前のレビュー結果など）・
    #     書き換える前のファイルを残したフォルダの中は読まない（Select-DesignFiles。読まなかったものは知らせる）
    $script:designSelection = Select-DesignFiles -Path $Path -NoRecurse:$NoRecurse -Exclude @($OutDir, $Checklist) -Extensions @('.xlsx', '.xlsm', '.xls') -IncludeWord
    Write-DesignSkipWarning $script:designSelection
    $files = New-Object System.Collections.Generic.List[object]
    foreach ($f in $script:designSelection.Files) { $files.Add([pscustomobject]@{ Path = $f.Path; Name = $f.Name; Relative = $f.Relative }) }
    return ,$files
}

function Test-NameLike {
    param([string] $Name, [string[]] $Patterns)
    foreach ($p in $Patterns) { if ($Name -like $p) { return $true } }
    return $false
}

# ---------------------------------------------------------------- ツールの確認（機械で確かめる観点）


# 観点表の「ツールの確認」に書く名前 → 確かめ方。Source は使うツール（提出前チェック・表記ゆれ・整合性）か、
# このスクリプトの中の確認（画面・テーブル）。Items はツールの指摘の項目のうち、この確認にあたるもの。Text は AI に伝える説明
$toolRules = [ordered]@{
    '版数'                     = @{ Source = '提出前チェック'; Items = @('表紙の版数', '表紙の更新日', 'ファイル名の版数', '改訂履歴の版数'); Text = '表紙・改訂履歴・ファイル名の版数と更新日の食い違い' }
    '残っている文言'           = @{ Source = '提出前チェック'; Items = @('残っている文言'); Text = 'TODO・TBD・未定・検討中などの消し忘れ' }
    '表記ゆれ'                 = @{ Source = '表記ゆれ'; Text = 'カタカナの長音（ー）や全角・半角の表記ゆれ' }
    '参照先なし'               = @{ Source = '整合性'; Items = @('参照先なし'); Text = '定義されていない ID の参照' }
    '重複定義'                 = @{ Source = '整合性'; Items = @('重複定義'); Text = '同じ ID の重複した定義' }
    '一覧との違い'             = @{ Source = '整合性'; Items = @('一覧との違い'); Text = '画面一覧・テーブル一覧と設計書の過不足・名前の違い' }
    'ヘッダの違い'             = @{ Source = '整合性'; Items = @('ヘッダの違い'); Text = 'シートのヘッダ・ファイル名の ID と名前の食い違い' }
    '組み合わせの違い'         = @{ Source = '整合性'; Items = @('組み合わせの違い'); Text = '画面遷移の一覧・イベント定義の遷移先・画面概要の遷移先の食い違い' }
    '属性の違い'               = @{ Source = '整合性'; Items = @('属性の違い'); Text = '画面項目・IF の項目と、参照先のカラムの型・桁数・必須の違い' }
    '表示制御'                 = @{ Source = '整合性'; Items = @('表示制御の空欄', '表示制御の書き方', '表示制御の過不足', '項目定義との矛盾', 'イベントとの食い違い', '戻す処理なし'); Text = '画面項目の表示・非活性・非表示（表示制御のシート）の空欄・過不足と、画面項目定義・イベント定義との食い違い' }
    '入力項目の型・桁数'       = @{ Source = '画面'; Text = '入力項目の型の空欄と、文字・数値の型の桁数の空欄' }
    '入力チェックのメッセージ' = @{ Source = '画面'; Text = 'メッセージIDの無い入力チェック' }
    '初期値'                   = @{ Source = '画面'; Text = '選択・日付・チェックボックスの入力項目の初期値の空欄' }
    'ボタンのイベント'         = @{ Source = '画面'; Text = 'イベント定義の無いボタン・リンク' }
    'メッセージ欄'             = @{ Source = '画面'; Text = '処理概要とメッセージ欄のメッセージIDの過不足' }
    '主キー'                   = @{ Source = 'テーブル'; Text = '主キーの無いテーブル' }
    'カラムの型・桁数'         = @{ Source = 'テーブル'; Text = 'カラムの型の空欄と、桁数・小数の要る型の桁数・小数の空欄' }
}

function Get-RuleText {
    #     ツールの確認の説明（種類を付けた確認は種類も添える）
    param($Spec)
    $t = $toolRules[$Spec.Name].Text
    if ($Spec.Kind) { $t += '（{0}）' -f $Spec.Kind }
    return $t
}

$script:consistencyRules = $null
function Get-ConsistencyRules {
    #     整合性チェック（Test-DesignConsistency.ps1）の参照のルール
    if ($null -eq $script:consistencyRules) {
        $script:consistencyRules = Read-CsvFile (Join-Path $PSScriptRoot 'consistency/参照のルール.csv') @('区分', '種類', '書き方') '参照のルール'
    }
    return , $script:consistencyRules
}

function Get-ConsistencyKinds {
    #     参照のルールの種類（テーブル・メッセージ・画面など）。ツールの確認の（ ）に書ける
    return @((Get-ConsistencyRules) | ForEach-Object { ([string]$_.'種類').Trim() } | Where-Object { $_ } | Select-Object -Unique)
}

function Get-MessagePattern {
    #     メッセージIDの書き方（参照のルールの、メッセージの参照の書き方。無ければ E0001 の形）
    $rule = @((Get-ConsistencyRules) | Where-Object { ([string]$_.'区分').Trim() -eq '参照' -and ([string]$_.'種類').Trim() -eq 'メッセージ' -and ([string]$_.'書き方').Trim() })[0]
    $pattern = '(?<![A-Za-z0-9])(?<id>[EWIQ]\d{4})(?![0-9])'
    if ($rule) { $pattern = ([string]$rule.'書き方').Trim() }
    return New-Object System.Text.RegularExpressions.Regex($pattern)
}

function Get-MessageIds {
    #     文の中のメッセージID（大文字。出てきた順・重複なし）
    param([string] $Text, [regex] $Rx)
    $ids = New-Object System.Collections.Generic.List[string]
    if (-not $Text) { return , $ids }
    foreach ($m in $Rx.Matches($Text.Normalize($nfkc))) {
        $id = $m.Value
        if ($m.Groups['id'].Success) { $id = $m.Groups['id'].Value }
        $id = $id.ToUpperInvariant()
        if (-not $ids.Contains($id)) { $ids.Add($id) }
    }
    return , $ids
}

function Split-CheckText {
    #     入力チェックの欄を、チェックごとに分ける（「、」「,」「;」・改行で区切る。括弧の中の区切りでは分けない）
    param([string] $Text)
    $parts = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $depth = 0
    foreach ($ch in ([string]$Text).ToCharArray()) {
        if ('（(「【['.IndexOf($ch) -ge 0) { $depth++ }
        elseif ('）)」】]'.IndexOf($ch) -ge 0 -and $depth -gt 0) { $depth-- }
        if (($depth -eq 0 -and '、,，;；'.IndexOf($ch) -ge 0) -or $ch -eq "`n" -or $ch -eq "`r") {
            if ($sb.ToString().Trim()) { $parts.Add($sb.ToString().Trim()) }
            [void]$sb.Clear()
            continue
        }
        [void]$sb.Append($ch)
    }
    if ($sb.ToString().Trim()) { $parts.Add($sb.ToString().Trim()) }
    return , $parts
}

function Add-ScreenRuleFindings {
    #     画面設計書の画面項目定義・イベント定義を確かめる（入力項目の型・桁数・入力チェックのメッセージ・初期値・ボタンのイベント・メッセージ欄）。
    #     Books は ConvertTo-DesignBooks の結果、Names は行う確認、Add は指摘を足すスクリプトブロック
    param($Books, [string[]] $Names, [scriptblock] $Add)
    $design = Read-ScreenDesign $Books
    $msgRx = Get-MessagePattern
    $cellOf = { param($o, [string] $Label) '{0}{1}' -f (ConvertTo-ColumnName $o.Columns[$Label]), $o.Row }
    # 確かめる列が表に無ければ、表ごとに 1 回だけ知らせる
    $missingColumn = @{}
    $needColumn = {
        param($o, [string] $Label, [string] $Rule, [string] $What)
        if ($o.Columns.ContainsKey($Label)) { return $true }
        $k = '{0}|{1}|{2}|{3}' -f $o.Path, $o.Sheet, $o.HeaderRow, $Label
        if (-not $missingColumn.ContainsKey($k)) {
            $missingColumn[$k] = $true
            $first = @($o.Columns.Values | Sort-Object)[0]
            & $Add $Rule $o.Path $o.Sheet ('{0}{1}' -f (ConvertTo-ColumnName $first), $o.HeaderRow) ('{0}の表に「{1}」の列がありません' -f $What, $Label)
        }
        return $false
    }
    # 画面のイベント・画面項目（画面IDが表ごとに違って、その画面のものが無ければ、同じファイルのもの）
    $sameScreen = {
        param($List, $o)
        $mine = @($List | Where-Object { $_.Path -eq $o.Path -and $_.Screen -eq $o.Screen })
        if ($mine.Count -eq 0) { $mine = @($List | Where-Object { $_.Path -eq $o.Path }) }
        return , $mine
    }
    $notInput = 'ボタン|リンク|ラベル'
    $digitType = '文字|英数|英字|数値|数字|カナ|かな|全角|半角|金額|整数|小数'
    $initialKind = '選択|プルダウン|ドロップダウン|コンボ|ラジオ|チェック|日付|カレンダー'
    foreach ($it in $design.Items) {
        $isInput = $it.Input -and $it.Kind -notmatch $notInput
        if ($Names -contains '入力項目の型・桁数' -and $isInput -and (& $needColumn $it '型' '入力項目の型・桁数' '画面項目定義')) {
            if (-not $it.Type) {
                & $Add '入力項目の型・桁数' $it.Path $it.Sheet (& $cellOf $it '型') ('入力項目「{0}」の型が空欄です' -f $it.Item)
            }
            elseif ($it.Type -match $digitType -and -not $it.DigitsText) {
                & $Add '入力項目の型・桁数' $it.Path $it.Sheet (& $cellOf $it '桁数') ('入力項目「{0}」（型: {1}）の桁数が空欄です' -f $it.Item, $it.Type)
            }
        }
        if ($Names -contains '初期値' -and $isInput -and $it.Kind -match $initialKind -and -not $it.Initial -and (& $needColumn $it '初期値' '初期値' '画面項目定義')) {
            & $Add '初期値' $it.Path $it.Sheet (& $cellOf $it '初期値') ('入力項目「{0}」（区分: {1}）の初期値が空欄です。初期値が無いときは「なし」と書いてください' -f $it.Item, $it.Kind)
        }
        if ($Names -contains '入力チェックのメッセージ' -and $it.Checks) {
            $parts = @((Split-CheckText $it.Checks) | Where-Object { $_ -notmatch '^(なし|無し|－|-|ー|―)$' -and -not $msgRx.IsMatch($_.Normalize($nfkc)) })
            if ($parts.Count -gt 0) {
                & $Add '入力チェックのメッセージ' $it.Path $it.Sheet (& $cellOf $it '入力チェック') ('入力項目「{0}」の入力チェック「{1}」に、エラーのときのメッセージIDがありません' -f $it.Item, ($parts -join '」「'))
            }
        }
        if ($Names -contains 'ボタンのイベント' -and $it.Kind -match 'ボタン|リンク') {
            # イベント名・トリガーに項目名（末尾の「ボタン」を除く）がある、またはイベント名が項目名の中にあるイベント
            $name = (ConvertTo-Label $it.Item) -replace '(ボタン|リンク)$', ''
            $hit = $false
            foreach ($e in (& $sameScreen $design.Events $it)) {
                $eventName = (ConvertTo-Label $e.Name) -replace '(ボタン|リンク)$', ''
                $trigger = ConvertTo-Label $e.Trigger
                if (($name -and ($eventName.Contains($name) -or $trigger.Contains($name))) -or ($eventName -and $name.Contains($eventName))) { $hit = $true; break }
            }
            if (-not $hit) {
                & $Add 'ボタンのイベント' $it.Path $it.Sheet (& $cellOf $it '項目名') ('{0}「{1}」のイベントが、イベント定義にありません' -f $it.Kind, $it.Item)
            }
        }
    }
    if ($Names -contains 'メッセージ欄') {
        foreach ($e in $design.Events) {
            if (-not (& $needColumn $e 'メッセージ' 'メッセージ欄' 'イベント定義')) { continue }
            $summary = Get-MessageIds $e.Summary $msgRx
            $column = Get-MessageIds $e.Messages $msgRx
            $missing = @($summary | Where-Object { -not $column.Contains($_) })
            if ($missing.Count -gt 0) {
                & $Add 'メッセージ欄' $e.Path $e.Sheet (& $cellOf $e 'メッセージ') ('イベント「{0}」の処理概要にあるメッセージ {1} が、メッセージ欄にありません' -f $e.Name, ($missing -join '・'))
            }
            # 画面の入力チェックで使うメッセージID
            $checkIds = @{}
            foreach ($it in (& $sameScreen $design.Items $e)) { foreach ($id in (Get-MessageIds $it.Checks $msgRx)) { $checkIds[$id] = $true } }
            $unused = @($column | Where-Object { -not $summary.Contains($_) -and -not $checkIds.ContainsKey($_) })
            if ($unused.Count -gt 0) {
                & $Add 'メッセージ欄' $e.Path $e.Sheet (& $cellOf $e 'メッセージ') ('イベント「{0}」のメッセージ欄の {1} が、処理概要にも画面の入力チェックにもありません' -f $e.Name, ($unused -join '・'))
            }
        }
    }
}

function Test-PrimaryKeyMark {
    #     主キーの欄の印（主キーの順番の数字・○・PK など）があるか
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim().ToUpperInvariant()
    return ($t -match '^[1-9]\d*$' -or $t -match '^(○|◯|〇|●|✓|✔|Y|YES|TRUE|PK)$')
}

function Add-TableRuleFindings {
    #     テーブル定義書のカラム定義を確かめる（主キー・カラムの型・桁数）。見出しと型は Export-TableDdl.ps1 と同じ ddl\ の CSV で決める
    param($Books, [string[]] $Names, [scriptblock] $Add)
    $ddl = Join-Path $PSScriptRoot 'ddl'
    $candidates = Read-HeaderCandidates (Join-Path $ddl 'テーブル定義書の見出し.csv') @('テーブルID', '物理名', 'データ型', '桁数', '小数', '主キー')
    # 桁数・小数の要る型（型の対応で {n}・{s} を使う型）
    $needLength = @{}; $needScale = @{}
    foreach ($r in (Read-CsvFile (Join-Path $ddl '型の対応.csv') @('定義書の型') '型の対応')) {
        $key = (([string]$r.'定義書の型').Normalize($nfkc) -replace '\s', '').ToUpperInvariant()
        $templates = @($r.PSObject.Properties | Where-Object { $_.Name -ne '定義書の型' } | ForEach-Object { [string]$_.Value })
        # […] の中の {n}（日時の秒の小数部の桁数 TIMESTAMP[({n})] など）は省けるので、桁数の要る型にしない
        if (@($templates | Where-Object { ($_ -replace '\[[^\[\]]*\]', '').Contains('{n}') }).Count -gt 0) { $needLength[$key] = $true }
        if (@($templates | Where-Object { $_.Contains('{s}') }).Count -gt 0) { $needScale[$key] = $true }
    }
    foreach ($book in $Books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # カラム定義の見出しの行（物理名とデータ型の見出しがそろっている行）
            $phys = $null; $type = $null
            foreach ($row in $grid.Rows) {
                $labels = @($grid.Labels[$row].Values)
                $p = @($candidates['物理名'] | Where-Object { $labels -contains $_ })[0]
                $t = @($candidates['データ型'] | Where-Object { $labels -contains $_ })[0]
                if ($p -and $t -and $p -ne $t) { $phys = $p; $type = $t; break }
            }
            if (-not $phys) { continue }
            # カラムのある最初の表（Find-DesignTables は表の一覧を 1 つのリストで返す）
            $table = @(foreach ($found in (Find-DesignTables $grid @($phys, $type) $phys)) { if ($found.Rows.Count -gt 0) { $found } })[0]
            if (-not $table) { continue }
            $cof = $table.ColumnOf
            $label = @{}
            foreach ($f in @('桁数', '小数', '主キー')) { $label[$f] = @($candidates[$f] | Where-Object { $cof.ContainsKey($_) })[0] }
            $tableId = $sheetName
            foreach ($n in $candidates['テーブルID']) { $c = Find-LabelCell $grid $n; if ($c -and $c.Value) { $tableId = $c.Value; break } }
            $at = { param([int] $Row, [string] $Label) '{0}{1}' -f (ConvertTo-ColumnName $cof[$Label]), $Row }
            $text = { param($dr, [string] $Label) if ($Label) { ([string]$dr.Cells[$cof[$Label]]).Normalize($nfkc).Trim() } else { '' } }
            if ($Names -contains '主キー') {
                if (-not $label['主キー']) {
                    & $Add '主キー' $book.Path $sheetName (& $at $table.HeaderRow $phys) ('テーブル {0} のカラム定義に、主キー（PK）の列がありません' -f $tableId)
                }
                elseif (@($table.Rows | Where-Object { Test-PrimaryKeyMark (& $text $_ $label['主キー']) }).Count -eq 0) {
                    & $Add '主キー' $book.Path $sheetName (& $at $table.HeaderRow $label['主キー']) ('テーブル {0} に主キーがありません（{1} の欄がどのカラムも空欄です）' -f $tableId, $grid.Cells[$table.HeaderRow][$cof[$label['主キー']]])
                }
            }
            if ($Names -contains 'カラムの型・桁数') {
                foreach ($dr in $table.Rows) {
                    $column = & $text $dr $phys
                    $typeText = & $text $dr $type
                    if (-not $typeText) {
                        & $Add 'カラムの型・桁数' $book.Path $sheetName (& $at $dr.Row $type) ('{0}.{1} のデータ型が空欄です' -f $tableId, $column)
                        continue
                    }
                    # 「VARCHAR(100)」「NUMERIC(11,2)」のように型に桁数を書いた形、桁数の欄の「11,2」も読む
                    $n = ''; $s = ''
                    $m = [regex]::Match($typeText, '^(.+?)\s*\(\s*(\d+)\s*(?:,\s*(\d+)\s*)?\)$')
                    if ($m.Success) { $typeText = $m.Groups[1].Value; $n = $m.Groups[2].Value; $s = $m.Groups[3].Value }
                    $lm = [regex]::Match((& $text $dr $label['桁数']), '^(\d+)(?:\s*[,.]\s*(\d+))?$')
                    if ($lm.Success) { $n = $lm.Groups[1].Value; if ($lm.Groups[2].Success) { $s = $lm.Groups[2].Value } }
                    $scaleText = & $text $dr $label['小数']
                    if ($scaleText -match '^\d+$') { $s = $scaleText }
                    $key = ($typeText -replace '\s', '').ToUpperInvariant()
                    if ($needLength.ContainsKey($key) -and -not $n) {
                        $where = $type
                        if ($label['桁数']) { $where = $label['桁数'] }
                        & $Add 'カラムの型・桁数' $book.Path $sheetName (& $at $dr.Row $where) ('{0}.{1}（{2}）の桁数が空欄です' -f $tableId, $column, $typeText)
                    }
                    elseif ($needScale.ContainsKey($key) -and $label['小数'] -and -not $s) {
                        & $Add 'カラムの型・桁数' $book.Path $sheetName (& $at $dr.Row $label['小数']) ('{0}.{1}（{2}）の小数の桁数が空欄です' -f $tableId, $column, $typeText)
                    }
                }
            }
        }
    }
}

function Invoke-ToolChecks {
    #     観点表の「ツールの確認」を、設計書に対して行う（ツールで確かめる観点で使う確認だけ）。
    #     戻り値: @{ Findings = @(@{ Rule（確認の名前）; Kind（整合性チェックの種類）; Path; Sheet; Cell; Detail }); Failed = 行えなかった確認の名前 → $true }
    param($Points, $Files)
    $result = [pscustomobject]@{ Findings = (New-Object System.Collections.Generic.List[object]); Failed = @{} }
    $sources = [ordered]@{}
    foreach ($p in $Points) {
        if (-not $p.ByTool) { continue }
        foreach ($r in $p.Rules) {
            $src = $toolRules[$r.Name].Source
            if (-not $sources.Contains($src)) { $sources[$src] = New-Object System.Collections.Generic.List[string] }
            if (-not $sources[$src].Contains($r.Name)) { $sources[$src].Add($r.Name) }
        }
    }
    if ($sources.Count -eq 0 -or @($Files).Count -eq 0) { return $result }
    $paths = [string[]]@($Files | ForEach-Object { $_.Path })
    $add = {
        param([string] $Rule, [string] $Path, [string] $Sheet, [string] $Cell, [string] $Detail, [string] $Kind)
        $result.Findings.Add([pscustomobject]@{ Rule = $Rule; Kind = $Kind; Path = $Path; Sheet = $Sheet; Cell = $Cell; Detail = $Detail })
    }
    $books = $null
    foreach ($src in $sources.Keys) {
        $names = $sources[$src]
        Write-Host ('  ツールの確認: {0}' -f ($names -join '・'))
        # ツールの指摘の項目 → 確認の名前
        $itemRule = @{}
        foreach ($n in $names) { foreach ($i in @($toolRules[$n].Items)) { if ($i) { $itemRule[$i] = $n } } }
        try {
            if ($src -eq '提出前チェック') {
                $checks = @()
                if ($names.Contains('版数')) { $checks += 'Version' }
                if ($names.Contains('残っている文言')) { $checks += 'Word' }
                foreach ($i in @(& (Join-Path $PSScriptRoot 'Test-ExcelDeliverable.ps1') -Path $paths -Check $checks 6>$null 3>$null)) {
                    if ($itemRule.ContainsKey($i.Item)) { & $add $itemRule[$i.Item] $i.Path $i.Sheet $i.Cell $i.Detail '' }
                }
            }
            elseif ($src -eq '表記ゆれ') {
                $termArgs = @{ Path = $paths }
                if (@($paths | Where-Object { Test-WordPath $_ }).Count -gt 0) { $termArgs['IncludeWord'] = $true }
                if ($Glossary) { $termArgs['Glossary'] = $Glossary }
                foreach ($v in @(& (Join-Path $PSScriptRoot 'Test-ExcelTerms.ps1') @termArgs 6>$null 3>$null)) {
                    # 揺れの表記ごと・ファイルごとに 1 件（場所は最初に出てくるセル）
                    foreach ($g in @($v.Occurrences | Group-InOrder Path)) {
                        $o = $g.Group[0]
                        & $add '表記ゆれ' $o.Path $o.Sheet $o.Cell ('「{0}」は「{1}」の表記ゆれです（{2}。このファイルで {3} か所）' -f $v.Variant, $v.Preferred, $v.Kind, $g.Count) ''
                    }
                }
            }
            elseif ($src -eq '整合性') {
                foreach ($i in @(& (Join-Path $PSScriptRoot 'Test-DesignConsistency.ps1') -Path $paths 6>$null 3>$null)) {
                    if ($i.Path -and $itemRule.ContainsKey($i.Item)) { & $add $itemRule[$i.Item] $i.Path $i.Sheet $i.Cell $i.Detail $i.Kind }
                }
            }
            else {
                if ($null -eq $books) { $books = ConvertTo-DesignBooks (Select-DesignHits @(& $searchScript -Pattern '\S' -Path $paths -Target Value -SkipHiddenSheets 6>$null 3>$null) $script:designSelection -NoWarning) }
                if ($src -eq '画面') { Add-ScreenRuleFindings $books $names $add }
                else { Add-TableRuleFindings $books $names $add }
            }
        }
        catch {
            Write-Warning ("ツールの確認（{0}）を行えませんでした。この確認を使う観点は「未回答」になります: {1}" -f ($names -join '・'), $_.Exception.Message)
            foreach ($n in $names) { $result.Failed[$n] = $true }
        }
    }
    Write-DesignSkipWarning $script:designSelection
    return $result
}

# ---------------------------------------------------------------- プロンプト


$promptHeader = @'
あなたはシステム開発の設計レビュー担当です。下の【設計書】を【レビュー観点】に沿って確認し、【回答形式】どおりに結果を出力してください。

【ルール】
- 観点ごとに判定してください。判定は OK／NG／要確認 のどれかです。
  - NG：観点を満たしていない箇所がある。箇所ごとに 1 行ずつ出力する。
  - 要確認：設計書の記載だけでは判断できない。理由を書く。
  - OK：問題がない。観点ごとに 1 行だけ出力し、場所・引用・指摘内容・根拠は「-」とする。
- すべての観点について、必ず 1 行以上出力してください。
- 場所は「シート名!セル番地」で書いてください。設計書の各行は「行番号| 列:値」の形で、例えば「12| K:必須」なら K12 です。
- 引用には、場所のセルに書かれている文字をそのまま写してください（長いときは先頭の 30 文字）。空欄のセルを指摘するときは「（空欄）」と書きます。
- 設計書に書かれていないことを推測で補わないでください。
- 【参照資料】は判定の参考にするだけで、指摘の対象にはしないでください。
- 回答形式の行以外（前置き、まとめ、表の見出し）は出力しないでください。

【回答形式】1 行に 1 件、区切りは半角の「|」
観点ID | 判定 | 場所 | 引用 | 指摘内容 | 根拠
（例）
V01 | NG | 画面項目定義!G12 | （空欄） | 桁数が記入されていない | 入力項目は文字・数値の場合に桁数が必要
V02 | NG | イベント定義!J9 | 検索ボタン押下 | 「押下」と「クリック」が混在している | 用語をそろえる
V03 | OK | - | - | - | -
'@
# Word の設計書（や参照資料）を渡すときに足す、場所の書き方
$wordRule = '- Word の設計書は、各行が「段落 12| 文」「表 1 の 3 行目| セル | セル」の形です。場所は「本文!段落 12」「本文!表 1 の 3 行目」のように書いてください。'
# 設計書のテキストの形の説明。Markdown の形（-TextFormat Markdown）で渡すときは、行の形の説明と置き換える
$rowsRule = '設計書の各行は「行番号| 列:値」の形で、例えば「12| K:必須」なら K12 です。'
$markdownRule = '設計書は Markdown で、「## 」の見出しがシート名です。箇条書きの「K12: 必須」はセル K12 です。表は 1 列目が行番号で、見出しの「K:桁数」の K が列です（例えば 12 行目の桁数なら K12）。'
# 回答の形（引用の列があるか）。依頼の一覧に残し、-ImportOnly で前の形の依頼を読むときに使う
$answerFormat = '引用付き'
# プロンプトの改行は OS によらず LF にそろえる。AppendLine や CRLF のスクリプトファイルのままだと
# Windows では 1 行ごとに 1 文字多く数えられ、同じ設計書でも依頼の分割数が変わってしまう
$promptHeader = $promptHeader.Replace("`r`n", "`n")

function New-PromptText {
    #     Word: 設計書か参照資料に Word の文書があるとき（場所の書き方を足す）。
    #     Format: 設計書のテキストの形（Rows か Markdown。-Baseline の変わった部分は、-TextFormat によらず Rows）
    param($Points, [string] $FileName, [string] $Body, [string] $PartNote, [string] $RefText, [switch] $Word, [string] $Format = $TextFormat)

    $sb = New-Object System.Text.StringBuilder
    $header = $promptHeader
    if ($Format -eq 'Markdown') { $header = $header.Replace($rowsRule, $markdownRule) }
    if ($Word) { $header = $header.Replace("`n- 設計書に書かれていないことを", "`n" + $wordRule + "`n- 設計書に書かれていないことを") }
    [void]$sb.Append($header).Append("`n")
    [void]$sb.Append("`n")
    [void]$sb.Append('【レビュー観点】').Append("`n")
    [void]$sb.Append('観点ID | 観点 | 判定基準').Append("`n")
    foreach ($p in $Points) {
        $criteria = $p.Criteria.Replace("`n", ' ')
        # ツールでも確かめる観点は、ツールの分を AI の判定から外す（同じ指摘が二重にならないよう）
        if ($p.ByTool) { $criteria += '（{0}は、ツールで確かめるので指摘しなくてよい）' -f (@($p.Rules | ForEach-Object { Get-RuleText $_ }) -join '、') }
        [void]$sb.Append(('{0} | {1} | {2}' -f $p.Id, $p.Point.Replace("`n", ' '), $criteria)).Append("`n")
    }
    [void]$sb.Append("`n")
    [void]$sb.Append("【設計書】$FileName$PartNote").Append("`n")
    [void]$sb.Append($Body)
    if ($RefText) {
        [void]$sb.Append("`n")
        [void]$sb.Append($RefText)
    }
    return $sb.ToString()
}

function Split-SheetTexts {
    #     シートのテキストを、1 つあたり Budget 文字以内のかたまりに分ける（改行は LF で数える）。
    #     1 シートで Budget を超える場合は行単位で分け、1 行で Budget を超える行は Split-LongLine で分ける。
    #     続きのかたまりに付け直す見出し（シート名・表の見出し）も Budget に数える。
    param($SheetTexts, [int] $Budget)

    $chunks = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    foreach ($name in $SheetTexts.Keys) {
        $text = $SheetTexts[$name]
        if ($current.Length -gt 0 -and $current.Length + $text.Length -gt $Budget) {
            $chunks.Add($current.ToString()); [void]$current.Clear()
        }
        if ($text.Length -le $Budget) { [void]$current.Append($text); continue }

        # 長いシートは行で分割し、続きには見出しを付け直す（Markdown の表の途中で分けたときは、表の見出しも付け直す）。
        # 見出しはシートのテキストの形に合わせる（-Baseline の変わった部分と Word の文書は、-TextFormat Markdown でも行の形）
        $continued = "### シート: $name（続き）"
        if ($text.StartsWith('## ')) { $continued = "## $name（続き）" }
        $prev = ''
        $tableHead = ''
        $top = 0   # かたまりの先頭に付け直した見出しの文字数
        foreach ($line in ($text -split "`r?`n")) {
            if (-not $line) { continue }
            $isSeparator = $line -match '^\|(\s*---\s*\|)+$'
            if (-not $line.StartsWith('|')) { $tableHead = '' }
            elseif ($isSeparator) { $tableHead = $prev + "`n" + $line }
            $rest = $line
            while ($rest) {
                if ($current.Length + $rest.Length + 2 -gt $Budget -and $current.Length -gt $top) {
                    $chunks.Add($current.ToString()); [void]$current.Clear()
                    [void]$current.Append($continued).Append("`n")
                    if ($isSeparator) { [void]$current.Append($prev).Append("`n") }
                    elseif ($tableHead) { [void]$current.Append($tableHead).Append("`n") }
                    $top = $current.Length
                }
                if ($current.Length + $rest.Length + 2 -le $Budget) { [void]$current.Append($rest).Append("`n"); break }
                # 見出しだけのかたまりにも入らない行は、入る分だけを入れ、残りは続きの行にして次のかたまりへ
                $cut = Split-LongLine -Line $rest -Size ($Budget - $current.Length - 2)
                [void]$current.Append($cut[0]).Append("`n")
                $rest = $cut[1]
            }
            $prev = $line
        }
    }
    if ($current.Length -gt 0) { $chunks.Add($current.ToString()) }
    return ,$chunks
}

function Split-LongLine {
    #     1 行を、Size 文字以内の頭と、残り（続きの行）に分ける。セルの区切り（ | ）で分け、1 つのセルが入らなければ
    #     セル内の改行（↵・<br>）か句読点の後ろで分ける（無ければ文字数で分ける）。
    #     続きの行には行の頭（「12| 」「段落 3| 」「| 12 | 」「- 」）を付け直し、分けたセルには列（「D:」「K12: 」）と「（続き）」を付ける。
    #     Markdown の表の行は、列の位置が変わらないよう、頭に入れたセルの分だけ空のセルを置く。戻り値: @(頭, 続きの行)
    param([string] $Line, [int] $Size)

    if ($Line.Length -le $Size) { return , @($Line, '') }
    $m = [regex]::Match($Line, '^(?:\| \d+ \| |- |★?\d+\| |★削除（旧 \d+ 行目）\| |(?:段落|表|テキストボックス) [^|]{0,20}\| )')
    $lead = $m.Value
    $body = $Line.Substring($lead.Length)
    $table = $lead.StartsWith('| ')
    $tail = ''
    if ($table -and $body.EndsWith(' |')) { $tail = ' |'; $body = $body.Substring(0, $body.Length - 2) }
    $cells = @($body -split ' \| ')
    # 続きの行の頭（Markdown の表は、頭に入れた n 個のセルの分の空のセルを足す）
    $nextLead = { param([int] $n) if ($table) { $lead + (' | ' * $n) } else { $lead } }

    $head = $lead
    $n = 0
    while ($n -lt $cells.Count) {
        $add = $cells[$n]
        if ($n -gt 0) { $add = ' | ' + $add }
        if ($head.Length + $add.Length + $tail.Length -gt $Size) { break }
        $head += $add
        $n++
    }
    $cell = $cells[$n]
    $after = ''
    if ($n + 1 -lt $cells.Count) { $after = ' | ' + ($cells[($n + 1)..($cells.Count - 1)] -join ' | ') }
    # 入らなかったセルが、続きの行に丸ごと入るなら、セルの区切りで分ける
    if ($n -gt 0 -and ((& $nextLead $n) + $cell + $tail).Length -le $Size) {
        return , @(($head + $tail), ((& $nextLead $n) + $cell + $after + $tail))
    }
    # セルの中で分ける。列（行の形の「D:」、箇条書きの「K12: 」）は頭に残し、続きにも付ける
    $label = ''
    if ($lead -match '^★?(\d+|削除（旧 \d+ 行目）)\| $') { $label = [regex]::Match($cell, '^[A-Z]{1,3}:').Value }
    elseif ($lead -eq '- ') { $label = [regex]::Match($cell, '^[A-Z]{1,3}\d+: ').Value }
    $sep = ''
    if ($n -gt 0) { $sep = ' | ' }
    $room = $Size - $head.Length - $sep.Length - $tail.Length
    # 少なくとも 10 文字は進める（Size が見出しより小さいときも、分けるたびに短くなるように）
    $least = $label.Length + '（続き）'.Length + 10
    if ($room -lt $least) {
        if ($n -gt 0) { return , @(($head + $tail), ((& $nextLead $n) + $cell + $after + $tail)) }
        $room = $least
    }
    if ($room -ge $cell.Length) {
        $next = ''
        if ($n + 1 -lt $cells.Count) { $next = (& $nextLead ($n + 1)) + ($cells[($n + 1)..($cells.Count - 1)] -join ' | ') + $tail }
        return , @(($head + $sep + $cell + $tail), $next)
    }
    $piece = $cell.Substring(0, $room)
    $at = -1
    foreach ($rx in @('↵|<br\s*/?>', '[。．！？]', '[、，,]|\s')) {
        $found = [regex]::Matches($piece, $rx)
        if ($found.Count -eq 0) { continue }
        $last = $found[$found.Count - 1]
        if ($last.Index + $last.Length -gt $label.Length -and $last.Index + $last.Length -ge $room / 2) { $at = $last.Index + $last.Length; break }
    }
    if ($at -lt 0) {
        $at = $room
        if ([char]::IsHighSurrogate($cell[$at - 1])) { $at-- }
    }
    return , @(($head + $sep + $cell.Substring(0, $at) + $tail), ((& $nextLead $n) + $label + '（続き）' + $cell.Substring($at) + $after + $tail))
}

function Get-TextHash {
    param([string] $Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').Substring(0, 16) }
    finally { $sha.Dispose() }
}

# ---------------------------------------------------------------- 回答の読み取り


function Split-AnswerLine {
    #     回答の 1 行を列に分ける。表をコピーしたタブ区切り（「|」で始まらない行）か、「|」区切り（Markdown の表・コードブロック）。
    #     「\|」は区切りにせず「|」に戻す。列の前後の空白は残す（引用の中の「|」で分かれた列を、元の形に戻せるように）
    param([string] $Line)
    if ($Line.Contains("`t") -and -not $Line.StartsWith('|')) { $fields = $Line.Split("`t") }
    elseif ($Line.Contains('|')) {
        $l = $Line
        if ($l.StartsWith('|')) { $l = $l.Substring(1) }
        if ($l.EndsWith('|') -and -not $l.EndsWith('\|')) { $l = $l.Substring(0, $l.Length - 1) }
        $fields = [regex]::Split($l, '(?<!\\)\|')
    }
    else { return , @() }
    return , @($fields | ForEach-Object { $_.Replace('\|', '|') })
}

function ConvertFrom-Answer {
    #     Copilot の回答から「観点ID | 判定 | 場所 | 引用 | 指摘内容 | 根拠」の行を取り出す。
    #     Markdown の表、コードブロック、表をコピーしたときのタブ区切りのどれでも読める。
    #     Quote を付けないとき（前の形の依頼）と、引用の列の無い 5 列の行は「観点ID | 判定 | 場所 | 指摘内容 | 根拠」として読む。
    #     引用の付いた形で列が 6 より多い行は、引用の中の「|」（Word の表の行・コード値の「1:有効|0:無効」など、設計書から写した文字）で
    #     分かれたとみて、多い列を引用に戻す（指摘内容と根拠は AI が書く文なので「|」はまず入らない）
    param([string] $Text, [string[]] $PointIds, [switch] $Quote)

    $known = @{}
    foreach ($id in $PointIds) { $known[$id.ToUpperInvariant()] = $id }

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('```')) { continue }
        if ($line -match '^\|?[\s:\-|]+\|?$') { continue }   # Markdown 表の区切り行

        $cols = Split-AnswerLine $line
        $fields = @($cols | ForEach-Object { $_.Trim() })
        if ($fields.Count -lt 2) { continue }

        $id = $fields[0].Trim('*', ' ').ToUpperInvariant()
        if (-not $known.ContainsKey($id)) { continue }

        $verdictRaw = $fields[1].Trim('*', ' ')
        if ($verdictRaw -match '^(OK|ＯＫ|○|〇|問題なし)$') { $verdict = 'OK' }
        elseif ($verdictRaw -match '^(NG|ＮＧ|×|指摘)') { $verdict = 'NG' }
        else { $verdict = '要確認' }

        $get = { param($i) if ($fields.Count -gt $i -and $fields[$i] -ne '-' -and $fields[$i] -ne '－') { $fields[$i] } else { '' } }
        # 指摘内容の列（引用の列があれば 1 つ右。列が多ければ、多い分も引用）
        $base = 3
        $quoteText = ''
        if ($Quote -and $fields.Count -ge 6) {
            $base = $fields.Count - 2
            $quoteText = (($cols[3..($base - 1)] -join '|').Trim()).Trim('「', '」', '"', '“', '”')
            if ($quoteText -eq '-' -or $quoteText -eq '－') { $quoteText = '' }
        }
        $reason = & $get ($base + 1)
        if ($fields.Count -gt $base + 2) { $reason = ($fields[($base + 1)..($fields.Count - 1)] -join ' | ') }

        $entries.Add([pscustomobject]@{
            PointId  = $known[$id]
            Verdict  = $verdict
            Location = & $get 2
            Quote    = $quoteText
            Finding  = & $get $base
            Reason   = $reason
        })
    }
    return ,$entries
}

function Resolve-Location {
    #     「シート名!セル」を解釈する。シート名が設計書に無い場合はリンクを付けない。
    param([string] $Location, [string[]] $SheetNames)

    if (-not $Location) { return $null }
    # Word の設計書の「本文!段落 12」「本文!表 1 の 3 行目」「本文!テキストボックス 1」（空白の有無と全角の数字はそろえる）
    $w = [regex]::Match($Location.Normalize($nfkc), '^\s*(?<sheet>[^!]+?)\s*!\s*(?<cell>(?:段落|表|テキストボックス)\s*\d.*?)\s*$')
    if ($w.Success -and $SheetNames -contains $w.Groups['sheet'].Value) {
        $cell = $w.Groups['cell'].Value -replace '\s', ''
        $cell = $cell -replace '^段落(\d+)$', '段落 $1' -replace '^表(\d+)の(\d+)行目$', '表 $1 の $2 行目' -replace '^テキストボックス(\d+)$', 'テキストボックス $1'
        return [pscustomobject]@{ Sheet = $w.Groups['sheet'].Value; Cell = $cell }
    }
    # Excel の「シート名!セル番地」（セル番地の全角の英数字・＄はそろえる。シート名はそのまま比べる）
    $m = [regex]::Match($Location, "^\s*'?(?<sheet>[^!']+?)'?\s*[!！]\s*[`$＄]?(?<col>[A-Za-zＡ-Ｚａ-ｚ]{1,3})[`$＄]?(?<row>[0-9０-９]+)")
    if ($m.Success) {
        $sheet = $m.Groups['sheet'].Value.Trim()
        if ($SheetNames -contains $sheet) {
            return [pscustomobject]@{ Sheet = $sheet; Cell = ($m.Groups['col'].Value + $m.Groups['row'].Value).Normalize($nfkc).ToUpperInvariant() }
        }
        return $null
    }
    # セル番地なしでシート名だけ書かれている場合はシートの先頭へ
    $name = $Location.Trim().Trim("'")
    if ($SheetNames -contains $name) { return [pscustomobject]@{ Sheet = $name; Cell = 'A1' } }
    return $null
}


function ConvertTo-QuoteKey {
    #     引用とセルの文字を比べる形（全角・半角をそろえ、空白・改行・末尾の … を除く）。
    #     改行は、行の形の ↵ も Markdown の形の <br> も除き、Markdown の表の「\|」は「|」に戻す
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Replace('↵', '').Replace('\|', '|') -replace '<br\s*/?>', '' -replace '\s', ''
    return ($t -replace '(…|\.\.\.)$', '')
}

function Read-QuoteCells {
    #     引用を確かめるための、設計書のセルの文字（値・コメント・図形。全シート）。
    #     戻り値: シート名 → セル番地 → @{ Cell; Col; Row; Text; Key（ConvertTo-QuoteKey の形） }
    param([string] $FilePath)
    $sheets = @{}
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value, Comment, Shape -IncludeWord 6>$null 3>$null)
    $seq = 0
    foreach ($h in $hits) {
        $seq++
        if (-not $sheets.ContainsKey($h.Sheet)) { $sheets[$h.Sheet] = @{} }
        $cells = $sheets[$h.Sheet]
        $cell = [string]$h.Cell
        if ($cells.ContainsKey($cell)) { $cells[$cell].Text += [string]$h.Text; continue }
        # Word の段落などは、文書の中の順番を行の代わりにする
        $m = [regex]::Match($cell, '^([A-Z]+)(\d+)$')
        $col = 0; $row = $seq
        if ($m.Success) { $col = Get-ColumnIndex $m.Groups[1].Value; $row = [int]$m.Groups[2].Value }
        $cells[$cell] = [pscustomobject]@{ Cell = $cell; Col = $col; Row = $row; Text = [string]$h.Text; Key = '' }
    }
    foreach ($cells in $sheets.Values) { foreach ($c in $cells.Values) { $c.Key = ConvertTo-QuoteKey $c.Text } }
    return $sheets
}

function Resolve-Quote {
    #     AI の引用を設計書のセル（Read-QuoteCells）と突き合わせる。Location は Resolve-Location の結果（シートが無ければ $null）、
    #     Sheet は AI が書いたシート名。戻り値: @{ State = 一致 / 場所を直した / 見つからない; Sheet; Cell（引用のあるセル）; Note }
    param($Cells, [string] $Quote, $Location, [string] $Sheet)
    $q = ConvertTo-QuoteKey $Quote
    $result = { param([string] $State, [string] $S, [string] $C, [string] $Note) [pscustomobject]@{ State = $State; Sheet = $S; Cell = $C; Note = $Note } }
    $here = $null
    if ($Location -and $Cells.ContainsKey($Location.Sheet) -and $Cells[$Location.Sheet].ContainsKey($Location.Cell)) { $here = $Cells[$Location.Sheet][$Location.Cell] }
    # 空欄の指摘は、場所のセルが空欄かを確かめる
    if ($q -match '^[(（]?空欄[)）]?$') {
        if (-not $Location) { return (& $result '見つからない' '' '' '空欄の場所のシートが設計書にありません') }
        if (-not $here -or -not $here.Key) { return (& $result '一致' $Location.Sheet $Location.Cell '') }
        return (& $result '見つからない' '' '' ('場所のセルは空欄ではありません（{0}）' -f $here.Text))
    }
    if ($here -and $here.Key.Contains($q)) { return (& $result '一致' $Location.Sheet $Location.Cell '') }
    $origin = @(0, 0)
    if ($Location) {
        $m = [regex]::Match($Location.Cell, '^([A-Z]+)(\d+)$')
        if ($here) { $origin = @($here.Col, $here.Row) }
        elseif ($m.Success) { $origin = @((Get-ColumnIndex $m.Groups[1].Value), [int]$m.Groups[2].Value) }
        # 行を引用したとき（同じ行のセルをつなげた文字にある）
        if ($q.Length -ge 2 -and $Cells.ContainsKey($Location.Sheet)) {
            $line = (@($Cells[$Location.Sheet].Values | Where-Object { $_.Row -eq $origin[1] } | Sort-Object Col | ForEach-Object { $_.Key }) -join '')
            if ($line.Contains($q)) { return (& $result '一致' $Location.Sheet $Location.Cell '') }
        }
    }
    # 1 文字の引用（○ など）は、ほかのセルにもよくあるので探さない
    if ($q.Length -lt 2) { return (& $result '見つからない' '' '' '引用が短く、場所のセルにもありません') }
    # 同じシートの、場所にいちばん近いセル。無ければほかのシート（AI が書いたシート名に近いものから）
    $order = @($Cells.Keys | Sort-Object @{ e = { if ($Location -and $_ -eq $Location.Sheet) { 0 } elseif ($Sheet -and ($_.Contains($Sheet) -or $Sheet.Contains($_))) { 1 } else { 2 } } }, @{ e = { $_ } })
    foreach ($s in $order) {
        $best = $null; $bestDistance = [double]::MaxValue
        foreach ($c in $Cells[$s].Values) {
            if (-not $c.Key.Contains($q)) { continue }
            $distance = [math]::Abs($c.Row - $origin[1]) * 1000 + [math]::Abs($c.Col - $origin[0])
            if ($distance -lt $bestDistance -or ($distance -eq $bestDistance -and [string]::CompareOrdinal($c.Cell, $best.Cell) -lt 0)) { $best = $c; $bestDistance = $distance }
        }
        if ($best) { return (& $result '場所を直した' $s $best.Cell '') }
    }
    return (& $result '見つからない' '' '' '引用が設計書に見つかりません（AI の誤った指摘かもしれません）')
}

# ---------------------------------------------------------------- 人が書く列（対応要否など）


# 指摘一覧のうち、人が書く列。レビュー結果を作り直しても消さない
$responseColumns = @('対応要否', '対応状況', '対応者', '対応内容')

function Get-FindingKey {
    #     指摘を見分けるキー（判定元・観点ID・設計書・場所・指摘内容。空白・改行の違いは見ない）
    param([string] $Source, [string] $PointId, [string] $Document, [string] $Location, [string] $Finding)
    return (@($Source, $PointId, $Document, $Location, $Finding) | ForEach-Object { ([string]$_).Replace('↵', '') -replace '\s', '' }) -join '|'
}

function Read-FindingSheet {
    #     レビュー結果の指摘一覧を、1 行ずつ 見出し → 文字 の辞書にする（判定元の列の無い、前の版のツールのレビュー結果は、どれも AI の指摘）
    param([string] $FilePath)
    $list = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $FilePath)) { return , $list }
    $rows = @{}
    foreach ($h in @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value -SheetNamePattern '^指摘一覧$' 6>$null 3>$null)) {
        $m = [regex]::Match([string]$h.Cell, '^([A-Z]+)(\d+)$')
        if (-not $m.Success) { continue }
        $r = [int]$m.Groups[2].Value
        if (-not $rows.ContainsKey($r)) { $rows[$r] = @{} }
        $rows[$r][$m.Groups[1].Value] = ([string]$h.Text).Replace('↵', "`n")
    }
    if (-not $rows.ContainsKey(1)) { return , $list }
    $header = $rows[1]
    foreach ($r in @($rows.Keys | Where-Object { $_ -gt 1 } | Sort-Object)) {
        $o = @{}
        foreach ($c in $rows[$r].Keys) { if ($header.ContainsKey($c)) { $o[$header[$c]] = $rows[$r][$c] } }
        if (@($header.Values) -notcontains '判定元') { $o['判定元'] = 'AI' }
        $list.Add($o)
    }
    return , $list
}

function Read-PreviousResponses {
    #     前のレビュー結果（指摘一覧）の、人が書く列を読む。
    #     戻り値: 指摘のキー（Get-FindingKey）→ 値（列名 → 文字）の待ち行列（同じキーが複数あれば上から順）
    param([string] $FilePath)
    $result = @{}
    foreach ($row in (Read-FindingSheet $FilePath)) {
        $values = @{}
        foreach ($name in $responseColumns) { if ($row[$name]) { $values[$name] = $row[$name] } }
        if ($values.Count -eq 0) { continue }
        $key = Get-FindingKey $row['判定元'] $row['観点ID'] $row['設計書'] $row['場所'] $row['指摘内容']
        if (-not $result.ContainsKey($key)) { $result[$key] = New-Object System.Collections.Generic.Queue[object] }
        $result[$key].Enqueue($values)
    }
    return $result
}

function Get-ResponseKind {
    #     対応要否の値を「要」「不要」に分ける（どちらでもなければ空）。「不要（誤指摘）」のような後ろの（ ）の補足は見ない
    param([string] $Text)
    $t = (([string]$Text).Normalize($nfkc) -replace '\s*\(.*$', '').Trim().ToUpperInvariant()
    if ($t -match '^(要|必要|対応要|要対応|要修正|修正要|対応する|修正する|する|あり|有|○|〇|◯|Y|YES|1)$') { return '要' }
    if ($t -match '^(不要|対応不要|修正不要|対応しない|修正しない|しない|なし|無|×|✕|N|NO|0|対象外|取下げ|取り下げ|誤指摘|重複)$') { return '不要' }
    return ''
}

# ---------------------------------------------------------------- 再確認（-Recheck: 前のレビューの指摘が直ったか）


$recheckHeader = @'
あなたはシステム開発の設計レビュー担当です。下の【前回の指摘】は、前の版の設計書へのレビューの指摘です。
直したあとの【設計書】を見て、指摘ごとに直っているかを判定し、【回答形式】どおりに結果を出力してください。

【ルール】
- 判定は 解消／未解消／要確認 のどれかです。
  - 解消：指摘された問題が直っている。場所・引用は「-」とする。
  - 未解消：問題が残っている。残っている場所と引用を書く。
  - 要確認：設計書の記載だけでは判断できない。理由を書く。
- 【前回の指摘】のすべてについて、必ず 1 行ずつ出力してください。前回の指摘に無い、新しい問題は書かないでください。
- 場所は「シート名!セル番地」で書いてください。設計書の各行は「行番号| 列:値」の形で、例えば「12| K:必須」なら K12 です。
- 引用には、場所のセルに書かれている文字をそのまま写してください（長いときは先頭の 30 文字）。空欄のセルを指摘するときは「（空欄）」と書きます。
- 回答形式の行以外（前置き、まとめ、表の見出し）は出力しないでください。

【回答形式】1 行に 1 件、区切りは半角の「|」
指摘No | 判定 | 場所 | 引用 | 説明
（例）
12 | 解消 | - | - | 桁数が記入された
15 | 未解消 | 画面項目定義!AB14 | （空欄） | 桁数がまだ空欄のまま
'@
$recheckHeader = $recheckHeader.Replace("`r`n", "`n")

function Read-PreviousFindings {
    #     前のレビュー結果（-Recheck）の、NG・要確認の指摘（対応要否が不要のものは除く）
    #     戻り値: @(@{ No; Source（判定元）; Rule（ツールの確認の名前）; PointId; Point; Document（設計書）; Location; Finding })
    param([string] $FilePath)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in (Read-FindingSheet $FilePath)) {
        if (@('NG', '要確認') -notcontains ([string]$row['判定']).Trim()) { continue }
        if ((Get-ResponseKind $row['対応要否']) -eq '不要') { continue }
        $source = ([string]$row['判定元']).Trim()
        $rule = ''
        $m = [regex]::Match($source, '^ツール（(.+)）$')
        if ($m.Success) { $rule = $m.Groups[1].Value }
        $list.Add([pscustomobject]@{
                No = ([string]$row['No']).Trim(); Source = $source; Rule = $rule; PointId = [string]$row['観点ID']; Point = [string]$row['観点']
                Document = [string]$row['設計書']; Location = [string]$row['場所']; Finding = [string]$row['指摘内容']
            })
    }
    return , $list
}

function Find-CurrentFile {
    #     前の指摘の設計書（相対パス）に当たる、今の設計書。版数を除いた相対パスが同じものを先に、無ければ版数を除いたファイル名が
    #     同じもの（1 つだけのとき）。無ければ $null
    param([string] $Document, $Files)
    $key = Get-VersionlessRelative $Document
    $same = @($Files | Where-Object { (Get-VersionlessRelative $_.Relative) -eq $key })
    if ($same.Count -gt 0) { return $same[0] }
    $key = Get-VersionlessName ([System.IO.Path]::GetFileName($Document.Replace('\', '/')))
    $same = @($Files | Where-Object { (Get-VersionlessName $_.Name) -eq $key })
    if ($same.Count -eq 1) { return $same[0] }
    return $null
}

function ConvertTo-ToolShape {
    #     ツールの指摘内容の形。数（か所の数・版数・日付など）は # にして、数だけが違う指摘を同じとみる（「」の中は名前なので変えない）。空白は見ない
    param([string] $Text)
    $t = ([string]$Text).Replace('↵', '') -replace '\s', ''
    return [regex]::Replace($t, '「[^」]*」|\d+(?:\.\d+)*', [System.Text.RegularExpressions.MatchEvaluator] { param($m) if ($m.Value.StartsWith('「')) { $m.Value } else { '#' } })
}

function Find-SameToolFindings {
    #     -Recheck: 前のツールの指摘（Previous）が、今のツールの指摘（Current。Invoke-ToolChecks の Findings）に残っているか。
    #     同じ確認・同じ設計書の指摘のうち、近いものから順に 1 対 1 で当てる。
    #       1. 同じ場所で同じ指摘内容  2. 同じ指摘内容（行を足して場所がずれた）
    #       3. 同じ場所で、数だけが違う指摘内容（版数・か所の数など）  4. 同じ場所で、「」の中（名前・残っている文言など）が同じ指摘内容
    #       5. 数だけが違う指摘内容で、場所の無い指摘か、表記ゆれ（場所は最初に出てくるセルなので、そこを直すと動く）
    #     戻り値: 前の指摘No → 今の指摘
    param($Previous, $Current)
    $key = { param([string] $Text) ([string]$Text).Replace('↵', '') -replace '\s', '' }
    $quotes = { param([string] $Text) (@([regex]::Matches([string]$Text, '「[^」]*」') | ForEach-Object { $_.Value }) -join '') }
    $now = @($Current | ForEach-Object {
            $location = $_.Sheet
            if ($_.Sheet -and $_.Cell) { $location = '{0}!{1}' -f $_.Sheet, $_.Cell }
            [pscustomobject]@{ Finding = $_; Rule = $_.Rule; Path = $_.Path; Location = (& $key $location); Detail = (& $key $_.Detail); Shape = (ConvertTo-ToolShape $_.Detail); Quotes = (& $quotes $_.Detail) }
        })
    $before = @($Previous | ForEach-Object {
            [pscustomobject]@{ No = $_.No; Rule = $_.Rule; Path = $_.Target; Location = (& $key $_.Location); Detail = (& $key $_.Finding); Shape = (ConvertTo-ToolShape $_.Finding); Quotes = (& $quotes $_.Finding) }
        })
    $level = {
        param($x, $y)
        $here = $x.Location -and $x.Location -eq $y.Location
        if ($here -and $x.Detail -eq $y.Detail) { return 5 }
        if ($x.Detail -eq $y.Detail) { return 4 }
        if ($here -and $x.Shape -eq $y.Shape) { return 3 }
        if ($here -and $x.Quotes -eq $y.Quotes) { return 2 }
        if ($x.Shape -eq $y.Shape -and ((-not $x.Location -and -not $y.Location) -or $x.Rule -eq '表記ゆれ')) { return 1 }
        return 0
    }
    $found = @{}
    $used = @{}
    foreach ($min in 5, 4, 3, 2, 1) {
        foreach ($x in $before) {
            if ($found.ContainsKey($x.No)) { continue }
            for ($i = 0; $i -lt $now.Count; $i++) {
                if ($used.ContainsKey($i) -or $now[$i].Rule -ne $x.Rule -or $now[$i].Path -ne $x.Path) { continue }
                if ((& $level $x $now[$i]) -ge $min) { $found[$x.No] = $now[$i].Finding; $used[$i] = $true; break }
            }
        }
    }
    return $found
}

function New-RecheckRequests {
    #     前の指摘（AI の指摘）を、設計書ごとに、指摘のあったシートの全体と一緒に依頼にする（入りきらなければ行で分け、指摘はその行のある部分に入れる）
    param($Findings, $Files)
    $requests = New-Object System.Collections.Generic.List[object]
    $no = 0
    foreach ($f in $Files) {
        $mine = @($Findings | Where-Object { $_.Source -eq 'AI' -and $_.Target -eq $f.Path })
        if ($mine.Count -eq 0) { continue }
        Write-Host ("  テキスト化: {0}" -f $f.Relative)
        $all = ConvertTo-DesignText -Cells (Get-ExcelCells -FilePath $f.Path)
        # 指摘のシート（シートの分からない指摘があれば全シート）
        $wanted = @{}
        foreach ($x in $mine) {
            $loc = Resolve-Location -Location $x.Location -SheetNames @($all.Keys)
            $x | Add-Member -NotePropertyName Sheet -NotePropertyValue $(if ($loc) { $loc.Sheet } else { '' }) -Force
            # 指摘の行（Excel は行番号、Word は「段落 12」などの場所）。分けた部分のうち、この行のあるものに入れる
            $place = ''
            if ($loc -and $loc.Cell -match '^[A-Z]+(\d+)$') { $place = $Matches[1] } elseif ($loc) { $place = $loc.Cell }
            $x | Add-Member -NotePropertyName Place -NotePropertyValue $place -Force
            if ($loc) { $wanted[$loc.Sheet] = $true } else { foreach ($k in $all.Keys) { $wanted[$k] = $true } }
        }
        $texts = [ordered]@{}
        foreach ($k in $all.Keys) { if ($wanted.ContainsKey($k)) { $texts[$k] = $all[$k] } }
        $word = Test-WordPath $f.Path
        $listText = { param($xs) ("【前回の指摘】`n指摘No | 観点 | 場所（前の版） | 指摘内容`n" + ((@($xs | ForEach-Object { '{0} | {1} | {2} | {3}' -f $_.No, ($_.Point -replace '\s+', ' '), $_.Location, ($_.Finding -replace '\s+', ' ') })) -join "`n") + "`n") }
        $build = {
            param($xs, [string] $Body, [string] $PartNote)
            $h = $recheckHeader
            # 設計書の形の説明（Word の文書は -TextFormat によらず段落・表の行の形）
            if ($TextFormat -eq 'Markdown' -and -not $word) { $h = $h.Replace($rowsRule, $markdownRule) }
            if ($word) { $h = $h.Replace("`n- 引用には", "`n" + $wordRule + "`n- 引用には") }
            $h + "`n`n" + (& $listText $xs) + "`n【設計書】" + $f.Name + $PartNote + "`n" + $Body
        }
        $docLength = 0
        foreach ($t in $texts.Values) { $docLength += $t.Length }
        $fixedLen = (& $build $mine '' '（9/9 の一部。この部分に無い指摘は要確認にしてください）').Length
        $budget = [math]::Max(1000, $MaxChars - $fixedLen)
        if ((& $build $mine '' '').Length + $docLength -le $MaxChars) { $chunks = Split-SheetTexts -SheetTexts $texts -Budget ([int]::MaxValue) }
        else { $chunks = Split-SheetTexts -SheetTexts $texts -Budget $budget }
        # 部分ごとの、シート → 行（その部分にある行）。シートの見出しは行の形「### シート: 」か Markdown の「## 」、
        # 行は行の形「12| 」、Word の「段落 12| 」「表 1 の 3 行目| 」、Markdown の表「| 12 | 」・箇条書き「- K12: 」
        $parts = @()
        foreach ($c in $chunks) {
            $rows = @{}; $sheet = ''
            foreach ($line in ($c -split "`n")) {
                $hm = [regex]::Match($line, '^(?:### シート: |## )(.+?)(?:（続き）)?$')
                if ($hm.Success) { $sheet = $hm.Groups[1].Value; if (-not $rows.ContainsKey($sheet)) { $rows[$sheet] = @{} }; continue }
                $rm = [regex]::Match($line, '^(?:(?<p>\d+|段落 \d+|表 \d+ の \d+ 行目|テキストボックス \d+)\||\| (?<p>\d+) \||- [A-Z]{1,3}(?<p>\d+): )')
                if ($rm.Success -and $sheet) { $rows[$sheet][$rm.Groups['p'].Value] = $true }
            }
            $parts += , $rows
        }
        $assigned = @{}
        foreach ($x in $mine) {
            $at = -1
            for ($i = 0; $i -lt $parts.Count; $i++) {
                if (-not $x.Sheet -or -not $parts[$i].ContainsKey($x.Sheet)) { continue }
                if ($at -lt 0) { $at = $i }
                if ($x.Place -and $parts[$i][$x.Sheet].ContainsKey($x.Place)) { $at = $i; break }
            }
            if ($at -lt 0) { $at = 0 }
            if (-not $assigned.ContainsKey($at)) { $assigned[$at] = New-Object System.Collections.Generic.List[object] }
            $assigned[$at].Add($x)
        }
        for ($i = 0; $i -lt $chunks.Count; $i++) {
            if (-not $assigned.ContainsKey($i)) { continue }
            $xs = $assigned[$i]
            $partNote = ''
            if ($chunks.Count -gt 1) { $partNote = '（{0}/{1} の一部。この部分に無い指摘は要確認にしてください）' -f ($i + 1), $chunks.Count }
            $no++
            $prompt = & $build $xs $chunks[$i] $partNote
            $requests.Add([pscustomobject]@{
                    Id = 'R{0:000}' -f $no; File = $f.Path; Relative = $f.Relative; Part = ('{0}/{1}' -f ($i + 1), $chunks.Count)
                    PointIds = @(); FindingNos = @($xs | ForEach-Object { $_.No }); Kind = '再確認'; Sheets = @($all.Keys)
                    Chars = $prompt.Length; Hash = (Get-TextHash $prompt); Prompt = $prompt
                })
        }
    }
    return , $requests
}

function ConvertFrom-RecheckAnswer {
    #     再確認の回答から「指摘No | 判定 | 場所 | 引用 | 説明」の行を取り出す（判定は 解消・未解消・要確認）。
    #     列が 5 より多い行は、ConvertFrom-Answer と同じく、多い列を引用に戻す
    param([string] $Text, [string[]] $Nos)
    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('```') -or $line -match '^\|?[\s:\-|]+\|?$') { continue }
        $cols = Split-AnswerLine $line
        $fields = @($cols | ForEach-Object { $_.Trim() })
        if ($fields.Count -lt 2) { continue }
        $no = ($fields[0].Trim('*', ' ') -replace '^(No\.?|NO\.?|№)\s*', '').Normalize($nfkc)
        if ($Nos -notcontains $no) { continue }
        $v = $fields[1].Trim('*', ' ').Normalize($nfkc)
        if ($v -match '^(未解消|未対応|NG|×|残)') { $verdict = '未解消' }
        elseif ($v -match '^(解消|対応済|OK|○|〇|直)') { $verdict = '解消' }
        else { $verdict = '要確認' }
        $get = { param($i) if ($fields.Count -gt $i -and $fields[$i] -ne '-' -and $fields[$i] -ne '－') { $fields[$i] } else { '' } }
        $quote = & $get 3
        $note = & $get 4
        if ($fields.Count -gt 5) {
            $quote = ($cols[3..($fields.Count - 2)] -join '|').Trim()
            $note = & $get ($fields.Count - 1)
        }
        if ($quote -eq '-' -or $quote -eq '－') { $quote = '' }
        $entries.Add([pscustomobject]@{ No = $no; Verdict = $verdict; Location = (& $get 2); Quote = $quote.Trim('「', '」', '"', '“', '”'); Note = $note })
    }
    return , $entries
}

# ---------------------------------------------------------------- 依頼（プロンプト）の作成


function New-ReviewRequests {
    param($Points, $Files)

    $points = $Points
    $files  = $Files

    # テキスト化は参照資料でも使うので 1 回だけ行う
    $texts = @{}
    $sheetNames = @{}
    $getText = {
        param($f)
        if (-not $texts.ContainsKey($f.Path)) {
            Write-Host ("  テキスト化: {0}" -f $f.Relative)
            $t = ConvertTo-DesignText -Cells (Get-ExcelCells -FilePath $f.Path)
            $texts[$f.Path] = $t
            $sheetNames[$f.Path] = @($t.Keys)
        }
        return $texts[$f.Path]
    }

    $requests = New-Object System.Collections.Generic.List[object]
    $no = 0
    foreach ($f in $files) {
        $targetPoints = @($points | Where-Object { $_.ByAi -and (Test-NameLike $f.Name $_.Targets) })
        if ($targetPoints.Count -eq 0) { continue }
        # -Baseline: 前の版から変わっていない設計書は依頼しない。変わった設計書は、変わった部分だけを渡す
        if ($f.Diff -eq '変更なし') { continue }
        $diffNote = ''
        # 設計書のテキストの形（前の版から変わった部分は、-TextFormat によらず行の形）
        $format = $TextFormat
        if ($f.Diff -eq '変更あり') {
            $format = 'Rows'
            $null = & $getText $f
            $sheetTexts = $f.DiffText
            $diffNote = '（前の版から変わった部分だけです。★の付いた行・図形・コメントだけを判定し、★の無い行は判定の参考にしてください。（旧: …）は前の版の値です。省いた部分にあるかもしれないことは指摘せず、判断できなければ要確認にしてください）'
        }
        else { $sheetTexts = & $getText $f }
        $docLength = 0
        foreach ($t in $sheetTexts.Values) { $docLength += $t.Length }

        $getRefText = {
            param([string[]] $Patterns)
            $sb = New-Object System.Text.StringBuilder
            foreach ($rf in $files) {
                if ($rf.Path -eq $f.Path) { continue }
                if (-not (Test-NameLike $rf.Name $Patterns)) { continue }
                [void]$sb.Append("【参照資料】$($rf.Name)（判定の参考。指摘の対象外）").Append("`n")
                foreach ($t in (& $getText $rf).Values) { [void]$sb.Append($t) }
            }
            if ($sb.Length -eq 0) { Write-Warning ("参照資料「{0}」に一致するファイルがありません" -f ($Patterns -join ', ')) }
            return $sb.ToString()
        }

        # 全観点と全参照資料が 1 回に収まるならまとめて依頼する。
        # 収まらない場合は、参照資料が同じ観点ごとに分ける（参照資料の要らない観点は設計書だけで依頼する）
        $groups = [ordered]@{}
        $allRefs = @($targetPoints | ForEach-Object { $_.Refs } | Sort-Object -Unique)
        # 設計書か参照資料が Word なら、プロンプトに Word の場所の書き方を足す
        $word = (Test-WordPath $f.Path) -or @($files | Where-Object { $_.Path -ne $f.Path -and (Test-WordPath $_.Path) -and (Test-NameLike $_.Name $allRefs) }).Count -gt 0
        $allRefText = ''
        if ($allRefs.Count -gt 0) { $allRefText = & $getRefText $allRefs }
        $allFixed = (New-PromptText -Points $targetPoints -FileName $f.Name -Body '' -PartNote $diffNote -RefText $allRefText -Word:$word -Format $format).Length
        if ($allFixed + $docLength -le $MaxChars) {
            $groups['*'] = [pscustomobject]@{ Points = $targetPoints; RefText = $allRefText }
        }
        else {
            foreach ($p in $targetPoints) {
                $key = ($p.Refs | Sort-Object) -join ','
                if (-not $groups.Contains($key)) {
                    $rt = ''
                    if ($key) { $rt = & $getRefText $p.Refs }
                    $groups[$key] = [pscustomobject]@{ Points = (New-Object System.Collections.Generic.List[object]); RefText = $rt }
                }
                $groups[$key].Points.Add($p)
            }
        }

        foreach ($key in $groups.Keys) {
            $gp = $groups[$key].Points
            $refText = $groups[$key].RefText

            $fixedLen = (New-PromptText -Points $gp -FileName $f.Name -Body '' -PartNote ($diffNote + '（9/9 の一部。他の部分は別途確認するため、この部分だけで判定してください）') -RefText $refText -Word:$word -Format $format).Length
            $budget = $MaxChars - $fixedLen
            if ($budget -lt 1000) {
                Write-Warning ("{0}: 観点と参照資料だけで {1} 文字あり、-MaxChars {2} に収まりません。-MaxChars を増やすか、参照設計書を絞ってください（観点 {3}）" -f $f.Relative, $fixedLen, $MaxChars, (($gp | ForEach-Object { $_.Id }) -join ', '))
                $budget = 1000
            }

            # 1 回に収まるなら分けない（分けたときの「n/m の一部」の注記の分を見込んで、収まるのに分けてしまわないよう）
            $wholeLen = (New-PromptText -Points $gp -FileName $f.Name -Body '' -PartNote $diffNote -RefText $refText -Word:$word -Format $format).Length + $docLength
            if ($wholeLen -le $MaxChars) { $chunks = Split-SheetTexts -SheetTexts $sheetTexts -Budget ([int]::MaxValue) }
            else { $chunks = Split-SheetTexts -SheetTexts $sheetTexts -Budget $budget }
            for ($i = 0; $i -lt $chunks.Count; $i++) {
                $partNote = $diffNote
                if ($chunks.Count -gt 1) {
                    $partNote += '（{0}/{1} の一部。他の部分は別途確認するため、この部分だけで判定してください）' -f ($i + 1), $chunks.Count
                }
                $no++
                $prompt = New-PromptText -Points $gp -FileName $f.Name -Body $chunks[$i] -PartNote $partNote -RefText $refText -Word:$word -Format $format
                $requests.Add([pscustomobject]@{
                    Id         = 'R{0:000}' -f $no
                    File       = $f.Path
                    Relative   = $f.Relative
                    Part       = ('{0}/{1}' -f ($i + 1), $chunks.Count)
                    PointIds   = @($gp | ForEach-Object { $_.Id })
                    Sheets     = $sheetNames[$f.Path]
                    Chars      = $prompt.Length
                    Hash       = Get-TextHash $prompt
                    Prompt     = $prompt
                })
            }
        }
    }
    return ,$requests
}

# ---------------------------------------------------------------- 実行


# ---- 観点表のチェック（どのモードでも最初に行う）

$checklistData = Read-Checklist -FilePath $Checklist
$designFiles = $null
$script:designSelection = $null
if ($Path) {
    $designFiles = Get-DesignFiles
    if ($designFiles.Count -eq 0 -and -not $CheckOnly) { throw '設計書（.xlsx / .xlsm / .xls / .docx / .docm）が見つかりません。' }
}
$hasError = Show-ChecklistIssues -List $checklistData -Issues (Test-Checklist -List $checklistData -Files $designFiles)
if ($CheckOnly) { return }
if ($hasError) { throw '観点表にエラーがあります。上の一覧を見て観点表を直してから、もう一度実行してください。' }
$points = $checklistData.Points

New-Item -ItemType Directory -Force -Path $promptDir, $answerDir | Out-Null

if ($ImportOnly) {
    if (-not (Test-Path -LiteralPath $manifestPath)) { throw "依頼の一覧がありません。先に -PromptOnly などで依頼を作成してください: $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $requests = @($manifest.Requests)
    if ($manifest.PSObject.Properties['Recheck'] -and $manifest.Recheck) { $recheckFindings = @($manifest.Recheck.Findings); $Recheck = $manifest.Recheck.Path }
    # 前の形の依頼（引用の列の無い回答形式）は、引用を読まない
    $quoteFormat = [bool]($manifest.PSObject.Properties['AnswerFormat'] -and $manifest.AnswerFormat -eq $answerFormat)
    if (-not $designFiles) {
        # 依頼を作ったときの設計書（古い依頼の一覧には無いので、依頼のある設計書）
        $designFiles = New-Object System.Collections.Generic.List[object]
        if ($manifest.PSObject.Properties['Files']) {
            foreach ($f in @($manifest.Files)) { $designFiles.Add([pscustomobject]@{ Path = $f.Path; Name = $f.Name; Relative = $f.Relative; Diff = [string]$f.Diff; Baseline = [string]$f.Baseline }) }
        }
        else {
            foreach ($r in $requests) {
                if (@($designFiles | Where-Object { $_.Path -eq $r.File }).Count -eq 0) { $designFiles.Add([pscustomobject]@{ Path = $r.File; Name = [System.IO.Path]::GetFileName($r.File); Relative = $r.Relative }) }
            }
        }
    }
    elseif ($manifest.PSObject.Properties['Files']) {
        # -Path を付けても、依頼を作ったときに前の版と比べた結果（-Baseline の変更なしなど）は、依頼の一覧のものを使う
        $saved = @{}
        foreach ($f in @($manifest.Files)) { $saved[[string]$f.Path] = $f }
        foreach ($f in $designFiles) {
            if (-not $saved.ContainsKey($f.Path)) { continue }
            $f | Add-Member -NotePropertyName Diff -NotePropertyValue ([string]$saved[$f.Path].Diff) -Force
            $f | Add-Member -NotePropertyName Baseline -NotePropertyValue ([string]$saved[$f.Path].Baseline) -Force
        }
    }
}
else {
    $quoteFormat = $true
    if ($Baseline) {
        Write-Host '前の版と比べています...'
        $pairs = Find-BaselineFiles $designFiles
        $compareScript = Join-Path $PSScriptRoot 'Compare-Excel.ps1'
        foreach ($f in $designFiles) {
            $diffState = '新規'; $diffText = $null; $oldPath = ''
            if ($pairs.ContainsKey($f.Path) -and (Test-WordPath $f.Path)) {
                # Word の文書は、変わった部分だけにせず全体を渡す
                $diffState = 'Word（全体）'; $oldPath = $pairs[$f.Path]
            }
            elseif ($pairs.ContainsKey($f.Path)) {
                $oldPath = $pairs[$f.Path]
                # 版数・更新日・改訂履歴だけの変更は、レビューしない
                $diffs = @(& $compareScript -Old $oldPath -New $f.Path -Ignore 版情報 6>$null 3>$null | Where-Object { @('書式', '画像', 'ファイル', '改訂履歴') -notcontains $_.Kind })
                $diffState = '変更なし'
                if ($diffs.Count -gt 0) {
                    $diffText = ConvertTo-DiffText -Cells (Get-ExcelCells -FilePath $f.Path) -Diffs $diffs
                    if ($diffText.Count -gt 0) { $diffState = '変更あり' }
                }
            }
            $f | Add-Member -NotePropertyName Diff -NotePropertyValue $diffState -Force
            $f | Add-Member -NotePropertyName DiffText -NotePropertyValue $diffText -Force
            $f | Add-Member -NotePropertyName Baseline -NotePropertyValue $oldPath -Force
        }
        # Group-Object の並びは PowerShell の版で違うので、決まった順に並べる
        $stateOrder = @('変更あり', '変更なし', '新規')
        $byState = @($designFiles | Group-Object Diff | Sort-Object { $i = [array]::IndexOf($stateOrder, $_.Name); if ($i -lt 0) { $stateOrder.Count } else { $i } }, Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
        Write-Host ("  前の版と比べた設計書: {0}（変更なしの設計書は AI に依頼しません）" -f ($byState -join '・'))
        if ($TextFormat -eq 'Markdown') { Write-Warning '前の版から変わった部分は、行の形（-TextFormat Rows）で渡します（Markdown は、新しい設計書と参照資料に使います）' }
    }
    Write-Host '依頼を作成しています...'
    $recheckFindings = $null
    if ($Recheck) {
        # 前のレビューの NG・要確認の指摘を、今の設計書と対応付けて、直ったかを AI に確かめてもらう（ツールの指摘はツールで確かめ直す）
        $recheckFindings = Read-PreviousFindings $Recheck
        foreach ($x in $recheckFindings) {
            $cur = Find-CurrentFile $x.Document $designFiles
            $x | Add-Member -NotePropertyName Target -NotePropertyValue $(if ($cur) { $cur.Path } else { '' }) -Force
        }
        Write-Host ("  前のレビューの指摘: {0} 件（AI {1} 件・ツール {2} 件。今の設計書が見つからない指摘 {3} 件）" -f $recheckFindings.Count, @($recheckFindings | Where-Object Source -eq 'AI').Count, @($recheckFindings | Where-Object Rule).Count, @($recheckFindings | Where-Object { -not $_.Target }).Count)
        $requests = (New-RecheckRequests -Findings $recheckFindings -Files $designFiles).ToArray()
    }
    else { $requests = (New-ReviewRequests -Points $points -Files $designFiles).ToArray() }
    # 読まなかった雛形のシートを知らせる
    Write-DesignSkipWarning $script:designSelection

    # 前回の依頼と内容が変わったものは、古い回答を使わない
    $oldHash = @{}
    if (Test-Path -LiteralPath $manifestPath) {
        foreach ($o in @((Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json).Requests)) { $oldHash[$o.Id] = $o.Hash }
    }
    foreach ($r in $requests) {
        [System.IO.File]::WriteAllText((Join-Path $promptDir "$($r.Id).txt"), $r.Prompt, $utf8Bom)
        $ans = Join-Path $answerDir "$($r.Id).txt"
        if ((Test-Path -LiteralPath $ans) -and $oldHash[$r.Id] -ne $r.Hash) {
            Move-Item -LiteralPath $ans -Destination ($ans + '.old') -Force
            Write-Warning ("{0} は設計書か観点が変わったため、前回の回答を {0}.txt.old に退避しました。" -f $r.Id)
        }
    }
    # 今回より番号が大きい古い依頼のファイルは片付ける
    Get-ChildItem -LiteralPath $promptDir -Filter 'R*.txt' | Where-Object { @($requests | ForEach-Object { "$($_.Id).txt" }) -notcontains $_.Name } | Remove-Item -Force

    $manifest = [pscustomobject]@{
        Created  = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
        MaxChars = $MaxChars
        # 配列（PointIds・FindingNos）は計算したプロパティにしない（Windows PowerShell 5.1 の ConvertTo-Json は、それを {"value": [...], "Count": n} にしてしまう）。
        # Kind・FindingNos は再確認の依頼だけにあり、ほかの依頼では空になる
        Requests = @($requests | Select-Object Id, File, Relative, Part, PointIds, Sheets, Chars, Hash, Kind, FindingNos)
        AnswerFormat = $answerFormat
        # -Recheck: 前のレビューの指摘（-ImportOnly で再確認の結果を作るため）
        Recheck  = $(if ($Recheck) { [pscustomobject]@{ Path = $Recheck; Findings = @($recheckFindings | Select-Object No, Source, Rule, PointId, Point, Document, Location, Finding, Target) } } else { $null })
        # -ImportOnly でツールの確認をする設計書（-Baseline のときは前の版と比べた結果も）
        Files    = @($designFiles | Select-Object Path, Name, Relative, @{ n = 'Diff'; e = { [string]$_.Diff } }, @{ n = 'Baseline'; e = { [string]$_.Baseline } })
    }
    [System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 5), $utf8Bom)

    Write-Host ''
    if ($requests.Count -eq 0 -and $Recheck) { Write-Host 'AI に確かめてもらう指摘はありません。' }
    elseif ($requests.Count -eq 0) {
        # 依頼の無いわけ（AI に依頼する観点が無い・AI の観点が当たる設計書が無い・その設計書が前の版から変わっていない）
        $aiFiles = @($designFiles | Where-Object { $df = $_; @($points | Where-Object { $_.ByAi -and (Test-NameLike $df.Name $_.Targets) }).Count -gt 0 })
        if (@($points | Where-Object { $_.ByAi }).Count -eq 0) { Write-Host 'AI に依頼する観点はありません（どの観点も、確かめ方がツールです）。' }
        elseif ($aiFiles.Count -eq 0) { Write-Host 'AI に依頼する観点はありません（AI で確かめる観点の対象設計書に当たる設計書がありません）。' }
        elseif (@($aiFiles | Where-Object { [string]$_.Diff -ne '変更なし' }).Count -eq 0) { Write-Host 'AI に依頼する設計書はありません（前の版から変わった設計書がありません）。' }
        else { Write-Host 'AI に依頼するものはありません。' }
    }
    else { Write-Host ("依頼 {0} 件を作成しました（1 件あたり最大 {1} 文字）: {2}" -f $requests.Count, (($requests | Measure-Object Chars -Maximum).Maximum), $promptDir) }
}

$pointById = @{}
foreach ($p in $points) { $pointById[$p.Id] = $p }

# ---- 対話モード: クリップボード経由で Copilot とやり取りする

function Test-KeyInput {
    #     キーの入力を読めるか（コンソールの画面か。PowerShell ISE や、入力をリダイレクトしたときは読めない）
    try { $null = [Console]::KeyAvailable; return $true } catch { return $false }
}

function Read-WatchKey {
    #     押されたキー（s・q・enter）。押されていない・読めないときは空
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Enter) { return 'enter' }
            $c = ([string]$k.KeyChar).ToLowerInvariant()
            if ($c -eq 's' -or $c -eq 'q') { return $c }
        }
    }
    catch { }
    return ''
}

function Test-PromptText {
    #     クリップボードの内容が、空か依頼（プロンプト）のままか（改行の違いは見ない）
    param([string] $Text, [string] $Prompt)
    if (-not $Text -or -not $Text.Trim()) { return $true }
    return ($Text.Replace("`r", '').Trim() -eq $Prompt.Replace("`r", '').Trim()) -or $Text.Contains('【回答形式】')
}

function Import-Answer {
    #     回答を読み取って answers に保存し、件数を知らせる。読み取れなければ $false
    param($Request, [string] $Answer)
    if ($Request.Kind -eq '再確認') {
        $checked = ConvertFrom-RecheckAnswer -Text $Answer -Nos $Request.FindingNos
        if ($checked.Count -eq 0) { return $false }
        [System.IO.File]::WriteAllText((Join-Path $answerDir "$($Request.Id).txt"), $Answer, $utf8Bom)
        Write-Host ("  取り込みました: 解消 {0} 件 / 未解消 {1} 件 / 要確認 {2} 件" -f @($checked | Where-Object Verdict -eq '解消').Count, @($checked | Where-Object Verdict -eq '未解消').Count, @($checked | Where-Object Verdict -eq '要確認').Count)
        return $true
    }
    $entries = ConvertFrom-Answer -Text $Answer -PointIds $Request.PointIds -Quote:$quoteFormat
    if ($entries.Count -eq 0) { return $false }
    [System.IO.File]::WriteAllText((Join-Path $answerDir "$($Request.Id).txt"), $Answer, $utf8Bom)
    $answered = @($entries | ForEach-Object { $_.PointId } | Sort-Object -Unique)
    $missing = @($Request.PointIds | Where-Object { $answered -notcontains $_ })
    Write-Host ("  取り込みました: NG {0} 件 / 要確認 {1} 件 / OK {2} 観点" -f @($entries | Where-Object Verdict -eq 'NG').Count, @($entries | Where-Object Verdict -eq '要確認').Count, @($entries | Where-Object Verdict -eq 'OK').Count)
    if ($missing.Count -gt 0) {
        Write-Warning ("回答に無い観点があります: {0}（結果では「未回答」になります。やり直す場合は answers\{1}.txt を消して再実行）" -f ($missing -join ', '), $Request.Id)
    }
    return $true
}

function Wait-Answer {
    #     クリップボードを見張り、回答がコピーされたら取り込む（Enter は要らない）。
    #     戻り値: 取り込んだら空、s・q キーを押したらその文字
    param($Request, [string] $Prompt)
    $last = $Prompt
    while ($true) {
        $key = Read-WatchKey
        if ($key -eq 's' -or $key -eq 'q') { return $key }
        $answer = [string](Get-Clipboard -Raw)
        if ($answer -ne $last -and -not (Test-PromptText $answer $Prompt)) {
            $last = $answer
            if (Import-Answer $Request $answer) { return '' }
            Write-Warning 'コピーした内容から「観点ID | 判定 | ...」の行を読み取れませんでした。回答全体をコピーし直すか、Copilot に「回答形式どおりに出し直して」と依頼してください（s キー = スキップ / q キー = 中断）。'
        }
        elseif ($key -eq 'enter') {
            Write-Warning 'クリップボードの内容が依頼のままです。Copilot の回答をコピーしてください。'
        }
        Start-Sleep -Milliseconds 500
    }
}

if (-not $PromptOnly -and -not $ImportOnly) {
    if (-not (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) {
        throw 'この環境ではクリップボードが使えません。-PromptOnly で依頼ファイルを作り、回答を answers フォルダに保存してから -ImportOnly を実行してください。'
    }
    # クリップボードを見張るか（Auto は、キーを読める画面なら見張る。読めない画面では s・q を押せないため、Enter で知らせる形にする）
    $watch = $AnswerMode -eq 'Watch' -or ($AnswerMode -eq 'Auto' -and (Test-KeyInput))

    $pending = @($requests | Where-Object { -not (Test-Path -LiteralPath (Join-Path $answerDir "$($_.Id).txt")) })
    Write-Host ''
    Write-Host ("未回答の依頼 {0} 件を順に処理します。" -f $pending.Count)
    if ($watch) {
        Write-Host '手順: Copilot に貼り付けて送信 → 回答をコピー（回答の下のコピー ボタン）すると、自動で取り込んで次の依頼に進みます'
        Write-Host '      s キー = この依頼をスキップ / q キー = 中断（回答済みの分は保存され、次回は続きから）'
    }
    else {
        Write-Host '手順: Copilot に貼り付けて送信 → 回答をコピー（回答の下のコピー ボタン）→ ここで Enter'
        Write-Host '      s = この依頼をスキップ / q = 中断（回答済みの分は保存され、次回は続きから）'
    }

    $i = 0
    foreach ($r in $pending) {
        $i++
        $prompt = [System.IO.File]::ReadAllText((Join-Path $promptDir "$($r.Id).txt"))
        $what = '観点 ' + ($r.PointIds -join ',')
        if ($r.Kind -eq '再確認') { $what = '指摘 ' + ($r.FindingNos -join ',') }
        $header = "[{0}/{1}] {2}  {3}（{4}）{5}  {6} 文字 → クリップボードにコピーしました" -f $i, $pending.Count, $r.Id, $r.Relative, $r.Part, $what, $r.Chars
        if ($watch) {
            Set-Clipboard -Value $prompt
            Write-Host ''
            Write-Host $header -ForegroundColor Cyan
            $key = Wait-Answer $r $prompt
        }
        else {
            while ($true) {
                Set-Clipboard -Value $prompt
                Write-Host ''
                Write-Host $header -ForegroundColor Cyan
                $key = Read-Host '回答をコピーしたら Enter'
                if ($key -eq 'q' -or $key -eq 's') { break }
                $answer = [string](Get-Clipboard -Raw)
                if (Test-PromptText $answer $prompt) {
                    Write-Warning 'クリップボードの内容が依頼のままです。Copilot の回答をコピーしてから Enter を押してください。'
                    continue
                }
                if (Import-Answer $r $answer) { break }
                Write-Warning '回答から「観点ID | 判定 | ...」の行を読み取れませんでした。回答全体をコピーし直すか、Copilot に「回答形式どおりに出し直して」と依頼してください。'
            }
        }
        if ($key -eq 'q') { Write-Host '中断しました。'; break }
    }
}

if ($PromptOnly) {
    Write-Host ''
    Write-Host '次の手順:'
    Write-Host ("  1. {0} の R001.txt から順に、内容を Copilot に貼り付けて送信する" -f $promptDir)
    Write-Host ("  2. 回答をコピーし、{0} に同じ名前（R001.txt など）で UTF-8 のテキストとして保存する" -f $answerDir)
    Write-Host '  3. -ImportOnly を付けて再実行し、レビュー結果を作る'
    return
}

# ---------------------------------------------------------------- 再確認の結果


if ($Recheck) {
    # AI の指摘: 回答から
    $result = @{}
    $requestRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $requests) {
        $ansPath = Join-Path $answerDir "$($r.Id).txt"
        $state = '未回答'
        if (Test-Path -LiteralPath $ansPath) {
            $state = '回答済み'
            foreach ($e in (ConvertFrom-RecheckAnswer -Text ([System.IO.File]::ReadAllText($ansPath)) -Nos @($r.FindingNos))) {
                if (-not $result.ContainsKey($e.No)) { $result[$e.No] = [pscustomobject]@{ Entry = $e; Request = $r } }
            }
        }
        $requestRows.Add(@{ Values = @($r.Id, $r.Relative, $r.Part, (@($r.FindingNos) -join ', '), $r.Chars, $state, "prompts\$($r.Id).txt") })
    }
    # ツールの指摘: 今の設計書をツールで確かめ直し、同じ確認の同じ指摘（数だけが違うものも）が残っているか
    $toolNow = $null
    $sameTool = @{}
    $files = @($designFiles | Where-Object { $_ -and (Test-Path -LiteralPath $_.Path) })
    if (@($recheckFindings | Where-Object { $_.Rule -and $_.Target }).Count -gt 0) {
        Write-Host ''
        Write-Host 'ツールで確かめ直しています...'
        $toolNow = Invoke-ToolChecks -Points @($points | Where-Object { $_.ByTool }) -Files $files
        $sameTool = Find-SameToolFindings @($recheckFindings | Where-Object { $_.Rule -and $_.Target }) $toolNow.Findings
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $counts = [ordered]@{ '解消' = 0; '未解消' = 0; '要確認' = 0; '未回答' = 0; '設計書なし' = 0 }
    $quoteCells = @{}
    foreach ($x in $recheckFindings) {
        $verdict = '未回答'; $location = ''; $quote = ''; $note = ''; $requestId = ''; $loc = $null
        $cur = @($files | Where-Object { $_.Path -eq $x.Target })[0]
        if (-not $cur) { $verdict = '設計書なし'; $note = '今の設計書に、この設計書（版数を除いた相対パス・名前）が見つからないか、同じ名前の設計書が複数あります' }
        elseif ($x.Rule) {
            $requestId = 'ツール'
            if (-not $toolNow -or $toolNow.Failed.ContainsKey($x.Rule)) { $note = 'ツールで確かめられませんでした' }
            else {
                $same = $sameTool[$x.No]
                if ($same) {
                    $verdict = '未解消'
                    $location = $same.Sheet
                    if ($same.Cell) { $location = '{0}!{1}' -f $same.Sheet, $same.Cell }
                    $loc = Resolve-Location -Location $location -SheetNames @($same.Sheet)
                    $note = '機械のチェック（{0}）で、同じ指摘がまだあります' -f $x.Rule
                    # 数などが変わったときは、今の指摘内容も出す
                    if (($same.Detail -replace '\s', '') -ne (([string]$x.Finding).Replace('↵', '') -replace '\s', '')) { $note += '（今: {0}）' -f $same.Detail }
                }
                else { $verdict = '解消'; $note = '機械のチェック（{0}）で、同じ指摘が無くなりました' -f $x.Rule }
            }
        }
        elseif ($result.ContainsKey($x.No)) {
            $hit = $result[$x.No]
            $verdict = $hit.Entry.Verdict; $location = $hit.Entry.Location; $quote = $hit.Entry.Quote; $note = $hit.Entry.Note; $requestId = $hit.Request.Id
            $loc = Resolve-Location -Location $location -SheetNames @($hit.Request.Sheets)
            # 残っている場所の引用を確かめ、違えば引用のあるセルに直す
            if ($verdict -eq '未解消' -and $quote) {
                if (-not $quoteCells.ContainsKey($cur.Path)) { $quoteCells[$cur.Path] = Read-QuoteCells $cur.Path }
                $qr = Resolve-Quote $quoteCells[$cur.Path] $quote $loc ''
                if ($qr.State -eq '場所を直した') { $note = ('{0}（AI の場所 {1} を、引用のあるセルに直しました）' -f $note, $location).Trim(); $location = '{0}!{1}' -f $qr.Sheet, $qr.Cell; $loc = [pscustomobject]@{ Sheet = $qr.Sheet; Cell = $qr.Cell } }
                elseif ($qr.State -eq '見つからない') { $note = ('{0}（{1}）' -f $note, $qr.Note).Trim() }
            }
        }
        else {
            $requestId = @($requests | Where-Object { @($_.FindingNos) -contains $x.No } | ForEach-Object { $_.Id })[0]
        }
        $counts[$verdict]++
        $row = @{ Values = @($x.No, $x.Source, $x.PointId, $x.Point, $x.Document, $x.Location, $x.Finding, $verdict, $location, $quote, $note, [string]$requestId) }
        if ($cur) {
            # リンクは今の設計書の、残っている場所（場所の列）。無ければ前の場所（前の場所の列。Word の文書は開くだけ）
            $column = 8
            if (-not $loc) {
                $column = 5
                $sheetNames = @(@($requests | Where-Object { $_.File -eq $cur.Path } | Select-Object -First 1 | ForEach-Object { $_.Sheets }))
                if ($sheetNames.Count -eq 0) {
                    if (-not $quoteCells.ContainsKey($cur.Path)) { $quoteCells[$cur.Path] = Read-QuoteCells $cur.Path }
                    $sheetNames = @($quoteCells[$cur.Path].Keys)
                }
                $loc = Resolve-Location -Location $x.Location -SheetNames $sheetNames
            }
            $target = ''
            if ($loc -and -not (Test-WordPath $cur.Path)) { $target = "'{0}'!{1}" -f $loc.Sheet.Replace("'", "''"), $loc.Cell }
            elseif (-not $loc) { $column = 4 }
            $row.Link = @{ Column = $column; Path = $cur.Path; Location = $target }
        }
        $rows.Add($row)
    }
    $sheets = @(
        @{
            Name = '再確認'
            Columns = @(
                @{ Header = '前のNo'; Width = 7 }, @{ Header = '判定元'; Width = 12 }, @{ Header = '観点ID'; Width = 8 }, @{ Header = '観点'; Width = 30 },
                @{ Header = '設計書'; Width = 30 }, @{ Header = '前の場所'; Width = 20 }, @{ Header = '前の指摘内容'; Width = 40 }, @{ Header = '結果'; Width = 9 },
                @{ Header = '場所'; Width = 20 }, @{ Header = '引用'; Width = 20 }, @{ Header = '説明'; Width = 40 }, @{ Header = '依頼No'; Width = 8 })
            Rows = $rows
        },
        @{
            Name = '依頼一覧'
            Columns = @(
                @{ Header = '依頼No'; Width = 8 }, @{ Header = '設計書'; Width = 45 }, @{ Header = '分割'; Width = 6 }, @{ Header = '指摘No'; Width = 30 },
                @{ Header = '文字数'; Width = 8 }, @{ Header = '状態'; Width = 9 }, @{ Header = 'プロンプト'; Width = 18 })
            Rows = $requestRows
        }
    )
    try { Export-Workbook -Sheets $sheets -FilePath $reportPath }
    catch { throw ("レビュー結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $reportPath, $_.Exception.Message) }
    Write-Host ''
    Write-Host ("再確認: {0}" -f (@($counts.Keys | Where-Object { $counts[$_] -gt 0 } | ForEach-Object { '{0} {1}' -f $_, $counts[$_] }) -join '・'))
    Write-Host "レビュー結果: $reportPath"
    return
}

# ---------------------------------------------------------------- 結果の集計


$findingRows = New-Object System.Collections.Generic.List[object]
$statusRows  = New-Object System.Collections.Generic.List[object]
$requestRows = New-Object System.Collections.Generic.List[object]

# ---- ツールの確認（確かめ方がツール・両方の観点）
$toolPoints = @($points | Where-Object { $_.ByTool })
$toolFiles = @()
$toolResult = $null
if ($toolPoints.Count -gt 0) {
    $toolFiles = @($designFiles | Where-Object { $_ -and (Test-Path -LiteralPath $_.Path) })
    $gone = @($designFiles | Where-Object { $_ -and -not (Test-Path -LiteralPath $_.Path) })
    if ($gone.Count -gt 0) { Write-Warning ("見つからない設計書は、ツールで確かめません: {0}" -f (($gone | ForEach-Object { $_.Relative }) -join ', ')) }
    Write-Host ''
    Write-Host 'ツールで確かめています...'
    $toolResult = Invoke-ToolChecks -Points $toolPoints -Files $toolFiles
}

# 設計書 × 観点 ごとに、分割された依頼の回答とツールの確認をまとめる
$coverage = [ordered]@{}
$findings = New-Object System.Collections.Generic.List[object]
$newCoverage = {
    param([string] $File, [string] $Relative, [string] $PointId)
    $key = "$File|$PointId"
    if (-not $coverage.Contains($key)) {
        $coverage[$key] = [pscustomobject]@{ File = $File; Relative = $Relative; PointId = $PointId; Verdicts = New-Object System.Collections.Generic.List[string]; Requests = New-Object System.Collections.Generic.List[string]; Missing = $false }
    }
    return $coverage[$key]
}
$droppedPoints = New-Object System.Collections.Generic.List[string]
foreach ($r in $requests) {
    $ansPath = Join-Path $answerDir "$($r.Id).txt"
    $state = '未回答'
    $entries = @()
    if (Test-Path -LiteralPath $ansPath) {
        $entries = ConvertFrom-Answer -Text ([System.IO.File]::ReadAllText($ansPath)) -PointIds $r.PointIds -Quote:$quoteFormat
        $state = '回答済み'
    }
    $requestRows.Add(@{ Values = @($r.Id, $r.Relative, $r.Part, ($r.PointIds -join ', '), $r.Chars, $state, "prompts\$($r.Id).txt") })

    foreach ($pointId in $r.PointIds) {
        # 依頼を作ったあとで観点表から無くした・無効にした観点は、回答があっても入れない
        if (-not $pointById.ContainsKey($pointId)) {
            if (-not $droppedPoints.Contains($pointId)) { $droppedPoints.Add($pointId) }
            continue
        }
        $cv = & $newCoverage $r.File $r.Relative $pointId
        $cv.Requests.Add($r.Id)
        $mine = @($entries | Where-Object PointId -eq $pointId)
        if ($mine.Count -eq 0) { $cv.Missing = $true }
        foreach ($e in $mine) {
            $cv.Verdicts.Add($e.Verdict)
            if ($e.Verdict -ne 'OK') {
                $findings.Add([pscustomobject]@{ Request = $r; Entry = $e; Source = 'AI'; Seq = $findings.Count })
            }
        }
    }
}

if ($droppedPoints.Count -gt 0) { Write-Warning ("観点表で有効でない観点の回答は、レビュー結果に入れません: {0}" -f ($droppedPoints -join ', ')) }

# ツールの指摘は、その確認を書いた観点のうち、観点表で先にあって設計書が対象のものに入れる。
# 対象の観点が無ければ（設計書どうしの食い違いを、一覧ではなく設計書の側で見つけたときなど）、その確認を書いた観点に入れる
$toolCounts = [ordered]@{}
if ($toolResult) {
    $relativeOf = @{}
    foreach ($f in $toolFiles) { $relativeOf[$f.Path] = $f.Relative }
    $byKey = @{}
    foreach ($tf in $toolResult.Findings) {
        if (-not $relativeOf.ContainsKey($tf.Path)) { continue }
        $name = [System.IO.Path]::GetFileName($tf.Path)
        $owners = @($toolPoints | Where-Object { @($_.Rules | Where-Object { $_.Name -eq $tf.Rule -and (-not $_.Kind -or $_.Kind -eq $tf.Kind) }).Count -gt 0 })
        $owner = @($owners | Where-Object { Test-NameLike $name $_.Targets })[0]
        if (-not $owner) { $owner = $owners[0] }
        if (-not $owner) { continue }
        $key = "$($tf.Path)|$($owner.Id)"
        if (-not $byKey.ContainsKey($key)) { $byKey[$key] = New-Object System.Collections.Generic.List[object] }
        $byKey[$key].Add($tf)
        $location = $tf.Sheet
        if ($tf.Sheet -and $tf.Cell) { $location = '{0}!{1}' -f $tf.Sheet, $tf.Cell }
        $request = [pscustomobject]@{ Id = 'ツール'; File = $tf.Path; Relative = $relativeOf[$tf.Path]; Sheets = @($tf.Sheet) }
        $entry = [pscustomobject]@{ PointId = $owner.Id; Verdict = 'NG'; Location = $location; Finding = $tf.Detail; Reason = ('機械のチェック（{0}）' -f $tf.Rule) }
        $findings.Add([pscustomobject]@{ Request = $request; Entry = $entry; Source = ('ツール（{0}）' -f $tf.Rule); Seq = $findings.Count })
        if (-not $toolCounts.Contains($tf.Rule)) { $toolCounts[$tf.Rule] = 0 }
        $toolCounts[$tf.Rule]++
    }
    foreach ($f in $toolFiles) {
        foreach ($p in $toolPoints) {
            # （変更なしの設計書でも、ツールでは確かめる）
            if (-not (Test-NameLike $f.Name $p.Targets)) { continue }
            $cv = & $newCoverage $f.Path $f.Relative $p.Id
            $cv.Requests.Add('ツール')
            $key = "$($f.Path)|$($p.Id)"
            if ($byKey.ContainsKey($key)) { foreach ($tf in $byKey[$key]) { $cv.Verdicts.Add('NG') } }
            elseif (@($p.Rules | Where-Object { $toolResult.Failed.ContainsKey($_.Name) }).Count -gt 0) { $cv.Missing = $true }
            else { $cv.Verdicts.Add('OK') }
        }
    }
    $countText = @($toolCounts.Keys | ForEach-Object { '{0} {1}' -f $_, $toolCounts[$_] }) -join '・'
    if ($countText) { Write-Host ("ツールの確認: 指摘 {0} 件（{1}）" -f @($findings | Where-Object { $_.Source -ne 'AI' }).Count, $countText) }
    else { Write-Host 'ツールの確認: 指摘はありません' }
}

# 設計書ごとに重要度の高い順。同じ観点では AI の指摘、ツールの指摘の順（読んだ順）
$sorted = $findings | Sort-Object @{ e = { $_.Request.Relative } }, @{ e = { $s = $pointById[$_.Entry.PointId].Severity; if ($severityOrder.ContainsKey($s)) { $severityOrder[$s] } else { 9 } } }, @{ e = { $_.Entry.PointId } }, @{ e = { $_.Seq } }
# 前のレビュー結果に人が書いた対応要否などは、同じ指摘に引き継ぐ
$previous = Read-PreviousResponses $reportPath
$keptCount = 0
# 観点ID → @{ Ai; Yes; No; Tool }（AI の指摘の採用率）
$adoption = @{}
# AI の指摘の引用を設計書のセルと突き合わせ、場所が違えば引用のあるセルに直す
$quoteCells = @{}
$quoteCounts = [ordered]@{}
$n = 0
foreach ($f in $sorted) {
    $n++
    $p = $pointById[$f.Entry.PointId]
    $location = $f.Entry.Location
    $loc = Resolve-Location -Location $location -SheetNames @($f.Request.Sheets)
    $note = ''
    $quote = ''; $check = ''
    if ($f.Source -eq 'AI' -and $quoteFormat) {
        $quote = [string]$f.Entry.Quote
        if (-not $quote) { $check = '引用なし' }
        elseif (-not (Test-Path -LiteralPath $f.Request.File)) { $check = '確かめていない'; $note = '設計書が見つからないため、引用を確かめていません' }
        else {
            if (-not $quoteCells.ContainsKey($f.Request.File)) { $quoteCells[$f.Request.File] = Read-QuoteCells $f.Request.File }
            $aiSheet = ''
            $sm = [regex]::Match([string]$location, "^\s*'?(?<sheet>[^!'！]+?)'?\s*[!！]")
            if ($sm.Success) { $aiSheet = $sm.Groups['sheet'].Value.Trim() }
            $qr = Resolve-Quote $quoteCells[$f.Request.File] $quote $loc $aiSheet
            $check = $qr.State
            if ($qr.State -eq '場所を直した') {
                $note = 'AI の場所: {0}（引用のあるセルに直しました）' -f $location
                $location = '{0}!{1}' -f $qr.Sheet, $qr.Cell
                $loc = [pscustomobject]@{ Sheet = $qr.Sheet; Cell = $qr.Cell }
            }
            elseif ($qr.State -eq '見つからない') { $note = $qr.Note }
        }
        if (-not $quoteCounts.Contains($check)) { $quoteCounts[$check] = 0 }
        $quoteCounts[$check]++
    }
    if ($location -and -not $loc -and -not $note) { $note = '場所のシート名・セル番地を設計書で確認してください' }
    if ($loc -and (Test-WordPath $f.Request.File)) { $location = '{0}!{1}' -f $loc.Sheet, $loc.Cell }
    $kept = @('', '', '', '')
    $key = Get-FindingKey $f.Source $p.Id $f.Request.Relative $location $f.Entry.Finding
    if ($previous.ContainsKey($key) -and $previous[$key].Count -gt 0) {
        $v = $previous[$key].Dequeue()
        $kept = @($responseColumns | ForEach-Object { if ($v.ContainsKey($_)) { $v[$_] } else { '' } })
        $keptCount++
    }
    if (-not $adoption.ContainsKey($p.Id)) { $adoption[$p.Id] = @{ Ai = 0; Yes = 0; No = 0; Tool = 0 } }
    if ($f.Source -eq 'AI') {
        $adoption[$p.Id].Ai++
        $kind = Get-ResponseKind $kept[0]
        if ($kind -eq '要') { $adoption[$p.Id].Yes++ } elseif ($kind -eq '不要') { $adoption[$p.Id].No++ }
    }
    else { $adoption[$p.Id].Tool++ }
    $row = @{ Values = @($n, $f.Entry.Verdict, $f.Source, $p.Severity, $p.Id, $p.Category, $p.Point, $f.Request.Relative, $location, $quote, $check, $f.Entry.Finding, $f.Entry.Reason, $f.Request.Id, $kept[0], $kept[1], $kept[2], $kept[3], $note) }
    if ($loc -and (Test-WordPath $f.Request.File)) {
        # Word の文書は開くだけ（セルの場所が無い）
        $row.Link = @{ Column = 8; Path = $f.Request.File; Location = '' }
    }
    elseif ($loc) {
        $row.Link = @{ Column = 8; Path = $f.Request.File; Location = "'{0}'!{1}" -f $loc.Sheet.Replace("'", "''"), $loc.Cell }
    }
    elseif ($f.Source -ne 'AI') {
        # 場所の無いツールの指摘（ファイル名の版数など）は、設計書を開く
        $row.Link = @{ Column = 7; Path = $f.Request.File; Location = '' }
    }
    $findingRows.Add($row)
}

# -Baseline で前の版から変わっていない設計書の、AI に依頼する観点（依頼していない）
foreach ($f in @($designFiles | Where-Object { $_ -and $_.PSObject.Properties['Diff'] -and $_.Diff -eq '変更なし' })) {
    foreach ($p in @($points | Where-Object { $_.ByAi -and (Test-NameLike $f.Name $_.Targets) })) {
        $cv = & $newCoverage $f.Path $f.Relative $p.Id
        $cv.Requests.Add('変更なし')
    }
}

# 設計書の順、観点表の順に並べる
$fileOrder = @{}
foreach ($f in $designFiles) { if ($f -and -not $fileOrder.ContainsKey($f.Path)) { $fileOrder[$f.Path] = $fileOrder.Count } }
foreach ($r in $requests) { if (-not $fileOrder.ContainsKey($r.File)) { $fileOrder[$r.File] = $fileOrder.Count } }
$pointOrder = @{}
foreach ($p in $points) { $pointOrder[$p.Id] = $pointOrder.Count }
$firstSheet = @{}
foreach ($r in $requests) { if (-not $firstSheet.ContainsKey($r.File) -and @($r.Sheets).Count -gt 0) { $firstSheet[$r.File] = @($r.Sheets)[0] } }
foreach ($cv in @($coverage.Values | Sort-Object @{ e = { $fileOrder[$_.File] } }, @{ e = { $pointOrder[$_.PointId] } })) {
    $p = $pointById[$cv.PointId]
    $ng = @($cv.Verdicts | Where-Object { $_ -eq 'NG' }).Count
    if ($ng -gt 0) { $verdict = 'NG' }
    elseif ($cv.Verdicts -contains '要確認') { $verdict = '要確認' }
    elseif ($cv.Missing) { $verdict = '未回答' }
    elseif ($cv.Verdicts.Count -eq 0 -and $cv.Requests -contains '変更なし') { $verdict = '変更なし' }
    else { $verdict = 'OK' }
    # 分割した依頼の一部だけ未回答なら、OK と言い切らない
    if ($cv.Missing -and $verdict -eq 'OK') { $verdict = '未回答' }
    # 依頼の無い設計書（ツールだけで確かめた設計書）は、ファイルを開くだけ
    $location = ''
    if ($firstSheet.ContainsKey($cv.File) -and -not (Test-WordPath $cv.File)) { $location = "'{0}'!A1" -f $firstSheet[$cv.File].Replace("'", "''") }
    $statusRows.Add(@{ Values = @($cv.Relative, $p.Id, $p.Category, $p.Point, $p.Severity, $verdict, $ng, ($cv.Requests -join ', ')); Link = @{ Column = 0; Path = $cv.File; Location = $location } })
}

# 観点ごとの、AI の指摘の採用率（人が書いた対応要否から。不要が多い観点は判定基準を見直す）
$adoptionRows = New-Object System.Collections.Generic.List[object]
foreach ($p in $points) {
    if (-not $adoption.ContainsKey($p.Id)) { continue }
    $a = $adoption[$p.Id]
    $decided = $a.Yes + $a.No
    $rate = ''
    if ($decided -gt 0) { $rate = [math]::Round($a.Yes / $decided, 3) }
    $hint = ''
    if ($a.Ai -gt 0 -and $decided -eq 0) { $hint = '対応要否（要・不要）を書くと、採用率が出ます' }
    elseif ($decided -gt 0 -and $a.No -ge 2 -and $a.Yes / $decided -lt 0.5) { $hint = '不要が多い観点です。判定基準を具体的にするか、確かめ方をツールにできないか見直してください' }
    $adoptionRows.Add(@{ Values = @($p.Id, $p.Category, $p.Point, $a.Ai, $a.Yes, $a.No, ($a.Ai - $decided), $rate, $a.Tool, $hint) })
}

$sheets = @(
    @{
        Name = '指摘一覧'
        Columns = @(
            @{ Header = 'No'; Width = 5 }, @{ Header = '判定'; Width = 8 }, @{ Header = '判定元'; Width = 12 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '観点ID'; Width = 8 },
            @{ Header = '分類'; Width = 10 }, @{ Header = '観点'; Width = 30 }, @{ Header = '設計書'; Width = 30 }, @{ Header = '場所'; Width = 20 },
            @{ Header = '引用'; Width = 24 }, @{ Header = '引用の確認'; Width = 11 }, @{ Header = '指摘内容'; Width = 45 }, @{ Header = '根拠'; Width = 35 }, @{ Header = '依頼No'; Width = 8 },
            @{ Header = '対応要否'; Width = 9 }, @{ Header = '対応状況'; Width = 9 }, @{ Header = '対応者'; Width = 10 }, @{ Header = '対応内容'; Width = 30 }, @{ Header = '備考'; Width = 25 })
        Rows = $findingRows
    },
    @{
        Name = '観点消化状況'
        Columns = @(
            @{ Header = '設計書'; Width = 40 }, @{ Header = '観点ID'; Width = 8 }, @{ Header = '分類'; Width = 10 }, @{ Header = '観点'; Width = 50 },
            @{ Header = '重要度'; Width = 7 }, @{ Header = '判定'; Width = 8 }, @{ Header = 'NG件数'; Width = 8 }, @{ Header = '依頼No'; Width = 14 })
        Rows = $statusRows
    },
    @{
        Name = '依頼一覧'
        Columns = @(
            @{ Header = '依頼No'; Width = 8 }, @{ Header = '設計書'; Width = 45 }, @{ Header = '分割'; Width = 6 }, @{ Header = '観点ID'; Width = 40 },
            @{ Header = '文字数'; Width = 8 }, @{ Header = '状態'; Width = 9 }, @{ Header = 'プロンプト'; Width = 18 })
        Rows = $requestRows
    },
    @{
        Name = '観点別の採用率'
        Columns = @(
            @{ Header = '観点ID'; Width = 8 }, @{ Header = '分類'; Width = 10 }, @{ Header = '観点'; Width = 40 }, @{ Header = 'AI の指摘'; Width = 9 },
            @{ Header = '要'; Width = 6 }, @{ Header = '不要'; Width = 6 }, @{ Header = '未記入'; Width = 7 }, @{ Header = '採用率'; Width = 8 },
            @{ Header = 'ツールの指摘'; Width = 10 }, @{ Header = '見方'; Width = 50 })
        Rows = $adoptionRows
    }
)

try {
    Export-Workbook -Sheets $sheets -FilePath $reportPath
}
catch {
    throw ("レビュー結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $reportPath, $_.Exception.Message)
}

$answeredCount = @($requestRows | Where-Object { $_.Values[5] -eq '回答済み' }).Count
$statusCounts = $statusRows | Group-InOrder { $_.Values[5] } | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
$toolNote = ''
if ($toolPoints.Count -gt 0) { $toolNote = '（うちツール {0} 件）' -f @($findings | Where-Object { $_.Source -ne 'AI' }).Count }
Write-Host ''
Write-Host ("依頼 {0} 件中 {1} 件回答済み / 観点の判定: {2} / 指摘 {3} 件{4}" -f $requests.Count, $answeredCount, ($statusCounts -join '・'), $findingRows.Count, $toolNote)
if ($quoteCounts.Count -gt 0) {
    Write-Host ("AI の指摘の引用: {0}" -f (@($quoteCounts.Keys | ForEach-Object { '{0} {1}' -f $_, $quoteCounts[$_] }) -join '・'))
    if ($quoteCounts.Contains('見つからない')) { Write-Host '  引用が見つからない指摘は、AI の誤りかもしれません。指摘一覧の「引用の確認」で絞り込んで確かめてください' }
}
if ($keptCount -gt 0) { Write-Host ("前のレビュー結果の対応要否・対応状況・対応者・対応内容を {0} 件引き継ぎました" -f $keptCount) }
$yes = 0; $no = 0
foreach ($a in $adoption.Values) { $yes += $a.Yes; $no += $a.No }
if ($yes + $no -gt 0) {
    Write-Host ("AI の指摘の採用率（対応要否から）: 要 {0}・不要 {1}（{2:0}%）" -f $yes, $no, ($yes * 100.0 / ($yes + $no)))
    $low = @($adoptionRows | Where-Object { $_.Values[9] -like '不要が多い*' } | ForEach-Object { '{0} {1}/{2}' -f $_.Values[0], $_.Values[4], ($_.Values[4] + $_.Values[5]) })
    if ($low.Count -gt 0) { Write-Host ("  不要が多い観点: {0}（観点別の採用率のシート）" -f ($low -join '・')) }
}
Write-Host "レビュー結果: $reportPath"

# .SYNOPSIS
#     Excel の設計書を、観点表に沿って AI（Microsoft 365 Copilot など）にレビューさせ、指摘一覧にまとめます。
#
# .DESCRIPTION
#     1. 設計書（Excel 方眼紙）をセル番地付きのテキストに変換し、観点表の観点と合わせて
#        Copilot に貼り付けるプロンプトを作ります。入力文字数の上限（-MaxChars）を超える設計書は
#        シート単位で分割します。
#     2. 対話モード（既定）では、プロンプトをクリップボードに入れて待ちます。Copilot に貼り付けて送信し、
#        回答をコピーすると、回答を読み取って保存し、次の依頼に進みます（Enter は要りません。-AnswerMode）。
#     3. 全回答から「レビュー結果.xlsx」（指摘一覧・観点消化状況・依頼一覧・観点別の採用率）を作ります。
#        指摘の場所は、クリックすると設計書の該当セルが開くリンクになります。
#        AI には指摘ごとに場所のセルの文字（引用）も書かせ、設計書と突き合わせます。場所が違えば引用のあるセルに直し、
#        引用が設計書に無ければ「見つからない」と知らせます（指摘一覧の「引用の確認」）。
#
#     観点表の「確かめ方」を「ツール」にした観点は AI に依頼せず、「ツールの確認」に書いた確認を
#     ツール（提出前チェック・表記ゆれ・整合性チェックと、画面項目定義・テーブル定義の確認）で行います。
#     「両方」の観点は AI にも依頼します（ツールで確かめる分は、AI には指摘しなくてよいと伝えます）。
#     ツールの指摘は、指摘一覧に判定元「ツール（確認の名前）」で入ります。数える・突き合わせるだけで決まる観点は、
#     AI より漏れが無く、依頼の数も減ります。ツールの確認（「ツールの確認」の欄に書く名前）:
#       版数                     表紙・改訂履歴・ファイル名の版数と更新日の食い違い（Test-ExcelDeliverable.ps1）
#       残っている文言           TODO・TBD・未定・検討中などの消し忘れ（Test-ExcelDeliverable.ps1）
#       表記ゆれ                 カタカナの長音・全角半角の揺れと、-Glossary の用語集の揺れ（Test-ExcelTerms.ps1）
#       参照先なし・重複定義・一覧との違い・ヘッダの違い・組み合わせの違い・属性の違い
#                                設計書どうしの食い違い（Test-DesignConsistency.ps1）。「参照先なし（メッセージ）」のように、
#                                （ ）で種類（参照のルールの種類）を付けると、その種類だけにします
#       表示制御                 画面項目の表示・非活性・非表示（表示制御のシート）の空欄・過不足と、画面項目定義・
#                                イベント定義との食い違い（Test-DesignConsistency.ps1）
#       入力項目の型・桁数       入力項目の型の空欄と、文字・数値の型の桁数の空欄
#       入力チェックのメッセージ メッセージIDの無い入力チェック
#       初期値                   選択・日付・チェックボックスの入力項目の初期値の空欄
#       ボタンのイベント         イベント定義の無いボタン・リンク
#       メッセージ欄             イベントの処理概要とメッセージ欄のメッセージIDの過不足
#       主キー                   主キーの無いテーブル
#       カラムの型・桁数         カラムの型の空欄と、桁数・小数の要る型（tools\ddl\型の対応.csv）の桁数・小数の空欄
#
#     途中で中断しても、回答済みの分は保存され、次回は続きから再開します。
#     レビュー結果を作り直しても、指摘一覧に人が書いた対応要否・対応状況・対応者・対応内容は、同じ指摘に引き継ぎます。
#     対応要否（要・不要）からは、観点ごとの AI の指摘の採用率を出します（不要が多い観点は、判定基準を見直す目安）。
#     設計書の読み取りには同じフォルダの Search-Excel.ps1 を使います（Excel 不要）。
#
# .PARAMETER Path
#     レビューする設計書のフォルダまたはファイル（複数指定可）。Excel（.xlsx / .xlsm / .xls）と Word（.docx / .docm）の設計書を読みます。
#     同じ文書の古い版・ツールが書き出したファイル・書き換える前のファイル（-BackupDir）と、非表示のシート・雛形のシート（ID が {{テーブルID}} などのまま）は読みません（読まなかったものは警告で知らせます）。
#
# .PARAMETER Checklist
#     観点表（xlsx または UTF-8 の csv）。省略時は review-checklists\詳細設計レビュー観点.xlsx。
#     列: 観点ID / 分類 / 観点 / 判定基準 / 重要度 / 対象設計書 / 参照設計書 / 有効 / 確かめ方 / ツールの確認
#     確かめ方（AI・ツール・両方）が空欄なら、ツールの確認があれば両方、無ければ AI です。
#
# .PARAMETER OutDir
#     プロンプト・回答・レビュー結果の保存先。省略時は、1 つ目の設計書のフォルダと同じ場所の design-review
#     （D:\設計書 なら D:\design-review。今いるフォルダによらない）。設計書を指定しないとき（-ImportOnly だけなど）は .\design-review
#
# .PARAMETER MaxChars
#     1 回のプロンプトの最大文字数。Copilot の入力上限に合わせてください（既定 8000）。
#
# .PARAMETER PromptOnly
#     プロンプトのファイルを作るだけで終了します（クリップボードを使えない環境向け）。
#
# .PARAMETER CheckOnly
#     観点表のチェックだけを行って終了します。観点表を編集したあとの確認に使います。
#     Path を指定すると、対象設計書・参照設計書のパターンが実際の設計書に当たるかも確認します。
#     （チェックは通常の実行時にも毎回行い、エラーがあれば中断します）
#
# .PARAMETER ImportOnly
#     answers フォルダに保存した回答から、レビュー結果だけを作り直します。
#
# .PARAMETER SheetNameExclude
#     テキスト化しないシート名の正規表現（例: '表紙|改訂履歴'）。文字数の節約に使います。
#
# .PARAMETER Glossary
#     表記ゆれ（ツールの確認）で使う用語集の CSV（正しい表記・揺れ）。Test-ExcelTerms.ps1 の -Glossary と同じです。
#
# .PARAMETER Baseline
#     前の版の設計書のフォルダ（またはファイル）。版数を除いた相対パス（無ければファイル名）で今の設計書と対応付け（Compare-Excel.ps1 で比べ）、
#     AI には前の版から変わった行（★。前の版の値を「旧: …」で添える）と、その節の見出しの行だけを渡します。
#     変わっていない設計書は AI に依頼しません（観点消化状況は「変更なし」）。版数・更新日・改訂履歴だけの変更は除きます。
#     ツールの確認は、変わっていない設計書も含めて行います。
#
# .PARAMETER AnswerMode
#     対話モードで回答を受け取る方法。
#       Auto（既定） キーを読める画面（PowerShell のコンソール）では Watch、読めない画面（PowerShell ISE など）では Enter
#       Watch        クリップボードを見張り、回答がコピーされたら取り込んで次の依頼に進む。s キーでスキップ、q キーで中断
#       Enter        回答をコピーして Enter を押す（s・q を入れて Enter でスキップ・中断）
#
# .PARAMETER Recheck
#     前のレビュー結果（レビュー結果.xlsx か、それのあるフォルダ）。設計書を直したあとに、前の NG・要確認の指摘
#     （対応要否が不要のものを除く）が直ったかを確かめます。AI の指摘は、指摘のあったシートと一緒に AI に渡して
#     解消／未解消／要確認を判定させ、ツールの指摘はツールで確かめ直します。結果は「再確認」のシートに出します。
#     設計書は、版数を除いた相対パス（無ければファイル名）で前の指摘と対応付けます。-OutDir には前のレビューとは別のフォルダを指定します。
#     -Baseline とは同時に使えません。
#
# .PARAMETER TextFormat
#     設計書を AI に渡すときのテキストの形。Rows（既定）は「行番号| 列:値 | 列:値」の行の形、Markdown は
#     Export-ExcelToMarkdown.ps1 と同じ形（項目定義などの表は Markdown の表、「■」の行は小見出し、それ以外は箇条書き）に、
#     行番号・列・セル番地を付けたものです。Markdown は表の列と見出しの対応が AI に伝わりやすい代わりに、文字数が 1 割ほど増えます。
#
# .PARAMETER NoRecurse
#     サブフォルダの設計書を対象にしません。
#
# .EXAMPLE
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計
#
# .EXAMPLE
#     # 入力上限が 16,000 文字のテナントで、表紙と改訂履歴を除いて依頼数を減らす
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -MaxChars 16000 -SheetNameExclude '表紙'
#
# .EXAMPLE
#     # 観点表を編集したあと、書き方の誤りがないか確かめる
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -CheckOnly
#
# .EXAMPLE
#     # 設計書を直したあと、前のレビューの指摘が直ったかを確かめる
#     .\Invoke-DesignReview.ps1 D:\Design\v1.1 -Recheck D:\design-review\レビュー結果.xlsx -OutDir D:\design-review-再確認
#
# .EXAMPLE
#     # プロンプトをファイルで受け渡す
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -PromptOnly
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -ImportOnly
