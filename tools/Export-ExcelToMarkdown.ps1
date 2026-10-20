[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $OutDir,

    [switch] $Force,

    [switch] $CellAddress,

    [switch] $TextConv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if ($OutDir) { $OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir) }
if ($TextConv) {
    # git の textconv から呼ばれるとき: 標準出力には Markdown だけを書く（進捗の表示も出さない）
    if ($OutDir) { throw '-TextConv と -OutDir は同時に指定できません' }
    $ProgressPreference = 'SilentlyContinue'
}

$extensions = @('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb')

# ---------------------------------------------------------------- 変換するファイル


$files = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) {
            $root = $item.FullName.TrimEnd('\', '/')
            $found = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue)
        }
        else {
            $root = [System.IO.Path]::GetDirectoryName($item.FullName)
            $found = @($item)
        }
        foreach ($f in $found) {
            if ($extensions -notcontains $f.Extension.ToLowerInvariant() -or $f.Name.StartsWith('~$') -or $seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
            $files.Add([pscustomobject]@{ Info = $f; Relative = $rel; Target = $null })
        }
    }
}
if ($TextConv -and $files.Count -ne 1) { throw '-TextConv には Excel ファイルを 1 つ指定してください' }

# -OutDir: 元のフォルダ構成のまま .md を作る。Excel より新しい .md があるファイルは変換しない（-Force で変換する）
$todo = New-Object System.Collections.Generic.List[object]
$skipped = 0
foreach ($f in $files) {
    if ($OutDir) {
        $relDir = [System.IO.Path]::GetDirectoryName($f.Relative)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Info.Name) + '.md'
        if ($relDir) { $f.Target = Join-Path (Join-Path $OutDir $relDir) $name } else { $f.Target = Join-Path $OutDir $name }
        if (-not $Force -and (Test-Path -LiteralPath $f.Target) -and (Get-Item -LiteralPath $f.Target).LastWriteTimeUtc -ge $f.Info.LastWriteTimeUtc) { $skipped++; continue }
    }
    $todo.Add($f)
}

# ---------------------------------------------------------------- Markdown にする


function ConvertTo-MarkdownText {
    #     セルの文字を Markdown の 1 行に収める（改行は <br>、表の中の | はエスケープ）
    param([string] $Text, [switch] $InTable)
    # Search-Excel はセル内の改行を ↵ で表す
    $t = ([string]$Text).Trim() -replace "\r\n|\r|\n|↵", '<br>'
    if ($InTable) { $t = $t.Replace('|', '\|') }
    return $t
}

function ConvertTo-SheetMarkdown {
    #     1 シートの値を Markdown にする。
    #       見出しの行（3 つ以上の値）の下に、見出しの列だけを使う行が続くところは表に、それ以外の行は箇条書きに、
    #       「■」で始まる行は小見出しにする
    param($Rows, [System.Collections.Generic.List[string]] $Lines)
    $rowNums = [int[]]@($Rows.Keys)
    [Array]::Sort($rowNums)
    $last = ''
    $i = 0
    while ($i -lt $rowNums.Count) {
        $cells = $Rows[$rowNums[$i]]
        if ($cells[0].Text.TrimStart().StartsWith('■')) {
            if ($last) { $Lines.Add('') }
            $Lines.Add('### ' + (ConvertTo-MarkdownText (@($cells | ForEach-Object { $_.Text.Trim() }) -join ' ')))
            $Lines.Add('')
            $last = ''
            $i++
            continue
        }
        $headerCols = @{}
        foreach ($c in $cells) { $headerCols[$c.Col] = $true }
        $j = $i + 1
        if ($cells.Count -ge 3) {
            while ($j -lt $rowNums.Count -and $rowNums[$j] -eq $rowNums[$j - 1] + 1) {
                $next = $Rows[$rowNums[$j]]
                if ($next[0].Text.TrimStart().StartsWith('■')) { break }
                $inside = $true
                foreach ($c in $next) { if (-not $headerCols.ContainsKey($c.Col)) { $inside = $false; break } }
                if (-not $inside) { break }
                $j++
            }
        }
        if ($j - $i -ge 2) {
            if ($last) { $Lines.Add('') }
            $cols = @($cells | ForEach-Object { $_.Col })
            $head = @($cells | ForEach-Object { ConvertTo-MarkdownText $_.Text -InTable })
            $sep = @($cols | ForEach-Object { '---' })
            if ($CellAddress) { $head = @('行') + $head; $sep = @('---') + $sep }
            $Lines.Add('| ' + ($head -join ' | ') + ' |')
            $Lines.Add('| ' + ($sep -join ' | ') + ' |')
            for ($k = $i + 1; $k -lt $j; $k++) {
                $byCol = @{}
                foreach ($c in $Rows[$rowNums[$k]]) { $byCol[$c.Col] = $c.Text }
                $vals = @($cols | ForEach-Object { ConvertTo-MarkdownText $byCol[$_] -InTable })
                if ($CellAddress) { $vals = @([string]$rowNums[$k]) + $vals }
                $Lines.Add('| ' + ($vals -join ' | ') + ' |')
            }
            $Lines.Add('')
            $last = ''
            $i = $j
            continue
        }
        if ($last -ne 'list' -and $Lines.Count -gt 0 -and $Lines[$Lines.Count - 1] -ne '') { $Lines.Add('') }
        if ($CellAddress) { $parts = @($cells | ForEach-Object { '{0}{1}: {2}' -f (ConvertTo-ColumnName $_.Col), $rowNums[$i], (ConvertTo-MarkdownText $_.Text) }) }
        else { $parts = @($cells | ForEach-Object { ConvertTo-MarkdownText $_.Text }) }
        $Lines.Add('- ' + ($parts -join ' | '))
        $last = 'list'
        $i++
    }
}

function ConvertTo-BookMarkdown {
    param([string] $Title, $Hits)
    $lines = New-Object System.Collections.Generic.List[string]
    if ($Title) { $lines.Add('# ' + $Title); $lines.Add('') }
    # シート → 行番号 → セル（列の順）。図形・コメントはシートごとに後ろへまとめる
    $sheets = [ordered]@{}
    foreach ($h in $Hits) {
        if (-not $sheets.Contains($h.Sheet)) { $sheets[$h.Sheet] = [pscustomobject]@{ Rows = @{}; Shapes = (New-Object System.Collections.Generic.List[object]); Comments = (New-Object System.Collections.Generic.List[object]) } }
        $s = $sheets[$h.Sheet]
        switch ($h.Kind) {
            '値' {
                $m = [regex]::Match([string]$h.Cell, '^([A-Z]+)(\d+)$')
                if (-not $m.Success) { break }
                $row = [int]$m.Groups[2].Value
                if (-not $s.Rows.ContainsKey($row)) { $s.Rows[$row] = New-Object System.Collections.Generic.List[object] }
                $s.Rows[$row].Add([pscustomobject]@{ Col = (Get-ColumnIndex $m.Groups[1].Value); Text = [string]$h.Text })
            }
            '図形' { $s.Shapes.Add($h) }
            'コメント' { $s.Comments.Add($h) }
        }
    }
    foreach ($name in $sheets.Keys) {
        $s = $sheets[$name]
        $lines.Add('## ' + $name)
        $lines.Add('')
        if ($s.Rows.Count -eq 0 -and $s.Shapes.Count -eq 0 -and $s.Comments.Count -eq 0) { $lines.Add('（空）'); $lines.Add(''); continue }
        foreach ($row in @($s.Rows.Keys)) {
            $sorted = @($s.Rows[$row] | Sort-Object Col)
            $s.Rows[$row] = $sorted
        }
        ConvertTo-SheetMarkdown $s.Rows $lines
        if ($lines[$lines.Count - 1] -ne '') { $lines.Add('') }
        if ($s.Shapes.Count -gt 0) {
            $lines.Add('### 図形')
            $lines.Add('')
            foreach ($h in $s.Shapes) {
                $label = $h.Name
                if ($CellAddress) { $label = '{0}（{1} 付近）' -f $h.Name, $h.Cell }
                $lines.Add(('- {0}: {1}' -f $label, (ConvertTo-MarkdownText $h.Text)))
            }
            $lines.Add('')
        }
        if ($s.Comments.Count -gt 0) {
            $lines.Add('### コメント')
            $lines.Add('')
            foreach ($h in $s.Comments) { $lines.Add(('- {0}: {1}' -f $h.Cell, (ConvertTo-MarkdownText $h.Text))) }
            $lines.Add('')
        }
    }
    while ($lines.Count -gt 0 -and $lines[$lines.Count - 1] -eq '') { $lines.RemoveAt($lines.Count - 1) }
    return (($lines -join "`n") + "`n")
}

$converted = 0
if ($todo.Count -gt 0) {
    $searchArgs = @{
        Pattern = '\S'
        Path = @($todo | ForEach-Object { [System.Management.Automation.WildcardPattern]::Escape($_.Info.FullName) })
        Target = @('Value', 'Comment', 'Shape', 'SheetName')
    }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null)
    $byPath = @{}
    foreach ($h in $hits) {
        if (-not $byPath.ContainsKey($h.Path)) { $byPath[$h.Path] = New-Object System.Collections.Generic.List[object] }
        $byPath[$h.Path].Add($h)
    }
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    foreach ($f in $todo) {
        $bookHits = $byPath[$f.Info.FullName]
        if (-not $bookHits) { continue }   # 読めなかったファイル（Search-Excel が警告済み）
        if ($TextConv) {
            # 一時ファイルの名前は毎回変わるので、見出しにファイル名を入れない
            $md = ConvertTo-BookMarkdown '' $bookHits
            $bytes = $utf8.GetBytes($md)
            $stdout = [Console]::OpenStandardOutput()
            $stdout.Write($bytes, 0, $bytes.Length)
            $stdout.Flush()
            return
        }
        $md = ConvertTo-BookMarkdown $f.Relative.Replace('\', '/') $bookHits
        if ($OutDir) {
            $dir = [System.IO.Path]::GetDirectoryName($f.Target)
            if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            [System.IO.File]::WriteAllText($f.Target, $md, $utf8)
        }
        else { $md }
        $converted++
    }
}

if ($OutDir) {
    Write-Host ''
    Write-Host ('{0} ファイルを Markdown にしました（Excel より新しい .md があり変換しなかったもの {1} ファイル）: {2}' -f $converted, $skipped, $OutDir)
}

# .SYNOPSIS
#     Excel（設計書など）を Markdown のテキストにします（Excel 不要）。Git での差分の表示や、AI への入力、全文検索に使えます。
#
# .DESCRIPTION
#     シートごとに、セルの値を次のように Markdown にします。
#
#       表         見出しの行（3 つ以上の値）の下に、見出しの列だけを使う行が続くところ（項目定義・カラム定義など）
#       小見出し   「■」で始まる行
#       箇条書き   それ以外の行（1 行を「値 | 値 | …」の 1 行に）
#       図形・コメント  シートの最後にまとめる
#
#     結合したセルは 1 つの値に、セル内の改行は <br> にします。既定ではセル番地を書かないので、行を挿入しても
#     差分はその行だけになります（-CellAddress で行番号・セル番地を付けられます）。
#
# .PARAMETER Path
#     Excel のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER OutDir
#     Markdown（.md）を保存するフォルダ。元のフォルダ構成のまま保存し、Excel より新しい .md があるファイルは変換しません。
#     指定しないときは Markdown の文字列を出力します。
#
# .PARAMETER Force
#     -OutDir に新しい .md があっても変換し直します。
#
# .PARAMETER CellAddress
#     表に行番号、箇条書きにセル番地、図形に位置を付けます（AI に場所を答えさせたいときなど）。
#
# .PARAMETER TextConv
#     git の textconv 用。1 つのファイルの Markdown を UTF-8 で標準出力に書きます（見出しにファイル名を入れません）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は変換しなくなります）。
#
# .EXAMPLE
#     .\Export-ExcelToMarkdown.ps1 D:\設計書 -OutDir D:\設計書_md
#
# .EXAMPLE
#     # git diff で Excel の変更を文字の差分で見る（.gitattributes に「*.xlsx diff=excel」を書いておく）
#     git config diff.excel.textconv "powershell -NoProfile -ExecutionPolicy Bypass -File C:/tools/Export-ExcelToMarkdown.ps1 -TextConv"
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
