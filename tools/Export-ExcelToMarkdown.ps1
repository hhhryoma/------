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

# ---------------------------------------------------------------- Markdown にする（ExcelTools.psm1 の ConvertTo-BookMarkdown）


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
            $md = ConvertTo-BookMarkdown '' $bookHits -CellAddress:$CellAddress
            $bytes = $utf8.GetBytes($md)
            $stdout = [Console]::OpenStandardOutput()
            $stdout.Write($bytes, 0, $bytes.Length)
            $stdout.Flush()
            return
        }
        $md = ConvertTo-BookMarkdown $f.Relative.Replace('\', '/') $bookHits -CellAddress:$CellAddress
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
