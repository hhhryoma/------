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
    $items = @(Get-PathItem $p)
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

# -OutDir: 元のフォルダ構成のまま .md を作る。Excel より新しい .md があるファイルは変換しない（-Force で変換する）。
# 同じ .md になる Excel（同じフォルダの 設計書.xlsx と 設計書.xlsm など）は、重ならないよう名前に拡張子も含める（設計書.xlsx.md）
$todo = New-Object System.Collections.Generic.List[object]
$skipped = 0
$sameName = @{}
if ($OutDir) {
    foreach ($f in $files) {
        $relDir = [System.IO.Path]::GetDirectoryName($f.Relative)
        $name = [System.IO.Path]::GetFileNameWithoutExtension($f.Info.Name) + '.md'
        if ($relDir) { $f.Target = Join-Path (Join-Path $OutDir $relDir) $name } else { $f.Target = Join-Path $OutDir $name }
        $key = $f.Target.ToUpperInvariant()
        $sameName[$key] = 1 + [int]$sameName[$key]
    }
}
foreach ($f in $files) {
    if ($OutDir) {
        if ($sameName[$f.Target.ToUpperInvariant()] -gt 1) {
            $f.Target = Join-Path ([System.IO.Path]::GetDirectoryName($f.Target)) ($f.Info.Name + '.md')
        }
        if (-not $Force -and (Test-Path -LiteralPath $f.Target) -and (Get-Item -LiteralPath $f.Target).LastWriteTimeUtc -ge $f.Info.LastWriteTimeUtc) { $skipped++; continue }
    }
    $todo.Add($f)
}

# ---------------------------------------------------------------- 画像


function Get-BookPictures {
    #     ブックの画像（シート名 → Name・Text（代替テキスト）・Cell・Width・Height・Hash・Hidden・Link）。.xlsx / .xlsm だけ
    #     （.xls などと、読めないブック（パスワード付きなど）は空）。
    #     AssetDir があれば、画像のファイルをそこに書き出して Link（.md からの相対パス AssetLink/ファイル名）を付け、
    #     前に書き出した画像のうち、今のブックに無いものは消す（画像が 1 つも無ければフォルダも消す）
    param([string] $FilePath, [string] $AssetDir, [string] $AssetLink)
    $result = @{}
    $drawings = $null
    try { $drawings = Read-BookDrawings $FilePath } catch { $drawings = $null }
    $written = @{}   # 画像のパーツ名 → 書き出したファイル名
    $names = @{}
    if ($drawings) {
        $pkg = $null
        try {
            foreach ($sheet in $drawings.Keys) {
                $list = New-Object System.Collections.Generic.List[object]
                foreach ($it in $drawings[$sheet]) {
                    if ($it.Kind -ne 'pic') { continue }
                    $link = ''
                    if ($AssetDir -and $it.Media) {
                        if (-not $written.ContainsKey($it.Media)) {
                            $written[$it.Media] = ''
                            if ($null -eq $pkg) { $pkg = Open-ZipPackage $FilePath }
                            $entry = Get-PackageEntry -Package $pkg -PartName $it.Media
                            if ($entry) {
                                # 名前はパーツの名前（image1.png など）。違うフォルダの同じ名前のパーツは番号を付けて分ける
                                $name = [System.IO.Path]::GetFileName($it.Media)
                                if ($names.ContainsKey($name.ToUpperInvariant())) { $name = '{0}_{1}' -f ($names.Count + 1), $name }
                                $names[$name.ToUpperInvariant()] = $true
                                if (-not (Test-Path -LiteralPath $AssetDir)) { New-Item -ItemType Directory -Force -Path $AssetDir | Out-Null }
                                $in = $entry.Open()
                                $out = [System.IO.File]::Create((Join-Path $AssetDir $name))
                                try { $in.CopyTo($out) } finally { $out.Dispose(); $in.Dispose() }
                                $written[$it.Media] = $name
                            }
                        }
                        if ($written[$it.Media]) { $link = $AssetLink + '/' + $written[$it.Media] }
                    }
                    $list.Add([pscustomobject]@{ Name = $it.Name; Text = $it.Descr; Cell = $it.Cell; Width = $it.Width; Height = $it.Height; Hash = $it.Hash; Hidden = $it.Hidden; Link = $link })
                }
                # 配列で渡す（New-Object で作った List を @() で囲むと、PowerShell 7.4 では止まることがある）
                if ($list.Count -gt 0) { $result[$sheet] = $list.ToArray() }
            }
        }
        finally { if ($pkg) { $pkg.Zip.Dispose() } }
    }
    if ($AssetDir -and (Test-Path -LiteralPath $AssetDir)) {
        foreach ($old in @(Get-ChildItem -LiteralPath $AssetDir -File)) { if (-not $names.ContainsKey($old.Name.ToUpperInvariant())) { Remove-Item -LiteralPath $old.FullName -Force } }
        if (@(Get-ChildItem -LiteralPath $AssetDir -Force).Count -eq 0) { Remove-Item -LiteralPath $AssetDir -Force }
    }
    return $result
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
        # 画像は一覧にする。-OutDir なら、.md と同じ名前の .assets フォルダ（設計書.assets）に画像のファイルも書き出してリンクする
        $assetDir = ''; $assetLink = ''
        if ($OutDir) {
            $assetLink = [System.IO.Path]::GetFileNameWithoutExtension($f.Target) + '.assets'
            $assetDir = Join-Path ([System.IO.Path]::GetDirectoryName($f.Target)) $assetLink
        }
        $pictures = Get-BookPictures $f.Info.FullName $assetDir $assetLink
        if ($TextConv) {
            # 一時ファイルの名前は毎回変わるので、見出しにファイル名を入れない
            $md = ConvertTo-BookMarkdown '' $bookHits -CellAddress:$CellAddress -Pictures $pictures
            $bytes = $utf8.GetBytes($md)
            $stdout = [Console]::OpenStandardOutput()
            $stdout.Write($bytes, 0, $bytes.Length)
            $stdout.Flush()
            return
        }
        $md = ConvertTo-BookMarkdown $f.Relative.Replace('\', '/') $bookHits -CellAddress:$CellAddress -Pictures $pictures
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
#       表         見出しの行（3 つ以上の値。「予定」の下に「開始」「終了」がある 2 行の見出しも）の下に、見出しの列を使う行が
#                  続くところ（項目定義・カラム定義など）。見出しの列より右の備考などの値は、見出しの空いた列にします。
#                  No などの番号が続く行は、表の途中の空の行（2 行まで）をはさんでも同じ表にします
#       小見出し   「■」で始まる行
#       箇条書き   それ以外の行（1 行を「値 | 値 | …」の 1 行に）
#       図形・画像・コメント  シートの最後にまとめる。画像は名前・代替テキストと、大きさ・中身の目印（差し替えると変わる）
#
#     結合したセルは 1 つの値に、セル内の改行は <br> にします。既定ではセル番地を書かないので、行を挿入しても
#     差分はその行だけになります（-CellAddress で行番号・セル番地を付けられます）。
#
# .PARAMETER Path
#     Excel のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER OutDir
#     Markdown（.md）を保存するフォルダ。元のフォルダ構成のまま保存し、Excel より新しい .md があるファイルは変換しません。
#     同じフォルダの、拡張子だけが違う同じ名前の Excel（設計書.xlsx と 設計書.xlsm など）を一緒に変換するときは、
#     名前に拡張子も含めます（設計書.xlsx.md・設計書.xlsm.md）。
#     画像（.xlsx / .xlsm）は、.md と同じ名前の .assets フォルダ（設計書.assets）に書き出し、Markdown からリンクします
#     （変換し直すと、ブックから無くなった画像のファイルは消します）。
#     指定しないときは Markdown の文字列を出力します。
#
# .PARAMETER Force
#     -OutDir に新しい .md があっても変換し直します。
#
# .PARAMETER CellAddress
#     表に行番号、箇条書きにセル番地、図形・画像に位置を付けます（AI に場所を答えさせたいときなど）。
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
