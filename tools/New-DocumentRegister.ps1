[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $Planned,

    [string] $Previous,

    [string[]] $Findings,

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $AllVersions,

    [string] $HistorySheet = '改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録',

    [string] $OutXlsx,

    [string] $OutCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

foreach ($name in @('Planned', 'Previous', 'OutXlsx', 'OutCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($Planned -and -not (Test-Path -LiteralPath $Planned)) { throw "-Planned の成果物の予定が見つかりません: $Planned" }
if ($Previous -and -not (Test-Path -LiteralPath $Previous)) { throw "-Previous の前回の成果物一覧が見つかりません: $Previous" }

Update-TypeData -TypeName 'DocumentRegister.Document' -DefaultDisplayPropertySet Folder, File, DocNo, Title, Version, Updated, Note -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 日付は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 表紙の見出し（全角・半角と空白をそろえて比べる）
$coverFields = [ordered]@{
    DocNo   = '^(文書番号|文書No\.?|ドキュメント番号|ドキュメントNo\.?|資料番号|管理番号|文書ID|文書コード)$'
    Title   = '^(文書名|ドキュメント名|資料名|成果物名|書名|タイトル|件名|文書タイトル)$'
    Version = '^(版数|版|バージョン|Version|Ver\.?|Rev\.?|改訂版数|改版番号)$'
    Created = '^(作成日|作成年月日|初版作成日)$'
    Updated = '^(最終更新日|更新日|最終改訂日|改訂日|改版日)$'
    Author  = '^(作成者|担当者|作成担当)$'
    System  = '^(システム名|プロジェクト名|案件名)$'
}
# 改訂履歴の見出し
$historyVersionLabel = $coverFields.Version
$historyDateLabel = '^(日付|改訂日|改版日|更新日|年月日|変更日)$'
$historyContentLabel = '^(改訂内容|変更内容|改版内容|内容|改訂箇所・内容|変更点)$'
$historyAuthorLabel = '^(改訂者|変更者|更新者|担当者|作成者|担当)$'
# 見出しの右のセルがこれらの見出しなら、値とみなさない（改訂履歴の表の見出しの行など）
$allLabelPatterns = @($coverFields.Values) + @($historyDateLabel, $historyContentLabel, $historyAuthorLabel)
# 見出しの無い表紙の文書名とみなす文字の終わり
$titleEndPattern = '(書|一覧|仕様|計画|報告|手順|資料|マニュアル|定義|設計)$'
# 一覧に載せないファイル
$skipPattern = '^(~\$|\.)|^(Thumbs\.db|desktop\.ini)$'

# ---------------------------------------------------------------- 表紙・改訂履歴を読む


function ConvertTo-VersionText {
    #     版数の文字（「第1.1版」「Ver.1.10」「v1.1」など）を「1.1」の形にする（無ければ空）。
    #     最初の数の前の 0 だけを除き、ほかは書いたまま（1.09 を 1.9 にしない）
    param([string] $Text)
    $m = [regex]::Match(([string]$Text).Normalize($nfkc), '(\d+(?:\.\d+)*)')
    if (-not $m.Success) { return '' }
    $parts = $m.Groups[1].Value.Split('.')
    $parts[0] = $parts[0].TrimStart('0')
    if (-not $parts[0]) { $parts[0] = '0' }
    return ($parts -join '.')
}

function ConvertTo-VersionKey {
    #     版数を、点で区切った数の並びで比べる形（数の前の 0 と、後ろの「.0」を除く。1.09 と 1.9、1 と 1.0 は同じ）
    param([string] $Version)
    $parts = @($Version.Split('.') | ForEach-Object { $t = $_.TrimStart('0'); if ($t) { $t } else { '0' } })
    return [regex]::Replace(($parts -join '.'), '(\.0)+$', '')
}

function Test-SameVersion {
    #     版数 A と B が同じか。どちらも「1.10」のような小数（点が 1 つまで）なら、数の大きさで比べる
    #     （数値のセルや Excel で開き直した CSV では、1.10 が 1.1、1.0 が 1 になるため）。ほかは点で区切った数の並びで比べる
    param([string] $A, [string] $B)
    if ($A -match '^\d+(\.\d+)?$' -and $B -match '^\d+(\.\d+)?$') { return [decimal]::Parse($A, $invariant) -eq [decimal]::Parse($B, $invariant) }
    return (ConvertTo-VersionKey $A) -eq (ConvertTo-VersionKey $B)
}

function Test-LowerVersion {
    #     版数 A が B より小さいか。どちらも小数（点が 1 つまで）なら、点で区切った数の並びでも、数の大きさでも小さいときだけ
    #     （1.9 → 1.10 は数では小さいが、版数としては大きい。数値のセルの 1.09 → 1.10（1.1）は、並びでは小さいが数では大きい）
    param([string] $A, [string] $B)
    $lower = $false
    try { $lower = (Get-NameVersion ('v' + $A)) -lt (Get-NameVersion ('v' + $B)) } catch { return $false }
    if ($lower -and $A -match '^\d+(\.\d+)?$' -and $B -match '^\d+(\.\d+)?$') { $lower = [decimal]::Parse($A, $invariant) -lt [decimal]::Parse($B, $invariant) }
    return [bool]$lower
}

function Get-FileNameVersion {
    #     ファイル名の版数（設計書_v1.1.xlsx・設計書 Ver1.1.xlsx・設計書 第1.1版.xlsx の「1.1」。「第2回」は版数にしない。無ければ空）
    param([string] $FileName)
    $m = [regex]::Match([System.IO.Path]::GetFileNameWithoutExtension($FileName).Normalize($nfkc), '(?i)(?<![a-z])v(?:er)?\.?\s*(\d+(?:\.\d+)*)|第\s*(\d+(?:\.\d+)*)\s*版')
    if (-not $m.Success) { return '' }
    return (ConvertTo-VersionText ($m.Groups[1].Value + $m.Groups[2].Value))
}

function ConvertTo-DateText {
    #     日付の文字を「2026/10/01」の形にする（読めなければそのまま）
    param([string] $Text)
    $d = ConvertTo-DateValue $Text $AsOf
    if ($d) { return $d.ToString('yyyy/MM/dd', $invariant) }
    return ([string]$Text).Trim()
}

function Test-HeadingText {
    param([string] $Text)
    $l = ConvertTo-Label $Text
    if (-not $l) { return $false }
    foreach ($p in $allLabelPatterns) { if ($l -match $p) { return $true } }
    return $false
}

function Read-CoverFields {
    #     表紙の行（行ごとのセルの文字の配列）から、見出しの右の値を読む（文書番号・文書名・版数・作成日・更新日・作成者・システム名）。
    #     見出しの無い文書名は「…書」「…一覧」などで終わる最初の文字、見出しの無い版数は「第1.1版」「Ver.1.1」の形の文字
    param($Rows)
    $result = @{}
    $titleGuess = ''
    foreach ($cells in $Rows) {
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $label = ConvertTo-Label $cells[$i]
            if (-not $label) { continue }
            $field = $null
            foreach ($k in $coverFields.Keys) { if ($label -match $coverFields[$k]) { $field = $k; break } }
            if ($field) {
                if (-not $result.ContainsKey($field) -and $i + 1 -lt $cells.Count -and ([string]$cells[$i + 1]).Trim() -and -not (Test-HeadingText $cells[$i + 1])) { $result[$field] = ([string]$cells[$i + 1]).Trim(); $i++ }
                continue
            }
            if (-not $result.ContainsKey('Version') -and $label -match '^(第\d+(\.\d+)*版|(Ver|Version|Rev)\.?\d+(\.\d+)*)$') { $result['Version'] = ([string]$cells[$i]).Trim(); continue }
            # 「画面設計書（SCR001 受注一覧）」のように、後ろの（ ）は除いて見る
            if (-not $titleGuess -and $label.Length -le 40 -and ($label -replace '[（(][^（）()]*[）)]$', '') -match $titleEndPattern) { $titleGuess = ([string]$cells[$i]).Trim() }
        }
    }
    if (-not $result.ContainsKey('Title') -and $titleGuess) { $result['Title'] = $titleGuess }
    return $result
}

function Read-HistoryEntry {
    #     改訂履歴の行から、版数のある最後の行（版数・日付・内容・改訂者）。帳票のヘッダにも「版数」「更新日」が並ぶことがあるので、
    #     下に版数の行がいちばん多い見出しの行を使う。無ければ $null
    param($Rows)
    $best = $null
    for ($h = 0; $h -lt $Rows.Count; $h++) {
        $cols = @{}
        for ($i = 0; $i -lt $Rows[$h].Count; $i++) {
            $label = ConvertTo-Label $Rows[$h][$i]
            if (-not $cols.ContainsKey('Version') -and $label -match $historyVersionLabel) { $cols['Version'] = $i }
            elseif (-not $cols.ContainsKey('Date') -and $label -match $historyDateLabel) { $cols['Date'] = $i }
            elseif (-not $cols.ContainsKey('Content') -and $label -match $historyContentLabel) { $cols['Content'] = $i }
            elseif (-not $cols.ContainsKey('Author') -and $label -match $historyAuthorLabel) { $cols['Author'] = $i }
        }
        if (-not $cols.ContainsKey('Version') -or -not $cols.ContainsKey('Date')) { continue }
        $entries = New-Object System.Collections.Generic.List[object]
        for ($r = $h + 1; $r -lt $Rows.Count; $r++) {
            $cells = $Rows[$r]
            $get = { param($k) if ($cols.ContainsKey($k) -and $cols[$k] -lt $cells.Count) { ([string]$cells[$cols[$k]]).Trim() } else { '' } }
            $version = ConvertTo-VersionText (& $get 'Version')
            if (-not $version) { continue }
            $entries.Add([pscustomobject]@{ Version = $version; Date = (ConvertTo-DateText (& $get 'Date')); Content = (& $get 'Content'); Author = (& $get 'Author') })
        }
        if ($entries.Count -gt 0 -and (-not $best -or $entries.Count -gt $best.Count)) { $best = $entries }
    }
    if (-not $best) { return $null }
    return $best[$best.Count - 1]
}

function Get-GridRows {
    #     シート（ConvertTo-DesignBooks）の行ごとのセルの文字（列の順。空のセルは詰める）
    param($Grid)
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $Grid) { return , $rows }
    foreach ($r in $Grid.Rows) { $rows.Add(@($Grid.Cols[$r] | ForEach-Object { [string]$Grid.Cells[$r][$_] })) }
    return , $rows
}

function Get-GridTableRows {
    #     シートの行ごとのセルの文字（列の位置のまま。空のセルは空の文字。改訂履歴の表のように、列の見出しで値を読むとき）
    param($Grid)
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $Grid) { return , $rows }
    $first = [int]::MaxValue; $last = 0
    foreach ($r in $Grid.Rows) { foreach ($c in $Grid.Cols[$r]) { if ($c -lt $first) { $first = $c }; if ($c -gt $last) { $last = $c } } }
    foreach ($r in $Grid.Rows) {
        $cells = $Grid.Cells[$r]
        $rows.Add(@(for ($c = $first; $c -le $last; $c++) { if ($cells.ContainsKey($c)) { [string]$cells[$c] } else { '' } }))
    }
    return , $rows
}

function Read-CoreProperties {
    #     ファイルのプロパティ（docProps/core.xml）のタイトル・作成者・更新日時。読めなければ空
    param([string] $FilePath)
    $props = @{ Title = ''; Author = ''; Updated = '' }
    try {
        $pkg = Open-ZipPackage $FilePath
        try {
            $doc = Read-PartXml -Package $pkg -PartName 'docProps/core.xml'
            if ($doc) {
                foreach ($pair in @(@('title', 'Title'), @('creator', 'Author'), @('modified', 'Updated'))) {
                    $node = $doc.SelectSingleNode("//*[local-name()='$($pair[0])']")
                    if ($node -and $node.InnerText.Trim()) { $props[$pair[1]] = $node.InnerText.Trim() }
                }
            }
        }
        finally { $pkg.Zip.Dispose() }
    }
    catch { }
    if ($props.Updated) { $props.Updated = ConvertTo-DateText $props.Updated }
    return $props
}

function Read-WordDocument {
    #     Word の表紙（本文の表の行と、「見出し：値」の段落）と改訂履歴（版数と日付の見出しのある表）
    param([string] $FilePath)
    $pkg = Open-ZipPackage $FilePath
    try {
        $doc = Read-PartXml -Package $pkg -PartName 'word/document.xml'
        if ($null -eq $doc) { throw 'Word の文書ではありません' }
        $nsm = New-WordNamespaceManager $doc
        $rows = New-Object System.Collections.Generic.List[object]
        $tables = [ordered]@{}
        foreach ($b in @(Get-WordBlocks $doc $nsm)) {
            if ($b.Kind -eq '表') {
                $rows.Add(@($b.Cells))
                $t = ($b.Cell -split ' の ')[0]
                if (-not $tables.Contains($t)) { $tables[$t] = New-Object System.Collections.Generic.List[object] }
                $tables[$t].Add(@($b.Cells))
            }
            elseif ($b.Text) {
                $m = [regex]::Match($b.Text, '^\s*([^:：]{1,12})\s*[:：]\s*(.+)$')
                if ($m.Success) { $rows.Add(@($m.Groups[1].Value, $m.Groups[2].Value)) } else { $rows.Add(@($b.Text)) }
            }
        }
    }
    finally { $pkg.Zip.Dispose() }
    $history = $null
    foreach ($t in $tables.Values) { $e = Read-HistoryEntry $t; if ($e) { $history = $e; break } }
    return [pscustomobject]@{ Cover = (Read-CoverFields $rows); History = $history }
}

function Get-PdfPages {
    #     PDF のページ数（ページのオブジェクトの数。数えられなければ空）。
    #     追記で保存した PDF は同じ番号のオブジェクトが何度も出てくるので、番号ごとに最後のものを見る
    param([string] $FilePath)
    try {
        $text = [System.Text.Encoding]::GetEncoding(28591).GetString([System.IO.File]::ReadAllBytes($FilePath))
        $isPage = @{}
        foreach ($m in [regex]::Matches($text, '(?s)(\d+)\s+\d+\s+obj\b(.*?)endobj')) { $isPage[$m.Groups[1].Value] = ($m.Groups[2].Value -match '/Type\s*/Page(?![s\w])') }
        $n = @($isPage.Values | Where-Object { $_ }).Count
        if ($n -gt 0) { return $n }
    }
    catch { }
    return ''
}

# ---------------------------------------------------------------- ファイルを集める


$files = New-Object System.Collections.Generic.List[object]
$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "成果物の置き場所が見つかりません: $p"; continue }
    foreach ($item in $items) {
        $root = $item.FullName
        $list = @($item)
        if ($item.PSIsContainer) { $list = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) }
        else { $root = $item.DirectoryName }
        foreach ($f in $list) {
            if ($f.Name -match $skipPattern) { continue }
            if (@($OutXlsx, $OutCsv, $Planned, $Previous) | Where-Object { $_ -and [string]::Equals($_, $f.FullName, [System.StringComparison]::OrdinalIgnoreCase) }) { continue }
            if (-not $seen.Add($f.FullName)) { continue }
            $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
            $files.Add([pscustomobject]@{ Info = $f; Relative = $rel })
        }
    }
}
# 相対パスの順（[Array]::Sort に PowerShell の配列を渡すと、写しが並べ替わるだけなので、List の Sort で並べる）
$files.Sort([System.Comparison[object]] { param($a, $b) [string]::Compare($a.Relative, $b.Relative, [System.StringComparison]::OrdinalIgnoreCase) })
$fileArray = $files.ToArray()

# Excel の表紙と改訂履歴の文字（まとめて Search-Excel で読む）
$excelExtensions = @('.xlsx', '.xlsm', '.xls', '.xlsb')
$books = @{}
$excelPaths = @($fileArray | Where-Object { $excelExtensions -contains $_.Info.Extension.ToLowerInvariant() } | ForEach-Object { $_.Info.FullName })
if ($excelPaths.Count -gt 0) {
    $searchArgs = @{ Pattern = '\S'; Path = $excelPaths; Target = 'Value' }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $grouped = ConvertTo-DesignBooks @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null)
    foreach ($k in $grouped.Keys) { $books[$k] = $grouped[$k] }
}

$docs = New-Object System.Collections.Generic.List[object]
foreach ($entry in $fileArray) {
    $f = $entry.Info
    $ext = $f.Extension.ToLowerInvariant()
    $kind = 'そのほか'
    if ($excelExtensions -contains $ext) { $kind = 'Excel' }
    elseif (@('.docx', '.docm', '.doc') -contains $ext) { $kind = 'Word' }
    elseif ($ext -eq '.pdf') { $kind = 'PDF' }
    $folder = [System.IO.Path]::GetDirectoryName($entry.Relative)
    if (-not $folder) { $folder = '' }
    $doc = [pscustomobject]@{
        Folder = $folder; File = $f.Name; Path = $f.FullName; Relative = $entry.Relative; Kind = $kind; Size = [math]::Ceiling($f.Length / 1024); Modified = $f.LastWriteTime
        Key = (Get-VersionlessName $entry.Relative); Order = (Get-NameVersion $f.Name); NameVersion = ''
        DocNo = ''; Title = ''; Version = ''; Created = ''; Updated = ''; Author = ''; System = ''; History = $null; Pages = ''
        Notes = (New-Object System.Collections.Generic.List[string]); Latest = $true; LatestFile = ''; Planned = ''; Note = ''
        Change = ''; OpenFindings = 0; OverdueFindings = 0
    }
    # 版数はファイル名、無ければフォルダ名（v1.0\設計書.xlsx など）
    if ($null -eq $doc.Order) { $doc.Order = Get-NameVersion $entry.Relative }
    $doc.NameVersion = Get-FileNameVersion $f.Name
    $cover = @{}
    try {
        if ($kind -eq 'Excel') {
            $book = $books[$f.FullName]
            if ($book) {
                $sheetNames = @($book.Sheets.Keys)
                $coverName = @($sheetNames | Where-Object { $_ -match '^\s*表紙\s*$' })[0]
                if (-not $coverName) { $coverName = @($sheetNames | Where-Object { $_ -notmatch $HistorySheet })[0] }
                if ($coverName) { $cover = Read-CoverFields (Get-GridRows $book.Sheets[$coverName]) }
                $historyName = @($sheetNames | Where-Object { $_ -match $HistorySheet })[0]
                if ($historyName) { $doc.History = Read-HistoryEntry (Get-GridTableRows $book.Sheets[$historyName]) }
            }
        }
        elseif ($kind -eq 'Word' -and $ext -ne '.doc') {
            $word = Read-WordDocument $f.FullName
            $cover = $word.Cover
            $doc.History = $word.History
        }
        elseif ($kind -eq 'PDF') { $doc.Pages = Get-PdfPages $f.FullName }
    }
    catch { $doc.Notes.Add(('読めません（{0}）' -f $_.Exception.Message)) }
    foreach ($k in @('DocNo', 'Title', 'Version', 'Created', 'Updated', 'Author', 'System')) { if ($cover.ContainsKey($k)) { $doc.$k = $cover[$k] } }
    if (@('.xlsx', '.xlsm', '.docx', '.docm') -contains $ext) {
        $props = Read-CoreProperties $f.FullName
        if (-not $doc.Title) { $doc.Title = $props.Title }
        if (-not $doc.Author) { $doc.Author = $props.Author }
    }
    if ($doc.Created) { $doc.Created = ConvertTo-DateText $doc.Created }
    if ($doc.Updated) { $doc.Updated = ConvertTo-DateText $doc.Updated }
    if (-not $doc.Title) { $doc.Title = [regex]::Replace([System.IO.Path]::GetFileNameWithoutExtension($f.Name), '[_\-\s]*(?:(?<![A-Za-z])[vV](?:[eE][rR])?\.?\s*\d+(?:\.\d+)*版?|第\s*\d+(?:\.\d+)*\s*版)$', '') }

    # 版数・更新日の食い違い（ファイル名・表紙・改訂履歴）
    $versions = [ordered]@{}
    if ($doc.NameVersion) { $versions['ファイル名'] = $doc.NameVersion }
    if ($doc.Version) { $v = ConvertTo-VersionText $doc.Version; if ($v) { $versions['表紙'] = $v } }
    if ($doc.History) { $versions['改訂履歴'] = $doc.History.Version }
    $distinct = New-Object System.Collections.Generic.List[string]
    foreach ($v in $versions.Values) { if (@($distinct | Where-Object { Test-SameVersion $_ $v }).Count -eq 0) { $distinct.Add($v) } }
    if ($distinct.Count -gt 1) {
        $doc.Notes.Add(('版数が違う（{0}）' -f ((@($versions.Keys | ForEach-Object { '{0} {1}' -f $_, $versions[$_] })) -join '・')))
    }
    if ($doc.Updated -and $doc.History -and $doc.History.Date -and $doc.Updated -ne $doc.History.Date) {
        $doc.Notes.Add(('更新日が違う（表紙 {0}・改訂履歴 {1}）' -f $doc.Updated, $doc.History.Date))
    }
    $docs.Add($doc)
}

# 同じ文書の古い版（版数だけが違うファイル）。版数の大きいもの、同じなら更新日時の新しいものを最新とみなす
foreach ($g in @($docs | Group-InOrder Key)) {
    if ($g.Count -lt 2) { continue }
    $sorted = @($g.Group | Sort-Object @{ e = { if ($_.Order) { $_.Order } else { New-Object System.Version(0, 0, 0, 0) } } }, Modified, Relative)
    $latest = $sorted[$sorted.Count - 1]
    foreach ($d in $sorted) {
        if ($d -eq $latest) { continue }
        $d.Latest = $false
        $d.LatestFile = $latest.Relative
    }
}

# ---------------------------------------------------------------- 成果物の予定と突き合わせる


$labels = { param([string[]] $Names) @($Names | ForEach-Object { ConvertTo-Label $_ }) }
# 文書番号を比べる形（全角・半角と大文字小文字をそろえる。数字だけの番号は前の 0 を除く（Excel で開き直した CSV の 001 は 1）。
# 「－」「-」など英数字の無いものは文書番号にしない（空）
$docNoKey = {
    param([string] $Text)
    $k = (ConvertTo-Label $Text).ToUpperInvariant()
    if ($k -notmatch '[\p{L}\p{N}]') { return '' }
    if ($k -match '^\d+$') { $k = $k.TrimStart('0'); if (-not $k) { $k = '0' } }
    return $k
}
$plannedRows = New-Object System.Collections.Generic.List[object]
$readLog = New-Object System.Collections.Generic.List[object]
if ($Planned) {
    $candidates = @{
        '成果物'   = & $labels @('成果物', '成果物名', '文書名', 'ドキュメント名', '資料名', '納品物', '納品物名', 'ファイル名', '名称')
        '文書番号' = & $labels @('文書番号', '文書No', '管理番号', '資料番号', 'ドキュメント番号')
        '予定日'   = & $labels @('予定日', '提出予定日', '納期', '期限', '提出日', '納品日')
        '担当'     = & $labels @('担当', '担当者', '作成者')
    }
    $records = Read-HeaderRecords -Inputs (Get-InputFiles @($Planned) '成果物の予定') -Candidates $candidates -Fields @('成果物', '文書番号', '予定日', '担当') `
        -Need @('成果物') -Key @('成果物', '文書番号') -What '成果物の予定' -NoExcel:$NoExcel -ReadLog $readLog
    $current = @($docs | Where-Object { $_.Latest })
    $key = { param([string] $Text) (ConvertTo-Label ([regex]::Replace(([string]$Text).Trim(), '\.[A-Za-z0-9]{2,5}$', ''))).ToUpperInvariant() }
    foreach ($rec in $records) {
        $name = ([string]$rec.Values['成果物']).Trim()
        $docNo = ([string]$rec.Values['文書番号']).Trim()
        if (-not $name -and -not $docNo) { continue }
        $matched = @()
        $no = & $docNoKey $docNo
        if ($no) { $matched = @($current | Where-Object { (& $docNoKey $_.DocNo) -eq $no }) }
        if ($matched.Count -eq 0 -and $name) {
            # 成果物の名前が、ファイル名（版数を除く）か文書名に含まれるもの（同じ名前があればそれだけ）
            $nk = & $key (Get-VersionlessName $name).Replace('{v}', '')
            $equal = @($current | Where-Object { (& $key $_.Title) -eq $nk -or (& $key ([System.IO.Path]::GetFileName($_.Key).Replace('{v}', ''))) -eq $nk })
            if ($equal.Count -gt 0) { $matched = $equal }
            elseif ($nk.Length -ge 2) { $matched = @($current | Where-Object { (& $key ([System.IO.Path]::GetFileName($_.Key).Replace('{v}', ''))).Contains($nk) -or (& $key $_.Title).Contains($nk) }) }
        }
        $due = ConvertTo-DateValue $rec.Values['予定日'] $AsOf
        $state = '提出済み'
        if ($matched.Count -eq 0) {
            $state = '未提出'
            if ($due -and $due -lt $AsOf) { $state = '未提出（予定日を過ぎています）' }
        }
        foreach ($d in $matched) { $d.Planned = $name }
        $plannedRows.Add([pscustomobject]@{ Name = $name; DocNo = $docNo; Due = $(if ($due) { $due.ToString('yyyy/MM/dd', $invariant) } else { ([string]$rec.Values['予定日']).Trim() }); Owner = ([string]$rec.Values['担当']).Trim(); State = $state; Docs = $matched; Place = $rec })
    }
    foreach ($d in $current) { if (-not $d.Planned) { $d.Notes.Add('予定に無い') } }
}

# ---------------------------------------------------------------- 前回の成果物一覧と比べる


function Get-DocVersion {
    #     比べる版数（表紙・改訂履歴・ファイル名の版数のうち、いちばん大きいもの。同じ版数なら桁の多い書き方（1.1 より 1.10）。無ければ空）
    param([string] $Cover, [string] $History, [string] $FileName)
    $best = ''
    $bestOrder = $null
    foreach ($v in @((ConvertTo-VersionText $Cover), (ConvertTo-VersionText $History), (Get-FileNameVersion $FileName))) {
        if (-not $v) { continue }
        $order = $null
        try { $order = Get-NameVersion ('v' + $v) } catch { }
        $same = $best -and (Test-SameVersion $v $best)
        if (-not $best -or ($same -and $v.Length -gt $best.Length) -or (-not $same -and $order -gt $bestOrder)) { $best = $v; $bestOrder = $order }
    }
    return $best
}

function Get-PairKey {
    #     前回と今回の成果物を突き合わせる形（版数を除いた場所と名前。Excel で開き直した CSV では「01」のフォルダが「1」になるので、
    #     数字だけのフォルダの前の 0 を除く）
    param([string] $Relative)
    return [regex]::Replace((Get-VersionlessName $Relative.Replace('\', '/')), '(?<=^|/)0+(?=\d+(?:/|$))', '')
}

function ConvertTo-ModifiedTime {
    #     前回の一覧の「ファイルの更新日時」（2026/10/01 09:30。Excel で開き直した 2026/10/1 9:30・シリアル値も）。
    #     Format は比べる形（時刻の無いものは日付だけで比べる）。読めなければ $null
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    if (-not $t) { return $null }
    try {
        $m = [regex]::Match($t, '^(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})日?(?:\s+(\d{1,2}):(\d{2}))?')
        if ($m.Success) {
            $value = New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value)
            if (-not $m.Groups[4].Success) { return [pscustomobject]@{ Value = $value; Format = 'yyyy/MM/dd' } }
            return [pscustomobject]@{ Value = $value.AddHours([int]$m.Groups[4].Value).AddMinutes([int]$m.Groups[5].Value); Format = 'yyyy/MM/dd HH:mm' }
        }
        $n = 0.0
        if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n) -and $n -ge 1 -and $n -lt 2958466) {
            # シリアル値は分に丸める
            $minutes = [long][math]::Round([DateTime]::FromOADate($n).Ticks / 600000000.0)
            $format = $(if ($n -ne [math]::Floor($n)) { 'yyyy/MM/dd HH:mm' } else { 'yyyy/MM/dd' })
            return [pscustomobject]@{ Value = (New-Object DateTime ($minutes * [long]600000000)); Format = $format }
        }
    }
    catch { }
    return $null
}

$changeRows = New-Object System.Collections.Generic.List[object]
$previousCount = 0
if ($Previous) {
    $fields = @('ファイル', 'フォルダ', '文書番号', '文書名', '版数', '最新の改訂', '更新日時')
    $candidates = @{
        'ファイル' = & $labels @('ファイル', 'ファイル名'); 'フォルダ' = & $labels @('フォルダ'); '文書番号' = & $labels @('文書番号'); '文書名' = & $labels @('文書名')
        '版数' = & $labels @('版数'); '最新の改訂' = & $labels @('最新の改訂'); '更新日時' = & $labels @('ファイルの更新日時')
    }
    # 成果物一覧と古い版のシートを読み、同じ文書（版数だけが違うファイル）は版数の大きいものを前回の版とする
    $records = Read-HeaderRecords -Inputs (Get-InputFiles @($Previous) '前回の成果物一覧') -Candidates $candidates -Fields $fields `
        -Need @('ファイル', '更新日時') -Key @('ファイル') -What '前回の成果物一覧' -NoExcel:$NoExcel -ReadLog $readLog
    if ($records.Count -eq 0) { Write-Warning "前回の成果物一覧に、ファイルとファイルの更新日時の見出しのある表がありません: $Previous" }
    $before = [ordered]@{}
    $unreadable = New-Object System.Collections.Generic.List[string]
    foreach ($rec in $records) {
        $v = $rec.Values
        $file = ([string]$v['ファイル']).Trim()
        if (-not $file) { continue }
        $folder = ([string]$v['フォルダ']).Trim()
        $relative = $file
        if ($folder) { $relative = ($folder + '/' + $file).Replace('\', '/') }
        $p = [pscustomobject]@{
            Folder = $folder; File = $file; Relative = $relative; DocNo = ([string]$v['文書番号']).Trim(); Title = ([string]$v['文書名']).Trim()
            Version = (Get-DocVersion $v['版数'] $v['最新の改訂'] $file); Modified = (ConvertTo-ModifiedTime $v['更新日時']); Order = (Get-NameVersion $file)
        }
        # 版数はファイル名、無ければフォルダ名（v1.0\設計書.xlsx と v1.1\設計書.xlsx。-AllVersions の一覧には両方ある）
        if ($null -eq $p.Order) { $p.Order = Get-NameVersion $relative }
        if (-not $p.Modified -and ([string]$v['更新日時']).Trim()) { $unreadable.Add(([string]$v['更新日時']).Trim()) }
        $k = Get-PairKey $relative
        if (-not $before.Contains($k) -or ($p.Order -and $p.Order -gt $before[$k].Order)) { $before[$k] = $p }
    }
    # 前の版で、和暦の設定のまま書いた一覧（08/10/20 10:00）など。版数を上げずに直したものを見逃すので知らせる
    if ($unreadable.Count -gt 0) {
        Write-Warning ('前回の一覧の {0} 行で、ファイルの更新日時を読めません（{1} など）。この行は、版数を上げずに直したかを確かめません' -f $unreadable.Count, $unreadable[0])
    }
    $previousCount = $before.Count

    # 同じ場所・名前（版数を除く）のもの。無ければ、文書番号が前回と今回で 1 つずつだけのもの（名前・場所を変えた文書）
    $current = @($docs | Where-Object { $_.Latest })
    $pairs = @{}
    $used = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($d in $current) {
        $k = Get-PairKey $d.Relative
        if ($before.Contains($k)) { $pairs[$d.Relative] = $before[$k]; [void]$used.Add($k) }
    }
    foreach ($d in $current) {
        if ($pairs.ContainsKey($d.Relative)) { continue }
        $no = & $docNoKey $d.DocNo
        if (-not $no) { continue }
        $same = @($before.Keys | Where-Object { -not $used.Contains($_) -and (& $docNoKey $before[$_].DocNo) -eq $no })
        $others = @($current | Where-Object { -not $pairs.ContainsKey($_.Relative) -and (& $docNoKey $_.DocNo) -eq $no })
        if ($same.Count -eq 1 -and $others.Count -eq 1) { $pairs[$d.Relative] = $before[$same[0]]; [void]$used.Add($same[0]) }
    }

    foreach ($d in $current) {
        $version = Get-DocVersion $d.Version $(if ($d.History) { $d.History.Version } else { '' }) $d.File
        $p = $pairs[$d.Relative]
        if (-not $p) {
            $d.Change = '新規'
            $changeRows.Add([pscustomobject]@{ Change = $d.Change; Doc = $d; Previous = $null; Version = $version })
            continue
        }
        $parts = New-Object System.Collections.Generic.List[string]
        if ((Get-PairKey $p.Relative) -ne (Get-PairKey $d.Relative)) { $parts.Add(('名前・場所が変わった（前回: {0}）' -f $p.Relative)) }
        if ($p.Version -and $version -and -not (Test-SameVersion $p.Version $version)) {
            $parts.Add(('版数 {0}→{1}' -f $p.Version, $version))
            if (Test-LowerVersion $version $p.Version) { $d.Notes.Add(('版数が前回（{0}）より小さい' -f $p.Version)) }
        }
        elseif ($p.Modified -and $d.Modified.ToString($p.Modified.Format, $invariant) -ne $p.Modified.Value.ToString($p.Modified.Format, $invariant)) {
            # 版数のある文書を、版数を変えずに直した（版数の無いファイルは「更新」だけ）
            if ($p.Version -and $version) { $parts.Add('版数を上げずに更新'); $d.Notes.Add('前回の一覧から版数を上げずに更新') }
            else { $parts.Add('更新') }
        }
        $d.Change = $parts -join '・'
        if ($d.Change) { $changeRows.Add([pscustomobject]@{ Change = $d.Change; Doc = $d; Previous = $p; Version = $version }) }
    }
    foreach ($k in $before.Keys) {
        if (-not $used.Contains($k)) { $changeRows.Add([pscustomobject]@{ Change = '無くなった'; Doc = $null; Previous = $before[$k]; Version = '' }) }
    }
}

# ---------------------------------------------------------------- 指摘管理表の未完了の指摘


function ConvertTo-NameKey {
    #     成果物の名前を比べる形（拡張子・版数を除き、全角・半角と空白をそろえる。フォルダは除かない（「外部I/F設計書」の / を残す））
    param([string] $Name)
    $n = ([string]$Name).Trim()
    $ext = [regex]::Match($n, '(?i)\.(xlsx|xlsm|xls|xlsb|docx|docm|doc|pdf|pptx|ppt)$')
    if ($ext.Success) { $n = $n.Substring(0, $ext.Index) }
    $n = [regex]::Replace($n.Normalize($nfkc), '[_\-\s]*(?:(?<![A-Za-z])[vV](?:[eE][rR])?\.?\s*\d+(?:\.\d+)*版?|第\s*\d+(?:\.\d+)*\s*版)', '')
    return (ConvertTo-Label $n).ToUpperInvariant()
}

$openFindings = New-Object System.Collections.Generic.List[object]
if ($Findings) {
    # 対応が終わった・指摘として数えない状態（Measure-ReviewQuality.ps1 と同じ）
    $closedPattern = '完了|済|クローズ|close|done|対応不要|不要|対処なし'
    $excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|誤指摘'
    $fields = @('設計書', '指摘内容', '重要度', '状態', '対応要否', '期限', '完了日', '対応者')
    $candidates = Read-HeaderCandidates (Join-Path $PSScriptRoot 'review-quality/指摘管理表の見出し.csv') $fields
    $skip = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($p in @($OutXlsx, $OutCsv, $Previous, $Planned)) { if ($p) { [void]$skip.Add($p) } }
    $records = Read-HeaderRecords -Inputs (Get-InputFiles $Findings '指摘管理表' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
        -Need @('指摘内容', '設計書') -Key @('指摘内容') -What '指摘管理表' -Skip $skip -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
    if ($records.Count -eq 0) { Write-Warning '指摘管理表が見つかりませんでした（指摘内容と、設計書（対象成果物）の見出しがある表を探します）' }

    # 文書番号・文書名・ファイル名（版数と拡張子を除く）→ 成果物。場所（フォルダと名前）は、フォルダの付いた指摘の対象と突き合わせる
    $index = @{}
    $order = @{}
    $pathKeyOf = @{}
    $latestDocs = @($docs | Where-Object { $_.Latest })
    foreach ($d in $latestDocs) {
        $order[$d.Relative] = $order.Count
        $pathKeyOf[$d.Relative] = ConvertTo-NameKey $d.Relative.Replace('\', '/')
        foreach ($k in @((ConvertTo-NameKey $d.File), (ConvertTo-NameKey $d.Title), $(if ($d.DocNo) { & $docNoKey $d.DocNo }))) {
            if (-not $k) { continue }
            if (-not $index.ContainsKey($k)) { $index[$k] = New-Object System.Collections.Generic.List[object] }
            if (-not $index[$k].Contains($d)) { $index[$k].Add($d) }
        }
    }
    foreach ($rec in $records) {
        $v = $rec.Values
        $content = ([string]$v['指摘内容']).Trim()
        if (-not $content) { continue }
        $status = ([string]$v['状態']).Trim()
        if ($status -match $excludedPattern) { continue }
        $closedOn = ConvertTo-DateValue $v['完了日'] $AsOf
        if (($status -match $closedPattern) -or (([string]$v['対応要否']).Trim() -match '^(不要|否|×|-)$') -or ((-not $status) -and $closedOn)) { continue }
        $name = ([string]$v['設計書']).Trim()
        $matched = New-Object System.Collections.Generic.List[object]
        # フォルダの付いた対象（結合テスト/テスト仕様書.xlsx・D:\納品物\…）は、まず成果物の場所と突き合わせる（ほかのフォルダの同じ名前に数えない）
        if ($name -match '[\\/]') {
            $pathKey = ConvertTo-NameKey $name.Replace('\', '/')
            foreach ($d in $latestDocs) {
                if ($pathKey -eq $pathKeyOf[$d.Relative] -or $pathKey.EndsWith('/' + $pathKeyOf[$d.Relative], [System.StringComparison]::Ordinal)) { $matched.Add($d) }
            }
        }
        # 無ければ、文書名（外部I/F設計書）・ファイル名・文書番号
        if ($matched.Count -eq 0) {
            foreach ($k in @((ConvertTo-NameKey $name), (ConvertTo-NameKey (Get-LeafName $name)), (& $docNoKey $name))) {
                if (-not $k -or -not $index.ContainsKey($k)) { continue }
                foreach ($d in $index[$k]) { if (-not $matched.Contains($d)) { $matched.Add($d) } }
            }
        }
        $due = ConvertTo-DateValue $v['期限'] $AsOf
        $overdue = [bool]($due -and $due -lt $AsOf)
        foreach ($d in $matched) { $d.OpenFindings++; if ($overdue) { $d.OverdueFindings++ } }
        $first = [int]::MaxValue
        foreach ($d in $matched) { if ($order[$d.Relative] -lt $first) { $first = $order[$d.Relative] } }
        $openFindings.Add([pscustomobject]@{
                Document = $name; Content = $content; Severity = ([string]$v['重要度']).Trim(); Status = $status; Owner = ([string]$v['対応者']).Trim()
                Due = $due; Overdue = $overdue; Docs = $matched.ToArray(); Order = $first; Seq = $openFindings.Count; Place = $rec
            })
    }
    # 成果物一覧の順（対応付かない指摘は最後）
    $openFindings.Sort([System.Comparison[object]] { param($a, $b) $c = $a.Order.CompareTo($b.Order); if ($c -eq 0) { $c = $a.Seq.CompareTo($b.Seq) }; $c })
    foreach ($d in $docs) {
        if ($d.OpenFindings -eq 0) { continue }
        $note = '未完了の指摘 {0} 件' -f $d.OpenFindings
        if ($d.OverdueFindings -gt 0) { $note += '（期限切れ {0} 件）' -f $d.OverdueFindings }
        $d.Notes.Add($note)
    }
}

# ---------------------------------------------------------------- 結果


foreach ($d in $docs) { $d.Note = $d.Notes -join '・' }
$listed = @($docs | Where-Object { $AllVersions -or $_.Latest })
$old = @($docs | Where-Object { -not $_.Latest })
foreach ($d in $listed) {
    if (-not $d.Latest) { $d.Note = (@(@('古い版（最新: {0}）' -f $d.LatestFile) + @($d.Notes)) -join '・') }
    $d.PSObject.TypeNames.Insert(0, 'DocumentRegister.Document')
}
$listed

# 「前回から」の列は -Previous のときだけ（いつもの一覧の列は変えない）
$headers = @('No', 'フォルダ', 'ファイル', '種類', '文書番号', '文書名', '版数', '作成日', '更新日', '作成者', '最新の改訂', 'ページ', 'サイズ（KB）', 'ファイルの更新日時')
$widths = @(5, 18, 34, 7, 14, 24, 7, 11, 11, 12, 34, 6, 9, 16)
if ($Previous) { $headers += '前回から'; $widths += 22 }
$headers += '確認'
$widths += 40
$registerRows = New-Object System.Collections.Generic.List[object]
$no = 0
foreach ($d in $listed) {
    $no++
    $history = ''
    if ($d.History) { $history = (@($d.History.Version, $d.History.Date, $d.History.Content, $d.History.Author) | Where-Object { $_ }) -join ' ' }
    $values = @([double]$no, $d.Folder, $d.File, $d.Kind, $d.DocNo, $d.Title, $d.Version, $d.Created, $d.Updated, $d.Author, $history, $d.Pages, [double]$d.Size, $d.Modified.ToString('yyyy/MM/dd HH:mm', $invariant))
    if ($Previous) { $values += $d.Change }
    $registerRows.Add([pscustomobject]@{ Doc = $d; Values = @($values + $d.Note) })
}

if ($OutCsv) {
    $lines = New-Object System.Collections.Generic.List[string]
    $quote = { param($v) $s = [string]$v; if ($s -match '[",\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s } }
    $lines.Add((@($headers | ForEach-Object { & $quote $_ })) -join ',')
    foreach ($r in $registerRows) { $lines.Add((@($r.Values | ForEach-Object { & $quote $_ })) -join ',') }
    [System.IO.File]::WriteAllText($OutCsv, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

if ($OutXlsx) {
    $columns = @(for ($i = 0; $i -lt $headers.Count; $i++) { @{ Header = $headers[$i]; Width = $widths[$i] } })
    $sheets = New-Object System.Collections.Generic.List[object]
    $sheets.Add(@{ Name = '成果物一覧'; Columns = $columns; Rows = @($registerRows | ForEach-Object { @{ Values = $_.Values; Link = @{ Column = 2; Path = $_.Doc.Path } } }) })
    if ($Planned) {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($p in $plannedRows) {
            $first = @($p.Docs)[0]
            $row = @{ Values = @($p.Name, $p.DocNo, $p.Due, $p.Owner, $p.State, ((@($p.Docs | ForEach-Object { $_.Relative })) -join '・'), $(if ($first) { $first.Version } else { '' })) }
            if ($first) { $row.Link = @{ Column = 5; Path = $first.Path } }
            $rows.Add($row)
        }
        foreach ($d in @($docs | Where-Object { $_.Latest -and -not $_.Planned })) { $rows.Add(@{ Values = @('', $d.DocNo, '', '', '予定に無い', $d.Relative, $d.Version); Link = @{ Column = 5; Path = $d.Path } }) }
        $sheets.Add(@{ Name = '予定との突き合わせ'; Columns = @(@{ Header = '成果物（予定）'; Width = 30 }, @{ Header = '文書番号'; Width = 14 }, @{ Header = '予定日'; Width = 11 }, @{ Header = '担当'; Width = 10 }, @{ Header = '状態'; Width = 26 }, @{ Header = 'ファイル'; Width = 50 }, @{ Header = '版数'; Width = 7 }); Rows = $rows })
    }
    if ($Previous) {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($c in $changeRows) {
            $p = $c.Previous
            $modified = ''
            if ($p -and $p.Modified) { $modified = $p.Modified.Value.ToString($p.Modified.Format, $invariant) }
            if ($c.Doc) {
                $d = $c.Doc
                $rows.Add(@{ Values = @($c.Change, $d.Folder, $d.File, $d.Title, $(if ($p) { $p.Version } else { '' }), $c.Version, $modified, $d.Modified.ToString('yyyy/MM/dd HH:mm', $invariant)); Link = @{ Column = 2; Path = $d.Path } })
            }
            else { $rows.Add(@{ Values = @($c.Change, $p.Folder, $p.File, $p.Title, $p.Version, '', $modified, '') }) }
        }
        $sheets.Add(@{ Name = '前回からの変化'; Columns = @(@{ Header = '変化'; Width = 30 }, @{ Header = 'フォルダ'; Width = 18 }, @{ Header = 'ファイル'; Width = 34 }, @{ Header = '文書名'; Width = 24 }, @{ Header = '前回の版数'; Width = 9 }, @{ Header = '今回の版数'; Width = 9 }, @{ Header = '前回の更新日時'; Width = 16 }, @{ Header = '今回の更新日時'; Width = 16 }); Rows = $rows })
    }
    if ($Findings) {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($f in $openFindings) {
            $place = $f.Place
            if ($place.Sheet) { $where = '{0} {1}!A{2}' -f $place.File, $place.Sheet, $place.Row } else { $where = '{0} {1} 行目' -f $place.File, $place.Row }
            $files = '（一覧に無い）'
            if ($f.Docs.Count -gt 0) { $files = (@($f.Docs | ForEach-Object { $_.Relative })) -join '・' }
            $row = @{ Values = @($files, $f.Document, $f.Content, $f.Severity, $f.Status, $(if ($f.Due) { $f.Due.ToString('yyyy/MM/dd', $invariant) } else { '' }), $(if ($f.Overdue) { '期限切れ' } else { '' }), $f.Owner, $where) }
            $row.Link = @{ Column = 8; Path = $place.Path }
            if ($place.Sheet) { $row.Link.Location = "'{0}'!A{1}" -f $place.Sheet.Replace("'", "''"), $place.Row }
            $rows.Add($row)
        }
        $sheets.Add(@{ Name = '未完了の指摘'; Columns = @(@{ Header = 'ファイル'; Width = 34 }, @{ Header = '設計書（指摘管理表）'; Width = 24 }, @{ Header = '指摘内容'; Width = 50 }, @{ Header = '重要度'; Width = 8 }, @{ Header = '状態'; Width = 10 }, @{ Header = '期限'; Width = 11 }, @{ Header = '期限切れ'; Width = 9 }, @{ Header = '対応者'; Width = 10 }, @{ Header = '指摘管理表の場所'; Width = 34 }); Rows = $rows })
    }
    if (-not $AllVersions -and $old.Count -gt 0) {
        $sheets.Add(@{ Name = '古い版'; Columns = @(@{ Header = 'フォルダ'; Width = 18 }, @{ Header = 'ファイル'; Width = 34 }, @{ Header = '版数'; Width = 7 }, @{ Header = 'ファイルの更新日時'; Width = 16 }, @{ Header = '最新の版'; Width = 40 }); Rows = @($old | ForEach-Object { @{ Values = @($_.Folder, $_.File, $_.Version, $_.Modified.ToString('yyyy/MM/dd HH:mm', $invariant), $_.LatestFile); Link = @{ Column = 1; Path = $_.Path } } }) })
    }
    if ($Planned -or $Previous -or $Findings) { $sheets.Add((New-ReadLogSheet $readLog)) }
    try { Export-Workbook -FilePath $OutXlsx -Sheets $sheets.ToArray() }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

$kinds = @(@('Excel', 'Word', 'PDF', 'そのほか') | ForEach-Object { $k = $_; $n = @($listed | Where-Object { $_.Kind -eq $k }).Count; if ($n -gt 0) { '{0} {1}' -f $k, $n } })
Write-Host ''
Write-Host ('成果物 {0} 件（{1}）。古い版 {2} 件' -f $listed.Count, ($kinds -join '・'), $old.Count)
$checks = @($listed | Where-Object { @($_.Notes | Where-Object { $_ -ne '予定に無い' }).Count -gt 0 })
if ($checks.Count -gt 0) { Write-Host ('版数・更新日などの確認が要るもの {0} 件: {1}' -f $checks.Count, ((@($checks | Select-Object -First 5 | ForEach-Object { $_.File })) -join '・')) }
if ($Planned) {
    $missing = @($plannedRows | Where-Object { $_.State -ne '提出済み' })
    $overdue = @($plannedRows | Where-Object { $_.State -like '*過ぎて*' })
    Write-Host ('予定 {0} 件: 提出済み {1}・未提出 {2}（予定日を過ぎたもの {3}）。予定に無い成果物 {4} 件' -f $plannedRows.Count, ($plannedRows.Count - $missing.Count), $missing.Count, $overdue.Count, @($docs | Where-Object { $_.Latest -and -not $_.Planned }).Count)
    foreach ($m in $missing) { Write-Host ('  {0}: {1}' -f $m.State, $m.Name) }
}
if ($Previous) {
    $counts = [ordered]@{
        '新規'               = @($changeRows | Where-Object { $_.Change -eq '新規' }).Count
        '版数が変わった'     = @($changeRows | Where-Object { $_.Change -match '版数 ' }).Count
        '版数を上げずに更新' = @($changeRows | Where-Object { $_.Change -match '版数を上げずに更新' }).Count
        '更新'               = @($changeRows | Where-Object { $_.Change -match '(^|・)更新$' }).Count
        '名前・場所が変わった' = @($changeRows | Where-Object { $_.Change -match '名前・場所' }).Count
        '無くなった'         = @($changeRows | Where-Object { $_.Change -eq '無くなった' }).Count
    }
    $summary = @($counts.Keys | Where-Object { $counts[$_] -gt 0 } | ForEach-Object { '{0} {1}' -f $_, $counts[$_] })
    if ($summary.Count -eq 0) { $summary = @('変わったものはありません') }
    Write-Host ('前回の一覧（{0}。成果物 {1} 件）から: {2}' -f [System.IO.Path]::GetFileName($Previous), $previousCount, ($summary -join '・'))
    foreach ($c in @($changeRows | Where-Object { $_.Change -match '版数を上げずに更新|無くなった' })) {
        $name = $(if ($c.Doc) { $c.Doc.Relative } else { $c.Previous.Relative })
        Write-Host ('  {0}: {1}' -f $(if ($c.Doc) { '版数を上げずに更新' } else { '無くなった' }), $name)
    }
}
if ($Findings) {
    $overdueCount = @($openFindings | Where-Object { $_.Overdue }).Count
    $unmatched = @($openFindings | Where-Object { $_.Docs.Count -eq 0 })
    $withFindings = @($docs | Where-Object { $_.OpenFindings -gt 0 }).Count
    Write-Host ('未完了の指摘 {0} 件（期限切れ {1} 件）: 成果物 {2} 件' -f $openFindings.Count, $overdueCount, $withFindings)
    if ($unmatched.Count -gt 0) { Write-Host ('  一覧の成果物と対応付かない指摘 {0} 件: {1}' -f $unmatched.Count, ((@($unmatched | ForEach-Object { if ($_.Document) { $_.Document } else { '（設計書なし）' } } | Select-Object -Unique)) -join '・')) }
}
if ($OutCsv) { Write-Host "一覧を保存しました: $OutCsv" }
if ($OutXlsx) { Write-Host "一覧を保存しました: $OutXlsx" }

# .SYNOPSIS
#     フォルダの成果物（設計書・報告書など）から、文書番号・文書名・版数・更新日・最新の改訂の一覧（成果物一覧・納品物一覧）を作ります（Excel 不要）。
#
# .DESCRIPTION
#     フォルダのファイルを集め、Excel・Word の表紙と改訂履歴から文書番号・文書名・版数・作成日・更新日・作成者・最新の改訂を読みます。
#     あわせて次を確かめ、「確認」の列に書きます。
#
#       版数が違う: ファイル名・表紙・改訂履歴の最後の行の版数が違う
#       更新日が違う: 表紙の更新日と、改訂履歴の最後の行の日付が違う
#       古い版: 版数だけが違う同じ文書（設計書_v1.0.xlsx と設計書_v1.1.xlsx など）の古いもの（一覧には最新だけを載せる）
#       予定に無い・未提出: -Planned の成果物の予定と突き合わせる
#       版数を上げずに更新: -Previous の前回の成果物一覧から、ファイルは更新されたのに版数が同じ
#       未完了の指摘: -Findings の指摘管理表で、対応の終わっていない指摘の数
#
#     表紙は「表紙」という名前のシート（無ければ最初のシート）、Word は本文の表と「見出し：値」の段落です。
#     見出し（文書番号・文書名・版数・作成日・最終更新日・作成者・システム名）の右の値を読みます。
#     文書名の見出しが無ければ、「…書」「…一覧」で終わる文字（後ろの（ ）は除いて見る）を文書名にします。
#
# .PARAMETER Path
#     成果物のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER Planned
#     成果物の予定（Excel か CSV。成果物名・文書名・ファイル名のどれかの見出しと、あれば文書番号・予定日・担当の見出しがある表）。
#     文書番号か名前で成果物と突き合わせ、未提出（予定日を過ぎたもの）と、予定に無い成果物を知らせます。
#
# .PARAMETER Previous
#     前回の成果物一覧（このツールの -OutXlsx か -OutCsv で保存したもの）。版数を除いた場所と名前で今回の成果物と突き合わせ、
#     前回からの変化（新規・版数 1.0→1.1・版数を上げずに更新・更新・名前・場所が変わった・無くなった）を書きます。
#     名前や場所を変えた文書は、文書番号で突き合わせます。更新は、版数の無いファイル（PDF など）が更新されたものです。
#
# .PARAMETER Findings
#     指摘管理表（Excel か CSV。ファイルかフォルダ。複数可）。対応の終わっていない指摘を、設計書（対象成果物）の列の
#     文書番号・文書名・ファイル名（版数と拡張子を除く）で成果物と突き合わせ、数を「確認」に書きます。
#     見出しと、対応が終わった・数えない状態（完了・済・対応不要・取下げなど）は Measure-ReviewQuality.ps1 と同じです。
#
# .PARAMETER AsOf
#     予定日と、指摘の期限を過ぎたかを決める基準日。既定は今日。
#
# .PARAMETER AllVersions
#     古い版も一覧に載せます（「確認」に古い版と最新のファイルを書きます）。
#
# .PARAMETER HistorySheet
#     改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#
# .PARAMETER OutXlsx
#     一覧を xlsx に保存します（成果物一覧（ファイルへのリンク付き）・予定との突き合わせ・前回からの変化・未完了の指摘・古い版）。
#
# .PARAMETER OutCsv
#     一覧を CSV（UTF-8・BOM 付き）に保存します。社内の様式に貼るとき用です。-Previous のときは「前回から」の列も足します。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の表紙は読まなくなります）。
#
# .EXAMPLE
#     .\New-DocumentRegister.ps1 D:\納品物 -Planned D:\成果物の予定.xlsx -OutXlsx D:\成果物一覧.xlsx
#
# .EXAMPLE
#     .\New-DocumentRegister.ps1 D:\納品物 -Previous D:\成果物一覧_前回.xlsx -Findings D:\レビュー -OutXlsx D:\成果物一覧.xlsx
#     前回の提出からの変化（版数を上げずに直した文書・無くなった文書）と、未完了の指摘のある成果物を知らせます。
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm / .docx は Excel・Word 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     PDF はページ数だけを数えます。.doc（古い Word）とそのほかのファイルは、名前・大きさ・更新日時だけを載せます。
