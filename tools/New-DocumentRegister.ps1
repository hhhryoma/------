[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $Planned,

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

foreach ($name in @('Planned', 'OutXlsx', 'OutCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($Planned -and -not (Test-Path -LiteralPath $Planned)) { throw "-Planned の成果物の予定が見つかりません: $Planned" }

Update-TypeData -TypeName 'DocumentRegister.Document' -DefaultDisplayPropertySet Folder, File, DocNo, Title, Version, Updated, Note -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

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
    #     版数の文字（「第1.1版」「Ver.1.10」「v1.1」など）を比べる形「1.1」にする（無ければ空）
    param([string] $Text)
    $m = [regex]::Match(([string]$Text).Normalize($nfkc), '(\d+(?:\.\d+)*)')
    if (-not $m.Success) { return '' }
    return (@($m.Groups[1].Value.Split('.') | ForEach-Object { [string][int64]$_ }) -join '.')
}

function ConvertTo-VersionKey {
    #     版数を比べる形（ConvertTo-VersionText の後ろの「.0」を除く。数値のセルの 1.0 は「1」と読めるので、1 と 1.0 を同じにする）
    param([string] $Version)
    return [regex]::Replace($Version, '(\.0)+$', '')
}

function ConvertTo-DateText {
    #     日付の文字を「2026/10/01」の形にする（読めなければそのまま）
    param([string] $Text)
    $d = ConvertTo-DateValue $Text $AsOf
    if ($d) { return $d.ToString('yyyy/MM/dd') }
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
            if (@($OutXlsx, $OutCsv, $Planned) | Where-Object { $_ -and [string]::Equals($_, $f.FullName, [System.StringComparison]::OrdinalIgnoreCase) }) { continue }
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
    }
    # 版数はファイル名、無ければフォルダ名（v1.0\設計書.xlsx など）
    if ($null -eq $doc.Order) { $doc.Order = Get-NameVersion $entry.Relative }
    $nm = [regex]::Match([System.IO.Path]::GetFileNameWithoutExtension($f.Name).Normalize($nfkc), '(?i)(?<![a-z])v(?:er)?\.?\s*(\d+(?:\.\d+)*)|第\s*(\d+(?:\.\d+)*)\s*版')
    if ($nm.Success) { $doc.NameVersion = ConvertTo-VersionText ($nm.Groups[1].Value + $nm.Groups[2].Value) }
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
    if (@($versions.Values | ForEach-Object { ConvertTo-VersionKey $_ } | Select-Object -Unique).Count -gt 1) {
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


$plannedRows = New-Object System.Collections.Generic.List[object]
$readLog = New-Object System.Collections.Generic.List[object]
if ($Planned) {
    $labels = { param([string[]] $Names) @($Names | ForEach-Object { ConvertTo-Label $_ }) }
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
        if ($docNo) { $matched = @($current | Where-Object { $_.DocNo -and (& $key $_.DocNo) -eq (& $key $docNo) }) }
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
        $plannedRows.Add([pscustomobject]@{ Name = $name; DocNo = $docNo; Due = $(if ($due) { $due.ToString('yyyy/MM/dd') } else { ([string]$rec.Values['予定日']).Trim() }); Owner = ([string]$rec.Values['担当']).Trim(); State = $state; Docs = $matched; Place = $rec })
    }
    foreach ($d in $current) { if (-not $d.Planned) { $d.Notes.Add('予定に無い') } }
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

$headers = @('No', 'フォルダ', 'ファイル', '種類', '文書番号', '文書名', '版数', '作成日', '更新日', '作成者', '最新の改訂', 'ページ', 'サイズ（KB）', 'ファイルの更新日時', '確認')
$registerRows = New-Object System.Collections.Generic.List[object]
$no = 0
foreach ($d in $listed) {
    $no++
    $history = ''
    if ($d.History) { $history = (@($d.History.Version, $d.History.Date, $d.History.Content, $d.History.Author) | Where-Object { $_ }) -join ' ' }
    $registerRows.Add([pscustomobject]@{ Doc = $d; Values = @([double]$no, $d.Folder, $d.File, $d.Kind, $d.DocNo, $d.Title, $d.Version, $d.Created, $d.Updated, $d.Author, $history, $d.Pages, [double]$d.Size, $d.Modified.ToString('yyyy/MM/dd HH:mm'), $d.Note) })
}

if ($OutCsv) {
    $lines = New-Object System.Collections.Generic.List[string]
    $quote = { param($v) $s = [string]$v; if ($s -match '[",\r\n]') { '"' + $s.Replace('"', '""') + '"' } else { $s } }
    $lines.Add((@($headers | ForEach-Object { & $quote $_ })) -join ',')
    foreach ($r in $registerRows) { $lines.Add((@($r.Values | ForEach-Object { & $quote $_ })) -join ',') }
    [System.IO.File]::WriteAllText($OutCsv, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

if ($OutXlsx) {
    $widths = @(5, 18, 34, 7, 14, 24, 7, 11, 11, 12, 34, 6, 9, 16, 40)
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
    if (-not $AllVersions -and $old.Count -gt 0) {
        $sheets.Add(@{ Name = '古い版'; Columns = @(@{ Header = 'フォルダ'; Width = 18 }, @{ Header = 'ファイル'; Width = 34 }, @{ Header = '版数'; Width = 7 }, @{ Header = 'ファイルの更新日時'; Width = 16 }, @{ Header = '最新の版'; Width = 40 }); Rows = @($old | ForEach-Object { @{ Values = @($_.Folder, $_.File, $_.Version, $_.Modified.ToString('yyyy/MM/dd HH:mm'), $_.LatestFile); Link = @{ Column = 1; Path = $_.Path } } }) })
    }
    if ($Planned) { $sheets.Add((New-ReadLogSheet $readLog)) }
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
# .PARAMETER AsOf
#     予定日を過ぎたかを決める基準日。既定は今日。
#
# .PARAMETER AllVersions
#     古い版も一覧に載せます（「確認」に古い版と最新のファイルを書きます）。
#
# .PARAMETER HistorySheet
#     改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#
# .PARAMETER OutXlsx
#     一覧を xlsx に保存します（成果物一覧（ファイルへのリンク付き）・予定との突き合わせ・古い版）。
#
# .PARAMETER OutCsv
#     一覧を CSV（UTF-8・BOM 付き）に保存します。社内の様式に貼るとき用です。
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
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm / .docx は Excel・Word 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     PDF はページ数だけを数えます。.doc（古い Word）とそのほかのファイルは、名前・大きさ・更新日時だけを載せます。
