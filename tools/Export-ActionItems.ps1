[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $OutXlsx,

    [string] $OutCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

foreach ($name in @('OutXlsx', 'OutCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'ActionItem' -DefaultDisplayPropertySet Kind, Title, Owner, Due, Meeting -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$wordNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
$date = '(?:\d{4}\s*[/\-.年]\s*)?\d{1,2}\s*[/月]\s*\d{1,2}\s*日?(?:\s*\([月火水木金土日]\))?'
$ignoreCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

# ---------------------------------------------------------------- 議事録を行にする


function New-Line {
    #     Row: Excel の行番号（表の終わりを、行が飛んだことで見分ける）。ほかは 0
    param([string] $Text, $Cells, $Place, [int] $Row = 0)
    return [pscustomobject]@{ Text = $Text; Cells = $Cells; Place = $Place; Row = $Row }
}

function Read-WordLines {
    #     Word（.docx）の段落と表の行を、文書の順に行にする（表の行はセルの並びも持つ）
    param([string] $File)
    $lines = New-Object System.Collections.Generic.List[object]
    $pkg = Open-ZipPackage $File
    try { $doc = Read-PartXml $pkg 'word/document.xml' }
    finally { $pkg.Zip.Dispose() }
    if (-not $doc) { return , $lines }
    $paragraphText = {
        param($p)
        $sb = New-Object System.Text.StringBuilder
        foreach ($n in $p.SelectNodes('.//*')) {
            if ($n.NamespaceURI -ne $wordNs) { continue }
            if ($n.LocalName -eq 't') { [void]$sb.Append($n.InnerText) }
            elseif ($n.LocalName -eq 'tab') { [void]$sb.Append(' ') }
            elseif ($n.LocalName -in @('br', 'cr')) { [void]$sb.Append("`n") }
        }
        $sb.ToString()
    }
    $body = $doc.DocumentElement.FirstChild
    while ($body -and $body.LocalName -ne 'body') { $body = $body.NextSibling }
    if (-not $body) { return , $lines }
    $para = 0; $table = 0
    foreach ($child in $body.ChildNodes) {
        if ($child.LocalName -eq 'p') {
            $para++
            foreach ($t in ((& $paragraphText $child) -split "`n")) { $lines.Add((New-Line $t $null ([pscustomobject]@{ File = (Split-Path -Leaf $File); Path = $File; Sheet = ''; Cell = ('{0} 段落目' -f $para) }))) }
        }
        elseif ($child.LocalName -eq 'tbl') {
            $table++; $row = 0
            foreach ($tr in $child.ChildNodes) {
                if ($tr.LocalName -ne 'tr') { continue }
                $row++
                $cells = @(foreach ($tc in $tr.ChildNodes) {
                        if ($tc.LocalName -ne 'tc') { continue }
                        (@(foreach ($p in $tc.ChildNodes) { if ($p.LocalName -eq 'p') { & $paragraphText $p } }) -join ' ').Trim()
                    })
                $lines.Add((New-Line ($cells -join ' ') $cells ([pscustomobject]@{ File = (Split-Path -Leaf $File); Path = $File; Sheet = ''; Cell = ('表{0} の {1} 行目' -f $table, $row) })))
            }
            $lines.Add((New-Line '' $null $null))   # 表の終わり
        }
    }
    return , $lines
}

function Read-TextLines {
    param([string] $File)
    $lines = New-Object System.Collections.Generic.List[object]
    $n = 0
    foreach ($t in ((Read-TextFile $File) -split "\r?\n")) {
        $n++
        # Markdown の表（| 内容 | 担当 | 期限 |）はセルに分ける。区切りの行（|---|）は読まない
        $cells = $null
        if ($t -match '^\s*\|.*\|\s*$') {
            if ($t -match '^\s*\|[\s:\-|]+\|\s*$') { continue }
            $cells = @($t.Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() })
        }
        $lines.Add((New-Line $t $cells ([pscustomobject]@{ File = (Split-Path -Leaf $File); Path = $File; Sheet = ''; Cell = ('{0} 行目' -f $n) })))
    }
    return , $lines
}

# 対象のファイル（Excel・Word・テキスト）
$excelPaths = New-Object System.Collections.Generic.List[string]
$otherFiles = New-Object System.Collections.Generic.List[string]
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        $files = @($item)
        if ($item.PSIsContainer) { $files = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) }
        foreach ($f in $files) {
            if ($f.Name.StartsWith('~$')) { continue }
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($OutCsv -and [string]::Equals($f.FullName, $OutCsv, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($f.Extension -match '^\.xls[xmb]?$') { $excelPaths.Add([System.Management.Automation.WildcardPattern]::Escape($f.FullName)) }
            elseif ($f.Extension -match '^\.(docx|txt|md)$') { $otherFiles.Add($f.FullName) }
        }
    }
}

# ファイル → 行の一覧（Excel はシートの行ごとに、セルを左から並べる）
$documents = New-Object System.Collections.Generic.List[object]
if ($excelPaths.Count -gt 0) {
    $searchArgs = @{ Pattern = '\S'; Path = $excelPaths.ToArray(); Target = 'Value'; NoRecurse = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $books = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
    foreach ($book in $books.Values) {
        $lines = New-Object System.Collections.Generic.List[object]
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            foreach ($row in $grid.Rows) {
                $cols = @($grid.Cols[$row])
                # 空のセルは読み取り結果に無いため、列の位置どおりに並べる（表の見出しと列を合わせる）
                $cells = @(for ($c = 1; $c -le $cols[-1]; $c++) { [string]$grid.Cells[$row][$c] })
                $lines.Add((New-Line ((@($cells | Where-Object { $_ })) -join ' ') $cells ([pscustomobject]@{ File = $book.File; Path = $book.Path; Sheet = $sheetName; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $cols[0]), $row) }) $row))
            }
            $lines.Add((New-Line '' $null $null))
        }
        $documents.Add([pscustomobject]@{ File = $book.File; Path = $book.Path; Lines = $lines })
    }
}
foreach ($f in $otherFiles) {
    try {
        if ($f -match '\.docx$') { $lines = Read-WordLines $f } else { $lines = Read-TextLines $f }
        $documents.Add([pscustomobject]@{ File = (Split-Path -Leaf $f); Path = $f; Lines = $lines })
    }
    catch { Write-Warning ('読めませんでした: {0} : {1}' -f $f, $_.Exception.Message) }
}

# ---------------------------------------------------------------- 宿題・決定事項を見つける


function ConvertTo-DueDate {
    #     期限の文字を日付にする。年の無い日付は会議の日の年（会議の日より 1 か月以上前になるときは翌年）
    param([string] $Text, [datetime] $MeetingDate)
    $d = ConvertTo-DateValue ($Text.Normalize($nfkc) -replace '\s*\([月火水木金土日]\)\s*$', '') $MeetingDate
    if ($d -and $Text -notmatch '\d{4}' -and $d -lt $MeetingDate.AddDays(-31)) { $d = $d.AddYears(1) }
    return $d
}

function Get-Friday {
    param([datetime] $Day, [int] $Weeks)
    $monday = $Day.Date.AddDays(-(([int]$Day.DayOfWeek + 6) % 7))
    return $monday.AddDays(4 + 7 * $Weeks)
}

function Split-ActionText {
    #     宿題の文から、担当・期限を取り出し、残りを件名にする
    param([string] $Text, [datetime] $MeetingDate)
    $t = ([string]$Text).Normalize($nfkc).Trim()
    $owner = ''; $dueText = ''; $note = ''
    $take = {
        param([string] $Pattern)
        $m = [regex]::Match($script:work, $Pattern, $ignoreCase)
        if ($m.Success) { $script:work = $script:work.Remove($m.Index, $m.Length).Insert($m.Index, ' ') }
        return $m
    }
    $script:work = $t
    # 先頭の [山田]
    $m = & $take '^\s*\[(?<o>[^\]\d]{1,10})\]'
    if ($m.Success) { $owner = $m.Groups['o'].Value }
    # （山田、10/30まで）・（担当: 山田 / 期限: 10/30）
    $m = & $take ('\(\s*(?:担当(?:者)?\s*[:：]?\s*)?(?<o>[^\s(),、:/\d][^\s(),、:/]{0,9}?)(?:さん|様)?\s*[、,/]\s*(?:期限\s*[:：]?\s*)?(?<d>' + $date + ')\s*(?:まで|迄)?\s*\)')
    if ($m.Success) { if (-not $owner) { $owner = $m.Groups['o'].Value }; $dueText = $m.Groups['d'].Value }
    if (-not $owner) {
        $m = & $take '(?:担当(?:者)?|対応者)\s*[:：]\s*(?<o>[^\s(),、/]+?)(?:さん|様)?(?=$|[\s(),、/]|期限)'
        if ($m.Success) { $owner = $m.Groups['o'].Value }
    }
    if (-not $owner) {
        $m = & $take '[@＠](?<o>[^\s(),、/]+)'
        if ($m.Success) { $owner = $m.Groups['o'].Value }
    }
    if (-not $dueText) {
        $m = & $take ('(?:期限|期日|締切|締め切り|〆切)\s*[:：]?\s*(?<d>' + $date + ')')
        if (-not $m.Success) { $m = & $take ('\(\s*(?<d>' + $date + ')\s*(?:まで|迄)?\s*\)') }
        if (-not $m.Success) { $m = & $take ('(?<d>' + $date + ')\s*(?:まで|迄)(?:に)?') }
        if ($m.Success) { $dueText = $m.Groups['d'].Value }
    }
    if (-not $owner) {
        # 「山田さんが確認する」「(鈴木さん)」
        $m = & $take '\(\s*(?<o>[^\s(),、/\d]{1,8}?)(?:さん|様)\s*\)'
        # 文の中の「様」は仕様・様式などの言葉に多いため、「さん」だけを見る
        if (-not $m.Success) { $m = [regex]::Match($script:work, '(?<o>[^\s、,。をはがにでとのへもや()\[\]:]{1,6})さん') }
        # 「お客様」「皆様」などは担当にしない
        if ($m.Success -and $m.Groups['o'].Value -notmatch '^(お客|客|皆|先方|各位|ご担当者|担当者)$') { $owner = $m.Groups['o'].Value }
    }
    $due = $null
    if ($dueText) { $due = ConvertTo-DueDate $dueText $MeetingDate }
    else {
        $m = & $take '(?:本日中|今日中)'
        if ($m.Success) { $due = $MeetingDate.Date }
        else {
            $m = & $take '明日(?:中|まで)?'
            if ($m.Success) { $due = $MeetingDate.Date.AddDays(1) }
            else {
                $m = & $take '今週(?:中|末)(?:に|まで)?'
                if ($m.Success) { $due = Get-Friday $MeetingDate 0 }
                else {
                    $m = & $take '来週(?:中|末)(?:に|まで)?'
                    if ($m.Success) { $due = Get-Friday $MeetingDate 1 }
                    else {
                        $m = & $take '次回(?:の会議|の定例|会議|定例)?(?:まで|までに)'
                        if ($m.Success) { $note = '次回まで' }
                    }
                }
            }
        }
    }
    $title = $script:work -replace '\(\s*\)', '' -replace '\s{2,}', ' '
    $title = $title.Trim().TrimEnd('。', '、', ',', ':', '：', '/').Trim()
    return [pscustomobject]@{ Title = $title; Owner = $owner.Trim(); Due = $due; Note = $note }
}

# 節の見出しの印（● ○ ・ は箇条書きの印に使われるため、見出しの印にしない）
$headingPattern = '^\s*(?:[■□◆◇▼▽▶#]+\s*|\d{1,2}\s*[\.．、)]\s*|[【\[]\s*)?(?<k>宿題|ToDo|To Do|アクションアイテム|アクション|AI|課題|次回までの対応|対応事項|持ち帰り事項|持ち帰り|決定事項|決定|合意事項)\s*(?:一覧)?\s*[】\]]?\s*[:：]?\s*$'
$markerPattern = '[【\[](?<k>宿題|TODO|ToDo|AI|アクション|課題|決定|決定事項)[】\]]'
$prefixPattern = '^\s*(?<k>TODO|ToDo|宿題|AI|アクション|決定)\s*[:：]\s*'
$bulletPattern = '^\s*(?:[・\-\*●○◯◦▪•‐－–—>]+|\d{1,3}\s*[\.．)]|\(\d{1,3}\)|[①-⑳])\s*'

function Get-KindOf {
    param([string] $Word)
    if ($Word -match '決定|合意') { return '決定事項' }
    return '宿題'
}

function Find-MeetingInfo {
    #     会議の日（日時・開催日の行、ファイル名の日付、最初のほうの年のある日付、ファイルの更新日の順）と会議名
    param($Document)
    $head = @($Document.Lines | Select-Object -First 40 | ForEach-Object { ([string]$_.Text).Normalize($nfkc) })
    $day = $null; $name = ''
    foreach ($t in $head) {
        $m = [regex]::Match($t, '(?:日時|開催日時|開催日|日付|実施日|会議日)\s*[:：]?\s*(?<d>\d{4}\s*[/\-.年]\s*\d{1,2}\s*[/\-.月]\s*\d{1,2})')
        if ($m.Success) { $day = ConvertTo-DateValue $m.Groups['d'].Value; break }
    }
    if (-not $day) {
        $m = [regex]::Match($Document.File, '(20\d{2})[-_.]?(\d{2})[-_.]?(\d{2})')
        if ($m.Success) { try { $day = New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value) } catch { } }
    }
    if (-not $day) {
        foreach ($t in @($head | Select-Object -First 10)) {
            $m = [regex]::Match($t, '\d{4}\s*[/\-.年]\s*\d{1,2}\s*[/\-.月]\s*\d{1,2}')
            if ($m.Success) { $day = ConvertTo-DateValue $m.Value; break }
        }
    }
    if (-not $day) { $day = (Get-Item -LiteralPath $Document.Path).LastWriteTime.Date }
    foreach ($t in $head) {
        $m = [regex]::Match($t, '^\s*(?:会議名|件名|会議)\s*[:：]\s*(?<n>.+?)\s*$')
        if ($m.Success) { $name = $m.Groups['n'].Value; break }
    }
    if (-not $name) { $name = [System.IO.Path]::GetFileNameWithoutExtension($Document.File) }
    return [pscustomobject]@{ Date = $day.Date; Name = $name }
}

$results = New-Object System.Collections.Generic.List[object]
function Add-Action {
    param([string] $Kind, $Parsed, $Meeting, $Line, [string] $Original)
    if (-not $Parsed.Title) { return }
    if ($Parsed.Title -match '^(なし|特になし|無し|特に無し|-|―|ー)$') { return }
    $notes = New-Object System.Collections.Generic.List[string]
    if ($Kind -eq '宿題') {
        if (-not $Parsed.Owner) { $notes.Add('担当なし') }
        if (-not $Parsed.Due) { if ($Parsed.Note) { $notes.Add($Parsed.Note) } else { $notes.Add('期限なし') } }
    }
    $item = [pscustomobject]@{
        Kind = $Kind; Title = $Parsed.Title; Owner = $Parsed.Owner; Due = $Parsed.Due; Note = ($notes -join '・')
        Meeting = $Meeting.Name; MeetingDate = $Meeting.Date; Place = $Line.Place; Original = $Original.Trim()
    }
    $item.PSObject.TypeNames.Insert(0, 'ActionItem')
    $results.Add($item)
    return $item
}

# 表の見出し（宿題の一覧を表で書いた議事録）。ConvertTo-Label して大文字にしたもの
$tableHeads = @{}
foreach ($pair in @(
        @('Content', @('内容', '件名', '項目', 'アクション', 'アクションアイテム', '宿題', 'TODO', '対応内容', '課題', '作業', '決定事項', '内容・対応')),
        @('Owner', @('担当', '担当者', '対応者', 'アクション担当', '担当者名')),
        @('Due', @('期限', '期日', '完了予定日', '対応期限', '予定日', '締切', '回答期限')),
        @('Kind', @('区分', '種別', '種類')))) {
    $tableHeads[$pair[0]] = @($pair[1] | ForEach-Object { (ConvertTo-Label $_).ToUpperInvariant() })
}
function Find-TableMap {
    param([string[]] $Cells)
    $map = @{}
    for ($i = 0; $i -lt $Cells.Count; $i++) {
        $label = (ConvertTo-Label $Cells[$i]).ToUpperInvariant()
        foreach ($k in $tableHeads.Keys) { if (-not $map.ContainsKey($k) -and $tableHeads[$k] -contains $label) { $map[$k] = $i } }
    }
    if ($map.ContainsKey('Content') -and ($map.ContainsKey('Owner') -or $map.ContainsKey('Due'))) { return $map }
    return $null
}

foreach ($doc in $documents) {
    $meeting = Find-MeetingInfo $doc
    $section = ''
    $tableMap = $null
    $last = $null
    $prevRow = 0
    foreach ($line in $doc.Lines) {
        $raw = [string]$line.Text
        $norm = $raw.Normalize($nfkc)
        $isHeading = $norm -match '^\s*([■□◆◇▼▽▶#]|【[^】]*】\s*$|以上\s*$)' -or [regex]::IsMatch($norm, $headingPattern, $ignoreCase)
        # 表の終わり: 表でない行・見出し・Excel で行が飛んだとき
        if (-not $line.Cells -or $isHeading -or ($line.Row -and $line.Row -ne $prevRow + 1)) { $tableMap = $null }
        $prevRow = $line.Row
        # 表（Word の表・Excel の行・Markdown の表）: 担当・期限の見出しのある表は、行をそのまま宿題にする
        if ($line.Cells -and -not $isHeading) {
            $map = Find-TableMap $line.Cells
            if ($map) { $tableMap = $map; continue }
            if ($tableMap) {
                $cell = { param($k) if ($tableMap.ContainsKey($k) -and $tableMap[$k] -lt $line.Cells.Count) { ([string]$line.Cells[$tableMap[$k]]).Trim() } else { '' } }
                $content = & $cell 'Content'
                if (-not $content) { continue }
                $kind = '宿題'
                if ((& $cell 'Kind') -match '決定|合意') { $kind = '決定事項' }
                elseif ($section -eq '決定事項') { $kind = '決定事項' }
                $dueCell = & $cell 'Due'
                $parsed = Split-ActionText $content $meeting.Date
                $ownerCell = (& $cell 'Owner') -replace '(さん|様)$', ''
                if ($ownerCell) { $parsed.Owner = $ownerCell }
                if ($dueCell) { $d = ConvertTo-DueDate $dueCell $meeting.Date; if ($d) { $parsed.Due = $d } elseif ($dueCell -match '次回') { $parsed.Note = '次回まで' } }
                $last = Add-Action $kind $parsed $meeting $line $raw
                continue
            }
        }
        if (-not $norm.Trim()) { continue }
        # 見出し（宿題・決定事項の節の始まり、ほかの節で終わり）
        $h = [regex]::Match($norm, $headingPattern, $ignoreCase)
        if ($h.Success) { $section = Get-KindOf $h.Groups['k'].Value; $last = $null; continue }
        if ($isHeading) { $section = ''; $last = $null; continue }
        # 文の中の印（【宿題】・TODO: など）
        $m = [regex]::Match($norm, $markerPattern, $ignoreCase)
        if ($m.Success) {
            # 印の後ろに文があれば、それが宿題（「次回は 10/27。【TODO】議事録を共有する」）
            $text = $norm.Substring($m.Index + $m.Length).Trim()
            if (-not $text) { $text = $norm.Substring(0, $m.Index) }
            $text = $text -replace $bulletPattern, ''
            $last = Add-Action (Get-KindOf $m.Groups['k'].Value) (Split-ActionText $text $meeting.Date) $meeting $line $raw
            continue
        }
        $m = [regex]::Match($norm, $prefixPattern, $ignoreCase)
        if ($m.Success) {
            $last = Add-Action (Get-KindOf $m.Groups['k'].Value) (Split-ActionText $norm.Substring($m.Length) $meeting.Date) $meeting $line $raw
            continue
        }
        if (-not $section) { continue }
        # 宿題・決定事項の節の中: 字下げした ※・→ の行は、前の項目の補足
        if ($last -and $norm -match '^\s+(?:※|→)\s*(?<n>.+)$') {
            $last.Note = (@($last.Note, $Matches['n'].Trim()) | Where-Object { $_ }) -join '・'
            continue
        }
        $text = $norm -replace $bulletPattern, ''
        $last = Add-Action $section (Split-ActionText $text $meeting.Date) $meeting $line $raw
    }
}

$results

# ---------------------------------------------------------------- 保存・表示


$actions = @($results | Where-Object { $_.Kind -eq '宿題' })
$decisions = @($results | Where-Object { $_.Kind -eq '決定事項' })
function Format-DateCell {
    param($Date)
    if ($null -eq $Date) { return '' }
    return $Date.ToString('yyyy/MM/dd', [System.Globalization.CultureInfo]::InvariantCulture)
}
function Format-PlaceText {
    param($Place)
    if (-not $Place) { return '' }
    if ($Place.Sheet) { return '{0} {1}!{2}' -f $Place.File, $Place.Sheet, $Place.Cell }
    return '{0} {1}' -f $Place.File, $Place.Cell
}

if ($OutCsv) {
    $i = 0
    $rows = @($actions | ForEach-Object {
            $i++
            [pscustomobject][ordered]@{ 'No' = $i; '区分' = '宿題'; '件名' = $_.Title; '担当' = $_.Owner; '期限' = (Format-DateCell $_.Due); '状態' = '未対応'; '起票日' = (Format-DateCell $_.MeetingDate); '会議' = $_.Meeting; 'メモ' = $_.Note; '出典' = (Format-PlaceText $_.Place) }
        })
    $text = ''
    if ($rows.Count -gt 0) { $text = ((($rows | ConvertTo-Csv -NoTypeInformation) -join "`r`n") + "`r`n") }
    else { $text = '"No","区分","件名","担当","期限","状態","起票日","会議","メモ","出典"' + "`r`n" }
    [System.IO.File]::WriteAllText($OutCsv, $text, (New-Object System.Text.UTF8Encoding($true)))
}
if ($OutXlsx) {
    $actionRows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($a in $actions) {
        $i++
        $row = @{ Values = @([double]$i, '宿題', $a.Title, $a.Owner, (Format-DateCell $a.Due), '未対応', (Format-DateCell $a.MeetingDate), $a.Meeting, $a.Note, (Format-PlaceText $a.Place), $a.Original) }
        if ($a.Place.Sheet) { $row.Link = @{ Column = 9; Path = $a.Place.Path; Location = ("'{0}'!{1}" -f $a.Place.Sheet.Replace("'", "''"), $a.Place.Cell) } }
        $actionRows.Add($row)
    }
    $decisionRows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($d in $decisions) {
        $i++
        $row = @{ Values = @([double]$i, $d.Title, (Format-DateCell $d.MeetingDate), $d.Meeting, (Format-PlaceText $d.Place), $d.Original) }
        if ($d.Place.Sheet) { $row.Link = @{ Column = 4; Path = $d.Place.Path; Location = ("'{0}'!{1}" -f $d.Place.Sheet.Replace("'", "''"), $d.Place.Cell) } }
        $decisionRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '宿題'; Columns = @(@{ Header = 'No'; Width = 5 }, @{ Header = '区分'; Width = 6 }, @{ Header = '件名'; Width = 50 }, @{ Header = '担当'; Width = 12 }, @{ Header = '期限'; Width = 11 }, @{ Header = '状態'; Width = 8 }, @{ Header = '起票日'; Width = 11 }, @{ Header = '会議'; Width = 24 }, @{ Header = 'メモ'; Width = 20 }, @{ Header = '出典'; Width = 30 }, @{ Header = '原文'; Width = 50 }); Rows = $actionRows },
            @{ Name = '決定事項'; Columns = @(@{ Header = 'No'; Width = 5 }, @{ Header = '決定事項'; Width = 60 }, @{ Header = '日付'; Width = 11 }, @{ Header = '会議'; Width = 24 }, @{ Header = '出典'; Width = 30 }, @{ Header = '原文'; Width = 50 }); Rows = $decisionRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ('議事録 {0} 件: 宿題 {1} 件（担当なし {2}・期限なし {3}）・決定事項 {4} 件' -f $documents.Count, $actions.Count,
    @($actions | Where-Object { -not $_.Owner }).Count, @($actions | Where-Object { -not $_.Due }).Count, $decisions.Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
if ($OutCsv) { Write-Host "宿題を CSV に保存しました（課題管理表に貼り付けるか、New-WeeklyReport.ps1 の -IssueList に渡せます）: $OutCsv" }

# .SYNOPSIS
#     議事録（Word・Excel・テキスト）から宿題（アクションアイテム）と決定事項を抜き出し、担当・期限付きの一覧にします（Excel 不要）。
#
# .DESCRIPTION
#     会議のあとに、議事録を読み返して課題管理表へ宿題を書き写す作業を機械でします。次の書き方を宿題・決定事項として読みます。
#
#       節            「■宿題」「【ToDo】」「アクションアイテム」「決定事項」などの見出しの下の行（次の ■ や【】の見出しまで）
#       印            行の中の【宿題】【TODO】【AI】【決定】や、行頭の「TODO:」「宿題:」
#       表            内容（件名・項目・アクション）と、担当か期限の見出しがある表（Word の表・Excel）
#
#     担当と期限は、行の中の「（山田、10/30まで）」「担当: 山田」「期限: 10/30」「@山田」「山田さんが」「[山田]」などから取り出し、
#     残りを件名にします。「今週中」「来週中」は会議の週・翌週の金曜日、「次回まで」はメモに書きます。
#     会議の日は「日時・開催日」の行、ファイル名の日付（20261020 など）、ファイルの更新日の順に探します（年の無い期限の年に使います）。
#
#     -OutCsv の CSV（No・区分・件名・担当・期限・状態・起票日など）は、課題管理表に貼り付けるか、
#     New-WeeklyReport.ps1 の -IssueList に渡して、週次報告で期限切れを追えます。
#
# .PARAMETER Path
#     議事録のファイルまたはフォルダ（.docx・.xlsx・.xlsm・.txt・.md。複数可）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（宿題・決定事項。Excel の議事録は該当セルへのリンク付き）。
#
# .PARAMETER OutCsv
#     宿題を CSV（UTF-8）に保存します。
#
# .PARAMETER NoRecurse
#     フォルダのサブフォルダを読みません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\Export-ActionItems.ps1 D:\PJ\議事録\20261020_定例.docx -OutXlsx D:\PJ\議事録\宿題.xlsx
#
# .EXAMPLE
#     # 議事録のフォルダから宿題を集め、週次報告で追う
#     .\Export-ActionItems.ps1 D:\PJ\議事録 -OutCsv D:\PJ\宿題.csv
#     .\New-WeeklyReport.ps1 -Wbs D:\PJ\WBS.xlsx -IssueList D:\PJ\宿題.csv -OutDir D:\PJ\週次報告
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel・Word は不要（.xls / .xlsb の議事録だけ Excel が必要）。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     書き方の揺れが大きい議事録では、取り出せない・件名に担当が残るなどがあります。結果を確かめてから課題管理表に移してください。
