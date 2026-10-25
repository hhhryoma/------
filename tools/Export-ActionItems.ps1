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
# 年のある日付（2026/10/30・2026-10-30・2026年10月30日・令和8年10月30日・R8.10.30）
$yearDate = '(?:\d{4}|(?:令和|平成|(?<![A-Za-z])[RH])\s*(?:元|\d{1,2}))\s*[/\-.年]\s*\d{1,2}\s*[/\-.月]\s*\d{1,2}'
# 期限の日付（年の無いものは 10/30・10月30日・10月末。後ろの曜日も）
$date = '(?:' + $yearDate + '|\d{1,2}\s*[/月]\s*\d{1,2}|\d{1,2}\s*月\s*末)\s*日?(?:\s*\([月火水木金土日]\))?'
# 会議の日から決める期限と「次回まで」。漢字に続く「明日」（説明日程・証明日）や「本日中」（基本日中）は読まない
$dueWord = '(?:(?<!\p{IsCJKUnifiedIdeographs})(?:(?<today>本日|今日)中|(?<tomorrow>明日)中?)|(?<week>今週|来週)(?:中|末|の?(?<day>[月火水木金土日])曜日?)|次回(?:の会議|の定例|会議|定例)?まで)'
$ignoreCase = [System.Text.RegularExpressions.RegexOptions]::IgnoreCase

# ---------------------------------------------------------------- 議事録を行にする


function New-Line {
    #     Row: Excel の行番号（表の終わりを、行が飛んだことで見分ける）。ほかは 0
    param([string] $Text, $Cells, $Place, [int] $Row = 0)
    return [pscustomobject]@{ Text = $Text; Cells = $Cells; Place = $Place; Row = $Row }
}

function Split-CellLines {
    #     Excel の行を、セルごと・セルの中の改行（Alt+Enter。読み取り結果では ↵）ごとの行にする（場所はそのセル）
    param($Line)
    $lines = New-Object System.Collections.Generic.List[object]
    for ($c = 0; $c -lt $Line.Cells.Count; $c++) {
        if (-not $Line.Cells[$c]) { continue }
        $place = [pscustomobject]@{ File = $Line.Place.File; Path = $Line.Place.Path; Sheet = $Line.Place.Sheet; Cell = ('{0}{1}' -f (ConvertTo-ColumnName ($c + 1)), $Line.Row) }
        foreach ($t in ([string]$Line.Cells[$c] -split '↵')) { $lines.Add((New-Line $t $null $place $Line.Row)) }
    }
    return , $lines
}

function Get-WordChildren {
    #     Word の要素の子（段落・表・行・セル）。内容コントロール（w:sdt）は、中身（w:sdtContent）の子をその場所に並べる
    param($Node)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($child in $Node.ChildNodes) {
        if ($child.LocalName -ne 'sdt') { $list.Add($child); continue }
        foreach ($c in $child.ChildNodes) { if ($c.LocalName -eq 'sdtContent') { $list.AddRange((Get-WordChildren $c)) } }
    }
    return , $list
}

function Read-WordLines {
    #     Word（.docx）の段落と表の行を、文書の順に行にする（表の行はセルの並びも持つ）。内容コントロールの中も読む。
    #     ExcelTools.psm1 の Get-WordBlocks は段落の文字の前後の空白を除く（字下げした ※・→ の行を見分けられない）ため使わない
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
    foreach ($child in (Get-WordChildren $body)) {
        if ($child.LocalName -eq 'p') {
            $para++
            foreach ($t in ((& $paragraphText $child) -split "`n")) { $lines.Add((New-Line $t $null ([pscustomobject]@{ File = (Split-Path -Leaf $File); Path = $File; Sheet = ''; Cell = ('{0} 段落目' -f $para) }))) }
        }
        elseif ($child.LocalName -eq 'tbl') {
            $table++; $row = 0
            foreach ($tr in (Get-WordChildren $child)) {
                if ($tr.LocalName -ne 'tr') { continue }
                $row++
                $cells = @(foreach ($tc in (Get-WordChildren $tr)) {
                        if ($tc.LocalName -ne 'tc') { continue }
                        (@(foreach ($p in (Get-WordChildren $tc)) { if ($p.LocalName -eq 'p') { & $paragraphText $p } }) -join ' ').Trim()
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
    #     期限の文字を日付にする。年の無い日付（10/30・10月30日・10月末）は会議の日の年（会議の日より 1 か月以上前になるときは翌年）。
    #     本日中・明日は会議の日・翌日、今週中・来週中は会議の週・翌週の金曜日（来週水曜などはその曜日）、10月末は月の最後の日
    #     （表のセルは、セル全体がこの書き方のときだけ）。「次回まで」などは $null
    param([string] $Text, [datetime] $MeetingDate)
    $t = $Text.Normalize($nfkc).Trim() -replace '\s*\([月火水木金土日]\)\s*$', ''
    $m = [regex]::Match($t, '^(?:' + $dueWord + '|(?<month>\d{1,2})\s*月\s*末日?)\s*(?:まで|迄)?\s*(?:に)?$')
    if ($m.Groups['today'].Success) { return $MeetingDate.Date }
    if ($m.Groups['tomorrow'].Success) { return $MeetingDate.Date.AddDays(1) }
    if ($m.Groups['week'].Success) {
        $weeks = 0
        if ($m.Groups['week'].Value -eq '来週') { $weeks = 1 }
        $day = 4
        if ($m.Groups['day'].Success) { $day = '月火水木金土日'.IndexOf($m.Groups['day'].Value[0]) }
        return Get-WeekDay $MeetingDate $weeks $day
    }
    if ($m.Groups['month'].Success) {
        $d = ConvertTo-DateValue ($m.Groups['month'].Value + '/1') $MeetingDate
        if ($d) { $d = $d.AddMonths(1).AddDays(-1) }
    }
    else { $d = ConvertTo-DateValue $t $MeetingDate }
    if ($d -and $t -match '^\d{1,2}\s*[/月]' -and $d -lt $MeetingDate.AddDays(-31)) { $d = $d.AddYears(1) }
    return $d
}

function Get-WeekDay {
    #     会議の週（Weeks = 0）・翌週（1）の曜日（DayOfWeek: 月曜日が 0、金曜日が 4）の日付
    param([datetime] $Day, [int] $Weeks, [int] $DayOfWeek)
    $monday = $Day.Date.AddDays(-(([int]$Day.DayOfWeek + 6) % 7))
    return $monday.AddDays($DayOfWeek + 7 * $Weeks)
}

function Join-OwnerNames {
    #     担当の名前（「山田、佐藤」「山田さん、佐藤さん」）を、さん・様を除いて「、」でつなぐ
    param([string] $Text)
    return (@($Text -split '\s*[、,]\s*' | ForEach-Object { ($_ -replace '(さん|様)$', '').Trim() } | Where-Object { $_ }) -join '、')
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
    # （山田、10/30まで）・（山田、来週中）・（担当: 山田 / 期限: 10/30）
    $m = & $take ('\(\s*(?:担当(?:者)?\s*[:：]?\s*)?(?<o>[^\s(),、:/\d][^\s(),、:/]{0,9}?)(?:さん|様)?\s*[、,/]\s*(?:期限\s*[:：]?\s*)?(?<d>' + $date + '|' + $dueWord + ')\s*(?:まで|迄)?\s*\)')
    if ($m.Success) { if (-not $owner) { $owner = $m.Groups['o'].Value }; $dueText = $m.Groups['d'].Value }
    if (-not $owner) {
        # 担当: 山田・担当: 山田、佐藤（2 人目からは、日付と期限の言葉を名前にしない）
        $m = & $take '(?:担当(?:者)?|対応者)\s*[:：]\s*(?<o>[^\s(),、/]+?(?:さん|様)?(?:\s*[、,]\s*(?!期限|期日|締切|本日|今日|明日|今週|来週|次回)[^\s(),、/:：\d]{1,10}?(?:さん|様)?)*)(?=$|[\s(),、/]|期限)'
        if ($m.Success) { $owner = Join-OwnerNames $m.Groups['o'].Value }
    }
    if (-not $owner) {
        $m = & $take '[@＠](?<o>[^\s(),、/]+)'
        if ($m.Success) { $owner = $m.Groups['o'].Value }
    }
    if (-not $dueText) {
        $m = & $take ('(?:期限|期日|締切|締め切り|〆切)\s*[:：]?\s*(?<d>' + $date + '|' + $dueWord + ')')
        if (-not $m.Success) { $m = & $take ('\(\s*(?<d>' + $date + ')\s*(?:まで|迄)?\s*\)') }
        if (-not $m.Success) { $m = & $take ('(?<d>' + $date + ')\s*(?:まで|迄)(?:に)?') }
        if ($m.Success) { $dueText = $m.Groups['d'].Value }
    }
    if (-not $owner) {
        # 「山田さんが確認する」「(鈴木さん)」「(山田さん、佐藤さん)」
        $m = & $take '\(\s*(?<o>[^\s(),、/\d]{1,8}?(?:さん|様)(?:\s*[、,]\s*[^\s(),、/\d]{1,8}?(?:さん|様))*)\s*\)'
        # 文の中の「様」は仕様・様式などの言葉に多いため、「さん」だけを見る
        if (-not $m.Success) { $m = [regex]::Match($script:work, '(?<o>[^\s、,。をはがにでとのへもや()\[\]:]{1,6})さん') }
        # 「お客様」「皆様」などは担当にしない
        if ($m.Success) { $o = Join-OwnerNames $m.Groups['o'].Value; if ($o -notmatch '^(お客|客|皆|先方|各位|ご担当者|担当者)$') { $owner = $o } }
    }
    # 日付の無い期限（本日中・明日・今週中・来週中・来週金曜・次回まで）
    if (-not $dueText) {
        $m = & $take ('(?<d>' + $dueWord + ')(?:まで|迄)?(?:に)?')
        if ($m.Success) { $dueText = $m.Groups['d'].Value }
    }
    $due = $null
    if ($dueText) { $due = ConvertTo-DueDate $dueText $MeetingDate }
    if (-not $due -and $dueText -match '^次回') { $note = '次回まで' }
    $title = $script:work -replace '\(\s*\)', '' -replace '\s{2,}', ' '
    $title = $title.Trim().TrimEnd('。', '、', ',', ':', '：', '/').Trim()
    return [pscustomobject]@{ Title = $title; Owner = $owner.Trim(); Due = $due; Note = $note }
}

# 節の見出しの印（● ○ ・ は箇条書きの印に使われるため、見出しの印にしない）。mark は、節の見出しと同じ印の行で節を終えるために使う
$headingPattern = '^\s*(?:(?<mark>[■□◆◇▼▽▶#])[■□◆◇▼▽▶#]*\s*|\d{1,2}\s*[\.．、)]\s*|[【\[]\s*)?(?<k>宿題|ToDo|To Do|アクションアイテム|アクション|AI|課題|次回までの対応|対応事項|持ち帰り事項|持ち帰り|決定事項|決定|合意事項)\s*(?:一覧)?\s*[】\]]?\s*[:：]?\s*$'
$markerPattern = '[【\[](?<k>宿題|TODO|ToDo|AI|アクション|課題|決定|決定事項)[】\]]'
$prefixPattern = '^\s*(?<k>TODO|ToDo|宿題|AI|アクション|決定)\s*[:：]\s*'
# 箇条書きの印（□ ☐ ◇ ▽ ▶ も。■ ◆ ▼ は見出しの印）
$bulletPattern = '^\s*(?:[・\-\*●○◯◦▪•‐－–—>□☐◇▽▶]+|\d{1,3}\s*[\.．)]|\(\d{1,3}\)|[①-⑳])\s*'

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
        # 年のある日付（和暦も）
        $m = [regex]::Match($t, '(?:日時|開催日時|開催日|日付|実施日|会議日)\s*[:：]?\s*(?<d>' + $yearDate + ')')
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
    #     Status: 表の状況・状態の列の文字（無ければ、宿題は「未対応」）。Opened は起票日（会議の日。同じ宿題をまとめたときは最初の会議の日）
    param([string] $Kind, $Parsed, $Meeting, $Line, [string] $Original, [string] $Status)
    if (-not $Parsed.Title) { return }
    if ($Parsed.Title -match '^(なし|特になし|無し|特に無し|-|―|ー)$') { return }
    $notes = New-Object System.Collections.Generic.List[string]
    if ($Kind -eq '宿題') {
        if (-not $Parsed.Owner) { $notes.Add('担当なし') }
        if (-not $Parsed.Due) { if ($Parsed.Note) { $notes.Add($Parsed.Note) } else { $notes.Add('期限なし') } }
        if (-not $Status) { $Status = '未対応' }
    }
    $item = [pscustomobject]@{
        Kind = $Kind; Title = $Parsed.Title; Owner = $Parsed.Owner; Due = $Parsed.Due; Note = ($notes -join '・'); Status = $Status
        Meeting = $Meeting.Name; MeetingDate = $Meeting.Date; Opened = $Meeting.Date; Place = $Line.Place; Original = $Original.Trim()
    }
    $item.PSObject.TypeNames.Insert(0, 'ActionItem')
    $results.Add($item)
    return $item
}

function Add-OwnerDue {
    #     担当・期限だけの行（「　担当：山田　期限：10/23」）を、前の宿題の担当・期限にする（前の宿題に無いものだけ）。メモの担当なし・期限なしも直す
    param($Item, $Parsed)
    $notes = @($Item.Note -split '・')
    if ($Parsed.Owner -and -not $Item.Owner) { $Item.Owner = $Parsed.Owner; $notes = @($notes | Where-Object { $_ -ne '担当なし' }) }
    if (($Parsed.Due -or $Parsed.Note) -and $notes -contains '期限なし') {
        $Item.Due = $Parsed.Due
        $notes = @($notes | ForEach-Object { if ($_ -ne '期限なし') { $_ } elseif ($Parsed.Note) { $Parsed.Note } })
    }
    $Item.Note = (@($notes | Where-Object { $_ })) -join '・'
}

# 表の見出し（宿題の一覧を表で書いた議事録）。ConvertTo-Label して大文字にしたもの
$tableHeads = @{}
foreach ($pair in @(
        @('Content', @('内容', '件名', '項目', 'アクション', 'アクションアイテム', '宿題', 'TODO', '対応内容', '課題', '作業', '決定事項', '内容・対応')),
        @('Owner', @('担当', '担当者', '対応者', 'アクション担当', '担当者名')),
        @('Due', @('期限', '期日', '完了予定日', '対応期限', '予定日', '締切', '回答期限')),
        @('Status', @('状況', '状態', '対応状況', '対応状態', 'ステータス')),
        @('Kind', @('区分', '種別', '種類')))) {
    $tableHeads[$pair[0]] = @($pair[1] | ForEach-Object { (ConvertTo-Label $_).ToUpperInvariant() })
}
function Find-TableMap {
    #     見出しのセルの改行（Excel の「対応↵期限」）は無いものとして比べる
    param([string[]] $Cells)
    $map = @{}
    for ($i = 0; $i -lt $Cells.Count; $i++) {
        $label = (ConvertTo-Label ([string]$Cells[$i]).Replace('↵', '')).ToUpperInvariant()
        foreach ($k in $tableHeads.Keys) { if (-not $map.ContainsKey($k) -and $tableHeads[$k] -contains $label) { $map[$k] = $i } }
    }
    if ($map.ContainsKey('Content') -and ($map.ContainsKey('Owner') -or $map.ContainsKey('Due'))) { return $map }
    return $null
}

foreach ($doc in $documents) {
    $meeting = Find-MeetingInfo $doc
    $section = ''; $sectionMark = ''
    $tableMap = $null
    $last = $null
    $prevRow = 0
    # セルの中で改行した Excel の行は、読みながらセルごと・改行ごとの行をすぐ後ろに差し込むため、番号で回す
    for ($n = 0; $n -lt $doc.Lines.Count; $n++) {
        $line = $doc.Lines[$n]
        $raw = [string]$line.Text
        $norm = $raw.Normalize($nfkc)
        # 見出し: ■ ◆ ▼ # 【…】・以上の行と、節の見出しと同じ印の行（「□宿題」の節の「□次回予定」）。□ ☐ ◇ ▽ ▶ は、ほかは箇条書きの印
        $isHeading = $norm -match '^\s*([■◆▼#]|【[^】]*】\s*$|以上\s*$)' -or [regex]::IsMatch($norm, $headingPattern, $ignoreCase) -or ($sectionMark -and $norm -match ('^\s*' + [regex]::Escape($sectionMark)))
        # 表の終わり: 表でない行・見出し・Excel で行が飛んだとき
        if (-not $line.Cells -or $isHeading -or ($line.Row -and $line.Row -ne $prevRow + 1)) { $tableMap = $null }
        $prevRow = $line.Row
        # 表（Word の表・Excel の行・Markdown の表）: 担当・期限の見出しのある表は、行をそのまま宿題にする
        if ($line.Cells -and -not $isHeading) {
            $map = Find-TableMap $line.Cells
            if ($map) { $tableMap = $map; continue }
            if ($tableMap) {
                # セルの中の改行（Excel の ↵）は空白にして 1 行にする
                $cell = { param($k) if ($tableMap.ContainsKey($k) -and $tableMap[$k] -lt $line.Cells.Count) { (([string]$line.Cells[$tableMap[$k]]) -replace '\s*↵\s*', ' ').Trim() } else { '' } }
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
                $last = Add-Action $kind $parsed $meeting $line $raw (& $cell 'Status')
                continue
            }
        }
        # 表でない Excel の行のうち、セルの中で改行したもの（Alt+Enter）は、セルごと・改行ごとの行にして続けて読む
        if ($line.Row -and $raw.Contains('↵')) { $doc.Lines.InsertRange($n + 1, (Split-CellLines $line)); continue }
        if (-not $norm.Trim()) { continue }
        # 見出し（宿題・決定事項の節の始まり、ほかの節で終わり）
        $h = [regex]::Match($norm, $headingPattern, $ignoreCase)
        if ($h.Success) { $section = Get-KindOf $h.Groups['k'].Value; $sectionMark = $h.Groups['mark'].Value; $last = $null; continue }
        if ($isHeading) { $section = ''; $sectionMark = ''; $last = $null; continue }
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
        $parsed = Split-ActionText $text $meeting.Date
        # 担当・期限だけの行（件名の無い行）は、前の宿題の担当・期限にする
        if (-not $parsed.Title -and $last -and $last.Kind -eq '宿題' -and ($parsed.Owner -or $parsed.Due -or $parsed.Note)) { Add-OwnerDue $last $parsed; continue }
        $last = Add-Action $section $parsed $meeting $line $raw
    }
}

# 前の会議の宿題の表を写した議事録が複数あるとき、同じ宿題（件名と担当が同じもの）は、会議の日の新しい議事録（同じ日なら後に読んだもの）のものだけを残す。
# 状態・期限はその議事録のもの、起票日は最初に出てきた会議の日（1 つの議事録の中の同じ宿題は、そのまま残す）
$keys = @(foreach ($a in $results) { if ($a.Kind -eq '宿題') { (ConvertTo-Label $a.Title).ToUpperInvariant() + '|' + (Get-OwnerKey $a.Owner) } else { '' } })
$latest = @{}
for ($i = 0; $i -lt $results.Count; $i++) {
    if (-not $keys[$i]) { continue }
    $a = $results[$i]
    $e = $latest[$keys[$i]]
    if (-not $e) { $latest[$keys[$i]] = @{ Path = $a.Place.Path; Date = $a.MeetingDate; Opened = $a.MeetingDate }; continue }
    if ($a.MeetingDate -lt $e.Opened) { $e.Opened = $a.MeetingDate }
    if ($a.MeetingDate -ge $e.Date) { $e.Path = $a.Place.Path; $e.Date = $a.MeetingDate }
}
$kept = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $results.Count; $i++) {
    $a = $results[$i]
    if ($keys[$i]) {
        $e = $latest[$keys[$i]]
        if ($a.Place.Path -ne $e.Path) { continue }
        $a.Opened = $e.Opened
    }
    $kept.Add($a)
}
$results = $kept

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
            [pscustomobject][ordered]@{ 'No' = $i; '区分' = '宿題'; '件名' = $_.Title; '担当' = $_.Owner; '期限' = (Format-DateCell $_.Due); '状態' = $_.Status; '起票日' = (Format-DateCell $_.Opened); '会議' = $_.Meeting; 'メモ' = $_.Note; '出典' = (Format-PlaceText $_.Place) }
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
        # 期限・起票日は日付のセル（Excel で並べ替え・フィルター・計算ができる）
        $row = @{ Values = @([double]$i, '宿題', $a.Title, $a.Owner, $a.Due, $a.Status, $a.Opened, $a.Meeting, $a.Note, (Format-PlaceText $a.Place), $a.Original) }
        if ($a.Place.Sheet) { $row.Link = @{ Column = 9; Path = $a.Place.Path; Location = ("'{0}'!{1}" -f $a.Place.Sheet.Replace("'", "''"), $a.Place.Cell) } }
        $actionRows.Add($row)
    }
    $decisionRows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($d in $decisions) {
        $i++
        $row = @{ Values = @([double]$i, $d.Title, $d.MeetingDate, $d.Meeting, (Format-PlaceText $d.Place), $d.Original) }
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
#       節            「■宿題」「【ToDo】」「アクションアイテム」「決定事項」などの見出しの下の行（次の ■・◆・▼・【】・# の見出しまで。
#                     □・☐・◇・▽・▶ は箇条書きの印。「□宿題」のように節の見出しに使った印なら、その印の行まで）
#       印            行の中の【宿題】【TODO】【AI】【決定】や、行頭の「TODO:」「宿題:」
#       表            内容（件名・項目・アクション）と、担当か期限の見出しがある表（Word の表・Excel）。状況・状態の列は宿題の状態にする
#
#     担当と期限は、行の中の「（山田、10/30まで）」「（山田、来週中）」「担当: 山田、佐藤」「期限: 10/30」「@山田」「山田さんが」「[山田]」などから
#     取り出し、残りを件名にします。「今週中」「来週中」は会議の週・翌週の金曜日（「来週水曜」はその曜日）、「10月末」は月の最後の日、
#     「次回まで」はメモに書きます。宿題の次の行が担当・期限だけなら、前の宿題の担当・期限にします。Excel のセルの中の改行は行の区切りにします。
#     会議の日は「日時・開催日」の行（和暦も）、ファイル名の日付（20261020 など）、ファイルの更新日の順に探します（年の無い期限の年に使います）。
#     同じ宿題（件名と担当が同じもの）が複数の議事録にあれば、会議の日の新しい議事録のものだけを残します（起票日は最初に出てきた会議の日）。
#
#     -OutCsv の CSV（No・区分・件名・担当・期限・状態・起票日など）は、課題管理表に貼り付けるか、
#     New-WeeklyReport.ps1 の -IssueList に渡して、週次報告で期限切れを追えます（状態が「完了」などの宿題は、終わったものとして数えます）。
#
# .PARAMETER Path
#     議事録のファイルまたはフォルダ（.docx・.xlsx・.xlsm・.txt・.md。複数可）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（宿題・決定事項。期限・起票日は日付のセル。Excel の議事録は該当セルへのリンク付き）。
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
