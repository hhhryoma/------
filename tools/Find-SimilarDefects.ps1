[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $BugList,

    [Parameter(Mandatory = $true, Position = 1)]
    [string[]] $DesignPath,

    [string[]] $Id,

    [ValidateRange(1, 1000)]
    [int] $MaxCandidates = 20,

    [string] $OutXlsx,

    [string] $OutPrompt,

    [string] $Answer,

    [string] $Headers,

    [string] $TableHeaders,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
if (-not $TableHeaders) { $TableHeaders = Join-Path $PSScriptRoot 'ddl/テーブル定義書の見出し.csv' }
foreach ($name in @('Headers', 'TableHeaders', 'OutXlsx', 'OutPrompt', 'Answer')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($Answer -and -not (Test-Path -LiteralPath $Answer)) { throw "-Answer のファイルが見つかりません: $Answer" }

$display = @('No', 'Bug', 'Kind', 'Target', 'Common', 'Score')
if ($Answer) { $display += 'Judgement' }
Update-TypeData -TypeName 'SimilarDefects.Candidate' -DefaultDisplayPropertySet $display -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
# 横展開の列がこの値の障害は探さない（横展開が要らないもの・済んだもの）
$skipPattern = '^(不要|否|なし|無し|無|対象外|-|－|ー|―|×|済|済み|完了|実施済|実施済み|対応済|対応済み)$'
# 状態がこの値の障害は、障害ではないので探さない（Measure-DefectRemoval.ps1 と同じ）
$excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|仕様どおり|仕様通り|障害ではない|非障害|バグではない|誤指摘'
# 言葉として比べない、どの設計書にも出てくる言葉
$stopWords = @('場合', '処理', '表示', '画面', '入力', '出力', '項目', '設定', '選択', '実行', '確認', '内容', '対応', '修正', '発生', '不具合', '障害', '原因', '対策', '結果', '機能', '必要', '可能', '不可', '部分', '全体', '指定', '中断', '遷移', '生成', '考慮', '漏れ', '誤り', '不足', '以上', '以下', '未満', '番号', '一覧', 'データ', 'ボタン', 'メッセージ', 'エラー', 'チェック', 'システム', 'テスト', 'ケース', 'マスタ', 'コード')
$wordRx = New-Object System.Text.RegularExpressions.Regex('[ァ-ヴー]{2,}|[一-龠々〆ヵヶ]{2,}|[A-Za-z][A-Za-z0-9_]{2,}')
$refRx = New-Object System.Text.RegularExpressions.Regex('(?<![A-Za-z0-9])[A-Za-z]{1,4}\d{2,6}(?![A-Za-z0-9])')

function Get-Words {
    #     文の言葉（2 文字以上のカタカナ・漢字の続き、3 文字以上の英数字）。どこにでも出てくる言葉と、ID（E0004 など）は除く
    param([string] $Text)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($m in $wordRx.Matches(([string]$Text).Normalize($nfkc))) {
        $w = $m.Value
        if ($w -match '^[A-Za-z]') { $w = $w.ToUpperInvariant() }
        if ($w -match '^[A-Z]+\d+$' -or $stopWords -contains $w -or $list.Contains($w)) { continue }
        $list.Add($w)
    }
    return , $list
}

function Get-Refs {
    #     文の ID（E0004・C001 など。大文字）
    param([string] $Text)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($m in $refRx.Matches(([string]$Text).Normalize($nfkc))) {
        $r = $m.Value.ToUpperInvariant()
        if (-not $list.Contains($r)) { $list.Add($r) }
    }
    return , $list
}

function Get-CheckKinds {
    #     入力チェックの欄の、チェックの種類（「必須(E0001)、日付形式(E0004)」→ 必須・日付形式）
    param([string] $Text)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($part in (([string]$Text).Normalize($nfkc) -split '[、,;/・\s]+')) {
        $k = ($part -replace '\([^)]*\)', '').Trim()
        if ($k -and -not $list.Contains($k)) { $list.Add($k) }
    }
    return , $list
}

function Get-ShortText {
    #     一覧に出す短い文（改行（Search-Excel の ↵）を除き、長ければ切る）
    param([string] $Text, [int] $Length = 120)
    $t = (([string]$Text) -replace '\s*(\r?\n|↵)\s*', ' ').Trim()
    if ($t.Length -gt $Length) { $t = $t.Substring(0, $Length) + '…' }
    return $t
}

# ---------------------------------------------------------------- 設計書を読む


$searchArgs = @{ Pattern = '\S'; Path = $DesignPath; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$books = ConvertTo-DesignBooks $hits
$design = Read-ScreenDesign $books
$tables = Read-TableDesign -Books $books -Candidates (Read-TableHeaderCandidates $TableHeaders)

# 設計の要素（画面の項目・イベント・テーブルのカラム）。Key は名前を比べる形
$elements = New-Object System.Collections.Generic.List[object]
$screenIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$screenNames = @{}   # 画面名（ConvertTo-Label 済み）→ 画面 ID
foreach ($x in @($design.Items.ToArray()) + @($design.Events.ToArray())) {
    if (-not $x.Screen) { continue }
    [void]$screenIds.Add($x.Screen)
    if ($x.ScreenName) { $screenNames[(ConvertTo-Label $x.ScreenName)] = $x.Screen }
}
foreach ($it in $design.Items) {
    if (-not $it.Item) { continue }
    $refs = Get-Refs ('{0} {1}' -f $it.Checks, $it.Note)
    $parts = New-Object System.Collections.Generic.List[string]
    if ($it.Type) { $parts.Add('型 ' + $it.Type) }
    if ($it.DigitsText) { $parts.Add('桁数 ' + $it.DigitsText) }
    if ($it.Required) { $parts.Add('必須') }
    if ($it.Io) { $parts.Add('入出力 ' + $it.Io) }
    if ($it.Checks) { $parts.Add('入力チェック ' + $it.Checks) }
    if ($it.Note) { $parts.Add('備考 ' + (Get-ShortText $it.Note 60)) }
    $elements.Add([pscustomobject]@{
            Kind = '項目'; Owner = $it.Screen; Name = $it.Item; Key = (ConvertTo-Label $it.Item); Label = ('{0} {1}' -f $it.Screen, $it.Item).Trim()
            Type = $it.Type; Digits = $it.Digits; Required = [bool]$it.Required; Input = [bool]$it.Input; Io = [string]$it.Io; Note = ([string]$it.Note).Normalize($nfkc).Trim()
            CheckKinds = (Get-CheckKinds $it.Checks); Messages = @($refs | Where-Object { $design.Messages.ContainsKey($_) }); Codes = @($refs | Where-Object { -not $design.Messages.ContainsKey($_) -and -not $screenIds.Contains($_) })
            Words = (Get-Words $it.Note); Text = ('{0} {1}' -f $it.Checks, $it.Note).Normalize($nfkc); Detail = ($parts -join '・')
            File = $it.File; Path = $it.Path; Sheet = $it.Sheet; Row = $it.Row; Seq = $elements.Count
        })
}
foreach ($ev in $design.Events) {
    if (-not $ev.Name) { continue }
    $refs = Get-Refs ('{0} {1}' -f $ev.Summary, $ev.Messages)
    $messages = @($refs | Where-Object { $design.Messages.ContainsKey($_) })
    $detail = '処理 ' + (Get-ShortText $ev.Summary)
    if ($ev.Next) { $detail += '・遷移先 ' + $ev.Next }
    $elements.Add([pscustomobject]@{
            Kind = 'イベント'; Owner = $ev.Screen; Name = $ev.Name; Key = (ConvertTo-Label $ev.Name); Label = ('{0} {1}（{2}）' -f $ev.Screen, $ev.Name, $ev.Id).Trim()
            Type = ''; Digits = $null; Required = $false; Input = $true; Io = ''; Note = ''
            CheckKinds = @(); Messages = $messages; Codes = @()
            Words = (Get-Words $ev.Summary); Text = ([string]$ev.Summary).Normalize($nfkc); Detail = $detail
            File = $ev.File; Path = $ev.Path; Sheet = $ev.Sheet; Row = $ev.Row; Seq = $elements.Count
        })
}
foreach ($t in $tables) {
    foreach ($c in $t.Columns) {
        $parts = New-Object System.Collections.Generic.List[string]
        if ($c.Logical) { $parts.Add('論理名 ' + $c.Logical) }
        $parts.Add('型 ' + $c.BaseType)
        if ($null -ne $c.Length) { $parts.Add('桁数 ' + $c.Length) }
        if ($c.NotNull) { $parts.Add('NOT NULL') }
        $fk = ''
        if ($c.FkTable) { $fk = ('{0}.{1}' -f $c.FkTable, $c.FkColumn).TrimEnd('.'); $parts.Add('外部キー ' + $fk) }
        if ($c.Note) { $parts.Add('備考 ' + (Get-ShortText $c.Note 60)) }
        $elements.Add([pscustomobject]@{
                Kind = 'カラム'; Owner = $t.Id; Name = $c.Name; Key = ([string]$c.Name).ToUpperInvariant(); Label = ('{0}.{1}（{2}）' -f $t.Id, $c.Name, $c.Logical)
                Type = ([string]$c.BaseType).ToUpperInvariant(); Digits = $c.Length; Required = [bool]$c.NotNull; Input = $true; Io = ''; Note = ([string]$c.Note).Normalize($nfkc).Trim()
                CheckKinds = @(); Messages = @(); Codes = @(); Fk = $fk; Logical = ([string]$c.Logical).Normalize($nfkc).Trim()
                Words = (Get-Words ('{0} {1}' -f $c.Logical, $c.Note)); Text = ('{0} {1}' -f $c.Logical, $c.Note).Normalize($nfkc); Detail = ($parts -join '・')
                File = $c.Place.File; Path = $c.Place.Path; Sheet = $c.Place.Sheet; Row = [int]([regex]::Match($c.Place.Cell, '\d+$').Value); Seq = $elements.Count
            })
    }
}
if ($elements.Count -eq 0) { Write-Warning ('設計書に、画面の項目・イベント・テーブルのカラムが見つかりません: {0}' -f ($DesignPath -join ', ')) }

# 言葉の出てくる要素の数。多くの要素に出てくる言葉と、2 文字の言葉（受注・更新など）は、同じ作りの手がかりにしない
$wordCount = @{}
foreach ($e in $elements) { foreach ($w in $e.Words) { if (-not $wordCount.ContainsKey($w)) { $wordCount[$w] = 0 }; $wordCount[$w]++ } }
$maxWordCount = [Math]::Max(3, [int][Math]::Ceiling($elements.Count * 0.1))
foreach ($e in $elements) {
    $e | Add-Member -NotePropertyName AllWords -NotePropertyValue (Get-Words $e.Text)
    $e.Words = @($e.Words | Where-Object { $_.Length -ge 3 -and $wordCount[$_] -le $maxWordCount })
}

# ---------------------------------------------------------------- 障害管理表を読む


$bugFields = @('障害ID', '機能', '障害内容', '原因', '原因詳細', '対策', '横展開', '状態')
$readLog = New-Object System.Collections.Generic.List[object]
$records = Read-HeaderRecords -Inputs (Get-InputFiles $BugList '障害管理表' -NoRecurse:$NoRecurse) -Candidates (Read-HeaderCandidates $Headers $bugFields) -Fields $bugFields `
    -Need @('障害ID', @('障害内容', '原因詳細')) -Key @('障害ID') -What '障害管理表' -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
$bugs = New-Object System.Collections.Generic.List[object]
$allCount = 0; $skipped = 0
foreach ($rec in $records) {
    $v = $rec.Values
    $bugId = ([string]$v['障害ID']).Normalize($nfkc).Trim()
    if (-not $bugId) { continue }
    $allCount++
    if ($Id) {
        if (@($Id | Where-Object { $bugId -like $_ }).Count -eq 0) { continue }
    }
    elseif (([string]$v['横展開']).Normalize($nfkc).Trim() -match $skipPattern -or ([string]$v['状態']) -match $excludedPattern) { $skipped++; continue }
    $cause = (@([string]$v['原因'], [string]$v['原因詳細']) | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' '
    $bugs.Add([pscustomobject]@{
            Id = $bugId; Function = ([string]$v['機能']).Trim(); Content = ([string]$v['障害内容']).Trim(); Cause = $cause; Measure = ([string]$v['対策']).Trim()
            Record = $rec; Screen = $null; Anchors = (New-Object System.Collections.Generic.List[object]); Candidates = @()
        })
}
if ($records.Count -eq 0) { Write-Warning '障害管理表が見つかりません（障害番号と、障害内容か原因詳細の見出しがある表を探します）' }
elseif ($Id -and $bugs.Count -eq 0) { Write-Warning ('障害番号 {0} の障害が見つかりません' -f ($Id -join ', ')) }

# ---------------------------------------------------------------- 障害の箇所を探す


function Sort-ByLength {
    #     長い順に並べる（同じ長さなら元の順）。Sort-Object は同じ値の並びが版で変わることがあるので使わない
    param($Items, [scriptblock] $Text)
    $list = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($x in $Items) { $list.Add([pscustomobject]@{ Item = $x; Length = ([string](& $Text $x)).Length; Index = $i }); $i++ }
    $list.Sort([System.Comparison[object]] { param($a, $b) if ($a.Length -ne $b.Length) { return $b.Length.CompareTo($a.Length) }; return $a.Index.CompareTo($b.Index) })
    return @($list | ForEach-Object { $_.Item })
}

function Find-BugScreen {
    #     障害の画面（機能の欄、無ければ障害内容などの文に、最初に出てくる画面 ID・画面名）
    param($Bug)
    foreach ($text in @($Bug.Function, ('{0} {1} {2}' -f $Bug.Content, $Bug.Cause, $Bug.Measure))) {
        $label = (ConvertTo-Label $text).ToUpperInvariant()
        if (-not $label) { continue }
        $first = $null; $at = [int]::MaxValue
        foreach ($sid in $screenIds) {
            $m = [regex]::Match($label, ('(?<![A-Z0-9]){0}(?![0-9])' -f [regex]::Escape($sid.ToUpperInvariant())))
            if ($m.Success -and $m.Index -lt $at) { $first = $sid; $at = $m.Index }
        }
        if ($first) { return $first }
        foreach ($name in (Sort-ByLength @($screenNames.Keys) { param($x) $x })) {
            $i = $label.IndexOf($name.ToUpperInvariant())
            if ($i -ge 0 -and $i -lt $at) { $first = $screenNames[$name]; $at = $i }
        }
        if ($first) { return $first }
    }
    return $null
}

function Find-Anchors {
    #     障害の箇所（障害内容・原因・対策に名前の出てくる項目・イベント・カラムと、それらに無いメッセージ ID）
    param($Bug)
    $anchors = New-Object System.Collections.Generic.List[object]
    $text = ConvertTo-Label ('{0} {1} {2}' -f $Bug.Content, $Bug.Cause, $Bug.Measure)
    # 画面名・画面 ID は項目名として探さない（「受注一覧画面」の「受注一覧」は、一覧の項目ではない）
    foreach ($name in (Sort-ByLength @($screenNames.Keys) { param($x) $x })) { $text = $text.Replace($name, '|') }
    foreach ($sid in $screenIds) { $text = [regex]::Replace($text, ('(?<![A-Za-z0-9]){0}(?![0-9])' -f [regex]::Escape($sid)), '|', 'IgnoreCase') }
    # 項目・イベントの名前（長い名前から。見つけた名前は文から除き、中の短い名前を見つけないようにする）
    $pool = @($elements | Where-Object { $_.Kind -ne 'カラム' -and $_.Key.Length -ge 2 -and (-not $Bug.Screen -or $_.Owner -eq $Bug.Screen) })
    foreach ($group in (Sort-ByLength @($pool | Group-InOrder Key) { param($x) $x.Name })) {
        if (-not $text.Contains($group.Name)) { continue }
        $text = $text.Replace($group.Name, '|')
        $events = @($group.Group | Where-Object Kind -eq 'イベント')
        if ($events.Count -gt 0) {
            # 画面が分からないときは、1 つの画面にしかない 4 文字以上のイベント名だけ（「登録でエラー」の「登録」は、ふつうの言葉のことが多い）
            if ($Bug.Screen -or ($group.Name.Length -ge 4 -and @($events | ForEach-Object { $_.Owner } | Select-Object -Unique).Count -eq 1)) { $anchors.Add($events[0]) }
            continue
        }
        $items = @($group.Group | Where-Object Input)
        if ($items.Count -eq 0) { $items = @($group.Group) }
        $anchors.Add($items[0])
    }
    # カラムの物理名（T_ORDER.ORDER_DATE・ORDER_DATE）
    $upper = (ConvertTo-Label ('{0} {1} {2}' -f $Bug.Content, $Bug.Cause, $Bug.Measure)).ToUpperInvariant()
    foreach ($group in @($elements | Where-Object { $_.Kind -eq 'カラム' -and $_.Key.Length -ge 3 } | Group-InOrder Key)) {
        if ($upper -notmatch ('(?<![A-Z0-9_]){0}(?![A-Z0-9_])' -f [regex]::Escape($group.Name))) { continue }
        $columns = @($group.Group | Where-Object { $upper -match ('(?<![A-Z0-9_]){0}\.{1}(?![A-Z0-9_])' -f [regex]::Escape($_.Owner.ToUpperInvariant()), [regex]::Escape($group.Name)) })
        if ($columns.Count -eq 0) { $columns = @($group.Group) }
        $anchors.Add($columns[0])
    }
    # メッセージ ID（障害の箇所の入力チェック・処理に無いもの。「E0006 が出ない」など）
    $used = @($anchors | ForEach-Object { $_.Messages })
    foreach ($m in (Get-Refs $upper)) {
        if (-not $design.Messages.ContainsKey($m) -or $used -contains $m) { continue }
        $msg = $design.Messages[$m]
        $anchors.Add([pscustomobject]@{ Kind = 'メッセージ'; Owner = ''; Name = $m; Key = $m; Label = ('{0}「{1}」' -f $m, (Get-ShortText $msg.Text 40)); Detail = ('メッセージ {0}「{1}」' -f $m, $msg.Text); Seq = -1; Messages = @($m) })
    }
    return , $anchors
}

# ---------------------------------------------------------------- 同じ作りの箇所を探す


function Measure-Similarity {
    #     障害の箇所（Anchor）と、ほかの要素（Other）の似ている点。似ていなければ $null。戻り値: @{ Score; Reasons }
    param($Anchor, $Other, [string[]] $CauseWords)
    $acc = @{ Score = 0 }
    $reasons = New-Object System.Collections.Generic.List[string]
    $add = { param([int] $Points, [string] $Reason) $acc.Score += $Points; $reasons.Add($Reason) }
    $sameName = $Anchor.Key -eq $Other.Key
    switch ($Anchor.Kind) {
        '項目' {
            # 入力の不具合は入力のある項目、表示の不具合は表示する項目
            if ($Anchor.Input -and -not $Other.Input) { return $null }
            if (-not $Anchor.Input -and $Other.Io -and $Other.Io -notmatch 'O') { return $null }
            $typed = $Anchor.Type -and $Anchor.Type -notmatch '^[-－ー―]$'
            $sameType = $typed -and $Anchor.Type -eq $Other.Type
            $checks = @($Anchor.CheckKinds | Where-Object { $Other.CheckKinds -contains $_ } | Select-Object -First 2)
            $codes = @($Anchor.Codes | Where-Object { $Other.Codes -contains $_ } | Select-Object -First 1)
            # 同じ作りとみなすのは、名前か型が同じもの（型の無い項目は、入力チェックかコードが同じもの）
            if (-not ($sameName -or $sameType -or (-not $typed -and ($checks.Count + $codes.Count) -gt 0))) { return $null }
            if ($sameName) { & $add 2 '同じ名前' }
            if ($sameType) { & $add 2 ('型 ' + $Anchor.Type) }
            foreach ($k in $checks) { & $add 2 ('入力チェック ' + $k) }
            foreach ($c in $codes) { & $add 2 ('コード ' + $c) }
            if ($null -ne $Anchor.Digits -and $Anchor.Digits -eq $Other.Digits) { & $add 1 ('桁数 ' + $Anchor.Digits) }
            if ($Anchor.Required -and $Other.Required) { & $add 1 '必須' }
            if ($Anchor.Note -and $Anchor.Note -eq $Other.Note) { & $add 1 ('備考が同じ（{0}）' -f (Get-ShortText $Anchor.Note 30)) }
            else { foreach ($w in @($Anchor.Words | Where-Object { $Other.Words -contains $_ -and $CauseWords -notcontains $_ } | Select-Object -First 2)) { & $add 1 ('備考 ' + $w) } }
        }
        'イベント' {
            if ($sameName) { & $add 3 '同じイベント名' }
            elseif ($Anchor.Key.Contains($Other.Key) -or $Other.Key.Contains($Anchor.Key)) { & $add 1 '似たイベント名' }
            foreach ($m in @($Anchor.Messages | Where-Object { $Other.Messages -contains $_ } | Select-Object -First 2)) { & $add 2 ('メッセージ ' + $m) }
            foreach ($w in @($Anchor.Words | Where-Object { $Other.Words -contains $_ -and $CauseWords -notcontains $_ } | Select-Object -First 3)) { & $add 1 ('処理 ' + $w) }
        }
        'カラム' {
            $sameType = $Anchor.Type -and $Anchor.Type -eq $Other.Type
            if (-not ($sameName -or $sameType)) { return $null }
            if ($sameName) { & $add 2 '同じ名前' }
            if ($sameType) { & $add 2 ('型 ' + $Anchor.Type) }
            if ($null -ne $Anchor.Digits -and $Anchor.Digits -eq $Other.Digits) { & $add 1 ('桁数 ' + $Anchor.Digits) }
            if ($Anchor.Required -and $Other.Required) { & $add 1 'NOT NULL' }
            if ($Anchor.Fk -and $Anchor.Fk -eq $Other.Fk) { & $add 2 ('外部キー ' + $Anchor.Fk) }
            $suffix = [regex]::Match($Anchor.Logical, '(日時|日付|年月日|日|金額|単価|数量|件数|コード|区分|フラグ|名|番号)$').Value
            if ($suffix -and $Other.Logical.EndsWith($suffix)) { & $add 1 ('論理名が「{0}」で終わる' -f $suffix) }
            foreach ($w in @($Anchor.Words | Where-Object { $Other.Words -contains $_ -and $CauseWords -notcontains $_ } | Select-Object -First 2)) { & $add 1 ('備考 ' + $w) }
        }
        'メッセージ' {
            if ($Other.Messages -notcontains $Anchor.Name) { return $null }
            & $add 3 ('メッセージ ' + $Anchor.Name)
        }
    }
    # 原因・対策に書いた言葉（バージョン・排他など）が、ほかの要素の入力チェック・備考・処理にあれば強い手がかり。
    # 2 文字の言葉は言葉として同じものだけ（「条件」を「検索条件」の中に見つけない）。上の共通点に出ている言葉は数えない
    $matched = @($CauseWords | Where-Object {
            $w = $_
            ($Other.AllWords -contains $w -or ($w.Length -ge 3 -and $Other.Text.Contains($w))) -and @($reasons | Where-Object { $_.Contains($w) }).Count -eq 0
        } | Select-Object -First 2)
    foreach ($w in $matched) { & $add 2 ('原因の言葉 ' + $w) }
    if ($acc.Score -lt 3) { return $null }
    return @{ Score = $acc.Score; Reasons = $reasons }
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($bug in $bugs) {
    $bug.Screen = Find-BugScreen $bug
    $bug.Anchors = Find-Anchors $bug
    if ($bug.Anchors.Count -eq 0) { continue }
    # 原因・対策の言葉（障害の箇所の名前と、多くの要素に出てくる言葉は除く）
    $anchorNames = @($bug.Anchors | ForEach-Object { $_.Key; ConvertTo-Label $_.Name })
    $causeWords = @((Get-Words ('{0} {1}' -f $bug.Cause, $bug.Measure)) | Where-Object {
            $w = $_
            $n = @($elements | Where-Object { $_.AllWords -contains $w -or ($w.Length -ge 3 -and $_.Text.Contains($w)) }).Count
            $n -ge 1 -and $n -le $maxWordCount -and $anchorNames -notcontains $w
        })
    $anchorSeqs = @($bug.Anchors | ForEach-Object { $_.Seq })
    $best = @{}   # 要素の Seq → @{ Element; Score; Reasons; Anchor }
    foreach ($a in $bug.Anchors) {
        $pool = @($elements | Where-Object { $anchorSeqs -notcontains $_.Seq })
        if ($a.Kind -eq 'メッセージ') { $pool = @($pool | Where-Object { $_.Kind -ne 'カラム' }) }
        else { $pool = @($pool | Where-Object { $_.Kind -eq $a.Kind }) }
        foreach ($e in $pool) {
            $sim = Measure-Similarity $a $e $causeWords
            if (-not $sim) { continue }
            if (-not $best.ContainsKey($e.Seq) -or $best[$e.Seq].Score -lt $sim.Score) { $best[$e.Seq] = @{ Element = $e; Score = $sim.Score; Reasons = $sim.Reasons; Anchor = $a } }
        }
    }
    # 点数の高い順（同じなら設計書の順）
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($k in $best.Keys) { $list.Add($best[$k]) }
    $list.Sort([System.Comparison[object]] { param($x, $y) if ($x.Score -ne $y.Score) { return $y.Score.CompareTo($x.Score) }; return $x.Element.Seq.CompareTo($y.Element.Seq) })
    $bug.Candidates = @($list.ToArray() | Select-Object -First $MaxCandidates)
    foreach ($c in $bug.Candidates) {
        $e = $c.Element
        $r = [pscustomobject]@{
            No = $results.Count + 1; Bug = $bug.Id; BugText = $bug.Content; Anchor = $c.Anchor.Label; AnchorDetail = $c.Anchor.Detail; Kind = $e.Kind; Target = $e.Label; Detail = $e.Detail
            Common = ($c.Reasons -join '・'); Score = $c.Score; Judgement = ''; Reason = ''
            File = $e.File; Location = ('{0} {1} 行目' -f $e.Sheet, $e.Row); Path = $e.Path; Sheet = $e.Sheet; Row = $e.Row
        }
        $r.PSObject.TypeNames.Insert(0, 'SimilarDefects.Candidate')
        $results.Add($r)
    }
}

# ---------------------------------------------------------------- Copilot に絞り込みを頼む


function ConvertTo-Cell {
    #     表のセルに入れる文字（改行と | を除く）
    param([string] $Text)
    return ($Text -replace '\r?\n', ' ' -replace '\|', '／').Trim()
}

function New-JudgementPrompt {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('あなたはシステム開発の品質管理の担当者です。障害の横展開（同じ原因の不具合が、ほかの箇所にもないか）で確かめる箇所を絞り込んでください。')
    $lines.Add('候補は、障害の箇所と同じ作り（型・入力チェック・処理など）の箇所を、設計書から機械的に集めたものです。')
    foreach ($bug in $bugs) {
        $rows = @($results | Where-Object { $_.Bug -eq $bug.Id })
        if ($rows.Count -eq 0) { continue }
        $lines.Add('')
        $lines.Add(('# 障害 {0}' -f $bug.Id))
        if ($bug.Function) { $lines.Add(('- 機能: {0}' -f (ConvertTo-Cell $bug.Function))) }
        $lines.Add(('- 障害内容: {0}' -f (ConvertTo-Cell $bug.Content)))
        if ($bug.Cause) { $lines.Add(('- 原因: {0}' -f (ConvertTo-Cell $bug.Cause))) }
        if ($bug.Measure) { $lines.Add(('- 対策: {0}' -f (ConvertTo-Cell $bug.Measure))) }
        foreach ($a in $bug.Anchors) { $lines.Add(('- 障害の箇所: {0}（{1}）' -f $a.Label, (ConvertTo-Cell $a.Detail))) }
        $lines.Add('')
        $lines.Add('| No | 候補の箇所 | 設計の内容 | 共通点 |')
        $lines.Add('| --- | --- | --- | --- |')
        foreach ($r in $rows) { $lines.Add(('| {0} | {1} | {2} | {3} |' -f $r.No, (ConvertTo-Cell $r.Target), (ConvertTo-Cell $r.Detail), (ConvertTo-Cell $r.Common))) }
    }
    $lines.Add('')
    $lines.Add('# 回答の形式')
    $lines.Add('候補の No ごとに 1 行ずつ、次の表だけを出力してください（表の前後に文を書かないでください）。')
    $lines.Add('| No | 判定 | 理由 |')
    $lines.Add('| --- | --- | --- |')
    $lines.Add('- 判定は「要確認」（同じ原因の不具合がありそうで、確かめるべき）か「不要」のどちらかにしてください。')
    $lines.Add('- 理由は、障害の原因と候補の設計の内容から、なぜそう判断したかを 100 文字以内で書いてください。設計書から分からないことは「～なら要確認」と書いてください。')
    return ($lines -join "`r`n")
}

function ConvertFrom-JudgementAnswer {
    #     回答の表（Markdown の | 区切り、または Copilot の表をコピーしたタブ区切り）から、No → 判定・理由
    param([string] $Text)
    $result = @{}
    foreach ($line in ($Text -split '\r?\n')) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t.Contains("`t")) { $cells = @($t -split "`t") }
        elseif ($t.StartsWith('|')) { $cells = @($t.Trim('|') -split '\|') }
        else { continue }
        $cells = @($cells | ForEach-Object { $_.Trim() })
        if ($cells.Count -lt 2) { continue }
        $no = $cells[0].Normalize($nfkc) -replace '^No\.?\s*', '' -replace '[.．。]$', ''
        if ($no -notmatch '^\d+$') { continue }
        $judgement = $cells[1].Normalize($nfkc).Trim()
        if ($judgement -match '不要|対象外|なし|無し') { $judgement = '不要' }
        elseif ($judgement -match '要確認|確認|要') { $judgement = '要確認' }
        $reason = ''
        if ($cells.Count -ge 3) { $reason = (($cells[2..($cells.Count - 1)] -join ' ') -replace '<br\s*/?>', ' ').Trim() }
        $result[[int]$no] = [pscustomobject]@{ Judgement = $judgement; Reason = $reason }
    }
    return $result
}

$answered = 0
if ($Answer) {
    $answers = ConvertFrom-JudgementAnswer ([System.IO.File]::ReadAllText($Answer))
    if ($answers.Count -eq 0) { Write-Warning "回答から「No | 判定 | 理由」の行を読み取れませんでした: $Answer" }
    foreach ($r in $results) {
        if (-not $answers.ContainsKey($r.No)) { continue }
        $r.Judgement = $answers[$r.No].Judgement
        $r.Reason = $answers[$r.No].Reason
        $answered++
    }
}

# ---------------------------------------------------------------- 結果


$results

if ($OutPrompt -and $results.Count -gt 0) {
    $dir = Split-Path -Parent $OutPrompt
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    [System.IO.File]::WriteAllText($OutPrompt, (New-JudgementPrompt), $utf8Bom)
}

if ($OutXlsx) {
    $candidateRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $row = @{ Values = @([double]$r.No, $r.Bug, $r.Anchor, $r.Kind, $r.Target, $r.Detail, $r.Common, [double]$r.Score, $r.Judgement, $r.Reason, '', ('{0} {1}' -f $r.File, $r.Location)) }
        if ($r.Path) { $row.Link = @{ Column = 5; Path = $r.Path; Location = ("'{0}'!A{1}" -f $r.Sheet.Replace("'", "''"), $r.Row) } }
        $candidateRows.Add($row)
    }
    $bugRows = New-Object System.Collections.Generic.List[object]
    foreach ($bug in $bugs) {
        $note = ''
        if ($bug.Anchors.Count -eq 0) { $note = '障害の箇所（画面の項目・イベント・カラム・メッセージ ID）が、障害内容・原因・対策に見つかりません' }
        elseif ($bug.Candidates.Count -eq 0) { $note = '同じ作りの箇所が見つかりません' }
        $row = @{ Values = @($bug.Id, $bug.Function, $bug.Content, $bug.Cause, (@($bug.Anchors | ForEach-Object { $_.Label }) -join '・'), [double]$bug.Candidates.Count, $note) }
        $rec = $bug.Record
        if ($rec.Sheet) { $row.Link = @{ Column = 1; Path = $rec.Path; Location = ("'{0}'!A{1}" -f $rec.Sheet.Replace("'", "''"), $rec.Row) } }
        else { $row.Link = @{ Column = 1; Path = $rec.Path } }
        $bugRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '横展開の候補'; Columns = @(@{ Header = 'No'; Width = 5 }, @{ Header = '障害ID'; Width = 10 }, @{ Header = '障害の箇所'; Width = 24 }, @{ Header = '種類'; Width = 8 }, @{ Header = '候補の箇所'; Width = 28 }, @{ Header = '設計の内容'; Width = 60 }, @{ Header = '共通点'; Width = 40 }, @{ Header = '点数'; Width = 6 }, @{ Header = 'AI の判定'; Width = 9 }, @{ Header = 'AI の理由'; Width = 40 }, @{ Header = '確認結果'; Width = 12 }, @{ Header = '設計書'; Width = 40 }); Rows = $candidateRows },
            @{ Name = '障害'; Columns = @(@{ Header = '障害ID'; Width = 10 }, @{ Header = '機能'; Width = 14 }, @{ Header = '障害内容'; Width = 50 }, @{ Header = '原因'; Width = 40 }, @{ Header = '障害の箇所'; Width = 30 }, @{ Header = '候補'; Width = 6 }, @{ Header = 'メモ'; Width = 50 }); Rows = $bugRows },
            (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-ReadLogSummary $readLog
$withCandidates = @($bugs | Where-Object { $_.Candidates.Count -gt 0 }).Count
$noAnchor = @($bugs | Where-Object { $_.Anchors.Count -eq 0 })
$line = '障害 {0} 件のうち、横展開を探した障害 {1} 件' -f $allCount, $bugs.Count
if ($skipped -gt 0) { $line += '（横展開が不要・済みのもの、障害ではないもの {0} 件を除く）' -f $skipped }
Write-Host $line
$kinds = @($results | Group-InOrder Kind | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
$line = '  候補あり {0} 件（候補 {1} か所' -f $withCandidates, $results.Count
if ($kinds.Count -gt 0) { $line += '。' + ($kinds -join '・') }
Write-Host ($line + '）')
if ($noAnchor.Count -gt 0) { Write-Host ('  障害の箇所が設計書に見つからない障害 {0} 件: {1}（障害内容・原因に、画面の項目名・イベント名・カラム名・メッセージ ID を書いてください）' -f $noAnchor.Count, (@($noAnchor | ForEach-Object { $_.Id }) -join '・')) }
if ($Answer) {
    $need = @($results | Where-Object Judgement -eq '要確認').Count
    $notNeed = @($results | Where-Object Judgement -eq '不要').Count
    Write-Host ('AI の判定: 要確認 {0}・不要 {1}（回答の無い候補 {2}）' -f $need, $notNeed, ($results.Count - $answered))
}
if ($OutPrompt -and $results.Count -gt 0) {
    Write-Host "Copilot への依頼を保存しました: $OutPrompt"
    Write-Host '  依頼を Copilot に貼り付けて送り、回答を UTF-8 のテキストに保存して -Answer に渡すと、AI の判定を一覧に入れます（障害管理表と設計書は変えずに実行してください）'
}
elseif ($OutPrompt) { Write-Host '候補が無いので、Copilot への依頼は作りませんでした' }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     障害の箇所と同じ作り（型・入力チェック・処理・カラムの型など）の箇所を設計書から集め、横展開の候補一覧を作ります（Excel 不要）。
#
# .DESCRIPTION
#     障害管理表の障害内容・原因・対策に名前の出てくる、画面の項目・イベント・テーブルのカラム・メッセージ ID を「障害の箇所」とし、
#     設計書から同じ作りの箇所を探して、似ている点（共通点）と点数を付けて一覧にします。
#
#       項目: 名前か型が同じで、入力チェック・コード・桁数・必須・備考が同じものほど高い点（入力の項目には入力のある項目だけ）
#       イベント: 同じイベント名・同じメッセージ・処理に同じ言葉があるもの
#       カラム: 名前か型が同じで、桁数・NOT NULL・外部キー・論理名の終わり（金額・日など）が同じもの
#       メッセージ: そのメッセージを使う項目・イベント
#
#     原因・対策に書いた言葉（うるう年・バージョンなど）がある箇所は、点数を高くします。
#     -OutPrompt で、候補を Copilot に絞り込ませる依頼を作り、回答を -Answer に渡すと AI の判定と理由を一覧に入れます。
#     横展開の列が「不要」「済」などの障害と、状態が「取下げ」「対象外」などの障害は探しません（-Id で指定したときは探します）。
#
# .PARAMETER BugList
#     障害管理表のファイルまたはフォルダ（Excel・CSV。複数可）。障害番号と、障害内容か原因詳細の見出しがある表を読みます。
#
# .PARAMETER DesignPath
#     設計書のフォルダまたはファイル（画面設計書・テーブル定義書・メッセージ一覧。複数可）。
#
# .PARAMETER Id
#     横展開を探す障害の障害番号（複数可。ワイルドカード可）。省略すると、横展開が不要・済みのもの以外のすべての障害。
#
# .PARAMETER MaxCandidates
#     障害ごとの候補の数の上限（点数の高い順。既定は 20）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（横展開の候補・障害・読み取り結果のシート）。
#
# .PARAMETER OutPrompt
#     候補を Copilot に絞り込ませる依頼のテキストを保存します。
#
# .PARAMETER Answer
#     Copilot の回答（「No | 判定 | 理由」の表）を保存したテキスト。AI の判定と理由を一覧に入れます。
#
# .PARAMETER Headers
#     障害管理表の見出しの候補の CSV。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER TableHeaders
#     テーブル定義書の見出しの候補の CSV。既定は tools\ddl\テーブル定義書の見出し.csv（Export-TableDdl.ps1 と同じ）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\Find-SimilarDefects.ps1 D:\テスト\障害管理表.xlsx D:\設計書 -Id B012 -OutXlsx D:\横展開.xlsx -OutPrompt D:\横展開の依頼.txt
#
# .EXAMPLE
#     .\Find-SimilarDefects.ps1 D:\テスト\障害管理表.xlsx D:\設計書 -Id B012 -Answer D:\横展開の回答.txt -OutXlsx D:\横展開.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     候補は設計書の書き方が同じものを機械的に集めたものです。同じ不具合があるかは、ソースやテストで確かめてください。
