[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $DesignPath,

    [ValidateRange(0.0, 1.0)]
    [double] $MinErrorRatio = 0.2,

    [string[]] $VagueWords = @('正しく', '正常に', '適切に', '問題なく', '問題ないこと', '期待通り', '期待どおり', '想定通り', '想定どおり', '仕様通り', '仕様どおり', 'うまく', 'きちんと'),

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
foreach ($name in @('Headers', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'TestSpec.Issue' -DefaultDisplayPropertySet File, Location, Level, Item, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# 異常系とみなす区分、区分の列が無いときに期待結果から異常系と推定する言葉、境界値のテストとみなす言葉
$errorKindPattern = '異常|エラー|準正常|例外|NG系|不正'
$errorTextPattern = 'エラー|[EW]\d{4}|不正|できない|されない|超え|超過|未入力|異常|警告|拒否|失敗'
$boundaryPattern = '桁|文字|最大|最小|上限|下限|境界|超え|超過|オーバー|ちょうど|未満|以上|以下'
# 必須項目の未入力のケース、型（日付・数値・半角）の不正な値のケースとみなす言葉
$requiredPattern = '未入力|入力なし|入力しない|空欄|空白|空の|空で|ブランク|null|必須'
$typePatterns = [ordered]@{
    '日付' = '形式|不正|存在しない|ありえない|日付でない|日付以外|文字|[0-9０-９]+月[0-9０-９]+日|2/30|2/31|13月|閏|うるう'
    '数値' = '数値以外|数字以外|文字|英字|記号|マイナス|負|小数|不正|範囲|0を|ゼロ|全角'
    '半角' = '全角|記号|不正|半角以外|カナ|漢字'
}
# 結果の欄に書かれる、未実施を表す言葉（これ以外の読めない値を知らせる）
$notRunPattern = '^(未実施|未|未着手|実施前|予定|-)$'

$fields = @('機能', 'テストID', 'テスト内容', '期待結果', '手順', '前提条件', '入力値', '区分', '結果', 'バグ番号', '予定日', '実施日', '工程')
$candidates = Read-HeaderCandidates $Headers $fields
$vagueRx = $null
if ($VagueWords) { $vagueRx = New-Object System.Text.RegularExpressions.Regex((@($VagueWords | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.Normalize($nfkc)) }) -join '|')) }

# ---------------------------------------------------------------- テスト仕様書を読む


# 読み取り結果（どのファイル・シートをどの見出しで読んだか）
$readLog = New-Object System.Collections.Generic.List[object]
$records = Read-HeaderRecords -Inputs (Get-InputFiles $Path 'テスト仕様書' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
    -Need @(@('期待結果', 'テスト内容'), @('テストID', '手順', '結果')) -Key @('テストID', '期待結果', 'テスト内容', '手順') -What 'テスト仕様書' `
    -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
$allCases = ConvertTo-TestCases $records
$cases = @($allCases | Where-Object { $_.Kind -ne '対象外' })
if ($cases.Count -eq 0) { Write-Warning 'テストケースが見つかりません（期待結果か確認内容と、テストID・手順・結果のどれかの見出しがある表を探します）' }

$issues = New-Object System.Collections.Generic.List[object]
function Add-Issue {
    param([string] $Level, [string] $Item, $Case, [string] $Field, [string] $Detail)
    $p = $Case.Place
    # 指摘した項目の列を指す（無ければ ID などの列）
    $col = $p.Column
    if ($Field -and $p.Columns[$Field]) { $col = $p.Columns[$Field] }
    $location = '{0} 行目' -f $p.Row
    if ($p.Sheet) { $location = '{0}!{1}{2}' -f $p.Sheet, (ConvertTo-ColumnName $col), $p.Row }
    $i = [pscustomobject]@{ File = $p.File; Location = $location; Level = $Level; Item = $Item; Function = $Case.Function; Id = $Case.Id; Detail = $Detail; Path = $p.Path; Sheet = $p.Sheet; Row = $p.Row; Col = $col; Seq = $issues.Count; Cell = $(if ($p.Sheet) { '{0}{1}' -f (ConvertTo-ColumnName $col), $p.Row } else { '' }) }
    $i.PSObject.TypeNames.Insert(0, 'TestSpec.Issue')
    $issues.Add($i)
}

# ---------------------------------------------------------------- ケースごとのチェック


function Get-Text {
    param($Case, [string] $Field)
    return ([string]$Case.Values[$Field]).Normalize($nfkc).Trim()
}

$seenIds = @{}
$seenContents = @{}
foreach ($c in $cases) {
    $labels = $c.Place.Labels
    $expected = Get-Text $c '期待結果'
    if ($labels['期待結果'] -and -not $expected) { Add-Issue 'エラー' '期待結果が空' $c '期待結果' '期待結果が書かれていません' }
    elseif ($expected -and $vagueRx) {
        # 「」で書いた画面の文言（メッセージなど）の中の「正しく」などは、あいまいな書き方ではない
        $m = $vagueRx.Match(($expected -replace '「[^「」]*」', ''))
        if ($m.Success) { Add-Issue '警告' 'あいまいな期待結果' $c '期待結果' ("「{0}」では、何を確かめれば合格かが分かりません（表示される値・メッセージなどを具体的に書いてください）: {1}" -f $m.Value, $expected) }
    }
    if ($labels['手順'] -and -not (Get-Text $c '手順')) { Add-Issue '警告' '手順が空' $c '手順' '手順（操作）が書かれていません' }

    # テスト ID の重複（同じ表の中）
    $rawId = Get-Text $c 'テストID'
    if ($rawId) {
        $idKey = '{0}|{1}|{2}' -f $c.Place.Path, $c.Place.Sheet, $rawId.ToUpperInvariant()
        if ($seenIds.ContainsKey($idKey)) { Add-Issue 'エラー' 'テストIDの重複' $c 'テストID' ("テストID {0} が {1} 行目と重複しています" -f $rawId, $seenIds[$idKey].Place.Row) }
        else { $seenIds[$idKey] = $c }
    }

    # 同じ内容のケース（前提条件・手順・入力値・期待結果がすべて同じ。機能ごと）
    $parts = @('前提条件', '手順', '入力値', '期待結果', 'テスト内容' | ForEach-Object { (ConvertTo-Label (Get-Text $c $_)) })
    if (($parts -join '').Length -gt 0) {
        $contentKey = '{0}|{1}' -f $c.Key, ($parts -join '|')
        if ($seenContents.ContainsKey($contentKey)) {
            $first = $seenContents[$contentKey]
            Add-Issue '警告' '同じ内容のケース' $c 'テストID' ("{0}（{1} 行目）と、前提条件・手順・入力値・期待結果が同じです" -f $first.Id, $first.Place.Row)
        }
        else { $seenContents[$contentKey] = $c }
    }

    # 結果の欄の読めない値
    $result = Get-Text $c '結果'
    if ($result -and $c.Kind -eq '未実施' -and $result -notmatch $notRunPattern) { Add-Issue '警告' '結果の値が読めない' $c '結果' ("結果「{0}」を OK・NG・保留・対象外・未実施のどれとも読めません" -f $result) }
}

# ---------------------------------------------------------------- 機能ごとのチェック（異常系の割合）


$functions = New-Object System.Collections.Generic.List[object]
foreach ($g in @($cases | Group-Object Key)) {
    $group = @($g.Group)
    $first = $group[0]
    $hasKind = [bool]$first.Place.Labels['区分']
    $errors = @($group | Where-Object {
            if ($hasKind) { (Get-Text $_ '区分') -match $errorKindPattern }
            else { ('{0} {1} {2}' -f (Get-Text $_ '期待結果'), (Get-Text $_ 'テスト内容'), (Get-Text $_ '入力値')) -match $errorTextPattern }
        }).Count
    $ratio = [math]::Round($errors / $group.Count, 3)
    $basis = '区分の列'
    if (-not $hasKind) { $basis = '期待結果の言葉から推定' }
    $functions.Add([pscustomobject]@{ Function = $first.Function; Key = $g.Name; Cases = $group.Count; ErrorCases = $errors; ErrorRatio = $ratio; Basis = $basis; First = $first; Boundary = '' })
    if ($group.Count -ge 5 -and $ratio -lt $MinErrorRatio) {
        Add-Issue '警告' '異常系が少ない' $first '' ("異常系のケースが {0} 件 / {1} 件（{2:0}%）で、{3:0}% 未満です（{4}）。入力チェック・エラー時の動きのケースを確かめてください" -f $errors, $group.Count, ($ratio * 100), ($MinErrorRatio * 100), $basis)
    }
}

function Get-NameVariants {
    #     項目名（ConvertTo-Label・大文字にしたもの）の、よくある別の書き方（語尾の 番号↔No・コード↔CD・区分↔KBN・日付↔日・名称↔名・長音）
    param([string] $Name)
    $pairs = @(@('番号', 'NO'), @('番号', '№'), @('番号', 'ナンバー'), @('コード', 'CD'), @('区分', 'KBN'), @('日付', '日'), @('名称', '名'))
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($p in $pairs) {
        foreach ($d in @(@($p[0], $p[1]), @($p[1], $p[0]))) {
            if ($Name.Length -gt $d[0].Length -and $Name.EndsWith($d[0])) { $list.Add($Name.Substring(0, $Name.Length - $d[0].Length) + $d[1]) }
        }
    }
    if ($Name.EndsWith('ー') -and $Name.Length -gt 2) { $list.Add($Name.Substring(0, $Name.Length - 1)) }
    elseif ($Name -match '[ザバタダラリパ]$') { $list.Add($Name + 'ー') }
    return @($list | Where-Object { $_.Length -ge 2 -and $_ -ne $Name } | Select-Object -Unique)
}

# ---------------------------------------------------------------- 入力項目の観点（-DesignPath の画面項目の必須・型・桁数とコード値）


$inputItems = New-Object System.Collections.Generic.List[object]
$codeValues = @{}   # コード区分名（ConvertTo-Label 済み）→ 表示名の一覧
$messages = @{}     # メッセージID（大文字）→ 文言と場所
if ($DesignPath) {
    $searchArgs = @{ Pattern = '\S'; Path = $DesignPath; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $designBooks = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
    foreach ($book in $designBooks.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # コード定義書: コード区分名ごとの表示名（選択の項目の値ごとのケースを確かめる）
            foreach ($t in (Find-DesignTables $grid @('コード区分名', '表示名') 'コード区分名')) {
                foreach ($dr in $t.Rows) {
                    $k = ConvertTo-Label ([string]$dr.Cells[$t.ColumnOf['コード区分名']])
                    $name = ([string]$dr.Cells[$t.ColumnOf['表示名']]).Trim()
                    if (-not $k -or -not $name) { continue }
                    if (-not $codeValues.ContainsKey($k)) { $codeValues[$k] = New-Object System.Collections.Generic.List[string] }
                    if (-not $codeValues[$k].Contains($name)) { $codeValues[$k].Add($name) }
                }
            }
            # メッセージ一覧: メッセージID → 文言（テストに書いたメッセージ ID と文言を確かめる）
            foreach ($t in (Find-DesignTables $grid @('メッセージID', 'メッセージ') 'メッセージID')) {
                foreach ($dr in $t.Rows) {
                    $id = ([string]$dr.Cells[$t.ColumnOf['メッセージID']]).Normalize($nfkc).Trim().ToUpperInvariant()
                    $text = ([string]$dr.Cells[$t.ColumnOf['メッセージ']]).Normalize($nfkc).Trim()
                    # 画面設計書のメッセージ欄（ID だけを書いた列）は文言ではない
                    if ($id -notmatch '^[A-Z]+\d+$' -or -not $text -or $text -match '^[A-Za-z]+\d+$' -or $messages.ContainsKey($id)) { continue }
                    $messages[$id] = [pscustomobject]@{ Text = $text; File = $book.File; Sheet = $sheetName; Row = $dr.Row }
                }
            }
            $screen = Find-LabelValue $grid '画面ID'
            if (-not $screen) { continue }
            foreach ($t in (Find-DesignTables $grid @('項目名', '桁数') '項目名')) {
                $cell = { param($dr, $name) if ($t.ColumnOf.ContainsKey($name)) { ([string]$dr.Cells[$t.ColumnOf[$name]]).Normalize($nfkc).Trim() } else { '' } }
                foreach ($dr in $t.Rows) {
                    # 出力だけの項目（入出力が O）は入力のテストが無い
                    if ($t.ColumnOf.ContainsKey('入出力') -and (& $cell $dr '入出力') -notmatch 'I') { continue }
                    $digits = & $cell $dr '桁数'
                    $inputItems.Add([pscustomobject]@{
                            Screen = $screen.Trim(); Item = ([string]$dr.Cells[$t.ColumnOf['項目名']]).Trim(); Digits = $(if ($digits -match '^\d+$') { [int]$digits } else { $null })
                            Required = ((& $cell $dr '必須') -match '^(○|〇|◯|必須|要|Y|YES|1)$'); Type = (& $cell $dr '型'); Kind = (& $cell $dr '区分')
                            File = $book.File; Sheet = $sheetName; Row = $dr.Row
                        })
                }
            }
        }
    }
    # 画面ごとの項目名（揺れとみなさないため）
    $screenItems = @{}
    foreach ($item in $inputItems) {
        if (-not $screenItems.ContainsKey($item.Screen)) { $screenItems[$item.Screen] = New-Object 'System.Collections.Generic.HashSet[string]' }
        [void]$screenItems[$item.Screen].Add((ConvertTo-Label $item.Item).ToUpperInvariant())
    }
    foreach ($item in $inputItems) {
        $screenRx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9]){0}(?![A-Za-z0-9])' -f [regex]::Escape($item.Screen)), 'IgnoreCase')
        $screenCases = @($cases | Where-Object { $screenRx.IsMatch(('{0} {1} {2}' -f $_.Function, $_.Place.File, $_.Place.Sheet).Normalize($nfkc)) })
        if ($screenCases.Count -eq 0) { continue }
        $itemLabel = ConvertTo-Label $item.Item
        # その項目の名前が出てくるケースの文字
        $itemTexts = @($screenCases | ForEach-Object { (@($_.Values.Values) -join ' ').Normalize($nfkc) } | Where-Object { (ConvertTo-Label $_).Contains($itemLabel) })
        $where = '{0} {1} {2} 行目' -f $item.File, $item.Sheet, $item.Row

        # 項目名の揺れ: 設計書の項目名と違う書き方（受注番号 → 受注No など）がケースに出てくる
        $upper = $itemLabel.ToUpperInvariant()
        foreach ($variant in (Get-NameVariants $upper)) {
            if ($screenItems[$item.Screen].Contains($variant)) { continue }
            # 別の書き方が設計書の項目名より短いときは、項目名の部分を除いて探す（「顧客名称」の中の「顧客名」を揺れとしない）
            $hit = @($screenCases | Where-Object {
                    $t = (ConvertTo-Label (@($_.Values.Values) -join ' ').Normalize($nfkc)).ToUpperInvariant()
                    if (-not $variant.Contains($upper)) { $t = $t.Replace($upper, '') }
                    $t.Contains($variant)
                })
            if ($hit.Count -gt 0) {
                $more = $(if ($hit.Count -gt 1) { '。ほかに {0} 件' -f ($hit.Count - 1) } else { '' })
                Add-Issue '警告' '項目名の揺れ' $hit[0] '' ("{0} のケースに「{1}」と書かれています。設計書の項目名は「{2}」です（{3}{4}）" -f $item.Screen, $variant, $item.Item, $where, $more)
                break
            }
        }

        # 必須の項目: 未入力のケース
        if ($item.Required -and @($itemTexts | Where-Object { $_ -match $requiredPattern }).Count -eq 0) {
            Add-Issue '警告' '必須のテストが無い' $screenCases[0] '' ("{0} の必須の入力項目「{1}」（{2}）を未入力にするケースが見つかりません" -f $item.Screen, $item.Item, $where)
        }
        # 型が日付・数値・半角の項目: 不正な値のケース
        $typeKey = @($typePatterns.Keys | Where-Object { $item.Type -match $_ })[0]
        if ($typeKey -and @($itemTexts | Where-Object { $_ -match $typePatterns[$typeKey] }).Count -eq 0) {
            Add-Issue '警告' '型のテストが無い' $screenCases[0] '' ("{0} の入力項目「{1}」（型 {2}。{3}）に、{4}として正しくない値を入れるケースが見つかりません" -f $item.Screen, $item.Item, $item.Type, $where, $typeKey)
        }
        # 選択の項目で、コード定義書に同じ名前のコード区分がある: 値（表示名）ごとのケース
        $codes = $codeValues[$itemLabel]
        if ($codes) {
            $missing = @($codes | Where-Object { $label = ConvertTo-Label $_; @($itemTexts | Where-Object { (ConvertTo-Label $_).Contains($label) }).Count -eq 0 })
            if ($missing.Count -gt 0) {
                Add-Issue '警告' 'コード値のテストが無い' $screenCases[0] '' ("{0} の入力項目「{1}」（{2}）の値のうち、{3} を選ぶケースが見つかりません（コード定義書の「{1}」の {4} 値）" -f $item.Screen, $item.Item, $where, (@($missing | ForEach-Object { '「' + $_ + '」' }) -join '・'), $codes.Count)
            }
        }
        # 桁数の境界: 選ぶ項目（コード値・選択・チェック）と日付は、値の確かめ方が別なので見ない
        if ($null -eq $item.Digits -or $codes -or $item.Kind -match '選択|ラジオ|チェック|プルダウン|リスト|コンボ' -or $item.Type -match '日付|時刻') { continue }
        $covered = @($itemTexts | Where-Object { $_ -match $boundaryPattern -or $_ -match ('(?<!\d){0}(?!\d)' -f $item.Digits) -or $_ -match ('(?<!\d){0}(?!\d)' -f ($item.Digits + 1)) })
        if ($covered.Count -eq 0) {
            Add-Issue '警告' '境界値のテストが無い' $screenCases[0] '' ("{0} の入力項目「{1}」（{2} 桁。{3} {4} {5} 行目）の桁数の境界（{2} 桁・{6} 桁）を確かめるケースが見つかりません" -f $item.Screen, $item.Item, $item.Digits, $item.File, $item.Sheet, $item.Row, ($item.Digits + 1))
        }
    }
}

# ---------------------------------------------------------------- メッセージ（-DesignPath のメッセージ一覧）


function Test-MessageText {
    #     テストに書いた文言が、メッセージ一覧の文言（{0} などは何かの文字）と合うか。空白と最後の「。」の違いは見ない
    param([string] $Template, [string] $Text)
    $strip = { param($t) ($t -replace '\s', '').TrimEnd('。') }
    $pattern = [regex]::Escape((& $strip $Template)) -replace '\\\{\d+\}', '.+?'
    return (& $strip $Text) -match ('^' + $pattern + '$')
}

if ($messages.Count -gt 0) {
    # メッセージ一覧の ID の形（E0001 なら E と 4 桁）だけを探す（画面 ID・要件 ID などを拾わないように）
    $shapes = @($messages.Keys | ForEach-Object { $m = [regex]::Match($_, '^([A-Z]+)(\d+)$'); '{0}\d{{{1}}}' -f $m.Groups[1].Value, $m.Groups[2].Value.Length } | Sort-Object -Unique)
    $idRx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9])(?:{0})(?![0-9])' -f ($shapes -join '|')), 'IgnoreCase')
    foreach ($c in $cases) {
        $field = ''
        if ($c.Values.ContainsKey('期待結果')) { $field = '期待結果' }
        # 同じ列を 2 つの項目として読むこともあるので、同じ文字は 1 回だけ見る
        $text = (@($c.Values.Values | Select-Object -Unique) -join ' ').Normalize($nfkc)
        $ids = @($idRx.Matches($text) | ForEach-Object { $_.Value.ToUpperInvariant() } | Select-Object -Unique)
        foreach ($id in @($ids | Where-Object { -not $messages.ContainsKey($_) })) {
            Add-Issue '警告' 'メッセージIDが一覧に無い' $c $field ("メッセージ ID {0} がメッセージ一覧にありません（誤記か、一覧への追加漏れ）" -f $id)
        }
        $known = @($ids | Where-Object { $messages.ContainsKey($_) })
        if ($known.Count -eq 0) { continue }
        # 「」で書いたメッセージらしい文言（。で終わる・ください／ません など）が、ケースのメッセージ ID の文言のどれとも合わない
        foreach ($quote in @([regex]::Matches($text, '「([^「」]+)」') | ForEach-Object { $_.Groups[1].Value.Trim() } | Select-Object -Unique)) {
            if ($quote -notmatch '。$|ください|ません|しました|ですか') { continue }
            if (@($known | Where-Object { Test-MessageText $messages[$_].Text $quote }).Count -gt 0) { continue }
            $defined = (@($known | ForEach-Object { '{0}「{1}」（{2} {3} {4} 行目）' -f $_, $messages[$_].Text, $messages[$_].File, $messages[$_].Sheet, $messages[$_].Row })) -join '・'
            Add-Issue '警告' 'メッセージの文言が違う' $c $field ("「{0}」が、メッセージ一覧の {1} と合いません" -f $quote, $defined)
        }
    }
}

# ---------------------------------------------------------------- 結果


$sorted = @($issues | Sort-Object @{ e = { @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }[$_.Level] } }, File, Sheet, Row, Col, Seq)
$sorted

if ($OutXlsx) {
    $issueRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.File, $i.Location, $i.Level, $i.Item, $i.Function, $i.Id, $i.Detail) }
        if ($i.Sheet) { $row.Link = @{ Column = 1; Path = $i.Path; Location = ("'{0}'!{1}" -f $i.Sheet.Replace("'", "''"), $i.Cell) } }
        $issueRows.Add($row)
    }
    $functionRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $functions) {
        $mine = @($issues | Where-Object { $_.Function -eq $f.Function })
        $functionRows.Add(@{ Values = @($f.Function, [double]$f.Cases, [double]$f.ErrorCases, [double]$f.ErrorRatio, $f.Basis, [double]@($mine | Where-Object Level -eq 'エラー').Count, [double]@($mine | Where-Object Level -eq '警告').Count) })
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '指摘'; Columns = @(@{ Header = 'ファイル'; Width = 28 }, @{ Header = '場所'; Width = 18 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 20 }, @{ Header = '機能'; Width = 18 }, @{ Header = 'テストID'; Width = 10 }, @{ Header = '内容'; Width = 70 }); Rows = $issueRows },
            @{ Name = '機能別'; Columns = @(@{ Header = '機能'; Width = 24 }, @{ Header = 'ケース'; Width = 7 }, @{ Header = '異常系'; Width = 7 }, @{ Header = '異常系の割合'; Width = 9 }, @{ Header = '異常系の数え方'; Width = 20 }, @{ Header = 'エラー'; Width = 7 }, @{ Header = '警告'; Width = 7 }); Rows = $functionRows }, (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ('テストケース {0} 件（機能 {1}）: エラー {2} 件 / 警告 {3} 件' -f $cases.Count, $functions.Count, @($issues | Where-Object Level -eq 'エラー').Count, @($issues | Where-Object Level -eq '警告').Count)
$byItem = @($issues | Group-Object Item | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
if ($byItem.Count -gt 0) { Write-Host ('内訳: {0}' -f ($byItem -join '・')) }
if ($DesignPath) { Write-Host ('入力項目 {0} 件（必須 {1}・桁数あり {2}）・コード区分 {3} 件・メッセージ {4} 件を確かめました' -f $inputItems.Count, @($inputItems | Where-Object Required).Count, @($inputItems | Where-Object { $null -ne $_.Digits }).Count, $codeValues.Count, $messages.Count) }
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テスト仕様書（テストケースの一覧）の書き方の抜け・あいまいさ・重複と、異常系・必須・型・境界値・コード値のケースの不足を見つけます（Excel 不要）。
#
# .DESCRIPTION
#     テスト仕様書（Excel・CSV）から、期待結果・手順・前提条件・入力値・区分・結果などの列を見出しの文字で探して読み、次を確かめます。
#
#       期待結果が空・あいまいな期待結果（「正しく」「適切に」「仕様どおり」など）・手順が空
#       テストIDの重複・同じ内容のケース（前提条件・手順・入力値・期待結果がすべて同じ）・結果の値が読めない
#       異常系が少ない機能（区分の列、無ければ期待結果の言葉から推定。-MinErrorRatio 未満）
#       境界値のテストが無い入力項目（-DesignPath の画面項目定義の桁数。項目名と、桁・文字・最大などの言葉か桁数の数字が出てくるケースを探す）
#       必須のテストが無い入力項目（必須 ○ の項目。項目名と、未入力・空欄などの言葉が出てくるケースを探す）
#       型のテストが無い入力項目（型が日付・数値・半角の項目。項目名と、その型として正しくない値の言葉が出てくるケースを探す）
#       コード値のテストが無い入力項目（コード定義書に同じ名前のコード区分がある項目。表示名ごとに、項目名と表示名が出てくるケースを探す）
#       メッセージ一覧に無いメッセージ ID と、メッセージ一覧と違う文言（-DesignPath のメッセージ一覧。「」で書いた文言を比べる）
#       項目名の揺れ（画面のケースに、画面項目定義の項目名の別の書き方（受注番号 → 受注No・顧客コード → 顧客CD など）が出てくる）
#
#     見出しの候補は tools\test-quality\見出し.csv（Measure-TestQuality.ps1 と同じ）で足せます。
#
# .PARAMETER Path
#     テスト仕様書のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER DesignPath
#     設計書のフォルダ・ファイル。画面項目定義の入力項目の必須・型・桁数とコード定義書から、未入力・不正な値・境界値・選ぶ値ごとのケースがあるかを確かめます。
#     メッセージ一覧から、ケースに書いたメッセージ ID と文言が一覧と合うかも確かめます。
#
# .PARAMETER MinErrorRatio
#     異常系のケースの割合の下限（既定 0.2）。ケースが 5 件以上の機能で、これより少なければ警告します。
#
# .PARAMETER VagueWords
#     あいまいな期待結果とみなす言葉。
#
# .PARAMETER Headers
#     見出しの候補の CSV。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘・機能別）。
#
# .EXAMPLE
#     .\Test-TestSpec.ps1 D:\テスト仕様書 -DesignPath D:\設計書 -OutXlsx D:\テスト仕様書のチェック.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     異常系・境界値は書かれた言葉から機械的に判定します。最終的にはレビューで確かめてください。
