[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path,

    [string] $Checklist,

    [string] $OutDir = '.\design-review',

    [ValidateRange(2000, 200000)]
    [int] $MaxChars = 8000,

    [switch] $PromptOnly,
    [switch] $ImportOnly,
    [switch] $CheckOnly,

    [string] $SheetNameExclude,

    [string] $Glossary,

    [ValidateSet('Rows', 'Markdown')]
    [string] $TextFormat = 'Rows',

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

$OutDir    = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$Checklist = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Checklist)
if ($Glossary) {
    $Glossary = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Glossary)
    if (-not (Test-Path -LiteralPath $Glossary)) { throw "用語集が見つかりません: $Glossary" }
}
$promptDir = Join-Path $OutDir 'prompts'
$answerDir = Join-Path $OutDir 'answers'
$manifestPath = Join-Path $OutDir 'requests.json'
$reportPath   = Join-Path $OutDir 'レビュー結果.xlsx'

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$nfkc = [System.Text.NormalizationForm]::FormKC

$severityOrder = @{ '高' = 0; '中' = 1; '低' = 2 }

# ---------------------------------------------------------------- 設計書のテキスト化


function Get-ExcelCells {
    #     Search-Excel.ps1 で全セル（値・コメント・図形・シート名）を読み出す。
    #     ブックの並び順（シート順、行優先）のまま返る。
    param([string] $FilePath)

    $warnings = $null
    $searchArgs = @{
        Pattern = '\S'
        Path    = [System.Management.Automation.WildcardPattern]::Escape($FilePath)
        Target  = @('Value', 'Comment', 'Shape', 'SheetName')
    }
    if ($SheetNameExclude) { $searchArgs['SheetNameExclude'] = $SheetNameExclude }

    $hits = @(& $searchScript @searchArgs -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)
    if ($warnings) { foreach ($w in $warnings) { Write-Warning $w.Message } }
    return $hits
}

function ConvertTo-DesignText {
    #     設計書を AI が読める行形式のテキストにする。
    #       「行番号| 列:値 | 列:値 ...」の形で、セル番地を AI が引用できるようにする。
    #     方眼紙の帳票ヘッダ（各シート 1〜3 行目）が前のシートと同じなら省略して文字数を節約する。
    #     返り値: シート名 → テキスト の順序付き辞書
    param($Cells)

    if ($TextFormat -eq 'Markdown') { return ConvertTo-DesignMarkdown $Cells }

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
        $cells = Get-ExcelCells -FilePath $FilePath | Where-Object { $_.Kind -eq '値' }
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
    foreach ($g in ($points | Group-Object { $_.Id.ToUpperInvariant() } | Where-Object { $_.Count -gt 1 })) {
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
    #     レビュー対象の設計書を列挙する（出力フォルダと観点表は除く）
    $files = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($p in $Path) {
        $items = @(Get-PathItem $p)
        if ($items.Count -eq 0) { Write-Warning "見つかりません: $p" }
        foreach ($item in $items) {
            if ($item.PSIsContainer) {
                $root  = $item.FullName.TrimEnd('\', '/')
                $found = Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue
            }
            else {
                $root  = [System.IO.Path]::GetDirectoryName($item.FullName)
                $found = @($item)
            }
            foreach ($f in $found) {
                if (@('.xlsx', '.xlsm', '.xls') -notcontains $f.Extension.ToLowerInvariant()) { continue }
                if ($f.Name.StartsWith('~$')) { continue }
                if ($f.FullName.StartsWith($OutDir, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ([string]::Equals($f.FullName, $Checklist, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($seen.ContainsKey($f.FullName)) { continue }
                $seen[$f.FullName] = $true
                $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
                $files.Add([pscustomobject]@{ Path = $f.FullName; Name = $f.Name; Relative = $rel })
            }
        }
    }
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
        if (@($templates | Where-Object { $_.Contains('{n}') }).Count -gt 0) { $needLength[$key] = $true }
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
            $table = @(Find-DesignTables $grid @($phys, $type) $phys)[0]
            if (-not $table -or $table.Rows.Count -eq 0) { continue }
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
                if ($Glossary) { $termArgs['Glossary'] = $Glossary }
                foreach ($v in @(& (Join-Path $PSScriptRoot 'Test-ExcelTerms.ps1') @termArgs 6>$null 3>$null)) {
                    # 揺れの表記ごと・ファイルごとに 1 件（場所は最初に出てくるセル）
                    foreach ($g in @($v.Occurrences | Group-Object Path)) {
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
                if ($null -eq $books) { $books = ConvertTo-DesignBooks @(& $searchScript -Pattern '\S' -Path $paths -Target Value 6>$null 3>$null) }
                if ($src -eq '画面') { Add-ScreenRuleFindings $books $names $add }
                else { Add-TableRuleFindings $books $names $add }
            }
        }
        catch {
            Write-Warning ("ツールの確認（{0}）を行えませんでした。この確認を使う観点は「未回答」になります: {1}" -f ($names -join '・'), $_.Exception.Message)
            foreach ($n in $names) { $result.Failed[$n] = $true }
        }
    }
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
# 回答の形（引用の列があるか）。依頼の一覧に残し、-ImportOnly で前の形の依頼を読むときに使う
$answerFormat = '引用付き'
# プロンプトの改行は OS によらず LF にそろえる。AppendLine や CRLF のスクリプトファイルのままだと
# Windows では 1 行ごとに 1 文字多く数えられ、同じ設計書でも依頼の分割数が変わってしまう
$promptHeader = $promptHeader.Replace("`r`n", "`n")
if ($TextFormat -eq 'Markdown') {
    $promptHeader = $promptHeader.Replace('設計書の各行は「行番号| 列:値」の形で、例えば「12| K:必須」なら K12 です。',
        '設計書は Markdown で、「## 」の見出しがシート名です。箇条書きの「K12: 必須」はセル K12 です。表は 1 列目が行番号で、見出しの「K:桁数」の K が列です（例えば 12 行目の桁数なら K12）。')
}

function New-PromptText {
    param($Points, [string] $FileName, [string] $Body, [string] $PartNote, [string] $RefText)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append($promptHeader).Append("`n")
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
    #     1 シートで Budget を超える場合は行単位で分ける。
    param($SheetTexts, [int] $Budget)

    $chunks = New-Object System.Collections.Generic.List[string]
    $current = New-Object System.Text.StringBuilder
    foreach ($name in $SheetTexts.Keys) {
        $text = $SheetTexts[$name]
        if ($current.Length -gt 0 -and $current.Length + $text.Length -gt $Budget) {
            $chunks.Add($current.ToString()); [void]$current.Clear()
        }
        if ($text.Length -le $Budget) { [void]$current.Append($text); continue }

        # 長いシートは行で分割し、続きには見出しを付け直す（Markdown の表の途中で分けたときは、表の見出しも付け直す）
        $prev = ''
        $tableHead = ''
        foreach ($line in ($text -split "`r?`n")) {
            if (-not $line) { continue }
            $isSeparator = $line -match '^\|(\s*---\s*\|)+$'
            if (-not $line.StartsWith('|')) { $tableHead = '' }
            elseif ($isSeparator) { $tableHead = $prev + "`n" + $line }
            if ($current.Length + $line.Length + 2 -gt $Budget -and $current.Length -gt 0) {
                $chunks.Add($current.ToString()); [void]$current.Clear()
                if ($TextFormat -eq 'Markdown') { [void]$current.Append("## $name（続き）").Append("`n") }
                else { [void]$current.Append("### シート: $name（続き）").Append("`n") }
                if ($isSeparator) { [void]$current.Append($prev).Append("`n") }
                elseif ($tableHead) { [void]$current.Append($tableHead).Append("`n") }
            }
            [void]$current.Append($line).Append("`n")
            $prev = $line
        }
    }
    if ($current.Length -gt 0) { $chunks.Add($current.ToString()) }
    return ,$chunks
}

function Get-TextHash {
    param([string] $Text)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text))).Replace('-', '').Substring(0, 16) }
    finally { $sha.Dispose() }
}

# ---------------------------------------------------------------- 回答の読み取り


function ConvertFrom-Answer {
    #     Copilot の回答から「観点ID | 判定 | 場所 | 引用 | 指摘内容 | 根拠」の行を取り出す。
    #     Markdown の表、コードブロック、表をコピーしたときのタブ区切りのどれでも読める。
    #     Quote を付けないとき（前の形の依頼）と、引用の列の無い 5 列の行は「観点ID | 判定 | 場所 | 指摘内容 | 根拠」として読む
    param([string] $Text, [string[]] $PointIds, [switch] $Quote)

    $known = @{}
    foreach ($id in $PointIds) { $known[$id.ToUpperInvariant()] = $id }

    $entries = New-Object System.Collections.Generic.List[object]
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('```')) { continue }
        if ($line -match '^\|?[\s:\-|]+\|?$') { continue }   # Markdown 表の区切り行

        if ($line.Contains('|')) { $fields = $line.Trim('|').Split('|') }
        elseif ($line.Contains("`t")) { $fields = $line.Split("`t") }
        else { continue }
        $fields = @($fields | ForEach-Object { $_.Trim() })
        if ($fields.Count -lt 2) { continue }

        $id = $fields[0].Trim('*', ' ').ToUpperInvariant()
        if (-not $known.ContainsKey($id)) { continue }

        $verdictRaw = $fields[1].Trim('*', ' ')
        if ($verdictRaw -match '^(OK|ＯＫ|○|〇|問題なし)$') { $verdict = 'OK' }
        elseif ($verdictRaw -match '^(NG|ＮＧ|×|指摘)') { $verdict = 'NG' }
        else { $verdict = '要確認' }

        $get = { param($i) if ($fields.Count -gt $i -and $fields[$i] -ne '-' -and $fields[$i] -ne '－') { $fields[$i] } else { '' } }
        # 指摘内容の列（引用の列があれば 1 つ右）
        $base = 3
        $quoteText = ''
        if ($Quote -and $fields.Count -ge 6) { $quoteText = (& $get 3).Trim('「', '」', '"', '“', '”'); $base = 4 }
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
    $m = [regex]::Match($Location, "^\s*'?(?<sheet>[^!']+?)'?\s*[!！]\s*\`$?(?<col>[A-Za-z]{1,3})\`$?(?<row>\d+)")
    if ($m.Success) {
        $sheet = $m.Groups['sheet'].Value.Trim()
        if ($SheetNames -contains $sheet) {
            return [pscustomobject]@{ Sheet = $sheet; Cell = $m.Groups['col'].Value.ToUpperInvariant() + $m.Groups['row'].Value }
        }
        return $null
    }
    # セル番地なしでシート名だけ書かれている場合はシートの先頭へ
    $name = $Location.Trim().Trim("'")
    if ($SheetNames -contains $name) { return [pscustomobject]@{ Sheet = $name; Cell = 'A1' } }
    return $null
}


function ConvertTo-QuoteKey {
    #     引用とセルの文字を比べる形（全角・半角をそろえ、空白・改行・末尾の … を除く）
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).Replace('↵', '') -replace '\s', ''
    return ($t -replace '(…|\.\.\.)$', '')
}

function Read-QuoteCells {
    #     引用を確かめるための、設計書のセルの文字（値・コメント・図形。全シート）。
    #     戻り値: シート名 → セル番地 → @{ Cell; Col; Row; Text; Key（ConvertTo-QuoteKey の形） }
    param([string] $FilePath)
    $sheets = @{}
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value, Comment, Shape 6>$null 3>$null)
    foreach ($h in $hits) {
        if (-not $sheets.ContainsKey($h.Sheet)) { $sheets[$h.Sheet] = @{} }
        $cells = $sheets[$h.Sheet]
        $cell = [string]$h.Cell
        if ($cells.ContainsKey($cell)) { $cells[$cell].Text += [string]$h.Text; continue }
        $m = [regex]::Match($cell, '^([A-Z]+)(\d+)$')
        $col = 0; $row = 0
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
        if ($m.Success) { $origin = @((Get-ColumnIndex $m.Groups[1].Value), [int]$m.Groups[2].Value) }
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
        $sheetTexts = & $getText $f
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
        $allRefText = ''
        if ($allRefs.Count -gt 0) { $allRefText = & $getRefText $allRefs }
        $allFixed = (New-PromptText -Points $targetPoints -FileName $f.Name -Body '' -PartNote '' -RefText $allRefText).Length
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

            $fixedLen = (New-PromptText -Points $gp -FileName $f.Name -Body '' -PartNote '（9/9 の一部。他の部分は別途確認するため、この部分だけで判定してください）' -RefText $refText).Length
            $budget = $MaxChars - $fixedLen
            if ($budget -lt 1000) {
                Write-Warning ("{0}: 観点と参照資料だけで {1} 文字あり、-MaxChars {2} に収まりません。-MaxChars を増やすか、参照設計書を絞ってください（観点 {3}）" -f $f.Relative, $fixedLen, $MaxChars, (($gp | ForEach-Object { $_.Id }) -join ', '))
                $budget = 1000
            }

            # 1 回に収まるなら分けない（分けたときの「n/m の一部」の注記の分を見込んで、収まるのに分けてしまわないよう）
            $wholeLen = (New-PromptText -Points $gp -FileName $f.Name -Body '' -PartNote '' -RefText $refText).Length + $docLength
            if ($wholeLen -le $MaxChars) { $chunks = Split-SheetTexts -SheetTexts $sheetTexts -Budget ([int]::MaxValue) }
            else { $chunks = Split-SheetTexts -SheetTexts $sheetTexts -Budget $budget }
            for ($i = 0; $i -lt $chunks.Count; $i++) {
                $partNote = ''
                if ($chunks.Count -gt 1) {
                    $partNote = '（{0}/{1} の一部。他の部分は別途確認するため、この部分だけで判定してください）' -f ($i + 1), $chunks.Count
                }
                $no++
                $prompt = New-PromptText -Points $gp -FileName $f.Name -Body $chunks[$i] -PartNote $partNote -RefText $refText
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
if ($Path) {
    $designFiles = Get-DesignFiles
    if ($designFiles.Count -eq 0 -and -not $CheckOnly) { throw '設計書（.xlsx / .xlsm / .xls）が見つかりません。' }
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
    # 前の形の依頼（引用の列の無い回答形式）は、引用を読まない
    $quoteFormat = [bool]($manifest.PSObject.Properties['AnswerFormat'] -and $manifest.AnswerFormat -eq $answerFormat)
    if (-not $designFiles) {
        # 依頼を作ったときの設計書（古い依頼の一覧には無いので、依頼のある設計書）
        $designFiles = New-Object System.Collections.Generic.List[object]
        if ($manifest.PSObject.Properties['Files']) {
            foreach ($f in @($manifest.Files)) { $designFiles.Add([pscustomobject]@{ Path = $f.Path; Name = $f.Name; Relative = $f.Relative }) }
        }
        else {
            foreach ($r in $requests) {
                if (@($designFiles | Where-Object { $_.Path -eq $r.File }).Count -eq 0) { $designFiles.Add([pscustomobject]@{ Path = $r.File; Name = [System.IO.Path]::GetFileName($r.File); Relative = $r.Relative }) }
            }
        }
    }
}
else {
    $quoteFormat = $true
    Write-Host '依頼を作成しています...'
    $requests = (New-ReviewRequests -Points $points -Files $designFiles).ToArray()

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
        Created  = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss')
        MaxChars = $MaxChars
        Requests = @($requests | Select-Object Id, File, Relative, Part, PointIds, Sheets, Chars, Hash)
        AnswerFormat = $answerFormat
        # -ImportOnly でツールの確認をする設計書
        Files    = @($designFiles | Select-Object Path, Name, Relative)
    }
    [System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 5), $utf8Bom)

    Write-Host ''
    if ($requests.Count -eq 0) { Write-Host 'AI に依頼する観点はありません（どの観点も、確かめ方がツールです）。' }
    else { Write-Host ("依頼 {0} 件を作成しました（1 件あたり最大 {1} 文字）: {2}" -f $requests.Count, (($requests | Measure-Object Chars -Maximum).Maximum), $promptDir) }
}

$pointById = @{}
foreach ($p in $points) { $pointById[$p.Id] = $p }

# ---- 対話モード: クリップボード経由で Copilot とやり取りする

if (-not $PromptOnly -and -not $ImportOnly) {
    if (-not (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) {
        throw 'この環境ではクリップボードが使えません。-PromptOnly で依頼ファイルを作り、回答を answers フォルダに保存してから -ImportOnly を実行してください。'
    }

    $pending = @($requests | Where-Object { -not (Test-Path -LiteralPath (Join-Path $answerDir "$($_.Id).txt")) })
    Write-Host ''
    Write-Host ("未回答の依頼 {0} 件を順に処理します。" -f $pending.Count)
    Write-Host '手順: Copilot に貼り付けて送信 → 回答をコピー（回答の下のコピー ボタン）→ ここで Enter'
    Write-Host '      s = この依頼をスキップ / q = 中断（回答済みの分は保存され、次回は続きから）'

    $i = 0
    foreach ($r in $pending) {
        $i++
        $prompt = [System.IO.File]::ReadAllText((Join-Path $promptDir "$($r.Id).txt"))
        while ($true) {
            Set-Clipboard -Value $prompt
            Write-Host ''
            Write-Host ("[{0}/{1}] {2}  {3}（{4}）観点 {5}  {6} 文字 → クリップボードにコピーしました" -f $i, $pending.Count, $r.Id, $r.Relative, $r.Part, ($r.PointIds -join ','), $r.Chars) -ForegroundColor Cyan
            $key = Read-Host '回答をコピーしたら Enter'
            if ($key -eq 'q') { break }
            if ($key -eq 's') { break }

            $answer = Get-Clipboard -Raw
            if (-not $answer -or $answer -eq $prompt) {
                Write-Warning 'クリップボードの内容が依頼のままです。Copilot の回答をコピーしてから Enter を押してください。'
                continue
            }
            $entries = ConvertFrom-Answer -Text $answer -PointIds $r.PointIds -Quote:$quoteFormat
            if ($entries.Count -eq 0) {
                Write-Warning '回答から「観点ID | 判定 | ...」の行を読み取れませんでした。回答全体をコピーし直すか、Copilot に「回答形式どおりに出し直して」と依頼してください。'
                continue
            }
            [System.IO.File]::WriteAllText((Join-Path $answerDir "$($r.Id).txt"), $answer, $utf8Bom)
            $answered = @($entries | ForEach-Object { $_.PointId } | Sort-Object -Unique)
            $missing = @($r.PointIds | Where-Object { $answered -notcontains $_ })
            Write-Host ("  取り込みました: NG {0} 件 / 要確認 {1} 件 / OK {2} 観点" -f @($entries | Where-Object Verdict -eq 'NG').Count, @($entries | Where-Object Verdict -eq '要確認').Count, @($entries | Where-Object Verdict -eq 'OK').Count)
            if ($missing.Count -gt 0) {
                Write-Warning ("回答に無い観点があります: {0}（結果では「未回答」になります。やり直す場合は answers\{1}.txt を消して再実行）" -f ($missing -join ', '), $r.Id)
            }
            break
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
    $row = @{ Values = @($n, $f.Entry.Verdict, $f.Source, $p.Severity, $p.Id, $p.Category, $p.Point, $f.Request.Relative, $location, $quote, $check, $f.Entry.Finding, $f.Entry.Reason, $f.Request.Id, '', '', '', '', $note) }
    if ($loc) {
        $row.Link = @{ Column = 8; Path = $f.Request.File; Location = "'{0}'!{1}" -f $loc.Sheet.Replace("'", "''"), $loc.Cell }
    }
    elseif ($f.Source -ne 'AI') {
        # 場所の無いツールの指摘（ファイル名の版数など）は、設計書を開く
        $row.Link = @{ Column = 7; Path = $f.Request.File; Location = '' }
    }
    $findingRows.Add($row)
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
    else { $verdict = 'OK' }
    # 分割した依頼の一部だけ未回答なら、OK と言い切らない
    if ($cv.Missing -and $verdict -eq 'OK') { $verdict = '未回答' }
    # 依頼の無い設計書（ツールだけで確かめた設計書）は、ファイルを開くだけ
    $location = ''
    if ($firstSheet.ContainsKey($cv.File)) { $location = "'{0}'!A1" -f $firstSheet[$cv.File].Replace("'", "''") }
    $statusRows.Add(@{ Values = @($cv.Relative, $p.Id, $p.Category, $p.Point, $p.Severity, $verdict, $ng, ($cv.Requests -join ', ')); Link = @{ Column = 0; Path = $cv.File; Location = $location } })
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
    }
)

try {
    Export-Workbook -Sheets $sheets -FilePath $reportPath
}
catch {
    throw ("レビュー結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $reportPath, $_.Exception.Message)
}

$answeredCount = @($requestRows | Where-Object { $_.Values[5] -eq '回答済み' }).Count
$statusCounts = $statusRows | Group-Object { $_.Values[5] } | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
$toolNote = ''
if ($toolPoints.Count -gt 0) { $toolNote = '（うちツール {0} 件）' -f @($findings | Where-Object { $_.Source -ne 'AI' }).Count }
Write-Host ''
Write-Host ("依頼 {0} 件中 {1} 件回答済み / 観点の判定: {2} / 指摘 {3} 件{4}" -f $requests.Count, $answeredCount, ($statusCounts -join '・'), $findingRows.Count, $toolNote)
if ($quoteCounts.Count -gt 0) {
    Write-Host ("AI の指摘の引用: {0}" -f (@($quoteCounts.Keys | ForEach-Object { '{0} {1}' -f $_, $quoteCounts[$_] }) -join '・'))
    if ($quoteCounts.Contains('見つからない')) { Write-Host '  引用が見つからない指摘は、AI の誤りかもしれません。指摘一覧の「引用の確認」で絞り込んで確かめてください' }
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
#        回答をコピーして Enter を押すと、回答を読み取って保存し、次の依頼に進みます。
#     3. 全回答から「レビュー結果.xlsx」（指摘一覧・観点消化状況・依頼一覧）を作ります。
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
#       入力項目の型・桁数       入力項目の型の空欄と、文字・数値の型の桁数の空欄
#       入力チェックのメッセージ メッセージIDの無い入力チェック
#       初期値                   選択・日付・チェックボックスの入力項目の初期値の空欄
#       ボタンのイベント         イベント定義の無いボタン・リンク
#       メッセージ欄             イベントの処理概要とメッセージ欄のメッセージIDの過不足
#       主キー                   主キーの無いテーブル
#       カラムの型・桁数         カラムの型の空欄と、桁数・小数の要る型（tools\ddl\型の対応.csv）の桁数・小数の空欄
#
#     途中で中断しても、回答済みの分は保存され、次回は続きから再開します。
#     設計書の読み取りには同じフォルダの Search-Excel.ps1 を使います（Excel 不要）。
#
# .PARAMETER Path
#     レビューする設計書のフォルダまたはファイル（複数指定可）。
#
# .PARAMETER Checklist
#     観点表（xlsx または UTF-8 の csv）。省略時は review-checklists\詳細設計レビュー観点.xlsx。
#     列: 観点ID / 分類 / 観点 / 判定基準 / 重要度 / 対象設計書 / 参照設計書 / 有効 / 確かめ方 / ツールの確認
#     確かめ方（AI・ツール・両方）が空欄なら、ツールの確認があれば両方、無ければ AI です。
#
# .PARAMETER OutDir
#     プロンプト・回答・レビュー結果の保存先。省略時は .\design-review
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
#     # プロンプトをファイルで受け渡す
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -PromptOnly
#     .\Invoke-DesignReview.ps1 D:\Design\詳細設計 -ImportOnly
