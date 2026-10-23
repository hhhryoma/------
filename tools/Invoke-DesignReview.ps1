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
$promptDir = Join-Path $OutDir 'prompts'
$answerDir = Join-Path $OutDir 'answers'
$manifestPath = Join-Path $OutDir 'requests.json'
$reportPath   = Join-Path $OutDir 'レビュー結果.xlsx'

$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$invariant = [System.Globalization.CultureInfo]::InvariantCulture

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


function Read-Checklist {
    #     観点表（xlsx / csv）を読む。1 行目または「観点ID」を含む行を見出しとし、列は見出しの文字で特定する。
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
        if (-not $p.Criteria.Trim()) {
            & $add '警告' $p.Row $p.Id '判定基準が空欄です（AI の判定がぶれやすくなります）'
        }
        else {
            # 「」や例示の中は判定基準そのものではないので除いて調べる
            $plain = [regex]::Replace($p.Criteria, '「[^」]*」|（例[^）]*）|\(例[^)]*\)', '')
            $found = @([regex]::Matches($plain, $vague) | ForEach-Object { $_.Value } | Sort-Object -Unique)
            if ($found.Count -gt 0) {
                & $add '警告' $p.Row $p.Id ("判定基準にあいまいな言葉があります（{0}）。何が書かれていれば OK かを具体的に書いてください" -f ($found -join '、'))
            }
        }
        if (($p.Point + $p.Criteria) -match '[|｜]') {
            & $add '警告' $p.Row $p.Id '観点・判定基準に「|」があります（プロンプトの区切りと紛れるため、別の記号にしてください）'
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

# ---------------------------------------------------------------- プロンプト


$promptHeader = @'
あなたはシステム開発の設計レビュー担当です。下の【設計書】を【レビュー観点】に沿って確認し、【回答形式】どおりに結果を出力してください。

【ルール】
- 観点ごとに判定してください。判定は OK／NG／要確認 のどれかです。
  - NG：観点を満たしていない箇所がある。箇所ごとに 1 行ずつ出力する。
  - 要確認：設計書の記載だけでは判断できない。理由を書く。
  - OK：問題がない。観点ごとに 1 行だけ出力し、場所・指摘内容・根拠は「-」とする。
- すべての観点について、必ず 1 行以上出力してください。
- 場所は「シート名!セル番地」で書いてください。設計書の各行は「行番号| 列:値」の形で、例えば「12| K:必須」なら K12 です。
- 設計書に書かれていないことを推測で補わないでください。
- 【参照資料】は判定の参考にするだけで、指摘の対象にはしないでください。
- 回答形式の行以外（前置き、まとめ、表の見出し）は出力しないでください。

【回答形式】1 行に 1 件、区切りは半角の「|」
観点ID | 判定 | 場所 | 指摘内容 | 根拠
（例）
V01 | NG | 画面項目定義!G12 | 桁数が記入されていない | 入力項目は文字・数値の場合に桁数が必要
V02 | OK | - | - | -
'@
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
        [void]$sb.Append(('{0} | {1} | {2}' -f $p.Id, $p.Point.Replace("`n", ' '), $p.Criteria.Replace("`n", ' '))).Append("`n")
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
    #     Copilot の回答から「観点ID | 判定 | 場所 | 指摘内容 | 根拠」の行を取り出す。
    #     Markdown の表、コードブロック、表をコピーしたときのタブ区切りのどれでも読める。
    param([string] $Text, [string[]] $PointIds)

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
        $reason = & $get 4
        if ($fields.Count -gt 5) { $reason = ($fields[4..($fields.Count - 1)] -join ' | ') }

        $entries.Add([pscustomobject]@{
            PointId  = $known[$id]
            Verdict  = $verdict
            Location = & $get 2
            Finding  = & $get 3
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
        $targetPoints = @($points | Where-Object { Test-NameLike $f.Name $_.Targets })
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

            $chunks = Split-SheetTexts -SheetTexts $sheetTexts -Budget $budget
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
}
else {
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
    }
    [System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 5), $utf8Bom)

    Write-Host ''
    Write-Host ("依頼 {0} 件を作成しました（1 件あたり最大 {1} 文字）: {2}" -f $requests.Count, (($requests | Measure-Object Chars -Maximum).Maximum), $promptDir)
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
            $entries = ConvertFrom-Answer -Text $answer -PointIds $r.PointIds
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

# 設計書 × 観点 ごとに、分割された依頼の回答をまとめる
$coverage = [ordered]@{}
$findings = New-Object System.Collections.Generic.List[object]
foreach ($r in $requests) {
    $ansPath = Join-Path $answerDir "$($r.Id).txt"
    $state = '未回答'
    $entries = @()
    if (Test-Path -LiteralPath $ansPath) {
        $entries = ConvertFrom-Answer -Text ([System.IO.File]::ReadAllText($ansPath)) -PointIds $r.PointIds
        $state = '回答済み'
    }
    $requestRows.Add(@{ Values = @($r.Id, $r.Relative, $r.Part, ($r.PointIds -join ', '), $r.Chars, $state, "prompts\$($r.Id).txt") })

    foreach ($pointId in $r.PointIds) {
        $key = "$($r.File)|$pointId"
        if (-not $coverage.Contains($key)) {
            $coverage[$key] = [pscustomobject]@{ File = $r.File; Relative = $r.Relative; PointId = $pointId; Verdicts = New-Object System.Collections.Generic.List[string]; Requests = New-Object System.Collections.Generic.List[string]; Missing = $false }
        }
        $cv = $coverage[$key]
        $cv.Requests.Add($r.Id)
        $mine = @($entries | Where-Object PointId -eq $pointId)
        if ($mine.Count -eq 0) { $cv.Missing = $true }
        foreach ($e in $mine) {
            $cv.Verdicts.Add($e.Verdict)
            if ($e.Verdict -ne 'OK') {
                $findings.Add([pscustomobject]@{ Request = $r; Entry = $e })
            }
        }
    }
}

$sorted = $findings | Sort-Object @{ e = { $_.Request.Relative } }, @{ e = { $s = $pointById[$_.Entry.PointId].Severity; if ($severityOrder.ContainsKey($s)) { $severityOrder[$s] } else { 9 } } }, @{ e = { $_.Entry.PointId } }
$n = 0
foreach ($f in $sorted) {
    $n++
    $p = $pointById[$f.Entry.PointId]
    $loc = Resolve-Location -Location $f.Entry.Location -SheetNames @($f.Request.Sheets)
    $note = ''
    if ($f.Entry.Location -and -not $loc) { $note = '場所のシート名・セル番地を設計書で確認してください' }
    $row = @{ Values = @($n, $f.Entry.Verdict, $p.Severity, $p.Id, $p.Category, $p.Point, $f.Request.Relative, $f.Entry.Location, $f.Entry.Finding, $f.Entry.Reason, $f.Request.Id, '', '', '', '', $note) }
    if ($loc) {
        $row.Link = @{ Column = 7; Path = $f.Request.File; Location = "'{0}'!{1}" -f $loc.Sheet.Replace("'", "''"), $loc.Cell }
    }
    $findingRows.Add($row)
}

foreach ($cv in $coverage.Values) {
    $p = $pointById[$cv.PointId]
    $ng = @($cv.Verdicts | Where-Object { $_ -eq 'NG' }).Count
    if ($ng -gt 0) { $verdict = 'NG' }
    elseif ($cv.Verdicts -contains '要確認') { $verdict = '要確認' }
    elseif ($cv.Missing) { $verdict = '未回答' }
    else { $verdict = 'OK' }
    # 分割した依頼の一部だけ未回答なら、OK と言い切らない
    if ($cv.Missing -and $verdict -eq 'OK') { $verdict = '未回答' }
    $statusRows.Add(@{ Values = @($cv.Relative, $p.Id, $p.Category, $p.Point, $p.Severity, $verdict, $ng, ($cv.Requests -join ', ')); Link = @{ Column = 0; Path = $cv.File; Location = "'{0}'!A1" -f ((@($requests | Where-Object File -eq $cv.File))[0].Sheets[0]).Replace("'", "''") } })
}

$sheets = @(
    @{
        Name = '指摘一覧'
        Columns = @(
            @{ Header = 'No'; Width = 5 }, @{ Header = '判定'; Width = 8 }, @{ Header = '重要度'; Width = 7 }, @{ Header = '観点ID'; Width = 8 },
            @{ Header = '分類'; Width = 10 }, @{ Header = '観点'; Width = 30 }, @{ Header = '設計書'; Width = 30 }, @{ Header = '場所'; Width = 20 },
            @{ Header = '指摘内容'; Width = 45 }, @{ Header = '根拠'; Width = 35 }, @{ Header = '依頼No'; Width = 8 },
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
Write-Host ''
Write-Host ("依頼 {0} 件中 {1} 件回答済み / 観点の判定: {2} / 指摘 {3} 件" -f $requests.Count, $answeredCount, ($statusCounts -join '・'), $findingRows.Count)
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
#
#     途中で中断しても、回答済みの分は保存され、次回は続きから再開します。
#     設計書の読み取りには同じフォルダの Search-Excel.ps1 を使います（Excel 不要）。
#
# .PARAMETER Path
#     レビューする設計書のフォルダまたはファイル（複数指定可）。
#
# .PARAMETER Checklist
#     観点表（xlsx または UTF-8 の csv）。省略時は review-checklists\詳細設計レビュー観点.xlsx。
#     列: 観点ID / 分類 / 観点 / 判定基準 / 重要度 / 対象設計書 / 参照設計書 / 有効
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
