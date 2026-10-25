[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [switch] $PromptOnly,

    [switch] $ImportOnly,

    [switch] $Force,

    [string] $Previous,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if ($PromptOnly -and $ImportOnly) { throw '-PromptOnly と -ImportOnly は同時に指定できません。' }

# 品質評価のブック（フォルダなら、その中で品質評価のシートがある一番新しい xlsx）
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (Test-Path -LiteralPath $Path -PathType Container) {
    $files = @(Get-ChildItem -LiteralPath $Path -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object @{ e = 'LastWriteTime'; Descending = $true }, Name)
    $book = $null
    foreach ($f in $files) {
        $pkg = Open-ZipPackage $f.FullName
        try { $hasSheet = (Get-WorkbookSheetParts $pkg).Sheets.Contains('品質評価') } finally { $pkg.Zip.Dispose() }
        if ($hasSheet) { $book = $f.FullName; break }
    }
    if (-not $book) { throw "品質評価のブック（「品質評価」のシートがある xlsx）が見つかりません: $Path" }
    $Path = $book
}
if (-not (Test-Path -LiteralPath $Path)) { throw "品質評価のブックが見つかりません: $Path" }
$base = Join-Path (Split-Path -Path $Path -Parent) ([System.IO.Path]::GetFileNameWithoutExtension($Path))
$promptPath = $base + '_定性分析の依頼.txt'
$answerPath = $base + '_定性分析の回答.txt'
$utf8Bom = New-Object System.Text.UTF8Encoding($true)
$nfkc = [System.Text.NormalizationForm]::FormKC

# ---------------------------------------------------------------- 品質評価のブックを読む


function Read-QualityMetrics {
    #     品質評価のシートの指標の行（定性分析・対策の列の場所も）
    param([string] $File)
    $fields = @('区分', '指標', '実績', '増減', '判定', '判定の考え方', '定性分析（原因）', '対策')
    $candidates = @{}
    foreach ($f in $fields) { $candidates[$f] = @(ConvertTo-Label $f) }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($rec in (Read-HeaderRecords -Inputs (Get-InputFiles @($File) '品質評価') -Candidates $candidates -Fields $fields -Need @('指標', '判定', '定性分析（原因）', '対策') -Key @('指標') -What '品質評価' -NoExcel:$NoExcel)) {
        if ($rec.Sheet -ne '品質評価') { continue }
        $v = $rec.Values
        if (-not ([string]$v['指標']).Trim()) { continue }
        $list.Add([pscustomobject]@{
                Area = ([string]$v['区分']).Trim(); Name = ([string]$v['指標']).Trim(); Value = ([string]$v['実績']).Trim(); Change = ([string]$v['増減']).Trim()
                Judgement = ([string]$v['判定']).Normalize($nfkc).Trim(); Basis = ([string]$v['判定の考え方']).Trim()
                Cause = ([string]$v['定性分析（原因）']).Trim(); Measure = ([string]$v['対策']).Trim()
                Row = $rec.Row; CauseColumn = $rec.Columns['定性分析（原因）']; MeasureColumn = $rec.Columns['対策']; No = 0
            })
    }
    if ($list.Count -eq 0) { throw "「品質評価」のシートに、指標・判定・定性分析（原因）・対策の見出しの表がありません: $File" }
    return , $list
}

function Find-PreviousBook {
    #     同じフォルダの、名前が「<タイトル>_日付（8 桁）」で日付が前の品質評価のブックのうち、一番近いもの（無ければ $null）
    param([string] $File)
    $m = [regex]::Match([System.IO.Path]::GetFileNameWithoutExtension($File), '^(.*)_(\d{8})$')
    if (-not $m.Success) { return $null }
    $older = @(Get-ChildItem -LiteralPath (Split-Path -Path $File -Parent) -Filter '*.xlsx' -File | ForEach-Object {
            $o = [regex]::Match($_.BaseName, '^(.*)_(\d{8})$')
            if ($o.Success -and $o.Groups[1].Value -eq $m.Groups[1].Value -and $o.Groups[2].Value -lt $m.Groups[2].Value) { [pscustomobject]@{ Date = $o.Groups[2].Value; Path = $_.FullName } }
        } | Sort-Object Date)
    if ($older.Count -eq 0) { return $null }
    return $older[-1].Path
}

$inputs = Get-InputFiles @($Path) '品質評価'
$metrics = Read-QualityMetrics $Path

# 前回の品質評価の定性分析・対策（対策が効いたかを、今回の実績と並べて確かめてもらう）
if ($Previous) { $Previous = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Previous) }
else { $Previous = Find-PreviousBook $Path }
$previousNotes = @()
if ($Previous) {
    if (-not (Test-Path -LiteralPath $Previous)) { throw "前回の品質評価のブックが見つかりません: $Previous" }
    $current = @{}
    foreach ($m in $metrics) { $current['{0}|{1}' -f $m.Area, $m.Name] = $m }
    $before = Read-QualityMetrics $Previous
    $previousNotes = @($before | Where-Object { $_.Cause -or $_.Measure } | ForEach-Object {
            [pscustomobject]@{ Before = $_; Now = $current['{0}|{1}' -f $_.Area, $_.Name] }
        })
}

$actionFields = @('区分', '対象', '内容')
$actionCandidates = @{}
foreach ($f in $actionFields) { $actionCandidates[$f] = @(ConvertTo-Label $f) }
$actions = @(Read-HeaderRecords -Inputs $inputs -Candidates $actionCandidates -Fields $actionFields -Need @('区分', '対象', '内容') -Key @('内容') -What '要対応' -NoExcel:$NoExcel |
        Where-Object { $_.Sheet -eq '要対応' -and ([string]$_.Values['内容']).Trim() })

# 下書きを書く指標: 判定が × か △（記入済みの行は -Force のときだけ）
$targets = @($metrics | Where-Object { @('×', '△') -contains $_.Judgement -and ($Force -or (-not $_.Cause -and -not $_.Measure)) })
$filled = @($metrics | Where-Object { @('×', '△') -contains $_.Judgement -and -not $Force -and ($_.Cause -or $_.Measure) }).Count
$i = 0
foreach ($m in $targets) { $i++; $m.No = $i }

# ---------------------------------------------------------------- 依頼の文


function ConvertTo-Cell {
    #     表のセルに入れる文字（改行・セル内の改行（Search-Excel の ↵）と | を除く）
    param([string] $Text)
    return ($Text -replace '\r?\n|↵', ' ' -replace '\|', '／').Trim()
}

function New-AnalysisPrompt {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('あなたはシステム開発の品質管理の担当者です。工程完了判定（品質会議）の資料として、品質評価の指標のうち基準から外れたものについて、定性分析（原因）と対策の下書きを書いてください。')
    $lines.Add('')
    $lines.Add('# 前提')
    $lines.Add(('- 資料: {0}' -f [System.IO.Path]::GetFileName($Path)))
    $lines.Add('- 判定は ○（基準どおり）・△（注意）・×（基準外）です。判定の考え方は、その指標の基準です。')
    $lines.Add('- 増減は前回の品質評価からの差です（空なら前回の値がありません）。')
    $lines.Add('')
    $lines.Add('# 基準から外れた指標')
    $lines.Add('| No | 区分 | 指標 | 実績 | 増減 | 判定 | 判定の考え方 |')
    $lines.Add('| --- | --- | --- | --- | --- | --- | --- |')
    foreach ($m in $targets) { $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} |' -f $m.No, (ConvertTo-Cell $m.Area), (ConvertTo-Cell $m.Name), (ConvertTo-Cell $m.Value), (ConvertTo-Cell $m.Change), $m.Judgement, (ConvertTo-Cell $m.Basis))) }
    if ($actions.Count -gt 0) {
        $lines.Add('')
        $lines.Add('# 要対応（基準から外れた設計書・機能、流出のあった工程など）')
        $lines.Add('| 区分 | 対象 | 内容 |')
        $lines.Add('| --- | --- | --- |')
        foreach ($a in $actions) { $lines.Add(('| {0} | {1} | {2} |' -f (ConvertTo-Cell $a.Values['区分']), (ConvertTo-Cell $a.Values['対象']), (ConvertTo-Cell $a.Values['内容']))) }
    }
    if ($previousNotes.Count -gt 0) {
        $lines.Add('')
        $lines.Add(('# 前回の定性分析と対策（{0}）' -f [System.IO.Path]::GetFileName($Previous)))
        $lines.Add('| 区分 | 指標 | 前回の実績 | 前回の判定 | 今回の実績 | 今回の判定 | 前回の定性分析（原因） | 前回の対策 |')
        $lines.Add('| --- | --- | --- | --- | --- | --- | --- | --- |')
        foreach ($n in $previousNotes) {
            $nowValue = '（今回なし）'; $nowJudgement = ''
            if ($n.Now) { $nowValue = $n.Now.Value; $nowJudgement = $n.Now.Judgement }
            $lines.Add(('| {0} | {1} | {2} | {3} | {4} | {5} | {6} | {7} |' -f (ConvertTo-Cell $n.Before.Area), (ConvertTo-Cell $n.Before.Name), (ConvertTo-Cell $n.Before.Value), $n.Before.Judgement, (ConvertTo-Cell $nowValue), $nowJudgement, (ConvertTo-Cell $n.Before.Cause), (ConvertTo-Cell $n.Before.Measure)))
        }
    }
    $others = @($metrics | Where-Object { @('×', '△') -notcontains $_.Judgement })
    if ($others.Count -gt 0) {
        $lines.Add('')
        $lines.Add('# そのほかの指標（参考）')
        $lines.Add('| 区分 | 指標 | 実績 | 判定 |')
        $lines.Add('| --- | --- | --- | --- |')
        foreach ($m in $others) { $lines.Add(('| {0} | {1} | {2} | {3} |' -f (ConvertTo-Cell $m.Area), (ConvertTo-Cell $m.Name), (ConvertTo-Cell $m.Value), $m.Judgement)) }
    }
    $lines.Add('')
    $lines.Add('# 回答の形式')
    $lines.Add('「基準から外れた指標」の No ごとに 1 行ずつ、次の表だけを出力してください（表の前後に文を書かないでください）。')
    $lines.Add('| No | 定性分析（原因） | 対策 |')
    $lines.Add('| --- | --- | --- |')
    $lines.Add('- 定性分析は、実績・増減・要対応・ほかの指標との関係から考えられる原因を書いてください。数字から断定できないことは「～の可能性がある（～で確かめる）」と書いてください。')
    $lines.Add('- 対策は、誰が・何を・いつまでにするかが分かる具体的な行動にしてください（例: 「設計リーダーが SCR002 の画面設計書を再レビューする（今週中）」）。')
    if ($previousNotes.Count -gt 0) { $lines.Add('- 前回の対策がある指標は、前回から今回への実績の変化で、対策が効いたかを定性分析に書いてください。効いていなければ、対策を変えてください。') }
    $lines.Add('- 1 つのセルは 200 文字以内で、セルの中で改行しないでください。')
    return ($lines -join "`r`n")
}

# ---------------------------------------------------------------- 回答を読む


function ConvertFrom-AnalysisAnswer {
    #     回答の表（Markdown の | 区切り、または Copilot の表をコピーしたタブ区切り）から、No → 定性分析・対策
    param([string] $Text)
    $result = @{}
    foreach ($line in ($Text -split '\r?\n')) {
        $t = $line.Trim()
        if (-not $t) { continue }
        if ($t.Contains("`t")) { $cells = @($t -split "`t") }
        elseif ($t.StartsWith('|')) { $cells = @($t.Trim('|') -split '\|') }
        else { continue }
        $cells = @($cells | ForEach-Object { $_.Trim() })
        if ($cells.Count -lt 3) { continue }
        $no = $cells[0].Normalize($nfkc) -replace '^No\.?\s*', '' -replace '[.．。]$', ''
        if ($no -notmatch '^\d+$') { continue }
        $result[[int]$no] = [pscustomobject]@{ Cause = ($cells[1] -replace '<br\s*/?>', ' ').Trim(); Measure = (($cells[2..($cells.Count - 1)] -join ' ') -replace '<br\s*/?>', ' ').Trim() }
    }
    return $result
}

function Read-AskedMetrics {
    #     保存した依頼の文の「基準から外れた指標」の表から、No と区分・指標（ConvertTo-Cell の形）を読む
    param([string] $Text)
    $list = New-Object System.Collections.Generic.List[object]
    $section = ''
    foreach ($line in ($Text -split '\r?\n')) {
        if ($line.StartsWith('# ')) { $section = $line.Trim(); continue }
        if ($section -ne '# 基準から外れた指標') { continue }
        $cells = @($line.Trim().Trim('|').Split('|') | ForEach-Object { $_.Trim() })
        if ($cells.Count -lt 3 -or $cells[0] -notmatch '^\d+$') { continue }
        $list.Add([pscustomobject]@{ No = [int]$cells[0]; Area = $cells[1]; Name = $cells[2] })
    }
    return , $list
}

# ---------------------------------------------------------------- 実行


# -ImportOnly: 回答の No は、依頼を作ったときの番号。今のブックから数え直すと、-Force の有無や、そのあとの記入で番号がずれるので、
# 保存した依頼の表の No・区分・指標で今のブックの指標に当てる（記入済みの指標は、-Force のときだけ書き直す）
$skipped = New-Object System.Collections.Generic.List[object]
$lost = New-Object System.Collections.Generic.List[object]
if ($ImportOnly -and (Test-Path -LiteralPath $promptPath)) {
    $asked = Read-AskedMetrics ([System.IO.File]::ReadAllText($promptPath))
    if ($asked.Count -gt 0) {
        $byKey = @{}
        foreach ($m in $metrics) {
            $key = '{0}|{1}' -f (ConvertTo-Cell $m.Area), (ConvertTo-Cell $m.Name)
            if (-not $byKey.ContainsKey($key)) { $byKey[$key] = New-Object System.Collections.Generic.Queue[object] }
            $byKey[$key].Enqueue($m)
        }
        $targets = New-Object System.Collections.Generic.List[object]
        foreach ($a in $asked) {
            $key = '{0}|{1}' -f $a.Area, $a.Name
            if (-not $byKey.ContainsKey($key) -or $byKey[$key].Count -eq 0) { $lost.Add($a); continue }
            $m = $byKey[$key].Dequeue()
            $m.No = $a.No
            if (-not $Force -and ($m.Cause -or $m.Measure)) { $skipped.Add($m); continue }
            $targets.Add($m)
        }
        $targets = $targets.ToArray()
        $filled = $skipped.Count
    }
}
if ($lost.Count -gt 0) { Write-Warning ('依頼の指標が、品質評価のシートにありません（ブックを作り直したときは、依頼も作り直してください）: {0}' -f ((@($lost | ForEach-Object { '{0}. {1}' -f $_.No, $_.Name })) -join '・')) }

if ($targets.Count -eq 0) {
    Write-Host ''
    if ($filled -gt 0) { Write-Host ('判定が × か △ の指標 {0} 件は、定性分析・対策が記入済みです（書き直すときは -Force）: {1}' -f $filled, $Path) }
    else { Write-Host ('判定が × か △ の指標がありません: {0}' -f $Path) }
    return
}

if (-not $ImportOnly) {
    $prompt = New-AnalysisPrompt
    [System.IO.File]::WriteAllText($promptPath, $prompt, $utf8Bom)
}
if ($PromptOnly) {
    Write-Host ''
    Write-Host ('判定が × か △ の指標 {0} 件の定性分析・対策の依頼を作りました: {1}' -f $targets.Count, $promptPath)
    if ($previousNotes.Count -gt 0) { Write-Host ('前回の品質評価の定性分析・対策 {0} 件も入れました: {1}' -f $previousNotes.Count, $Previous) }
    Write-Host '次の手順:'
    Write-Host '  1. 依頼の内容を Copilot に貼り付けて送信する'
    Write-Host ('  2. 回答をコピーし、{0} に UTF-8 のテキストとして保存する' -f $answerPath)
    Write-Host '  3. -ImportOnly を付けて再実行し、品質評価のブックに書き込む'
    return
}

if ($ImportOnly) {
    if (-not (Test-Path -LiteralPath $answerPath)) { throw "回答のファイルがありません。先に -PromptOnly で依頼を作り、Copilot の回答を保存してください: $answerPath" }
    $answers = ConvertFrom-AnalysisAnswer ([System.IO.File]::ReadAllText($answerPath))
    if ($answers.Count -eq 0) { throw "回答から「No | 定性分析（原因） | 対策」の行を読み取れませんでした: $answerPath" }
}
else {
    if (-not (Get-Command Set-Clipboard -ErrorAction SilentlyContinue)) {
        throw 'この環境ではクリップボードが使えません。-PromptOnly で依頼ファイルを作り、回答を保存してから -ImportOnly を実行してください。'
    }
    Write-Host ''
    Write-Host '手順: Copilot に貼り付けて送信 → 回答をコピー（回答の下のコピー ボタン）→ ここで Enter（q = 中断）'
    while ($true) {
        Set-Clipboard -Value $prompt
        Write-Host ('判定が × か △ の指標 {0} 件の依頼（{1} 文字）→ クリップボードにコピーしました' -f $targets.Count, $prompt.Length) -ForegroundColor Cyan
        $key = Read-Host '回答をコピーしたら Enter'
        if ($key -eq 'q') { Write-Host '中断しました。'; return }
        $answer = Get-Clipboard -Raw
        if (-not $answer -or $answer -eq $prompt) {
            Write-Warning 'クリップボードの内容が依頼のままです。Copilot の回答をコピーしてから Enter を押してください。'
            continue
        }
        $answers = ConvertFrom-AnalysisAnswer $answer
        if ($answers.Count -eq 0) {
            Write-Warning '回答から「No | 定性分析（原因） | 対策」の行を読み取れませんでした。回答全体をコピーし直すか、Copilot に「回答形式どおりに出し直して」と依頼してください。'
            continue
        }
        [System.IO.File]::WriteAllText($answerPath, $answer, $utf8Bom)
        break
    }
}

# 品質評価のシートの定性分析（原因）・対策の列に書く（ほかのセル・シートはそのまま）
$written = New-Object System.Collections.Generic.List[object]
$pkg = Open-ZipPackage $Path
$changes = @{}
try {
    $parts = Get-WorkbookSheetParts $pkg
    $doc = Read-PartXmlForEdit $pkg $parts.Sheets['品質評価']
    foreach ($m in $targets) {
        $a = $answers[$m.No]
        if (-not $a) { continue }
        $rowEl = Get-SheetRowElement $doc $m.Row -Create
        foreach ($pair in @(@($m.CauseColumn, $a.Cause), @($m.MeasureColumn, $a.Measure))) {
            $cell = Get-SheetCellElement $doc $rowEl $m.Row $pair[0] -Create
            Set-SheetCellValue $doc $cell $pair[1]
        }
        $written.Add($m)
    }
    $changes[$parts.Sheets['品質評価'].ToLowerInvariant()] = ConvertTo-XmlString $doc
}
finally { $pkg.Zip.Dispose() }
try { Save-ZipPackage -FilePath $Path -Changes $changes }
catch { throw ("品質評価のブックに書き込めませんでした（Excel で開いていませんか）: {0} : {1}" -f $Path, $_.Exception.Message) }

$written | ForEach-Object { [pscustomobject]@{ No = $_.No; Area = $_.Area; Name = $_.Name; Judgement = $_.Judgement; Cause = $answers[$_.No].Cause; Measure = $answers[$_.No].Measure } }

Write-Host ''
Write-Host ('定性分析・対策の下書きを {0} 件書き込みました（判定が × か △ の指標 {1} 件）: {2}' -f $written.Count, $targets.Count, $Path)
$missing = @($targets | Where-Object { -not $answers[$_.No] })
if ($missing.Count -gt 0) { Write-Warning ('回答に無い指標があります: {0}' -f ((@($missing | ForEach-Object { '{0}. {1}' -f $_.No, $_.Name })) -join '・')) }
if ($skipped.Count -gt 0) { Write-Warning ('記入済みの指標には書き込んでいません（書き直すときは -Force）: {0}' -f ((@($skipped | ForEach-Object { '{0}. {1}' -f $_.No, $_.Name })) -join '・')) }
Write-Host 'AI の下書きです。会議の前に、事実と合っているかを確かめて書き直してください。'

# .SYNOPSIS
#     品質評価のブック（New-QualityReport.ps1）の、判定が × か △ の指標について、定性分析（原因）と対策の下書きを AI（Microsoft 365 Copilot など）に書かせ、ブックに書き込みます（Excel 不要）。
#
# .DESCRIPTION
#     「品質評価」のシートの × と △ の指標（実績・増減・判定の考え方）と、「要対応」のシート、ほかの指標を 1 つの依頼の文にまとめます。
#     既定では依頼をクリップボードにコピーし、Copilot の回答（No | 定性分析（原因） | 対策 の表）をコピーして Enter を押すと、
#     品質評価のシートの定性分析（原因）・対策の列に書き込みます。ほかのセル・シート（リンク・グラフ）はそのままです。
#
#     依頼と回答は、ブックと同じフォルダに「<ブック名>_定性分析の依頼.txt」「<ブック名>_定性分析の回答.txt」として保存します。
#     記入済みの行は書き換えません（-Force で書き直します）。
#     前回の品質評価のブック（同じフォルダの、名前の日付が前のもの）に定性分析・対策があれば、今回の実績と並べて依頼に入れます。
#
# .PARAMETER Path
#     品質評価のブック（品質評価_日付.xlsx）。フォルダを指定すると、その中で「品質評価」のシートがある一番新しい xlsx を使います。
#
# .PARAMETER PromptOnly
#     依頼のファイルを作るだけにします（クリップボードが使えないとき）。回答を「_定性分析の回答.txt」に保存してから -ImportOnly を実行します。
#
# .PARAMETER ImportOnly
#     保存した回答のファイルを読んで、ブックに書き込みます。回答の No は、保存した依頼の No の指標に当てます
#     （記入済みの指標に書き込むのは、-Force を付けたときだけです）。
#
# .PARAMETER Force
#     定性分析・対策が記入済みの指標も、依頼に含めて書き直します。
#
# .PARAMETER Previous
#     前回の品質評価のブック。省略時は、同じフォルダで名前の日付（品質評価_20261008.xlsx など）が前のもののうち一番近いもの。
#     前回の定性分析・対策と、前回と今回の実績・判定を依頼に入れ、対策が効いたかも書いてもらいます。
#
# .PARAMETER NoExcel
#     Excel を起動しません（品質評価のブックは .xlsx なので、通常は不要です）。
#
# .EXAMPLE
#     .\Invoke-QualityAnalysis.ps1 D:\品質評価\品質評価_20261022.xlsx
#
# .EXAMPLE
#     # クリップボードが使えないとき
#     .\Invoke-QualityAnalysis.ps1 D:\品質評価 -PromptOnly
#     .\Invoke-QualityAnalysis.ps1 D:\品質評価 -ImportOnly
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     AI の下書きです。会議の前に、事実と合っているかを確かめて書き直してください。
