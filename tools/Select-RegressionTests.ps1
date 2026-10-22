[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Old,

    [Parameter(Mandatory = $true)]
    [string] $New,

    [Parameter(Mandatory = $true)]
    [string[]] $TestPath,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
foreach ($name in @('Old', 'New', 'Headers', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'RegressionTest.Case' -DefaultDisplayPropertySet Priority, Function, Id, Reason -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 変更した行の「名前」とみなす列（項目名・イベント名・メッセージID など）。テストケースにこの名前が出てくれば、そのケースを選ぶ
$nameLabels = @('項目名', '項目ID', 'イベント名', 'イベントID', 'イベント', 'メッセージID', 'コード区分名', '表示名', '論理名', '物理名', '画面ID', '画面名', 'テーブル名')
# 設計書の文書の情報（更新日など）の変更は、テストに関係しない
$ignoredSheets = '^(表紙|改訂履歴|変更履歴|更新履歴)$'
$ignoredLabels = '^(作成者|作成日|更新者|更新日|版数|最終更新日|文書番号)$'

# ---------------------------------------------------------------- 設計書の差分


Write-Host '設計書の差分を取っています…'
$compareArgs = @{ Old = $Old; New = $New }
if ($NoRecurse) { $compareArgs['NoRecurse'] = $true }
if ($NoExcel) { $compareArgs['NoExcel'] = $true }
$diffs = @(& (Join-Path $PSScriptRoot 'Compare-Excel.ps1') @compareArgs 6>$null)

function Read-Books {
    #     設計書のフォルダ・ファイルを、パス → シート → 行・列 の形で読む
    param([string] $Target)
    $searchArgs = @{ Pattern = '\S'; Path = [System.Management.Automation.WildcardPattern]::Escape($Target); Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    return ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
}
$newBooks = Read-Books $New
$oldBooks = Read-Books $Old

function Get-BookScreen {
    #     ブックの画面 ID（画面設計書の「画面ID」の欄）。無ければ空
    param($Book)
    if (-not $Book) { return '' }
    foreach ($grid in $Book.Sheets.Values) {
        $v = Find-LabelValue $grid '画面ID'
        if ($v -and $v -match '^[A-Za-z]+\d+$') { return $v.Trim() }
    }
    return ''
}

function Get-ChangeTarget {
    #     変更したセルの行が、表（名前の列のある見出しの下）の行なら、その行の名前と列の見出しを返す。
    #     表の外なら、左にある見出しのセルの文字を返す
    param($Grid, [int] $Row, [int] $Col)
    $result = [pscustomobject]@{ Names = @(); Attribute = '' }
    if (-not $Grid -or -not $Grid.Cells.ContainsKey($Row)) { return $result }
    # 一番近い上の見出しの行（名前の列のどれかがあり、3 列以上ある行）
    $header = $null
    foreach ($r in $Grid.Rows) {
        if ($r -ge $Row) { break }
        $labels = @($Grid.Labels[$r].Values)
        if ($labels.Count -ge 3 -and @($nameLabels | Where-Object { $labels -contains $_ }).Count -gt 0) { $header = $r }
        # 「■」の節の始まりで表は終わる
        if (@($Grid.Cells[$r].Values | Where-Object { $_.TrimStart().StartsWith('■') }).Count -gt 0) { $header = $null }
    }
    if ($null -ne $header) {
        $labelOf = $Grid.Labels[$header]
        $names = @(foreach ($label in $nameLabels) {
                foreach ($c in $labelOf.Keys) { if ($labelOf[$c] -eq $label -and $Grid.Cells[$Row].ContainsKey($c)) { $Grid.Cells[$Row][$c].Trim() } }
            })
        $attribute = ''
        if ($labelOf.ContainsKey($Col)) { $attribute = $Grid.Cells[$header][$Col].Trim() }
        $result.Names = @($names | Where-Object { $_ } | Select-Object -Unique)
        $result.Attribute = $attribute
        return $result
    }
    # 表の外: 同じ行の左にある見出し
    $left = @($Grid.Cols[$Row] | Where-Object { $_ -lt $Col })
    if ($left.Count -gt 0) { $result.Attribute = $Grid.Cells[$Row][$left[-1]].Trim() }
    return $result
}

$changes = New-Object System.Collections.Generic.List[object]
foreach ($d in $diffs) {
    if (@('値', '数式', '行', 'シート', '連番') -notcontains $d.Kind) { continue }
    # 削除したセル・行は旧のブック、それ以外は新のブック
    $deleted = $d.Change -eq '削除'
    $sheet = $(if ($deleted) { $d.OldSheet } else { $d.Sheet })
    $location = $(if ($deleted) { $d.OldLocation } else { $d.Location })
    if (-not $sheet) { $sheet = $d.Sheet }
    if ($sheet -match $ignoredSheets) { continue }
    $book = $(if ($deleted) { $oldBooks[$d.OldPath] } else { $newBooks[$d.Path] })
    if (-not $book) { $book = $(if ($d.Path) { $newBooks[$d.Path] } else { $oldBooks[$d.OldPath] }) }
    $grid = $null
    if ($book -and $sheet -and $book.Sheets.Contains($sheet)) { $grid = $book.Sheets[$sheet] }
    $target = [pscustomobject]@{ Names = @(); Attribute = '' }
    $m = [regex]::Match([string]$location, '([A-Z]+)(\d+)$')
    if ($grid -and $m.Success) { $target = Get-ChangeTarget $grid ([int]$m.Groups[2].Value) (Get-ColumnIndex $m.Groups[1].Value) }
    if ($target.Attribute -match $ignoredLabels) { continue }
    $screen = Get-BookScreen $(if ($newBooks[$d.Path]) { $newBooks[$d.Path] } else { $oldBooks[$d.OldPath] })
    $what = $(if ($target.Names.Count -gt 0) { '「{0}」' -f ($target.Names -join '・') } else { '' })
    if ($target.Attribute -and $d.Kind -ne '行') { $what += 'の' + $target.Attribute }
    $changes.Add([pscustomobject]@{
            File = $d.File; Screen = $screen; Sheet = $sheet; Location = $location; Kind = $d.Kind; Change = $d.Change
            Names = $target.Names; Attribute = $target.Attribute; What = $what.Trim(); Old = $d.Old; New = $d.New; Cases = 0; Seq = $changes.Count
        })
}

# ---------------------------------------------------------------- テスト仕様書


$fields = @('機能', 'テストID', 'テスト内容', '期待結果', '手順', '前提条件', '入力値', '区分', '結果', 'バグ番号', '予定日', '実施日', '工程')
$candidates = Read-HeaderCandidates $Headers $fields
$readLog = New-Object System.Collections.Generic.List[object]
$records = Read-HeaderRecords -Inputs (Get-InputFiles $TestPath 'テスト仕様書' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
    -Need @(@('期待結果', 'テスト内容'), @('テストID', '手順', '結果')) -Key @('テストID', '期待結果', 'テスト内容', '手順') -What 'テスト仕様書' `
    -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
$allCases = ConvertTo-TestCases $records
$cases = @($allCases | Where-Object { $_.Kind -ne '対象外' })
if ($cases.Count -eq 0) { Write-Warning 'テストケースが見つかりません（期待結果か確認内容と、テストID・手順・結果のどれかの見出しがある表を探します）' }
foreach ($c in $cases) {
    $c | Add-Member -NotePropertyName 'Text' -NotePropertyValue (ConvertTo-Label (@($c.Values.Values | Select-Object -Unique) -join ' ').Normalize($nfkc))
    $c | Add-Member -NotePropertyName 'Where' -NotePropertyValue ('{0} {1} {2}' -f $c.Function, $c.Place.File, $c.Place.Sheet).Normalize($nfkc)
}

# ---------------------------------------------------------------- 選ぶ


# ケース → 選んだ理由（高: 変更した項目などの名前がケースに出てくる / 中: 変更した画面のケース）
$selected = [ordered]@{}
function Add-Selection {
    param($Case, [string] $Priority, [string] $Reason)
    $key = '{0}|{1}|{2}' -f $Case.Place.Path, $Case.Place.Sheet, $Case.Place.Row
    if (-not $selected.Contains($key)) { $selected[$key] = [pscustomobject]@{ Case = $Case; Priority = $Priority; Reasons = New-Object System.Collections.Generic.List[string] } }
    $s = $selected[$key]
    if ($Priority -eq '高') { $s.Priority = '高' }
    # 理由ごとに、名前が出てくる（高）か同じ画面（中）かを書く
    $text = '[{0}] {1}' -f $Priority, $Reason
    if (-not $s.Reasons.Contains($text)) { $s.Reasons.Add($text) }
}

foreach ($ch in $changes) {
    $reason = '{0} {1}（{2}）{3}を{4}' -f $ch.File, $ch.Location, $ch.Kind, $ch.What, $ch.Change
    $reason = ($reason -replace '\s+', ' ').Trim()
    $inScreen = $cases
    if ($ch.Screen) {
        $rx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9]){0}(?![A-Za-z0-9])' -f [regex]::Escape($ch.Screen)), 'IgnoreCase')
        $inScreen = @($cases | Where-Object { $rx.IsMatch($_.Where) })
    }
    $labels = @($ch.Names | ForEach-Object { ConvertTo-Label $_.Normalize($nfkc) } | Where-Object { $_.Length -ge 2 })
    $hits = 0
    foreach ($c in $inScreen) {
        $named = @($labels | Where-Object { $c.Text.Contains($_) }).Count -gt 0
        if ($named) { Add-Selection $c '高' $reason; $hits++ }
        elseif ($ch.Screen) { Add-Selection $c '中' $reason; $hits++ }
    }
    $ch.Cases = $hits
}

$result = @($selected.Values | Sort-Object @{ e = { if ($_.Priority -eq '高') { 0 } else { 1 } } }, @{ e = { $_.Case.Place.File } }, @{ e = { $_.Case.Place.Sheet } }, @{ e = { $_.Case.Place.Row } } | ForEach-Object {
        $p = $_.Case.Place
        $o = [pscustomobject]@{
            Priority = $_.Priority; Function = $_.Case.Function; Id = $_.Case.Id; Reason = ($_.Reasons -join ' ／ '); File = $p.File
            Location = $(if ($p.Sheet) { '{0}!{1}{2}' -f $p.Sheet, (ConvertTo-ColumnName $p.Column), $p.Row } else { '{0} 行目' -f $p.Row })
            Expected = ([string]$_.Case.Values['期待結果']).Trim(); Path = $p.Path; Sheet = $p.Sheet; Row = $p.Row; Col = $p.Column
        }
        $o.PSObject.TypeNames.Insert(0, 'RegressionTest.Case')
        $o
    })
$result

# ---------------------------------------------------------------- 結果


$untested = @($changes | Where-Object { $_.Cases -eq 0 })
if ($OutXlsx) {
    $caseRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $result) {
        $row = @{ Values = @($r.Priority, $r.Function, $r.Id, $r.Expected, $r.Reason, $r.File, $r.Location) }
        if ($r.Sheet) { $row.Link = @{ Column = 7; Path = $r.Path; Location = ("'{0}'!{1}{2}" -f $r.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $r.Col), $r.Row) } }
        $caseRows.Add($row)
    }
    $changeRows = New-Object System.Collections.Generic.List[object]
    foreach ($ch in $changes) { $changeRows.Add(@{ Values = @($ch.File, $ch.Screen, $ch.Location, $ch.Kind, $ch.Change, ($ch.Names -join '・'), $ch.Attribute, [string]$ch.Old, [string]$ch.New, [double]$ch.Cases) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '再テストするケース'; Columns = @(@{ Header = '優先'; Width = 6 }, @{ Header = '機能'; Width = 14 }, @{ Header = 'テストID'; Width = 10 }, @{ Header = '期待結果'; Width = 40 }, @{ Header = '選んだ理由（設計書の変更）'; Width = 70 }, @{ Header = 'テスト仕様書'; Width = 24 }, @{ Header = '場所'; Width = 16 }); Rows = $caseRows },
            @{ Name = '設計書の変更'; Columns = @(@{ Header = '設計書'; Width = 34 }, @{ Header = '画面'; Width = 8 }, @{ Header = '場所'; Width = 20 }, @{ Header = '種類'; Width = 6 }, @{ Header = '変更'; Width = 6 }, @{ Header = '名前'; Width = 20 }, @{ Header = '列'; Width = 12 }, @{ Header = '旧'; Width = 30 }, @{ Header = '新'; Width = 30 }, @{ Header = '選んだケース'; Width = 9 }); Rows = $changeRows }, (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ('設計書の変更 {0} 件（差分 {1} 件。表紙・改訂履歴・作成日などは除く）: 再テストするケース {2} 件（高 {3}・中 {4}）/ テストケース {5} 件' -f $changes.Count, $diffs.Count, $result.Count, @($result | Where-Object Priority -eq '高').Count, @($result | Where-Object Priority -eq '中').Count, $cases.Count)
if ($untested.Count -gt 0) {
    Write-Host ('テストが見つからない変更 {0} 件: {1}' -f $untested.Count, ((@($untested | Select-Object -First 5 | ForEach-Object { '{0} {1}{2}' -f $_.File, $_.Location, $(if ($_.What) { ' ' + $_.What } else { '' }) })) -join '・'))
}
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書の旧版と新版の差分から、再テストするテストケースを選びます（Excel 不要）。
#
# .DESCRIPTION
#     Compare-Excel.ps1 で設計書の差分を取り、変更したセルが表のどの行・列か（画面項目定義の「支払方法」の「桁数」など）を調べて、
#     テスト仕様書のケースを選びます。改修案件の回帰テストの範囲を決めるときに使います。
#
#       高  変更した行の名前（項目名・イベント名・メッセージID・コード区分名・表示名・論理名など）がケースに出てくる
#       中  画面設計書の変更で、その画面のケース（機能・ファイル名・シート名に画面 ID が出てくる）
#
#     表紙・改訂履歴のシートと、作成者・更新日・版数などの欄の変更は除きます。
#     テストが見つからない変更（テストを足すか、影響が無いことを確かめるもの）も知らせます。
#
# .PARAMETER Old
#     旧版の設計書のフォルダまたはファイル。
#
# .PARAMETER New
#     新版の設計書のフォルダまたはファイル。
#
# .PARAMETER TestPath
#     テスト仕様書のファイル・フォルダ（Excel・CSV）。
#
# .PARAMETER Headers
#     テスト仕様書の見出しの候補の CSV。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（再テストするケース（テスト仕様書の該当セルへのリンク付き）・設計書の変更・読み取り結果）。
#
# .EXAMPLE
#     .\Select-RegressionTests.ps1 -Old D:\設計書_v1.0 -New D:\設計書_v1.1 -TestPath D:\テスト仕様書 -OutXlsx D:\再テスト.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Compare-Excel.ps1 が必要です。
#     選び方は機械的な目安です。共通部品・テーブルの変更などの影響は、最終的には人が確かめてください。
