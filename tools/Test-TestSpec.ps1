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
# 結果の欄に書かれる、未実施を表す言葉（これ以外の読めない値を知らせる）
$notRunPattern = '^(未実施|未|未着手|実施前|予定|-)$'

$fields = @('機能', 'テストID', 'テスト内容', '期待結果', '手順', '前提条件', '入力値', '区分', '結果', 'バグ番号', '予定日', '実施日', '工程')
$candidates = Read-HeaderCandidates $Headers $fields
$vagueRx = $null
if ($VagueWords) { $vagueRx = New-Object System.Text.RegularExpressions.Regex((@($VagueWords | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.Normalize($nfkc)) }) -join '|')) }

# ---------------------------------------------------------------- テスト仕様書を読む


$records = Read-HeaderRecords -Inputs (Get-InputFiles $Path 'テスト仕様書' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
    -Need @(@('期待結果', 'テスト内容'), @('テストID', '手順', '結果')) -Key @('テストID', '期待結果', 'テスト内容', '手順') -What 'テスト仕様書' `
    -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel
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
        $m = $vagueRx.Match($expected)
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

# ---------------------------------------------------------------- 境界値（-DesignPath の画面項目の桁数）


$boundaryItems = New-Object System.Collections.Generic.List[object]
if ($DesignPath) {
    $searchArgs = @{ Pattern = '\S'; Path = $DesignPath; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $designBooks = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
    foreach ($book in $designBooks.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            $screen = Find-LabelValue $grid '画面ID'
            if (-not $screen) { continue }
            foreach ($t in (Find-DesignTables $grid @('項目名', '桁数') '項目名')) {
                foreach ($dr in $t.Rows) {
                    $digits = ([string]$dr.Cells[$t.ColumnOf['桁数']]).Normalize($nfkc).Trim()
                    if ($digits -notmatch '^\d+$') { continue }
                    # 出力だけの項目（入出力が O）は入力の境界値が無い
                    if ($t.ColumnOf.ContainsKey('入出力') -and ([string]$dr.Cells[$t.ColumnOf['入出力']]).Normalize($nfkc).Trim() -notmatch 'I') { continue }
                    $boundaryItems.Add([pscustomobject]@{ Screen = $screen.Trim(); Item = ([string]$dr.Cells[$t.ColumnOf['項目名']]).Trim(); Digits = [int]$digits; File = $book.File; Sheet = $sheetName; Row = $dr.Row })
                }
            }
        }
    }
    foreach ($item in $boundaryItems) {
        $screenRx = New-Object System.Text.RegularExpressions.Regex(('(?<![A-Za-z0-9]){0}(?![A-Za-z0-9])' -f [regex]::Escape($item.Screen)), 'IgnoreCase')
        $screenCases = @($cases | Where-Object { $screenRx.IsMatch(('{0} {1} {2}' -f $_.Function, $_.Place.File, $_.Place.Sheet).Normalize($nfkc)) })
        if ($screenCases.Count -eq 0) { continue }
        $itemLabel = ConvertTo-Label $item.Item
        $covered = @($screenCases | Where-Object {
                $text = (@($_.Values.Values) -join ' ').Normalize($nfkc)
                (ConvertTo-Label $text).Contains($itemLabel) -and ($text -match $boundaryPattern -or $text -match ('(?<!\d){0}(?!\d)' -f $item.Digits) -or $text -match ('(?<!\d){0}(?!\d)' -f ($item.Digits + 1)))
            })
        if ($covered.Count -eq 0) {
            Add-Issue '警告' '境界値のテストが無い' $screenCases[0] '' ("{0} の入力項目「{1}」（{2} 桁。{3} {4} {5} 行目）の桁数の境界（{2} 桁・{6} 桁）を確かめるケースが見つかりません" -f $item.Screen, $item.Item, $item.Digits, $item.File, $item.Sheet, $item.Row, ($item.Digits + 1))
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
            @{ Name = '機能別'; Columns = @(@{ Header = '機能'; Width = 24 }, @{ Header = 'ケース'; Width = 7 }, @{ Header = '異常系'; Width = 7 }, @{ Header = '異常系の割合'; Width = 9 }, @{ Header = '異常系の数え方'; Width = 20 }, @{ Header = 'エラー'; Width = 7 }, @{ Header = '警告'; Width = 7 }); Rows = $functionRows })
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
if ($DesignPath) { Write-Host ('境界値: 桁数のある入力項目 {0} 件を確かめました' -f $boundaryItems.Count) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テスト仕様書（テストケースの一覧）の書き方の抜け・あいまいさ・重複と、異常系・境界値のケースの不足を見つけます（Excel 不要）。
#
# .DESCRIPTION
#     テスト仕様書（Excel・CSV）から、期待結果・手順・前提条件・入力値・区分・結果などの列を見出しの文字で探して読み、次を確かめます。
#
#       期待結果が空・あいまいな期待結果（「正しく」「適切に」「仕様どおり」など）・手順が空
#       テストIDの重複・同じ内容のケース（前提条件・手順・入力値・期待結果がすべて同じ）・結果の値が読めない
#       異常系が少ない機能（区分の列、無ければ期待結果の言葉から推定。-MinErrorRatio 未満）
#       境界値のテストが無い入力項目（-DesignPath の画面項目定義の桁数。項目名と、桁・文字・最大などの言葉か桁数の数字が出てくるケースを探す）
#
#     見出しの候補は tools\test-quality\見出し.csv（Measure-TestQuality.ps1 と同じ）で足せます。
#
# .PARAMETER Path
#     テスト仕様書のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER DesignPath
#     設計書のフォルダ・ファイル。画面項目定義の入力項目の桁数から、境界値のケースがあるかを確かめます。
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
