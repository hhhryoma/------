[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateSet('Katakana', 'Abbreviation', 'Word')]
    [string[]] $Kind = @('Katakana', 'Abbreviation', 'Word'),

    [ValidateRange(1, 1000)]
    [int] $MinCount = 3,

    [string[]] $Exclude,

    [switch] $IncludeWord,

    [string] $OutXlsx,

    [string] $OutCsv,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

foreach ($name in @('OutXlsx', 'OutCsv')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'GlossaryDraft.Term' -DefaultDisplayPropertySet Term, Kind, Count, Files, Variants -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$kindNames = @{ Katakana = 'カタカナ語'; Abbreviation = '略語'; Word = '用語' }
# 用語にしない、どの設計書にも出てくる言葉
$stopWords = @('場合', '処理', '表示', '画面', '入力', '出力', '項目', '設定', '選択', '実行', '確認', '内容', '対応', '修正', '発生', '結果', '機能', '必要', '可能', '不可', '部分', '全体', '指定', '中断', '遷移', '生成', '以上', '以下', '未満', '以外', '一覧', '概要', '備考', '名称', '番号', '区分', '日付', '時刻', '件数', '文字', '数字', '半角', '全角', '桁数', '必須', '任意', '初期', '初期値', '最大', '最小', '新規', '追加', '更新', '削除', '登録', '検索', '参照', '作成', '変更', '改訂', '版数', '作成日', '更新日', '作成者', '更新者', '承認', '承認者', '対象', '対象外', '取得', '設定値', '有無', '不要', '同上')

function Get-KatakanaKey {
    #     カタカナ語の表記ゆれをまとめる形（全角・半角、語末の長音、中黒の違いを除く）
    param([string] $Text)
    return $Text.Normalize($nfkc).TrimEnd('ー').Replace('・', '')
}

# 用語にしない言葉（表記ゆれ・大文字小文字の違うものも除く）
$excludeSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($e in @($Exclude | ForEach-Object { ([string]$_) -split '[,、]' })) {
    $t = $e.Normalize($nfkc).Trim()
    if ($t) { [void]$excludeSet.Add($t); [void]$excludeSet.Add((Get-KatakanaKey $t)) }
}

$katakanaRx = New-Object System.Text.RegularExpressions.Regex('[ァ-ヴｦ-ﾟー・･ｰ]{2,}')
$abbreviationRx = New-Object System.Text.RegularExpressions.Regex('(?<![A-Za-zＡ-Ｚａ-ｚ0-9０-９_])[A-ZＡ-Ｚ][A-ZＡ-Ｚ0-9０-９]{1,5}(?![A-Za-zＡ-Ｚａ-ｚ0-9０-９_])')
$wordRx = New-Object System.Text.RegularExpressions.Regex('[一-龠々〆ヵヶ][一-龠々〆ヵヶァ-ヴー]+|[ァ-ヴー]+[一-龠々〆ヵヶ][一-龠々〆ヵヶァ-ヴー]*')

# ---------------------------------------------------------------- 設計書を読む


$searchArgs = @{ Pattern = '\S'; Path = $Path; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
if ($IncludeWord) { $searchArgs['IncludeWord'] = $true }
$hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })

# 表記（そのままの文字）→ @{ Kind; Key（表記ゆれをまとめる形）; Text; Count; Files; First（最初に出てきたセル）; FirstSeq }
$terms = [ordered]@{}
$hitSeq = 0
function Add-Term {
    param([string] $TermKind, [string] $Text, [string] $Key, $Hit, [int] $Seq)
    $id = '{0}|{1}' -f $TermKind, $Text
    if (-not $terms.Contains($id)) {
        $terms[$id] = [pscustomobject]@{ Kind = $TermKind; Key = $Key; Text = $Text; Count = 0; Files = (New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)); First = $Hit; FirstSeq = $Seq }
    }
    $terms[$id].Count++
    [void]$terms[$id].Files.Add($Hit.Path)
}

# カタカナの続き（中黒のあるものは、中黒の無い形も出てくるかで数え方を決めるので、あとで数える）
$katakanaRuns = New-Object System.Collections.Generic.List[object]
$plainKeys = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($h in $hits) {
    $hitSeq++
    $text = [string]$h.Text
    if (-not $text) { continue }
    if ($Kind -contains 'Katakana') {
        foreach ($m in $katakanaRx.Matches($text)) {
            $raw = $m.Value.Trim('・', '･')
            if (-not $raw) { continue }
            $key = Get-KatakanaKey $raw
            if ($raw -notmatch '[・･]') { [void]$plainKeys.Add($key) }
            $katakanaRuns.Add([pscustomobject]@{ Raw = $raw; Key = $key; Hit = $h; Seq = $hitSeq })
        }
    }
    if ($Kind -contains 'Abbreviation') {
        foreach ($m in $abbreviationRx.Matches($text)) {
            $norm = $m.Value.Normalize($nfkc)
            # 数字の入った ID（SCR001・E0001 など）は略語にしない
            if ($norm -match '\d') { continue }
            Add-Term 'Abbreviation' $m.Value $norm.ToUpperInvariant() $h $hitSeq
        }
    }
    if ($Kind -contains 'Word') {
        foreach ($m in $wordRx.Matches($text.Normalize($nfkc))) {
            $w = $m.Value
            if ($w.Length -gt 12 -or $stopWords -contains $w) { continue }
            Add-Term 'Word' $w $w $h $hitSeq
        }
    }
}
foreach ($run in $katakanaRuns) {
    # 「データ・ベース」は、「データベース」も出てくればその表記ゆれ。無ければ「ユーザー・グループ」のような並びなので 1 語ずつ数える
    $parts = @($run.Raw)
    if ($run.Raw -match '[・･]' -and -not $plainKeys.Contains($run.Key)) { $parts = @($run.Raw -split '[・･]' | Where-Object { $_ }) }
    foreach ($part in $parts) {
        # 3 文字以上（半角カタカナは全角にして数える）で、長音のほかに 2 文字以上
        $norm = $part.Normalize($nfkc)
        if ($norm.Length -lt 3 -or $norm.Replace('ー', '').Replace('・', '').Length -lt 2) { continue }
        # 表記ゆれ: 全角・半角、語末の長音、中黒の有無
        Add-Term 'Katakana' $part (Get-KatakanaKey $part) $run.Hit $run.Seq
    }
}

# ---------------------------------------------------------------- 用語にまとめる


$results = New-Object System.Collections.Generic.List[object]
$variantGroups = New-Object System.Collections.Generic.List[object]
foreach ($k in @('Katakana', 'Abbreviation', 'Word')) {
    if ($Kind -notcontains $k) { continue }
    $groups = @($terms.Values | Where-Object Kind -eq $k | Group-InOrder Key)
    foreach ($g in $groups) {
        $variants = New-Object System.Collections.Generic.List[object]
        foreach ($t in $g.Group) { $variants.Add($t) }
        # 推奨の表記: 多く使われている表記（同じ数なら長音のある方）。半角カタカナ・全角英数字は推奨にしない（Test-ExcelTerms.ps1 と同じ）
        $variants.Sort([System.Comparison[object]] {
                param($a, $b)
                $wa = [int]($a.Text -cne $a.Text.Normalize($nfkc))
                $wb = [int]($b.Text -cne $b.Text.Normalize($nfkc))
                if ($wa -ne $wb) { return $wa.CompareTo($wb) }
                if ($a.Count -ne $b.Count) { return $b.Count.CompareTo($a.Count) }
                $pa = [int]$a.Text.EndsWith('ー')
                $pb = [int]$b.Text.EndsWith('ー')
                if ($pa -ne $pb) { return $pb.CompareTo($pa) }
                return [string]::CompareOrdinal($a.Text, $b.Text)
            })
        $total = ($variants | Measure-Object Count -Sum).Sum
        if ($total -lt $MinCount) { continue }
        $best = $variants[0]
        $term = $best.Text
        $others = @($variants | Select-Object -Skip 1)
        if ($term -cne $term.Normalize($nfkc)) {
            # 半角カタカナ・全角英数字の表記しか無いときは、全角カタカナ・半角英数字にしたものを用語にする
            $term = $term.Normalize($nfkc)
            $others = $variants.ToArray()
        }
        if ($excludeSet.Contains($term) -or $excludeSet.Contains($g.Name)) { continue }
        $files = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($v in $variants) { foreach ($f in $v.Files) { [void]$files.Add($f) } }
        # 最初の場所: いちばん先に出てきた表記のセル
        $first = $variants[0].First; $firstSeq = $variants[0].FirstSeq
        foreach ($v in $variants) { if ($v.FirstSeq -lt $firstSeq) { $first = $v.First; $firstSeq = $v.FirstSeq } }
        $r = [pscustomobject]@{
            Term = $term; Kind = $kindNames[$k]; Count = [int]$total; Files = $files.Count
            Variants = (@($others | ForEach-Object { '{0}（{1}）' -f $_.Text, $_.Count }) -join '・')
            File = $first.File; Location = ('{0}!{1}' -f $first.Sheet, $first.Cell); Path = $first.Path; Sheet = $first.Sheet; Cell = $first.Cell
        }
        $r.PSObject.TypeNames.Insert(0, 'GlossaryDraft.Term')
        $results.Add($r)
        if ($others.Count -gt 0) { $variantGroups.Add([pscustomobject]@{ Term = $term; Kind = $kindNames[$k]; Others = @($others | ForEach-Object { $_.Text }); Detail = ((@($variants.ToArray() | ForEach-Object { '{0}（{1}）' -f $_.Text, $_.Count })) -join '・') }) }
    }
}
# 種類の順に、多く出てくる順（同じなら文字の順）
$sorted = New-Object System.Collections.Generic.List[object]
foreach ($r in $results) { $sorted.Add($r) }
$kindOrder = @{ 'カタカナ語' = 0; '略語' = 1; '用語' = 2 }
$sorted.Sort([System.Comparison[object]] {
        param($a, $b)
        if ($kindOrder[$a.Kind] -ne $kindOrder[$b.Kind]) { return $kindOrder[$a.Kind].CompareTo($kindOrder[$b.Kind]) }
        if ($a.Count -ne $b.Count) { return $b.Count.CompareTo($a.Count) }
        return [string]::CompareOrdinal($a.Term, $b.Term)
    })

# ---------------------------------------------------------------- 結果


$sorted

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $sorted) {
        $row = @{ Values = @($r.Term, '', $r.Kind, [double]$r.Count, [double]$r.Files, $r.Variants, ('{0} {1}' -f $r.File, $r.Location)) }
        if ($r.Path) {
            if ($r.Cell -match '^[A-Z]+\d+$') { $row.Link = @{ Column = 6; Path = $r.Path; Location = ("'{0}'!{1}" -f $r.Sheet.Replace("'", "''"), $r.Cell) } }
            else { $row.Link = @{ Column = 6; Path = $r.Path } }
        }
        $rows.Add($row)
    }
    $variantRows = New-Object System.Collections.Generic.List[object]
    foreach ($v in $variantGroups) { $variantRows.Add(@{ Values = @($v.Term, ($v.Others -join '/'), $v.Kind, $v.Detail) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '用語集'; Columns = @(@{ Header = '用語'; Width = 22 }, @{ Header = '意味'; Width = 50 }, @{ Header = '種類'; Width = 10 }, @{ Header = '出現数'; Width = 8 }, @{ Header = '設計書の数'; Width = 9 }, @{ Header = '表記ゆれ'; Width = 30 }, @{ Header = '最初の場所'; Width = 50 }); Rows = $rows },
            @{ Name = '表記ゆれ'; Columns = @(@{ Header = '正しい表記'; Width = 22 }, @{ Header = '揺れ'; Width = 30 }, @{ Header = '種類'; Width = 10 }, @{ Header = '出現数'; Width = 40 }); Rows = $variantRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}
if ($OutCsv) {
    # Test-ExcelTerms.ps1 の -Glossary の形（正しい表記・揺れ・備考）。表記ゆれのある用語だけ
    $dir = Split-Path -Parent $OutCsv
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $lines = @(@($variantGroups | ForEach-Object { [pscustomobject][ordered]@{ '正しい表記' = $_.Term; '揺れ' = ($_.Others -join '/'); '備考' = $_.Detail } }) | ConvertTo-Csv -NoTypeInformation)
    if ($lines.Count -eq 0) { $lines = @('"正しい表記","揺れ","備考"') }
    [System.IO.File]::WriteAllText($OutCsv, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
}

Write-Host ''
$byKind = @($sorted | Group-InOrder Kind | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
Write-Host ('{0} ファイルから用語 {1}（{2}）。{3} 回以上出てくるもの' -f @($hits | ForEach-Object { $_.Path } | Select-Object -Unique).Count, $sorted.Count, ($byKind -join '・'), $MinCount)
if ($variantGroups.Count -gt 0) {
    $shown = @($variantGroups | Select-Object -First 10 | ForEach-Object { '{0}／{1}' -f $_.Term, ($_.Others -join '／') })
    $more = ''
    if ($variantGroups.Count -gt 10) { $more = ' ほか {0} 組' -f ($variantGroups.Count - 10) }
    Write-Host ('表記ゆれ {0} 組: {1}{2}' -f $variantGroups.Count, ($shown -join '・'), $more)
}
if ($OutXlsx) { Write-Host "用語集の下書きを保存しました: $OutXlsx" }
if ($OutCsv) { Write-Host "表記ゆれの用語集（Test-ExcelTerms.ps1 の -Glossary の形）を保存しました: $OutCsv" }

# .SYNOPSIS
#     設計書からカタカナ語・略語・よく出る用語を集め、用語集の下書きと、表記ゆれの用語集（Test-ExcelTerms.ps1 の -Glossary）を作ります（Excel 不要）。
#
# .DESCRIPTION
#     設計書（Excel。-IncludeWord で Word も）のセルの文字から、次の言葉を集めて数えます。
#
#       カタカナ語: 3 文字以上のカタカナ（ユーザー・ステータスなど。半角カタカナは全角にして数える）。
#                   「ユーザー・グループ」のような中黒の並びは 1 語ずつ数える（「データベース」もある「データ・ベース」は 1 語）
#       略語: 英大文字だけの 2〜6 文字（ID・CSV・API など。全角も。数字の入った ID（SCR001 など）は除く）
#       用語: 漢字を含む、漢字とカタカナの続き（受注番号・顧客マスタなど。どこにでも出てくる言葉は除く）
#
#     全角・半角、語末の長音（ユーザ／ユーザー）、中黒（データ・ベース／データベース）だけが違う表記は、
#     1 つの用語の「表記ゆれ」としてまとめ、多く使われている表記を用語にします（同じ数なら長音のある方）。
#     半角カタカナ・全角英数字は用語にしません（それしか無ければ、全角カタカナ・半角英数字にしたものを用語にします）。
#     -MinCount 回以上出てくるものを、種類ごとに多い順に出力します。意味の列は空欄なので、人が書いてください。
#     -OutCsv の表記ゆれの用語集は、Test-ExcelTerms.ps1 の -Glossary に渡して、表記ゆれの置換リストを作れます。
#
# .PARAMETER Path
#     設計書のフォルダまたはファイル（複数可）。
#
# .PARAMETER Kind
#     集める言葉の種類（Katakana: カタカナ語 / Abbreviation: 略語 / Word: 用語）。既定はすべて。
#
# .PARAMETER MinCount
#     用語にする出現数の下限（表記ゆれを合わせた数。既定は 3）。
#
# .PARAMETER Exclude
#     用語にしない言葉（複数可。カンマ区切りでも書ける）。表記ゆれ（全角・半角、語末の長音、中黒）と英字の大文字小文字の違うものも除きます。
#
# .PARAMETER IncludeWord
#     Word の文書（.docx / .docm）も読みます。
#
# .PARAMETER OutXlsx
#     用語集の下書き（用語・意味（空欄）・種類・出現数・設計書の数・表記ゆれ・最初の場所）と、表記ゆれの一覧を xlsx に保存します。
#
# .PARAMETER OutCsv
#     表記ゆれのある用語を、Test-ExcelTerms.ps1 の -Glossary の形（正しい表記・揺れ・備考。UTF-8・BOM 付き）の CSV に保存します。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\New-GlossaryDraft.ps1 D:\設計書 -IncludeWord -OutXlsx D:\用語集の下書き.xlsx -OutCsv D:\用語集.csv
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     言葉は文字の種類（カタカナ・漢字・英字）の続きで区切るだけなので、「受注番号入力」のように長い続きは 1 つの言葉になります。
