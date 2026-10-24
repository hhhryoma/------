[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string] $Glossary,

    [switch] $NoAuto,

    [ValidateSet('Value', 'Comment', 'Shape', 'SheetName')]
    [string[]] $Target = @('Value', 'Comment', 'Shape'),

    [switch] $NoRecurse,

    [switch] $IncludeWord,

    [switch] $NoExcel,

    [string] $OutCsv,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

foreach ($name in @('OutCsv', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($NoAuto -and -not $Glossary) { throw '-NoAuto を付ける場合は、用語集（-Glossary）を指定してください。' }

Update-TypeData -TypeName 'ExcelTerms.Variant' -DefaultDisplayPropertySet Preferred, Variant, Count, PreferredCount, Kind -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
# 自動で見つける語: カタカナ（半角カタカナを含む）の語と、英字を含む英数字の語。
# 置換リスト（-OutCsv）では、前後にこの文字（カタカナは全角・半角とも）が続かない箇所だけを置換する
$kanaChars = '[ァ-ヴーｦ-ﾟ]'
$alnumChars = '[A-Za-z0-9Ａ-Ｚａ-ｚ０-９]'
$kanaRx = New-Object System.Text.RegularExpressions.Regex('[ァ-ヴー]+|[ｦ-ﾟ]+')
$alnumRx = New-Object System.Text.RegularExpressions.Regex($alnumChars + '+')

# ---------------------------------------------------------------- 用語集


$entries = New-Object System.Collections.Generic.List[object]
if ($Glossary) {
    $Glossary = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Glossary)
    if (-not (Test-Path -LiteralPath $Glossary)) { throw "用語集が見つかりません: $Glossary" }
    foreach ($row in @(Read-TextFile $Glossary | ConvertFrom-Csv)) {
        $names = @($row.PSObject.Properties | ForEach-Object { $_.Name })
        if ($names -notcontains '正しい表記' -or $names -notcontains '揺れ') { throw "用語集には「正しい表記」「揺れ」の列が必要です: $Glossary" }
        $preferred = ([string]$row.'正しい表記').Trim()
        if (-not $preferred) { continue }
        $variants = @(([string]$row.'揺れ') -split '[/／、,|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -cne $preferred })
        if ($variants.Count -eq 0) { continue }
        $entries.Add([pscustomobject]@{
            Preferred = $preferred
            Variants  = $variants
            # 揺れの一致のうち、正しい表記の一部になっている箇所は数えない（「ユーザ」が「ユーザー」の中にある場合など）
            Protect   = New-Object System.Text.RegularExpressions.Regex([regex]::Escape($preferred))
            Rx        = New-Object System.Text.RegularExpressions.Regex((@($variants | Sort-Object Length -Descending | ForEach-Object { [regex]::Escape($_) }) -join '|'))
        })
    }
}

# ---------------------------------------------------------------- 文字を集める


$searchArgs = @{ Pattern = '\S'; Path = $Path; Target = $Target }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
if ($IncludeWord) { $searchArgs['IncludeWord'] = $true }
$hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$fileCount = @($hits | ForEach-Object { $_.Path } | Sort-Object -Unique).Count

function Test-WordPath {
    param([string] $FilePath)
    return @('.docx', '.docm') -contains [System.IO.Path]::GetExtension($FilePath).ToLowerInvariant()
}

# 表記 → 出現箇所の一覧（用語集と自動の検出で別々に数える）。表記は大文字・小文字を区別する
function New-OrdinalTable { return , (New-Object System.Collections.Hashtable([System.StringComparer]::Ordinal)) }
$glossaryOcc = New-OrdinalTable
$autoOcc = New-OrdinalTable
$addOccurrence = {
    param($Table, [string] $Form, $Hit)
    if (-not $Table.ContainsKey($Form)) { $Table[$Form] = New-Object System.Collections.Generic.List[object] }
    $Table[$Form].Add($Hit)
}
$glossaryHits = @{}            # 用語集の行の番号 → 見つかった揺れの表記
$autoGroups = New-OrdinalTable  # 正規化した語 → 見つかった表記
foreach ($h in $hits) {
    $text = $h.Text
    $preferredAt = $null  # 正しい表記 → この文字の中の一致（Found）と、数えた揺れの一致（Variants）。正しい表記があるときだけ作る
    for ($ei = 0; $ei -lt $entries.Count; $ei++) {
        $e = $entries[$ei]
        $protected = @($e.Protect.Matches($text))
        if ($protected.Count -gt 0) {
            if ($null -eq $preferredAt) { $preferredAt = New-OrdinalTable }
            if (-not $preferredAt.ContainsKey($e.Preferred)) { $preferredAt[$e.Preferred] = @{ Found = $protected; Variants = New-Object System.Collections.Generic.List[object] } }
        }
        foreach ($m in $e.Rx.Matches($text)) {
            $inside = $false
            foreach ($p in $protected) { if ($m.Index -ge $p.Index -and $m.Index + $m.Length -le $p.Index + $p.Length) { $inside = $true; break } }
            if ($inside) { continue }
            & $addOccurrence $glossaryOcc $m.Value $h
            if ($protected.Count -gt 0) { $preferredAt[$e.Preferred].Variants.Add($m) }
            if (-not $glossaryHits.ContainsKey($ei)) { $glossaryHits[$ei] = New-OrdinalTable }
            $glossaryHits[$ei][$m.Value] = $true
        }
    }
    # 正しい表記は、同じ正しい表記の行が複数あっても 1 回だけ数える。数えた揺れの一部になっている箇所
    # （正しい表記「サーバ」・揺れ「サーバー」のときの、「サーバー」の中の「サーバ」）は数えない
    if ($null -ne $preferredAt) {
        foreach ($pf in $preferredAt.Keys) {
            foreach ($p in $preferredAt[$pf].Found) {
                $inside = $false
                foreach ($m in $preferredAt[$pf].Variants) { if ($p.Index -ge $m.Index -and $p.Index + $p.Length -le $m.Index + $m.Length) { $inside = $true; break } }
                if (-not $inside) { & $addOccurrence $glossaryOcc $pf $h }
            }
        }
    }
    if ($NoAuto) { continue }
    # カタカナの語は、全角・半角と長音（ー）の違いを無視した形でまとめる（長音を除いて 2 文字以上の語だけ）
    foreach ($m in $kanaRx.Matches($text)) {
        $key = 'カナ:' + $m.Value.Normalize($nfkc).Replace('ー', '')
        if ($key.Length -lt 5) { continue }
        & $addOccurrence $autoOcc $m.Value $h
        if (-not $autoGroups.ContainsKey($key)) { $autoGroups[$key] = New-OrdinalTable }
        $autoGroups[$key][$m.Value] = $true
    }
    # 英字を含む英数字の語は、全角・半角の違いを無視した形でまとめる（大文字・小文字は区別する）
    foreach ($m in $alnumRx.Matches($text)) {
        $n = $m.Value.Normalize($nfkc)
        if ($n.Length -lt 2 -or $n -notmatch '[A-Za-z]') { continue }
        $key = '英数:' + $n
        & $addOccurrence $autoOcc $m.Value $h
        if (-not $autoGroups.ContainsKey($key)) { $autoGroups[$key] = New-OrdinalTable }
        $autoGroups[$key][$m.Value] = $true
    }
}

# ---------------------------------------------------------------- 結果


$results = New-Object System.Collections.Generic.List[object]
$count = { param($Table, $f) if ($Table.ContainsKey($f)) { $Table[$f].Count } else { 0 } }
$newResult = {
    param($Table, [string] $Preferred, [string] $Variant, [string] $Kind)
    $list = $Table[$Variant]
    $examples = @($list | Select-Object -First 3 | ForEach-Object { '{0} {1}' -f $_.File, $_.Location }) -join '、'
    $r = [pscustomobject]@{
        Preferred = $Preferred; Variant = $Variant; Count = $list.Count; PreferredCount = (& $count $Table $Preferred)
        Kind = $Kind; Files = @($list | ForEach-Object { $_.Path } | Sort-Object -Unique).Count; Examples = $examples; Occurrences = $list.ToArray()
    }
    $r.PSObject.TypeNames.Insert(0, 'ExcelTerms.Variant')
    $results.Add($r)
}

$reported = New-OrdinalTable
$glossaryPreferred = New-OrdinalTable
foreach ($e in $entries) { $glossaryPreferred[$e.Preferred] = $true }
for ($ei = 0; $ei -lt $entries.Count; $ei++) {
    if (-not $glossaryHits.ContainsKey($ei)) { continue }
    foreach ($v in $glossaryHits[$ei].Keys) { & $newResult $glossaryOcc $entries[$ei].Preferred $v '用語集'; $reported[$v] = $true }
}
foreach ($key in $autoGroups.Keys) {
    $forms = @($autoGroups[$key].Keys)
    if ($forms.Count -lt 2) { continue }
    # 推奨: 用語集の正しい表記があればそれ。無ければ、全角の英数字と半角カタカナを除いた表記のうち多く使われている表記で、
    #       同じ数なら長い方（長音のある方）。全角の英数字・半角カタカナしか無ければ、半角・全角にした表記ごとに件数を合わせて比べる
    $preferred = @($forms | Where-Object { $glossaryPreferred.ContainsKey($_) })[0]
    if (-not $preferred) {
        $scores = New-OrdinalTable
        foreach ($f in $forms) { if ($f -ceq $f.Normalize($nfkc)) { $scores[$f] = & $count $autoOcc $f } }
        if ($scores.Count -eq 0) {
            foreach ($f in $forms) { $n = $f.Normalize($nfkc); $scores[$n] = [int]$scores[$n] + (& $count $autoOcc $f) }
        }
        $preferred = @($scores.Keys | Sort-Object @{ e = { $scores[$_] }; Descending = $true }, @{ e = { $_.Length }; Descending = $true }, @{ e = { $_ } })[0]
    }
    foreach ($f in $forms) {
        if ($f -ceq $preferred -or $reported.ContainsKey($f)) { continue }
        $kind = '長音（ー）'
        if ($f.Normalize($nfkc) -ceq $preferred.Normalize($nfkc)) { $kind = '全角・半角' }
        & $newResult $autoOcc $preferred $f $kind
    }
}
$sorted = @($results | Sort-Object @{ e = { $_.Kind -ne '用語集' } }, @{ e = { $_.Count }; Descending = $true }, Preferred, Variant -CaseSensitive)
$sorted

if ($OutCsv) {
    # Replace-Excel.ps1 -ListPath にそのまま渡せる置換リスト（UTF-8・BOM 付き。Excel で開いて直してから使える）。
    # 置換は上の行から順に行うので、長い表記から並べる（短い表記で、長い表記の一部を先に書き換えない）。
    # 自動の検出の揺れは、見つけたときと同じ語の単位で置換する正規表現の行にする（「モニタ」→「モニター」で
    # 「モニタリング」を「モニターリング」にしない）。用語集の揺れは、そのままの文字の行
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('置換前,置換後,正規表現,種類,件数,例')
    $q = { param($t) '"' + ([string]$t).Replace('"', '""') + '"' }
    foreach ($r in @($sorted | Sort-Object @{ e = { $_.Variant.Length }; Descending = $true }, @{ e = { $_.Kind -ne '用語集' } }, @{ e = { $_.Count }; Descending = $true }, Preferred, Variant -CaseSensitive)) {
        $find = $r.Variant
        $isRegex = ''
        if ($r.Kind -ne '用語集') {
            $chars = $kanaChars; $caseOption = ''
            # 英数字の語は大文字・小文字の違う語（ＩＤ と ｉｄ）を別の語として見つけるので、置換でも大文字・小文字を区別する
            if ($alnumRx.IsMatch($r.Variant)) { $chars = $alnumChars; $caseOption = '(?-i)' }
            $find = '{2}(?<!{0}){1}(?!{0})' -f $chars, [regex]::Escape($r.Variant), $caseOption
            $isRegex = '○'
        }
        $lines.Add((@((& $q $find), (& $q $r.Preferred), $isRegex, (& $q $r.Kind), $r.Count, (& $q $r.Examples)) -join ','))
    }
    [System.IO.File]::WriteAllText($OutCsv, ($lines -join "`r`n") + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
}

if ($OutXlsx) {
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $detailRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $sorted) {
        $summaryRows.Add(@{ Values = @($r.Preferred, $r.Variant, $r.Kind, $r.Count, $r.PreferredCount, $r.Files, $r.Examples) })
        foreach ($o in $r.Occurrences) {
            $loc = ''
            # Word の文書は開くだけ（セルの場所が無い）
            if ($o.Sheet -and $o.Cell -and -not (Test-WordPath $o.Path)) { $loc = "'{0}'!{1}" -f $o.Sheet.Replace("'", "''"), $o.Cell }
            $detailRows.Add(@{ Values = @($r.Variant, $r.Preferred, $o.File, $o.Location, $o.Kind, $o.Text); Link = @{ Column = 3; Path = $o.Path; Location = $loc } })
        }
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{
                Name = '表記ゆれ'
                Columns = @(@{ Header = '推奨'; Width = 24 }, @{ Header = '揺れ'; Width = 24 }, @{ Header = '種類'; Width = 12 }, @{ Header = '件数'; Width = 8 },
                    @{ Header = '推奨の件数'; Width = 10 }, @{ Header = 'ファイル数'; Width = 10 }, @{ Header = '例'; Width = 60 })
                Rows = $summaryRows
            },
            @{
                Name = '出現箇所'
                Columns = @(@{ Header = '揺れ'; Width = 24 }, @{ Header = '推奨'; Width = 24 }, @{ Header = 'ファイル'; Width = 36 }, @{ Header = '場所'; Width = 20 },
                    @{ Header = '種類'; Width = 8 }, @{ Header = '文字'; Width = 60 })
                Rows = $detailRows
            })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$total = 0
foreach ($r in $sorted) { $total += $r.Count }
Write-Host ("{0} ファイルを調べました: 表記ゆれ {1} 種類（{2} か所）" -f $fileCount, $sorted.Count, $total)
if ($OutCsv) { Write-Host "置換リストを保存しました: $OutCsv （確認・修正してから Replace-Excel.ps1 -ListPath で使えます）" }
$wordFiles = @($sorted | ForEach-Object { $_.Occurrences } | Where-Object { Test-WordPath $_.Path } | ForEach-Object { $_.File } | Sort-Object -Unique)
if ($wordFiles.Count -gt 0) { Write-Host ("Word の文書の揺れ（{0}）は Replace-Excel.ps1 では直せません。Word の置換（Ctrl+H）で直してください" -f ($wordFiles -join '、')) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     複数の Excel ファイルの表記ゆれ（ユーザ／ユーザー、ＩＤ／ID など）を見つけます（Excel 不要）。
#
# .DESCRIPTION
#     セルの値・コメント・図形の文字を読み、次の 2 つの方法で表記ゆれを見つけます。-IncludeWord で Word の文書も調べます。
#
#       用語集（-Glossary）  CSV の「正しい表記」と「揺れ」（/ や 、 で区切って複数可）に沿って、揺れの表記を数えます。
#                            揺れの一致のうち、正しい表記の一部になっている箇所（「ユーザー」の中の「ユーザ」）は数えません。
#       自動                 カタカナの語で長音（ー）の有無だけが違うもの（サーバ／サーバー、インタフェース／インターフェース）、
#                            全角・半角だけが違うもの（ＩＤ／ID、ﾕｰｻﾞｰ／ユーザー）を見つけます（-NoAuto で無効）。
#
#     結果は揺れの表記ごとに 1 行で、推奨の表記（用語集の正しい表記、または多く使われている表記）と件数を出力します。
#     -OutCsv で保存した置換リストは、確認・修正してから Replace-Excel.ps1 -ListPath で一括置換に使えます。
#
# .PARAMETER Path
#     対象のファイルまたはフォルダ（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER Glossary
#     用語集の CSV。「正しい表記」「揺れ」の列が必要です（UTF-8 / Shift_JIS）。
#
# .PARAMETER NoAuto
#     自動の検出（長音・全角半角）をしません。用語集だけで確かめます。
#
# .PARAMETER Target
#     調べる対象。Value（セルの値）/ Comment（コメント）/ Shape（図形）/ SheetName（シート名）。既定は Value, Comment, Shape。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER IncludeWord
#     Word の文書（.docx / .docm）の本文・表・脚注・コメント・テキストボックスも調べます（「送る」の表記ゆれはこれを付けて実行します）。
#     Word の揺れは Replace-Excel.ps1 では直せないので、Word の置換（Ctrl+H）で直してください。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は調べられなくなります）。
#
# .PARAMETER OutCsv
#     置換リスト（置換前・置換後・正規表現・種類・件数・例）を CSV に保存します。Replace-Excel.ps1 -ListPath で使えます。
#     長い表記の行から並べ、自動で見つけた揺れは、見つけた語だけを置換する正規表現の行にします（「モニタリング」の「モニタ」は置換しない）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（表記ゆれの一覧と、出現箇所へのリンク付きの一覧）。
#
# .EXAMPLE
#     .\Test-ExcelTerms.ps1 D:\設計書
#
# .EXAMPLE
#     # 用語集で確かめ、置換リストを作って一括で直す
#     .\Test-ExcelTerms.ps1 D:\設計書 -Glossary .\用語集.csv -OutCsv D:\置換リスト.csv
#     .\Replace-Excel.ps1 -ListPath D:\置換リスト.csv -Path D:\設計書 -WhatIf
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     自動の検出は「どちらかに統一したほうがよい表記」の候補です。意味の違う語が混ざることもあるので、確認してから直してください。
