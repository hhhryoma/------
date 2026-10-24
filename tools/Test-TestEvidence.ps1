[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $EvidencePath,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
foreach ($name in @('Headers', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
# エビデンスの置き場所を指定しなければ、テスト仕様書と同じ場所から探す（テスト仕様書のブックのほかのシート・同じフォルダの画像など）
if (-not $EvidencePath) { $EvidencePath = $Path }

Update-TypeData -TypeName 'TestEvidence.Issue' -DefaultDisplayPropertySet File, Location, Level, Item, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# エビデンスとして読むファイル。.xlsx / .xlsm はシートごと（画像と文字を見る）、画像は 1 枚ずつ、そのほかはファイルごと（中は見ない）
$bookExtensions = @('.xlsx', '.xlsm')
$imageExtensions = @('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tif', '.tiff')
$fileExtensions = @('.xls', '.xlsb', '.pdf', '.docx', '.docm', '.doc', '.pptx', '.ppt', '.txt', '.log', '.csv', '.htm', '.html', '.mp4', '.wmv', '.avi', '.mov')
# エビデンスの列の値のうち、要らないもの（なし・-）と、あるとだけ書いたもの（あり・○）。ほかの値はシート名・ファイル名として探す
$notNeededPattern = '^(-|―|ー|なし|無し|無|不要|対象外|N/?A)$'
$claimedPattern = '^(あり|有|有り|○|〇|◯|済|済み|取得済み?|OK)$'
# ケースの始まりの目印（行のいちばん左のセルの「■UT-001」「【No.1】」「テストID: UT-001」など。全角・半角と大文字小文字をそろえてから比べる）
$markerRx = New-Object System.Text.RegularExpressions.Regex('^(?<mark>[■□●◆◇▼▽★☆※#【\[]+)?\s*(?<label>(?:テストケース|テスト|ケース|試験|確認)?(?:ID|NO\.?|番号)|テストケース|テスト|ケース|試験)?\s*[:.]?\s*(?<id>[A-Z0-9][A-Z0-9_\-.]*)')
# 名前の言葉のうち、どの機能にも付く一般的なもの（版数・テスト仕様書・エビデンスなど。機能を見分けるのに使わない）
$genericWordPattern = '^(V?[\d.]+|UT|IT|ST|SIT|UAT|PT|(単体|結合|総合|システム|受入|運用)?(テスト|試験)?(仕様書?|項目書?|ケース|結果|成績書?|報告書?|エビデンス|証跡|キャプチャ|ハードコピー|スクリーンショット)?)$'
# 同じ画像がこの数より多くのケースのエビデンスにあれば、ロゴなどの共通の画像とみなす（使い回しの警告にしない）
$commonImageCases = 4

# ---------------------------------------------------------------- 名前を比べる


function ConvertTo-CaseKey {
    #     テストID・シート名・ファイル名・目印を比べる形。全角・半角と大文字小文字をそろえ、前に付けた「■」「No.」「ケース」「テストID:」と
    #     空白を除き、数字の前の 0 を除く（UT-001 と UT-1、No.01 と 1 を同じにする）
    param([string] $Text)
    $t = ([string]$Text).Normalize($nfkc).ToUpperInvariant().Trim()
    $t = $t -replace '^[■□●◆◇▼▽★☆※#【\[(]+\s*', ''
    $t = $t -replace '^(テストケース|テスト|ケース|試験|確認)?(ID|NO\.?|番号)?\s*[:.]?\s*(?=\d)', ''
    $t = $t -replace '^(テストケース|テスト|ケース|試験|確認)?(ID|NO\.?|番号)\s*:\s*', ''
    $t = ($t -replace '[】\])]+$', '') -replace '\s', ''
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -lt $t.Length) {
        if (-not [char]::IsDigit($t[$i])) { [void]$sb.Append($t[$i]); $i++; continue }
        $j = $i
        while ($j -lt $t.Length -and [char]::IsDigit($t[$j])) { $j++ }
        $digits = $t.Substring($i, $j - $i).TrimStart('0')
        if (-not $digits) { $digits = '0' }
        [void]$sb.Append($digits)
        $i = $j
    }
    return $sb.ToString()
}

function Get-NameWords {
    #     名前を「_」・空白などで分けた言葉（全角・半角と大文字小文字をそろえる）。一般的な言葉（テスト仕様書・エビデンス・版数など）と 1 文字のものは除く
    param([string[]] $Names)
    $words = New-Object System.Collections.Generic.List[string]
    foreach ($n in $Names) {
        foreach ($w in (([string]$n).Normalize($nfkc).ToUpperInvariant() -split '[\\/_\s・、,()\[\]【】「」]+')) {
            if ($w.Length -ge 2 -and $w -notmatch $genericWordPattern -and -not $words.Contains($w)) { $words.Add($w) }
        }
    }
    return , $words
}

function ConvertTo-FileKey {
    #     ファイルの名前を比べる形（フォルダと、エビデンスとして読む拡張子を除き、全角・半角と空白、大文字小文字をそろえる）
    param([string] $Name)
    $n = ([string]$Name).Trim().Split(@('\', '/'))[-1]
    if ($n -match '\.([A-Za-z0-9]+)$' -and ($bookExtensions + $imageExtensions + $fileExtensions) -contains ('.' + $Matches[1].ToLowerInvariant())) { $n = $n.Substring(0, $n.Length - $Matches[0].Length) }
    return (ConvertTo-Label $n).ToUpperInvariant()
}

function Add-ToIndex {
    param([hashtable] $Index, [string] $Key, $Value)
    if (-not $Key) { return }
    if (-not $Index.ContainsKey($Key)) { $Index[$Key] = New-Object System.Collections.Generic.List[object] }
    if (-not $Index[$Key].Contains($Value)) { $Index[$Key].Add($Value) }
}

# ---------------------------------------------------------------- テスト仕様書を読む


$fields = @('機能', 'テストID', 'テスト内容', '結果', 'エビデンス')
$candidates = Read-HeaderCandidates $Headers $fields
# 読み取り結果（どのファイル・シートをどの見出しで読んだか）
$readLog = New-Object System.Collections.Generic.List[object]
$records = Read-HeaderRecords -Inputs (Get-InputFiles $Path 'テスト仕様書' -NoRecurse:$NoRecurse) -Candidates $candidates -Fields $fields `
    -Need @('結果', @('テストID', 'エビデンス')) -Key @('テストID', 'テスト内容', 'エビデンス') -What 'テスト仕様書' `
    -Exclude $OutXlsx -NoRecurse:$NoRecurse -NoExcel:$NoExcel -ReadLog $readLog
$cases = ConvertTo-TestCases $records
if ($cases.Count -eq 0) { Write-Warning 'テストケースが見つかりません（結果と、テストID かエビデンスの見出しがある表を探します）' }

# ケース。名前で探すため、テストID と、機能とテストID をつないだもの（SCR001_UT-001 など）で引けるようにする
$caseInfo = New-Object System.Collections.Generic.List[object]
$byKey = @{}
$byQualified = @{}
foreach ($c in $cases) {
    $evidence = ([string]$c.Values['エビデンス']).Normalize($nfkc).Trim()
    $info = [pscustomobject]@{
        Case = $c; Id = $c.Id; Function = $c.Function; Kind = $c.Kind; Key = ''; Evidence = $evidence
        NotNeeded = ($evidence -match $notNeededPattern); Claimed = ($evidence -match $claimedPattern)
        Words = (Get-NameWords @($c.Function, [System.IO.Path]::GetFileNameWithoutExtension($c.Place.File), $c.Place.Sheet))
        Units = (New-Object System.Collections.Generic.List[object]); Missing = (New-Object System.Collections.Generic.List[string]); Seq = $caseInfo.Count
    }
    $rawId = ([string]$c.Values['テストID']).Trim()
    if ($rawId) {
        $info.Key = ConvertTo-CaseKey $rawId
        Add-ToIndex $byKey $info.Key $info
        $functionKey = ConvertTo-CaseKey $c.Function
        if ($functionKey -and $info.Key) { foreach ($sep in @('', '_', '-')) { Add-ToIndex $byQualified ($functionKey + $sep + $info.Key) $info } }
    }
    $caseInfo.Add($info)
}

# テスト仕様書の表のあるシート（ブックのパス|シート名 → そのシートのケース。行の順）。表の中の画像は、その行のケースのエビデンスにする
$specPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$tableSheets = @{}
foreach ($info in $caseInfo) {
    $p = $info.Case.Place
    [void]$specPaths.Add($p.Path)
    if ($p.Sheet) { Add-ToIndex $tableSheets ($p.Path + '|' + $p.Sheet) $info }
}

function Find-CaseCandidates {
    #     名前（シート名・ファイル名・目印）に合うケースの候補。テストID、機能とテストID をつないだもの（SCR001_UT-001 など）、
    #     後ろの「_2」「(2)」を除いたもの（同じケースの 2 枚目）、最後の「_」・空白のあと（受注一覧_No1 など。前が機能の名前のものだけ）の順に探す
    param([string] $Name)
    $key = ConvertTo-CaseKey $Name
    if (-not $key) { return @() }
    foreach ($k in @($key, ($key -replace '[_\-(]\d+\)?$', ''))) {
        if (-not $k) { continue }
        if ($byKey.ContainsKey($k)) { return $byKey[$k].ToArray() }
        if ($byQualified.ContainsKey($k)) { return $byQualified[$k].ToArray() }
    }
    $parts = @(([string]$Name).Normalize($nfkc).Trim() -split '[_\s]+' | Where-Object { $_ })
    if ($parts.Count -lt 2) { return @() }
    $tailKey = ConvertTo-CaseKey $parts[-1]
    if (-not $byKey.ContainsKey($tailKey)) { return @() }
    $head = ($parts[0..($parts.Count - 2)] -join '_').ToUpperInvariant()
    return @($byKey[$tailKey] | Where-Object { @($_.Words | Where-Object { $head.Contains($_) }).Count -gt 0 })
}

# ---------------------------------------------------------------- エビデンスを読む


$units = New-Object System.Collections.Generic.List[object]

function New-Unit {
    #     エビデンスの 1 単位（ブック・シート・目印からの行・テスト仕様書の表の中の画像・画像のファイル・そのほかのファイル）
    #     Context は、ケースを見分けるための置き場所の名前（フォルダ・ファイル・シート）
    param([string] $Kind, [string] $FilePath, [string] $Sheet, [string] $Name, [string] $Cell, [int] $Row)
    $location = [System.IO.Path]::GetFileName($FilePath)
    if ($Sheet) { $location = '{0}!{1}' -f $Sheet, $Cell }
    $context = '{0}\{1}' -f [System.IO.Path]::GetDirectoryName($FilePath), [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    if ($Sheet) { $context += '\' + $Sheet }
    $unit = [pscustomobject]@{
        Kind = $Kind; File = [System.IO.Path]::GetFileName($FilePath); Path = $FilePath; Sheet = $Sheet; Name = $Name; Cell = $Cell; Row = $Row; Location = $location
        Context = $context.Normalize($nfkc).ToUpperInvariant(); SheetKey = (ConvertTo-Label $Sheet).ToUpperInvariant(); FileKey = (ConvertTo-FileKey $FilePath)
        Images = (New-Object System.Collections.Generic.List[object]); Texts = 0
        Cases = (New-Object System.Collections.Generic.List[object]); Candidates = @(); Seq = $units.Count
    }
    $units.Add($unit)
    return $unit
}

function Add-UnitCase {
    param($Unit, $Info)
    if ($Unit.Cases.Contains($Info)) { return }
    $Unit.Cases.Add($Info)
    $Info.Units.Add($Unit)
    $Unit.Candidates = @()
}

function Select-Case {
    #     候補のケースから、エビデンスの置き場所に合うものを選ぶ。同じブックにある表のケースと、機能・テスト仕様書の名前の言葉が
    #     置き場所の名前（フォルダ・ファイル・シート）に長く含まれるケースを選ぶ。戻り値: いちばん合うケース（1 つに決まらなければすべて）
    param($Candidates, $Unit)
    $list = @($Candidates)
    if ($list.Count -le 1) { return $list }
    $best = New-Object System.Collections.Generic.List[object]
    $bestScore = -1
    foreach ($c in $list) {
        $score = 0
        if ([string]::Equals($c.Case.Place.Path, $Unit.Path, [System.StringComparison]::OrdinalIgnoreCase)) { $score += 1000 }
        foreach ($w in $c.Words) { if ($Unit.Context.Contains($w)) { $score += $w.Length } }
        if ($score -gt $bestScore) { $best.Clear(); $bestScore = $score }
        if ($score -eq $bestScore) { $best.Add($c) }
    }
    return $best.ToArray()
}

function Set-UnitCase {
    #     エビデンスを、名前で見つけた候補のケースに結び付ける（1 つに決まらなければ候補を残す）
    param($Unit, $Candidates)
    $chosen = @(Select-Case $Candidates $Unit)
    if ($chosen.Count -eq 1) { Add-UnitCase $Unit $chosen[0] }
    else { $Unit.Candidates = $chosen }
}

function Find-PlaceCandidates {
    #     ファイルの置き場所に合うケースの候補。フォルダの名前（ケースごとのフォルダ。UT-001\1.png など）、合わなければファイルの名前
    param($File)
    $found = @(Find-CaseCandidates $File.Directory.Name)
    if ($found.Count -eq 0) { $found = @(Find-CaseCandidates ([System.IO.Path]::GetFileNameWithoutExtension($File.Name))) }
    return $found
}

function Get-FileHashText {
    #     画像のファイルの中身の要約（ExcelTools.psm1 の Read-BookDrawings の画像の Hash と同じ形）
    param([string] $FilePath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $fs = [System.IO.File]::OpenRead($FilePath)
        try { $bytes = $sha.ComputeHash($fs) } finally { $fs.Dispose() }
    }
    finally { $sha.Dispose() }
    return ([System.BitConverter]::ToString($bytes)).Replace('-', '').Substring(0, 16).ToLowerInvariant()
}

function Get-CellCount {
    #     シートの文字のあるセルの数（From〜To 行）
    param($Grid, [int] $From = 0, [int] $To = [int]::MaxValue)
    if (-not $Grid) { return 0 }
    $n = 0
    foreach ($row in $Grid.Rows) { if ($row -ge $From -and $row -le $To) { $n += $Grid.Cols[$row].Count } }
    return $n
}

function Find-Markers {
    #     シートの中のケースの始まりの目印（行のいちばん左のセルの「■UT-001」「【No.1】」「テストID: UT-001」など）。ケースに合うものだけ。
    #     数字だけの目印（1・2 など）は、■ や No. などを付けたものだけ（表の行番号と見分けるため）
    param($Grid)
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $Grid) { return , $list }
    foreach ($row in $Grid.Rows) {
        $col = $Grid.Cols[$row][0]
        $text = ([string]$Grid.Cells[$row][$col]).Normalize($nfkc).Trim()
        if ($text.Length -gt 60) { continue }
        $m = $markerRx.Match($text.ToUpperInvariant())
        if (-not $m.Success) { continue }
        $id = $m.Groups['id'].Value.TrimEnd('.', '-', '_')
        if ($id -match '^[\d.\-]+$' -and -not ($m.Groups['mark'].Success -or $m.Groups['label'].Success)) { continue }
        $found = @(Find-CaseCandidates $id)
        if ($found.Count -eq 0) { continue }
        $list.Add([pscustomobject]@{ Row = $row; Cell = ('{0}{1}' -f (ConvertTo-ColumnName $col), $row); Text = $text; Candidates = $found })
    }
    return , $list
}

function Add-ImageTo {
    param($Unit, $Picture)
    $Unit.Images.Add([pscustomobject]@{ Hash = $Picture.Hash; Cell = $Picture.Cell })
}

function Add-BookUnits {
    #     .xlsx / .xlsm のエビデンス。フォルダかブックの名前がケースに合えば、ブック全体を 1 つのエビデンスにする。
    #     ほかはシートごと。シートの名前がケースに合わなければ、シートの中の目印で分ける。テスト仕様書の表のシートは、表の中の画像を行のケースに
    param($File, $Book)
    try { $drawings = Read-BookDrawings $File.FullName }
    catch { Write-Warning ("エビデンスのブックを読めません: {0} : {1}" -f $File.FullName, $_.Exception.Message); return }
    $grids = @{}
    if ($Book) { foreach ($s in $Book.Sheets.Keys) { $grids[$s] = $Book.Sheets[$s] } }
    if (-not $specPaths.Contains($File.FullName)) {
        $found = @(Find-PlaceCandidates $File)
        if ($found.Count -gt 0) {
            $unit = New-Unit 'ブック' $File.FullName '' ([System.IO.Path]::GetFileNameWithoutExtension($File.Name)) '' 0
            foreach ($sheet in $drawings.Keys) {
                foreach ($pic in @($drawings[$sheet] | Where-Object { $_.Kind -eq 'pic' -and -not $_.Hidden })) { Add-ImageTo $unit $pic }
                $unit.Texts += Get-CellCount $grids[$sheet]
            }
            Set-UnitCase $unit $found
            return
        }
    }
    foreach ($sheet in $drawings.Keys) {
        $pics = @($drawings[$sheet] | Where-Object { $_.Kind -eq 'pic' -and -not $_.Hidden })
        $grid = $grids[$sheet]
        $tableKey = $File.FullName + '|' + $sheet
        if ($tableSheets.ContainsKey($tableKey)) {
            # 表の中（ケースの行より下）の画像は、その行か、上でいちばん近い行のケースのエビデンス。表より上の画像（ロゴなど）は見ない
            $rows = @($tableSheets[$tableKey] | Sort-Object { $_.Case.Place.Row })
            $inline = @{}
            foreach ($pic in $pics) {
                $owner = $null
                foreach ($info in $rows) { if ($info.Case.Place.Row -le $pic.Anchor.R1) { $owner = $info } else { break } }
                if (-not $owner) { continue }
                if (-not $inline.ContainsKey($owner.Seq)) {
                    $inline[$owner.Seq] = New-Unit '表の中' $File.FullName $sheet $sheet $pic.Cell $pic.Anchor.R1
                    $inline[$owner.Seq].Texts = 1
                    Add-UnitCase $inline[$owner.Seq] $owner
                }
                Add-ImageTo $inline[$owner.Seq] $pic
            }
            continue
        }
        $found = @(Find-CaseCandidates $sheet)
        if ($found.Count -gt 0) {
            $unit = New-Unit 'シート' $File.FullName $sheet $sheet 'A1' 1
            foreach ($pic in $pics) { Add-ImageTo $unit $pic }
            $unit.Texts = Get-CellCount $grid
            Set-UnitCase $unit $found
            continue
        }
        $markers = Find-Markers $grid
        # 目印の無いシート、目印より上の部分は、どのケースか分からないエビデンス（画像があれば知らせる）
        $top = [int]::MaxValue
        if ($markers.Count -gt 0) { $top = $markers[0].Row }
        $rest = New-Unit 'シート' $File.FullName $sheet $sheet 'A1' 1
        foreach ($pic in @($pics | Where-Object { $_.Anchor.R1 -lt $top })) { Add-ImageTo $rest $pic }
        $rest.Texts = Get-CellCount $grid 0 ($top - 1)
        for ($i = 0; $i -lt $markers.Count; $i++) {
            $mk = $markers[$i]
            $end = [int]::MaxValue
            if ($i + 1 -lt $markers.Count) { $end = $markers[$i + 1].Row - 1 }
            $unit = New-Unit '目印' $File.FullName $sheet $mk.Text $mk.Cell $mk.Row
            foreach ($pic in @($pics | Where-Object { $_.Anchor.R1 -ge $mk.Row -and $_.Anchor.R1 -le $end })) { Add-ImageTo $unit $pic }
            # 目印のセルは数えない
            $unit.Texts = (Get-CellCount $grid $mk.Row $end) - 1
            Set-UnitCase $unit $mk.Candidates
        }
    }
}

# エビデンスのファイル（パスの順。~$ で始まる一時ファイル・結果の xlsx・テスト仕様書の CSV などは除く）
$fileList = New-Object System.Collections.Generic.List[object]
$seenFiles = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($p in $EvidencePath) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "エビデンスの置き場所が見つかりません: $p"; continue }
    foreach ($item in $items) {
        $list = @($item)
        if ($item.PSIsContainer) { $list = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) }
        foreach ($f in $list) {
            $ext = $f.Extension.ToLowerInvariant()
            if ($f.Name.StartsWith('~$') -or ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase))) { continue }
            if (($bookExtensions + $imageExtensions + $fileExtensions) -notcontains $ext) { continue }
            if ($specPaths.Contains($f.FullName) -and $bookExtensions -notcontains $ext) { continue }
            if ($seenFiles.Add($f.FullName)) { $fileList.Add($f) }
        }
    }
}
# パスの順（[Array]::Sort に PowerShell の配列を渡すと、写しが並べ替わるだけなので、List の Sort で並べる）
$fileList.Sort([System.Comparison[object]] { param($a, $b) [string]::Compare($a.FullName, $b.FullName, [System.StringComparison]::OrdinalIgnoreCase) })
$files = $fileList.ToArray()

# ブックの文字（目印と、エビデンスが空でないかを見るため）
$books = @{}
$bookPaths = @($files | Where-Object { $bookExtensions -contains $_.Extension.ToLowerInvariant() } | ForEach-Object { $_.FullName })
if ($bookPaths.Count -gt 0) {
    $hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') -Pattern '\S' -Path $bookPaths -Target Value -NoExcel 6>$null)
    $grouped = ConvertTo-DesignBooks $hits
    foreach ($k in $grouped.Keys) { $books[$k] = $grouped[$k] }
}

foreach ($f in $files) {
    $ext = $f.Extension.ToLowerInvariant()
    if ($bookExtensions -contains $ext) { Add-BookUnits $f $books[$f.FullName]; continue }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
    if ($imageExtensions -contains $ext) {
        $unit = New-Unit '画像' $f.FullName '' $base '' 0
        try { $unit.Images.Add([pscustomobject]@{ Hash = (Get-FileHashText $f.FullName); Cell = '' }) }
        catch { Write-Warning ("画像を読めません: {0} : {1}" -f $f.FullName, $_.Exception.Message) }
    }
    else {
        $unit = New-Unit 'ファイル' $f.FullName '' $base '' 0
        if ($f.Length -gt 0) { $unit.Texts = 1 }
    }
    $found = @(Find-PlaceCandidates $f)
    if ($found.Count -gt 0) { Set-UnitCase $unit $found }
}

# ---------------------------------------------------------------- エビデンスの列に書いたもの


function Select-RefUnits {
    #     エビデンスの列に書いたシート・ファイルのエビデンスのうち、このケースに結び付けたもの、無ければどのケースのものでもないもの（目印で分けたシートなど）
    param($Info, $List)
    $mine = @($List | Where-Object { $_.Cases.Contains($Info) })
    if ($mine.Count -gt 0) { return $mine }
    $free = @($List | Where-Object { $_.Cases.Count -eq 0 -and @($_.Candidates).Count -le 1 -and ($_.Images.Count -gt 0 -or $_.Texts -gt 0) })
    if ($free.Count -gt 0) { return $free }
    return @($List | Select-Object -First 1)
}

function Find-RefUnits {
    #     エビデンスの列に書いたもの（「ブック!シート」「シート!セル」・シート名・ファイル名）のエビデンス
    param($Info, [string] $Ref)
    $book = ''; $sheet = $Ref
    $parts = @($Ref -split '!')
    if ($parts.Count -ge 2) {
        if ($parts[-1] -match '^\$?[A-Za-z]{1,3}\$?\d+$') { $sheet = $parts[-2]; if ($parts.Count -ge 3) { $book = $parts[-3] } }
        else { $sheet = $parts[-1]; $book = $parts[-2] }
    }
    $sheetKey = (ConvertTo-Label $sheet.Trim("'")).ToUpperInvariant()
    $bookKey = ''
    if ($book) { $bookKey = ConvertTo-FileKey $book.Trim("'", '[', ']') }
    $bySheet = @($units | Where-Object { $_.Sheet -and $_.SheetKey -eq $sheetKey -and (-not $bookKey -or $_.FileKey -eq $bookKey) })
    if ($bySheet.Count -gt 0) {
        # 同じ名前のシートが幾つかのブックにあれば、テスト仕様書と同じブックか、名前の合うブックのもの
        $books = @($bySheet | ForEach-Object { $_.Path } | Select-Object -Unique)
        if ($books.Count -gt 1) {
            $same = @($bySheet | Where-Object { [string]::Equals($_.Path, $Info.Case.Place.Path, [System.StringComparison]::OrdinalIgnoreCase) })
            if ($same.Count -gt 0) { $bySheet = $same }
            else {
                $bestPath = $null; $bestScore = -1
                foreach ($u in $bySheet) {
                    $score = 0
                    foreach ($w in $Info.Words) { if ($u.Context.Contains($w)) { $score += $w.Length } }
                    if ($score -gt $bestScore) { $bestScore = $score; $bestPath = $u.Path }
                }
                $bySheet = @($bySheet | Where-Object { $_.Path -eq $bestPath })
            }
        }
        return @(Select-RefUnits $Info $bySheet)
    }
    if ($book) { return @() }
    $refKey = ConvertTo-FileKey $Ref
    return @(Select-RefUnits $Info @($units | Where-Object { $_.FileKey -eq $refKey }))
}

foreach ($info in $caseInfo) {
    if (-not $info.Evidence -or $info.NotNeeded -or $info.Claimed) { continue }
    foreach ($ref in @($info.Evidence -split '[、,;\r\n↵]+' | ForEach-Object { $_.Trim() } | Where-Object { $_ })) {
        if ($ref -match $notNeededPattern -or $ref -match $claimedPattern) { continue }
        $found = @(Find-RefUnits $info $ref)
        if ($found.Count -eq 0) { $info.Missing.Add($ref); continue }
        foreach ($u in $found) { Add-UnitCase $u $info }
    }
}

# ---------------------------------------------------------------- 確かめる


$issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    #     指摘を足す。場所は、ケースの指摘はテスト仕様書のセル（エビデンスの列があればその列）、エビデンスの指摘はエビデンスの場所
    param([string] $Level, [string] $Item, $Info, $Unit, [string] $Detail)
    if ($Info) {
        $p = $Info.Case.Place
        $col = $p.Column
        if ($p.Columns['エビデンス']) { $col = $p.Columns['エビデンス'] }
        $location = '{0} 行目' -f $p.Row
        $cell = ''
        if ($p.Sheet) { $cell = '{0}{1}' -f (ConvertTo-ColumnName $col), $p.Row; $location = '{0}!{1}' -f $p.Sheet, $cell }
        $i = [pscustomobject]@{ File = $p.File; Location = $location; Level = $Level; Item = $Item; Function = $Info.Function; Id = $Info.Id; Detail = $Detail; Path = $p.Path; Sheet = $p.Sheet; Row = $p.Row; Col = $col; Seq = $issues.Count; Cell = $cell }
    }
    else {
        $i = [pscustomobject]@{ File = $Unit.File; Location = $Unit.Location; Level = $Level; Item = $Item; Function = ''; Id = ''; Detail = $Detail; Path = $Unit.Path; Sheet = $Unit.Sheet; Row = $Unit.Row; Col = 0; Seq = $issues.Count; Cell = $Unit.Cell }
        if ($Unit.Cell -match '^[A-Z]+') { $i.Col = Get-ColumnIndex $Matches[0] }
    }
    $i.PSObject.TypeNames.Insert(0, 'TestEvidence.Issue')
    $issues.Add($i)
}

function Get-UnitList {
    #     エビデンスの場所の一覧（「UT-001.png・エビデンス!B5」）
    param($List, [int] $Max = 3)
    $all = @($List | ForEach-Object { $_ })
    $text = (@($all | Select-Object -First $Max | ForEach-Object { if ($_.Sheet) { '{0} {1}' -f $_.File, $_.Location } else { $_.File } })) -join '・'
    if ($all.Count -gt $Max) { $text += (' ほか {0} 件' -f ($all.Count - $Max)) }
    return $text
}

foreach ($info in $caseInfo) {
    $done = @('OK', 'NG') -contains $info.Kind
    $result = ([string]$info.Case.Values['結果']).Trim()
    foreach ($ref in $info.Missing) {
        Add-Issue $(if ($done) { 'エラー' } else { '警告' }) 'エビデンスが見つからない' $info $null ("エビデンスの列の「{0}」が見つかりません（シート名・ファイル名と、エビデンスの置き場所を確かめてください）" -f $ref)
    }
    if ($info.Units.Count -eq 0) {
        if (-not $done -or $info.NotNeeded -or $info.Missing.Count -gt 0) { continue }
        $detail = '結果が {0} ですが、エビデンスが見つかりません' -f $result
        if ($info.Claimed) { $detail = 'エビデンスの列は「{0}」ですが、エビデンスが見つかりません' -f $info.Evidence }
        if (-not $info.Key) { $detail += '（テストID が無いため、シート名・ファイル名では探せません。エビデンスの列にシート名・ファイル名を書いてください）' }
        else { $detail += '（シート名・ファイル名・目印をテストID にしてください）' }
        Add-Issue 'エラー' 'エビデンスが無い' $info $null $detail
        continue
    }
    if (@($info.Units | Where-Object { $_.Images.Count -gt 0 -or $_.Texts -gt 0 }).Count -eq 0) {
        Add-Issue '警告' 'エビデンスが空' $info $null ('エビデンス（{0}）に、画像も文字もありません' -f (Get-UnitList $info.Units))
    }
    if ($info.Kind -eq '未実施' -or $info.Kind -eq '保留') {
        $what = '空欄'
        if ($result) { $what = '「{0}」' -f $result }
        Add-Issue '警告' '実施していないケースのエビデンス' $info $null ('結果が{0}ですが、エビデンスがあります（{1}）。結果の書き忘れではありませんか' -f $what, (Get-UnitList $info.Units))
    }
    elseif ($info.Kind -eq '対象外') {
        Add-Issue '情報' '対象外のケースのエビデンス' $info $null ('結果が「{0}」ですが、エビデンスがあります（{1}）' -f $result, (Get-UnitList $info.Units))
    }
}

# 同じ画像の使い回し（違うケースのエビデンスに同じ画像がある）。同じエビデンスを幾つかのケースで使うのは使い回しとみなさない
$byHash = [ordered]@{}
foreach ($u in $units) {
    if ($u.Cases.Count -eq 0) { continue }
    foreach ($h in @($u.Images | ForEach-Object { $_.Hash } | Where-Object { $_ } | Select-Object -Unique)) {
        if (-not $byHash.Contains($h)) { $byHash[$h] = New-Object System.Collections.Generic.List[object] }
        $byHash[$h].Add($u)
    }
}
$reused = @{}
foreach ($h in $byHash.Keys) {
    $us = $byHash[$h]
    $involved = New-Object System.Collections.Generic.List[object]
    foreach ($u in $us) { foreach ($c in $u.Cases) { if (-not $involved.Contains($c)) { $involved.Add($c) } } }
    if ($involved.Count -lt 2) { continue }
    if ($involved.Count -gt $commonImageCases) {
        Add-Issue '情報' '多くのケースにある画像' $null $us[0] ('同じ画像が {0} 件のケースのエビデンスにあります（{1} ほか）。ロゴなどの共通の画像なら問題ありません' -f $involved.Count, ((@($involved | Sort-Object Seq | Select-Object -First 3 | ForEach-Object { $_.Id })) -join '・'))
        continue
    }
    foreach ($u in $us) {
        foreach ($c in $u.Cases) {
            foreach ($other in $us) {
                if ($other -eq $u) { continue }
                foreach ($c2 in $other.Cases) {
                    if ($c2 -eq $c -or $u.Cases.Contains($c2)) { continue }
                    if (-not $reused.ContainsKey($c.Seq)) { $reused[$c.Seq] = New-Object System.Collections.Generic.List[object] }
                    if (-not $reused[$c.Seq].Contains($other)) { $reused[$c.Seq].Add($other) }
                }
            }
        }
    }
}
foreach ($info in $caseInfo) {
    if (-not $reused.ContainsKey($info.Seq)) { continue }
    $others = @($reused[$info.Seq] | Sort-Object Seq)
    $names = @($others | ForEach-Object { '{0}（{1}）' -f ((@($_.Cases | ForEach-Object { $_.Id })) -join '・'), (Get-UnitList @($_) 1) })
    Add-Issue '警告' '同じ画像の使い回し' $info $null ('エビデンスの画像が、{0}のエビデンスの画像と同じです（撮り直し忘れ・貼り間違いではありませんか）' -f ($names -join '、'))
}

# どのケースのエビデンスか分からないもの（画像のあるもの）と、決められないもの
foreach ($u in $units) {
    if ($u.Cases.Count -gt 0) { continue }
    if (@($u.Candidates).Count -gt 1) {
        $list = @($u.Candidates | ForEach-Object { '{0}（{1}）' -f $_.Id, $_.Function }) -join '・'
        Add-Issue '情報' 'どのケースか決められない' $null $u ('「{0}」に合うケースが {1} 件あります（{2}）。ブック・フォルダの名前に機能の名前を入れるか、エビデンスの列にシート名・ファイル名を書いてください' -f $u.Name, @($u.Candidates).Count, $list)
    }
    elseif ($u.Images.Count -gt 0) {
        Add-Issue '情報' 'どのケースか分からないエビデンス' $null $u ('画像が {0} 枚ありますが、どのケースのエビデンスか分かりません（シート名・ファイル名・目印をテストID にするか、エビデンスの列に書いてください）' -f $u.Images.Count)
    }
}

# ---------------------------------------------------------------- 結果


$sorted = @($issues | Sort-Object @{ e = { @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }[$_.Level] } }, File, Sheet, Row, Col, Seq)
$sorted

function Get-UnitLink {
    param($Unit, [int] $Column)
    $link = @{ Column = $Column; Path = $Unit.Path }
    if ($Unit.Sheet) { $link.Location = "'{0}'!{1}" -f $Unit.Sheet.Replace("'", "''"), $Unit.Cell }
    return $link
}

if ($OutXlsx) {
    $issueRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.File, $i.Location, $i.Level, $i.Item, $i.Function, $i.Id, $i.Detail) }
        $row.Link = @{ Column = 1; Path = $i.Path }
        if ($i.Sheet) { $row.Link.Location = "'{0}'!{1}" -f $i.Sheet.Replace("'", "''"), $i.Cell }
        $issueRows.Add($row)
    }
    # 結果の列は「テストの結果」にする（この結果のブックをテスト仕様書のフォルダに保存しても、次からテスト仕様書として読まないように）
    $caseRows = New-Object System.Collections.Generic.List[object]
    foreach ($info in $caseInfo) {
        $p = $info.Case.Place
        $state = '－'
        if ($info.Units.Count -gt 0) { $state = 'あり' } elseif ($info.NotNeeded) { $state = '不要' } elseif (@('OK', 'NG') -contains $info.Kind) { $state = '無し' }
        $location = '{0} 行目' -f $p.Row
        if ($p.Sheet) { $location = '{0}!{1}{2}' -f $p.Sheet, (ConvertTo-ColumnName $p.Column), $p.Row }
        $images = 0
        foreach ($u in $info.Units) { $images += $u.Images.Count }
        $row = @{ Values = @($p.File, $location, $info.Function, $info.Id, ([string]$info.Case.Values['結果']).Trim(), $info.Evidence, $state, [double]$images, (Get-UnitList $info.Units 5)) }
        if ($info.Units.Count -gt 0) { $row.Link = Get-UnitLink $info.Units[0] 8 }
        else {
            $row.Link = @{ Column = 1; Path = $p.Path }
            if ($p.Sheet) { $row.Link.Location = "'{0}'!{1}{2}" -f $p.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $p.Column), $p.Row }
        }
        $caseRows.Add($row)
    }
    $unitRows = New-Object System.Collections.Generic.List[object]
    foreach ($u in $units) {
        if ($u.Cases.Count -eq 0 -and @($u.Candidates).Count -eq 0 -and $u.Images.Count -eq 0) { continue }
        $note = ''
        if (@($u.Candidates).Count -gt 1) { $note = '候補: ' + ((@($u.Candidates | ForEach-Object { '{0}（{1}）' -f $_.Id, $_.Function })) -join '・') }
        elseif ($u.Cases.Count -eq 0) { $note = 'どのケースか分からない' }
        $row = @{ Values = @($u.File, $(if ($u.Sheet) { $u.Location } else { '' }), $u.Kind, [double]$u.Images.Count, [double]$u.Texts, ((@($u.Cases | ForEach-Object { $_.Id })) -join '・'), $note) }
        $row.Link = Get-UnitLink $u 0
        $unitRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '指摘'; Columns = @(@{ Header = 'ファイル'; Width = 28 }, @{ Header = '場所'; Width = 18 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 24 }, @{ Header = '機能'; Width = 18 }, @{ Header = 'テストID'; Width = 10 }, @{ Header = '内容'; Width = 80 }); Rows = $issueRows },
            @{ Name = 'ケース別'; Columns = @(@{ Header = 'ファイル'; Width = 28 }, @{ Header = '場所'; Width = 16 }, @{ Header = '機能'; Width = 18 }, @{ Header = 'テストID'; Width = 10 }, @{ Header = 'テストの結果'; Width = 9 }, @{ Header = 'エビデンスの列'; Width = 16 }, @{ Header = 'エビデンス'; Width = 9 }, @{ Header = '画像'; Width = 6 }, @{ Header = 'エビデンスの場所'; Width = 60 }); Rows = $caseRows },
            @{ Name = 'エビデンス別'; Columns = @(@{ Header = 'ファイル'; Width = 30 }, @{ Header = '場所'; Width = 20 }, @{ Header = '種類'; Width = 9 }, @{ Header = '画像'; Width = 6 }, @{ Header = '文字のセル'; Width = 9 }, @{ Header = 'ケース'; Width = 20 }, @{ Header = '備考'; Width = 50 }); Rows = $unitRows }, (New-ReadLogSheet $readLog))
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

$doneCases = @($caseInfo | Where-Object { @('OK', 'NG') -contains $_.Kind })
$kinds = @(@('ブック', 'シート', '目印', '表の中', '画像', 'ファイル') | ForEach-Object {
        $k = $_
        $n = @($units | Where-Object { $_.Kind -eq $k -and ($_.Cases.Count -gt 0 -or @($_.Candidates).Count -gt 0 -or $_.Images.Count -gt 0) }).Count
        if ($n -gt 0) { '{0} {1}' -f $k, $n }
    })
$imageCount = 0
foreach ($u in $units) { $imageCount += $u.Images.Count }
Write-Host ''
Write-Host ('テストケース {0} 件（実施済み {1} 件）: エビデンスあり {2} 件・無し {3} 件・不要 {4} 件' -f $caseInfo.Count, $doneCases.Count,
    @($doneCases | Where-Object { $_.Units.Count -gt 0 }).Count, @($doneCases | Where-Object { $_.Units.Count -eq 0 -and -not $_.NotNeeded }).Count, @($doneCases | Where-Object { $_.Units.Count -eq 0 -and $_.NotNeeded }).Count)
$kindText = '無し'
if ($kinds.Count -gt 0) { $kindText = $kinds -join '・' }
Write-Host ('エビデンス: {0}（画像 {1} 枚）。どのケースか分からないもの {2}・決められないもの {3}' -f $kindText, $imageCount,
    @($units | Where-Object { $_.Cases.Count -eq 0 -and @($_.Candidates).Count -le 1 -and $_.Images.Count -gt 0 }).Count, @($units | Where-Object { $_.Cases.Count -eq 0 -and @($_.Candidates).Count -gt 1 }).Count)
Write-Host ('エラー {0} 件 / 警告 {1} 件 / 情報 {2} 件' -f @($issues | Where-Object Level -eq 'エラー').Count, @($issues | Where-Object Level -eq '警告').Count, @($issues | Where-Object Level -eq '情報').Count)
$byItem = @($issues | Group-Object Item | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })
if ($byItem.Count -gt 0) { Write-Host ('内訳: {0}' -f ($byItem -join '・')) }
Write-ReadLogSummary $readLog
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テストのエビデンス（画面のハードコピーなど）の抜け・空・使い回しを、テスト仕様書の結果と突き合わせて見つけます（Excel 不要）。
#
# .DESCRIPTION
#     テスト仕様書（Excel・CSV）から、テストID・結果・エビデンスの列を見出しの文字で探して読み、エビデンスと結び付けて次を確かめます。
#
#       エビデンスが無い（結果が OK・NG のケース。エビデンスの列が「-」「不要」なら要らないものとする）
#       エビデンスの列に書いたシート・ファイルが見つからない
#       エビデンスが空（画像も文字も無い）
#       実施していないケース（結果が空欄・保留）のエビデンス（結果の書き忘れ）
#       同じ画像の使い回し（違うケースのエビデンスに同じ画像がある。撮り直し忘れ・貼り間違い）
#       どのケースのエビデンスか分からない画像・決められないもの（情報）
#
#     エビデンスとケースは、次のように結び付けます（テストID は全角・半角、No.・■ などの印、数字の前の 0 を除いて比べます）。
#
#       シート名・ファイル名・フォルダ名がテストID（UT-001・No.1・SCR001_UT-001・UT-001_2（2 枚目）・受注一覧_No1 など）
#       ブックの名前がテストID なら、ブック全体がそのケースのエビデンス
#       シートの中の目印（行のいちばん左のセルの「■UT-001」「【No.1】」「テストID: UT-001」など）から次の目印までの画像
#       テスト仕様書の表の中に貼った画像（その行のケース）
#       テスト仕様書のエビデンスの列に書いたシート名・ファイル名（「ブック!シート」も可）
#
#     機能ごとに同じテストID を使っているときは、エビデンスのブック・フォルダ・シートの名前に機能の名前を入れてください。
#     見出しの候補は tools\test-quality\見出し.csv（Measure-TestQuality.ps1 と同じ）で足せます。
#
# .PARAMETER Path
#     テスト仕様書のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER EvidencePath
#     エビデンスのファイルまたはフォルダ（複数可）。省略するとテスト仕様書と同じ場所（テスト仕様書のブックのほかのシートなど）。
#     Excel（.xlsx / .xlsm）はシートごとに画像と文字を、画像（.png .jpg など）は中身を見ます。.xls・.pdf などはファイルの名前だけを見ます。
#
# .PARAMETER Headers
#     見出しの候補の CSV。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb のテスト仕様書は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘・ケース別・エビデンス別）。
#
# .EXAMPLE
#     .\Test-TestEvidence.ps1 D:\テスト仕様書 -EvidencePath D:\エビデンス -OutXlsx D:\エビデンスのチェック.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     エビデンスの中身（正しい画面か）は見ません。結び付け方を決めて、最後はレビューで確かめてください。
