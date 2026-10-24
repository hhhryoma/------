[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $Baseline,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'ScreenNumbers.Issue' -DefaultDisplayPropertySet File, Location, Level, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# ---------------------------------------------------------------- 番号の読み方


# 1 文字で番号を表す文字（先頭の文字の番号・それが表す番号・文字の数）
$numberChars = @(
    @(0x2460, 1, 20),   # ①〜⑳
    @(0x3251, 21, 15),  # ㉑〜㉟
    @(0x32B1, 36, 15),  # ㊱〜㊿
    @(0x24EA, 0, 1),    # ⓪
    @(0x2776, 1, 10),   # ❶〜❿
    @(0x24EB, 11, 10),  # ⓫〜⓴
    @(0x2780, 1, 10),   # ➀〜➉
    @(0x278A, 1, 10),   # ➊〜➓
    @(0x24F5, 1, 10),   # ⓵〜⓾
    @(0x24FF, 0, 1),    # ⓿
    @(0x2474, 1, 20),   # ⑴〜⒇
    @(0x2488, 1, 20)    # ⒈〜⒛
)

function ConvertFrom-NumberText {
    #     番号だけを書いた文字（⑤・⑤⑥・5・(5)・【5】・No.5・#5 など）の番号。番号のほかに文字があれば何も返さない
    param([string] $Text)
    $t = ([string]$Text).Trim()
    if (-not $t) { return }
    $numbers = New-Object System.Collections.Generic.List[int]
    foreach ($ch in $t.ToCharArray()) {
        if ([char]::IsWhiteSpace($ch)) { continue }
        $code = [int]$ch
        $n = -1
        foreach ($r in $numberChars) { if ($code -ge $r[0] -and $code -lt $r[0] + $r[2]) { $n = $r[1] + $code - $r[0]; break } }
        if ($n -lt 0) { $numbers.Clear(); break }
        $numbers.Add($n)
    }
    if ($numbers.Count -gt 0) { return $numbers.ToArray() }
    $m = [regex]::Match($t.Normalize($nfkc), '^(?:[(\[【<]\s*(?<n>\d{1,3})\s*[)\]】>]|(?:No\.?|#)?\s*(?<n>\d{1,3}))$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($m.Success) { return [int]$m.Groups['n'].Value }
}

function Format-Numbers {
    #     番号の一覧の表示（1〜12・15 のように）
    param([int[]] $Numbers)
    $sorted = @($Numbers | Sort-Object -Unique)
    $parts = New-Object System.Collections.Generic.List[string]
    $i = 0
    while ($i -lt $sorted.Count) {
        $j = $i
        while ($j + 1 -lt $sorted.Count -and $sorted[$j + 1] -eq $sorted[$j] + 1) { $j++ }
        if ($j -gt $i) { $parts.Add(('{0}〜{1}' -f $sorted[$i], $sorted[$j])) } else { $parts.Add([string]$sorted[$i]) }
        $i = $j + 1
    }
    return $parts -join '・'
}

function ConvertTo-ScreenKey {
    #     画面ID・画面名・シート名を比べるための形（全角・半角と大文字・小文字をそろえ、空白を除く）
    param([string] $Text)
    return (ConvertTo-Label $Text).ToUpperInvariant()
}

function Get-OneLine {
    param([string] $Text)
    return (([string]$Text).Trim() -replace '\s*\r?\n\s*', ' ')
}

# ---------------------------------------------------------------- 設計書を読む


function Get-ExcelFilePaths {
    #     フォルダ・ファイルの中の .xlsx / .xlsm（図形・画像を読むブック）
    param([string[]] $Paths)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        $items = @(Get-PathItem $p)
        if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
        foreach ($item in $items) {
            if ($item.PSIsContainer) { $files = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Sort-Object FullName) }
            else { $files = @($item) }
            foreach ($f in $files) {
                if (@('.xlsx', '.xlsm') -notcontains $f.Extension.ToLowerInvariant() -or $f.Name.StartsWith('~$') -or $list.Contains($f.FullName)) { continue }
                if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                $list.Add($f.FullName)
            }
        }
    }
    return , $list
}

function Read-ScreenLayouts {
    #     設計書を読み、スクリーンショット（画像）の上に番号を重ねたシートと、画面項目定義を集める。
    #     戻り値: @{ Layouts; Items; Books（読んだ .xlsx / .xlsm の数） }
    #       Layout: File・Path・Sheet・SheetIndex・Ids・Names（シートかブックの「画面ID」「画面名」。比べる形）・SheetKey・
    #               SheetCut（31 文字で切れたシート名か）・Pictures・Markers（重ねた番号）
    #       Marker: Text・Name（図形の名前）・Numbers・Picture（Pictures の何番目か）・Cell・RelX RelY（画像の中での中心の位置の割合）・Size
    #       Items: 画面項目定義の行（Read-ScreenDesign）と、比べる形の IdKey・NameKey・SheetKey・Number（No の番号。無ければ -1）
    param([string[]] $Paths)
    $searchArgs = @{ Pattern = '\S'; Path = $Paths; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
    $books = ConvertTo-DesignBooks $hits
    $bookByPath = New-Object 'System.Collections.Generic.Dictionary[string,object]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($k in $books.Keys) { $bookByPath[$k] = $books[$k] }

    # いくつものブックにある同じ名前のシート（「画面レイアウト」など）は、画面の名前とみなさない
    $sheetCount = @{}
    foreach ($b in $books.Values) { foreach ($s in $b.Sheets.Keys) { $k = ConvertTo-ScreenKey $s; $sheetCount[$k] = 1 + [int]$sheetCount[$k] } }
    $sheetKeyOf = { param([string] $Name) $k = ConvertTo-ScreenKey $Name; if ([int]$sheetCount[$k] -gt 1) { '' } else { $k } }

    # 「画面ID」「画面名」の欄（シートに無ければ、ブックの中で 1 つだけの値）
    $labels = @{}
    foreach ($bp in $books.Keys) {
        $b = $books[$bp]
        $perSheet = @{}
        $bookIds = New-Object 'System.Collections.Generic.HashSet[string]'
        $bookNames = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($s in $b.Sheets.Keys) {
            $id = ConvertTo-ScreenKey (Find-LabelValue $b.Sheets[$s] '画面ID')
            $nm = ConvertTo-ScreenKey (Find-LabelValue $b.Sheets[$s] '画面名')
            $perSheet[$s] = @($id, $nm)
            if ($id) { [void]$bookIds.Add($id) }
            if ($nm) { [void]$bookNames.Add($nm) }
        }
        $bookId = ''; if ($bookIds.Count -eq 1) { $bookId = @($bookIds)[0] }
        $bookName = ''; if ($bookNames.Count -eq 1) { $bookName = @($bookNames)[0] }
        $labels[$bp] = @{ Sheets = $perSheet; Id = $bookId; Name = $bookName }
    }

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($it in (Read-ScreenDesign $books -AnySheet).Items) {
        $number = -1
        $ns = @(ConvertFrom-NumberText $it.No)
        if ($ns.Count -eq 1) { $number = $ns[0] }
        $items.Add([pscustomobject]@{
                Item = $it; Path = $it.Path; Sheet = $it.Sheet; Number = $number
                IdKey = (ConvertTo-ScreenKey $it.Screen); NameKey = (ConvertTo-ScreenKey $it.ScreenName); SheetKey = (& $sheetKeyOf $it.Sheet)
            })
    }

    $layouts = New-Object System.Collections.Generic.List[object]
    $files = Get-ExcelFilePaths $Paths
    foreach ($file in $files) {
        $drawings = $null
        try { $drawings = Read-BookDrawings $file -Bounds }
        catch { Write-Warning ("{0}: 図形・画像を読めませんでした: {1}" -f [System.IO.Path]::GetFileName($file), $_.Exception.Message) }
        if (-not $drawings) { continue }
        $book = $null
        if ($bookByPath.ContainsKey($file)) { $book = $bookByPath[$file] }
        $index = -1
        foreach ($sheet in $drawings.Keys) {
            $index++
            $all = $drawings[$sheet]
            # 非表示の図形と、非表示のグループの中の図形は読まない
            $hiddenGroups = @($all | Where-Object { $_.Hidden -and $_.Kind -eq 'grpSp' } | ForEach-Object { $_.Path + '/' })
            $shown = @($all | Where-Object { $it = $_; -not $it.Hidden -and $it.Bounds -and @($hiddenGroups | Where-Object { $it.Path.StartsWith($_) }).Count -eq 0 })
            $pictures = @($shown | Where-Object { $_.Kind -eq 'pic' -and $_.Bounds.W -gt 0 -and $_.Bounds.H -gt 0 })
            if ($pictures.Count -eq 0) { continue }
            $markers = New-Object System.Collections.Generic.List[object]
            foreach ($it in $shown) {
                if ($it.Kind -ne 'sp' -or -not $it.Text) { continue }
                $numbers = @(ConvertFrom-NumberText $it.Text)
                if ($numbers.Count -eq 0) { continue }
                # スクリーンショットの外の図形（凡例など）は数えない。中心が画像の上にあるものだけ
                $b = $it.Bounds
                $cx = $b.X + $b.W / 2; $cy = $b.Y + $b.H / 2
                $pi = -1
                for ($i = 0; $i -lt $pictures.Count; $i++) {
                    $p = $pictures[$i].Bounds
                    if ($cx -ge $p.X -and $cx -le $p.X + $p.W -and $cy -ge $p.Y -and $cy -le $p.Y + $p.H) { $pi = $i; break }
                }
                if ($pi -lt 0) { continue }
                $p = $pictures[$pi].Bounds
                $markers.Add([pscustomobject]@{
                        Text = (Get-OneLine $it.Text); Name = $it.Name; Numbers = [int[]]$numbers; Picture = $pi
                        Cell = (Get-SheetCell $b.Geometry $cx $cy); RelX = ($cx - $p.X) / $p.W; RelY = ($cy - $p.Y) / $p.H; Size = [math]::Max($b.W, $b.H)
                        Items = (New-Object System.Collections.Generic.List[object]); OldItems = (New-Object System.Collections.Generic.List[object])
                        Results = (New-Object System.Collections.Generic.List[string])
                    })
            }
            if ($markers.Count -eq 0) { continue }
            $ids = New-Object 'System.Collections.Generic.HashSet[string]'
            $names = New-Object 'System.Collections.Generic.HashSet[string]'
            $lb = $null
            foreach ($k in $labels.Keys) { if ([string]::Equals($k, $file, [System.StringComparison]::OrdinalIgnoreCase)) { $lb = $labels[$k]; break } }
            if ($lb) {
                $own = $lb.Sheets[$sheet]
                if ($own -and ($own[0] -or $own[1])) { if ($own[0]) { [void]$ids.Add($own[0]) }; if ($own[1]) { [void]$names.Add($own[1]) } }
                else { if ($lb.Id) { [void]$ids.Add($lb.Id) }; if ($lb.Name) { [void]$names.Add($lb.Name) } }
            }
            $layouts.Add([pscustomobject]@{
                    File = [System.IO.Path]::GetFileName($file); Path = $file; Sheet = $sheet; SheetIndex = $index
                    Ids = $ids; Names = $names; SheetKey = (& $sheetKeyOf $sheet); SheetCut = ($sheet.Length -ge 31)
                    Pictures = $pictures; Markers = $markers; Map = $null; Source = ''
                })
        }
    }
    return [pscustomobject]@{ Layouts = $layouts; Items = $items; Books = $files.Count }
}

# ---------------------------------------------------------------- 画面項目定義と対応付ける


function Get-ScreenIdentity {
    #     画面項目定義の行がどの画面のものか（画面ID、無ければ画面名、無ければブックとシート）
    param($Entry)
    if ($Entry.IdKey) { return 'ID:' + $Entry.IdKey }
    if ($Entry.NameKey) { return 'NAME:' + $Entry.NameKey }
    return 'SHEET:{0}|{1}' -f $Entry.Path, $Entry.Sheet
}

function Test-LayoutKey {
    #     画面項目定義の行（Entry）が、画面（Layout）の画面ID・画面名・シート名と合うか
    param($Layout, $Entry)
    if ($Entry.IdKey -and $Layout.Ids.Contains($Entry.IdKey)) { return $true }
    if ($Entry.NameKey -and $Layout.Names.Contains($Entry.NameKey)) { return $true }
    $sk = $Layout.SheetKey
    if (-not $sk) { return $false }
    if ($Entry.IdKey -eq $sk -or $Entry.NameKey -eq $sk -or $Entry.SheetKey -eq $sk) { return $true }
    # シート名は 31 文字までなので、長い画面名は途中で切れている
    return ($Layout.SheetCut -and $Entry.NameKey -and $Entry.NameKey.StartsWith($sk))
}

function Find-ScreenItems {
    #     画面（Layout）の画面項目定義の行を、同じシート → 同じブックの同じ画面 → ほかのブックの同じ画面 →
    #     （画面が 1 つだけの）同じブック の順に探す。戻り値: @{ Items; Problem（見つからない・候補が複数のときの説明） }
    param($Layout, $Items)
    $sameBook = @($Items | Where-Object { [string]::Equals($_.Path, $Layout.Path, [System.StringComparison]::OrdinalIgnoreCase) })
    $found = @($sameBook | Where-Object { $_.Sheet -eq $Layout.Sheet })
    if ($found.Count -gt 0) { return @{ Items = $found } }
    $found = @($sameBook | Where-Object { $e = $_; ($e.IdKey -and $Layout.Ids.Contains($e.IdKey)) -or ($e.NameKey -and $Layout.Names.Contains($e.NameKey)) })
    if ($found.Count -gt 0) { return @{ Items = $found } }
    $others = @($Items | Where-Object { -not [string]::Equals($_.Path, $Layout.Path, [System.StringComparison]::OrdinalIgnoreCase) -and (Test-LayoutKey $Layout $_) })
    $screens = @($others | ForEach-Object { Get-ScreenIdentity $_ } | Select-Object -Unique)
    if ($screens.Count -eq 1) { return @{ Items = $others } }
    if ($screens.Count -gt 1) {
        $where = @($others | ForEach-Object { '{0}（{1}）' -f $_.Item.File, $_.Sheet } | Select-Object -Unique | Select-Object -First 3) -join '・'
        return @{ Items = @(); Problem = ('この画面の画面項目定義の候補が複数あるため（{0}）、重ねた番号を確かめていません' -f $where) }
    }
    $bookScreens = @($sameBook | ForEach-Object { Get-ScreenIdentity $_ } | Select-Object -Unique)
    if ($bookScreens.Count -eq 1) { return @{ Items = $sameBook } }
    if ($bookScreens.Count -gt 1) { return @{ Items = @(); Problem = '同じブックに画面項目定義が複数あり、この画面のものが決まらないため、重ねた番号を確かめていません（シートに「画面ID」「画面名」の欄を書くと対応付けます）' } }
    return @{ Items = @(); Problem = 'この画面の画面項目定義（No・項目名・桁数の表）が見つからないため、重ねた番号を確かめていません' }
}

function Set-LayoutItems {
    #     画面ごとに、画面項目定義の No の番号 → 行（Map）を作る（同じ No が 2 行あれば先の行）
    param($Read)
    foreach ($layout in $Read.Layouts) {
        $r = Find-ScreenItems $layout $Read.Items
        $layout.Source = $r.Problem
        if ($r.Items.Count -eq 0) { continue }
        $map = @{}
        foreach ($e in $r.Items) { if ($e.Number -ge 0 -and -not $map.ContainsKey($e.Number)) { $map[$e.Number] = $e.Item } }
        if ($map.Count -eq 0) {
            $layout.Source = '画面項目定義（{0} の {1}）の No の列に番号が無いため、重ねた番号を確かめていません' -f $r.Items[0].Item.File, $r.Items[0].Sheet
            continue
        }
        $layout.Map = $map
    }
}

# ---------------------------------------------------------------- 確かめる


$script:Issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    param([ValidateSet('エラー', '警告', '情報')] [string] $Level, [string] $Item, [string] $Detail, $Layout, [string] $Cell)
    $issue = [pscustomobject]@{
        File = $Layout.File; Location = ('{0}!{1}' -f $Layout.Sheet, $Cell); Level = $Level; Item = $Item; Detail = $Detail
        Sheet = $Layout.Sheet; Cell = $Cell; Path = $Layout.Path; SheetIndex = $Layout.SheetIndex; Row = 0; Col = 0
    }
    $m = [regex]::Match($Cell, '^([A-Z]+)(\d+)$')
    if ($m.Success) { $issue.Col = Get-ColumnIndex $m.Groups[1].Value; $issue.Row = [int]$m.Groups[2].Value }
    $issue.PSObject.TypeNames.Insert(0, 'ScreenNumbers.Issue')
    $script:Issues.Add($issue)
}

function Get-MarkerLabel {
    #     指摘に書く図形の示し方（⑤（楕円 5））
    param($Marker)
    if ($Marker.Name) { return '{0}（{1}）' -f $Marker.Text, $Marker.Name }
    return $Marker.Text
}

$current = Read-ScreenLayouts $Path
Set-LayoutItems $current

foreach ($layout in $current.Layouts) {
    if (-not $layout.Map) {
        Add-Issue '情報' '定義なし' ('{0}（重ねた番号 {1} 個）' -f $layout.Source, $layout.Markers.Count) $layout $layout.Markers[0].Cell
        foreach ($m in $layout.Markers) { $m.Results.Add('未確認') }
        continue
    }
    $range = Format-Numbers ([int[]]@($layout.Map.Keys))
    foreach ($m in $layout.Markers) {
        foreach ($n in $m.Numbers) { $m.Items.Add($layout.Map[$n]) }
        $missing = @($m.Numbers | Where-Object { -not $layout.Map.ContainsKey($_) })
        if ($missing.Count -eq 0) { continue }
        $m.Results.Add('定義に無い番号')
        $what = ''
        if ($m.Numbers.Count -gt 1) { $what = '（{0}）' -f (($missing | ForEach-Object { "No $_" }) -join '・') }
        Add-Issue 'エラー' '定義に無い番号' ('重ねた番号 {0}{1}は、画面項目定義に無い番号です（画面項目定義の No: {2}）' -f (Get-MarkerLabel $m), $what, $range) $layout $m.Cell
    }
    # 同じ画像の上の同じ番号（わざと 2 か所に付けることもあるので警告）
    $groups = [ordered]@{}
    foreach ($m in $layout.Markers) {
        foreach ($n in ($m.Numbers | Select-Object -Unique)) {
            $k = '{0}|{1}' -f $m.Picture, $n
            if (-not $groups.Contains($k)) { $groups[$k] = New-Object System.Collections.Generic.List[object] }
            $groups[$k].Add($m)
        }
    }
    foreach ($k in $groups.Keys) {
        $g = $groups[$k]
        if ($g.Count -lt 2) { continue }
        foreach ($m in $g) { if (-not $m.Results.Contains('番号の重複')) { $m.Results.Add('番号の重複') } }
        $n = [int]($k.Split('|')[1])
        $places = @($g | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Cell }) -join '、'
        Add-Issue '警告' '番号の重複' ('No {0} の番号が {1} か所にあります（{2}）' -f $n, $g.Count, $places) $layout $g[1].Cell
    }
}

# ---------------------------------------------------------------- 前の版と比べる（-Baseline）


function Find-OldLayout {
    #     前の版の同じ画面（版数を除いたファイル名とシート名 → 画面ID・画面名 → シート名 の順）
    param($Layout, $OldLayouts)
    $name = Get-VersionlessName ([System.IO.Path]::GetFileName($Layout.Path))
    $found = @($OldLayouts | Where-Object { $_.Sheet -eq $Layout.Sheet -and (Get-VersionlessName ([System.IO.Path]::GetFileName($_.Path))) -eq $name })
    if ($found.Count -eq 1) { return $found[0] }
    $found = @($OldLayouts | Where-Object { $_.Ids.Overlaps($Layout.Ids) -or $_.Names.Overlaps($Layout.Names) })
    if ($found.Count -eq 1) { return $found[0] }
    if ($Layout.SheetKey) {
        $found = @($OldLayouts | Where-Object { $_.SheetKey -eq $Layout.SheetKey })
        if ($found.Count -eq 1) { return $found[0] }
    }
    return $null
}

function Find-OldPicture {
    #     前の版の、同じスクリーンショット（同じ図形・同じ画像、または同じ順番）の番号。無ければ -1
    param($OldLayout, $Layout, [int] $Index)
    $p = $Layout.Pictures[$Index]
    for ($i = 0; $i -lt $OldLayout.Pictures.Count; $i++) { if ($OldLayout.Pictures[$i].Key -eq $p.Key) { return $i } }
    if ($p.Hash) { for ($i = 0; $i -lt $OldLayout.Pictures.Count; $i++) { if ($OldLayout.Pictures[$i].Hash -eq $p.Hash) { return $i } } }
    if ($OldLayout.Pictures.Count -eq $Layout.Pictures.Count) { return $Index }
    return -1
}

$hasBaseline = @($Baseline | Where-Object { $_ }).Count -gt 0
if ($hasBaseline) {
    $old = Read-ScreenLayouts $Baseline
    Set-LayoutItems $old
    foreach ($layout in $current.Layouts) {
        if (-not $layout.Map) { continue }
        $prev = Find-OldLayout $layout $old.Layouts
        if (-not $prev) {
            Add-Issue '情報' '前の版なし' '前の版に、この画面（番号を重ねたスクリーンショット）が見つからないため、番号の指す項目が変わったかは確かめていません' $layout $layout.Markers[0].Cell
            continue
        }
        if (-not $prev.Map) { continue }
        # 画像の中の同じ場所に重ねた番号どうしを比べる（動かした番号は、付け直したものとみなして比べない）
        $used = New-Object 'System.Collections.Generic.HashSet[int]'
        foreach ($m in $layout.Markers) {
            $op = Find-OldPicture $prev $layout $m.Picture
            if ($op -lt 0) { continue }
            $pic = $layout.Pictures[$m.Picture].Bounds
            $bi = -1; $bestD = [double]::MaxValue
            for ($j = 0; $j -lt $prev.Markers.Count; $j++) {
                $o = $prev.Markers[$j]
                if ($o.Picture -ne $op -or $used.Contains($j)) { continue }
                $d = [math]::Sqrt([math]::Pow(($m.RelX - $o.RelX) * $pic.W, 2) + [math]::Pow(($m.RelY - $o.RelY) * $pic.H, 2))
                if ($d -le [math]::Max(4, [math]::Min($m.Size, $o.Size) / 2) -and $d -lt $bestD) { $bi = $j; $bestD = $d }
            }
            if ($bi -lt 0) { continue }
            [void]$used.Add($bi)
            $best = $prev.Markers[$bi]
            foreach ($n in $best.Numbers) { $m.OldItems.Add($prev.Map[$n]) }
            if ($best.Numbers.Count -ne $m.Numbers.Count) { continue }
            for ($i = 0; $i -lt $m.Numbers.Count; $i++) {
                $was = $prev.Map[$best.Numbers[$i]]
                $now = $layout.Map[$m.Numbers[$i]]
                if (-not $was -or -not $now -or (ConvertTo-Label $was.Item) -eq (ConvertTo-Label $now.Item)) { continue }
                if (-not $m.Results.Contains('指す項目の変化')) { $m.Results.Add('指す項目の変化') }
                $detail = '重ねた番号 {0}が指す項目が、前の版の「{1}」から「{2}」に変わりました' -f (Get-MarkerLabel $m), (Get-OneLine $was.Item), (Get-OneLine $now.Item)
                if ($best.Numbers[$i] -ne $m.Numbers[$i]) { $detail += '（前の版の番号は {0}）' -f $best.Text }
                $moved = @($layout.Map.Keys | Where-Object { (ConvertTo-Label $layout.Map[$_].Item) -eq (ConvertTo-Label $was.Item) } | Sort-Object)
                if ($moved.Count -gt 0) { $detail += '。「{0}」は、いまは No {1} です' -f (Get-OneLine $was.Item), $moved[0] }
                Add-Issue '警告' '指す項目の変化' ($detail + '。番号が古いままでないか確かめてください') $layout $m.Cell
            }
        }
    }
}

# ---------------------------------------------------------------- 出力


$levelOrder = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
$sorted = @($script:Issues | Sort-Object @{ e = { $_.File } }, @{ e = { $_.SheetIndex } }, @{ e = { $_.Row } }, @{ e = { $_.Col } }, @{ e = { $levelOrder[$_.Level] } })
$sorted

$markerCount = 0
foreach ($l in $current.Layouts) { $markerCount += $l.Markers.Count }

function Format-ItemNames {
    #     番号の一覧に書く、番号が指す項目名（無い番号は「（無い番号）」）
    param($Items)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($it in $Items) { if ($it) { $names.Add((Get-OneLine $it.Item)) } else { $names.Add('（無い番号）') } }
    return $names -join '・'
}

if ($OutXlsx) {
    $link = { param($l, $cell) @{ Path = $l.Path; Location = ("'{0}'!{1}" -f $l.Sheet.Replace("'", "''"), $cell) } }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $rows.Add(@{ Values = @($i.File, $i.Level, $i.Item, $i.Location, $i.Detail); Link = @{ Column = 3; Path = $i.Path; Location = ("'{0}'!{1}" -f $i.Sheet.Replace("'", "''"), $i.Cell) } })
    }
    $columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '場所'; Width = 22 }, @{ Header = '番号'; Width = 8 }, @{ Header = '図形'; Width = 16 }, @{ Header = '項目名（画面項目定義）'; Width = 28 })
    if ($hasBaseline) { $columns += @{ Header = '前の版の項目名'; Width = 28 } }
    $columns += @{ Header = '結果'; Width = 24 }
    $markerRows = New-Object System.Collections.Generic.List[object]
    foreach ($l in $current.Layouts) {
        foreach ($m in $l.Markers) {
            $values = @($l.File, ('{0}!{1}' -f $l.Sheet, $m.Cell), $m.Text, $m.Name, (Format-ItemNames $m.Items))
            if ($hasBaseline) { $values += (Format-ItemNames $m.OldItems) }
            $result = $m.Results -join '・'
            if (-not $result) { $result = 'OK' }
            $values += $result
            $l2 = & $link $l $m.Cell
            $markerRows.Add(@{ Values = $values; Link = @{ Column = 1; Path = $l2.Path; Location = $l2.Location } })
        }
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '指摘一覧'; Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 16 }, @{ Header = '場所'; Width = 22 }, @{ Header = '内容'; Width = 90 }); Rows = $rows },
            @{ Name = '番号の一覧'; Columns = $columns; Rows = $markerRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
if ($current.Layouts.Count -eq 0) {
    Write-Host ('{0} ファイルを読みましたが、スクリーンショット（画像）の上に重ねた番号が見つかりませんでした' -f $current.Books)
}
else {
    Write-Host ('{0} ファイルを読みました（番号を重ねたスクリーンショットのあるシート {1}・重ねた番号 {2} 個）' -f $current.Books, $current.Layouts.Count, $markerCount)
}
Write-Host ('エラー {0} 件 / 警告 {1} 件 / 情報 {2} 件' -f @($sorted | Where-Object Level -eq 'エラー').Count, @($sorted | Where-Object Level -eq '警告').Count, @($sorted | Where-Object Level -eq '情報').Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     画面レイアウトのスクリーンショットに重ねた番号（①・No.5 など）を、画面項目定義の No と突き合わせます（Excel 不要。ファイルは変更しません）。
#
# .DESCRIPTION
#     設計書の .xlsx / .xlsm を読み、スクリーンショット（画像）の上に重ねた図形のうち、番号だけを書いたもの
#     （①〜㊿・❶・⑤⑥・丸や四角の中の 5・(5)・No.5 など）を集めて、その画面の画面項目定義の No と比べます。
#     説明の文も書いた吹き出し・文字の無い枠・スクリーンショットの外の図形は読みません。グループの中の図形は読みます。
#
#       定義に無い番号   重ねた番号が、画面項目定義の No に無い（項目を消した・番号を振り直したのに、番号が古いまま。エラー）
#       番号の重複       同じスクリーンショットの上に、同じ番号が 2 つ以上ある（わざと付けることもあるので警告）
#       指す項目の変化   -Baseline の前の版と比べて、同じ場所に重ねた番号が指す項目が変わった（途中に項目を足して No がずれたのに、
#                        番号を直していないなど。重ねた図形の文字は変わらないので、いちばん気づきにくい誤り。警告）
#       定義なし         画面項目定義が見つからない・決まらないため、確かめていない（情報）
#       前の版なし       -Baseline に同じ画面が見つからないため、指す項目の変化を確かめていない（情報）
#
#     画面項目定義は、No・項目名・桁数の見出しのある表です（テスト仕様書の下書きなどと同じ読み方）。
#     スクリーンショットのあるシートごとに、次の順で探します。
#       1. 同じシートの表
#       2. 同じブックの、「画面ID」か「画面名」の欄が同じシートの表（欄がシートに無ければ、ブックの中で 1 つだけの値を使う）
#       3. ほかのブックの、画面ID・画面名が同じ表（シート名が画面名・画面ID と同じものも。31 文字で切れたシート名は前の部分で比べる）
#       4. 同じブックに 1 つの画面の表しか無ければ、その表
#     画面項目定義がスクリーンショットと別のブックにあるときは、両方を -Path に指定してください。
#
#     前の版（-Baseline）とは、版数を除いたファイル名とシート名、画面ID・画面名、シート名の順に画面を対応付け、
#     スクリーンショットの中の同じ場所（図形の大きさの半分以内）に重ねた番号どうしを比べます。番号を動かしたものは、付け直したものとみなして比べません。
#
#     結果は 1 指摘 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧（指摘一覧と、重ねた番号ごとに指す項目を並べた番号の一覧）に保存できます。
#
# .PARAMETER Path
#     画面設計書のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#     画面項目定義が別のブックにあるときは、そのブックも含めて指定してください。
#
# .PARAMETER Baseline
#     前の版の画面設計書のフォルダまたはファイル（複数可）。指定すると、同じ場所に重ねた番号が指す項目が変わったものを知らせます。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の画面項目定義は読まなくなります。重ねた番号は .xlsx / .xlsm だけから読みます）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘一覧と番号の一覧。場所をクリックすると該当セルが開きます）。
#
# .EXAMPLE
#     .\Test-ScreenNumbers.ps1 D:\設計書\画面設計書
#
# .EXAMPLE
#     # 前の版と比べて、項目を足して番号がずれたのに直していない丸数字も見つける
#     .\Test-ScreenNumbers.ps1 D:\設計書\v1.1 -Baseline D:\設計書\v1.0 -OutXlsx D:\画面の番号.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     図形の位置は、列の幅と行の高さから求めます。スクリーンショットの端にかかる番号は、中心が画像の上にあるかで決めます。
