[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [Parameter(Mandatory = $true)]
    [string[]] $Wireframe,

    [switch] $Record,

    [string[]] $Screen,

    [string] $RecordFile,

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
if ($Screen -and -not $Record) { throw '-Screen は -Record と一緒に指定してください（記録する画面を選びます）' }

Update-TypeData -TypeName 'Screenshots.Screen' -DefaultDisplayPropertySet Screen, Status, Location, Detail -Force -WhatIf:$false

$hasher = [System.Security.Cryptography.SHA256]::Create()

function Get-ShortHash {
    #     文字の要約（SHA-256 の先頭 16 桁）
    param([string] $Text)
    return ([System.BitConverter]::ToString($hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').Substring(0, 16).ToLowerInvariant()
}

function ConvertTo-ScreenKey {
    #     画面名・シート名を比べるための形（全角・半角と大文字・小文字をそろえ、空白を除く）
    param([string] $Text)
    return (ConvertTo-Label $Text).ToUpperInvariant()
}

function Get-ExcelFilePaths {
    #     フォルダ・ファイルの中の .xlsx / .xlsm（.xls・.xlsb は中を読めないので知らせて除く）
    param([string[]] $Paths, [string] $What)
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        $items = @(Get-PathItem $p)
        if ($items.Count -eq 0) { Write-Warning "${What}が見つかりません: $p"; continue }
        foreach ($item in $items) {
            if ($item.PSIsContainer) { $files = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue | Sort-Object FullName) }
            else { $files = @($item) }
            foreach ($f in $files) {
                $ext = $f.Extension.ToLowerInvariant()
                if ($f.Name.StartsWith('~$') -or $list.Contains($f.FullName)) { continue }
                if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
                if (@('.xls', '.xlsb') -contains $ext) { Write-Warning ("{0}: {1} は読めません（.xlsx に保存し直すと確かめられます）" -f $f.Name, $ext); continue }
                if (@('.xlsx', '.xlsm') -contains $ext) { $list.Add($f.FullName) }
            }
        }
    }
    return , $list
}

# ---------------------------------------------------------------- 元のワイヤーフレームを読む


function Read-StyleKeys {
    #     セルの書式の番号 → 書式の中身（表示形式・フォント・塗りつぶし・罫線・配置）の文字（書式の番号が振り直されても同じになる）
    param($Package, [string] $StylesPart)
    $keys = @{}
    $doc = $null
    if ($StylesPart) { $doc = Read-PartXml -Package $Package -PartName $StylesPart }
    if (-not $doc) { return $keys }
    $parts = @{}
    foreach ($kind in @(@('fonts', 'font'), @('fills', 'fill'), @('borders', 'border'))) {
        $parts[$kind[1]] = @($doc.SelectNodes("//*[local-name()='$($kind[0])']/*[local-name()='$($kind[1])']") | ForEach-Object { $_.OuterXml })
    }
    $formats = @{}
    foreach ($n in $doc.SelectNodes("//*[local-name()='numFmts']/*[local-name()='numFmt']")) { $formats[$n.GetAttribute('numFmtId')] = $n.GetAttribute('formatCode') }
    $pick = { param([string] $Kind, [string] $Id) $i = 0; if ([int]::TryParse($Id, [ref]$i) -and $i -lt $parts[$Kind].Count) { $parts[$Kind][$i] } else { '' } }
    $index = 0
    foreach ($xf in $doc.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
        $nf = $xf.GetAttribute('numFmtId')
        if ($formats.ContainsKey($nf)) { $nf = $formats[$nf] }
        $keys[$index] = '{0}|{1}|{2}|{3}|{4}' -f $nf, (& $pick 'font' $xf.GetAttribute('fontId')), (& $pick 'fill' $xf.GetAttribute('fillId')), (& $pick 'border' $xf.GetAttribute('borderId')), $xf.InnerXml
        $index++
    }
    return $keys
}

function Read-SharedStringList {
    #     共有文字列の一覧（番号の順）
    param($Package, [string] $Part)
    $list = New-Object System.Collections.Generic.List[string]
    $doc = $null
    if ($Part) { $doc = Read-PartXml -Package $Package -PartName $Part }
    if ($doc) { foreach ($si in $doc.SelectNodes("//*[local-name()='si']")) { $list.Add(-join @($si.SelectNodes(".//*[local-name()='t' and not(ancestor::*[local-name()='rPh'])]") | ForEach-Object { $_.InnerText })) } }
    return , $list
}

function Get-WireframeKey {
    #     元のシートの、印刷範囲の中の見た目の要約（セルの値と書式・結合・列の幅・行の高さ・図形と画像の位置・大きさ・文字・書式）。
    #     印刷範囲が無ければシート全体。範囲の外のメモを直しても、要約は変わらない
    param($Package, [string] $SheetPart, $Areas, $Styles, $Strings, [int] $Mdw, [hashtable] $Images)
    $doc = Read-PartXml -Package $Package -PartName $SheetPart
    if (-not $doc) { return '' }
    $inArea = {
        param([int] $Col, [int] $Row)
        if ($Areas.Count -eq 0) { return $true }
        foreach ($a in $Areas) { if ($Col -ge $a.C0 -and $Col -le $a.C1 -and $Row -ge $a.R0 -and $Row -le $a.R1) { return $true } }
        return $false
    }
    $lines = New-Object System.Collections.Generic.List[string]
    $plain = ''
    if ($Styles.ContainsKey(0)) { $plain = $Styles[0] }
    foreach ($c in $doc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']/*[local-name()='c'][@r]")) {
        $ref = $c.GetAttribute('r')
        $m = [regex]::Match($ref, '^([A-Z]+)(\d+)$')
        if (-not $m.Success -or -not (& $inArea (Get-ColumnIndex $m.Groups[1].Value) ([int]$m.Groups[2].Value))) { continue }
        $value = ''
        $v = $c.SelectSingleNode("*[local-name()='v']")
        switch ($c.GetAttribute('t')) {
            's' { if ($v) { $i = 0; if ([int]::TryParse($v.InnerText, [ref]$i) -and $i -lt $Strings.Count) { $value = $Strings[$i] } } }
            'inlineStr' { $value = $c.InnerText }
            default { if ($v) { $value = $v.InnerText } }
        }
        $style = ''
        $s = 0
        if ([int]::TryParse($c.GetAttribute('s'), [ref]$s) -and $Styles.ContainsKey($s)) { $style = $Styles[$s] }
        if (-not $value -and (-not $style -or $style -eq $plain)) { continue }
        $lines.Add(('C|{0}|{1}|{2}' -f $ref, $value, $style))
    }
    foreach ($mc in $doc.SelectNodes("//*[local-name()='mergeCells']/*[local-name()='mergeCell']")) {
        $m = [regex]::Match($mc.GetAttribute('ref'), '^([A-Z]+)(\d+)')
        if ($m.Success -and (& $inArea (Get-ColumnIndex $m.Groups[1].Value) ([int]$m.Groups[2].Value))) { $lines.Add('M|' + $mc.GetAttribute('ref')) }
    }
    # 列の幅・行の高さ（範囲の中の列・行だけ）
    $c0 = 1; $c1 = 256; $r0 = 1; $r1 = 1048576
    if ($Areas.Count -gt 0) {
        $c0 = ($Areas | Measure-Object C0 -Minimum).Minimum; $c1 = ($Areas | Measure-Object C1 -Maximum).Maximum
        $r0 = ($Areas | Measure-Object R0 -Minimum).Minimum; $r1 = ($Areas | Measure-Object R1 -Maximum).Maximum
    }
    foreach ($col in $doc.SelectNodes("//*[local-name()='cols']/*[local-name()='col']")) {
        $lo = [math]::Max([int]$col.GetAttribute('min'), $c0); $hi = [math]::Min([int]$col.GetAttribute('max'), $c1)
        if ($lo -le $hi) { $lines.Add(('W|{0}|{1}|{2}|{3}' -f $lo, $hi, $col.GetAttribute('width'), $col.GetAttribute('hidden'))) }
    }
    foreach ($row in $doc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row'][@ht or @hidden]")) {
        $r = 0
        if ([int]::TryParse($row.GetAttribute('r'), [ref]$r) -and $r -ge $r0 -and $r -le $r1) { $lines.Add(('H|{0}|{1}|{2}' -f $r, $row.GetAttribute('ht'), $row.GetAttribute('hidden'))) }
    }
    # 図形・画像（範囲にかかるもの。位置は範囲の左上から）
    $items = Read-SheetDrawingItems $Package $SheetPart $Images
    if ($items.Count -gt 0) {
        $geometry = Get-SheetGeometry $doc $Mdw
        Set-DrawingBounds $items $geometry
        $ox = 0; $oy = 0; $ax1 = [double]::MaxValue; $ay1 = [double]::MaxValue
        if ($Areas.Count -gt 0) {
            $p0 = Get-SheetPoint $geometry $c0 $r0
            $p1 = Get-SheetPoint $geometry ($c1 + 1) ($r1 + 1)
            $ox = $p0[0]; $oy = $p0[1]; $ax1 = $p1[0]; $ay1 = $p1[1]
        }
        foreach ($it in $items) {
            $b = $it.Bounds
            if ($b.X -ge $ax1 -or $b.Y -ge $ay1 -or $b.X + $b.W -le $ox -or $b.Y + $b.H -le $oy) { continue }
            $lines.Add(('D|{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $it.Kind, [math]::Round($b.X - $ox), [math]::Round($b.Y - $oy), [math]::Round($b.W), [math]::Round($b.H), $it.Text, $it.Hash, $it.Look, $it.Hidden))
        }
    }
    return (Get-ShortHash ($lines -join "`n"))
}

function Read-Wireframes {
    #     元のワイヤーフレームのシート（表示されているワークシート）と、印刷範囲の中の見た目の要約
    param([string[]] $Paths)
    $sheets = New-Object System.Collections.Generic.List[object]
    foreach ($file in (Get-ExcelFilePaths $Paths '元のワイヤーフレーム')) {
        try {
            $pkg = Open-ZipPackage $file
            try {
                $parts = Get-WorkbookSheetParts $pkg
                $wb = Read-PartXml -Package $pkg -PartName $parts.WorkbookPart
                $rels = Get-PartRelationships -Package $pkg -PartName $parts.WorkbookPart
                $stylesPart = @($rels | Where-Object { $_.Kind -eq 'styles' } | ForEach-Object { $_.Target })[0]
                $styles = Read-StyleKeys $pkg $stylesPart
                $strings = Read-SharedStringList $pkg (@($rels | Where-Object { $_.Kind -eq 'sharedStrings' } | ForEach-Object { $_.Target })[0])
                $mdw = Get-MaxDigitWidth $pkg $stylesPart
                $images = @{}
                $areas = @{}
                foreach ($dn in $wb.SelectNodes("//*[local-name()='definedNames']/*[local-name()='definedName']")) {
                    if ($dn.GetAttribute('name') -ne '_xlnm.Print_Area' -or -not $dn.GetAttribute('localSheetId')) { continue }
                    $areas[[int]$dn.GetAttribute('localSheetId')] = Get-PrintAreas $dn.InnerText
                }
                $index = -1
                foreach ($s in $wb.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
                    $index++
                    $name = $s.GetAttribute('name')
                    $state = $s.GetAttribute('state')
                    if (($state -and $state -ne 'visible') -or -not $parts.Sheets.Contains($name)) { continue }
                    $part = $parts.Sheets[$name]
                    if ($part -notmatch 'worksheets/') { continue }
                    $area = $areas[$index]
                    if (-not $area) { $area = New-Object System.Collections.Generic.List[object] }
                    $sheets.Add([pscustomobject]@{
                            File = [System.IO.Path]::GetFileName($file); Path = $file; Sheet = $name; Key = (ConvertTo-ScreenKey $name); Cut = ($name.Length -ge 31)
                            HasArea = ($area.Count -gt 0); Area = (@($area | ForEach-Object { '{0}{1}:{2}{3}' -f (ConvertTo-ColumnName $_.C0), $_.R0, (ConvertTo-ColumnName $_.C1), $_.R1 }) -join ',')
                            Look = (Get-WireframeKey $pkg $part $area $styles $strings $mdw $images); Used = $false
                        })
                }
            }
            finally { $pkg.Zip.Dispose() }
        }
        catch { Write-Warning ("{0}: 読めませんでした: {1}" -f [System.IO.Path]::GetFileName($file), $_.Exception.Message) }
    }
    return , $sheets
}

# ---------------------------------------------------------------- 設計書のスクリーンショットを読む


function Read-DesignScreens {
    #     設計書の、スクリーンショット（画像）を貼ったシート（1 シートに 1 枚。いくつもあればいちばん大きいもの）
    #     Names: 画面名の候補（シート名・「画面名」の欄・版数を除いたファイル名。比べる形）
    param([string[]] $Paths)
    $searchArgs = @{ Pattern = '\S'; Path = $Paths; Target = 'Value' }
    if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
    if ($NoExcel) { $searchArgs['NoExcel'] = $true }
    $hits = @(& $searchScript @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
    $books = ConvertTo-DesignBooks $hits
    $screens = New-Object System.Collections.Generic.List[object]
    foreach ($file in (Get-ExcelFilePaths $Paths '設計書')) {
        $drawings = $null
        try { $drawings = Read-BookDrawings $file -Bounds }
        catch { Write-Warning ("{0}: 図形・画像を読めませんでした: {1}" -f [System.IO.Path]::GetFileName($file), $_.Exception.Message) }
        if (-not $drawings) { continue }
        $book = $null
        foreach ($k in $books.Keys) { if ([string]::Equals($k, $file, [System.StringComparison]::OrdinalIgnoreCase)) { $book = $books[$k]; break } }
        # 「画面名」の欄（シートに無ければ、ブックの中で 1 つだけの値）
        $bookNames = @()
        if ($book) { $bookNames = @($book.Sheets.Keys | ForEach-Object { Find-LabelValue $book.Sheets[$_] '画面名' } | Where-Object { $_ } | ForEach-Object { $_.Trim() } | Select-Object -Unique) }
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension((Get-VersionlessName ([System.IO.Path]::GetFileName($file))).Replace('{v}', ''))
        $index = -1
        foreach ($sheet in $drawings.Keys) {
            $index++
            $all = $drawings[$sheet]
            $hiddenGroups = @($all | Where-Object { $_.Hidden -and $_.Kind -eq 'grpSp' } | ForEach-Object { $_.Path + '/' })
            $shown = @($all | Where-Object { $it = $_; -not $it.Hidden -and $it.Bounds -and @($hiddenGroups | Where-Object { $it.Path.StartsWith($_) }).Count -eq 0 })
            # ロゴなどの小さい画像は、スクリーンショットとみなさない
            $pic = @($shown | Where-Object { $_.Kind -eq 'pic' -and $_.Bounds.W -ge 240 -and $_.Bounds.H -ge 160 } | Sort-Object { $_.Bounds.W * $_.Bounds.H } -Descending)[0]
            if (-not $pic) { continue }
            $p = $pic.Bounds
            # 画像の上に重ねた図形（丸数字・枠・吹き出しなど）
            $overlays = @($shown | Where-Object { $_.Kind -eq 'sp' -or $_.Kind -eq 'cxnSp' } | Where-Object { $cx = $_.Bounds.X + $_.Bounds.W / 2; $cy = $_.Bounds.Y + $_.Bounds.H / 2; $cx -ge $p.X -and $cx -le $p.X + $p.W -and $cy -ge $p.Y -and $cy -le $p.Y + $p.H }).Count
            $label = $null
            if ($book -and $book.Sheets.Contains($sheet)) { $label = Find-LabelValue $book.Sheets[$sheet] '画面名' }
            if (-not $label -and $bookNames.Count -eq 1) { $label = $bookNames[0] }
            if ($label) { $label = $label.Trim() }
            $names = @((ConvertTo-ScreenKey $sheet), (ConvertTo-ScreenKey $label)) | Where-Object { $_ } | Select-Object -Unique
            $shownName = $sheet
            if ($label) { $shownName = $label }
            $screens.Add([pscustomobject]@{
                    File = [System.IO.Path]::GetFileName($file); Path = $file; Sheet = $sheet; SheetIndex = $index; Name = $shownName
                    Names = @($names); SheetCut = ($sheet.Length -ge 31); FileKey = (ConvertTo-ScreenKey $baseName)
                    Picture = $pic; Overlays = $overlays
                })
        }
    }
    return , $screens
}

function Find-Wireframe {
    #     画面の元のシート（シート名が、設計書のシート名・「画面名」の欄と同じもの → ファイル名に含まれるもの）。
    #     シート名は 31 文字までなので、長い画面名は前の部分で比べる
    param($Shot, $Sheets)
    $found = @($Sheets | Where-Object {
            $w = $_
            ($Shot.Names -contains $w.Key) -or
            ($w.Cut -and @($Shot.Names | Where-Object { $_.StartsWith($w.Key) }).Count -gt 0) -or
            ($Shot.SheetCut -and $w.Key.StartsWith($Shot.Names[0]))
        })
    if ($found.Count -eq 0 -and $Shot.FileKey) {
        # ファイル名に画面名が含まれるもの（いちばん長い名前）
        $found = @($Sheets | Where-Object { $_.Key -and $Shot.FileKey.Contains($_.Key) } | Sort-Object { $_.Key.Length } -Descending | Select-Object -First 1)
    }
    return , $found
}

# ---------------------------------------------------------------- 記録


$recordColumns = @('画面', '元のファイル', '元のシート', '元の内容', '設計書のファイル', '設計書のシート', '画像', '記録日時')

if (-not $RecordFile) {
    $first = @(Get-PathItem $Wireframe[0])[0]
    if (-not $first) { throw "元のワイヤーフレームが見つかりません: $($Wireframe[0])" }
    $dir = $first.FullName
    if (-not $first.PSIsContainer) { $dir = Split-Path -Parent $first.FullName }
    $RecordFile = Join-Path $dir 'スクリーンショットの記録.csv'
}
$RecordFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RecordFile)

$records = New-Object System.Collections.Generic.List[object]
if (Test-Path -LiteralPath $RecordFile) {
    foreach ($r in (Read-CsvFile $RecordFile $recordColumns '記録のファイル')) { $records.Add($r) }
}

function Find-Record {
    #     画面の記録（版数を除いた設計書のファイル名とシート名 → 画面の名前）
    param($Shot)
    $name = Get-VersionlessName $Shot.File
    $found = @($records | Where-Object { $_.'設計書のシート' -eq $Shot.Sheet -and (Get-VersionlessName $_.'設計書のファイル') -eq $name })
    if ($found.Count -gt 0) { return $found[-1] }
    $found = @($records | Where-Object { $_.'画面' -eq $Shot.Name })
    if ($found.Count -eq 1) { return $found[0] }
    return $null
}

# ---------------------------------------------------------------- 確かめる


$wires = Read-Wireframes $Wireframe
$screens = Read-DesignScreens $Path

function Get-ScreenState {
    #     画面の状態と説明
    param($S)
    $w = $S.Wire
    if (-not $w) { return @('元のシートなし', '元のワイヤーフレームに、この画面のシートが見つかりません（元のシート名を画面名にしてください）') }
    if ($S.Wires.Count -gt 1) { return @('元のシートが複数', ('元のワイヤーフレームに、この画面のシートが {0} つあります（{1}）' -f $S.Wires.Count, (@($S.Wires | ForEach-Object { '{0} の {1}' -f $_.File, $_.Sheet }) -join '、'))) }
    $rec = $S.Record
    if (-not $rec) { return @('記録なし', 'まだ記録していません。いまのスクリーンショットが元と合っていれば、記録してください（-Record）') }
    $source = $rec.'元の内容' -ne $w.Look
    $picture = $rec.'画像' -ne $S.Picture.Hash
    if (-not $source -and -not $picture) { return @('最新', '') }
    if ($source -and -not $picture) { return @('元が更新された', ('元のワイヤーフレームが、記録（{0}）のあとで変わりました。スクリーンショットを撮り直して貼り替え、記録してください' -f $rec.'記録日時')) }
    $detail = 'スクリーンショットを、記録（{0}）のあとで貼り替えました。元と合っていれば、記録してください（-Record）' -f $rec.'記録日時'
    if ($source) { $detail = '元のワイヤーフレームが変わり、スクリーンショットも貼り替えました（記録 {0} のあと）。撮り直したものか確かめて、記録してください（-Record）' -f $rec.'記録日時' }
    return @('貼り替えた', $detail)
}

foreach ($s in $screens) {
    $found = Find-Wireframe $s $wires
    $s | Add-Member -NotePropertyName Wires -NotePropertyValue $found
    $w = $null
    if ($found.Count -gt 0) { $w = $found[0]; foreach ($f in $found) { $f.Used = $true } }
    $s | Add-Member -NotePropertyName Wire -NotePropertyValue $w
    $s | Add-Member -NotePropertyName Record -NotePropertyValue (Find-Record $s)
}

# 記録する（-Record）: 記録の無い画面と貼り替えた画面。元だけが変わった画面は、貼り替えるまで記録しない（-Screen で選べば記録する）
$targets = New-Object System.Collections.Generic.List[object]
$recorded = New-Object System.Collections.Generic.List[string]
$skipped = New-Object System.Collections.Generic.List[string]
if ($Record) {
    foreach ($s in $screens) {
        if (-not $s.Wire -or $s.Wires.Count -gt 1) { continue }
        $state = (Get-ScreenState $s)[0]
        if ($Screen) {
            $hit = @($Screen | Where-Object { $s.Name -like $_ -or $s.Sheet -like $_ }).Count -gt 0
            if (-not $hit) { continue }
        }
        elseif ($state -eq '最新') { continue }
        elseif ($state -eq '元が更新された') { $skipped.Add($s.Name); continue }
        $targets.Add($s)
    }
    # -WhatIf のときは、記録を書かず、状態も記録の前のまま出す
    if ($targets.Count -gt 0 -and $PSCmdlet.ShouldProcess($RecordFile, ('スクリーンショットの記録を書く（{0}）' -f (@($targets | ForEach-Object { $_.Name }) -join '・')))) {
        $now = (Get-Date).ToString('yyyy/MM/dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        foreach ($s in $targets) {
            $row = [pscustomobject][ordered]@{
                '画面' = $s.Name; '元のファイル' = $s.Wire.File; '元のシート' = $s.Wire.Sheet; '元の内容' = $s.Wire.Look
                '設計書のファイル' = $s.File; '設計書のシート' = $s.Sheet; '画像' = $s.Picture.Hash; '記録日時' = $now
            }
            if ($s.Record) { [void]$records.Remove($s.Record) }
            $records.Add($row)
            $s.Record = $row
            $recorded.Add($s.Name)
        }
        $csv = @($records | Select-Object $recordColumns | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
        [System.IO.File]::WriteAllText($RecordFile, $csv + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
    }
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($s in $screens) {
    $state = Get-ScreenState $s
    $detail = $state[1]
    if ($state[0] -eq '貼り替えた' -and $s.Overlays -gt 0) { $detail += '。画像の上に重ねた図形が {0} 個あります。位置がずれていないか確かめてください' -f $s.Overlays }
    if ($s.Wire -and -not $s.Wire.HasArea) { $detail = ($detail + '（元のシートに印刷範囲が無いため、シート全体で比べています）').TrimStart() }
    $r = [pscustomobject]@{
        Screen = $s.Name; Status = $state[0]; Location = ('{0} {1}!{2}' -f $s.File, $s.Sheet, $s.Picture.Cell); Detail = $detail
        File = $s.File; Path = $s.Path; Sheet = $s.Sheet; Cell = $s.Picture.Cell; SheetIndex = $s.SheetIndex
        Wireframe = $(if ($s.Wire) { '{0} {1}' -f $s.Wire.File, $s.Wire.Sheet } else { '' }); WireframePath = $(if ($s.Wire) { $s.Wire.Path } else { '' }); WireframeSheet = $(if ($s.Wire) { $s.Wire.Sheet } else { '' })
        Recorded = $(if ($s.Record) { $s.Record.'記録日時' } else { '' })
    }
    $r.PSObject.TypeNames.Insert(0, 'Screenshots.Screen')
    $results.Add($r)
}
$results

$order = @('元が更新された', '貼り替えた', '記録なし', '元のシートなし', '元のシートが複数', '最新')
if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($results | Sort-Object @{ e = { $order.IndexOf($_.Status) } }, File, SheetIndex)) {
        $row = @{ Values = @($r.Screen, $r.Status, ('{0} {1}' -f $r.File, $r.Sheet), $r.Wireframe, $r.Recorded, $r.Detail); Link = @{ Column = 2; Path = $r.Path; Location = ("'{0}'!{1}" -f $r.Sheet.Replace("'", "''"), $r.Cell) } }
        $rows.Add($row)
    }
    $wireRows = New-Object System.Collections.Generic.List[object]
    foreach ($w in $wires) {
        $use = '使っている'
        if (-not $w.Used) { $use = '設計書に無い' }
        $area = $w.Area
        if (-not $w.HasArea) { $area = '（印刷範囲なし）' }
        $wireRows.Add(@{ Values = @($w.File, $w.Sheet, $area, $use); Link = @{ Column = 1; Path = $w.Path; Location = ("'{0}'!A1" -f $w.Sheet.Replace("'", "''")) } })
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '画面ごとの状態'; Columns = @(@{ Header = '画面'; Width = 24 }, @{ Header = '状態'; Width = 14 }, @{ Header = '設計書'; Width = 40 }, @{ Header = '元のワイヤーフレーム'; Width = 36 }, @{ Header = '記録日時'; Width = 16 }, @{ Header = '内容'; Width = 80 }); Rows = $rows },
            @{ Name = '元のシート'; Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = 'シート'; Width = 24 }, @{ Header = '印刷範囲'; Width = 20 }, @{ Header = '設計書'; Width = 14 }); Rows = $wireRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$counts = @($order | ForEach-Object { $st = $_; $n = @($results | Where-Object Status -eq $st).Count; if ($n -gt 0) { '{0} {1}' -f $st, $n } }) -join '・'
if (-not $counts) { $counts = 'スクリーンショットを貼ったシートが見つかりません' }
Write-Host ('スクリーンショットのある画面 {0}: {1}' -f $results.Count, $counts)
$noArea = @($wires | Where-Object { $_.Used -and -not $_.HasArea })
if ($noArea.Count -gt 0) { Write-Host ('印刷範囲の無い元のシート {0}: {1}（撮る範囲を印刷範囲にすると、範囲の外のメモを直しても古いと出なくなります）' -f $noArea.Count, (@($noArea | Select-Object -First 5 | ForEach-Object { $_.Sheet }) -join '、')) }
$unused = @($wires | Where-Object { -not $_.Used })
if ($unused.Count -gt 0) { Write-Host ('設計書に無い元のシート {0}: {1}' -f $unused.Count, (@($unused | Select-Object -First 5 | ForEach-Object { $_.Sheet }) -join '、')) }
if ($Record) {
    if ($recorded.Count -gt 0) { Write-Host ('{0} 画面を記録しました: {1}' -f $recorded.Count, $RecordFile) }
    elseif ($targets.Count -eq 0) { Write-Host '記録する画面はありませんでした' }
    if ($skipped.Count -gt 0) { Write-Host ('元が更新されたままの画面は記録していません（貼り替えてから記録してください。貼り替えの要らない変更なら -Screen で画面を選んで記録できます）: {0}' -f ($skipped -join '、')) }
}
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書に貼ったスクリーンショットが、元のワイヤーフレーム（Excel）より古くなっていないかを確かめます（Excel 不要。設計書は変更しません）。
#
# .DESCRIPTION
#     設計書の画面レイアウトに貼ったスクリーンショットと、元のワイヤーフレームのシート（シート名が画面名）を対応付け、
#     記録（-Record で作る CSV）と比べて、画面ごとに次のどれかを出します。
#
#       最新               記録のあと、元もスクリーンショットも変わっていない
#       元が更新された     元のワイヤーフレームが記録のあとで変わった（撮り直して貼り替える）
#       貼り替えた         スクリーンショットを記録のあとで貼り替えた（元と合っていれば記録し直す。重ねた図形があれば、位置を確かめるよう知らせる）
#       記録なし           まだ記録していない
#       元のシートなし     元のワイヤーフレームに、この画面のシートが無い
#       元のシートが複数   元のワイヤーフレームに、この画面のシートが 2 つ以上ある
#
#     元のシートは、印刷範囲の中だけを比べます（セルの値と書式・結合・列の幅・行の高さ・図形と画像）。範囲の外のメモを直したり、
#     開いて保存し直したりしただけでは「元が更新された」になりません。印刷範囲の無いシートは、シート全体で比べます。
#
#     画面の対応付け: 設計書のスクリーンショットのあるシートの、シート名 → 「画面名」の欄（シートに無ければ、ブックの中で 1 つだけの値）
#     → ファイル名 の順に、元のシート名と比べます。シート名は 31 文字までなので、長い画面名は前の部分で比べます。
#     スクリーンショットは、1 シートに 1 枚（いくつもあれば、いちばん大きい画像）とみなします。
#
#     記録は、元のワイヤーフレームのフォルダの「スクリーンショットの記録.csv」に書きます（設計書には何も書きません）。
#
# .PARAMETER Path
#     スクリーンショットを貼った設計書のフォルダまたはファイル（複数可・ワイルドカード可）。
#
# .PARAMETER Wireframe
#     元のワイヤーフレーム（Excel）のフォルダまたはファイル（複数可）。シート名を画面名にしてください。
#
# .PARAMETER Record
#     スクリーンショットを貼った画面を記録します。記録するのは、記録の無い画面と、記録のあとで貼り替えた画面です。
#     元だけが変わった画面は、貼り替えるまで記録しません（-Screen で選んだ画面は、状態によらず記録します）。
#
# .PARAMETER Screen
#     -Record で記録する画面の名前（複数可・ワイルドカード可）。元を直したが、スクリーンショットの貼り替えが要らないときに使います。
#
# .PARAMETER RecordFile
#     記録のファイル（CSV）。既定は、-Wireframe の最初のフォルダの「スクリーンショットの記録.csv」。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の設計書の「画面名」の欄は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（画面ごとの状態と、元のシートの一覧（印刷範囲・設計書で使っているか）。場所をクリックすると該当のシートが開きます）。
#
# .EXAMPLE
#     # 貼り替えたあとに記録する（初めては、いまの画面をすべて記録する）
#     .\Test-Screenshots.ps1 D:\設計書\画面設計書 -Wireframe D:\ワイヤーフレーム -Record
#
# .EXAMPLE
#     # 古くなったスクリーンショットを確かめる
#     .\Test-Screenshots.ps1 D:\設計書\画面設計書 -Wireframe D:\ワイヤーフレーム -OutXlsx D:\スクリーンショットの状態.xlsx
#
# .EXAMPLE
#     # 元の注記だけを直した（貼り替えの要らない）画面を、記録し直す
#     .\Test-Screenshots.ps1 D:\設計書\画面設計書 -Wireframe D:\ワイヤーフレーム -Record -Screen '受注登録'
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     スクリーンショットの中身（絵）は比べません。記録のときの画像と同じかどうかだけを見ます。
