[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateSet('View', 'Hidden', 'Comment', 'Revision', 'Word', 'Property', 'Link', 'Error', 'Macro', 'Print', 'Version')]
    [string[]] $Check = @('View', 'Hidden', 'Comment', 'Revision', 'Word', 'Property', 'Link', 'Error', 'Macro', 'Print', 'Version'),

    [string] $Words = 'TODO|TBD|FIXME|要確認|未定|検討中|仮置き|（仮）|\(仮\)|XXX|○○|××',

    [string] $AllowedNames,

    [ValidateRange(10, 400)]
    [int] $Zoom = 100,

    [string] $FileNamePattern,

    [switch] $RequirePageNumber,

    [string] $HeaderFooterText,

    [string] $HistorySheet = '改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録',

    [switch] $NoRecurse,

    [switch] $Fix,

    [string] $BackupDir,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$resetScript = Join-Path $PSScriptRoot 'Reset-ExcelView.ps1'
if ($BackupDir) { $BackupDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupDir) }

foreach ($p in @($Words, $AllowedNames, $FileNamePattern, $HeaderFooterText, $HistorySheet)) {
    if ($p) {
        try { [void][regex]::new($p) }
        catch { throw "正規表現が不正です: $p 詳細: $($_.Exception.InnerException.Message)" }
    }
}
if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'ExcelDeliverable.Issue' -DefaultDisplayPropertySet File, Location, Level, Detail -Force -WhatIf:$false

$checkSet = @{}
foreach ($c in $Check) { $checkSet[$c] = $true }

$errorValues = @('#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!')
$maxPerSheet = 20   # エラー値などを 1 シートで個別に挙げる上限（超えた分はまとめて 1 件にする）

# 用紙の幅と高さ（インチ）。pageSetup の paperSize の番号ごと。無い場合は Letter（Excel の既定）
$paperSizes = @{
    1 = @('レター', 8.5, 11); 5 = @('リーガル', 8.5, 14); 8 = @('A3', 11.69, 16.54); 9 = @('A4', 8.27, 11.69); 11 = @('A5', 5.83, 8.27)
    12 = @('B4', 10.12, 14.33); 13 = @('B5', 7.17, 10.12)
}

# ---------------------------------------------------------------- 指摘の記録


$script:Issues = New-Object System.Collections.Generic.List[object]

function Add-Issue {
    param(
        [ValidateSet('エラー', '警告', '情報')] [string] $Level,
        [string] $Category,
        [string] $Item,
        [string] $Detail,
        [string] $Sheet = '',
        [string] $Cell = ''
    )
    $location = $Sheet
    if ($Sheet -and $Cell) { $location = '{0}!{1}' -f $Sheet, $Cell }
    $issue = [pscustomobject]@{
        File     = $script:CurrentFile
        Location = $location
        Level    = $Level
        Category = $Category
        Item     = $Item
        Detail   = $Detail
        Sheet    = $Sheet
        Cell     = $Cell
        Path     = $script:CurrentPath
    }
    $issue.PSObject.TypeNames.Insert(0, 'ExcelDeliverable.Issue')
    $script:Issues.Add($issue)
}

# ファイルのプロパティのうち、個人名・会社名が残りやすいもの（要素名 → 表示の名前）
$propertyLabels = [ordered]@{ 'creator' = '作成者'; 'lastModifiedBy' = '最終更新者'; 'Company' = '会社名'; 'Manager' = '管理者' }

function Test-AllowedName {
    param([string] $Value)
    return ($AllowedNames -and $Value -match $AllowedNames)
}

# ---------------------------------------------------------------- シートの読み取り


function Read-SheetFacts {
    #     シートの XML を 1 回だけ先頭から読み、チェックに必要な情報を集める。
    #     表示（sheetView / pane / selection）、非表示の行・列、エラー値のセル、ハイパーリンク
    #     印刷の設定（pageSetup / pageMargins / headerFooter / 列の幅）と、値のあるセルの範囲・印刷範囲の外のセル
    param($Package, [string] $PartName, $PrintAreas)

    $facts = [pscustomobject]@{
        HasView = $false; TabSelected = $false; Zoom = 100; TopLeft = 'A1'; ViewMode = 'normal'
        Pane = $null; Selections = New-Object System.Collections.Generic.List[object]
        HiddenRows = 0; HiddenCols = 0
        Errors = New-Object System.Collections.Generic.List[object]
        Links = New-Object System.Collections.Generic.List[object]
        PageSetup = @{}; FitToPage = $false; MarginLeft = 0.7; MarginRight = 0.7
        HeaderFooter = @{}; HasColBreaks = $false
        Cols = New-Object System.Collections.Generic.List[object]; DefaultColWidth = 0.0; BaseColWidth = 8
        ValueCells = 0; MaxValueCol = 0; Outside = 0; OutsideRefs = New-Object System.Collections.Generic.List[string]
    }

    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return $facts }
    try {
        $inFirstView = $false
        $viewDone = $false
        $errorCell = $null
        $cellRef = $null
        $counted = $false
        $advance = $true
        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
                if ($reader.LocalName -eq 'sheetView' -and $inFirstView) { $inFirstView = $false; $viewDone = $true }
                elseif ($reader.LocalName -eq 'c') { $errorCell = $null; $cellRef = $null }
                continue
            }
            if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }

            switch ($reader.LocalName) {
                'sheetView' {
                    # ブックの表示（workbookViewId）が複数あっても、最初のものだけを見る
                    if (-not $viewDone -and -not $facts.HasView) {
                        $facts.HasView = $true
                        $facts.TabSelected = @('1', 'true') -contains $reader.GetAttribute('tabSelected')
                        $z = $reader.GetAttribute('zoomScale')
                        if ($z) { $facts.Zoom = [int]$z }
                        $t = $reader.GetAttribute('topLeftCell')
                        if ($t) { $facts.TopLeft = $t.ToUpperInvariant() }
                        $v = $reader.GetAttribute('view')
                        if ($v) { $facts.ViewMode = $v }
                        $inFirstView = -not $reader.IsEmptyElement
                        if (-not $inFirstView) { $viewDone = $true }
                    }
                }
                'pane' {
                    if ($inFirstView) {
                        $facts.Pane = [pscustomobject]@{
                            XSplit = [double]('0' + $reader.GetAttribute('xSplit'))
                            YSplit = [double]('0' + $reader.GetAttribute('ySplit'))
                            TopLeft = ([string]$reader.GetAttribute('topLeftCell')).ToUpperInvariant()
                            ActivePane = [string]$reader.GetAttribute('activePane')
                            State = [string]$reader.GetAttribute('state')
                        }
                    }
                }
                'selection' {
                    if ($inFirstView) {
                        $ac = $reader.GetAttribute('activeCell')
                        if (-not $ac) { $ac = 'A1' }
                        $facts.Selections.Add([pscustomobject]@{ Pane = [string]$reader.GetAttribute('pane'); ActiveCell = $ac.ToUpperInvariant() })
                    }
                }
                'col' {
                    if (@('1', 'true') -contains $reader.GetAttribute('hidden')) {
                        $facts.HiddenCols += [int]$reader.GetAttribute('max') - [int]$reader.GetAttribute('min') + 1
                    }
                    $w = 0.0
                    [void][double]::TryParse([string]$reader.GetAttribute('width'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$w)
                    $facts.Cols.Add([pscustomobject]@{ Min = [int]$reader.GetAttribute('min'); Max = [int]$reader.GetAttribute('max'); Width = $w; Hidden = (@('1', 'true') -contains $reader.GetAttribute('hidden')) })
                }
                'sheetFormatPr' {
                    $w = 0.0
                    if ([double]::TryParse([string]$reader.GetAttribute('defaultColWidth'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$w)) { $facts.DefaultColWidth = $w }
                    $b = $reader.GetAttribute('baseColWidth')
                    if ($b) { $facts.BaseColWidth = [int]$b }
                }
                'pageSetUpPr' { $facts.FitToPage = @('1', 'true') -contains $reader.GetAttribute('fitToPage') }
                'pageMargins' {
                    $l = 0.0; $r = 0.0
                    if ([double]::TryParse([string]$reader.GetAttribute('left'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$l)) { $facts.MarginLeft = $l }
                    if ([double]::TryParse([string]$reader.GetAttribute('right'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$r)) { $facts.MarginRight = $r }
                }
                'pageSetup' {
                    foreach ($a in @('orientation', 'paperSize', 'scale', 'fitToWidth', 'fitToHeight')) { $facts.PageSetup[$a] = $reader.GetAttribute($a) }
                }
                { @('oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter') -contains $_ } {
                    $facts.HeaderFooter[$reader.LocalName] = $reader.ReadElementContentAsString()
                    $advance = $false
                }
                'colBreaks' { $facts.HasColBreaks = $true }
                'is' { if ($cellRef -and -not $counted) { Add-ValueCell $facts $cellRef $PrintAreas; $counted = $true } }
                'row' {
                    if (@('1', 'true') -contains $reader.GetAttribute('hidden')) { $facts.HiddenRows++ }
                }
                'c' {
                    if ($reader.GetAttribute('t') -eq 'e' -and -not $reader.IsEmptyElement) { $errorCell = $reader.GetAttribute('r') }
                    $cellRef = $null
                    if (-not $reader.IsEmptyElement) { $cellRef = $reader.GetAttribute('r') }
                    $counted = $false
                }
                'v' {
                    if ($cellRef -and -not $counted) { Add-ValueCell $facts $cellRef $PrintAreas; $counted = $true }
                    if ($errorCell) {
                        $facts.Errors.Add([pscustomobject]@{ Cell = $errorCell; Value = $reader.ReadElementContentAsString() })
                        $errorCell = $null
                        $advance = $false
                    }
                }
                'hyperlink' {
                    $facts.Links.Add([pscustomobject]@{
                        Ref = $reader.GetAttribute('ref')
                        RelId = $reader.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                        Location = $reader.GetAttribute('location')
                    })
                }
            }
        }
    }
    finally { $reader.Dispose() }
    return $facts
}

function Add-ValueCell {
    #     値のあるセルを数え、いちばん右の列と、印刷範囲の外にあるセルを覚える
    param($Facts, [string] $Ref, $PrintAreas)
    $m = [regex]::Match($Ref, '^([A-Z]+)(\d+)$')
    if (-not $m.Success) { return }
    $col = Get-ColumnIndex $m.Groups[1].Value
    $row = [int]$m.Groups[2].Value
    $Facts.ValueCells++
    if ($col -gt $Facts.MaxValueCol) { $Facts.MaxValueCol = $col }
    if ($PrintAreas -and $PrintAreas.Count -gt 0) {
        foreach ($a in $PrintAreas) { if ($row -ge $a.R0 -and $row -le $a.R1 -and $col -ge $a.C0 -and $col -le $a.C1) { return } }
        $Facts.Outside++
        if ($Facts.OutsideRefs.Count -lt 5) { $Facts.OutsideRefs.Add($Ref) }
    }
}

function Get-PrintAreas {
    #     名前の定義 _xlnm.Print_Area（'シート'!$A$1:$H$40,'シート'!$J$1:$K$5）を、行・列の範囲の一覧にする
    param([string] $Text)
    $areas = New-Object System.Collections.Generic.List[object]
    foreach ($m in [regex]::Matches($Text, '\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?(?=,|$)')) {
        $c0 = Get-ColumnIndex $m.Groups[1].Value; $r0 = [int]$m.Groups[2].Value
        $c1 = $c0; $r1 = $r0
        if ($m.Groups[3].Success) { $c1 = Get-ColumnIndex $m.Groups[3].Value; $r1 = [int]$m.Groups[4].Value }
        $areas.Add([pscustomobject]@{ R0 = $r0; R1 = $r1; C0 = $c0; C1 = $c1 })
    }
    return , $areas
}

function Get-MaxDigitWidth {
    #     列の幅（文字数）をピクセルにするための、標準フォントの数字 1 文字の幅（ピクセル）の目安。
    #     日本語のフォント（游ゴシック・ＭＳ Ｐゴシックなど）は 11 ポイントで 8、欧文のフォント（Calibri など）は 7
    param($Package, [string] $StylesPart)
    $doc = $null
    if ($StylesPart) { $doc = Read-PartXml -Package $Package -PartName $StylesPart }
    if ($null -eq $doc) { return 7 }
    $font = $doc.SelectSingleNode("//*[local-name()='fonts']/*[local-name()='font']")
    if ($null -eq $font) { return 7 }
    $name = ''; $size = 11.0
    $n = $font.SelectSingleNode("*[local-name()='name']")
    if ($n) { $name = $n.GetAttribute('val') }
    $z = $font.SelectSingleNode("*[local-name()='sz']")
    if ($z) { [void][double]::TryParse($z.GetAttribute('val'), [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$size) }
    $base = 7
    if ($name -match '[^\x00-\x7F]|Yu Gothic|Meiryo|MS P?Gothic|MS P?Mincho') { $base = 8 }
    return [math]::Max(4, [math]::Round($base * $size / 11))
}

function Test-PrintSetup {
    #     印刷の設定をチェックする（表示されているワークシートで、値のあるセルがあるものだけ）
    param($Sheets, [int] $Mdw)

    $printable = @($Sheets | Where-Object { $_.Facts -and $_.State -eq 'visible' -and $_.IsWorksheet -and $_.Facts.ValueCells -gt 0 })
    foreach ($s in $printable) {
        $f = $s.Facts
        # 印刷範囲の外に残ったデータ
        if ($s.PrintAreas.Count -gt 0 -and $f.Outside -gt 0) {
            $more = ''
            if ($f.Outside -gt $f.OutsideRefs.Count) { $more = ' など' }
            Add-Issue '警告' '印刷' '印刷範囲の外' ("印刷範囲の外に値のあるセルが {0} 個あります（{1}{2}）。印刷されません" -f $f.Outside, ($f.OutsideRefs -join '、'), $more) -Sheet $s.Name -Cell $f.OutsideRefs[0]
        }

        # 横が 1 ページに収まるか（「次のページ数に合わせて印刷」や改ページを設定していれば意図どおりとみなす）
        if (-not $f.FitToPage -and -not $f.HasColBreaks) {
            $c0 = 1; $c1 = $f.MaxValueCol
            if ($s.PrintAreas.Count -gt 0) { $c0 = ($s.PrintAreas | Measure-Object C0 -Minimum).Minimum; $c1 = ($s.PrintAreas | Measure-Object C1 -Maximum).Maximum }
            $default = $f.DefaultColWidth
            if ($default -le 0) { $default = $f.BaseColWidth + 0.71 }
            $px = 0.0
            for ($c = $c0; $c -le $c1; $c++) {
                $w = $default
                foreach ($col in $f.Cols) { if ($c -ge $col.Min -and $c -le $col.Max) { $w = $col.Width; if ($col.Hidden) { $w = 0 }; break } }
                $px += $w * $Mdw
            }
            $scale = 100
            if ($f.PageSetup['scale']) { $scale = [int]$f.PageSetup['scale'] }
            $paperNo = 1
            if ($f.PageSetup['paperSize']) { $paperNo = [int]$f.PageSetup['paperSize'] }
            $paper = $paperSizes[$paperNo]
            if ($paper) {
                $paperWidth = $paper[1]
                if ($f.PageSetup['orientation'] -eq 'landscape') { $paperWidth = $paper[2] }
                $printableInch = $paperWidth - $f.MarginLeft - $f.MarginRight
                $contentInch = $px * 0.75 / 72 * $scale / 100
                # 列の幅の見積もりには誤差があるので、1 割はみ出すまでは 1 ページとみなす
                if ($printableInch -gt 0 -and $contentInch -gt $printableInch * 1.1) {
                    $pages = [math]::Ceiling($contentInch / $printableInch)
                    $orient = '縦'
                    if ($f.PageSetup['orientation'] -eq 'landscape') { $orient = '横' }
                    Add-Issue '警告' '印刷' '横のはみ出し' ("印刷すると横が約 {0} ページに分かれます（{1}・{2}向き・{3}%）。「ページ設定」の「次のページ数に合わせて印刷」で横 1 ページにしてください" -f $pages, $paper[0], $orient, $scale) -Sheet $s.Name
                }
            }
        }

        # ヘッダー／フッター
        $hf = (@($f.HeaderFooter.Values) -join "`n")
        if ($hf -match '&[Zz]') { Add-Issue '警告' '情報漏れ' 'ヘッダー／フッター' 'ヘッダー／フッターにファイルの場所（&Z）が印刷されます' -Sheet $s.Name }
        if ($checkSet['Word'] -and $Words) {
            $plain = [regex]::Replace($hf, '&(?:"[^"]*"|\d+|[A-Za-z&])', ' ')
            $m = [regex]::Match($plain, $Words, 'IgnoreCase')
            if ($m.Success) { Add-Issue '警告' '消し忘れ' '残っている文言' ("ヘッダー／フッターに「{0}」が残っています" -f $m.Value) -Sheet $s.Name }
        }
        if ($RequirePageNumber -and $hf -notmatch '&[Pp]') { Add-Issue '警告' '印刷' 'ページ番号' 'ヘッダー／フッターにページ番号がありません' -Sheet $s.Name }
        if ($HeaderFooterText -and $hf -notmatch $HeaderFooterText) { Add-Issue '警告' '印刷' 'ヘッダー／フッター' ("ヘッダー／フッターに「{0}」がありません" -f $HeaderFooterText) -Sheet $s.Name }
    }

    # 用紙のサイズ・向きがシートごとにばらばら
    if ($printable.Count -gt 1) {
        foreach ($kind in @(@('paperSize', '用紙サイズ'), @('orientation', '用紙の向き'))) {
            $groups = [ordered]@{}
            foreach ($s in $printable) {
                $v = $s.Facts.PageSetup[$kind[0]]
                if ($kind[0] -eq 'paperSize') {
                    if (-not $v) { $v = '1' }
                    $name = $paperSizes[[int]$v]
                    if ($name) { $v = $name[0] } else { $v = "番号 $v" }
                }
                elseif ($v -eq 'landscape') { $v = '横' } else { $v = '縦' }
                if (-not $groups.Contains($v)) { $groups[$v] = New-Object System.Collections.Generic.List[string] }
                $groups[$v].Add($s.Name)
            }
            if ($groups.Count -gt 1) {
                $parts = @($groups.Keys | ForEach-Object { '{0}: {1}' -f $_, ($groups[$_] -join '、') })
                Add-Issue '情報' '印刷' $kind[1] ("{0}がシートによって違います（{1}）" -f $kind[1], ($parts -join ' / '))
            }
        }
    }
}

function ConvertTo-VersionText {
    #     版数の文字（「第1.1版」「Ver.1.10」「v1.1」など）を比較用の「1.1」の形にする（無ければ $null）
    param([string] $Text)
    if (-not $Text) { return $null }
    $t = $Text.Normalize([System.Text.NormalizationForm]::FormKC)
    $m = [regex]::Match($t, '(\d+(?:\.\d+)*)')
    if (-not $m.Success) { return $null }
    return (@($m.Groups[1].Value.Split('.') | ForEach-Object { [string][int64]$_ }) -join '.')
}

function ConvertTo-DateText {
    #     日付の文字（2026/10/1・2026-10-01・2026年10月1日）を「2026/10/01」の形にする（無ければ $null）
    param([string] $Text)
    if (-not $Text) { return $null }
    $t = $Text.Normalize([System.Text.NormalizationForm]::FormKC)
    $m = [regex]::Match($t, '(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})')
    if (-not $m.Success) { return $null }
    return '{0}/{1:D2}/{2:D2}' -f [int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value
}

function Compare-VersionText {
    #     版数どうしを数として比べる（-1 / 0 / 1）
    param([string] $A, [string] $B)
    $x = @($A.Split('.') | ForEach-Object { [int64]$_ }); $y = @($B.Split('.') | ForEach-Object { [int64]$_ })
    for ($i = 0; $i -lt [math]::Max($x.Count, $y.Count); $i++) {
        $p = 0; $q = 0
        if ($i -lt $x.Count) { $p = $x[$i] }
        if ($i -lt $y.Count) { $q = $y[$i] }
        if ($p -lt $q) { return -1 }
        if ($p -gt $q) { return 1 }
    }
    return 0
}

function Test-VersionConsistency {
    #     表紙の版数・更新日と、改訂履歴の最後の行と、ファイル名の版数が合っているかをチェックする
    param([string] $FilePath, $Sheets, [string] $FileName)

    $visible = @($Sheets | Where-Object { $_.State -eq 'visible' -and $_.IsWorksheet })
    $cover = @($visible | Where-Object { $_.Name -match '^\s*表紙\s*$' })[0]
    if (-not $cover) { $cover = $visible[0] }
    $hist = @($Sheets | Where-Object { $_.IsWorksheet -and $_.Name -match $HistorySheet })[0]
    $names = @(@($cover, $hist) | Where-Object { $_ } | ForEach-Object { [regex]::Escape($_.Name) } | Select-Object -Unique)
    if ($names.Count -eq 0) { return }

    # 表紙と改訂履歴のセルの値（日付は yyyy/m/d の文字）を Search-Excel で読む
    $cells = @{}
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($FilePath)) -Target Value -NoExcel `
        -SheetNamePattern ('^(?:{0})$' -f ($names -join '|')) -WarningAction SilentlyContinue 6>$null)
    foreach ($h in $hits) {
        $m = [regex]::Match($h.Cell, '^([A-Z]+)(\d+)$')
        if (-not $cells.ContainsKey($h.Sheet)) { $cells[$h.Sheet] = @{} }
        $row = [int]$m.Groups[2].Value
        if (-not $cells[$h.Sheet].ContainsKey($row)) { $cells[$h.Sheet][$row] = New-Object System.Collections.Generic.List[object] }
        $cells[$h.Sheet][$row].Add([pscustomobject]@{ Col = (Get-ColumnIndex $m.Groups[1].Value); Cell = $h.Cell; Text = $h.Text.Trim() })
    }
    $norm = { param($t) $t.Normalize([System.Text.NormalizationForm]::FormKC) -replace '\s', '' }
    $versionLabel = '^(版数|版|改訂版数|バージョン|Version|Ver\.?|Rev\.?|改版番号)$'
    $dateLabel = '^(最終更新日|更新日|改訂日|改版日|最終改訂日)$'

    # ---- 表紙: 見出し（「版数」など）の右の値、または「第1.1版」「Ver.1.1」の形の文字
    $coverVersion = $null; $coverVersionCell = $null; $coverDate = $null; $coverDateCell = $null
    if ($cover -and $cells.ContainsKey($cover.Name)) {
        foreach ($row in ($cells[$cover.Name].Keys | Sort-Object)) {
            $list = @($cells[$cover.Name][$row] | Sort-Object Col)
            for ($i = 0; $i -lt $list.Count; $i++) {
                $label = & $norm $list[$i].Text
                if (-not $coverVersion -and $label -match $versionLabel -and $i + 1 -lt $list.Count) {
                    $coverVersion = ConvertTo-VersionText $list[$i + 1].Text; $coverVersionCell = $list[$i + 1].Cell
                }
                elseif (-not $coverDate -and $label -match $dateLabel -and $i + 1 -lt $list.Count) {
                    $coverDate = ConvertTo-DateText $list[$i + 1].Text; $coverDateCell = $list[$i + 1].Cell
                }
                elseif (-not $coverVersion -and $label -match '^(第\d+(\.\d+)*版|(Ver|Version|Rev)\.?\d+(\.\d+)*)$') {
                    $coverVersion = ConvertTo-VersionText $label; $coverVersionCell = $list[$i].Cell
                }
            }
        }
    }

    # ---- 改訂履歴: 「版数」と「改訂日」などの見出しがある行の下で、版数のある最後の行。
    #      シートの上部（帳票のヘッダ）にも「版数」「更新日」が並ぶことがあるので、下に版数の行がいちばん多い見出しの行を使う
    $histVersion = $null; $histDate = $null
    if ($hist -and $cells.ContainsKey($hist.Name)) {
        $rows = @($cells[$hist.Name].Keys | Sort-Object)
        $best = $null
        foreach ($row in $rows) {
            $vCol = 0; $dCol = 0
            foreach ($c in $cells[$hist.Name][$row]) {
                $label = & $norm $c.Text
                if ($label -match $versionLabel) { $vCol = $c.Col }
                elseif ($label -match '^(日付|改訂日|改版日|更新日|年月日|変更日)$') { $dCol = $c.Col }
            }
            if (-not $vCol -or -not $dCol) { continue }
            $entries = New-Object System.Collections.Generic.List[object]
            foreach ($r in @($rows | Where-Object { $_ -gt $row })) {
                $vc = @($cells[$hist.Name][$r] | Where-Object { $_.Col -eq $vCol })[0]
                if (-not $vc) { continue }
                $v = ConvertTo-VersionText $vc.Text
                if (-not $v) { continue }
                $dc = @($cells[$hist.Name][$r] | Where-Object { $_.Col -eq $dCol })[0]
                $d = $null
                if ($dc) { $d = ConvertTo-DateText $dc.Text }
                $entries.Add([pscustomobject]@{ Version = $v; Text = $vc.Text; Cell = $vc.Cell; Date = $d })
            }
            if (-not $best -or $entries.Count -gt $best.Count) { $best = $entries }
        }
        if ($best -and $best.Count -gt 0) {
            for ($i = 1; $i -lt $best.Count; $i++) {
                if ((Compare-VersionText $best[$i].Version $best[$i - 1].Version) -le 0) {
                    Add-Issue '警告' '版数' '改訂履歴の版数' ("改訂履歴の版数（{0}）が前の行（{1}）より大きくありません" -f $best[$i].Text, $best[$i - 1].Text) -Sheet $hist.Name -Cell $best[$i].Cell
                }
            }
            $histVersion = $best[$best.Count - 1].Version
            $histDate = $best[$best.Count - 1].Date
        }
    }

    # ---- 突き合わせ
    if ($coverVersion -and $histVersion -and (Compare-VersionText $coverVersion $histVersion) -ne 0) {
        Add-Issue '警告' '版数' '表紙の版数' ("表紙の版数（{0}）と、改訂履歴の最後の版数（{1}）が違います" -f $coverVersion, $histVersion) -Sheet $cover.Name -Cell $coverVersionCell
    }
    if ($coverDate -and $histDate -and $coverDate -ne $histDate) {
        Add-Issue '警告' '版数' '表紙の更新日' ("表紙の更新日（{0}）と、改訂履歴の最後の改訂日（{1}）が違います" -f $coverDate, $histDate) -Sheet $cover.Name -Cell $coverDateCell
    }
    $latest = $histVersion
    if (-not $latest) { $latest = $coverVersion }
    $fm = [regex]::Match([System.IO.Path]::GetFileNameWithoutExtension($FileName).Normalize([System.Text.NormalizationForm]::FormKC), '(?i)(?:(?<![a-z])v(?:er)?\.?\s*|第\s*)(\d+(?:\.\d+)*)')
    if ($latest -and $fm.Success) {
        $fileVersion = ConvertTo-VersionText $fm.Groups[1].Value
        if ((Compare-VersionText $fileVersion $latest) -ne 0) {
            $where = '改訂履歴の最後の版数'
            if (-not $histVersion) { $where = '表紙の版数' }
            Add-Issue '警告' '版数' 'ファイル名の版数' ("ファイル名の版数（{0}）と、{1}（{2}）が違います" -f $fileVersion, $where, $latest)
        }
    }
}

function Get-ExpectedCell {
    #     見出しを固定（ウィンドウ枠の固定）しているシートで、スクロール領域の左上になるセル
    param($Pane)
    if ($Pane.TopLeft) { return $Pane.TopLeft }
    return (ConvertTo-ColumnName ([int]$Pane.XSplit + 1)) + ([int]$Pane.YSplit + 1)
}

# ---------------------------------------------------------------- 1 ファイルのチェック


function Test-Workbook {
    param([string] $FilePath)

    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) {
            if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break }
        }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        if ($null -eq $wbDoc) { throw 'Excel ブックではありません（workbook.xml が見つかりません）' }

        $wbRels = Get-PartRelationships -Package $pkg -PartName $wbPart
        $relById = @{}
        foreach ($rel in $wbRels) { $relById[$rel.Id] = $rel }

        # 印刷範囲（シートの番号 → 範囲）
        $printAreas = @{}
        foreach ($dn in $wbDoc.SelectNodes("//*[local-name()='definedNames']/*[local-name()='definedName']")) {
            if ($dn.GetAttribute('name') -ne '_xlnm.Print_Area' -or -not $dn.GetAttribute('localSheetId')) { continue }
            $printAreas[[int]$dn.GetAttribute('localSheetId')] = Get-PrintAreas $dn.InnerText
        }
        $sheets = New-Object System.Collections.Generic.List[object]
        $index = 0
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            $state = $s.GetAttribute('state')
            if (-not $state) { $state = 'visible' }
            $areas = $printAreas[$index]
            if (-not $areas) { $areas = New-Object System.Collections.Generic.List[object] }
            $isWorksheet = $rel -and (@('worksheet', 'macrosheet') -contains $rel.Kind)
            $sheets.Add([pscustomobject]@{ Name = $s.GetAttribute('name'); State = $state; Rel = $rel; Facts = $null; PrintAreas = $areas; IsWorksheet = $isWorksheet })
            $index++
        }

        # ---- 非表示シート
        if ($checkSet['Hidden']) {
            foreach ($s in $sheets) {
                if ($s.State -eq 'hidden') { Add-Issue '警告' '消し忘れ' '非表示シート' "非表示のシート「$($s.Name)」があります" -Sheet $s.Name }
                elseif ($s.State -eq 'veryHidden') { Add-Issue '警告' '消し忘れ' '非表示シート' "VBA でしか再表示できない非表示のシート「$($s.Name)」があります（veryHidden）" -Sheet $s.Name }
            }
        }

        # ---- シートごとに読む
        $needFacts = $checkSet['View'] -or $checkSet['Hidden'] -or $checkSet['Error'] -or $checkSet['Link'] -or $checkSet['Print']
        foreach ($s in $sheets) {
            if ($null -eq $s.Rel) { continue }
            $sheetRels = Get-PartRelationships -Package $pkg -PartName $s.Rel.Target
            if ($needFacts) { $s.Facts = Read-SheetFacts -Package $pkg -PartName $s.Rel.Target -PrintAreas $s.PrintAreas }
            $f = $s.Facts
            $isWorksheet = @('worksheet', 'macrosheet') -contains $s.Rel.Kind

            if ($checkSet['Hidden'] -and $f -and $s.State -eq 'visible') {
                if ($f.HiddenRows -gt 0) { Add-Issue '情報' '消し忘れ' '非表示の行・列' "非表示の行が $($f.HiddenRows) 行あります" -Sheet $s.Name }
                if ($f.HiddenCols -gt 0) { Add-Issue '情報' '消し忘れ' '非表示の行・列' "非表示の列が $($f.HiddenCols) 列あります" -Sheet $s.Name }
            }

            if ($checkSet['Error'] -and $f -and $f.Errors.Count -gt 0) {
                $n = 0
                foreach ($e in $f.Errors) {
                    $n++
                    if ($n -le $maxPerSheet) { Add-Issue 'エラー' '壊れ' 'エラー値' "セルがエラー値（$($e.Value)）です" -Sheet $s.Name -Cell $e.Cell }
                }
                if ($f.Errors.Count -gt $maxPerSheet) {
                    Add-Issue 'エラー' '壊れ' 'エラー値' ("ほかにもエラー値のセルが {0} 個あります" -f ($f.Errors.Count - $maxPerSheet)) -Sheet $s.Name
                }
            }

            if ($checkSet['Link'] -and $f) {
                $linkRels = @{}
                foreach ($r in @($sheetRels | Where-Object { $_.Kind -eq 'hyperlink' })) { $linkRels[$r.Id] = $r }
                foreach ($l in $f.Links) {
                    if (-not $l.RelId) { continue }
                    # 外部リンクの Target は Resolve-PartName を通さない生の値が必要なので、rels を直接読む
                    $target = Get-ExternalTarget -Package $pkg -SourcePart $s.Rel.Target -RelId $l.RelId
                    if ($target -match '^(file:|\\\\|[A-Za-z]:\\)') {
                        Add-Issue '警告' '情報漏れ' 'ファイルへのリンク' "社内のファイルやフォルダへのリンクがあります（$target）" -Sheet $s.Name -Cell $l.Ref
                    }
                }
            }

            if ($checkSet['Comment']) {
                # スレッド形式のコメントにも互換用のメモ（comments）が必ず付くので、メモだけを数える
                foreach ($rel in @($sheetRels | Where-Object { $_.Kind -eq 'comments' })) {
                    $doc = Read-PartXml -Package $pkg -PartName $rel.Target
                    if ($null -eq $doc) { continue }
                    foreach ($c in $doc.SelectNodes("//*[local-name()='comment']")) {
                        $text = $c.InnerText
                        if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '…' }
                        Add-Issue '警告' '消し忘れ' 'コメント' "コメント（メモ）が残っています: $text" -Sheet $s.Name -Cell $c.GetAttribute('ref')
                    }
                }
            }

            # ---- 表示の状態（表示されているシートだけ）
            if ($checkSet['View'] -and $f -and $f.HasView -and $s.State -eq 'visible') {
                if ($f.Zoom -ne $Zoom) { Add-Issue '警告' '体裁' '倍率' "倍率が $($f.Zoom)% です（$Zoom% にしてください）" -Sheet $s.Name }
                if ($f.ViewMode -ne 'normal') {
                    $modeName = @{ pageBreakPreview = '改ページプレビュー'; pageLayout = 'ページレイアウト' }[$f.ViewMode]
                    if (-not $modeName) { $modeName = $f.ViewMode }
                    Add-Issue '情報' '体裁' '表示モード' "表示が「$modeName」です" -Sheet $s.Name
                }

                if ($isWorksheet) {
                    $frozen = $f.Pane -and ($f.Pane.State -eq 'frozen' -or $f.Pane.State -eq 'frozenSplit')
                    # 選択セル: 見出しを固定しているシートは、スクロール領域の左上でもよい
                    $okCells = @('A1')
                    if ($frozen) { $okCells += (Get-ExpectedCell $f.Pane) }
                    $active = 'A1'
                    $activePane = ''
                    if ($f.Pane) { $activePane = $f.Pane.ActivePane }
                    $sel = @($f.Selections | Where-Object { $_.Pane -eq $activePane })
                    if ($sel.Count -eq 0) { $sel = $f.Selections.ToArray() }
                    if ($sel.Count -gt 0) { $active = $sel[0].ActiveCell }
                    if ($okCells -notcontains $active) {
                        Add-Issue '警告' '体裁' '選択セル' "選択セルが $active です（A1 にしてください）" -Sheet $s.Name -Cell $active
                    }
                    # 表示位置: 左上が A1、見出しを固定している場合はスクロール領域が固定位置の直後
                    $scrolled = $f.TopLeft -ne 'A1'
                    if ($frozen -and $f.Pane.TopLeft -and $f.Pane.TopLeft -ne (ConvertTo-ColumnName ([int]$f.Pane.XSplit + 1)) + ([int]$f.Pane.YSplit + 1)) { $scrolled = $true }
                    if ($scrolled) { Add-Issue '警告' '体裁' '表示位置' 'スクロールした状態で保存されています（左上を表示してください）' -Sheet $s.Name }
                }
            }
        }

        if ($checkSet['View']) {
            # 先頭シート: 最初の「表示されている」シートが開いた状態か
            $bv = $wbDoc.SelectSingleNode("//*[local-name()='bookViews']/*[local-name()='workbookView']")
            $activeTab = 0
            if ($null -ne $bv -and $bv.GetAttribute('activeTab')) { $activeTab = [int]$bv.GetAttribute('activeTab') }
            $firstVisible = -1
            for ($i = 0; $i -lt $sheets.Count; $i++) { if ($sheets[$i].State -eq 'visible') { $firstVisible = $i; break } }
            if ($firstVisible -ge 0 -and $activeTab -ne $firstVisible -and $activeTab -lt $sheets.Count) {
                Add-Issue '警告' '体裁' '先頭シート' ("開いたときに先頭ではないシート「{0}」が表示されます（「{1}」を表示してください）" -f $sheets[$activeTab].Name, $sheets[$firstVisible].Name) -Sheet $sheets[$activeTab].Name
            }
            # 作業グループ: 複数のシートが選択されたまま保存すると、開いた人の編集が全シートに及ぶ
            $grouped = @($sheets | Where-Object { $_.Facts -and $_.Facts.TabSelected })
            if ($grouped.Count -gt 1) {
                Add-Issue '警告' '体裁' '作業グループ' ("複数のシートが選択されたままです（作業グループ）: {0}" -f (($grouped | ForEach-Object { $_.Name }) -join '、'))
            }
        }

        if ($checkSet['Print']) {
            $stylesPart = @($wbRels | Where-Object { $_.Kind -eq 'styles' } | ForEach-Object { $_.Target })[0]
            Test-PrintSetup -Sheets $sheets -Mdw (Get-MaxDigitWidth $pkg $stylesPart)
        }

        # ---- ブック全体
        if ($checkSet['Link']) {
            foreach ($rel in @($wbRels | Where-Object { $_.Kind -eq 'externalLink' })) {
                $target = Get-ExternalTarget -Package $pkg -SourcePart $rel.Target -Kind 'externalLinkPath'
                if (-not $target) { $target = '（参照先不明）' }
                Add-Issue '警告' '情報漏れ' '外部リンク' "他のブックを参照しています（$target）"
            }
            foreach ($dn in $wbDoc.SelectNodes("//*[local-name()='definedNames']/*[local-name()='definedName']")) {
                $name = $dn.GetAttribute('name')
                if ($name -like '_xlnm.*') { continue }   # 印刷範囲などの組み込みの名前
                if ($dn.InnerText -match '#REF!') { Add-Issue 'エラー' '壊れ' '名前の定義' "名前「$name」の参照先が壊れています（#REF!）" }
                elseif ($dn.InnerText -match '\[\d+\]') { Add-Issue '警告' '情報漏れ' '名前の定義' "名前「$name」が他のブックを参照しています" }
            }
        }

        if ($checkSet['Property']) { Test-PackageProperties $pkg }

        if ($checkSet['Macro'] -and (Get-PackageEntry -Package $pkg -PartName 'xl/vbaProject.bin')) {
            Add-Issue '情報' 'マクロ' 'マクロ' 'マクロ（VBA）が含まれています'
        }
    }
    finally {
        $pkg.Zip.Dispose()
    }
    if ($checkSet['Version']) { Test-VersionConsistency -FilePath $FilePath -Sheets $sheets -FileName ([System.IO.Path]::GetFileName($FilePath)) }
}

function Test-PackageProperties {
    #     ファイルのプロパティ（docProps）に残っている個人名・会社名（Excel・Word 共通）
    param($Package)
    foreach ($part in @('docProps/core.xml', 'docProps/app.xml')) {
        $doc = Read-PartXml -Package $Package -PartName $part
        if ($null -eq $doc) { continue }
        foreach ($key in $propertyLabels.Keys) {
            $node = $doc.SelectSingleNode("//*[local-name()='$key']")
            if ($null -eq $node) { continue }
            $value = $node.InnerText.Trim()
            if (-not $value -or (Test-AllowedName $value)) { continue }
            Add-Issue '警告' '情報漏れ' 'プロパティ' ("ファイルのプロパティに{0}「{1}」が残っています" -f $propertyLabels[$key], $value)
        }
    }
}

# ---------------------------------------------------------------- Word（.docx / .docm）


function Get-WordText {
    #     段落・ラン（w:p・w:r）の文字（w:t。削除した文字 w:delText は含めない）
    param($Node, $Nsm)
    return ((@($Node.SelectNodes('.//w:t', $Nsm) | ForEach-Object { $_.InnerText })) -join '').Trim()
}

function Test-WordDocument {
    #     Word の提出前チェック: 変更履歴・コメント・隠し文字・蛍光ペン・残っている文言・社内のファイルへのリンク・テンプレートの場所・プロパティ・マクロ
    param([string] $FilePath)
    $ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    $pkg = Open-ZipPackage $FilePath
    try {
        $doc = Read-PartXml -Package $pkg -PartName 'word/document.xml'
        if ($null -eq $doc) { throw 'Word の本文（word/document.xml）がありません' }
        $nsm = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
        $nsm.AddNamespace('w', $ns)
        # 本文・ヘッダー・フッター・脚注・文末脚注
        $parts = New-Object System.Collections.Generic.List[object]
        $parts.Add([pscustomobject]@{ Name = '本文'; Part = 'word/document.xml'; Doc = $doc })
        foreach ($e in @($pkg.Entries.Keys | Sort-Object)) {
            $name = $null
            if ($e -match '^word/header\d*\.xml$') { $name = 'ヘッダー' }
            elseif ($e -match '^word/footer\d*\.xml$') { $name = 'フッター' }
            elseif ($e -eq 'word/footnotes.xml') { $name = '脚注' }
            elseif ($e -eq 'word/endnotes.xml') { $name = '文末脚注' }
            if (-not $name) { continue }
            $d = Read-PartXml -Package $pkg -PartName $e
            if ($d) { $parts.Add([pscustomobject]@{ Name = $name; Part = $e; Doc = $d }) }
        }
        $count = { param([string] $XPath) $n = 0; foreach ($p in $parts) { $n += $p.Doc.SelectNodes($XPath, $nsm).Count }; $n }

        if ($checkSet['Revision']) {
            $kinds = [ordered]@{ '挿入' = (& $count '//w:ins'); '削除' = (& $count '//w:del'); '移動' = (& $count '//w:moveFrom')
                '書式の変更' = (& $count '//w:rPrChange | //w:pPrChange | //w:tblPrChange | //w:trPrChange | //w:tcPrChange | //w:sectPrChange') }
            $found = @($kinds.Keys | Where-Object { $kinds[$_] -gt 0 } | ForEach-Object { '{0} {1}' -f $_, $kinds[$_] })
            if ($found.Count -gt 0) { Add-Issue '警告' '消し忘れ' '変更履歴' ("変更履歴が残っています（{0}）。すべて承諾するか元に戻してから提出してください" -f ($found -join '・')) }
            $settings = Read-PartXml -Package $pkg -PartName 'word/settings.xml'
            if ($settings -and $settings.SelectSingleNode('//*[local-name()="trackRevisions"]')) { Add-Issue '情報' '消し忘れ' '変更履歴' '変更履歴の記録がオンのまま保存されています' }
        }

        if ($checkSet['Comment']) {
            $comments = Read-PartXml -Package $pkg -PartName 'word/comments.xml'
            if ($comments) {
                $list = @($comments.SelectNodes('//*[local-name()="comment"]'))
                $n = 0
                foreach ($c in $list) {
                    $n++
                    if ($n -gt $maxPerSheet) { break }
                    $text = ((@($c.SelectNodes('.//*[local-name()="t"]') | ForEach-Object { $_.InnerText })) -join '').Trim()
                    if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '…' }
                    $author = $c.GetAttribute('author', $ns)
                    Add-Issue '警告' '消し忘れ' 'コメント' ("コメントが残っています（{0}）: {1}" -f $author, $text)
                }
                if ($list.Count -gt $maxPerSheet) { Add-Issue '警告' '消し忘れ' 'コメント' ("ほかにもコメントが {0} 件あります" -f ($list.Count - $maxPerSheet)) }
            }
            # 蛍光ペン（レビューの印のまま残りやすい）
            $marks = 0
            foreach ($p in $parts) { foreach ($h in $p.Doc.SelectNodes('//w:rPr/w:highlight', $nsm)) { if ($h.GetAttribute('val', $ns) -ne 'none') { $marks++ } } }
            if ($marks -gt 0) { Add-Issue '情報' '消し忘れ' '蛍光ペン' ("蛍光ペンの色が {0} か所に残っています（レビューの印なら消してください）" -f $marks) }
        }

        if ($checkSet['Hidden']) {
            foreach ($p in $parts) {
                $runs = @($p.Doc.SelectNodes('//w:r[w:rPr/w:vanish]', $nsm) | Where-Object {
                        $v = $_.SelectSingleNode('w:rPr/w:vanish', $nsm).GetAttribute('val', $ns)
                        $v -notin @('0', 'false', 'off') -and (Get-WordText $_ $nsm)
                    })
                if ($runs.Count -eq 0) { continue }
                $text = (@($runs | Select-Object -First 3 | ForEach-Object { Get-WordText $_ $nsm }) -join ' / ')
                if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '…' }
                Add-Issue '警告' '消し忘れ' '隠し文字' ("隠し文字が {0} か所あります: {1}" -f $runs.Count, $text) -Sheet $p.Name
            }
        }

        if ($checkSet['Word'] -and $Words) {
            foreach ($p in $parts) {
                $index = 0; $n = 0
                foreach ($para in $p.Doc.SelectNodes('//w:p', $nsm)) {
                    $index++
                    $text = Get-WordText $para $nsm
                    if (-not $text) { continue }
                    $m = [regex]::Match($text, $Words, 'IgnoreCase')
                    if (-not $m.Success) { continue }
                    $n++
                    if ($n -gt $maxPerSheet) { continue }
                    if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '…' }
                    Add-Issue '警告' '消し忘れ' '残っている文言' ("「{0}」が残っています: {1}" -f $m.Value, $text) -Sheet $p.Name -Cell ('段落 {0}' -f $index)
                }
                if ($n -gt $maxPerSheet) { Add-Issue '警告' '消し忘れ' '残っている文言' ("ほかにも消し忘れの文言のある段落が {0} 個あります" -f ($n - $maxPerSheet)) -Sheet $p.Name }
            }
        }

        if ($checkSet['Link']) {
            foreach ($p in $parts) {
                $slash = $p.Part.LastIndexOf('/')
                $rels = Read-PartXml -Package $pkg -PartName ($p.Part.Substring(0, $slash) + '/_rels/' + $p.Part.Substring($slash + 1) + '.rels')
                if (-not $rels) { continue }
                foreach ($r in $rels.SelectNodes('//*[local-name()="Relationship"]')) {
                    $target = $r.GetAttribute('Target')
                    if ($r.GetAttribute('TargetMode') -eq 'External' -and $target -match '^(file:|\\\\|[A-Za-z]:\\)') {
                        Add-Issue '警告' '情報漏れ' 'ファイルへのリンク' "社内のファイルやフォルダへのリンクがあります（$target）" -Sheet $p.Name
                    }
                }
            }
            # 文書のテンプレートの場所（ユーザー名やサーバー名が入りやすい）
            $settingsRels = Read-PartXml -Package $pkg -PartName 'word/_rels/settings.xml.rels'
            if ($settingsRels) {
                foreach ($r in $settingsRels.SelectNodes('//*[local-name()="Relationship"]')) {
                    $target = $r.GetAttribute('Target')
                    if ($r.GetAttribute('Type') -like '*/attachedTemplate' -and $target -match '^(file:|\\\\|[A-Za-z]:\\)') {
                        Add-Issue '情報' '情報漏れ' 'テンプレートの場所' "文書のテンプレートの場所が残っています（$target）"
                    }
                }
            }
        }

        if ($checkSet['Property']) { Test-PackageProperties $pkg }
        if ($checkSet['Macro'] -and (Get-PackageEntry -Package $pkg -PartName 'word/vbaProject.bin')) { Add-Issue '情報' 'マクロ' 'マクロ' 'マクロ（VBA）が含まれています' }
    }
    finally { $pkg.Zip.Dispose() }
}

function Get-ExternalTarget {
    #     リレーションの Target をそのまま返す（外部のパスや URL はパーツ名として解決しない）。
    #     RelId を指定するとその Id の Target、Kind を指定すると種類が一致する最初の Target。
    param($Package, [string] $SourcePart, [string] $RelId, [string] $Kind)

    $slash = $SourcePart.LastIndexOf('/')
    $relsName = $SourcePart.Substring(0, $slash) + '/_rels/' + $SourcePart.Substring($slash + 1) + '.rels'
    $doc = Read-PartXml -Package $Package -PartName $relsName
    if ($null -eq $doc) { return $null }
    foreach ($r in $doc.SelectNodes("//*[local-name()='Relationship']")) {
        $type = $r.GetAttribute('Type')
        if ($RelId -and $r.GetAttribute('Id') -eq $RelId) { return $r.GetAttribute('Target') }
        if ($Kind -and $type.EndsWith('/' + $Kind)) { return $r.GetAttribute('Target') }
    }
    return $null
}

function Get-FileKind {
    # 先頭 4 バイトで実体を判定する
    param([string] $FilePath)
    $share = [System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete
    $fs = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, $share)
    try {
        $buf = New-Object byte[] 4
        if ($fs.Read($buf, 0, 4) -lt 4) { return 'Unknown' }
        if ($buf[0] -eq 0x50 -and $buf[1] -eq 0x4B) { return 'Zip' }
        if ($buf[0] -eq 0xD0 -and $buf[1] -eq 0xCF -and $buf[2] -eq 0x11 -and $buf[3] -eq 0xE0) { return 'Ole' }
        return 'Unknown'
    }
    finally { $fs.Dispose() }
}

# ---------------------------------------------------------------- 対象ファイルの列挙と実行


$files = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($p in $Path) {
    $items = @(Get-Item -Path $p -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
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
            if (@('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb', '.docx', '.docm') -notcontains $f.Extension.ToLowerInvariant()) { continue }
            if ($f.Name.StartsWith('~$')) { continue }
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $files.Add([pscustomobject]@{ Info = $f; Relative = $f.FullName.Substring($root.Length).TrimStart('\', '/') })
        }
    }
}
if ($files.Count -eq 0) { Write-Warning 'チェックする Excel・Word ファイルがありません。'; return }

function Invoke-FileChecks {
    #     1 ファイルをチェックし、指摘を $script:Issues に足す
    param($Entry)
    $script:CurrentFile = $Entry.Relative
    $script:CurrentPath = $Entry.Info.FullName

    if ($FileNamePattern -and $Entry.Info.Name -notmatch $FileNamePattern) {
        Add-Issue '警告' 'ファイル' 'ファイル名' "ファイル名が命名規則（$FileNamePattern）に合っていません"
    }

    $ext = $Entry.Info.Extension.ToLowerInvariant()
    try {
        $kind = Get-FileKind -FilePath $Entry.Info.FullName
        if (@('.docx', '.docm') -contains $ext) {
            if ($kind -eq 'Ole') { Add-Issue '情報' 'ファイル' 'パスワード' 'パスワードで保護（暗号化）されているためチェックしていません'; return }
            if ($kind -ne 'Zip') { Add-Issue '情報' 'ファイル' '形式' 'docx / docm 形式ではないためチェックしていません'; return }
            Test-WordDocument -FilePath $Entry.Info.FullName
            return
        }
        if ($ext -eq '.xlsb' -or ($kind -eq 'Ole' -and $ext -eq '.xls') -or $kind -eq 'Unknown') {
            Add-Issue '情報' 'ファイル' '形式' 'xlsx / xlsm 形式ではないためチェックしていません（xlsx に保存し直すとチェックできます）'
            return
        }
        if ($kind -eq 'Ole') {
            Add-Issue '情報' 'ファイル' 'パスワード' 'パスワードで保護（暗号化）されているためチェックしていません'
            return
        }

        Test-Workbook -FilePath $Entry.Info.FullName

        if ($checkSet['Word'] -and $Words) {
            $hits = @(& $searchScript -Pattern $Words -Path ([System.Management.Automation.WildcardPattern]::Escape($Entry.Info.FullName)) -Target Value, Shape -NoExcel -WarningAction SilentlyContinue 6>$null)
            foreach ($h in $hits) {
                $m = [regex]::Match($h.Text, $Words, 'IgnoreCase')
                $text = $h.Text
                if ($text.Length -gt 40) { $text = $text.Substring(0, 40) + '…' }
                Add-Issue '警告' '消し忘れ' '残っている文言' ("「{0}」が残っています: {1}" -f $m.Value, $text) -Sheet $h.Sheet -Cell $h.Cell
            }
        }
    }
    catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $msg = $ex.Message
        if ($ex -is [System.IO.InvalidDataException]) { $msg = "ファイルが壊れているか Excel 形式ではありません（$msg）" }
        Add-Issue 'エラー' 'ファイル' '読み取り' "読み取れませんでした: $msg"
    }
}

# ---------------------------------------------------------------- 機械で直せる指摘を直す（-Fix）


function Repair-Properties {
    #     ファイルのプロパティの作成者・最終更新者・会社名・管理者（-AllowedNames に一致しないもの）を消す。返り値: 消した項目の名前
    param($Package, [hashtable] $Changes)
    $cleared = New-Object System.Collections.Generic.List[string]
    foreach ($part in @('docProps/core.xml', 'docProps/app.xml')) {
        $doc = Read-PartXmlForEdit -Package $Package -PartName $part
        if ($null -eq $doc) { continue }
        $edited = $false
        foreach ($key in $propertyLabels.Keys) {
            $node = $doc.SelectSingleNode("//*[local-name()='$key']")
            if ($null -eq $node) { continue }
            $value = $node.InnerText.Trim()
            if (-not $value -or (Test-AllowedName $value)) { continue }
            $node.InnerText = ''
            $cleared.Add($propertyLabels[$key])
            $edited = $true
        }
        if ($edited) { $Changes[$part.ToLowerInvariant()] = ConvertTo-XmlString $doc }
    }
    return , $cleared
}

function Repair-HeaderFooterPath {
    #     ヘッダー・フッターから、ファイルの場所（&Z）を消す（文字の &&Z はそのまま）。返り値: 直したシートの数
    param($Package, [hashtable] $Changes)
    $count = 0
    foreach ($part in (Get-WorkbookSheetParts $Package).Sheets.Values) {
        $doc = Read-PartXmlForEdit -Package $Package -PartName $part
        if ($null -eq $doc) { continue }
        $edited = $false
        foreach ($hf in @($doc.SelectNodes("//*[local-name()='headerFooter']/*"))) {
            $new = [regex]::Replace($hf.InnerText, '(?<!&)((?:&&)*)&[Zz]', '$1')
            if ($new -cne $hf.InnerText) { $hf.InnerText = $new; $edited = $true }
        }
        if ($edited) { $Changes[$part.ToLowerInvariant()] = ConvertTo-XmlString $doc; $count++ }
    }
    return $count
}

function Repair-Deliverable {
    #     機械で直せる指摘を直す。返り値: 直したことの説明
    param($Entry, $Issues)
    $done = New-Object System.Collections.Generic.List[string]
    $path = $Entry.Info.FullName
    $changes = @{}
    $pkg = Open-ZipPackage -FilePath $path
    try {
        if (@($Issues | Where-Object Item -eq 'プロパティ').Count -gt 0) {
            $cleared = Repair-Properties $pkg $changes
            if ($cleared.Count -gt 0) { $done.Add(('プロパティの{0}を消去' -f ($cleared -join '・'))) }
        }
        if (@($Issues | Where-Object { $_.Item -eq 'ヘッダー／フッター' -and $_.Detail -like '*&Z*' }).Count -gt 0) {
            if ((Repair-HeaderFooterPath $pkg $changes) -gt 0) { $done.Add('ヘッダー／フッターのファイルの場所（&Z）を削除') }
        }
    }
    finally { $pkg.Zip.Dispose() }
    if ($changes.Count -gt 0) { Save-ZipPackage -FilePath $path -Changes $changes }
    # 表示の状態（倍率・選択セル・先頭シートなど）は Reset-ExcelView.ps1 で直す
    if (@($Issues | Where-Object { $_.Category -eq '体裁' -and $_.Item -ne '表示モード' }).Count -gt 0) {
        $r = @(& $resetScript -Path ([System.Management.Automation.WildcardPattern]::Escape($path)) -Zoom $Zoom -Confirm:$false -WarningAction SilentlyContinue 6>$null)
        if (@($r | Where-Object Result -eq '修正').Count -gt 0) { $done.Add('表示の状態（倍率・選択セル・先頭シートなど）を直した') }
    }
    return , $done
}

function Test-FileLocked {
    # Excel で開いているファイルは書き換えられないので、先に確かめる
    param([string] $FilePath)
    try {
        $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Dispose()
        return $false
    }
    catch { return $true }
}

$script:Fixes = New-Object System.Collections.Generic.List[object]
$index = 0
foreach ($entry in $files) {
    $index++
    Write-Progress -Activity '提出前チェック' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) -PercentComplete ([int](100 * $index / $files.Count))
    Invoke-FileChecks $entry
    if (-not $Fix) { continue }

    # 機械で直せる指摘: プロパティの個人名・会社名、ヘッダー／フッターのファイルの場所、表示の状態
    $path = $entry.Info.FullName
    $fixable = @($script:Issues | Where-Object { $_.Path -eq $path -and ($_.Item -eq 'プロパティ' -or ($_.Item -eq 'ヘッダー／フッター' -and $_.Detail -like '*&Z*') -or ($_.Category -eq '体裁' -and $_.Item -ne '表示モード')) })
    if ($fixable.Count -eq 0) { continue }
    $what = @($fixable | ForEach-Object { $_.Item } | Sort-Object -Unique) -join '・'
    if ($entry.Info.IsReadOnly) {
        $script:Fixes.Add([pscustomobject]@{ File = $entry.Relative; Result = 'スキップ'; Detail = '読み取り専用のファイルです' })
        continue
    }
    if (-not $PSCmdlet.ShouldProcess($entry.Relative, "指摘を直す（$what）")) {
        $script:Fixes.Add([pscustomobject]@{ File = $entry.Relative; Result = '修正予定'; Detail = $what })
        continue
    }
    try {
        if (Test-FileLocked $path) { throw 'ファイルを開いているため書き換えられません（Excel で閉じてから実行してください）' }
        if ($BackupDir) {
            $dest = Join-Path $BackupDir $entry.Relative
            New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
            Copy-Item -LiteralPath $path -Destination $dest -Force
        }
        $done = Repair-Deliverable $entry $fixable
        # 直したあとの状態でチェックし直す（出力するのは残っている指摘）
        $keep = @($script:Issues | Where-Object { $_.Path -ne $path })
        $script:Issues.Clear()
        foreach ($k in $keep) { $script:Issues.Add($k) }
        Invoke-FileChecks $entry
        $script:Fixes.Add([pscustomobject]@{ File = $entry.Relative; Result = '修正'; Detail = ($done -join '・') })
    }
    catch {
        $script:Fixes.Add([pscustomobject]@{ File = $entry.Relative; Result = '失敗'; Detail = $_.Exception.Message })
        Write-Warning ("{0}: 直せませんでした: {1}" -f $entry.Relative, $_.Exception.Message)
    }
}
Write-Progress -Activity '提出前チェック' -Completed

# ---------------------------------------------------------------- 結果


$levelOrder = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
$categoryOrder = @{ 'ファイル' = 0; '壊れ' = 1; '情報漏れ' = 2; '版数' = 3; '消し忘れ' = 4; '体裁' = 5; '印刷' = 6; 'マクロ' = 7 }
$sorted = @($script:Issues | Sort-Object @{ e = { $_.File } }, @{ e = { $levelOrder[$_.Level] } }, @{ e = { $categoryOrder[$_.Category] } }, @{ e = { $_.Sheet } })
$sorted

$byFile = @{}
foreach ($i in $script:Issues) {
    if (-not $byFile.ContainsKey($i.Path)) { $byFile[$i.Path] = @{ 'エラー' = 0; '警告' = 0; '情報' = 0 } }
    $byFile[$i.Path][$i.Level]++
}
$ngFiles = @($byFile.Keys | Where-Object { $byFile[$_]['エラー'] + $byFile[$_]['警告'] -gt 0 }).Count

if ($OutXlsx) {
    $issueRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.File, $i.Level, $i.Category, $i.Item, $i.Location, $i.Detail, $i.Path) }
        # 場所が無い指摘（ブック全体）もファイルが開けるよう、先頭シートの A1 にリンクする
        $sheet = $i.Sheet
        $cell = $i.Cell
        if (-not $cell) { $cell = 'A1' }
        if ($sheet -and $i.Path -notmatch '\.doc[xm]$') { $row.Link = @{ Column = 4; Path = $i.Path; Location = ("'{0}'!{1}" -f $sheet.Replace("'", "''"), $cell) } }
        else { $row.Link = @{ Column = 0; Path = $i.Path; Location = '' } }
        $issueRows.Add($row)
    }
    $fileRows = New-Object System.Collections.Generic.List[object]
    foreach ($entry in $files) {
        $c = $byFile[$entry.Info.FullName]
        if (-not $c) { $c = @{ 'エラー' = 0; '警告' = 0; '情報' = 0 } }
        $result = 'OK'
        if ($c['エラー'] + $c['警告'] -gt 0) { $result = '要対応' }
        $fileRows.Add(@{ Values = @($entry.Relative, $result, $c['エラー'], $c['警告'], $c['情報'], $entry.Info.FullName); Link = @{ Column = 0; Path = $entry.Info.FullName; Location = '' } })
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{
                Name = '指摘一覧'
                Columns = @(@{ Header = 'ファイル'; Width = 40 }, @{ Header = '区分'; Width = 7 }, @{ Header = '分類'; Width = 10 }, @{ Header = '項目'; Width = 16 },
                    @{ Header = '場所'; Width = 22 }, @{ Header = '内容'; Width = 70 }, @{ Header = 'フルパス'; Width = 60 })
                Rows = $issueRows
            },
            @{
                Name = 'ファイル別'
                Columns = @(@{ Header = 'ファイル'; Width = 50 }, @{ Header = '結果'; Width = 9 }, @{ Header = 'エラー'; Width = 8 }, @{ Header = '警告'; Width = 8 },
                    @{ Header = '情報'; Width = 8 }, @{ Header = 'フルパス'; Width = 60 })
                Rows = $fileRows
            })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
Write-Host ("{0} ファイル中 {1} ファイルが要対応（エラー {2} 件 / 警告 {3} 件 / 情報 {4} 件）" -f $files.Count, $ngFiles,
    @($script:Issues | Where-Object Level -eq 'エラー').Count, @($script:Issues | Where-Object Level -eq '警告').Count, @($script:Issues | Where-Object Level -eq '情報').Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }
if ($Fix) {
    $fixed = @($script:Fixes | Where-Object Result -eq '修正')
    if ($script:Fixes.Count -eq 0) { Write-Host '機械で直せる指摘はありませんでした' }
    foreach ($x in $script:Fixes) { Write-Host ('  {0}  {1}: {2}' -f $x.Result, $x.File, $x.Detail) }
    if ($fixed.Count -gt 0) {
        Write-Host ('{0} ファイルを直しました（上の件数は直したあとに残っている指摘です）' -f $fixed.Count)
        if ($BackupDir) { Write-Host "直す前のファイル: $BackupDir" }
    }
}

# .SYNOPSIS
#     Excel・Word ファイルを社外へ提出する前に、体裁の崩れ・消し忘れ・情報漏れ・壊れた箇所がないかをまとめて確認します。
#
# .DESCRIPTION
#     フォルダ配下の xlsx / xlsm を Excel を使わずに読み、次の項目をチェックします。ファイルは変更しません。
#
#       体裁       選択セルが A1 でない、倍率が既定と違う、スクロールしたまま、先頭シートが表示されない、
#                  複数シートの選択（作業グループ）のまま、改ページプレビュー表示（情報）
#       消し忘れ   非表示シート、コメント（メモ）、TODO などの文言、非表示の行・列（情報）
#       情報漏れ   作成者・最終更新者・会社名などのプロパティ、他ブックへの外部リンク、社内パスへのハイパーリンク
#       壊れ       エラー値（#REF! など）のセル、参照先が壊れた名前の定義
#       版数       表紙の版数・更新日と、改訂履歴の最後の行・ファイル名の版数の食い違い、改訂履歴の版数の逆戻り
#       印刷       印刷すると横がはみ出す、印刷範囲の外のデータ、ヘッダー／フッターのファイルの場所（&Z）や TODO、
#                  用紙のサイズ・向きがシートによって違う（情報）
#       マクロ     VBA を含むブック（情報）
#
#     Word（.docx / .docm）は、Word を使わずに次の項目をチェックします（場所は「本文!段落 12」のように示します）。
#
#       消し忘れ   変更履歴（挿入・削除・移動・書式の変更）と、変更履歴の記録がオンのまま（情報）、コメント、隠し文字、
#                  TODO などの文言（本文・ヘッダー・フッター・脚注）、蛍光ペン（情報）
#       情報漏れ   作成者・最終更新者・会社名などのプロパティ、社内パスへのハイパーリンク、テンプレートの場所（情報）
#       マクロ     VBA を含む文書（情報）
#
#     結果は 1 指摘 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#     区分は エラー（必ず直す）／警告（原則直す）／情報（意図どおりか確認する）の 3 つです。
#
#     -Fix を付けると、機械で直せる指摘（プロパティの個人名・会社名、ヘッダー／フッターのファイルの場所、
#     表示の状態。Word はプロパティだけ）を直してから、直したあとの状態で残っている指摘を出力します（このときだけファイルを書き換えます）。
#
# .PARAMETER Path
#     チェックするフォルダまたはファイル（複数指定可、ワイルドカード可）。
#
# .PARAMETER Check
#     チェックする項目。View（体裁）/ Hidden（非表示・隠し文字）/ Comment（コメント・蛍光ペン）/ Revision（Word の変更履歴）/
#     Word（残っている文言）/ Property（プロパティ）/ Link（外部リンク・名前の定義・テンプレートの場所）/ Error（エラー値）/
#     Macro（マクロ）/ Print（印刷）/ Version（表紙と改訂履歴の版数）。既定はすべて。
#
# .PARAMETER Words
#     消し忘れとして探す文言の正規表現（大文字小文字は区別しません）。
#     既定は TODO / TBD / FIXME / 要確認 / 未定 / 検討中 / 仮置き / （仮） / XXX / ○○ / ××
#
# .PARAMETER AllowedNames
#     プロパティに残っていてもよい名前の正規表現（例: 自社名や部署名）。一致するものは指摘しません。
#
# .PARAMETER Zoom
#     期待する表示倍率（%）。既定は 100。
#
# .PARAMETER FileNamePattern
#     ファイル名の命名規則の正規表現。一致しないファイルを警告します（例: '^[A-Z]{2}-\d{2}-\d{2}-\d{4}_.+\.xlsx$'）。
#
# .PARAMETER RequirePageNumber
#     ヘッダー／フッターにページ番号（&P）が無いシートを警告します。
#
# .PARAMETER HeaderFooterText
#     ヘッダー／フッターに必ず入れる文字の正規表現（例: '社外秘|Confidential'）。無いシートを警告します。
#
# .PARAMETER HistorySheet
#     版数のチェックで使う、改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#     表紙は「表紙」という名前のシート、無ければ最初に表示されているシートです。
#
# .PARAMETER NoRecurse
#     サブフォルダをチェックしません。
#
# .PARAMETER Fix
#     機械で直せる指摘を直します。
#       プロパティ          作成者・最終更新者・会社名・管理者を消す（-AllowedNames に一致するものは残す）
#       ヘッダー／フッター  ファイルの場所（&Z）を消す（ファイル名 &F などはそのまま）
#       体裁                倍率・選択セル・表示位置・先頭シート・作業グループを直す（Reset-ExcelView.ps1 と同じ）
#     コメント・非表示シート・外部リンク・残っている文言など、内容が消える・変わるものは直しません。
#     読み取り専用のファイルは書き換えません。-WhatIf で直すファイルと項目を確かめられます。
#
# .PARAMETER BackupDir
#     -Fix で書き換える前のファイルを、このフォルダに同じ相対パスで残します。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘一覧・ファイル別。場所をクリックすると該当セルが開きます）。
#
# .EXAMPLE
#     .\Test-ExcelDeliverable.ps1 D:\提出物
#
# .EXAMPLE
#     # 自社名はプロパティに残っていてよいとし、結果を保存する
#     .\Test-ExcelDeliverable.ps1 D:\提出物 -AllowedNames '株式会社サンプル' -OutXlsx D:\提出前チェック.xlsx
#
# .EXAMPLE
#     # 体裁だけを確認する
#     .\Test-ExcelDeliverable.ps1 D:\提出物 -Check View
#
# .EXAMPLE
#     # 印刷と版数だけを、ページ番号と「社外秘」の表示も含めて確認する
#     .\Test-ExcelDeliverable.ps1 D:\提出物 -Check Print, Version -RequirePageNumber -HeaderFooterText '社外秘'
#
# .EXAMPLE
#     # 機械で直せる指摘を直す（直す前のファイルを残す）。まず -WhatIf で確かめる
#     .\Test-ExcelDeliverable.ps1 D:\提出物 -Fix -AllowedNames '株式会社サンプル' -BackupDir D:\直す前 -WhatIf
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
