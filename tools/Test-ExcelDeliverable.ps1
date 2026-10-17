[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateSet('View', 'Hidden', 'Comment', 'Word', 'Property', 'Link', 'Error', 'Macro')]
    [string[]] $Check = @('View', 'Hidden', 'Comment', 'Word', 'Property', 'Link', 'Error', 'Macro'),

    [string] $Words = 'TODO|TBD|FIXME|要確認|未定|検討中|仮置き|（仮）|\(仮\)|XXX|○○|××',

    [string] $AllowedNames,

    [ValidateRange(10, 400)]
    [int] $Zoom = 100,

    [string] $FileNamePattern,

    [switch] $NoRecurse,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

foreach ($p in @($Words, $AllowedNames, $FileNamePattern)) {
    if ($p) {
        try { [void][regex]::new($p) }
        catch { throw "正規表現が不正です: $p 詳細: $($_.Exception.InnerException.Message)" }
    }
}
if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'ExcelDeliverable.Issue' -DefaultDisplayPropertySet File, Location, Level, Detail -Force

$checkSet = @{}
foreach ($c in $Check) { $checkSet[$c] = $true }

$errorValues = @('#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!')
$maxPerSheet = 20   # エラー値などを 1 シートで個別に挙げる上限（超えた分はまとめて 1 件にする）

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

function Test-AllowedName {
    param([string] $Value)
    return ($AllowedNames -and $Value -match $AllowedNames)
}

# ---------------------------------------------------------------- シートの読み取り


function Read-SheetFacts {
    #     シートの XML を 1 回だけ先頭から読み、チェックに必要な情報を集める。
    #     表示（sheetView / pane / selection）、非表示の行・列、エラー値のセル、ハイパーリンク
    param($Package, [string] $PartName)

    $facts = [pscustomobject]@{
        HasView = $false; TabSelected = $false; Zoom = 100; TopLeft = 'A1'; ViewMode = 'normal'
        Pane = $null; Selections = New-Object System.Collections.Generic.List[object]
        HiddenRows = 0; HiddenCols = 0
        Errors = New-Object System.Collections.Generic.List[object]
        Links = New-Object System.Collections.Generic.List[object]
    }

    $reader = New-PartXmlReader -Package $Package -PartName $PartName
    if ($null -eq $reader) { return $facts }
    try {
        $inFirstView = $false
        $viewDone = $false
        $errorCell = $null
        $advance = $true
        while ($true) {
            if ($advance -and -not $reader.Read()) { break }
            $advance = $true
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
                if ($reader.LocalName -eq 'sheetView' -and $inFirstView) { $inFirstView = $false; $viewDone = $true }
                elseif ($reader.LocalName -eq 'c') { $errorCell = $null }
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
                }
                'row' {
                    if (@('1', 'true') -contains $reader.GetAttribute('hidden')) { $facts.HiddenRows++ }
                }
                'c' {
                    if ($reader.GetAttribute('t') -eq 'e' -and -not $reader.IsEmptyElement) { $errorCell = $reader.GetAttribute('r') }
                }
                'v' {
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

        $sheets = New-Object System.Collections.Generic.List[object]
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            $state = $s.GetAttribute('state')
            if (-not $state) { $state = 'visible' }
            $sheets.Add([pscustomobject]@{ Name = $s.GetAttribute('name'); State = $state; Rel = $rel; Facts = $null })
        }

        # ---- 非表示シート
        if ($checkSet['Hidden']) {
            foreach ($s in $sheets) {
                if ($s.State -eq 'hidden') { Add-Issue '警告' '消し忘れ' '非表示シート' "非表示のシート「$($s.Name)」があります" -Sheet $s.Name }
                elseif ($s.State -eq 'veryHidden') { Add-Issue '警告' '消し忘れ' '非表示シート' "VBA でしか再表示できない非表示のシート「$($s.Name)」があります（veryHidden）" -Sheet $s.Name }
            }
        }

        # ---- シートごとに読む
        $needFacts = $checkSet['View'] -or $checkSet['Hidden'] -or $checkSet['Error'] -or $checkSet['Link']
        foreach ($s in $sheets) {
            if ($null -eq $s.Rel) { continue }
            $sheetRels = Get-PartRelationships -Package $pkg -PartName $s.Rel.Target
            if ($needFacts) { $s.Facts = Read-SheetFacts -Package $pkg -PartName $s.Rel.Target }
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
                    if ($sel.Count -eq 0) { $sel = @($f.Selections) }
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

        if ($checkSet['Property']) {
            $labels = [ordered]@{ 'creator' = '作成者'; 'lastModifiedBy' = '最終更新者'; 'Company' = '会社名'; 'Manager' = '管理者' }
            foreach ($part in @('docProps/core.xml', 'docProps/app.xml')) {
                $doc = Read-PartXml -Package $pkg -PartName $part
                if ($null -eq $doc) { continue }
                foreach ($key in $labels.Keys) {
                    $node = $doc.SelectSingleNode("//*[local-name()='$key']")
                    if ($null -eq $node) { continue }
                    $value = $node.InnerText.Trim()
                    if (-not $value -or (Test-AllowedName $value)) { continue }
                    Add-Issue '警告' '情報漏れ' 'プロパティ' ("ファイルのプロパティに{0}「{1}」が残っています" -f $labels[$key], $value)
                }
            }
        }

        if ($checkSet['Macro'] -and (Get-PackageEntry -Package $pkg -PartName 'xl/vbaProject.bin')) {
            Add-Issue '情報' 'マクロ' 'マクロ' 'マクロ（VBA）が含まれています'
        }
    }
    finally {
        $pkg.Zip.Dispose()
    }
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
            if (@('.xlsx', '.xlsm', '.xltx', '.xltm', '.xls', '.xlsb') -notcontains $f.Extension.ToLowerInvariant()) { continue }
            if ($f.Name.StartsWith('~$')) { continue }
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $files.Add([pscustomobject]@{ Info = $f; Relative = $f.FullName.Substring($root.Length).TrimStart('\', '/') })
        }
    }
}
if ($files.Count -eq 0) { Write-Warning 'チェックする Excel ファイルがありません。'; return }

$index = 0
foreach ($entry in $files) {
    $index++
    $script:CurrentFile = $entry.Relative
    $script:CurrentPath = $entry.Info.FullName
    Write-Progress -Activity '提出前チェック' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) -PercentComplete ([int](100 * $index / $files.Count))

    if ($FileNamePattern -and $entry.Info.Name -notmatch $FileNamePattern) {
        Add-Issue '警告' 'ファイル' 'ファイル名' "ファイル名が命名規則（$FileNamePattern）に合っていません"
    }

    $ext = $entry.Info.Extension.ToLowerInvariant()
    try {
        $kind = Get-FileKind -FilePath $entry.Info.FullName
        if ($ext -eq '.xlsb' -or ($kind -eq 'Ole' -and $ext -eq '.xls') -or $kind -eq 'Unknown') {
            Add-Issue '情報' 'ファイル' '形式' 'xlsx / xlsm 形式ではないためチェックしていません（xlsx に保存し直すとチェックできます）'
            continue
        }
        if ($kind -eq 'Ole') {
            Add-Issue '情報' 'ファイル' 'パスワード' 'パスワードで保護（暗号化）されているためチェックしていません'
            continue
        }

        Test-Workbook -FilePath $entry.Info.FullName

        if ($checkSet['Word'] -and $Words) {
            $hits = @(& $searchScript -Pattern $Words -Path ([System.Management.Automation.WildcardPattern]::Escape($entry.Info.FullName)) -Target Value, Shape -NoExcel -WarningAction SilentlyContinue 6>$null)
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
Write-Progress -Activity '提出前チェック' -Completed

# ---------------------------------------------------------------- 結果


$levelOrder = @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }
$categoryOrder = @{ 'ファイル' = 0; '壊れ' = 1; '情報漏れ' = 2; '消し忘れ' = 3; '体裁' = 4; 'マクロ' = 5 }
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
        if ($sheet) { $row.Link = @{ Column = 4; Path = $i.Path; Location = ("'{0}'!{1}" -f $sheet.Replace("'", "''"), $cell) } }
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

# .SYNOPSIS
#     Excel ファイルを社外へ提出する前に、体裁の崩れ・消し忘れ・情報漏れ・壊れた箇所がないかをまとめて確認します。
#
# .DESCRIPTION
#     フォルダ配下の xlsx / xlsm を Excel を使わずに読み、次の項目をチェックします。ファイルは変更しません。
#
#       体裁       選択セルが A1 でない、倍率が既定と違う、スクロールしたまま、先頭シートが表示されない、
#                  複数シートの選択（作業グループ）のまま、改ページプレビュー表示（情報）
#       消し忘れ   非表示シート、コメント（メモ）、TODO などの文言、非表示の行・列（情報）
#       情報漏れ   作成者・最終更新者・会社名などのプロパティ、他ブックへの外部リンク、社内パスへのハイパーリンク
#       壊れ       エラー値（#REF! など）のセル、参照先が壊れた名前の定義
#       マクロ     VBA を含むブック（情報）
#
#     結果は 1 指摘 1 行のオブジェクトで出力し、-OutXlsx で該当セルへのリンク付きの一覧に保存できます。
#     区分は エラー（必ず直す）／警告（原則直す）／情報（意図どおりか確認する）の 3 つです。
#
# .PARAMETER Path
#     チェックするフォルダまたはファイル（複数指定可、ワイルドカード可）。
#
# .PARAMETER Check
#     チェックする項目。View（体裁）/ Hidden（非表示）/ Comment（コメント）/ Word（残っている文言）/
#     Property（プロパティ）/ Link（外部リンク・名前の定義）/ Error（エラー値）/ Macro（マクロ）。既定はすべて。
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
# .PARAMETER NoRecurse
#     サブフォルダをチェックしません。
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
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
