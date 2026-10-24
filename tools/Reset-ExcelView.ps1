[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [ValidateRange(10, 400)]
    [int] $Zoom = 100,

    [switch] $NormalView,

    [switch] $KeepTimestamp,

    [string] $BackupDir,

    [switch] $Force,

    [switch] $NoRecurse
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$checker = Join-Path $PSScriptRoot 'Test-ExcelDeliverable.ps1'

if ($BackupDir) {
    $BackupDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupDir)
}

# ---------------------------------------------------------------- XML の書き換え（文字列のまま、必要な属性だけ）


function Set-XmlAttribute {
    #     開始タグ（<tag ...> または <tag .../>）の属性を設定する。Value が $null なら属性を消す。
    #     Value に型を付けないこと（[string] にすると $null が空文字になり、空の属性を書いてしまう）
    param([string] $Tag, [string] $Name, $Value)

    $pattern = '\s' + [regex]::Escape($Name) + '="[^"]*"'
    $Tag = [regex]::Replace($Tag, $pattern, '')
    if ($null -eq $Value) { return $Tag }
    $insert = ' {0}="{1}"' -f $Name, $Value
    if ($Tag.EndsWith('/>')) { return $Tag.Substring(0, $Tag.Length - 2) + $insert + '/>' }
    return $Tag.Substring(0, $Tag.Length - 1) + $insert + '>'
}

function Get-XmlAttribute {
    param([string] $Tag, [string] $Name)
    $m = [regex]::Match($Tag, '\s' + [regex]::Escape($Name) + '="([^"]*)"')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

function Update-SheetViews {
    #     シート XML の <sheetViews> だけを書き換えた文字列を返す（ほかの部分は 1 文字も変えない）。
    #       ・選択セルを A1（見出しを固定しているシートはスクロール領域の左上）に
    #       ・表示位置を左上に、倍率を $Zoom に
    #       ・タブの選択は Selected が $true のシートだけに（作業グループの解除）
    #       ・-NormalView なら改ページプレビューなどを標準表示に
    param([string] $Xml, [bool] $Selected)

    $m = [regex]::Match($Xml, '<(\w+:)?sheetViews\b[^>]*>.*?</(\w+:)?sheetViews>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $m.Success) { return $Xml }
    $block = $m.Value

    $block = [regex]::Replace($block, '<(\w+:)?sheetView\b[^>]*?/?>', {
        param($t)
        $tag = $t.Value
        $tag = Set-XmlAttribute $tag 'topLeftCell' $null
        $tag = Set-XmlAttribute $tag 'tabSelected' $null
        if ($Selected) { $tag = Set-XmlAttribute $tag 'tabSelected' '1' }
        $tag = Set-XmlAttribute $tag 'zoomScaleNormal' $null
        if ($Zoom -eq 100) { $tag = Set-XmlAttribute $tag 'zoomScale' $null }
        else { $tag = Set-XmlAttribute $tag 'zoomScale' ([string]$Zoom) }
        if ($NormalView) { $tag = Set-XmlAttribute $tag 'view' $null }
        return $tag
    })

    # ウィンドウ枠の固定: スクロール領域を固定位置の直後に戻し、そこを選択する
    $homeCell = 'A1'
    $activePane = $null
    $paneMatch = [regex]::Match($block, '<(\w+:)?pane\b[^>]*/>')
    if ($paneMatch.Success) {
        $pane = $paneMatch.Value
        $state = Get-XmlAttribute $pane 'state'
        if ($state -eq 'frozen' -or $state -eq 'frozenSplit') {
            $x = [int][double]('0' + (Get-XmlAttribute $pane 'xSplit'))
            $y = [int][double]('0' + (Get-XmlAttribute $pane 'ySplit'))
            $homeCell = (ConvertTo-ColumnName ($x + 1)) + ($y + 1)
            $activePane = Get-XmlAttribute $pane 'activePane'
            $block = $block.Replace($pane, (Set-XmlAttribute $pane 'topLeftCell' $homeCell))
        }
    }

    $block = [regex]::Replace($block, '<(\w+:)?selection\b[^>]*/>', {
        param($t)
        $tag = $t.Value
        $pane = Get-XmlAttribute $tag 'pane'
        # 固定したシートで、カーソルの無い枠の選択はそのままにする
        if ($activePane -and $pane -and $pane -ne $activePane) { return $tag }
        $tag = Set-XmlAttribute $tag 'activeCellId' $null
        $tag = Set-XmlAttribute $tag 'activeCell' $homeCell
        $tag = Set-XmlAttribute $tag 'sqref' $homeCell
        return $tag
    })

    return $Xml.Substring(0, $m.Index) + $block + $Xml.Substring($m.Index + $m.Length)
}

function Update-WorkbookView {
    #     workbook.xml の <workbookView> で、開いたときに表示するシートを FirstVisible 番目にする
    param([string] $Xml, [int] $FirstVisible)
    return [regex]::Replace($Xml, '<(\w+:)?workbookView\b[^>]*?/?>', {
        param($t)
        $tag = Set-XmlAttribute $t.Value 'firstSheet' $null
        if ($FirstVisible -eq 0) { return (Set-XmlAttribute $tag 'activeTab' $null) }
        return (Set-XmlAttribute $tag 'activeTab' ([string]$FirstVisible))
    })
}

# ---------------------------------------------------------------- 1 ファイルの修正


function Read-EntryText {
    param($Entry)
    $sr = New-Object System.IO.StreamReader($Entry.Open(), (New-Object System.Text.UTF8Encoding($false)), $true)
    try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
}

function Repair-Workbook {
    #     表示の状態を直した内容を一時ファイルに書き、元のファイルと置き換える。
    #     書き換えるのは workbook.xml と各シートの XML の表示に関する部分だけで、ほかのパーツはそのままコピーする
    param([string] $FilePath)

    # 書き換えるパーツと、その新しい内容を先に決める
    $pkg = Open-ZipPackage -FilePath $FilePath
    $changes = @{}
    try {
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) { if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break } }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        $relById = @{}
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName $wbPart)) { $relById[$rel.Id] = $rel }

        $sheetParts = New-Object System.Collections.Generic.List[object]
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            if ($null -eq $rel) { continue }
            $state = $s.GetAttribute('state')
            $sheetParts.Add([pscustomobject]@{ Part = $rel.Target; Visible = (-not $state -or $state -eq 'visible') })
        }
        $firstVisible = 0
        for ($i = 0; $i -lt $sheetParts.Count; $i++) { if ($sheetParts[$i].Visible) { $firstVisible = $i; break } }

        $text = Read-EntryText (Get-PackageEntry -Package $pkg -PartName $wbPart)
        $new = Update-WorkbookView -Xml $text -FirstVisible $firstVisible
        if ($new -ne $text) { $changes[$wbPart.ToLowerInvariant()] = $new }

        for ($i = 0; $i -lt $sheetParts.Count; $i++) {
            $entry = Get-PackageEntry -Package $pkg -PartName $sheetParts[$i].Part
            if ($null -eq $entry) { continue }
            $text = Read-EntryText $entry
            $new = Update-SheetViews -Xml $text -Selected ($i -eq $firstVisible)
            if ($new -ne $text) { $changes[$sheetParts[$i].Part.ToLowerInvariant()] = $new }
        }
    }
    finally { $pkg.Zip.Dispose() }

    if ($changes.Count -eq 0) { return 0 }

    # 書き換えたパーツだけを差し替えて、元のファイルに上書きする（代替データストリームやアクセス権は保つ）
    Save-ZipPackage -FilePath $FilePath -Changes $changes
    return $changes.Count
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

# ---------------------------------------------------------------- 実行


$files = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
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
            if (@('.xlsx', '.xlsm', '.xltx', '.xltm') -notcontains $f.Extension.ToLowerInvariant()) { continue }
            if ($f.Name.StartsWith('~$')) { continue }
            if ($BackupDir -and $f.FullName.StartsWith($BackupDir, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $files.Add([pscustomobject]@{ Info = $f; Root = $root; Relative = $f.FullName.Substring($root.Length).TrimStart('\', '/') })
        }
    }
}
if ($files.Count -eq 0) { Write-Warning '対象の Excel ファイル（.xlsx / .xlsm）がありません。'; return }

$results = New-Object System.Collections.Generic.List[object]
$index = 0
foreach ($entry in $files) {
    $index++
    $file = $entry.Info
    Write-Progress -Activity '表示の状態を直しています' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) -PercentComplete ([int](100 * $index / $files.Count))

    $result = '変更なし'
    $detail = ''
    try {
        # 直す必要があるかは、提出前チェックの「体裁」と同じ基準で判定する
        $found = @(& $checker -Path ([System.Management.Automation.WildcardPattern]::Escape($file.FullName)) -Check View -Zoom $Zoom 6>$null)
        $issues = @($found | Where-Object { $_.Category -eq '体裁' -and ($NormalView -or $_.Item -ne '表示モード') })
        $unreadable = @($found | Where-Object { $_.Category -eq 'ファイル' })
        if ($file.IsReadOnly) {
            # 確定版などで意図的に読み取り専用にしていることが多いので、書き換えない
            $result = 'スキップ'
            $detail = '読み取り専用のファイルです（読み取り専用を外すと直せます）'
        }
        elseif ($unreadable.Count -gt 0) {
            $result = 'スキップ'
            $detail = $unreadable[0].Detail
        }
        elseif ($issues.Count -eq 0 -and -not $Force) {
            $result = '変更なし'
        }
        else {
            $detail = (@($issues | ForEach-Object { $_.Item } | Sort-Object -Unique) -join '・')
            if ($PSCmdlet.ShouldProcess($entry.Relative, "表示の状態を直す（$detail）")) {
                if (Test-FileLocked $file.FullName) { throw 'ファイルを開いているため書き換えられません（Excel で閉じてから実行してください）' }
                $lastWrite = $file.LastWriteTimeUtc
                if ($BackupDir) {
                    $dest = Join-Path $BackupDir $entry.Relative
                    New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
                    Copy-Item -LiteralPath $file.FullName -Destination $dest -Force
                }
                $null = Repair-Workbook -FilePath $file.FullName
                if ($KeepTimestamp) { [System.IO.File]::SetLastWriteTimeUtc($file.FullName, $lastWrite) }
                $result = '修正'
            }
            else {
                $result = '修正予定'
            }
        }
    }
    catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $result = '失敗'
        $detail = $ex.Message
        Write-Warning ("{0}: {1}" -f $entry.Relative, $detail)
    }
    $results.Add([pscustomobject]@{ File = $entry.Relative; Result = $result; Detail = $detail; Path = $file.FullName })
}
Write-Progress -Activity '表示の状態を直しています' -Completed

$results
Write-Host ''
$counts = $results | Group-InOrder Result | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
Write-Host ("{0} ファイル: {1}" -f $files.Count, ($counts -join ' / '))
if ($BackupDir -and @($results | Where-Object Result -eq '修正').Count -gt 0) { Write-Host "修正前のファイル: $BackupDir" }

# .SYNOPSIS
#     Excel ファイルを、開いたときに先頭シートの A1 が倍率 100% で表示される状態に直します（Excel 不要）。
#
# .DESCRIPTION
#     提出前の「体裁をそろえる」作業を一括で行います。Book1.xlsm の「体裁整える」マクロの代わりです。
#
#       ・開いたときに、表示されている最初のシートが表示されるようにする
#       ・各シートの選択セルを A1 に、表示位置を左上に、倍率を 100% にする
#         （ウィンドウ枠を固定したシートは、スクロール領域の左上を選択する）
#       ・複数シートの選択（作業グループ）を解除する
#       ・-NormalView を付けると、改ページプレビューなどを標準表示に戻す
#
#     直す必要があるかを Test-ExcelDeliverable.ps1 と同じ基準で判定し、必要なファイルだけを書き換えます。
#     書き換えるのはファイル内部の表示に関する設定だけで、セルの値・書式・マクロなどには触れません。
#     Excel を使わないため、マクロ付きのブックでもマクロは実行されません。
#
# .PARAMETER Path
#     対象のフォルダまたはファイル（複数指定可、ワイルドカード可）。xlsx / xlsm / xltx / xltm が対象です。
#
# .PARAMETER Zoom
#     そろえる倍率（%）。既定は 100。
#
# .PARAMETER NormalView
#     改ページプレビュー・ページレイアウト表示のシートを標準表示に戻します。
#
# .PARAMETER KeepTimestamp
#     ファイルの更新日時を書き換え前のままにします（表示を直しただけで「更新された」ように見えないように）。
#
# .PARAMETER BackupDir
#     書き換える前のファイルを、このフォルダにフォルダ構成ごとコピーしておきます。
#
# .PARAMETER Force
#     直す必要がないと判定したファイルも書き換えます。
#
# .PARAMETER NoRecurse
#     サブフォルダを対象にしません。
#
# .EXAMPLE
#     # 何が直されるかだけ確認する
#     .\Reset-ExcelView.ps1 D:\提出物 -WhatIf
#
# .EXAMPLE
#     # 修正前のファイルを残し、更新日時は変えずに直す
#     .\Reset-ExcelView.ps1 D:\提出物 -BackupDir D:\修正前 -KeepTimestamp
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1・Test-ExcelDeliverable.ps1・Search-Excel.ps1 が必要です。
#     Excel で開いているファイルは書き換えられません（失敗として表示されます）。読み取り専用のファイルは書き換えません。
#     書き換えは元のファイルの中身を置き換える形で行うため、アクセス権・作成日時・インターネットから取得したファイルの印
#     （保護ビューで開くための情報）はそのまま残ります。
