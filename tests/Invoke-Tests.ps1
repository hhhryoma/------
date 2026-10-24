[CmdletBinding()]
param(
    # 実行するテスト名を絞り込む正規表現（例: -Filter 'Search-Excel'）
    [string] $Filter
)

#     tools 以下のスクリプトのテストを実行する。外部モジュール（Pester など）は使わない。
#     テストは tests\*.Tests.ps1 に Test-Case で書く。失敗が 1 件でもあれば終了コード 1。

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression

$script:RepoRoot  = Split-Path $PSScriptRoot -Parent
$script:ToolsDir  = Join-Path $script:RepoRoot 'tools'
$script:Fixtures  = Join-Path $PSScriptRoot 'fixtures'
$script:OnWindows = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows
$script:Results   = New-Object System.Collections.Generic.List[object]
$script:TempDirs  = New-Object System.Collections.Generic.List[string]

# ---------------------------------------------------------------- テストの書き方

function Test-Case {
    param(
        [Parameter(Mandatory = $true, Position = 0)] [string] $Name,
        [Parameter(Mandatory = $true, Position = 1)] [scriptblock] $Body,
        [switch] $WindowsOnly,
        [switch] $NonWindowsOnly
    )
    if ($Filter -and $Name -notmatch $Filter) { return }

    if (($WindowsOnly -and -not $script:OnWindows) -or ($NonWindowsOnly -and $script:OnWindows)) {
        $script:Results.Add([pscustomobject]@{ Name = $Name; Status = 'SKIP'; Message = 'この OS では実行しない'; Seconds = 0 })
        Write-Host ("  [SKIP] {0}" -f $Name) -ForegroundColor DarkGray
        return
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $null = & $Body
        $sw.Stop()
        $script:Results.Add([pscustomobject]@{ Name = $Name; Status = 'PASS'; Message = ''; Seconds = $sw.Elapsed.TotalSeconds })
        Write-Host ("  [PASS] {0} ({1:N1}s)" -f $Name, $sw.Elapsed.TotalSeconds) -ForegroundColor Green
    }
    catch {
        $sw.Stop()
        $where = ''
        if ($_.InvocationInfo -and $_.InvocationInfo.ScriptName) {
            $where = ' [{0}:{1}]' -f (Split-Path $_.InvocationInfo.ScriptName -Leaf), $_.InvocationInfo.ScriptLineNumber
        }
        $msg = $_.Exception.Message + $where
        $script:Results.Add([pscustomobject]@{ Name = $Name; Status = 'FAIL'; Message = $msg; Seconds = $sw.Elapsed.TotalSeconds })
        Write-Host ("  [FAIL] {0}" -f $Name) -ForegroundColor Red
        Write-Host ("         {0}" -f $msg) -ForegroundColor Red
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string] $Message = '')
    $e = @($Expected) -join ' | '
    $a = @($Actual) -join ' | '
    if ($e -cne $a) { throw ("{0} 期待値: [{1}] 実際: [{2}]" -f $Message, $e, $a) }
}

function Assert-True {
    param($Condition, [string] $Message = '条件が成り立ちません')
    if (-not $Condition) { throw $Message }
}

function Assert-Match {
    param([string] $Text, [string] $Pattern, [string] $Message = '')
    if ($Text -notmatch $Pattern) { throw ("{0} /{1}/ に一致しません: [{2}]" -f $Message, $Pattern, $Text) }
}

function Assert-NotMatch {
    param([string] $Text, [string] $Pattern, [string] $Message = '')
    if ($Text -match $Pattern) { throw ("{0} /{1}/ に一致してはいけません: [{2}]" -f $Message, $Pattern, $Text) }
}

function Assert-Throws {
    param([scriptblock] $Body, [string] $Pattern = '.')
    try { $null = & $Body } catch {
        if ($_.Exception.Message -notmatch $Pattern) { throw ("例外のメッセージが /{0}/ に一致しません: {1}" -f $Pattern, $_.Exception.Message) }
        return
    }
    throw '例外が発生しませんでした'
}

# ---------------------------------------------------------------- 共通の補助


function New-TestDir {
    $d = Join-Path ([System.IO.Path]::GetTempPath()) ('xltools-test-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $d | Out-Null
    $script:TempDirs.Add($d)
    return $d
}

function Copy-Fixture {
    #     フィクスチャ（tests\fixtures からの相対パス、または絶対パス）を一時フォルダへコピーして、
    #     コピー先のパスを返す（テストがフィクスチャを汚さないように）。
    #     コピー先は必ず一時フォルダの直下（元の名前の最後の部分だけを使う）
    param([string] $Source)
    if (-not [System.IO.Path]::IsPathRooted($Source)) { $Source = Join-Path $script:Fixtures $Source }
    $Source = (Resolve-Path -LiteralPath $Source).ProviderPath
    $dest = Join-Path (New-TestDir) (Split-Path $Source -Leaf)
    Copy-Item -LiteralPath $Source -Destination $dest -Recurse
    return $dest
}

function ConvertTo-Slash {
    # 相対パスの区切りを OS によらず / にそろえる
    param([string] $Path)
    return $Path.Replace('\', '/')
}

function Read-XlsxSheet {
    #     xlsx の 1 シートを読む（テスト用の簡易版）。
    #     返り値: Rows（行番号 → 列 → 文字列）、Links（ref / location / target）、Sheets（シート名の一覧）
    param([string] $Path, [int] $Index = 1)

    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $read = {
            param($n)
            $e = $zip.GetEntry($n)
            if ($null -eq $e) { return $null }
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { $x = New-Object System.Xml.XmlDocument; $x.LoadXml($sr.ReadToEnd()); return $x } finally { $sr.Dispose() }
        }

        $sst = @()
        $sstDoc = & $read 'xl/sharedStrings.xml'
        if ($sstDoc) { $sst = @($sstDoc.GetElementsByTagName('si') | ForEach-Object { $_.InnerText }) }

        $wb = & $read 'xl/workbook.xml'
        $sheets = @($wb.GetElementsByTagName('sheet') | ForEach-Object { $_.GetAttribute('name') })

        $doc = & $read "xl/worksheets/sheet$Index.xml"
        $rows = @{}
        foreach ($c in $doc.GetElementsByTagName('c')) {
            $m = [regex]::Match($c.GetAttribute('r'), '^([A-Z]+)(\d+)$')
            $row = [int]$m.Groups[2].Value
            if (-not $rows.ContainsKey($row)) { $rows[$row] = @{} }
            $t = $c.GetAttribute('t')
            if ($t -eq 'inlineStr') { $v = $c.InnerText }
            elseif ($t -eq 's') { $v = $sst[[int]$c.InnerText] }
            else { $v = $c.InnerText }
            $rows[$row][$m.Groups[1].Value] = $v
        }

        $targets = @{}
        $relDoc = & $read "xl/worksheets/_rels/sheet$Index.xml.rels"
        if ($relDoc) { foreach ($r in $relDoc.GetElementsByTagName('Relationship')) { $targets[$r.GetAttribute('Id')] = $r.GetAttribute('Target') } }
        $links = @()
        foreach ($h in $doc.GetElementsByTagName('hyperlink')) {
            $rid = $h.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $links += [pscustomobject]@{ Ref = $h.GetAttribute('ref'); Location = $h.GetAttribute('location'); Target = $targets[$rid] }
        }

        return [pscustomobject]@{ Rows = $rows; Links = $links; Sheets = $sheets }
    }
    finally { $fs.Dispose() }
}

function Read-XlsxCharts {
    #     xlsx のグラフ（xl/charts/*.xml）を読む（テスト用）。返り値: @{ Title; Refs = 参照する範囲の一覧; Xml }
    param([string] $Path)
    $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $charts = @()
        foreach ($e in @($zip.Entries | Where-Object { $_.FullName -like 'xl/charts/*.xml' } | Sort-Object { [int]($_.Name -replace '\D', '') })) {
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { $xml = $sr.ReadToEnd() } finally { $sr.Dispose() }
            [xml]$doc = $xml   # 正しい XML か
            $title = ([regex]::Match($xml, '<c:title><c:tx><c:rich>.*?<a:t>([^<]*)</a:t>')).Groups[1].Value
            $charts += [pscustomobject]@{ Title = $title; Refs = @([regex]::Matches($xml, '<c:f>([^<]*)</c:f>') | ForEach-Object { $_.Groups[1].Value }); Xml = $xml }
        }
        return $charts
    }
    finally { $fs.Dispose() }
}

function Set-CellValues {
    #     ブックのシートのセルの値を書き換える（テスト用。$null で空にする）
    param([string] $File, [string] $Sheet, [hashtable] $Values)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $pkg = Open-ZipPackage $File
    try {
        $part = (Get-WorkbookSheetParts $pkg).Sheets[$Sheet]
        $doc = Read-PartXmlForEdit $pkg $part
        foreach ($ref in $Values.Keys) {
            $row = [int]($ref -replace '^[A-Z]+', '')
            $rowEl = Get-SheetRowElement $doc $row -Create
            Set-SheetCellValue $doc (Get-SheetCellElement $doc $rowEl $row (Get-ColumnIndex $ref) -Create) $Values[$ref]
        }
        $changes = @{ $part.ToLowerInvariant() = (ConvertTo-XmlString $doc) }
    }
    finally { $pkg.Zip.Dispose() }
    Save-ZipPackage -FilePath $File -Changes $changes
}

function New-WordDoc {
    #     段落と表だけの最小の Word 文書を作る。Body: 文字は段落、配列の配列は表（行ごとのセル）
    param([string] $Path, [object[]] $Body)
    $esc = { param($t) [System.Security.SecurityElement]::Escape($t) }
    $p = { param($t) '<w:p><w:r><w:t xml:space="preserve">' + (& $esc $t) + '</w:t></w:r></w:p>' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in $Body) {
        if ($b -is [string]) { [void]$sb.Append((& $p $b)); continue }
        [void]$sb.Append('<w:tbl>')
        foreach ($row in $b) { [void]$sb.Append('<w:tr>' + ((@($row | ForEach-Object { '<w:tc>' + (& $p $_) + '</w:tc>' })) -join '') + '</w:tr>') }
        [void]$sb.Append('</w:tbl>')
    }
    $xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $sb.ToString() + '</w:body></w:document>'
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $fs = [System.IO.File]::Create($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        $sw = New-Object System.IO.StreamWriter($zip.CreateEntry('word/document.xml').Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($xml); $sw.Dispose()
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

function New-SheetsBook {
    #     複数シートのブックを作る（テスト用）。Sheets: @{ 名前 = @(見出しの配列, 行の配列...) } の順序付き辞書
    param([string] $Path, $Sheets)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $list = @()
    foreach ($name in $Sheets.Keys) {
        $header = $Sheets[$name][0]
        $rows = @($Sheets[$name] | Select-Object -Skip 1)
        $list += @{ Name = $name; Columns = @($header | ForEach-Object { @{ Header = $_; Width = 10 } }); Rows = @($rows | ForEach-Object { @{ Values = $_ } }) }
    }
    Export-Workbook -FilePath $Path -Sheets $list
}

# テスト用の小さな画像（40×30 の PNG）。Retake は Old と絵は同じで、データ（圧縮の仕方）だけが違う（同じ画面を撮り直したとき）
$script:TestPng = @{
    Old    = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAYElEQVR4nO2XMQpAIQxDE+n9r1wH4QviomALv8kidXk0pEXp7shQS6ECsHGQDEMOj22pX+vrMM1qgcNkm7ujhN9Gsp7VAv8fvBunkKVdz2qBBRZY4GvNXR35pgfAcn+nDgTjEEBUFCIpAAAAAElFTkSuQmCC'
    New    = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAbUlEQVR4nO2XwQoAIQhEZxb//5ftEHQIYSlaXdK5RF0eYzoUVRURekKoAKQvJN2QvcYy7b/WcBhW6gK7SYyzpQ7fbcl8pS7w/WBrnFxC2wJv6XX4Jzv57rjABb4PfCxAVuPuB4493/QAmO7v1ACh5BRCnP+rmgAAAABJRU5ErkJggg=='
    Retake = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAgUlEQVR4Ae2WSw6AMAhEHeP9r4yams5CE7GhsHC6YkEY+vgEmNlS8dYK0VNza8IA0jJojC/hQzWHef9hGWoJp7UYm4uSnzp8dA2oxgQ+2RLqyYAZ/mmcRieEUR2WauyAFOMi1DEcHVGE2gEpxuV/qLmr+8Ubg/ItCnLu+HsaZTXeAQTjEEApC71sAAAAAElFTkSuQmCC'
    Logo   = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAlklEQVR4nO2WSw6AIAxEqfHexpPXpZt+pgWhMXRpOvMYLARi5raijiXUDZ5ZJ95KN4nf+cqMJ7lTrfE6V+BsNU6NNlvgkFFUooIT1JBQBqepuFwAd1JBkzIXyJC4iFWZxBs8DZy78cWyrcokboNCuyZy4k42Ile3Os0GhdY/TrBxiTNcIXao2X/6vK2T31wfVaVz/HPwA6K+OS+gTz8VAAAAAElFTkSuQmCC'
    # 400×300 の灰色の画像（スクリーンショットの代わり）
    Screen = 'iVBORw0KGgoAAAANSUhEUgAAAZAAAAEsCAAAAADI3LoeAAABsUlEQVR42u3RAQ0AAAjDsEvDvxok4IJA0klYUzpVLAAiIEAEBIiAABEQIAIiIEAEBIiAABEQIAIiIEAE5D1I61RAgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAgIEAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAAREAEBIiBABASIgAARECAWABEQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAREQIAICRECACAgQAQEiIAICRECACAgQrTYc8MQ056+ClAAAAABJRU5ErkJggg=='
}

function Add-TestDrawing {
    #     ブックの Sheet 番目のシートに図形・画像を置く（テスト用）。Items の要素（列・行は 0 から数える）:
    #       @{ Kind = 'pic'; Name; Png（$script:TestPng のキー）; From = @(列, 行); To = @(列, 行) }
    #       @{ Kind = 'sp'; Name; Text（省略可）; From; To }
    #       @{ Kind = 'grpSp'; Name; From; To; Children = @( @{ Kind; Name; Text; Png; Off = @(x, y); Ext = @(cx, cy) } ) }（中の座標は 0〜1000）
    #       Hidden = $true で非表示にする。シートに置いたものは Ext = @(幅, 高さ)（ピクセル）で表示の大きさ（xfrm）を書く
    #       画像は Crop = @(左, 上, 右, 下)（1/1000 %）でトリミングする
    param([string] $Path, $Items, [int] $Sheet = 1)
    $media = New-Object System.Collections.Generic.List[string]
    $script:TestShapeId = 1
    $obj = {
        param($it, [string] $Xfrm)
        $script:TestShapeId++
        $id = $script:TestShapeId
        if ($it.Hidden) { $id = '{0}" hidden="1' -f $id }
        switch ($it.Kind) {
            'pic' {
                $media.Add($it.Png)
                $crop = ''
                if ($it.Crop) { $crop = '<a:srcRect l="{0}" t="{1}" r="{2}" b="{3}"/>' -f $it.Crop[0], $it.Crop[1], $it.Crop[2], $it.Crop[3] }
                '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{0}" name="{1}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rIdImg{2}"/>{4}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>{3}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>' -f $id, $it.Name, $media.Count, $Xfrm, $crop
            }
            'sp' {
                $body = ''
                if ($it.Text) { $body = '<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>{0}</a:t></a:r></a:p></xdr:txBody>' -f $it.Text }
                '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="{0}" name="{1}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>{2}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>{3}</xdr:sp>' -f $id, $it.Name, $Xfrm, $body
            }
            'grpSp' {
                $kids = (@($it.Children | ForEach-Object { & $obj $_ ('<a:xfrm><a:off x="{0}" y="{1}"/><a:ext cx="{2}" cy="{3}"/></a:xfrm>' -f $_.Off[0], $_.Off[1], $_.Ext[0], $_.Ext[1]) })) -join ''
                '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="{0}" name="{1}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/><a:chOff x="0" y="0"/><a:chExt cx="1000" cy="1000"/></a:xfrm></xdr:grpSpPr>{2}</xdr:grpSp>' -f $id, $it.Name, $kids
            }
        }
    }
    $anchors = foreach ($it in $Items) {
        $ext = @(0, 0)
        if ($it.Ext) { $ext = @(($it.Ext[0] * 9525), ($it.Ext[1] * 9525)) }
        $body = & $obj $it ('<a:xfrm><a:off x="0" y="0"/><a:ext cx="{0}" cy="{1}"/></a:xfrm>' -f $ext[0], $ext[1])
        '<xdr:twoCellAnchor><xdr:from><xdr:col>{0}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{3}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>{4}<xdr:clientData/></xdr:twoCellAnchor>' -f $it.From[0], $it.From[1], $it.To[0], $it.To[1], $body
    }
    $head = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    $drawing = $head + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + ($anchors -join '') + '</xdr:wsDr>'
    $relsHead = $head + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    $drawingRels = $relsHead + ((@(for ($i = 1; $i -le $media.Count; $i++) { '<Relationship Id="rIdImg{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{1}_{0}.png"/>' -f $i, $Sheet })) -join '') + '</Relationships>'
    $sheetRels = $relsHead + ('<Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{0}.xml"/>' -f $Sheet) + '</Relationships>'
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        $put = { param([string] $Name, [byte[]] $Bytes) $e = $zip.GetEntry($Name); if ($e) { $e.Delete() }; $st = $zip.CreateEntry($Name).Open(); $st.Write($Bytes, 0, $Bytes.Length); $st.Dispose() }
        $read = { param([string] $Name) $sr = New-Object System.IO.StreamReader($zip.GetEntry($Name).Open()); try { $sr.ReadToEnd() } finally { $sr.Dispose() } }
        & $put ('xl/drawings/drawing{0}.xml' -f $Sheet) $utf8.GetBytes($drawing)
        & $put ('xl/drawings/_rels/drawing{0}.xml.rels' -f $Sheet) $utf8.GetBytes($drawingRels)
        for ($i = 0; $i -lt $media.Count; $i++) { & $put ('xl/media/image{0}_{1}.png' -f $Sheet, ($i + 1)) ([Convert]::FromBase64String($script:TestPng[$media[$i]])) }
        & $put ('xl/worksheets/_rels/sheet{0}.xml.rels' -f $Sheet) $utf8.GetBytes($sheetRels)
        $sheetXml = (& $read ('xl/worksheets/sheet{0}.xml' -f $Sheet)).Replace('</worksheet>', '<drawing r:id="rIdDrw1"/></worksheet>')
        & $put ('xl/worksheets/sheet{0}.xml' -f $Sheet) $utf8.GetBytes($sheetXml)
        $types = & $read '[Content_Types].xml'
        if (-not $types.Contains('Extension="png"')) { $types = $types.Replace('<Default Extension="xml"', '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"') }
        $types = $types.Replace('</Types>', ('<Override PartName="/xl/drawings/drawing{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>' -f $Sheet))
        & $put '[Content_Types].xml' $utf8.GetBytes($types)
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

function Edit-ZipText {
    #     xlsx の中の 1 つのパーツの文字を書き換える（テスト用）
    param([string] $Path, [string] $Name, [scriptblock] $Change)
    $fs = [System.IO.File]::Open($Path, 'Open', 'ReadWrite')
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Update)
        $entry = $zip.GetEntry($Name)
        $sr = New-Object System.IO.StreamReader($entry.Open())
        $text = $sr.ReadToEnd(); $sr.Dispose()
        $text = & $Change $text
        $entry.Delete()
        $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($Name).Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($text); $sw.Dispose()
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

function Get-HostText {
    #     Write-Host の出力（情報ストリーム）も含めて文字列で受け取る
    param([scriptblock] $Body)
    return ((& $Body 6>&1 3>&1 | ForEach-Object { "$_" }) -join "`n")
}

# ---------------------------------------------------------------- 実行


Write-Host ("PowerShell {0} ({1})" -f $PSVersionTable.PSVersion, $PSVersionTable.PSEdition)
try {
    foreach ($file in (Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' | Sort-Object Name)) {
        Write-Host ''
        Write-Host ("== {0}" -f $file.Name)
        . $file.FullName
    }
}
finally {
    foreach ($d in $script:TempDirs) { Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue }
}

$pass = @($script:Results | Where-Object Status -eq 'PASS').Count
$fail = @($script:Results | Where-Object Status -eq 'FAIL').Count
$skip = @($script:Results | Where-Object Status -eq 'SKIP').Count
Write-Host ''
Write-Host ("成功 {0} / 失敗 {1} / スキップ {2}" -f $pass, $fail, $skip)
if ($fail -gt 0) {
    Write-Host '失敗したテスト:' -ForegroundColor Red
    foreach ($r in ($script:Results | Where-Object Status -eq 'FAIL')) { Write-Host ("  {0}: {1}" -f $r.Name, $r.Message) -ForegroundColor Red }
    exit 1
}
if ($pass -eq 0) { Write-Host '実行されたテストがありません。' -ForegroundColor Yellow; exit 1 }
exit 0
