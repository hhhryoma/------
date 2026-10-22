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
