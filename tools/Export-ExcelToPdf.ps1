[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [string] $OutputRoot,

    [switch] $NoRecurse,
    [switch] $PerSheet,
    [switch] $SkipExisting,
    [switch] $FitToWidth,
    [switch] $IncludeHiddenSheets,

    [string] $SheetNamePattern,
    [string] $SheetNameExclude,

    [switch] $CountOnly,

    [switch] $IgnoreSmallPages,

    [ValidateRange(0.01, 1.0)]
    [double] $SmallPageRatio = 0.2,

    [ValidateSet('Auto', 'Full', 'Simple')]
    [string] $OpenMode = 'Auto',

    [string] $LogPath,

    [string[]] $Extensions = @('.xls', '.xlsx', '.xlsm', '.xlsb')
)

$ErrorActionPreference = 'Stop'

# ---- Excel 定数（COM の列挙体を PowerShell から参照しないで済むよう定義） ----
$xlTypePDF         = 0
$xlQualityStandard = 0
$xlSheetVisible    = -1
$xlSheetHidden     = 0
$msoAutomationSecurityForceDisable = 3

# ---------------------------------------------------------------- ヘルパー

function ConvertTo-SafeFileName {
    param([string] $Name)

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    $result = $sb.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($result)) { return 'sheet' }
    return $result
}

function Get-UniquePath {
    param([string] $CandidatePath)

    if (-not (Test-Path -LiteralPath $CandidatePath)) { return $CandidatePath }

    $dir  = [System.IO.Path]::GetDirectoryName($CandidatePath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($CandidatePath)
    $ext  = [System.IO.Path]::GetExtension($CandidatePath)

    for ($n = 2; $n -lt 1000; $n++) {
        $try = Join-Path $dir ("{0}({1}){2}" -f $base, $n, $ext)
        if (-not (Test-Path -LiteralPath $try)) { return $try }
    }
    throw "出力先の連番が上限に達しました: $CandidatePath"
}

function Test-SheetNameMatch {
    #     シート名が -SheetNamePattern / -SheetNameExclude の条件を満たすか判定する。
    #     いずれも .NET 正規表現、大文字小文字は区別しない。
    param([string] $Name)

    if ($SheetNamePattern -and -not [regex]::IsMatch($Name, $SheetNamePattern, 'IgnoreCase')) { return $false }
    if ($SheetNameExclude -and [regex]::IsMatch($Name, $SheetNameExclude, 'IgnoreCase')) { return $false }
    return $true
}

function Test-HasPrintableContent {
    param($Sheet)

    # グラフシートなど UsedRange を持たないシートは「印刷対象あり」と判定する
    try {
        if ($Sheet.Shapes.Count -gt 0) { return $true }
        $used = $Sheet.UsedRange
        if ($null -eq $used) { return $false }
        if ($used.Count -gt 1) { return $true }
        return -not [string]::IsNullOrWhiteSpace([string]$used.Value2)
    }
    catch {
        return $true
    }
}

function Get-PdfObjectWindow {
    # PDF 内の指定位置を含むオブジェクト（N 0 obj ... endobj）の範囲を切り出す
    param([string] $Text, [int] $Index)

    $start = $Text.LastIndexOf(' obj', $Index)
    if ($start -lt 0) { $start = [math]::Max(0, $Index - 4000) }
    $end = $Text.IndexOf('endobj', $Index)
    if ($end -lt 0) { $end = [math]::Min($Text.Length, $Index + 4000) }
    if ($end -le $start) { return '' }
    return $Text.Substring($start, $end - $start)
}

function Get-PdfMediaBoxSize {
    # /MediaBox [x1 y1 x2 y2] から用紙サイズ（mm）を求める。PDF の単位は 1/72 インチ
    param([string] $Text)

    $m = [regex]::Match($Text, '/MediaBox\s*\[\s*(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s+(-?[\d.]+)\s*\]')
    if (-not $m.Success) { return $null }

    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    try {
        $x1 = [double]::Parse($m.Groups[1].Value, $ci)
        $y1 = [double]::Parse($m.Groups[2].Value, $ci)
        $x2 = [double]::Parse($m.Groups[3].Value, $ci)
        $y2 = [double]::Parse($m.Groups[4].Value, $ci)
    }
    catch { return $null }

    return [pscustomobject]@{
        WidthMm  = [math]::Abs($x2 - $x1) / 72.0 * 25.4
        HeightMm = [math]::Abs($y2 - $y1) / 72.0 * 25.4
    }
}

function Get-PdfPageSizes {
    #     PDF を解析し、ページごとの用紙サイズ（mm）と面積を返す。
    #     ページ個別に /MediaBox が無い場合はページツリー側の値を継承する。
    #     解析できなければ $null。
    param([string] $PdfPath)

    try {
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        $text  = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)

        $pageMatches = [regex]::Matches($text, '/Type\s*/Page(?![sA-Za-z])')
        if ($pageMatches.Count -eq 0) { return $null }

        # 継承元（/Type /Pages 側）の既定サイズ
        $defaultSize = $null
        $pagesNode = [regex]::Match($text, '/Type\s*/Pages\b')
        if ($pagesNode.Success) {
            $defaultSize = Get-PdfMediaBoxSize -Text (Get-PdfObjectWindow -Text $text -Index $pagesNode.Index)
        }

        $pages = New-Object System.Collections.Generic.List[object]
        $no = 0
        foreach ($m in $pageMatches) {
            $no++
            $size = Get-PdfMediaBoxSize -Text (Get-PdfObjectWindow -Text $text -Index $m.Index)
            if ($null -eq $size) { $size = $defaultSize }
            if ($null -eq $size) { $size = [pscustomobject]@{ WidthMm = 0.0; HeightMm = 0.0 } }

            $pages.Add([pscustomobject]@{
                Page   = $no
                Width  = [math]::Round($size.WidthMm, 1)
                Height = [math]::Round($size.HeightMm, 1)
                Area   = $size.WidthMm * $size.HeightMm
            })
        }
        return $pages
    }
    catch {
        Write-Verbose ("ページサイズを解析できませんでした: {0} : {1}" -f $PdfPath, $_.Exception.Message)
        return $null
    }
}

function Get-PdfPageCountFallback {
    #     ページ個別の解析ができない場合に、総ページ数だけを求める。
    #     オブジェクトストリームで圧縮された PDF 向けに、ページツリーの /Count を読む。
    param([string] $PdfPath)

    try {
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        $text  = [System.Text.Encoding]::GetEncoding(28591).GetString($bytes)

        $pageMatches = [regex]::Matches($text, '/Type\s*/Page(?![sA-Za-z])')
        if ($pageMatches.Count -gt 0) { return $pageMatches.Count }

        $max = 0
        foreach ($pattern in @('/Type\s*/Pages\b[^>]{0,400}?/Count\s+(\d+)',
                               '/Count\s+(\d+)[^>]{0,400}?/Type\s*/Pages\b')) {
            foreach ($hit in [regex]::Matches($text, $pattern)) {
                $v = [int]$hit.Groups[1].Value
                if ($v -gt $max) { $max = $v }
            }
        }
        if ($max -gt 0) { return $max }
        return $null
    }
    catch { return $null }
}

function Measure-PdfPages {
    #     PDF のページ数を数える。
    #     -IgnoreSmall を指定すると、そのPDF内で最大のページ面積に対して
    #     -SmallRatio 未満の面積しかないページを集計から除外する。
    #     返り値: Total（総ページ数） / Counted（集計対象） / Small（除外数）
    param(
        [string] $PdfPath,
        [switch] $IgnoreSmall,
        [double] $SmallRatio = 0.2
    )

    $empty = [pscustomobject]@{ Total = $null; Counted = $null; Small = 0 }

    if ([string]::IsNullOrEmpty($PdfPath)) { return $empty }
    if (-not (Test-Path -LiteralPath $PdfPath)) { return $empty }

    $pages = Get-PdfPageSizes -PdfPath $PdfPath

    if ($null -eq $pages) {
        # ページ個別の解析ができない場合は総数だけ返す（サイズ判定は不可）
        $total = Get-PdfPageCountFallback -PdfPath $PdfPath
        if ($IgnoreSmall -and $null -ne $total) {
            Write-Verbose ("ページサイズを解析できないため、小サイズ除外は適用されません: {0}" -f $PdfPath)
        }
        return [pscustomobject]@{ Total = $total; Counted = $total; Small = 0 }
    }

    $total = $pages.Count
    if (-not $IgnoreSmall) {
        return [pscustomobject]@{ Total = $total; Counted = $total; Small = 0 }
    }

    $maxArea = ($pages | Measure-Object -Property Area -Maximum).Maximum
    if ($null -eq $maxArea -or $maxArea -le 0) {
        return [pscustomobject]@{ Total = $total; Counted = $total; Small = 0 }
    }

    $threshold = $maxArea * $SmallRatio
    $small = @($pages | Where-Object { $_.Area -lt $threshold })

    foreach ($sp in $small) {
        Write-Verbose ("小サイズのため除外: {0} p.{1} {2}x{3}mm" -f
            [System.IO.Path]::GetFileName($PdfPath), $sp.Page, $sp.Width, $sp.Height)
    }

    return [pscustomobject]@{
        Total   = $total
        Counted = $total - $small.Count
        Small   = $small.Count
    }
}

function Get-SheetPageCount {
    # PDF から取得できなかった場合の保険。Excel にページ割りを計算させる（遅い）
    param($Sheet)
    try { return [int]$Sheet.PageSetup.Pages.Count } catch { return $null }
}

$script:WorkbookOpenMode = 'Full'

function Initialize-WorkbookOpenMode {
    #     Workbooks.Open の呼び出し形式を起動時に決める。
    #
    #     省略可能引数（[Type]::Missing）の扱いは PowerShell と Excel のバージョンで差があり、
    #     環境によっては「Workbooks クラスの Open プロパティを取得できません」で失敗する。
    #     自分で作った空ブックで試すので、パスワード入力ダイアログが出る心配はない。
    param($Excel, [string] $Requested)

    if ($Requested -ne 'Auto') {
        $script:WorkbookOpenMode = $Requested
        Write-Verbose "Workbooks.Open の形式: $Requested (指定)"
        return
    }

    $probe = Join-Path $env:TEMP ('xl2pdf_probe_{0}.xlsx' -f [guid]::NewGuid().ToString('N'))
    $mode  = 'Full'
    try {
        $tmp = $Excel.Workbooks.Add()
        try { $tmp.SaveAs($probe, 51) } finally { $tmp.Close($false) }

        try {
            $wb = $Excel.Workbooks.Open($probe, 0, $true, [Type]::Missing, 'x', 'x', $true)
            $wb.Close($false)
            $mode = 'Full'
        }
        catch {
            Write-Verbose ("省略可能引数つきの Open が使えません: {0}" -f $_.Exception.Message)
            $mode = 'Simple'
        }
    }
    catch {
        Write-Verbose ("Open 形式を判定できなかったため Full を使います: {0}" -f $_.Exception.Message)
        $mode = 'Full'
    }
    finally {
        if (Test-Path -LiteralPath $probe) {
            Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
        }
    }

    $script:WorkbookOpenMode = $mode
    Write-Verbose "Workbooks.Open の形式: $mode (自動判定)"
    if ($mode -eq 'Simple') {
        Write-Warning 'この環境では Workbooks.Open の省略可能引数が使えないため、簡易形式で開きます。パスワード保護されたファイルで入力ダイアログが表示される場合があります。'
    }
}

function Open-WorkbookReadOnly {
    param($Excel, [string] $FilePath)

    # Excel はフルパス 218 文字を超えるブックを開けない（OS の 260 文字より手前で失敗する）
    if ($FilePath.Length -gt 218) {
        throw ("パスが長すぎます（{0} 文字）。Excel はフルパス 218 文字までしか開けません。" -f $FilePath.Length)
    }
    if (-not (Test-Path -LiteralPath $FilePath)) {
        throw "ファイルが見つかりません（OneDrive のオンライン専用ファイルの可能性があります）: $FilePath"
    }

    if ($script:WorkbookOpenMode -eq 'Full') {
        # ダミーのパスワードを渡すことで、保護ファイルは入力待ちにならず例外になる
        return $Excel.Workbooks.Open($FilePath, 0, $true, [Type]::Missing, 'x', 'x', $true)
    }
    return $Excel.Workbooks.Open($FilePath, 0, $true)
}

function Set-FitToOnePageWide {
    param($Workbook, $Excel)

    # PageSetup はプロパティ 1 つごとにプリンターと通信するため非常に遅い。
    # PrintCommunication を切ると劇的に速くなる（Excel 2010 以降）。
    try { $Excel.PrintCommunication = $false } catch { }
    try {
        foreach ($sh in $Workbook.Sheets) {
            try {
                $ps = $sh.PageSetup
                $ps.Zoom = $false
                $ps.FitToPagesWide = 1
                $ps.FitToPagesTall = $false
            }
            catch { }
        }
    }
    finally {
        try { $Excel.PrintCommunication = $true } catch { }
    }
}

function New-ResultRecord {
    param(
        [string] $SourcePath,
        [string] $PdfPath,
        [string] $Status,
        [string] $Detail,
        $PageCount,
        [int] $SmallPages = 0,
        [double] $Seconds
    )
    return [pscustomobject]@{
        ファイル = $SourcePath
        出力先   = $PdfPath
        ページ数 = $PageCount
        除外ページ = $SmallPages
        結果     = $Status
        詳細     = $Detail
        秒       = [math]::Round($Seconds, 1)
    }
}

# ---------------------------------------------------------------- 結果の表示

function Write-RunSummary {
    #     処理結果の一覧とサマリを表示し、-LogPath があれば CSV に保存する。
    #     スクリプトスコープの $results / $root / $CountOnly などを参照する。
    $ok      = @($results | Where-Object { $_.結果 -eq '成功' }).Count
    $skipped = @($results | Where-Object { $_.結果 -eq 'スキップ' }).Count
    $failed  = @($results | Where-Object { $_.結果 -eq '失敗' }).Count

    $pageSum = ($results | Where-Object { $null -ne $_.ページ数 } | Measure-Object -Property ページ数 -Sum).Sum
    if ($null -eq $pageSum) { $pageSum = 0 }
    $smallSum = ($results | Measure-Object -Property 除外ページ -Sum).Sum
    if ($null -eq $smallSum) { $smallSum = 0 }
    $unknownPages = @($results | Where-Object { $_.結果 -eq '成功' -and $null -eq $_.ページ数 }).Count

    $view = @(
        @{ Name = 'ファイル'; Expression = { $_.ファイル.Substring($root.Length).TrimStart('\') } }
    )
    if (-not $CountOnly) {
        $view += @{ Name = '出力PDF'; Expression = { if ($_.出力先) { [System.IO.Path]::GetFileName($_.出力先) } else { '-' } } }
    }
    $view += @{ Name = 'ページ'; Expression = { if ($null -ne $_.ページ数) { $_.ページ数 } else { '-' } } }
    if ($IgnoreSmallPages) {
        $view += @{ Name = '除外'; Expression = { if ($_.除外ページ -gt 0) { $_.除外ページ } else { '' } } }
    }
    $view += '結果'
    $view += '詳細'

    Write-Host ""
    if ($CountOnly) {
        Write-Host "==== ページ数一覧 ========================================================"
    }
    else {
        Write-Host "==== 出力ファイル一覧 ===================================================="
    }
    $results | Select-Object $view | Format-Table -AutoSize | Out-Host

    Write-Host ("成功 {0} / スキップ {1} / 失敗 {2}" -f $ok, $skipped, $failed)
    Write-Host ("合計ページ数 {0}" -f $pageSum)
    if ($IgnoreSmallPages) {
        Write-Host ("小サイズとして除外したページ {0}（最大ページ面積の {1:P0} 未満）" -f $smallSum, $SmallPageRatio)
        if ($smallSum -eq 0) {
            Write-Host "※ 除外が0件の場合は -SmallPageRatio を上げるか、-Verbose で各ページの寸法を確認してください"
        }
    }
    if ($unknownPages -gt 0) {
        Write-Host ("※ うち {0} 件はページ数を取得できませんでした（一覧では - と表示）" -f $unknownPages)
    }

    if ($LogPath) {
        # PowerShell 5.1 の UTF8 は BOM 付き、7 以降は BOM 無しなので明示的に切り替える
        if ($PSVersionTable.PSVersion.Major -ge 6) { $enc = 'utf8BOM' } else { $enc = 'UTF8' }
        $results | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding $enc
        Write-Host ("ログ: {0}" -f $LogPath)
    }
}

# ---------------------------------------------------------------- 入力の検証

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "フォルダが見つかりません: $Path"
}

# 正規表現は実行前に検証しておく（全ファイル処理してから落ちるのを防ぐ）
foreach ($item in @(@{ Name = 'SheetNamePattern'; Value = $SheetNamePattern },
                    @{ Name = 'SheetNameExclude'; Value = $SheetNameExclude })) {
    if ($item.Value) {
        try { [void][regex]::IsMatch('', $item.Value) }
        catch { throw ("-{0} の正規表現が不正です: {1}  ({2})" -f $item.Name, $item.Value, $_.Exception.Message) }
    }
}
$root = (Resolve-Path -LiteralPath $Path).ProviderPath.TrimEnd('\')

if ($OutputRoot) {
    if (-not (Test-Path -LiteralPath $OutputRoot)) {
        if ($PSCmdlet.ShouldProcess($OutputRoot, '出力フォルダを作成')) {
            New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
        }
    }
    if (Test-Path -LiteralPath $OutputRoot) {
        $OutputRoot = (Resolve-Path -LiteralPath $OutputRoot).ProviderPath.TrimEnd('\')
    }
}

# ---------------------------------------------------------------- 対象ファイルの収集

if ($CountOnly) {
    # PDF を出力せず、既にある PDF のページ数を数えるだけのモード
    $Extensions = @('.pdf')
    foreach ($opt in @('OutputRoot', 'PerSheet', 'FitToWidth', 'SkipExisting', 'IncludeHiddenSheets',
                       'SheetNamePattern', 'SheetNameExclude')) {
        if ($PSBoundParameters.ContainsKey($opt)) {
            Write-Warning "-CountOnly 指定時、-$opt は無視されます。"
        }
    }
}

$gci = @{
    LiteralPath   = $root
    File          = $true
    ErrorAction   = 'SilentlyContinue'
    ErrorVariable = 'scanErrors'
}
if (-not $NoRecurse) { $gci['Recurse'] = $true }

$files = @(
    Get-ChildItem @gci |
        Where-Object {
            $Extensions -contains $_.Extension.ToLowerInvariant() -and
            $_.Name -notlike '~$*'
        } |
        Sort-Object FullName
)

if ($scanErrors) {
    foreach ($e in $scanErrors) { Write-Warning ("走査できないパス: " + $e.TargetObject) }
}

if ($files.Count -eq 0) {
    if ($CountOnly) { Write-Warning "PDF が見つかりません: $root" }
    else { Write-Warning "対象ファイルがありません: $root" }
    return
}

if ($CountOnly) {
    Write-Host ("対象 {0} PDF / 走査元 {1}" -f $files.Count, $root)
}
else {
    Write-Host ("対象 {0} ファイル / 走査元 {1}" -f $files.Count, $root)
}

$results = New-Object System.Collections.Generic.List[object]

# ---------------------------------------------------------------- ページ数カウントのみ

if ($CountOnly) {
    $index = 0
    foreach ($file in $files) {
        $index++
        Write-Progress -Activity 'ページ数カウント' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $file.Name) -PercentComplete ([int](100 * $index / $files.Count))

        $pi = Measure-PdfPages -PdfPath $file.FullName -IgnoreSmall:$IgnoreSmallPages -SmallRatio $SmallPageRatio

        if ($null -eq $pi.Counted) {
            $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $file.FullName `
                -Status '失敗' -Detail 'ページ数を取得できませんでした（破損または未対応の PDF）' `
                -PageCount $null -SmallPages 0 -Seconds 0))
        }
        else {
            $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $file.FullName `
                -Status '成功' -Detail '' -PageCount $pi.Counted -SmallPages $pi.Small -Seconds 0))
        }
    }
    Write-Progress -Activity 'ページ数カウント' -Completed

    Write-RunSummary
    if (@($results | Where-Object { $_.結果 -eq '失敗' }).Count -gt 0) { exit 1 }
    return
}

# ---------------------------------------------------------------- Excel 起動

# 既存の EXCEL プロセスを覚えておき、後始末では自分が起動した分だけを終了させる
$pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

$excel   = $null
$ownPids = @()

try {
    try {
        $excel = New-Object -ComObject Excel.Application
    }
    catch {
        throw "Excel を起動できません。デスクトップ版 Excel がインストールされているか確認してください。 詳細: $($_.Exception.Message)"
    }

    $pidsAfter = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $ownPids   = @($pidsAfter | Where-Object { $pidsBefore -notcontains $_ })

    $excel.Visible            = $false
    $excel.DisplayAlerts      = $false
    $excel.AskToUpdateLinks   = $false
    $excel.EnableEvents       = $false
    $excel.ScreenUpdating     = $false
    try {
        $excel.AutomationSecurity = $msoAutomationSecurityForceDisable
    }
    catch {
        Write-Warning ("AutomationSecurity を設定できませんでした（マクロが無効化されません）: {0}" -f $_.Exception.Message)
    }

    Initialize-WorkbookOpenMode -Excel $excel -Requested $OpenMode

    $index = 0
    foreach ($file in $files) {
        $index++
        Write-Progress -Activity 'PDF 出力' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $file.Name) -PercentComplete ([int](100 * $index / $files.Count))

        # 出力先フォルダを決める
        if ($OutputRoot) {
            $srcDir = [System.IO.Path]::GetDirectoryName($file.FullName)
            $relDir = $srcDir.Substring($root.Length).TrimStart('\')
            if ([string]::IsNullOrEmpty($relDir)) { $outDir = $OutputRoot }
            else { $outDir = Join-Path $OutputRoot $relDir }
        }
        else {
            $outDir = $file.DirectoryName
        }

        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $wb = $null
        $exported = New-Object System.Collections.Generic.List[object]

        try {
            if (-not (Test-Path -LiteralPath $outDir)) {
                if ($PSCmdlet.ShouldProcess($outDir, '出力フォルダを作成')) {
                    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
                }
            }

            # ブック単位のときは、開く前にスキップ判定して無駄な起動を避ける
            if (-not $PerSheet) {
                $pdfPath = Join-Path $outDir ($file.BaseName + '.pdf')
                if ($SkipExisting -and (Test-Path -LiteralPath $pdfPath)) {
                    $pdfItem = Get-Item -LiteralPath $pdfPath
                    if ($pdfItem.LastWriteTime -ge $file.LastWriteTime) {
                        $pi = Measure-PdfPages -PdfPath $pdfPath -IgnoreSmall:$IgnoreSmallPages -SmallRatio $SmallPageRatio
                        $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $pdfPath `
                            -Status 'スキップ' -Detail '既存PDFが新しい' `
                            -PageCount $pi.Counted -SmallPages $pi.Small -Seconds 0))
                        continue
                    }
                }
            }

            if (-not $PSCmdlet.ShouldProcess($file.FullName, 'PDF 出力')) {
                $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $outDir `
                    -Status 'WhatIf' -Detail '' -PageCount $null -Seconds 0))
                continue
            }

            $wb = Open-WorkbookReadOnly -Excel $excel -FilePath $file.FullName

            if ($FitToWidth) { Set-FitToOnePageWide -Workbook $wb -Excel $excel }

            if ($PerSheet) {
                foreach ($sh in $wb.Sheets) {
                    $isVisible = ($sh.Visible -eq $xlSheetVisible)
                    if (-not $isVisible) {
                        if (-not $IncludeHiddenSheets) { continue }
                        $sh.Visible = $xlSheetVisible   # 非表示シートは出力できないため一時的に表示
                    }
                    if (-not (Test-SheetNameMatch $sh.Name)) { continue }
                    if (-not (Test-HasPrintableContent $sh)) { continue }

                    $sheetPdf = Join-Path $outDir ("{0}_{1}.pdf" -f $file.BaseName, (ConvertTo-SafeFileName $sh.Name))

                    if ($SkipExisting -and (Test-Path -LiteralPath $sheetPdf)) {
                        $existing = Get-Item -LiteralPath $sheetPdf
                        if ($existing.LastWriteTime -ge $file.LastWriteTime) {
                            $pi = Measure-PdfPages -PdfPath $sheetPdf -IgnoreSmall:$IgnoreSmallPages -SmallRatio $SmallPageRatio
                            $exported.Add([pscustomobject]@{
                                Path   = $sheetPdf
                                Pages  = $pi.Counted
                                Small  = $pi.Small
                                Status = 'スキップ'
                                Detail = '既存PDFが新しい'
                            })
                            continue
                        }
                    }
                    else {
                        $sheetPdf = Get-UniquePath $sheetPdf
                    }

                    $sh.ExportAsFixedFormat($xlTypePDF, $sheetPdf, $xlQualityStandard, $true, $false, [Type]::Missing, [Type]::Missing, $false)

                    $pi = Measure-PdfPages -PdfPath $sheetPdf -IgnoreSmall:$IgnoreSmallPages -SmallRatio $SmallPageRatio
                    $pages = $pi.Counted
                    if ($null -eq $pages) { $pages = Get-SheetPageCount $sh }

                    $exported.Add([pscustomobject]@{
                        Path = $sheetPdf; Pages = $pages; Small = $pi.Small; Status = '成功'; Detail = ''
                    })
                }

                if ($exported.Count -eq 0) {
                    $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath '' `
                        -Status 'スキップ' -Detail $(if ($SheetNamePattern -or $SheetNameExclude) { '条件に一致する出力対象シートなし' } else { '出力対象シートなし' }) -PageCount $null `
                        -Seconds $sw.Elapsed.TotalSeconds))
                }
                else {
                    foreach ($item in $exported) {
                        $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $item.Path `
                            -Status $item.Status -Detail $item.Detail -PageCount $item.Pages `
                            -SmallPages $item.Small -Seconds $sw.Elapsed.TotalSeconds))
                    }
                }
            }
            else {
                # シート名で絞る場合、対象外のシートを一時的に非表示にする。
                # 非表示シートは ExportAsFixedFormat の対象外になるため、
                # シート選択 API を使わずに出力範囲を限定できる。
                # 読み取り専用で開いており保存しないので、元ファイルは変更されない。
                $noSheetMatch = $false
                if ($SheetNamePattern -or $SheetNameExclude) {
                    $matched = 0
                    foreach ($sh in $wb.Sheets) {
                        if ($sh.Visible -eq $xlSheetVisible -and (Test-SheetNameMatch $sh.Name)) { $matched++ }
                    }

                    if ($matched -eq 0) {
                        # 1 枚も一致しない。Excel は全シートを非表示にできないため、
                        # 非表示化はせずこのブックごとスキップする
                        $noSheetMatch = $true
                    }
                    else {
                        foreach ($sh in $wb.Sheets) {
                            if ($sh.Visible -eq $xlSheetVisible -and -not (Test-SheetNameMatch $sh.Name)) {
                                try { $sh.Visible = $xlSheetHidden }
                                catch { Write-Verbose ("シートを非表示にできませんでした: {0}" -f $sh.Name) }
                            }
                        }
                    }
                }

                # 印刷対象が 1 つも無いブックは ExportAsFixedFormat が例外になるため事前に判定する
                $printableSheets = New-Object System.Collections.Generic.List[object]
                if (-not $noSheetMatch) {
                    foreach ($sh in $wb.Sheets) {
                        if ($sh.Visible -eq $xlSheetVisible -and (Test-HasPrintableContent $sh)) {
                            $printableSheets.Add($sh)
                        }
                    }
                }

                if ($printableSheets.Count -eq 0) {
                    $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath '' `
                        -Status 'スキップ' -Detail $(if ($SheetNamePattern -or $SheetNameExclude) { '条件に一致する印刷対象シートなし' } else { '印刷対象なし（空ブック／全シート非表示）' }) `
                        -PageCount $null -Seconds $sw.Elapsed.TotalSeconds))
                }
                else {
                    $wb.ExportAsFixedFormat($xlTypePDF, $pdfPath, $xlQualityStandard, $true, $false, [Type]::Missing, [Type]::Missing, $false)

                    $pi = Measure-PdfPages -PdfPath $pdfPath -IgnoreSmall:$IgnoreSmallPages -SmallRatio $SmallPageRatio
                    $pages = $pi.Counted
                    if ($null -eq $pages) {
                        # PDF から読めなかった場合のみ Excel に計算させる
                        $sum = 0
                        $got = $false
                        foreach ($sh in $printableSheets) {
                            $n = Get-SheetPageCount $sh
                            if ($null -ne $n) { $sum += $n; $got = $true }
                        }
                        if ($got) { $pages = $sum }
                    }

                    $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $pdfPath `
                        -Status '成功' -Detail '' -PageCount $pages -SmallPages $pi.Small `
                        -Seconds $sw.Elapsed.TotalSeconds))
                }
            }
        }
        catch {
            Write-Warning ("失敗: {0} : {1}" -f $file.FullName, $_.Exception.Message)
            $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath '' `
                -Status '失敗' -Detail $_.Exception.Message -PageCount $null `
                -Seconds $sw.Elapsed.TotalSeconds))
        }
        finally {
            if ($null -ne $wb) {
                try { $wb.Close($false) } catch { }
                try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
                $wb = $null
            }
            $sw.Stop()
        }
    }
}
finally {
    Write-Progress -Activity 'PDF 出力' -Completed

    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch { }
        $excel = $null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    # Quit しても残ることがあるため、自分が起動したプロセスだけ確実に落とす
    foreach ($procId in $ownPids) {
        $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
        if ($p) {
            Start-Sleep -Seconds 2
            $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
            if ($p) {
                Write-Verbose "残存した Excel プロセスを終了します (PID: $procId)"
                Stop-Process -Id $procId -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# ---------------------------------------------------------------- 結果

Write-RunSummary
if (@($results | Where-Object { $_.結果 -eq '失敗' }).Count -gt 0) { exit 1 }

# ================================================================
# 以下はコメントベースヘルプ（Get-Help .\Export-ExcelToPdf.ps1 -Full で表示）
# スクリプトでは末尾に置いても有効。param() より前には何も置かないこと。
# ================================================================
# .SYNOPSIS
#     フォルダ配下の Excel ファイルを PDF に一括出力します。
#
# .DESCRIPTION
#     指定フォルダ（既定でサブフォルダも含む）内の .xls / .xlsx / .xlsm / .xlsb を
#     Excel COM で開き、ExportAsFixedFormat で PDF に出力します。
#     処理後に、ファイルごとの出力先とページ数の一覧を表示します。
#
#     - マクロは強制無効で開くため、xlsm を安全に処理できます
#     - 元ファイルは読み取り専用で開き、保存しません
#     - 既に起動しているユーザーの Excel は操作せず、専用インスタンスを使います
#     - パスワード保護ファイルは入力待ちで止まらず、エラーとして記録します
#
# .PARAMETER Path
#     走査するフォルダ。
#
# .PARAMETER OutputRoot
#     PDF の出力先ルート。省略時は元ファイルと同じ場所に出力します。
#     指定した場合は Path 配下のフォルダ構成を再現します。
#
# .PARAMETER NoRecurse
#     サブフォルダを走査しません。
#
# .PARAMETER PerSheet
#     ブック単位ではなくシート単位で出力します（ファイル名_シート名.pdf）。
#
# .PARAMETER SkipExisting
#     出力先 PDF が既に存在し、元ファイルより新しい場合はスキップします。
#     スキップした場合も、既存 PDF のページ数を一覧に表示します。
#
# .PARAMETER FitToWidth
#     全シートを「横 1 ページに収める」設定にしてから出力します（処理は遅くなります）。
#
# .PARAMETER IncludeHiddenSheets
#     非表示シートも出力対象にします（PerSheet 指定時のみ有効）。
#
# .PARAMETER SheetNamePattern
#     処理対象とするシート名の正規表現（.NET 正規表現、大文字小文字は区別しません）。
#     指定すると、一致するシートだけを PDF 出力し、ページ数もそのシートのみ数えます。
#     ブック単位の出力では、対象外のシートを一時的に非表示にして出力範囲から外します
#     （読み取り専用で開いており保存しないため、元ファイルは変更されません）。
#     例: -SheetNamePattern '^(画面|帳票)'
#
# .PARAMETER SheetNameExclude
#     除外するシート名の正規表現。SheetNamePattern と併用でき、除外が優先されます。
#     例: -SheetNameExclude '表紙|改訂履歴|目次'
#
# .PARAMETER CountOnly
#     PDF を出力せず、フォルダ配下にある既存の PDF のページ数を数えるだけのモードです。
#     Excel を起動しないため高速で、Excel がインストールされていない環境でも動作します。
#     このモードでは Path 配下の *.pdf が対象になります（Excel ファイルは見ません）。
#
# .PARAMETER IgnoreSmallPages
#     用紙サイズが極端に小さいページをページ数の集計から除外します。
#     ブック内のシートごとに用紙設定が異なり、意図しない小さなページが混ざる場合に使います。
#     PDF 自体は変更せず、集計から外すだけです。
#
# .PARAMETER SmallPageRatio
#     「小さいページ」と判定するしきい値。そのPDF内で最大のページ面積に対する比率で、
#     既定は 0.2（最大ページの面積の 20% 未満なら除外）。-Verbose を付けると
#     除外されたページの寸法が表示されるので、しきい値の調整に使えます。
#
# .PARAMETER OpenMode
#     Workbooks.Open の呼び出し形式。既定の Auto は起動時に自動判定します。
#     「Workbooks クラスの Open プロパティを取得できません」というエラーが出る場合は
#     Simple を指定してください。
#
# .PARAMETER LogPath
#     処理結果（ページ数を含む）を CSV（UTF-8 BOM 付き）で保存します。
#
# .EXAMPLE
#     .\Export-ExcelToPdf.ps1 -Path D:\Docs -OutputRoot D:\Pdf -LogPath D:\Pdf\result.csv
#
# .EXAMPLE
#     .\Export-ExcelToPdf.ps1 -Path D:\Docs -PerSheet -FitToWidth
#
# .EXAMPLE
#     # 「画面」で始まるシートだけを対象にページ数を数える
#     .\Export-ExcelToPdf.ps1 -Path D:\Docs -SheetNamePattern '^画面'
#
# .EXAMPLE
#     # 表紙・改訂履歴を除いたページ数を数える
#     .\Export-ExcelToPdf.ps1 -Path D:\Docs -SheetNameExclude '表紙|改訂履歴|目次'
#
# .EXAMPLE
#     # PDF は作らず、既存 PDF のページ数だけ数える
#     .\Export-ExcelToPdf.ps1 -Path D:\Pdf -CountOnly
#
# .EXAMPLE
#     # 何が出力されるかだけ確認する
#     .\Export-ExcelToPdf.ps1 -Path D:\Docs -WhatIf
#
# .NOTES
#     要件: Windows + Excel（デスクトップ版）がインストールされていること。
#           対話的なデスクトップセッションで実行してください。
