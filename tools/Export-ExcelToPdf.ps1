<#
.SYNOPSIS
    フォルダ配下の Excel ファイルを PDF に一括出力します。

.DESCRIPTION
    指定フォルダ（既定でサブフォルダも含む）内の .xls / .xlsx / .xlsm / .xlsb を
    Excel COM で開き、ExportAsFixedFormat で PDF に出力します。
    処理後に、ファイルごとの出力先とページ数の一覧を表示します。

    - マクロは強制無効で開くため、xlsm を安全に処理できます
    - 元ファイルは読み取り専用で開き、保存しません
    - 既に起動しているユーザーの Excel は操作せず、専用インスタンスを使います
    - パスワード保護ファイルは入力待ちで止まらず、エラーとして記録します

.PARAMETER Path
    走査するフォルダ。

.PARAMETER OutputRoot
    PDF の出力先ルート。省略時は元ファイルと同じ場所に出力します。
    指定した場合は Path 配下のフォルダ構成を再現します。

.PARAMETER NoRecurse
    サブフォルダを走査しません。

.PARAMETER PerSheet
    ブック単位ではなくシート単位で出力します（ファイル名_シート名.pdf）。

.PARAMETER SkipExisting
    出力先 PDF が既に存在し、元ファイルより新しい場合はスキップします。
    スキップした場合も、既存 PDF のページ数を一覧に表示します。

.PARAMETER FitToWidth
    全シートを「横 1 ページに収める」設定にしてから出力します（処理は遅くなります）。

.PARAMETER IncludeHiddenSheets
    非表示シートも出力対象にします（PerSheet 指定時のみ有効）。

.PARAMETER LogPath
    処理結果（ページ数を含む）を CSV（UTF-8 BOM 付き）で保存します。

.EXAMPLE
    .\Export-ExcelToPdf.ps1 -Path D:\Docs -OutputRoot D:\Pdf -LogPath D:\Pdf\result.csv

.EXAMPLE
    .\Export-ExcelToPdf.ps1 -Path D:\Docs -PerSheet -FitToWidth

.EXAMPLE
    # 何が出力されるかだけ確認する
    .\Export-ExcelToPdf.ps1 -Path D:\Docs -WhatIf

.NOTES
    要件: Windows + Excel（デスクトップ版）がインストールされていること。
          対話的なデスクトップセッションで実行してください。
#>
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

    [string] $LogPath,

    [string[]] $Extensions = @('.xls', '.xlsx', '.xlsm', '.xlsb')
)

$ErrorActionPreference = 'Stop'

# ---- Excel 定数（COM の列挙体を PowerShell から参照しないで済むよう定義） ----
$xlTypePDF         = 0
$xlQualityStandard = 0
$xlSheetVisible    = -1
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

function Get-PdfPageCount {
    <#
        生成された PDF のページ数を、外部ツールなしで取得する。
        1. /Type /Page の出現数を数える（Excel が出力する PDF はこれで取れる）
        2. 取れない場合はページツリーの /Count を読む
           （オブジェクトストリームで圧縮された PDF への保険）
        取得できなければ $null を返す。
    #>
    param([string] $PdfPath)

    if ([string]::IsNullOrEmpty($PdfPath)) { return $null }
    if (-not (Test-Path -LiteralPath $PdfPath)) { return $null }

    try {
        $bytes = [System.IO.File]::ReadAllBytes($PdfPath)
        # バイト列を 1 バイト 1 文字として扱う（PDF 構造部は ASCII のため）
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
    catch {
        Write-Verbose ("ページ数を取得できませんでした: {0} : {1}" -f $PdfPath, $_.Exception.Message)
        return $null
    }
}

function Get-SheetPageCount {
    # PDF から取得できなかった場合の保険。Excel にページ割りを計算させる（遅い）
    param($Sheet)
    try { return [int]$Sheet.PageSetup.Pages.Count } catch { return $null }
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
        [double] $Seconds
    )
    return [pscustomobject]@{
        ファイル = $SourcePath
        出力先   = $PdfPath
        ページ数 = $PageCount
        結果     = $Status
        詳細     = $Detail
        秒       = [math]::Round($Seconds, 1)
    }
}

# ---------------------------------------------------------------- 入力の検証

if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    throw "フォルダが見つかりません: $Path"
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
    Write-Warning "対象ファイルがありません: $root"
    return
}

Write-Host ("対象 {0} ファイル / 走査元 {1}" -f $files.Count, $root)

# ---------------------------------------------------------------- Excel 起動

# 既存の EXCEL プロセスを覚えておき、後始末では自分が起動した分だけを終了させる
$pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })

$excel   = $null
$ownPids = @()
$results = New-Object System.Collections.Generic.List[object]

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
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

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
                        $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $pdfPath `
                            -Status 'スキップ' -Detail '既存PDFが新しい' `
                            -PageCount (Get-PdfPageCount -PdfPath $pdfPath) -Seconds 0))
                        continue
                    }
                }
            }

            if (-not $PSCmdlet.ShouldProcess($file.FullName, 'PDF 出力')) {
                $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $outDir `
                    -Status 'WhatIf' -Detail '' -PageCount $null -Seconds 0))
                continue
            }

            # ダミーのパスワードを渡すことで、保護ファイルは入力待ちにならず例外になる
            $wb = $excel.Workbooks.Open(
                $file.FullName, 0, $true, [Type]::Missing, 'x', 'x', $true,
                [Type]::Missing, [Type]::Missing, $false, $false,
                [Type]::Missing, $false, $true, [Type]::Missing)

            if ($FitToWidth) { Set-FitToOnePageWide -Workbook $wb -Excel $excel }

            if ($PerSheet) {
                foreach ($sh in $wb.Sheets) {
                    $isVisible = ($sh.Visible -eq $xlSheetVisible)
                    if (-not $isVisible) {
                        if (-not $IncludeHiddenSheets) { continue }
                        $sh.Visible = $xlSheetVisible   # 非表示シートは出力できないため一時的に表示
                    }
                    if (-not (Test-HasPrintableContent $sh)) { continue }

                    $sheetPdf = Join-Path $outDir ("{0}_{1}.pdf" -f $file.BaseName, (ConvertTo-SafeFileName $sh.Name))

                    if ($SkipExisting -and (Test-Path -LiteralPath $sheetPdf)) {
                        $existing = Get-Item -LiteralPath $sheetPdf
                        if ($existing.LastWriteTime -ge $file.LastWriteTime) {
                            $exported.Add([pscustomobject]@{
                                Path = $sheetPdf
                                Pages = (Get-PdfPageCount -PdfPath $sheetPdf)
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

                    $pages = Get-PdfPageCount -PdfPath $sheetPdf
                    if ($null -eq $pages) { $pages = Get-SheetPageCount $sh }

                    $exported.Add([pscustomobject]@{
                        Path = $sheetPdf; Pages = $pages; Status = '成功'; Detail = ''
                    })
                }

                if ($exported.Count -eq 0) {
                    $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath '' `
                        -Status 'スキップ' -Detail '出力対象シートなし' -PageCount $null `
                        -Seconds $sw.Elapsed.TotalSeconds))
                }
                else {
                    foreach ($item in $exported) {
                        $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath $item.Path `
                            -Status $item.Status -Detail $item.Detail -PageCount $item.Pages `
                            -Seconds $sw.Elapsed.TotalSeconds))
                    }
                }
            }
            else {
                # 印刷対象が 1 つも無いブックは ExportAsFixedFormat が例外になるため事前に判定する
                $printableSheets = New-Object System.Collections.Generic.List[object]
                foreach ($sh in $wb.Sheets) {
                    if ($sh.Visible -eq $xlSheetVisible -and (Test-HasPrintableContent $sh)) {
                        $printableSheets.Add($sh)
                    }
                }

                if ($printableSheets.Count -eq 0) {
                    $results.Add((New-ResultRecord -SourcePath $file.FullName -PdfPath '' `
                        -Status 'スキップ' -Detail '印刷対象なし（空ブック／全シート非表示）' `
                        -PageCount $null -Seconds $sw.Elapsed.TotalSeconds))
                }
                else {
                    $wb.ExportAsFixedFormat($xlTypePDF, $pdfPath, $xlQualityStandard, $true, $false, [Type]::Missing, [Type]::Missing, $false)

                    $pages = Get-PdfPageCount -PdfPath $pdfPath
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
                        -Status '成功' -Detail '' -PageCount $pages -Seconds $sw.Elapsed.TotalSeconds))
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

# ---------------------------------------------------------------- 結果の一覧

$ok      = @($results | Where-Object { $_.結果 -eq '成功' }).Count
$skipped = @($results | Where-Object { $_.結果 -eq 'スキップ' }).Count
$failed  = @($results | Where-Object { $_.結果 -eq '失敗' }).Count

$pageSum = ($results | Where-Object { $null -ne $_.ページ数 } | Measure-Object -Property ページ数 -Sum).Sum
if ($null -eq $pageSum) { $pageSum = 0 }
$unknownPages = @($results | Where-Object { $_.結果 -eq '成功' -and $null -eq $_.ページ数 }).Count

$view = @(
    @{ Name = 'ファイル'; Expression = { $_.ファイル.Substring($root.Length).TrimStart('\') } },
    @{ Name = '出力PDF';  Expression = { if ($_.出力先) { [System.IO.Path]::GetFileName($_.出力先) } else { '-' } } },
    @{ Name = 'ページ';   Expression = { if ($null -ne $_.ページ数) { $_.ページ数 } else { '-' } } },
    '結果',
    '詳細'
)

Write-Host ""
Write-Host "==== 出力ファイル一覧 ===================================================="
$results | Select-Object $view | Format-Table -AutoSize | Out-Host

Write-Host ("成功 {0} / スキップ {1} / 失敗 {2}" -f $ok, $skipped, $failed)
Write-Host ("合計ページ数 {0}" -f $pageSum)
if ($unknownPages -gt 0) {
    Write-Host ("※ うち {0} 件はページ数を取得できませんでした（一覧では - と表示）" -f $unknownPages)
}

if ($LogPath) {
    # PowerShell 5.1 の UTF8 は BOM 付き、7 以降は BOM 無しなので明示的に切り替える
    if ($PSVersionTable.PSVersion.Major -ge 6) { $enc = 'utf8BOM' } else { $enc = 'UTF8' }
    $results | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding $enc
    Write-Host ("ログ: {0}" -f $LogPath)
}

if ($failed -gt 0) { exit 1 }
