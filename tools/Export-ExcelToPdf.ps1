<#
.SYNOPSIS
    フォルダ配下の Excel ファイルを PDF に一括出力します。

.DESCRIPTION
    指定フォルダ（既定でサブフォルダも含む）内の .xls / .xlsx / .xlsm / .xlsb を
    Excel COM で開き、ExportAsFixedFormat で PDF に出力します。

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

.PARAMETER FitToWidth
    全シートを「横 1 ページに収める」設定にしてから出力します（処理は遅くなります）。

.PARAMETER IncludeHiddenSheets
    非表示シートも出力対象にします（PerSheet 指定時のみ有効）。

.PARAMETER LogPath
    処理結果を CSV（UTF-8 BOM 付き）で保存します。

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
    LiteralPath = $root
    File        = $true
    ErrorAction = 'SilentlyContinue'
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

    $excel.Visible          = $false
    $excel.DisplayAlerts    = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents     = $false
    $excel.ScreenUpdating   = $false
    $excel.AutomationSecurity = $msoAutomationSecurityForceDisable

    $index = 0
    foreach ($file in $files) {
        $index++
        Write-Progress -Activity 'PDF 出力' `
                       -Status ("{0}/{1}  {2}" -f $index, $files.Count, $file.Name) `
                       -PercentComplete ([int](100 * $index / $files.Count))

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
        $exported = @()

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
                        $results.Add([pscustomobject]@{
                            ファイル = $file.FullName; 出力先 = $pdfPath
                            結果 = 'スキップ'; 詳細 = '既存PDFが新しい'; 秒 = 0
                        })
                        continue
                    }
                }
            }

            if (-not $PSCmdlet.ShouldProcess($file.FullName, 'PDF 出力')) {
                $results.Add([pscustomobject]@{
                    ファイル = $file.FullName; 出力先 = $outDir
                    結果 = 'WhatIf'; 詳細 = ''; 秒 = 0
                })
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
                        if ($existing.LastWriteTime -ge $file.LastWriteTime) { continue }
                    }
                    else {
                        $sheetPdf = Get-UniquePath $sheetPdf
                    }

                    $sh.ExportAsFixedFormat($xlTypePDF, $sheetPdf, $xlQualityStandard, $true, $false,
                                            [Type]::Missing, [Type]::Missing, $false)
                    $exported += $sheetPdf
                }

                if ($exported.Count -eq 0) {
                    $results.Add([pscustomobject]@{
                        ファイル = $file.FullName; 出力先 = ''
                        結果 = 'スキップ'; 詳細 = '出力対象シートなし'; 秒 = [math]::Round($sw.Elapsed.TotalSeconds, 1)
                    })
                }
                else {
                    foreach ($p in $exported) {
                        $results.Add([pscustomobject]@{
                            ファイル = $file.FullName; 出力先 = $p
                            結果 = '成功'; 詳細 = ''; 秒 = [math]::Round($sw.Elapsed.TotalSeconds, 1)
                        })
                    }
                }
            }
            else {
                # 印刷対象が 1 つも無いブックは ExportAsFixedFormat が例外になるため事前に判定する
                $printable = 0
                foreach ($sh in $wb.Sheets) {
                    if ($sh.Visible -eq $xlSheetVisible -and (Test-HasPrintableContent $sh)) { $printable++ }
                }

                if ($printable -eq 0) {
                    $results.Add([pscustomobject]@{
                        ファイル = $file.FullName; 出力先 = ''
                        結果 = 'スキップ'; 詳細 = '印刷対象なし（空ブック／全シート非表示）'
                        秒 = [math]::Round($sw.Elapsed.TotalSeconds, 1)
                    })
                }
                else {
                    $wb.ExportAsFixedFormat($xlTypePDF, $pdfPath, $xlQualityStandard, $true, $false,
                                            [Type]::Missing, [Type]::Missing, $false)
                    $results.Add([pscustomobject]@{
                        ファイル = $file.FullName; 出力先 = $pdfPath
                        結果 = '成功'; 詳細 = ''; 秒 = [math]::Round($sw.Elapsed.TotalSeconds, 1)
                    })
                }
            }
        }
        catch {
            Write-Warning ("失敗: {0} : {1}" -f $file.FullName, $_.Exception.Message)
            $results.Add([pscustomobject]@{
                ファイル = $file.FullName; 出力先 = ''
                結果 = '失敗'; 詳細 = $_.Exception.Message
                秒 = [math]::Round($sw.Elapsed.TotalSeconds, 1)
            })
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

$ok      = @($results | Where-Object { $_.結果 -eq '成功' }).Count
$skipped = @($results | Where-Object { $_.結果 -eq 'スキップ' }).Count
$failed  = @($results | Where-Object { $_.結果 -eq '失敗' }).Count

Write-Host ""
Write-Host ("完了  成功 {0} / スキップ {1} / 失敗 {2}" -f $ok, $skipped, $failed)

if ($LogPath) {
    # PowerShell 5.1 の UTF8 は BOM 付き、7 以降は BOM 無しなので明示的に切り替える
    if ($PSVersionTable.PSVersion.Major -ge 6) { $enc = 'utf8BOM' } else { $enc = 'UTF8' }
    $results | Export-Csv -LiteralPath $LogPath -NoTypeInformation -Encoding $enc
    Write-Host ("ログ: {0}" -f $LogPath)
}

if ($failed -gt 0) { exit 1 }
