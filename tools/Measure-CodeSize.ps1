[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [string] $Baseline,

    [ValidateSet('Folder', 'File')]
    [string] $GroupBy = 'Folder',

    [ValidateRange(1, 10)]
    [int] $Depth = 1,

    [string[]] $Exclude = @('.git', '.svn', 'node_modules', 'bin', 'obj', 'target', 'build', 'dist', '.vs', '.idea'),

    [string] $OutCsv,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

foreach ($name in @('OutCsv', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'CodeSize.Group' -DefaultDisplayPropertySet Group, Files, Lines, Comment, Blank, Steps, Added, Changed, Deleted -Force -WhatIf:$false

function Get-SourceFiles {
    #     数えるファイル（言語の拡張子のもの）。相対パス → FileInfo
    param([string] $Root)
    $result = [ordered]@{}
    $rootItem = Get-Item -LiteralPath $Root
    if (-not $rootItem.PSIsContainer) {
        if (Get-SourceLanguage $rootItem.Extension) { $result[$rootItem.Name] = $rootItem }
        return $result
    }
    $rootPath = $rootItem.FullName.TrimEnd('\', '/')
    $excludeSet = @($Exclude | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($f in @(Get-ChildItem -LiteralPath $rootPath -File -Recurse -ErrorAction SilentlyContinue | Sort-Object FullName)) {
        if (-not (Get-SourceLanguage $f.Extension)) { continue }
        $rel = $f.FullName.Substring($rootPath.Length).TrimStart('\', '/')
        $segments = $rel -split '[\\/]'
        if (@($segments[0..($segments.Count - 2)] | Where-Object { $excludeSet -contains $_.ToLowerInvariant() }).Count -gt 0) { continue }
        $result[($rel -replace '\\', '/')] = $f
    }
    return $result
}

function Get-GroupName {
    param([string] $Relative)
    if ($GroupBy -eq 'File') { return $Relative }
    $segments = @($Relative -split '/')
    if ($segments.Count -le 1) { return '（直下）' }
    $n = [math]::Min($Depth, $segments.Count - 1)
    return ($segments[0..($n - 1)] -join '/')
}

function Compare-CodeLines {
    #     ステップの行の増減（同じ文字の行の数の差）から、新規・修正・削除の行数を出す。
    #     増えた行と減った行のうち、少ない方の数を修正とみなす。行の文字は大文字・小文字も区別して比べる（userId → userID も修正）
    param($Old, $New)
    $count = New-Object System.Collections.Hashtable ([System.StringComparer]::Ordinal)
    foreach ($l in $Old) { if ($count.ContainsKey($l)) { $count[$l]-- } else { $count[$l] = -1 } }
    foreach ($l in $New) { if ($count.ContainsKey($l)) { $count[$l]++ } else { $count[$l] = 1 } }
    $added = 0; $removed = 0
    foreach ($v in $count.Values) { if ($v -gt 0) { $added += $v } elseif ($v -lt 0) { $removed -= $v } }
    $changed = [math]::Min($added, $removed)
    return [pscustomobject]@{ Added = $added - $changed; Changed = $changed; Deleted = $removed - $changed }
}

$files = Get-SourceFiles $Path
$baseFiles = [ordered]@{}
if ($Baseline) { $baseFiles = Get-SourceFiles $Baseline }

$fileResults = New-Object System.Collections.Generic.List[object]
foreach ($rel in $files.Keys) {
    $f = $files[$rel]
    $lang = Get-SourceLanguage $f.Extension
    $m = Measure-SourceLines $f.FullName $lang
    $diff = $null
    if ($Baseline) {
        $old = @()
        if ($baseFiles.Contains($rel)) { $old = (Measure-SourceLines $baseFiles[$rel].FullName $lang).Code }
        $diff = Compare-CodeLines $old $m.Code
    }
    $fileResults.Add([pscustomobject]@{
        File = $rel; Group = (Get-GroupName $rel); Language = $lang.Name; Lines = $m.Lines; Blank = $m.Blank; Comment = $m.Comment; Steps = $m.Steps
        Added = $(if ($diff) { $diff.Added } else { $null }); Changed = $(if ($diff) { $diff.Changed } else { $null }); Deleted = $(if ($diff) { $diff.Deleted } else { $null })
        Status = $(if (-not $Baseline) { '' } elseif ($baseFiles.Contains($rel)) { $(if ($diff.Added + $diff.Changed + $diff.Deleted -gt 0) { '修正' } else { '変更なし' }) } else { '新規' })
    })
}
# 前の版にだけあるファイル（削除）
if ($Baseline) {
    foreach ($rel in $baseFiles.Keys) {
        if ($files.Contains($rel)) { continue }
        $f = $baseFiles[$rel]
        $lang = Get-SourceLanguage $f.Extension
        $m = Measure-SourceLines $f.FullName $lang
        $fileResults.Add([pscustomobject]@{ File = $rel; Group = (Get-GroupName $rel); Language = $lang.Name; Lines = 0; Blank = 0; Comment = 0; Steps = 0; Added = 0; Changed = 0; Deleted = $m.Steps; Status = '削除' })
    }
}

$groups = New-Object System.Collections.Generic.List[object]
foreach ($g in @($fileResults | Group-Object Group | Sort-Object Name)) {
    $items = @($g.Group)
    $sum = { param($p) ($items | Measure-Object -Property $p -Sum).Sum }
    $steps = [int](& $sum 'Steps')
    $r = [pscustomobject]@{
        Group = $g.Name; Files = @($items | Where-Object { $_.Status -ne '削除' }).Count; Lines = [int](& $sum 'Lines'); Blank = [int](& $sum 'Blank'); Comment = [int](& $sum 'Comment')
        Steps = $steps; Kloc = [math]::Round($steps / 1000.0, 3)
        Added = $(if ($Baseline) { [int](& $sum 'Added') } else { $null }); Changed = $(if ($Baseline) { [int](& $sum 'Changed') } else { $null }); Deleted = $(if ($Baseline) { [int](& $sum 'Deleted') } else { $null })
    }
    $r.PSObject.TypeNames.Insert(0, 'CodeSize.Group')
    $groups.Add($r)
}
$groups

# ---------------------------------------------------------------- 出力


if ($OutCsv) {
    # Measure-TestQuality.ps1 の -SizeCsv にそのまま渡せる形（機能・KLOC の列）
    $rows = foreach ($r in $groups) {
        $o = [ordered]@{ '機能' = $r.Group; 'KLOC' = $r.Kloc; 'ステップ数' = $r.Steps; 'ファイル数' = $r.Files; '行数' = $r.Lines; 'コメント' = $r.Comment; '空行' = $r.Blank }
        if ($Baseline) { $o['新規'] = $r.Added; $o['修正'] = $r.Changed; $o['削除'] = $r.Deleted }
        [pscustomobject]$o
    }
    $csvText = (@($rows) | ConvertTo-Csv -NoTypeInformation) -join "`r`n"
    [System.IO.File]::WriteAllText($OutCsv, $csvText + "`r`n", (New-Object System.Text.UTF8Encoding($true)))
}

if ($OutXlsx) {
    $num = { param($v) if ($null -eq $v) { '' } else { [double]$v } }
    $groupRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $groups) { $groupRows.Add(@{ Values = @($r.Group, [double]$r.Files, [double]$r.Lines, [double]$r.Comment, [double]$r.Blank, [double]$r.Steps, [double]$r.Kloc, (& $num $r.Added), (& $num $r.Changed), (& $num $r.Deleted)) }) }
    $fileRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $fileResults) { $fileRows.Add(@{ Values = @($f.File, $f.Group, $f.Language, $f.Status, [double]$f.Lines, [double]$f.Comment, [double]$f.Blank, [double]$f.Steps, (& $num $f.Added), (& $num $f.Changed), (& $num $f.Deleted)) }) }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '機能別'; Columns = @(@{ Header = '機能'; Width = 30 }, @{ Header = 'ファイル数'; Width = 9 }, @{ Header = '行数'; Width = 9 }, @{ Header = 'コメント'; Width = 9 }, @{ Header = '空行'; Width = 8 }, @{ Header = 'ステップ数'; Width = 10 }, @{ Header = 'KLOC'; Width = 8 }, @{ Header = '新規'; Width = 8 }, @{ Header = '修正'; Width = 8 }, @{ Header = '削除'; Width = 8 }); Rows = $groupRows },
            @{ Name = 'ファイル別'; Columns = @(@{ Header = 'ファイル'; Width = 50 }, @{ Header = '機能'; Width = 20 }, @{ Header = '言語'; Width = 10 }, @{ Header = '状態'; Width = 8 }, @{ Header = '行数'; Width = 8 }, @{ Header = 'コメント'; Width = 9 }, @{ Header = '空行'; Width = 8 }, @{ Header = 'ステップ数'; Width = 10 }, @{ Header = '新規'; Width = 8 }, @{ Header = '修正'; Width = 8 }, @{ Header = '削除'; Width = 8 }); Rows = $fileRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$totalSteps = ($groups | Measure-Object -Property Steps -Sum).Sum
if ($null -eq $totalSteps) { $totalSteps = 0 }
Write-Host ('{0} ファイル: ステップ数 {1}（{2} KLOC）・コメント {3}・空行 {4}' -f @($fileResults | Where-Object { $_.Status -ne '削除' }).Count, $totalSteps, [math]::Round($totalSteps / 1000.0, 3), ($groups | Measure-Object -Property Comment -Sum).Sum, ($groups | Measure-Object -Property Blank -Sum).Sum)
if ($Baseline) {
    $statusText = (@($fileResults | Group-Object Status | Sort-Object Name | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count })) -join '・'
    Write-Host ('前の版との差: 新規 {0}・修正 {1}・削除 {2} ステップ（ファイル: {3}）' -f ($groups | Measure-Object -Property Added -Sum).Sum, ($groups | Measure-Object -Property Changed -Sum).Sum, ($groups | Measure-Object -Property Deleted -Sum).Sum, $statusText)
}
if ($OutCsv) { Write-Host "規模の CSV を保存しました: $OutCsv" }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     ソースコードのステップ数（コメント・空行を除いた行数）を数え、前の版との新規・修正・削除の行数を出します（ステップカウンター。Excel 不要）。
#
# .DESCRIPTION
#     フォルダの中のソースコードを言語ごとのコメントの書き方で読み、行数・コメント・空行・ステップ数を、機能（フォルダ）ごとに数えます。
#     バグ密度・テスト密度の分母（KLOC）に使えます。-OutCsv の CSV は Measure-TestQuality.ps1 の -SizeCsv にそのまま渡せます。
#
#     対応する言語: C 系（C・C++・C#・Java・JavaScript・TypeScript・Go・Kotlin・Swift・Scala・Rust・CSS など）・PHP・SQL・VB・
#     Python・Ruby・シェル・PowerShell・HTML / XML（JSP・Vue など）・COBOL（固定形式）
#
#     -Baseline に前の版のフォルダを指定すると、同じ相対パスのファイルどうしで、ステップの行の増減から新規・修正・削除を数えます。
#     増えた行と減った行のうち少ない方を修正とみなします（行の並べ替えは数えません）。
#
# .PARAMETER Path
#     数えるソースコードのフォルダ（またはファイル）。
#
# .PARAMETER Baseline
#     前の版のソースコードのフォルダ。指定すると新規・修正・削除の行数を出します。
#
# .PARAMETER GroupBy
#     まとめる単位。Folder（既定。-Depth 階層目までのフォルダ）か File。
#
# .PARAMETER Depth
#     Folder でまとめるときのフォルダの階層（既定 1。-Path の直下のフォルダごと）。
#
# .PARAMETER Exclude
#     数えないフォルダの名前（既定は .git・node_modules・bin・obj・target・build・dist など）。
#
# .PARAMETER OutCsv
#     機能ごとの規模を CSV（UTF-8・BOM 付き）に保存します。列は 機能・KLOC・ステップ数・ファイル数・行数・コメント・空行（・新規・修正・削除）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（機能別・ファイル別）。
#
# .EXAMPLE
#     .\Measure-CodeSize.ps1 D:\src -OutCsv D:\規模.csv
#     .\Measure-TestQuality.ps1 D:\テスト -BugList D:\障害管理表.xlsx -SizeCsv D:\規模.csv
#
# .EXAMPLE
#     # 前の版との差（新規・修正・削除）
#     .\Measure-CodeSize.ps1 D:\src_v1.1 -Baseline D:\src_v1.0 -OutXlsx D:\ステップ数.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1 が必要です。
#     文字列（"…"・'…' など。HTML・XML は除く）の中と、行コメントの後ろの「/*」などはコメントの始まりにしません。
#     複数の行にわたる文字列（Java のテキストブロックなど）の中の「/*」などは、コメントの始まりとして数えることがあります（行単位で判定する目安です）。
