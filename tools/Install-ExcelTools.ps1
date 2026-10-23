[CmdletBinding()]
param(
    [string] $Source,

    [string] $InstallDir,

    [string[]] $Action,

    [string] $OutDir,

    [string] $SendToDir,

    [switch] $NoSendTo,

    [string] $ShortcutDir,

    [switch] $Desktop,

    [switch] $Update,

    [switch] $CheckOnly,

    [int] $CheckHours = 12,

    [switch] $Uninstall
)

$ErrorActionPreference = 'Stop'

$isWin = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows
if (-not $InstallDir) {
    $base = $env:LOCALAPPDATA
    if (-not $base) { $base = Join-Path $HOME '.local/share' }
    $InstallDir = Join-Path $base 'ExcelTools'
}
$InstallDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($InstallDir)
$toolsDir = Join-Path $InstallDir 'tools'
$infoPath = Join-Path $InstallDir 'install.json'

# ---------------------------------------------------------------- ヘルパー


function Get-ToolFiles {
    #     tools フォルダのファイル（相対パス → FileInfo）。Office の一時ファイルは除く
    param([string] $Dir)
    $files = [ordered]@{}
    $root = (Get-Item -LiteralPath $Dir).FullName.TrimEnd('\', '/')
    foreach ($f in @(Get-ChildItem -LiteralPath $root -Recurse -File | Sort-Object FullName)) {
        if ($f.Name.StartsWith('~$')) { continue }
        $files[$f.FullName.Substring($root.Length + 1)] = $f
    }
    return $files
}

function Get-ToolsFingerprint {
    #     フォルダの中身の目印（ファイルの相対パス・大きさ・更新日時から作る）。変わっていれば更新がある
    param([string] $Dir)
    $sb = New-Object System.Text.StringBuilder
    $files = Get-ToolFiles $Dir
    foreach ($rel in $files.Keys) { [void]$sb.Append($rel.Replace('\', '/')).Append('|').Append($files[$rel].Length).Append('|').Append($files[$rel].LastWriteTimeUtc.Ticks).Append("`n") }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($sb.ToString())) } finally { $sha.Dispose() }
    return (-join ($hash | Select-Object -First 8 | ForEach-Object { $_.ToString('x2') }))
}

function Test-ToolsFolder {
    param([string] $Dir)
    return (Test-Path -LiteralPath (Join-Path $Dir 'ExcelTools.psm1')) -and (Test-Path -LiteralPath (Join-Path $Dir 'Invoke-SendTo.ps1'))
}

function Read-InstallInfo {
    if (-not (Test-Path -LiteralPath $infoPath)) { return $null }
    try { return (Get-Content -LiteralPath $infoPath -Raw -Encoding UTF8 | ConvertFrom-Json) } catch { return $null }
}

function Save-InstallInfo {
    #     配布元・中身の目印・最後に確かめた日時と、「送る」の登録の指定（更新のときに同じ指定で登録し直す）
    param([string] $From, [string] $Fingerprint, [datetime] $Checked)
    $info = [ordered]@{
        Source = $From; Fingerprint = $Fingerprint; Installed = (Get-Date).ToString('yyyy/MM/dd HH:mm:ss'); LastCheck = $Checked.ToString('yyyy-MM-ddTHH:mm:ss')
        Action = @($Action); OutDir = $OutDir; SendToDir = $SendToDir; NoSendTo = [bool]$NoSendTo; ShortcutDir = $ShortcutDir; Desktop = [bool]$Desktop
    }
    [System.IO.File]::WriteAllText($infoPath, ($info | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
}

function Sync-ToolsFolder {
    #     配布元の tools フォルダを、この PC の tools フォルダにそろえる（変わったファイルだけを写し、配布元に無いファイルは消す）
    param([string] $From, [string] $To)
    $src = Get-ToolFiles $From
    if (-not (Test-Path -LiteralPath $To)) { New-Item -ItemType Directory -Force -Path $To | Out-Null }
    $dst = Get-ToolFiles $To
    $copied = 0; $removed = 0
    foreach ($rel in $src.Keys) {
        $s = $src[$rel]
        if ($dst.Contains($rel) -and $dst[$rel].Length -eq $s.Length -and $dst[$rel].LastWriteTimeUtc -eq $s.LastWriteTimeUtc) { continue }
        $target = Join-Path $To $rel
        $dir = Split-Path -Parent $target
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        Copy-Item -LiteralPath $s.FullName -Destination $target -Force
        # 次の比較のため、更新日時を配布元とそろえる（Copy-Item が日時を写さない環境がある）
        (Get-Item -LiteralPath $target).LastWriteTimeUtc = $s.LastWriteTimeUtc
        # 共有フォルダやダウンロードから写したファイルの「ブロック」を外す（実行のたびに確認が出ないように）
        if ($isWin -and (Get-Command Unblock-File -ErrorAction SilentlyContinue)) { Unblock-File -LiteralPath $target -ErrorAction SilentlyContinue }
        $copied++
    }
    foreach ($rel in $dst.Keys) {
        if (-not $src.Contains($rel)) { Remove-Item -LiteralPath $dst[$rel].FullName -Force; $removed++ }
    }
    # 空になったフォルダを消す
    foreach ($d in @(Get-ChildItem -LiteralPath $To -Recurse -Directory | Sort-Object { $_.FullName.Length } -Descending)) {
        if (@(Get-ChildItem -LiteralPath $d.FullName -Force).Count -eq 0) { Remove-Item -LiteralPath $d.FullName -Force }
    }
    return [pscustomobject]@{ Copied = $copied; Removed = $removed; Total = $src.Count }
}

function Register-SendTo {
    #     この PC の tools フォルダの Install-SendTo.ps1 で「送る」に登録する
    if ($NoSendTo) { return }
    if (-not $isWin) { Write-Warning '「送る」メニューは Windows でのみ登録できます（-NoSendTo で省けます）'; return }
    $sendArgs = @{}
    if ($Action) { $sendArgs['Action'] = $Action }
    if ($OutDir) { $sendArgs['OutDir'] = $OutDir }
    if ($SendToDir) { $sendArgs['SendToDir'] = $SendToDir }
    $null = & (Join-Path $toolsDir 'Install-SendTo.ps1') @sendArgs 6>$null
}

function Get-LauncherDirs {
    #     画面（Start-ExcelTools.ps1）のショートカットを置くフォルダ（既定はスタートメニュー。-Desktop ならデスクトップにも）
    $dirs = @()
    if ($ShortcutDir) { $dirs += $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ShortcutDir) }
    else { $dirs += [Environment]::GetFolderPath('Programs') }
    if ($Desktop) { $dirs += [Environment]::GetFolderPath('Desktop') }
    return @($dirs | Where-Object { $_ })
}

function Register-Launcher {
    #     スタートメニュー（とデスクトップ）に、画面を開くショートカット「Excel ツール」を作る
    if ($NoSendTo -or -not $isWin) { return }
    $gui = Join-Path $toolsDir 'Start-ExcelTools.ps1'
    if (-not (Test-Path -LiteralPath $gui)) { return }
    Import-Module (Join-Path $toolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Add-ShellLinkType
    $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    foreach ($dir in Get-LauncherDirs) {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $linkArgs = '-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}"' -f $gui
        [ExcelTools.ShellLink]::Save((Join-Path $dir 'Excel ツール.lnk'), $exe, $linkArgs, $toolsDir, 'Excel ツールの画面を開きます（やりたいことを選んで、ファイルを入れて実行）')
    }
}

function Unregister-Launcher {
    #     画面のショートカットを消す（このツールのものだけ）
    if (-not $isWin) { return }
    Import-Module (Join-Path $toolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Add-ShellLinkType
    foreach ($dir in Get-LauncherDirs) {
        $link = Join-Path $dir 'Excel ツール.lnk'
        if (-not (Test-Path -LiteralPath $link)) { continue }
        try { $info = [ExcelTools.ShellLink]::Read($link) } catch { continue }
        if ($info.Arguments -match 'Start-ExcelTools\.ps1') { Remove-Item -LiteralPath $link -Force }
    }
}

# ---------------------------------------------------------------- 入れたときの指定


# 指定が無ければ、入れたときの「送る」の登録の指定を使う（更新・削除も同じ指定で行う）
$info = Read-InstallInfo
if ($info) {
    if (-not $PSBoundParameters.ContainsKey('Action') -and $info.Action) { $Action = @($info.Action | Where-Object { $_ }) }
    if (-not $PSBoundParameters.ContainsKey('OutDir') -and $info.OutDir) { $OutDir = $info.OutDir }
    if (-not $PSBoundParameters.ContainsKey('SendToDir') -and $info.SendToDir) { $SendToDir = $info.SendToDir }
    if (-not $PSBoundParameters.ContainsKey('NoSendTo') -and $info.NoSendTo) { $NoSendTo = [switch]$true }
    if (-not $PSBoundParameters.ContainsKey('ShortcutDir') -and $info.ShortcutDir) { $ShortcutDir = $info.ShortcutDir }
    if (-not $PSBoundParameters.ContainsKey('Desktop') -and $info.Desktop) { $Desktop = [switch]$true }
}

# ---------------------------------------------------------------- 削除


if ($Uninstall) {
    if (-not $info) { throw "Install-ExcelTools.ps1 で入れたツールが見つかりません: $InstallDir" }
    if ($isWin -and -not $NoSendTo -and (Test-Path -LiteralPath (Join-Path $toolsDir 'Install-SendTo.ps1'))) {
        $sendArgs = @{ Uninstall = $true }
        if ($SendToDir) { $sendArgs['SendToDir'] = $SendToDir }
        $null = & (Join-Path $toolsDir 'Install-SendTo.ps1') @sendArgs 6>$null
    }
    if (-not $NoSendTo -and (Test-Path -LiteralPath (Join-Path $toolsDir 'ExcelTools.psm1'))) { Unregister-Launcher }
    Remove-Item -LiteralPath $InstallDir -Recurse -Force
    Write-Host "Excel ツールを削除しました（「送る」の項目とスタートメニューの「Excel ツール」も消しました）: $InstallDir"
    [pscustomobject]@{ Result = '削除'; InstallDir = $InstallDir; Source = $null; Copied = 0; Removed = 0 }
    return
}

# ---------------------------------------------------------------- 更新の確認（Invoke-SendTo.ps1 が実行の後に呼ぶ）


if ($Update -or $CheckOnly) {
    if (-not $info) { throw "Install-ExcelTools.ps1 で入れたツールが見つかりません: $InstallDir" }
    if (-not $Source) { $Source = $info.Source }
    # 確認は $CheckHours 時間に 1 回まで（配布元が遠いと時間がかかるため）。PowerShell 7 の ConvertFrom-Json は日時の文字を日時にする
    $last = [datetime]::MinValue
    if ($info.LastCheck -is [datetime]) { $last = $info.LastCheck }
    elseif ($info.LastCheck) { try { $last = [datetime]::ParseExact([string]$info.LastCheck, 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } catch { } }
    if ($CheckOnly -and $last.AddHours($CheckHours) -gt (Get-Date)) {
        [pscustomobject]@{ Result = '確認済み'; InstallDir = $InstallDir; Source = $Source; Copied = 0; Removed = 0 }
        return
    }
    if (-not (Test-Path -LiteralPath $Source) -or -not (Test-ToolsFolder $Source)) {
        if ($CheckOnly) {
            Save-InstallInfo $info.Source $info.Fingerprint (Get-Date)
            [pscustomobject]@{ Result = '配布元に届かない'; InstallDir = $InstallDir; Source = $Source; Copied = 0; Removed = 0 }
            return
        }
        throw "配布元の tools フォルダが見つかりません: $Source"
    }
    $fp = Get-ToolsFingerprint $Source
    if ($fp -eq $info.Fingerprint) {
        Save-InstallInfo $info.Source $fp (Get-Date)
        [pscustomobject]@{ Result = '最新'; InstallDir = $InstallDir; Source = $Source; Copied = 0; Removed = 0 }
        return
    }
    $r = Sync-ToolsFolder $Source $toolsDir
    Save-InstallInfo $Source $fp (Get-Date)
    Register-SendTo
    Register-Launcher
    Write-Host ("Excel ツールを新しい版に更新しました（{0} ファイル。削除 {1}）: {2}" -f $r.Copied, $r.Removed, $Source)
    [pscustomobject]@{ Result = '更新'; InstallDir = $InstallDir; Source = $Source; Copied = $r.Copied; Removed = $r.Removed }
    return
}

# ---------------------------------------------------------------- 入れる


if (-not $Source) {
    $Source = $PSScriptRoot
    # この PC に入れた tools フォルダから実行したときは、覚えている配布元から入れ直す
    if ($info -and [string]::Equals($Source.TrimEnd('\', '/'), $toolsDir.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) { $Source = $info.Source }
}
$Source = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Source).TrimEnd('\', '/')
if (-not (Test-ToolsFolder $Source)) { throw "配布する tools フォルダではありません（ExcelTools.psm1 と Invoke-SendTo.ps1 が要ります）: $Source" }
if ([string]::Equals($Source, $toolsDir.TrimEnd('\', '/'), [StringComparison]::OrdinalIgnoreCase)) { throw "配布元と入れる先が同じです。配布元（共有フォルダ）の Install-ExcelTools.cmd から実行してください: $Source" }

New-Item -ItemType Directory -Force -Path $InstallDir | Out-Null
$r = Sync-ToolsFolder $Source $toolsDir
Save-InstallInfo $Source (Get-ToolsFingerprint $Source) (Get-Date)
Register-SendTo
Register-Launcher
$result = '入れた'
if ($info) { $result = '入れ直した' }
Write-Host ("Excel ツールを{0}: {1}（{2} ファイル中 {3} ファイルを写しました）" -f @{ '入れた' = '入れました'; '入れ直した' = '入れ直しました' }[$result], $toolsDir, $r.Total, $r.Copied)
if (-not $NoSendTo -and $isWin) {
    Write-Host 'スタートメニューの「Excel ツール」から、ツールの画面を開けます（やりたいことを選んで、ファイルを入れて実行）。'
    Write-Host 'エクスプローラーでファイルを右クリック →「送る」→「Excel …」からも使えます。'
}
Write-Host ("配布元: {0}（ここが更新されると、「送る」で使ったあとに自動で新しい版にします）" -f $Source)
[pscustomobject]@{ Result = $result; InstallDir = $InstallDir; Source = $Source; Copied = $r.Copied; Removed = $r.Removed }

# .SYNOPSIS
#     共有フォルダなどに置いた Excel ツールを、この PC に入れて「送る」に登録します。更新・削除もします（管理者の権限は不要）。
#
# .DESCRIPTION
#     管理者が tools フォルダを共有フォルダ（配布元）に置き、担当者は配布元の Install-ExcelTools.cmd をダブルクリックするだけで使い始められます。
#
#       ・配布元の tools フォルダを、この PC の「%LOCALAPPDATA%\ExcelTools\tools」に写し、右クリックの「送る」に各ツールを登録します
#         （共有フォルダから直接動かすより速く、ネットワークが切れても使えます）
#       ・スタートメニューに、ツールの画面（Start-ExcelTools.ps1）を開く「Excel ツール」を作ります（-Desktop でデスクトップにも）
#       ・配布元が更新されると、「送る」でツールを使ったあとに自動で新しい版を写します（確認は 12 時間に 1 回まで。次の実行から新しい版になります）
#       ・もう一度実行すると、すぐに入れ直します。-Uninstall で「送る」の項目・スタートメニューの「Excel ツール」とこの PC の写しを消します
#
# .PARAMETER Source
#     配布元の tools フォルダ。既定はこのスクリプトのあるフォルダ（この PC に入れた写しから実行したときは、覚えている配布元）。
#
# .PARAMETER InstallDir
#     この PC に入れるフォルダ。既定は「%LOCALAPPDATA%\ExcelTools」（tools フォルダと install.json を置きます）。
#
# .PARAMETER Action
#     「送る」に登録する処理（Install-SendTo.ps1 の -Action）。既定はすべて。
#
# .PARAMETER OutDir
#     「送る」の結果を保存するフォルダ（Install-SendTo.ps1 の -OutDir）。既定は「ドキュメント\Excel ツールの結果」。
#
# .PARAMETER SendToDir
#     ショートカットを作るフォルダ（Install-SendTo.ps1 の -SendToDir）。既定は、ログインしている人の「送る」のフォルダ。
#
# .PARAMETER NoSendTo
#     「送る」とスタートメニューに登録しません（ツールを写すだけ）。
#
# .PARAMETER ShortcutDir
#     ツールの画面を開くショートカット「Excel ツール」を作るフォルダ。既定はスタートメニュー（プログラム）。
#
# .PARAMETER Desktop
#     ツールの画面を開くショートカットを、デスクトップにも作ります。
#
# .PARAMETER Update
#     配布元が変わっていれば、変わったファイルだけを写します（変わっていなければ何もしません）。
#
# .PARAMETER CheckOnly
#     -Update と同じですが、前の確認から -CheckHours 時間たっていなければ配布元を見ません（Invoke-SendTo.ps1 が実行のあとに呼びます）。
#
# .PARAMETER CheckHours
#     -CheckOnly で配布元を見る間隔（時間）。既定は 12。
#
# .PARAMETER Uninstall
#     「送る」の項目と、この PC に入れたツールを消します。
#
# .EXAMPLE
#     # 担当者: 配布元の Install-ExcelTools.cmd をダブルクリックする（PowerShell からなら次のとおり）
#     \\server\share\ExcelTools\tools\Install-ExcelTools.ps1
#
# .EXAMPLE
#     # 提出前チェックと差分だけを「送る」に登録し、結果を共有フォルダに保存する
#     \\server\share\ExcelTools\tools\Install-ExcelTools.ps1 -Action Check, Compare -OutDir \\server\share\チェック結果
#
# .EXAMPLE
#     # すぐに新しい版にする / 消す
#     & "$env:LOCALAPPDATA\ExcelTools\tools\Install-ExcelTools.ps1" -Update
#     & "$env:LOCALAPPDATA\ExcelTools\tools\Install-ExcelTools.ps1" -Uninstall
#
# .NOTES
#     要件: PowerShell 5.1 以降。「送る」の登録は Windows のみ。Excel は不要です。管理者の権限は要りません。
