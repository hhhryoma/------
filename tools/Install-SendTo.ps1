[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string[]] $Action,

    [string] $SendToDir,

    [string] $Prefix = 'Excel ',

    [string] $OutDir,

    [ValidateSet('powershell', 'pwsh')]
    [string] $Shell = 'powershell',

    [switch] $Uninstall,

    [switch] $Force
)

$ErrorActionPreference = 'Stop'

if (-not (($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows)) { throw '「送る」メニューは Windows でのみ使えます' }

$launcher = Join-Path $PSScriptRoot 'Invoke-SendTo.ps1'
$all = @(& $launcher -Action List)
$selected = $all
if ($Action) {
    $known = @($all | ForEach-Object { $_.Action })
    $unknown = @($Action | Where-Object { $known -notcontains $_ })
    if ($unknown.Count -gt 0) { throw ('登録できない -Action です: {0}（使えるもの: {1}）' -f ($unknown -join ', '), ($known -join ', ')) }
    $selected = @($all | Where-Object { $Action -contains $_.Action })
}

if (-not $SendToDir) { $SendToDir = [Environment]::GetFolderPath('SendTo') }
$SendToDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SendToDir)
if (-not (Test-Path -LiteralPath $SendToDir -PathType Container)) { throw "「送る」のフォルダが見つかりません: $SendToDir" }

if ($Shell -eq 'pwsh') {
    $cmd = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if (-not $cmd) { throw 'pwsh.exe（PowerShell 7）が見つかりません' }
    $exe = $cmd.Source
}
else {
    $exe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
}

# ショートカット（.lnk）の読み書き。WScript.Shell はファイル名を ANSI のコードページで扱うため、
# 日本語の名前は日本語版以外の Windows で保存できない。Unicode の IShellLinkW を直接使う
if (-not ('ExcelTools.ShellLink' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace ExcelTools {
    [ComImport, Guid("000214F9-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellLinkW {
        void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cch, IntPtr pfd, uint fFlags);
        void GetIDList(out IntPtr ppidl);
        void SetIDList(IntPtr pidl);
        void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cch);
        void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cch);
        void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
        void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cch);
        void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
        void GetHotkey(out short pwHotkey);
        void SetHotkey(short wHotkey);
        void GetShowCmd(out int piShowCmd);
        void SetShowCmd(int iShowCmd);
        void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cch, out int piIcon);
        void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath, int iIcon);
        void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, uint dwReserved);
        void Resolve(IntPtr hwnd, uint fFlags);
        void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
    }

    [ComImport, Guid("00021401-0000-0000-C000-000000000046")]
    class CShellLink { }

    public class ShellLinkInfo {
        public string TargetPath;
        public string Arguments;
        public string WorkingDirectory;
        public string Description;
    }

    public static class ShellLink {
        public static void Save(string path, string targetPath, string arguments, string workingDirectory, string description) {
            IShellLinkW link = (IShellLinkW)new CShellLink();
            try {
                link.SetPath(targetPath);
                link.SetArguments(arguments);
                link.SetWorkingDirectory(workingDirectory);
                link.SetDescription(description);
                ((IPersistFile)link).Save(path, true);
            }
            finally { Marshal.FinalReleaseComObject(link); }
        }

        public static ShellLinkInfo Read(string path) {
            IShellLinkW link = (IShellLinkW)new CShellLink();
            try {
                ((IPersistFile)link).Load(path, 0);
                ShellLinkInfo info = new ShellLinkInfo();
                StringBuilder sb = new StringBuilder(32768);
                link.GetPath(sb, sb.Capacity, IntPtr.Zero, 0);
                info.TargetPath = sb.ToString();
                sb.Length = 0;
                link.GetArguments(sb, sb.Capacity);
                info.Arguments = sb.ToString();
                sb.Length = 0;
                link.GetWorkingDirectory(sb, sb.Capacity);
                info.WorkingDirectory = sb.ToString();
                sb.Length = 0;
                link.GetDescription(sb, sb.Capacity);
                info.Description = sb.ToString();
                return info;
            }
            finally { Marshal.FinalReleaseComObject(link); }
        }
    }
}
'@
}

$results = New-Object System.Collections.Generic.List[object]

function Get-LauncherAction {
    #     このツールのショートカットなら、その -Action を返す（ほかのショートカット・読めないファイルは $null）
    param([string] $LinkPath)
    try { $link = [ExcelTools.ShellLink]::Read($LinkPath) } catch { return $null }
    $m = [regex]::Match($link.Arguments, 'Invoke-SendTo\.ps1"?\s+-Action\s+(\w+)')
    if ($m.Success) { return $m.Groups[1].Value }
    return $null
}

if ($Uninstall) {
    foreach ($f in @(Get-ChildItem -LiteralPath $SendToDir -Filter '*.lnk' -File)) {
        $act = Get-LauncherAction $f.FullName
        if (-not $act) { continue }
        if ($Action -and $Action -notcontains $act) { continue }
        $status = '削除'
        if ($PSCmdlet.ShouldProcess($f.Name, 'ショートカットを削除')) { Remove-Item -LiteralPath $f.FullName -Force }
        else { $status = '削除予定' }
        $results.Add([pscustomobject]@{ Name = $f.BaseName; Status = $status; Detail = ''; Path = $f.FullName })
    }
}
else {
    # 引数の最後が \" だと引用符の終わりと解釈されないため、末尾の \ は外す
    $argText = ''
    if ($OutDir) { $argText = ' -OutDir "{0}"' -f $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir).TrimEnd('\') }
    foreach ($a in $selected) {
        $name = $Prefix + $a.Label
        $linkPath = Join-Path $SendToDir ($name + '.lnk')
        $status = '作成'
        if (Test-Path -LiteralPath $linkPath) {
            if (-not (Get-LauncherAction $linkPath) -and -not $Force) {
                $results.Add([pscustomobject]@{ Name = $name; Status = 'スキップ'; Detail = '同じ名前のほかのショートカットがあります（-Force で置き換えます）'; Path = $linkPath })
                continue
            }
            $status = '更新'
        }
        if ($PSCmdlet.ShouldProcess($name, "ショートカットを$status")) {
            $linkArgs = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -Action {1}{2}' -f $launcher, $a.Action, $argText
            [ExcelTools.ShellLink]::Save($linkPath, $exe, $linkArgs, $PSScriptRoot, $a.Detail)
        }
        else { $status += '予定' }
        $results.Add([pscustomobject]@{ Name = $name; Status = $status; Detail = $a.Detail; Path = $linkPath })
    }
}

$results
Write-Host ''
if ($results.Count -eq 0) { Write-Host 'ショートカットはありませんでした。' }
else { Write-Host ('{0}: {1}' -f $SendToDir, (($results | Group-Object Status | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join ' / ')) }

# .SYNOPSIS
#     エクスプローラーの右クリック →「送る」メニューに、各ツールを登録します（Excel 不要）。
#
# .DESCRIPTION
#     「送る」のフォルダに、Invoke-SendTo.ps1 を呼ぶショートカット（「Excel 提出前チェック」など）を作ります。
#     ファイルやフォルダを選んで右クリック →「送る」→ 項目を選ぶと、そのツールを実行し、結果の Excel を開きます。
#     PowerShell に慣れていない人にも、この 1 回の登録だけで使ってもらえます。
#
#     ショートカットは tools フォルダの場所を覚えるため、tools フォルダを動かしたら、もう一度実行してください。
#     同じ名前のほかのショートカットは置き換えません（-Force で置き換えます）。
#
# .PARAMETER Action
#     登録する処理（既定はすべて）。Check・ResetView・Compare・SpVersion・Search・Terms・Consistency・TestDraft・Size・Estimate・Pdf。
#     一覧は .\Invoke-SendTo.ps1 -Action List で確かめられます。
#
# .PARAMETER SendToDir
#     ショートカットを作るフォルダ。既定は、ログインしている人の「送る」のフォルダ（shell:sendto）。
#
# .PARAMETER Prefix
#     ショートカットの名前の頭に付ける文字。既定は「Excel 」（「送る」メニューで並んで表示されるように）。
#
# .PARAMETER OutDir
#     結果のファイルを保存するフォルダ。既定は「ドキュメント\Excel ツールの結果」。
#
# .PARAMETER Shell
#     ツールを動かす PowerShell。powershell（Windows PowerShell 5.1、既定）または pwsh（PowerShell 7）。
#
# .PARAMETER Uninstall
#     このツールのショートカットを「送る」から削除します（-Action を付けると、その処理のものだけ）。
#
# .PARAMETER Force
#     同じ名前のほかのショートカットがあっても置き換えます。
#
# .EXAMPLE
#     # すべての処理を「送る」に登録する
#     .\Install-SendTo.ps1
#
# .EXAMPLE
#     # 提出前チェックと差分だけを登録し、結果を共有フォルダに保存する
#     .\Install-SendTo.ps1 -Action Check, Compare -OutDir \\server\share\チェック結果
#
# .EXAMPLE
#     # 登録したショートカットを削除する
#     .\Install-SendTo.ps1 -Uninstall
#
# .NOTES
#     要件: Windows、PowerShell 5.1 以降。Excel は不要です。管理者の権限は要りません（自分の「送る」だけに登録します）。
#     ショートカットは -ExecutionPolicy Bypass で PowerShell を起動するため、実行ポリシーの設定を変える必要はありません。
