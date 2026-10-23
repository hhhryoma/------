[CmdletBinding()]
param(
    # 「送る」は選んだファイルを引数の最後に並べるため、名前の無い引数はすべてここに入れる
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]] $Path,

    [string] $Tool,

    [string] $SettingsPath
)

$ErrorActionPreference = 'Stop'

if (-not (($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows)) { throw 'Excel ツールの画面は Windows でのみ使えます（ほかの OS では、各ツールを PowerShell から実行してください）' }

# ドラッグ＆ドロップとファイルを選ぶ画面には STA が要る。MTA で起動されたら、STA で起動し直す
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA) {
    $exe = (Get-Process -Id $PID).Path
    $relaunch = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
    if ($Tool) { $relaunch += @('-Tool', ('"{0}"' -f $Tool)) }
    if ($SettingsPath) { $relaunch += @('-SettingsPath', ('"{0}"' -f $SettingsPath.TrimEnd('\'))) }
    foreach ($p in @($Path)) { if ($p) { $relaunch += ('"{0}"' -f $p.TrimEnd('\')) } }
    Start-Process -FilePath $exe -ArgumentList $relaunch -WindowStyle Hidden
    return
}

Import-Module (Join-Path $PSScriptRoot 'ExcelToolsGui.psm1') -Force -DisableNameChecking
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# 「送る」から渡されたパスを直す（ドライブの直下 "D:\" は D:" になって届く）
$paths = @($Path | Where-Object { $_ -and $_.Trim() } | ForEach-Object {
        $v = $_.Trim().Trim('"')
        if ($v -match '^[A-Za-z]:$') { $v += '\' }
        $v
    })

$windowArgs = @{ Path = $paths }
if ($Tool) { $windowArgs['Tool'] = $Tool }
if ($SettingsPath) { $windowArgs['SettingsPath'] = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SettingsPath) }
$form = New-GuiWindow @windowArgs
try { [System.Windows.Forms.Application]::Run($form) }
finally { $form.Dispose() }

# Install-ExcelTools.ps1 で入れたツールなら、閉じたあとに配布元が更新されていないかを確かめる（12 時間に 1 回まで。次に開いたときから新しい版になる）
$installDir = Split-Path -Parent $PSScriptRoot
if (Test-Path -LiteralPath (Join-Path $installDir 'install.json')) {
    try { $null = & (Join-Path $PSScriptRoot 'Install-ExcelTools.ps1') -InstallDir $installDir -CheckOnly 6>$null } catch { }
}

# .SYNOPSIS
#     Excel ツールを画面（ウィンドウ）から使います。やりたいことを選び、ファイルやフォルダを入れて「実行」を押すだけで、PowerShell を使わずに各ツールを実行できます。
#
# .DESCRIPTION
#     左の一覧から、やりたいこと（提出前チェック・差分・検索など）を選ぶと、右に入力欄が出ます。
#
#       ・ファイルやフォルダは、「ファイル…」「フォルダ…」のボタンで選ぶか、エクスプローラーから画面にドラッグ＆ドロップします
#       ・結果の保存先を空欄にすると、「ドキュメント\Excel ツールの結果」に「名前_日時.xlsx」で保存し、終わったら開きます
#       ・ファイルを書き換える処理（置換・表示を整えるなど）は、最初は「試しに実行」で、書き換える箇所を確かめられます
#       ・「詳細な設定も表示する」で、各ツールのすべての引数を入力できます（説明は各欄の下とマウスを重ねたときに出ます）
#       ・実行中の様子（ツールの出力）は下の記録の欄に出ます。「中止」で止められます
#       ・入力は保存され、次に開いたときに出ます（%APPDATA%\ExcelTools\画面の設定.json）
#       ・「コマンドをコピー」で、同じことをする PowerShell のコマンドを写せます
#
#     並べる処理と、最初から見せる入力欄は tools\gui\ツールの一覧.psd1 に書きます。入力欄は各ツールの引数とヘルプから作ります。
#
# .PARAMETER Path
#     最初のパスの欄に入れるファイル・フォルダ（「送る」から渡されます）。
#
# .PARAMETER Tool
#     最初に選ぶ処理（ツールの一覧の Id。Check・Compare・Search など）。省略時は前に使った処理。
#
# .PARAMETER SettingsPath
#     入力を保存するファイル。既定は「%APPDATA%\ExcelTools\画面の設定.json」。
#
# .EXAMPLE
#     .\Start-ExcelTools.ps1
#
# .EXAMPLE
#     # 提出前チェックを選んだ状態で、フォルダを入れて開く
#     .\Start-ExcelTools.ps1 -Tool Check D:\提出物
#
# .NOTES
#     要件: Windows の PowerShell 5.1 以降（PowerShell 7 でも動作）。同じフォルダに ExcelToolsGui.psm1・各ツールと gui フォルダが必要です。
#     エクスプローラーからは、同じフォルダの Start-ExcelTools.cmd をダブルクリックして開けます。
#     Install-ExcelTools.ps1 で入れると、スタートメニューの「Excel ツール」と、「送る」→「Excel ツールの画面で開く」からも開けます。
