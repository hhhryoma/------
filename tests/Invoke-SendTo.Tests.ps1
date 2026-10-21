#     Invoke-SendTo.ps1・Install-SendTo.ps1 のテスト（右クリックの「送る」から使う）

$SendTo = Join-Path $script:ToolsDir 'Invoke-SendTo.ps1'
$InstallSendTo = Join-Path $script:ToolsDir 'Install-SendTo.ps1'

function Invoke-SendTo {
    param([hashtable] $Options, [string[]] $Paths)
    $null = & $SendTo @Options -NoOpen -NoPause @Paths 6>$null
}

function Get-ResultFiles {
    param([string] $Dir)
    return , @(Get-ChildItem -LiteralPath $Dir -Filter '*.xlsx' -File | ForEach-Object { $_.FullName })
}

function Get-SheetValues {
    #     結果のブックの 1 シート目の値を、行ごとに | でつないだ文字列の配列にする
    param([string] $Path)
    $sheet = Read-XlsxSheet $Path
    return , @($sheet.Rows.Keys | Sort-Object | ForEach-Object {
            $row = $sheet.Rows[$_]
            (@($row.Keys | Sort-Object { $_.Length }, { $_ } | ForEach-Object { $row[$_] }) -join '|')
        })
}

function New-VersionBook {
    param([string] $Path, [string] $Name)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    Export-Workbook -FilePath $Path -Sheets @(@{
        Name = 'Data'
        Columns = @(@{ Header = 'ID'; Width = 10 }, @{ Header = '名前'; Width = 20 })
        Rows = @(@{ Values = @('1', $Name) }, @{ Values = @('2', '顧客') })
    })
}

Test-Case 'Invoke-SendTo: -Action List で「送る」に登録する処理の一覧を返す' {
    $list = @(& $SendTo -Action List)
    Assert-Equal @('Check', 'ResetView', 'Compare', 'Search', 'Terms', 'Consistency', 'Size', 'Estimate', 'Pdf') @($list | ForEach-Object { $_.Action })
    Assert-True (@($list | Where-Object { -not $_.Label -or -not $_.Detail }).Count -eq 0) '名前・説明のない処理'
}

Test-Case 'Invoke-SendTo: 提出前チェックの結果を保存する（「送る」から届く引用符の残ったパスも読む）' {
    $dir = Join-Path (New-TestDir) '提出物 [最終]'
    New-Item -ItemType Directory -Path $dir | Out-Null
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/good.xlsx') -Destination $dir
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/bad.xlsx') -Destination $dir
    $outDir = Join-Path (New-TestDir) '結果'
    Invoke-SendTo @{ Action = 'Check'; OutDir = $outDir } @(($dir + '"'))
    $files = Get-ResultFiles $outDir
    Assert-Equal 1 $files.Count
    Assert-Match (Split-Path $files[0] -Leaf) '^提出前チェック_\d{8}_\d{6}\.xlsx$'
    Assert-True (@((Get-SheetValues $files[0]) | Where-Object { $_ -like 'bad.xlsx|*' }).Count -gt 0) 'bad.xlsx の指摘がない'
    Assert-True (@((Get-SheetValues $files[0]) | Where-Object { $_ -like 'good.xlsx|*' }).Count -eq 0) 'good.xlsx に指摘がある'

    # 同じ秒にもう一度実行しても、前の結果を上書きしない
    Invoke-SendTo @{ Action = 'Check'; OutDir = $outDir } @($dir)
    Assert-True ((Get-ResultFiles $outDir).Count -ge 2) '結果を上書きした'
}

Test-Case 'Invoke-SendTo: 差分は 2 つを選び、版数（無ければ更新日時）の古い方を旧にする' {
    $dir = New-TestDir
    $v10 = Join-Path $dir '受注_v1.0.xlsx'
    $v11 = Join-Path $dir '受注_v1.1.xlsx'
    New-VersionBook $v10 '受注'
    New-VersionBook $v11 '受注明細'
    # 新しい版の方が更新日時が古くても、版数で決める
    (Get-Item -LiteralPath $v11).LastWriteTime = (Get-Date).AddDays(-1)
    $outDir = Join-Path $dir '結果'
    Invoke-SendTo @{ Action = 'Compare'; OutDir = $outDir } @($v11, $v10)
    $rows = Get-SheetValues (Get-ResultFiles $outDir)[0]
    Assert-True (@($rows | Where-Object { $_ -match '\|受注\|受注明細$' }).Count -eq 1) ($rows -join ' / ')

    # 版数が無ければ更新日時の古い方を旧にする
    $a = Join-Path $dir 'a.xlsx'; $b = Join-Path $dir 'b.xlsx'
    Copy-Item -LiteralPath $v10 -Destination $b
    Copy-Item -LiteralPath $v11 -Destination $a
    (Get-Item -LiteralPath $b).LastWriteTime = (Get-Date).AddDays(-2)
    $outDir2 = Join-Path $dir '結果2'
    Invoke-SendTo @{ Action = 'Compare'; OutDir = $outDir2 } @($a, $b)
    $rows = Get-SheetValues (Get-ResultFiles $outDir2)[0]
    Assert-True (@($rows | Where-Object { $_ -match '\|受注\|受注明細$' }).Count -eq 1) ($rows -join ' / ')

    Assert-Throws { Invoke-SendTo @{ Action = 'Compare'; OutDir = $outDir } @($v10) } '2 つ選んでください'
    Assert-Throws { Invoke-SendTo @{ Action = 'Compare'; OutDir = $outDir } @($v10, $dir) } 'ファイルどうし'
}

Test-Case 'Invoke-SendTo: 検索は入力した文字そのものを探す。対象が無い・見つからないときはエラーを知らせる' {
    $dir = New-TestDir
    New-VersionBook (Join-Path $dir '受注 (1).xlsx') '受注.明細'
    New-VersionBook (Join-Path $dir '顧客.xlsx') '受注X明細'
    $outDir = Join-Path $dir '結果'
    Invoke-SendTo @{ Action = 'Search'; OutDir = $outDir; Text = '受注.明細' } @($dir)
    $rows = Get-SheetValues (Get-ResultFiles $outDir)[0]
    Assert-True (@($rows | Where-Object { $_ -match '受注 \(1\)\.xlsx' }).Count -eq 1) ($rows -join ' / ')
    Assert-True (@($rows | Where-Object { $_ -match '顧客\.xlsx' }).Count -eq 0) '正規表現として探した'

    Assert-Throws { Invoke-SendTo @{ Action = 'Check'; OutDir = $outDir } @() } '選んで'
    Assert-Throws { Invoke-SendTo @{ Action = 'Check'; OutDir = $outDir } @((Join-Path $dir 'none.xlsx')) } '見つかりません'
}

Test-Case 'Invoke-SendTo: 表示を整えるは、選んだファイルを上書きする' {
    $path = Copy-Fixture 'deliverable/bad.xlsx'
    $before = (Get-FileHash -LiteralPath $path).Hash
    Invoke-SendTo @{ Action = 'ResetView' } @($path)
    Assert-True ($before -ne (Get-FileHash -LiteralPath $path).Hash) '書き換えていない'
    $issues = @(& (Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1') -Path $path -Check View 6>$null | Where-Object { $_.Item -ne '表示モード' })
    Assert-Equal 0 $issues.Count ((@($issues | ForEach-Object { $_.Detail })) -join ' / ')
}

Test-Case 'Install-SendTo: Windows 以外では使えないと知らせる' -NonWindowsOnly {
    Assert-Throws { & $InstallSendTo -SendToDir (New-TestDir) } 'Windows でのみ'
}

Test-Case 'Install-SendTo: 「送る」にショートカットを作り、ほかのショートカットは置き換えない。-Uninstall で自分のものだけ消す' -WindowsOnly {
    $sendTo = New-TestDir
    $wsh = New-Object -ComObject WScript.Shell
    # 同じ名前のほかのショートカット
    $other = $wsh.CreateShortcut((Join-Path $sendTo 'Excel 差分（2 つ選ぶ）.lnk'))
    $other.TargetPath = Join-Path $env:SystemRoot 'notepad.exe'
    $other.Save()

    $results = @(& $InstallSendTo -SendToDir $sendTo -WhatIf 6>$null)
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $sendTo -Filter '*.lnk').Count '-WhatIf で作った'
    Assert-Equal 8 @($results | Where-Object Status -eq '作成予定').Count

    $results = @(& $InstallSendTo -SendToDir $sendTo -Action Check, Compare, Search 6>$null)
    Assert-Equal @('Excel 提出前チェック|作成', 'Excel 差分（2 つ選ぶ）|スキップ', 'Excel 検索|作成') @($results | ForEach-Object { '{0}|{1}' -f $_.Name, $_.Status })
    $link = $wsh.CreateShortcut((Join-Path $sendTo 'Excel 提出前チェック.lnk'))
    Assert-Match $link.TargetPath '\\WindowsPowerShell\\v1\.0\\powershell\.exe$'
    Assert-Equal ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Action Check' -f $SendTo) $link.Arguments
    Assert-Equal $script:ToolsDir $link.WorkingDirectory

    # もう一度実行すると更新、-Force でほかのショートカットも置き換える
    $results = @(& $InstallSendTo -SendToDir $sendTo -Action Check, Compare -Force -OutDir 'C:\結果\' 6>$null)
    Assert-Equal @('Excel 提出前チェック|更新', 'Excel 差分（2 つ選ぶ）|更新') @($results | ForEach-Object { '{0}|{1}' -f $_.Name, $_.Status })
    $link = $wsh.CreateShortcut((Join-Path $sendTo 'Excel 差分（2 つ選ぶ）.lnk'))
    Assert-Equal ('-NoProfile -ExecutionPolicy Bypass -File "{0}" -Action Compare -OutDir "C:\結果"' -f $SendTo) $link.Arguments

    # 自分のものだけを消す
    $other = $wsh.CreateShortcut((Join-Path $sendTo 'メモ帳.lnk'))
    $other.TargetPath = Join-Path $env:SystemRoot 'notepad.exe'
    $other.Save()
    $results = @(& $InstallSendTo -SendToDir $sendTo -Uninstall 6>$null)
    Assert-Equal 3 $results.Count
    Assert-Equal @('メモ帳.lnk') @(Get-ChildItem -LiteralPath $sendTo -Filter '*.lnk' | ForEach-Object { $_.Name })

    Assert-Throws { & $InstallSendTo -SendToDir $sendTo -Action Nothing 6>$null } '登録できない -Action'
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh)
}

Test-Case 'Install-SendTo: 作ったショートカットのコマンドで、空白や日本語を含むパスのファイルをチェックできる' -WindowsOnly {
    $sendTo = New-TestDir
    $null = & $InstallSendTo -SendToDir $sendTo -Action Check 6>$null
    $wsh = New-Object -ComObject WScript.Shell
    $link = $wsh.CreateShortcut((Join-Path $sendTo 'Excel 提出前チェック.lnk'))
    $dir = Join-Path (New-TestDir) '提出 物'
    New-Item -ItemType Directory -Path $dir | Out-Null
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/bad.xlsx') -Destination (Join-Path $dir '設計書 v1.0.xlsx')
    $outDir = Join-Path (New-TestDir) '結果'
    # 「送る」は、ショートカットの引数の後ろに、選んだファイルを引用符で囲んで付けて起動する
    $argLine = '{0} -OutDir "{1}" -NoOpen -NoPause "{2}"' -f $link.Arguments, $outDir, (Join-Path $dir '設計書 v1.0.xlsx')
    # エクスプローラーから起動したときと同じく、PowerShell 7 のモジュールのパスを引き継がないようにする
    $modulePath = $env:PSModulePath
    try {
        $env:PSModulePath = [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')
        $p = Start-Process -FilePath $link.TargetPath -ArgumentList $argLine -Wait -PassThru -WindowStyle Hidden
        $failed = Start-Process -FilePath $link.TargetPath -ArgumentList ('{0} -NoOpen -NoPause' -f $link.Arguments) -Wait -PassThru -WindowStyle Hidden
    }
    finally { $env:PSModulePath = $modulePath }
    Assert-Equal 0 $p.ExitCode
    $files = Get-ResultFiles $outDir
    Assert-Equal 1 $files.Count
    Assert-True (@((Get-SheetValues $files[0]) | Where-Object { $_ -like '設計書 v1.0.xlsx|*' }).Count -gt 0) '指摘がない'

    # ファイルを選ばずに起動したとき（失敗）は終了コードで知らせる
    Assert-Equal 1 $failed.ExitCode
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wsh)
}

Test-Case 'Invoke-SendTo / Install-SendTo: Get-Help でヘルプが表示される' {
    foreach ($tool in @($SendTo, $InstallSendTo)) {
        $h = Get-Help $tool
        Assert-Match $h.Synopsis '送る'
        Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description -and $_.name -notin @('WhatIf', 'Confirm') }).Count -eq 0) "説明のないパラメーター: $tool"
    }
}
