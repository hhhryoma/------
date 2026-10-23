#     Install-ExcelTools.ps1 のテスト（共有フォルダの tools を、この PC に入れて更新する）

$InstallTools = Join-Path $script:ToolsDir 'Install-ExcelTools.ps1'

function New-ToolsShare {
    #     配布元（共有フォルダ）の代わりに、tools フォルダの写しを作る
    $share = Join-Path (New-TestDir) 'share/tools'
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $share) | Out-Null
    Copy-Item -LiteralPath $script:ToolsDir -Destination $share -Recurse
    return $share
}

function Set-LastCheck {
    #     前の確認を昔にする（-CheckOnly の間隔を過ぎたことにする）
    param([string] $InstallDir)
    $path = Join-Path $InstallDir 'install.json'
    $text = [System.IO.File]::ReadAllText($path) -replace '"LastCheck":\s*"[^"]*"', '"LastCheck": "2000-01-01T00:00:00"'
    [System.IO.File]::WriteAllText($path, $text)
}

Test-Case 'Install-ExcelTools: 配布元の tools をこの PC に写し、配布元が変わったら変わったファイルだけを写す' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $r = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    Assert-Equal '入れた' $r.Result
    $count = @(Get-ChildItem -LiteralPath $share -Recurse -File).Count
    Assert-Equal $count @(Get-ChildItem -LiteralPath (Join-Path $dir 'tools') -Recurse -File).Count '全部を写す'
    $info = Get-Content -LiteralPath (Join-Path $dir 'install.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Equal (ConvertTo-Slash $share) (ConvertTo-Slash $info.Source)
    Assert-True $info.NoSendTo '登録の指定を覚える'

    # 変わっていなければ何もしない
    Assert-Equal '最新' (& (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null).Result

    # 配布元を更新する（書き換え・追加・削除）
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記'
    Set-Content -LiteralPath (Join-Path $share 'term-lists/new.csv') -Value 'a,b' -Encoding UTF8
    Remove-Item -LiteralPath (Join-Path $share 'Measure-CodeSize.ps1')
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null
    Assert-Equal '更新|2|1' ('{0}|{1}|{2}' -f $r.Result, $r.Copied, $r.Removed)
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '追記'
    Assert-True (Test-Path -LiteralPath (Join-Path $dir 'tools/term-lists/new.csv'))
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $dir 'tools/Measure-CodeSize.ps1'))) '配布元に無いファイルは消す'

    # -CheckOnly は、前の確認から間隔がたつまで配布元を見ない
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記 2'
    Assert-Equal '確認済み' (& (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -CheckOnly 6>$null).Result
    Set-LastCheck $dir
    Assert-Equal '更新' (& (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -CheckOnly 6>$null).Result

    # 配布元に届かないときは、知らせずに次の機会にする
    Set-LastCheck $dir
    Rename-Item -LiteralPath $share -NewName 'tools-moved'
    Assert-Equal '配布元に届かない' (& (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -CheckOnly 6>$null).Result
    Assert-Throws { & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null } '配布元の tools フォルダが見つかりません'
}

Test-Case 'Install-ExcelTools: 入れたツールを「送る」から使ったあとに、配布元の更新を写す' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '更新した版'
    Set-LastCheck $dir
    $data = New-TestDir
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/bad.xlsx') -Destination (Join-Path $data 'a.xlsx')
    $text = Get-HostText { & (Join-Path $dir 'tools/Invoke-SendTo.ps1') -Action Search -Text 'x' -OutDir (Join-Path $data 'out') -NoOpen -NoPause (Join-Path $data 'a.xlsx') | Out-Null }
    Assert-Match $text '新しい版に更新しました（次の実行から使います）: 1 ファイル'
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '更新した版'
}

Test-Case 'Install-ExcelTools: 入れ直し・削除と、使えない指定を知らせる' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    Assert-Throws { & $InstallTools -Source (New-TestDir) -InstallDir $dir -NoSendTo 6>$null } '配布する tools フォルダではありません'
    Assert-Throws { & $InstallTools -InstallDir $dir -Uninstall 6>$null } '入れたツールが見つかりません'
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    # この PC に入れた写しから実行すると、覚えている配布元から入れ直す
    Remove-Item -LiteralPath (Join-Path $dir 'tools/Search-Excel.ps1')
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir 6>$null
    Assert-Equal '入れ直した|1' ('{0}|{1}' -f $r.Result, $r.Copied)
    Assert-Throws { & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -Source (Join-Path $dir 'tools') -InstallDir $dir 6>$null } '配布元と入れる先が同じです'
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Uninstall 6>$null
    Assert-Equal '削除' $r.Result
    Assert-True (-not (Test-Path -LiteralPath $dir))
}

Test-Case 'Install-ExcelTools: 「送る」には、この PC に入れた写しのツールを登録し、更新のときも同じ指定で登録し直す' -WindowsOnly {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $links = New-TestDir
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -SendToDir $links -Action Check, Compare 6>$null
    $made = @(Get-ChildItem -LiteralPath $links -Filter '*.lnk' | ForEach-Object { $_.BaseName } | Sort-Object)
    Assert-Equal @('Excel 差分（2 つ選ぶ）', 'Excel 提出前チェック') $made
    $link = [ExcelTools.ShellLink]::Read((Join-Path $links 'Excel 提出前チェック.lnk'))
    Assert-Match $link.Arguments ([regex]::Escape((Join-Path $dir 'tools\Invoke-SendTo.ps1')))
    Remove-Item -LiteralPath (Join-Path $links '*.lnk')
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記'
    $null = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null
    Assert-Equal 2 @(Get-ChildItem -LiteralPath $links -Filter '*.lnk').Count '更新のときも、入れたときの -Action で登録する'
    $null = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Uninstall 6>$null
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $links -Filter '*.lnk').Count
}

Test-Case 'Install-ExcelTools: Get-Help でヘルプが表示される' {
    $h = Get-Help $InstallTools
    Assert-Match $h.Synopsis 'この PC に入れて'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
