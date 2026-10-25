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
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $share 'term-lists/new.csv') -Value 'a,b' -Encoding UTF8
    Remove-Item -LiteralPath (Join-Path $share 'Measure-CodeSize.ps1')
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null
    Assert-Equal '更新|2|1' ('{0}|{1}|{2}' -f $r.Result, $r.Copied, $r.Removed)
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '追記'
    Assert-True (Test-Path -LiteralPath (Join-Path $dir 'tools/term-lists/new.csv'))
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $dir 'tools/Measure-CodeSize.ps1'))) '配布元に無いファイルは消す'

    # -CheckOnly は、前の確認から間隔がたつまで配布元を見ない
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記 2' -Encoding UTF8
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
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '更新した版' -Encoding UTF8
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

Test-Case 'Install-ExcelTools: 「送る」とスタートメニューの「Excel ツール」には、この PC に入れた写しのツールを登録し、更新のときも同じ指定で登録し直す' -WindowsOnly {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $links = New-TestDir
    $menu = Join-Path (New-TestDir) 'スタートメニュー'
    $text = Get-HostText { $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -SendToDir $links -ShortcutDir $menu -Action Check, Compare }
    Assert-Match $text 'スタートメニューの「Excel ツール」から、ツールの画面を開けます'
    $made = @(Get-ChildItem -LiteralPath $links -Filter '*.lnk' | ForEach-Object { $_.BaseName } | Sort-Object)
    Assert-Equal @('Excel 差分（2 つ選ぶ）', 'Excel 提出前チェック') $made
    $link = [ExcelTools.ShellLink]::Read((Join-Path $links 'Excel 提出前チェック.lnk'))
    Assert-Match $link.Arguments ([regex]::Escape((Join-Path $dir 'tools\Invoke-SendTo.ps1')))
    $gui = [ExcelTools.ShellLink]::Read((Join-Path $menu 'Excel ツール.lnk'))
    Assert-Equal ('-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "{0}"' -f (Join-Path $dir 'tools\Start-ExcelTools.ps1')) $gui.Arguments
    Get-ChildItem -LiteralPath $links -Filter '*.lnk' | Remove-Item
    Remove-Item -LiteralPath (Join-Path $menu 'Excel ツール.lnk')
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '追記' -Encoding UTF8
    $null = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null
    Assert-Equal 2 @(Get-ChildItem -LiteralPath $links -Filter '*.lnk').Count '更新のときも、入れたときの -Action で登録する'
    Assert-True (Test-Path -LiteralPath (Join-Path $menu 'Excel ツール.lnk')) '更新のときも、入れたときのフォルダに画面のショートカットを作る'
    $null = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Uninstall 6>$null
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $links -Filter '*.lnk').Count
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $menu -Filter '*.lnk').Count '画面のショートカットも消す'
}

Test-Case 'Install-ExcelTools: 削除では入れたもの（tools と install.json）だけを消し、ほかのファイルのあるフォルダは残す。このスクリプトで入れていない tools フォルダには入れない' {
    $share = New-ToolsShare
    # 自分のファイルのあるフォルダに入れてから消す
    $dir = Join-Path (New-TestDir) '仕事'
    New-Item -ItemType Directory -Force -Path (Join-Path $dir '報告書') | Out-Null
    Set-Content -LiteralPath (Join-Path $dir '報告書/週報.txt') -Value '自分のファイル' -Encoding UTF8
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    $text = Get-HostText { & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Uninstall | Out-Null }
    Assert-Match $text 'ほかのファイルがあるため、フォルダは残しました'
    Assert-Equal @('報告書') @(Get-ChildItem -LiteralPath $dir -Force | ForEach-Object { $_.Name }) 'tools と install.json だけを消す'
    Assert-Equal '自分のファイル' ([System.IO.File]::ReadAllText((Join-Path $dir '報告書/週報.txt')).Trim())

    # 自分のスクリプトを置いた tools フォルダには入れない（配布元に無いファイルを消してしまうため）
    $own = Join-Path (New-TestDir) '作業'
    New-Item -ItemType Directory -Force -Path (Join-Path $own 'tools') | Out-Null
    Set-Content -LiteralPath (Join-Path $own 'tools/自作マクロ.ps1') -Value '# 自分のスクリプト' -Encoding UTF8
    Assert-Throws { & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $own -NoSendTo 6>$null } 'Install-ExcelTools\.ps1 で入れたものではない tools フォルダがあります'
    Assert-Equal @('自作マクロ.ps1') @(Get-ChildItem -LiteralPath (Join-Path $own 'tools') -Force | ForEach-Object { $_.Name })
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $own 'install.json'))) '入れたことにしない'
}

Test-Case 'Install-ExcelTools: 相対パスの -OutDir・-SendToDir・-ShortcutDir は、いまのフォルダからの絶対パスにして覚える（更新のときに tools フォルダからのパスにならない）' {
    $share = New-ToolsShare
    $base = New-TestDir
    $dir = Join-Path $base 'ExcelTools'
    $work = Join-Path $base 'work'
    New-Item -ItemType Directory -Path $work | Out-Null
    Push-Location -LiteralPath $work
    try { $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -OutDir '結果' -SendToDir 'links' -ShortcutDir 'メニュー' -NoSendTo 6>$null }
    finally { Pop-Location }
    $info = Get-Content -LiteralPath (Join-Path $dir 'install.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-Equal @((Join-Path $work '結果'), (Join-Path $work 'links'), (Join-Path $work 'メニュー')) @($info.OutDir, $info.SendToDir, $info.ShortcutDir)
}

Test-Case 'Install-ExcelTools: 更新で写せないファイル（開いている Excel など）があっても残りは写し、写せなかったファイルは次の確認で写し直す' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    foreach ($rel in @('Compare-Excel.ps1', 'estimate/係数表_サンプル.csv', 'Search-Excel.ps1')) { Add-Content -LiteralPath (Join-Path $share $rel) -Value '# 新しい版' -Encoding UTF8 }
    $data = New-TestDir
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/bad.xlsx') -Destination (Join-Path $data 'a.xlsx')
    # 係数表の見本を、ほかのアプリが書き込ませずに開いている
    $open = [System.IO.File]::Open((Join-Path $dir 'tools/estimate/係数表_サンプル.csv'), 'Open', 'ReadWrite', 'None')
    try {
        $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update -WarningVariable warned -WarningAction SilentlyContinue 6>$null
        # 「送る」で使ったあとの確認でも（確認の間隔を待たずに）写し直し、写せなかったファイルを知らせる
        $text = Get-HostText { & (Join-Path $dir 'tools/Invoke-SendTo.ps1') -Action Search -Text 'x' -OutDir (Join-Path $data 'out') -NoOpen -NoPause (Join-Path $data 'a.xlsx') | Out-Null }
    }
    finally { $open.Dispose() }
    Assert-Equal '一部を更新|2' ('{0}|{1}' -f $r.Result, $r.Copied)
    Assert-Match ($r.Failed -join ' / ') '係数表_サンプル\.csv'
    Assert-Match ($warned -join ' / ') '写せなかったファイルがあります.*係数表_サンプル\.csv'
    foreach ($rel in @('Compare-Excel.ps1', 'Search-Excel.ps1')) { Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir "tools/$rel"))) '# 新しい版' "写せないファイルの前後の $rel も写す" }
    Assert-Match $text '写せなかったファイルがあります.*係数表_サンプル\.csv'
    # ファイルを閉じたあとの次の確認で写す
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -CheckOnly 6>$null
    Assert-Equal '更新|1' ('{0}|{1}' -f $r.Result, $r.Copied)
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/estimate/係数表_サンプル.csv'))) '# 新しい版'
}

Test-Case 'Install-ExcelTools: 入れている間に配布元が変わっても、次の確認で写す（配布元の目印は写す前に取る）' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    # Search-Excel.ps1 を写したところで、管理者が配布元の README.md（写し終えたもの）を更新したことにする
    # （このテストの中だけの Copy-Item。入れるスクリプトの Copy-Item はこれを呼ぶ）
    $readme = Join-Path $share 'README.md'
    function Copy-Item {
        param([string] $LiteralPath, [string] $Destination, [switch] $Force)
        Microsoft.PowerShell.Management\Copy-Item -LiteralPath $LiteralPath -Destination $Destination -Force:$Force
        if ((Split-Path -Leaf $LiteralPath) -eq 'Search-Excel.ps1') { Add-Content -LiteralPath $readme -Value '入れている間の更新' -Encoding UTF8 }
    }
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    Remove-Item -Path Function:\Copy-Item
    Assert-NotMatch ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '入れている間の更新'
    $r = & (Join-Path $dir 'tools/Install-ExcelTools.ps1') -InstallDir $dir -Update 6>$null
    Assert-Equal '更新|1' ('{0}|{1}' -f $r.Result, $r.Copied)
    Assert-Match ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '入れている間の更新'
}

Test-Case 'Install-ExcelTools: 「送る」→「ツールの画面で開く」では更新を確かめない（開いた画面が読むファイルを書き換えない。画面を閉じたあとに確かめる）' {
    $share = New-ToolsShare
    $dir = Join-Path (New-TestDir) 'ExcelTools'
    $null = & (Join-Path $share 'Install-ExcelTools.ps1') -InstallDir $dir -NoSendTo 6>$null
    Add-Content -LiteralPath (Join-Path $share 'README.md') -Value '更新した版' -Encoding UTF8
    Set-LastCheck $dir
    $text = Get-HostText { & (Join-Path $dir 'tools/Invoke-SendTo.ps1') -Action Gui -NoOpen (New-TestDir) | Out-Null }
    Assert-Match $text '画面を開くコマンド'
    Assert-NotMatch $text '更新しました'
    Assert-NotMatch ([System.IO.File]::ReadAllText((Join-Path $dir 'tools/README.md'))) '更新した版'
}

Test-Case 'Install-ExcelTools: Get-Help でヘルプが表示される' {
    $h = Get-Help $InstallTools
    Assert-Match $h.Synopsis 'この PC に入れて'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
