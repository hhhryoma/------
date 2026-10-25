#     Start-ExcelTools.ps1（画面）と ExcelToolsGui.psm1 のテスト。
#     入力欄・引数・実行は画面を使わずにどの OS でも、画面（Windows Forms）は Windows で、表示せずに操作して確かめる

$GuiModule = Join-Path $script:ToolsDir 'ExcelToolsGui.psm1'
$StartTool = Join-Path $script:ToolsDir 'Start-ExcelTools.ps1'
Import-Module $GuiModule -Force -DisableNameChecking

function Get-TestEntry {
    param([string] $Id)
    return @((Get-GuiCatalog).Tools | Where-Object Id -eq $Id)[0]
}

function Get-FieldKeys {
    #     入力欄を「名前:種類（[] は複数・* は必須）」の並びにする
    param($Fields)
    return @($Fields | ForEach-Object { '{0}:{1}{2}{3}' -f $_.Name, $_.Kind, $(if ($_.Multiple) { '[]' } else { '' }), $(if ($_.Required) { '*' } else { '' }) })
}

Test-Case 'Start-ExcelTools: ツールの一覧のすべての処理で入力欄を作れ、導入と「送る」のほかのツールはすべて一覧にある' {
    $catalog = Get-GuiCatalog
    $listed = @($catalog.Tools | ForEach-Object { $_.Tool } | Sort-Object -Unique)
    $tools = @(Get-ChildItem -LiteralPath $script:ToolsDir -Filter '*.ps1' | Where-Object { $_.Name -notmatch '^(Install-|Invoke-SendTo|Start-ExcelTools)' } | ForEach-Object { $_.Name })
    Assert-Equal @() @($tools | Where-Object { $listed -notcontains $_ }) '一覧に無いツール（gui\ツールの一覧.psd1 に足してください）'
    $kinds = @('Text', 'File', 'Folder', 'Path', 'SaveFile', 'SaveFolder', 'Number', 'Date', 'Switch', 'Choice')
    foreach ($t in $catalog.Tools) {
        $fields = Get-GuiFields $t
        Assert-Equal @($t.Main) @($fields | Where-Object Main | ForEach-Object { $_.Name }) "$($t.Id) の最初の入力欄"
        Assert-Equal @() @($fields | Where-Object { $kinds -notcontains $_.Kind -or -not $_.Label } | ForEach-Object { $_.Name }) "$($t.Id) の入力欄の種類と名前"
        foreach ($o in $t.Output) { Assert-Equal 'SaveFile' (@($fields | Where-Object Name -eq $o)[0].Kind) "$($t.Id) の結果の欄 $o" }
        foreach ($o in $t.Open) { Assert-True (@($fields | Where-Object Name -eq $o).Count -eq 1) "$($t.Id) の開く欄 $o" }
        Assert-True ($catalog.Groups.Name -contains $t.Group) "$($t.Id) の分類"
    }
}

Test-Case 'Start-ExcelTools: 入力欄の種類・名前・必須は、ツールの引数とヘルプとツールの一覧から決まる' {
    $search = Get-GuiFields (Get-TestEntry 'Search')
    Assert-Equal @('Pattern:Text*', 'Path:Path[]*', 'SimpleMatch:Switch', 'IncludeWord:Switch', 'OutXlsx:SaveFile') (Get-FieldKeys @($search | Where-Object Main))
    Assert-Equal '探す文字' $search[0].Label 'ツールの一覧の名前'
    $target = @($search | Where-Object Name -eq 'Target')[0]
    Assert-Equal 'Choice|True|Value' ('{0}|{1}|{2}' -f $target.Kind, $target.Multiple, $target.Choices[0]) '選択肢（ValidateSet）は選ぶ欄'
    Assert-Equal '.xlsx' @($search | Where-Object Name -eq 'OutXlsx')[0].Extension
    Assert-True (@($search | Where-Object Name -eq 'Open').Count -eq 0) '画面で使えない引数（-Open）は出さない'
    Assert-Equal $true (@($search | Where-Object Name -eq 'SimpleMatch')[0].Default) '最初の値'

    # 引数のセットごとの処理（差分は旧・新、版を続けて比べるは版の一覧）
    Assert-Equal @('Old:Path*', 'New:Path*', 'OutHtml:SaveFile', 'OutXlsx:SaveFile') (Get-FieldKeys @(Get-GuiFields (Get-TestEntry 'Compare') | Where-Object Main))
    Assert-Equal @('Versions:Path[]*', 'OutHtml:SaveFile', 'OutXlsx:SaveFile') (Get-FieldKeys @(Get-GuiFields (Get-TestEntry 'CompareVersions') | Where-Object Main))
    $bump = Get-GuiFields (Get-TestEntry 'Bump')
    Assert-True (@($bump | Where-Object { @('Bump', 'Template', 'Book') -contains $_.Name }).Count -eq 0) '版を上げるは -Bump を付けて渡し、ほかのセットの引数は出さない'

    # ヘルプから: 日付・数・名前（ファイルを選ぶ欄にしない）・保存先のフォルダ・ファイルとフォルダ
    $progress = Get-GuiFields (Get-TestEntry 'Progress')
    Assert-Equal 'Date' @($progress | Where-Object Name -eq 'AsOf')[0].Kind
    Assert-Equal 'File' @($progress | Where-Object Name -eq 'FillTemplate')[0].Kind 'テンプレートのファイル'
    $newDoc = Get-GuiFields (Get-TestEntry 'NewDoc')
    Assert-Equal 'Text|Text' ('{0}|{1}' -f @($newDoc | Where-Object Name -eq 'FileName')[0].Kind, @($newDoc | Where-Object Name -eq 'ListSheet')[0].Kind) 'ファイル名・シートの名前は文字'
    Assert-Equal 'SaveFolder' @($newDoc | Where-Object Name -eq 'OutDir')[0].Kind
    Assert-Equal 'Number' @(Get-GuiFields (Get-TestEntry 'Effort') | Where-Object Name -eq 'HoursPerDay')[0].Kind
    Assert-Equal 'Path' @(Get-GuiFields (Get-TestEntry 'Quality') | Where-Object Name -eq 'Config')[0].Kind '括弧の中の「。」で文を切らない（設定ファイルか、それを置いたフォルダ）'
    Assert-Equal 'SaveFile' @(Get-GuiFields (Get-TestEntry 'Pdf') | Where-Object Name -eq 'LogPath')[0].Kind '「保存します」の CSV は保存先'
    Assert-Equal @('Path:Path[]*', 'Checklist:File', 'OutDir:SaveFolder*', 'Baseline:Path', 'Recheck:Path') (Get-FieldKeys @(Get-GuiFields (Get-TestEntry 'Review') | Where-Object Main)) 'AI レビューは前の版・前のレビュー結果（ファイルかフォルダ）も最初から出す'
    Assert-Equal '改訂履歴のシート名の正規表現' @(Get-GuiFields (Get-TestEntry 'Compare') | Where-Object Name -eq 'HistorySheet')[0].Label 'ヘルプからの名前（長いときは最後の「、」のあと）'
    Assert-Equal '結果を CSV に保存します' @(Get-GuiFields (Get-TestEntry 'Search') | Where-Object Name -eq 'OutCsv')[0].Label '括弧を除いても、英数字と日本語の間の空白は残す（CSV（UTF-8・BOM 付き）に保存）'
    Assert-Equal '設計書の一覧を Excel に保存します' (Get-GuiLabel -Help '設計書（Word を含む）の一覧を Excel（xlsx）に保存します。続きの文。' -Name 'OutXlsx') '日本語の間の括弧は詰める'
}

Test-Case 'Start-ExcelTools: 入力から引数を作る（空は渡さず、必須の欠け・数と日付の誤りを知らせ、ワイルドカードのパスは逃がす）' {
    $search = Get-TestEntry 'Search'
    $fields = Get-GuiFields $search
    $r = ConvertTo-GuiArguments $search $fields @{ Pattern = '受注'; Path = @('D:\設計書[旧]', ' "D:\b" ', ''); SimpleMatch = $true; IncludeWord = $false; Target = [string[]]@('Value', 'Comment'); OutXlsx = '' }
    Assert-Equal 0 $r.Errors.Count ($r.Errors -join ' / ')
    Assert-Equal @('Pattern', 'Path', 'SimpleMatch', 'Target') @($r.Arguments.Keys) '空・選ばないスイッチは渡さない'
    Assert-Equal @('D:\設計書`[旧`]', 'D:\b') $r.Arguments['Path'] 'ワイルドカードを使える引数のパスは逃がし、前後の空白と引用符は除く'
    Assert-Equal @('Value', 'Comment') $r.Arguments['Target']

    $r = ConvertTo-GuiArguments $search $fields @{ Pattern = ''; Path = @() }
    Assert-Equal @('「探す文字」を入れてください', '「探すファイル・フォルダ」を入れてください') $r.Errors
    Assert-Equal @('Pattern', 'Path') $r.ErrorFields '誤りのある欄'
    $r = ConvertTo-GuiArguments $search $fields @{ Pattern = ' "OK" '; Path = @('D:') }
    Assert-Equal ' "OK" ' $r.Arguments['Pattern'] '探す文字は入力のまま（前後の空白と引用符を除かない）'

    # 置換: 全角の空白を半角に・空欄なら消す
    $replace = Get-TestEntry 'Replace'
    $replaceFields = Get-GuiFields $replace
    $r = ConvertTo-GuiArguments $replace $replaceFields @{ Pattern = '　'; Replacement = ' '; Path = @('D:') }
    Assert-Equal '[　]→[ ]' ('[{0}]→[{1}]' -f $r.Arguments['Pattern'], $r.Arguments['Replacement']) '空白だけの入力も渡す'
    $r = ConvertTo-GuiArguments $replace $replaceFields @{ Pattern = '(仮)'; Replacement = ''; Path = @('D:') }
    Assert-True ($r.Arguments.Contains('Replacement') -and $r.Arguments['Replacement'] -eq '') '置き換えた後の文字が空欄なら、空の文字を渡す（消す）'
    Assert-Match (Format-GuiCommandLine 'Replace-Excel.ps1' $r.Arguments) "-Replacement ''"

    $effort = Get-TestEntry 'Effort'
    $r = ConvertTo-GuiArguments $effort (Get-GuiFields $effort) @{ Path = @('a.xlsx'); HoursPerDay = '７.５'; Tolerance = 'たくさん'; AsOf = '2026/10/20' }
    Assert-Equal @('「超過の見込みとみなす、予定工数を超える割合」は数で入れてください') $r.Errors
    Assert-Equal @('Tolerance') $r.ErrorFields
    Assert-Equal '7.5' ([string]$r.Arguments['HoursPerDay']) '全角の数字も読む'
    Assert-True ($r.Arguments['AsOf'] -is [datetime] -and $r.Arguments['AsOf'].Day -eq 20) '日付の引数には日付を渡す'
    $wareki = Invoke-InWareki { ConvertTo-GuiArguments $effort (Get-GuiFields $effort) @{ Path = @('a.xlsx'); AsOf = '2026/10/20' } }
    Assert-Equal '2026-10-20' $wareki.Arguments['AsOf'].ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) '和暦の設定でも、入れた日付は西暦として読む'

    $consistency = Get-TestEntry 'Consistency'
    $r = ConvertTo-GuiArguments $consistency (Get-GuiFields $consistency) @{ Path = @('x'); Where = 'T_ORDER, E0001、C005' }
    Assert-Equal @('T_ORDER', 'E0001', 'C005') $r.Arguments['Where'] '複数の文字はカンマ・読点で区切る'

    $bump = Get-TestEntry 'Bump'
    $r = ConvertTo-GuiArguments $bump (Get-GuiFields $bump) @{ Path = @('a_v1.0.xlsx'); Description = '誤記の修正' }
    Assert-Equal $true $r.Arguments['Bump'] 'いつも付ける引数（Fixed）'
}

Test-Case 'Start-ExcelTools: 結果の既定の保存先（名前_日時。同じ名前があれば連番）と、同じことをするコマンド・PowerShell の画面で実行するスクリプト' {
    $dir = New-TestDir
    $compare = Get-TestEntry 'Compare'
    $fields = Get-GuiFields $compare
    $now = Get-Date '2026-10-23 10:15:00'
    [System.IO.File]::WriteAllText((Join-Path $dir '差分_20261023_101500.html'), '')
    $outs = Get-GuiDefaultOutputs $compare $fields @{ OutXlsx = 'D:\自分で決めた.xlsx' } -Dir $dir -Now $now
    Assert-Equal @('OutHtml') @($outs.Keys) '入れてある欄はそのまま'
    Assert-Equal '差分_20261023_101500(2).html' (Split-Path -Leaf $outs['OutHtml'])

    $arguments = [ordered]@{ Old = "D:\旧'版"; New = 'D:\新'; KeyColumn = [string[]]@('物理名', '項目ID'); NoRecurse = $true; AsOf = (Get-Date '2026-10-20'); MinBugs = 5 }
    Assert-Equal "& 'C:\tools\Compare-Excel.ps1' -Old 'D:\旧''版' -New 'D:\新' -KeyColumn '物理名', '項目ID' -NoRecurse -AsOf ([datetime]'2026-10-20') -MinBugs 5" (Format-GuiCommandLine 'C:\tools\Compare-Excel.ps1' $arguments)
    Assert-Equal "([datetime]'2026-10-20')" (Invoke-InWareki { ConvertTo-GuiLiteral ([datetime]'2026-10-20') }) '和暦の設定でも、日付は西暦で書く'
    Assert-Equal "([datetime]'2026-10-20')|09:05:03" (Invoke-InDashDate { '{0}|{1}' -f (ConvertTo-GuiLiteral ([datetime]'2026-10-20')), (Format-GuiTime ([datetime]'2026-10-20 09:05:03')) }) '日付・時刻の区切りを変えた設定でも、- と : で書く'

    # フォルダに保存する処理の結果: 実行を始めてから書いた Excel を、文字や古いファイルより先にする
    $results = Join-Path $dir '報告'
    New-Item -ItemType Directory -Path (Join-Path $results 'sub') -Force | Out-Null
    $start = Get-Date
    foreach ($n in @('古い.xlsx', '報告.txt', '報告.xlsx', 'メモ.docx', 'sub/中.xlsx')) { [System.IO.File]::WriteAllText((Join-Path $results $n), 'x') }
    (Get-Item -LiteralPath (Join-Path $results '古い.xlsx')).LastWriteTime = $start.AddDays(-1)
    Assert-Equal '報告.xlsx' (Split-Path -Leaf (Find-GuiNewResult $results $start.AddSeconds(-2)))
    Remove-Item -LiteralPath (Join-Path $results '報告.xlsx')
    Assert-Equal '報告.txt' (Split-Path -Leaf (Find-GuiNewResult $results $start.AddSeconds(-2))) 'Excel が無ければ文字'
    Assert-Equal $null (Find-GuiNewResult $results $start.AddMinutes(1)) '実行のあとに書いたものが無い'

    # 旧・新の順（版数があれば版数、無ければ更新日時）
    $v10 = Join-Path $dir '画面設計書_v1.10.xlsx'
    $v9 = Join-Path $dir '画面設計書_v1.9.xlsx'
    foreach ($p in @($v10, $v9)) { [System.IO.File]::WriteAllText($p, 'x') }
    Assert-Equal @($v9, $v10) (Get-GuiOldNewOrder @($v10, $v9)) '版数は数で比べる（1.9 → 1.10）'
    $a = Join-Path $dir 'a.xlsx'
    $b = Join-Path $dir 'b.xlsx'
    foreach ($p in @($a, $b)) { [System.IO.File]::WriteAllText($p, 'x') }
    (Get-Item -LiteralPath $a).LastWriteTime = (Get-Date).AddDays(-1)
    Assert-Equal @($a, $b) (Get-GuiOldNewOrder @($b, $a)) '版数が無ければ更新日時の古い方が旧'

    $old = Join-Path $dir '実行_old.ps1'
    [System.IO.File]::WriteAllText($old, '')
    (Get-Item -LiteralPath $old).LastWriteTime = (Get-Date).AddDays(-2)
    $script = New-GuiConsoleScript 'C:\tools\Invoke-QualityAnalysis.ps1' $arguments '定性分析' -Dir $dir
    Assert-True (-not (Test-Path -LiteralPath $old)) '前に作った古いスクリプトは消す'
    $text = [System.IO.File]::ReadAllText($script)
    Assert-Match $text ([regex]::Escape("`$arguments['KeyColumn'] = @('物理名', '項目ID')"))
    Assert-Match $text 'Enter キーを押すと閉じます'
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$null, [ref]$errors)
    Assert-Equal 0 @($errors).Count 'PowerShell のスクリプトとして読める'
}

Test-Case 'Start-ExcelTools: ツールを別の実行空間で動かし、Write-Host・警告・出力の件数と表を受け取り、中止・エラーも知らせる' {
    $dir = New-TestDir
    $search = Get-TestEntry 'Search'
    $out = Join-Path $dir '検索.xlsx'
    $built = ConvertTo-GuiArguments $search (Get-GuiFields $search) @{ Pattern = 'TODO'; Path = @((Join-Path $script:Fixtures 'grep/sample.xlsx')); SimpleMatch = $true; OutXlsx = $out }
    $r = Wait-GuiToolRun (Start-GuiToolRun $search.ToolPath $built.Arguments) -TimeoutSeconds 120
    Assert-True $r.Succeeded $r.Error
    Assert-True (@($r.Messages | Where-Object { $_.Text -match '1 ファイル中 1 ファイルで \d+ 件ヒット' }).Count -eq 1) (($r.Messages | ForEach-Object { $_.Text }) -join ' / ')
    Assert-True ($r.Summary.Count -gt 0) '出力の件数'
    Assert-Match $r.Summary.Text 'Location'
    Assert-True (Test-Path -LiteralPath $out) '結果のファイル'

    $script = Join-Path $dir 'tool.ps1'
    [System.IO.File]::WriteAllText($script, "param([string] `$Name)`r`nWrite-Host ""始めます: `$Name"" -ForegroundColor Cyan`r`nWrite-Warning '注意'`r`nthrow '読めません'`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $r = Wait-GuiToolRun (Start-GuiToolRun $script @{ Name = '設計書' }) -TimeoutSeconds 60
    Assert-Equal 'False|読めません' ('{0}|{1}' -f $r.Succeeded, $r.Error)
    Assert-Equal @('始めます: 設計書|Cyan', '警告: 注意|DarkYellow') @($r.Messages | ForEach-Object { '{0}|{1}' -f $_.Text, $_.Color })

    # 失敗を exit で知らせるツールは失敗にする
    $exit = Join-Path $dir 'exit.ps1'
    [System.IO.File]::WriteAllText($exit, "Write-Host '1 件失敗'`r`nexit 1`r`n")
    $r = Wait-GuiToolRun (Start-GuiToolRun $exit @{}) -TimeoutSeconds 60
    Assert-Equal 'False|1' ('{0}|{1}' -f $r.Succeeded, $r.Summary.ExitCode)
    Assert-Match $r.Error '終了コード 1'

    $slow = Join-Path $dir 'slow.ps1'
    [System.IO.File]::WriteAllText($slow, "Start-Sleep -Seconds 60`r`n")
    $run = Start-GuiToolRun $slow @{}
    Start-Sleep -Milliseconds 500
    $null = Stop-GuiToolRun $run
    $r = Wait-GuiToolRun $run -TimeoutSeconds 30
    Assert-Equal 'True|中止しました' ('{0}|{1}' -f $r.Stopped, $r.Error)

    # 画面を閉じるときは、ツールの後片付け（finally で Excel を閉じるなど）が終わるまで待つ
    $marker = Join-Path $dir '片付けた.txt'
    $cleanup = Join-Path $dir 'cleanup.ps1'
    [System.IO.File]::WriteAllText($cleanup, ("try {{ Start-Sleep -Seconds 60 }}`r`nfinally {{ [System.Threading.Thread]::Sleep(800); [System.IO.File]::WriteAllText('{0}', 'x') }}`r`n" -f $marker.Replace("'", "''")), (New-Object System.Text.UTF8Encoding($true)))
    $run = Start-GuiToolRun $cleanup @{}
    Start-Sleep -Milliseconds 500
    Assert-True (Stop-GuiToolRun $run -WaitSeconds 30) '止まるまで待つ'
    Assert-True (Test-Path -LiteralPath $marker) '後片付けが終わっている'
    Assert-True (Wait-GuiToolRun $run -TimeoutSeconds 10).Stopped '中止'
}

Test-Case 'Start-ExcelTools: ツールが中で呼んだツールの、6>$null・3>$null で捨てた Write-Host・警告は記録の欄に出さない' {
    $dir = New-TestDir
    $utf8 = New-Object System.Text.UTF8Encoding($true)
    $inner = Join-Path $dir 'inner.ps1'
    [System.IO.File]::WriteAllText($inner, "Write-Host '中のツール'`r`nWrite-Information '中の情報'`r`nWrite-Warning '中の警告'`r`n'中の出力'`r`n", $utf8)
    $outer = Join-Path $dir 'outer.ps1'
    [System.IO.File]::WriteAllText($outer, ("Write-Host '始めます' -ForegroundColor Cyan`r`n`$x = @(& '{0}' 6>`$null 3>`$null)`r`nWrite-Warning ('受け取った ' + `$x.Count)`r`nWrite-Host '終わり'`r`n'結果'`r`n" -f $inner.Replace("'", "''")), $utf8)
    $r = Wait-GuiToolRun (Start-GuiToolRun $outer @{}) -TimeoutSeconds 60
    Assert-True $r.Succeeded $r.Error
    Assert-Equal @('始めます|Cyan', '警告: 受け取った 1|DarkYellow', '終わり|') @($r.Messages | ForEach-Object { '{0}|{1}' -f $_.Text, $_.Color })
    Assert-Equal 1 $r.Summary.Count '出力の件数'
}

Test-Case 'Start-ExcelTools: ファイル・フォルダを選ぶ画面を開く場所（名前だけ・使えない文字でも止まらない）' {
    $dir = New-TestDir
    $file = Join-Path $dir '設計書.xlsx'
    [System.IO.File]::WriteAllText($file, 'x')
    $r = Get-GuiBrowseStart $dir
    Assert-Equal "$dir|" ('{0}|{1}' -f $r.Dir, $r.Name) 'フォルダはそのフォルダから'
    $r = Get-GuiBrowseStart $file
    Assert-Equal "$dir|設計書.xlsx" ('{0}|{1}' -f $r.Dir, $r.Name) 'ファイルはそのフォルダから、名前を入れて'
    $r = Get-GuiBrowseStart (Join-Path $dir '新しい結果.xlsx')
    Assert-Equal "$dir|新しい結果.xlsx" ('{0}|{1}' -f $r.Dir, $r.Name) 'まだ無いファイルも、あるフォルダから'
    $r = Get-GuiBrowseStart '結果.xlsx'
    Assert-Equal '|結果.xlsx' ('{0}|{1}' -f $r.Dir, $r.Name) '名前だけ（フォルダがわからない）'
    $r = Get-GuiBrowseStart (Join-Path (Join-Path $dir '無いフォルダ') 'a.xlsx')
    Assert-Equal '|a.xlsx' ('{0}|{1}' -f $r.Dir, $r.Name) '無いフォルダ'
    $r = Get-GuiBrowseStart 'a<b>|c'
    Assert-True ($null -eq $r.Dir) '使えない文字'
    Assert-Equal '|' ('{0}|{1}' -f (Get-GuiBrowseStart '').Dir, (Get-GuiBrowseStart ' ').Name) '空'
}

Test-Case 'Start-ExcelTools: 入力を保存して次に読む（結果の保存先と日付は保存しない）。やりたいことを言葉で探す' {
    $dir = New-TestDir
    $path = Join-Path $dir '画面の設定.json'
    $effort = Get-TestEntry 'Effort'
    $keep = Get-GuiValuesToSave $effort (Get-GuiFields $effort) @{ Path = [string[]]@('D:\工数', 'D:\日報'); Wbs = [string[]]@(); AsOf = (Get-Date); NoRecurse = $true; Unit = '時間'; OutXlsx = 'D:\r.xlsx' }
    $settings = Read-GuiSettings $path
    Assert-Equal 'True|False' ('{0}|{1}' -f $settings.OpenResult, $settings.ShowAdvanced) '保存が無いときの既定'
    $settings.Tools['Effort'] = $keep
    $settings.LastTool = 'Effort'
    $settings.ShowAdvanced = $true
    $settings.Window = '-8,20,1200,900,max'
    $settings.Splitter = 340
    Save-GuiSettings $settings $path
    $back = Read-GuiSettings $path
    Assert-Equal 'Effort|True|-8,20,1200,900,max|340' ('{0}|{1}|{2}|{3}' -f $back.LastTool, $back.ShowAdvanced, $back.Window, $back.Splitter) '画面の位置・大きさ・左の一覧の幅も覚える'
    $v = $back.Tools['Effort']
    Assert-Equal @('D:\工数', 'D:\日報') $v['Path']
    Assert-Equal '時間|True' ('{0}|{1}' -f $v['Unit'], $v['NoRecurse'])
    Assert-True (-not $v.ContainsKey('OutXlsx') -and -not $v.ContainsKey('AsOf')) '結果の保存先と日付は保存しない'

    # 最近実行した処理（新しい順・同じものは 1 つ・5 つまで）
    foreach ($id in @('Search', 'Check', 'Compare', 'Search', 'Terms', 'Size', 'Progress')) { Add-GuiRecent $back $id }
    Assert-Equal @('Progress', 'Size', 'Terms', 'Search', 'Compare') $back.Recent
    Save-GuiSettings $back $path
    Assert-Equal @('Progress', 'Size', 'Terms', 'Search', 'Compare') (Read-GuiSettings $path).Recent

    $catalog = Get-GuiCatalog
    Assert-Equal @('Compare', 'SpVersion') @(Find-GuiTool $catalog '差分 画面' | ForEach-Object { $_.Id }) '空白で区切った言葉のすべてを含むもの'
    Assert-Equal @('Check') @(Find-GuiTool $catalog '提出前' | ForEach-Object { $_.Id })
}

Test-Case 'Start-ExcelTools: 画面で処理を選び、入力して実行すると、記録と結果が出る。入力の誤り・試しに実行・ドロップ・PowerShell の画面での実行' -WindowsOnly {
    $dir = New-TestDir
    $form = New-GuiWindow -SettingsPath (Join-Path $dir '画面の設定.json') -Tool 'Search' -Quiet
    try {
        $s = Get-GuiState
        Assert-Equal '文字を探す（検索）' $s.Title.Text
        Assert-Equal @('Pattern', 'Path', 'SimpleMatch', 'IncludeWord', 'OutXlsx') @($s.Bindings | Where-Object { $_.Field.Main } | ForEach-Object { $_.Field.Name })
        Assert-Equal $true $s.Bindings[2].Input.Checked '最初の値（入力した文字そのものを探す）'
        # 表示していない画面では Visible が false になるので、画面が覚えている状態で確かめる
        Assert-True (-not $s.WhatIfShown) '書き換えない処理には「試しに実行」を出さない'
        Assert-True (@($s.Tree.Nodes | ForEach-Object { $_.Text }) -contains '担当者向け') '一覧'

        # 実行: 結果の欄が空なら既定の保存先を入れ、記録に出力が出て、結果を開けるようになる
        Set-GuiValue 'Pattern' 'TODO'
        Set-GuiValue 'Path' @((Join-Path $script:Fixtures 'grep/sample.xlsx'))
        $out = Join-Path $dir '検索.xlsx'
        Set-GuiValue 'OutXlsx' $out
        Invoke-GuiRun
        Wait-GuiRun -TimeoutSeconds 120
        Assert-Match $s.Log.Text '件ヒット'
        Assert-Match $s.Log.Text '終わりました'
        Assert-Equal $out $s.ResultPath
        Assert-True $s.OpenButton.Enabled '結果を開く'
        Assert-Equal '最近使ったもの|文字を探す（検索）' ('{0}|{1}' -f $s.Tree.Nodes[0].Text, $s.Tree.Nodes[0].Nodes[0].Text) '実行した処理は一覧の先頭の「最近使ったもの」に出る'

        # 入力の誤りは実行しない
        Select-GuiTool 'Compare'
        Invoke-GuiRun
        Assert-Equal @('「旧（変更前）のファイル・フォルダ」を入れてください', '「新（変更後）のファイル・フォルダ」を入れてください') $s.LastErrors
        Assert-True ($null -eq $s.Run) '実行していない'
        Assert-Equal 'MistyRose|Window' ('{0}|{1}' -f $s.Bindings[0].Input.BackColor.Name, $s.Bindings[2].Input.BackColor.Name) '誤りのある欄の色を変える'
        Assert-Equal '' $s.Bindings[2].Input.Text '誤りがあるときは、結果の欄に既定の保存先を入れない'

        # 前の入力を戻す
        Select-GuiTool 'Search'
        Assert-Equal 'TODO' $s.Bindings[0].Input.Text

        # 書き換える処理は、最初は「試しに実行」（ファイルは変わらない）
        Select-GuiTool 'ResetView'
        Assert-True ($s.WhatIfShown -and $s.WhatIf.Checked) '試しに実行'
        $book = Copy-Fixture 'deliverable/bad.xlsx'
        $before = (Get-FileHash -LiteralPath $book).Hash
        Add-GuiPaths @($book)
        Assert-Equal @($book) @($s.Bindings[0].Input.Items) 'ドロップしたファイルは空いているパスの欄に入る'
        Invoke-GuiRun
        Wait-GuiRun -TimeoutSeconds 120
        Assert-Match $s.Log.Text '試しに実行'
        Assert-Equal $before (Get-FileHash -LiteralPath $book).Hash 'ファイルは書き換えない'

        # 詳細な設定
        $s.ShowAdvanced.Checked = $false
        Assert-True (-not $s.AdvancedShown) '詳細な設定は隠す'
        $s.ShowAdvanced.Checked = $true
        Assert-True $s.AdvancedShown '詳細な設定を出す'
        # 隠した詳細な設定に入れてある欄の数を出す
        Select-GuiTool 'Search'
        Set-GuiValue 'Target' @('Value')
        $s.ShowAdvanced.Checked = $false
        Assert-Equal '1|詳細な設定も表示する（1 件入れています）' ('{0}|{1}' -f $s.AdvancedCount, $s.ShowAdvanced.Text)
        Set-GuiValue 'Target' @()
        $s.ShowAdvanced.Checked = $true
        Assert-Equal '0|詳細な設定も表示する' ('{0}|{1}' -f $s.AdvancedCount, $s.ShowAdvanced.Text)

        # 2 つ落とすと旧・新に分けて入れる（版数の古い方が旧）
        Select-GuiTool 'Compare'
        $v1 = Join-Path $dir '設計書_v1.0.xlsx'
        $v2 = Join-Path $dir '設計書_v1.1.xlsx'
        foreach ($p in @($v1, $v2)) { Copy-Item -LiteralPath (Join-Path $script:Fixtures 'grep/sample.xlsx') -Destination $p }
        Add-GuiPaths @($v2, $v1)
        Assert-Equal "$v1|$v2" ('{0}|{1}' -f $s.Bindings[0].Input.Text, $s.Bindings[1].Input.Text)

        # フォルダに保存する処理は、そのフォルダに書いた結果のファイルを開けるようにする
        Select-GuiTool 'Markdown'
        Set-GuiValue 'Path' @((Join-Path $script:Fixtures 'grep/sample.xlsx'))
        $mdDir = Join-Path $dir 'md'
        Set-GuiValue 'OutDir' $mdDir
        Invoke-GuiRun
        Wait-GuiRun -TimeoutSeconds 120
        Assert-Equal (Join-Path $mdDir 'sample.md') $s.ResultPath ($s.Log.Text)
        Assert-Equal $mdDir $s.ResultDir

        # 途中で入力が要るツールは PowerShell の画面で実行する（テストでは起動しない）
        Select-GuiTool 'Analysis'
        Assert-Equal 'PowerShell の画面で実行' $s.RunButton.Text
        Set-GuiValue 'Path' $dir
        Invoke-GuiRun
        Assert-True (Test-Path -LiteralPath $s.LastConsoleScript) 'PowerShell の画面で実行するスクリプト'

        # 閉じるときに画面の大きさを覚える
        Save-GuiWindowBounds
        # 画面より大きい大きさは Windows が縮めるので、いまの大きさと比べる
        Assert-Equal ('{0},{1},{2},{3}' -f $form.Bounds.X, $form.Bounds.Y, $form.Bounds.Width, $form.Bounds.Height) $s.Settings.Window

        # やりたいことを言葉で探す
        Update-GuiTree '検索'
        $labels = @($s.Tree.Nodes | ForEach-Object { $_.Nodes } | ForEach-Object { $_.Nodes } | ForEach-Object { $_.Text })
        Assert-True ($labels -contains '文字を探す（検索）') ($labels -join ', ')

        # 探す欄の Enter は実行せず、見つかった処理を選ぶ
        $s.Find.Text = '提出前'
        $key = New-Object System.Windows.Forms.KeyEventArgs([System.Windows.Forms.Keys]::Enter)
        Invoke-GuiFindKey $key
        Assert-Equal 'Check|True' ('{0}|{1}' -f $s.Entry.Id, $key.SuppressKeyPress)
        Assert-True ($null -eq $s.Run) '実行しない'
        $s.Find.Text = '見つからない言葉'
        Assert-True (-not (Select-GuiFound)) '見つからなければ選び直さない'
        Assert-Equal 'Check' $s.Entry.Id
    }
    finally { $form.Dispose() }
}

Test-Case 'Start-ExcelTools: 1 つの数の欄は区切らずに読み（8,000 は 8000）、ほかのカンマ（7,5）は誤りとして知らせる。複数の数の欄は区切って読む' {
    $review = Get-TestEntry 'Review'
    $r = ConvertTo-GuiArguments $review (Get-GuiFields $review) @{ Path = @('a.xlsx'); OutDir = 'out'; MaxChars = '8,000' }
    Assert-Equal 0 $r.Errors.Count ($r.Errors -join ' / ')
    Assert-Equal '8000|Int32' ('{0}|{1}' -f $r.Arguments['MaxChars'], $r.Arguments['MaxChars'].GetType().Name) '3 桁ごとのカンマを除いて読む'
    $r = ConvertTo-GuiArguments $review (Get-GuiFields $review) @{ Path = @('a.xlsx'); OutDir = 'out'; MaxChars = '１２，０００' }
    Assert-Equal '12000' ([string]$r.Arguments['MaxChars']) '全角の数字とカンマも読む'
    $data = Get-TestEntry 'TestData'
    $r = ConvertTo-GuiArguments $data (Get-GuiFields $data) @{ Path = @('a.xlsx'); OutDir = 'out'; Rows = '1,000' }
    Assert-Equal '1000' ([string]$r.Arguments['Rows'])

    $effort = Get-TestEntry 'Effort'
    $r = ConvertTo-GuiArguments $effort (Get-GuiFields $effort) @{ Path = @('a.xlsx'); HoursPerDay = '7,5' }
    Assert-Equal @('「1 人日の時間」は数で入れてください') $r.Errors '小数点のつもりのカンマは、黙って 7 にしない'
    Assert-Equal @('HoursPerDay') $r.ErrorFields
    $r = ConvertTo-GuiArguments $effort (Get-GuiFields $effort) @{ Path = @('a.xlsx'); HoursPerDay = '1,234.5'; Tolerance = '8 000' }
    Assert-Equal '1234.5' ([string]$r.Arguments['HoursPerDay'])
    Assert-Equal @('Tolerance') $r.ErrorFields '空白で区切った 1 つの数の欄も誤り'

    $limits = Get-TestEntry 'Limits'
    $r = ConvertTo-GuiArguments $limits (Get-GuiFields $limits) @{ Path = @('a.csv'); Percentiles = '10, 90' }
    Assert-Equal @('10', '90') @($r.Arguments['Percentiles'] | ForEach-Object { [string]$_ }) '複数の数の欄は、カンマ・読点・空白で区切る'
}

Test-Case 'Start-ExcelTools: 画面が入れた結果の保存先のまま、もう一度実行すると、新しい名前にする（前の結果を上書きしない）' {
    $dir = New-TestDir
    $search = Get-TestEntry 'Search'
    $fields = Get-GuiFields $search
    $now = Get-Date '2026-10-23 10:15:00'
    $first = Get-GuiDefaultOutputs $search $fields @{ OutXlsx = '' } -Dir $dir -Now $now
    Assert-Equal '検索_20261023_101500.xlsx' (Split-Path -Leaf $first['OutXlsx'])
    [System.IO.File]::WriteAllText($first['OutXlsx'], 'x')
    # 1 回目のあと、結果の欄には画面が入れた名前が出ている
    $again = Get-GuiDefaultOutputs $search $fields @{ OutXlsx = $first['OutXlsx'] } -Dir $dir -Now $now -AutoFilled $first
    Assert-Equal '検索_20261023_101500(2).xlsx' (Split-Path -Leaf $again['OutXlsx'])
    $later = Get-GuiDefaultOutputs $search $fields @{ OutXlsx = $first['OutXlsx'] } -Dir $dir -Now $now.AddMinutes(5) -AutoFilled $first
    Assert-Equal '検索_20261023_102000.xlsx' (Split-Path -Leaf $later['OutXlsx'])
    # 自分で入れ直した保存先はそのまま
    $typed = Get-GuiDefaultOutputs $search $fields @{ OutXlsx = (Join-Path $dir '自分で決めた.xlsx') } -Dir $dir -Now $now -AutoFilled $first
    Assert-Equal 0 $typed.Count '入れ直した保存先'
}

Test-Case 'Start-ExcelTools: 同じことをするコマンド・スクリプトは、日付の引数と、Word などが書く右の引用符（U+2019）を含む文字でも、そのまま動く' {
    $dir = New-TestDir
    $tool = Join-Path $dir "Customer$([char]0x2019)s tool.ps1"
    [System.IO.File]::WriteAllText($tool, "param([datetime] `$AsOf, [string] `$Path, [string[]] `$Names)`r`n'{0}|{1}|{2}' -f `$AsOf.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture), `$Path, (`$Names -join '+')`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $path = "D:\資料\Customer$([char]0x2019)s 要件$([char]0x2018)案$([char]0x201A)$([char]0x201B).xlsx"
    $arguments = [ordered]@{ AsOf = (Get-Date '2026-10-20'); Path = $path; Names = [string[]]@("O'Brien", 'x') }
    $command = Format-GuiCommandLine $tool $arguments
    Assert-Equal "2026-10-20|$path|O'Brien+x" (Invoke-Expression $command) "コマンドをコピーで写すコマンド: $command"
    Assert-Equal '2026-10-20' (Invoke-Expression (ConvertTo-GuiLiteral ([datetime]'2026-10-20'))).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) '式でも日付'

    # PowerShell の画面で実行するスクリプト（AI レビュー・品質評価の定性分析）
    $script = New-GuiConsoleScript $tool $arguments '定性分析' -Dir $dir
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$null, [ref]$errors)
    Assert-Equal 0 @($errors).Count (@($errors | ForEach-Object { $_.Message }) -join ' / ')
    # スクリプトの引数の行を実行すると、元と同じ値になる
    $arguments = @{}
    foreach ($line in @([System.IO.File]::ReadAllLines($script) | Where-Object { $_ -like '$arguments`[*' })) { Invoke-Expression $line }
    Assert-Equal "2026-10-20|$path|O'Brien+x" ('{0}|{1}|{2}' -f $arguments['AsOf'].ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture), $arguments['Path'], ($arguments['Names'] -join '+'))
}

Test-Case 'Start-ExcelTools: 試しに実行（-WhatIf）では、雛形から作るファイルを記録の欄で確かめられる（ファイルは作らない）' {
    $dir = New-TestDir
    $tpl = Join-Path $dir '画面設計書_雛形.xlsx'
    New-SheetsBook $tpl ([ordered]@{ '画面概要' = @(@('項目', '値'), @('画面ID', '{{画面ID}}')) })
    $list = Join-Path $dir 'list.csv'
    @('画面ID', 'SCR001', 'SCR002') | Set-Content -LiteralPath $list -Encoding UTF8
    $out = Join-Path $dir 'out'
    $entry = Get-TestEntry 'NewDoc'
    $built = ConvertTo-GuiArguments $entry (Get-GuiFields $entry) @{ Template = $tpl; List = $list; FileName = '{{画面ID}}.xlsx'; OutDir = $out }
    Assert-Equal 0 $built.Errors.Count ($built.Errors -join ' / ')
    # 「試しに実行」を選んでいるときに画面が付ける引数
    $built.Arguments['WhatIf'] = $true
    $r = Wait-GuiToolRun (Start-GuiToolRun $entry.ToolPath $built.Arguments) -TimeoutSeconds 120
    Assert-True $r.Succeeded $r.Error
    Assert-Match ([string]$r.Summary.Text) 'SCR001\.xlsx\s+作成予定' '作るファイル'
    Assert-Match ([string]$r.Summary.Text) 'SCR002\.xlsx\s+作成予定'
    Assert-True (@($r.Messages | Where-Object { $_.Text -match '^2 ファイルを作る予定です' }).Count -eq 1) (($r.Messages | ForEach-Object { $_.Text }) -join ' / ')
    Assert-True (@($r.Messages | Where-Object { $_.Text -match '件ヒット' }).Count -eq 0) ('中で呼んだ検索の行は出さない: ' + (($r.Messages | ForEach-Object { $_.Text }) -join ' / '))
    Assert-True (-not (Test-Path -LiteralPath $out)) 'ファイルは作らない'
}

Test-Case 'Start-ExcelTools: Get-Help でヘルプが表示される' {
    $h = Get-Help $StartTool
    Assert-Match $h.Synopsis '画面'
    Assert-Equal 3 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
