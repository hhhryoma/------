#     Export-ExcelToMarkdown.ps1 のテスト。Excel は docs\design-sample と tests\fixtures\grep

$MarkdownTool = Join-Path $script:ToolsDir 'Export-ExcelToMarkdown.ps1'
$scr001 = Join-Path $script:RepoRoot 'docs/design-sample/DD02_画面設計書_SCR001_受注一覧.xlsx'

function Invoke-Markdown {
    param([hashtable] $Options)
    return ((@(& $MarkdownTool @Options 6>$null) -join "`n").Replace("`r`n", "`n"))
}

Test-Case 'Export-ExcelToMarkdown: 設計書を Markdown にする（表・「■」の小見出し・箇条書き）' {
    $md = Invoke-Markdown @{ Path = $scr001 }
    Assert-Match $md '^# DD02_画面設計書_SCR001_受注一覧\.xlsx\n'
    Assert-Match $md '(?m)^## 画面項目定義$'
    Assert-Match $md '(?m)^### ■ 画面項目定義$'
    Assert-Match $md '(?m)^\| No \| 項目名 \| 項目ID \| 区分 \| 入出力 \| 型 \| 桁数 \| 必須 \| 初期値 \| 入力チェック \| 対応項目 \| 備考 \|$'
    Assert-Match $md '(?m)^\| 1 \| 受注番号 \| orderNo \| テキスト \| I \| 半角英数 \| 12 \|  \|  \| 半角英数（E0003） \| T_ORDER\.ORDER_NO \| 前方一致 \|$'
    Assert-Match $md '(?m)^- 画面ID \| SCR001$' '表でない行は箇条書き'
    Assert-Match $md '(?m)^\| EV02 \| 検索 \|' 'イベント定義の表'
    Assert-NotMatch $md '(?m)^\| 行 \|' '既定ではセル番地を付けない'
}

Test-Case 'Export-ExcelToMarkdown: -CellAddress で行番号・セル番地を付け、図形・コメント・セル内の改行も書く' {
    $md = Invoke-Markdown @{ Path = (Join-Path $script:Fixtures 'grep/sample.xlsx'); CellAddress = $true }
    Assert-Match $md '(?m)^\| 行 \| 顧客ID \| 会社名 \|'
    Assert-Match $md '(?m)^\| 4 \|  \| 複数行<br>の文字列 TODO \|'
    Assert-Match $md '(?m)^- テキスト ボックス 1（G3 付近）: 注意：TODO 図形内<br>2行目$'
    Assert-Match $md '(?m)^- A2: 要確認：住所が古い$'
    Assert-Match $md '(?m)^## 作業用\n\n- A1: 隠しTODO$'
}

Test-Case 'Export-ExcelToMarkdown: -OutDir にフォルダ構成のまま保存し、変わっていない Excel は変換しない' {
    $src = New-TestDir
    New-Item -ItemType Directory -Path (Join-Path $src 'sub') | Out-Null
    Copy-Item -LiteralPath $scr001 -Destination (Join-Path $src '画面設計書.xlsx')
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'grep/sub/inline.xlsx') -Destination (Join-Path $src 'sub/inline.xlsx')
    $out = Join-Path (New-TestDir) 'md'
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out }
    Assert-Match $text '2 ファイルを Markdown にしました（Excel より新しい \.md があり変換しなかったもの 0 ファイル）'
    $mdFile = Join-Path $out '画面設計書.md'
    Assert-True (Test-Path -LiteralPath (Join-Path (Join-Path $out 'sub') 'inline.md')) 'サブフォルダ'
    $md = [System.IO.File]::ReadAllText($mdFile, [System.Text.Encoding]::UTF8)
    Assert-Match $md '^# 画面設計書\.xlsx\n'
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out }
    Assert-Match $text '0 ファイルを Markdown にしました（Excel より新しい \.md があり変換しなかったもの 2 ファイル）'
    (Get-Item -LiteralPath (Join-Path $src '画面設計書.xlsx')).LastWriteTimeUtc = [DateTime]::UtcNow.AddMinutes(5)
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out }
    Assert-Match $text '1 ファイルを Markdown にしました（Excel より新しい \.md があり変換しなかったもの 1 ファイル）'
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out -Force }
    Assert-Match $text '2 ファイルを Markdown にしました'
}

Test-Case 'Export-ExcelToMarkdown: [ ] を含む名前のファイルも Markdown にする（ワイルドカードとして読まない）' {
    $src = New-TestDir
    Copy-Item -LiteralPath $scr001 -Destination (Join-Path $src '画面設計書[旧].xlsx')
    $out = Join-Path (New-TestDir) 'md'
    $text = Get-HostText { & $MarkdownTool -Path (Join-Path $src '画面設計書[旧].xlsx') -OutDir $out 3>$null }
    Assert-Match $text '1 ファイルを Markdown にしました'
    Assert-True (Test-Path -LiteralPath (Join-Path $out '画面設計書[旧].md')) $text
}

Test-Case 'Export-ExcelToMarkdown: -TextConv は Markdown だけを UTF-8 で標準出力に書く（git の textconv 用）' {
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = (Get-Process -Id $PID).Path
    $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -TextConv "{1}"' -f $MarkdownTool, $scr001
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $ms = New-Object System.IO.MemoryStream
    $p.StandardOutput.BaseStream.CopyTo($ms)
    $err = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    Assert-Equal 0 $p.ExitCode $err
    $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($ms.ToArray())
    Assert-Match $text '^## 表紙\n' 'ファイル名の見出しは付けない（git が渡す一時ファイルの名前は毎回変わるため）'
    $expected = (Invoke-Markdown @{ Path = $scr001 }) -replace '^# [^\n]*\n\n', ''
    Assert-Equal $expected.TrimEnd() $text.TrimEnd()
}

Test-Case 'Export-ExcelToMarkdown: -OutDir で、拡張子だけが違う同じ名前の Excel は .md の名前に拡張子も含めて別々に保存する' {
    $src = New-TestDir
    Copy-Item -LiteralPath $scr001 -Destination (Join-Path $src '設計書.xlsx')
    Copy-Item -LiteralPath (Join-Path $script:Fixtures 'grep/sub/inline.xlsx') -Destination (Join-Path $src '設計書.xlsm')
    Copy-Item -LiteralPath $scr001 -Destination (Join-Path $src '画面設計書.xlsx')
    $out = Join-Path (New-TestDir) 'md'
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out }
    Assert-Match $text '3 ファイルを Markdown にしました（Excel より新しい \.md があり変換しなかったもの 0 ファイル）'
    Assert-Equal 3 @(Get-ChildItem -LiteralPath $out -File).Count '.md の数'
    foreach ($name in @('設計書.xlsx', '設計書.xlsm')) {
        $md = [System.IO.File]::ReadAllText((Join-Path $out ($name + '.md')), [System.Text.Encoding]::UTF8)
        Assert-Match $md ('^# ' + [regex]::Escape($name) + '\n')
    }
    Assert-True (Test-Path -LiteralPath (Join-Path $out '画面設計書.md')) '重ならない名前はそのまま'
    $text = Get-HostText { & $MarkdownTool -Path $src -OutDir $out }
    Assert-Match $text '0 ファイルを Markdown にしました（Excel より新しい \.md があり変換しなかったもの 3 ファイル）'
}

Test-Case 'Export-ExcelToMarkdown: 2 行の見出し（「予定」の下に「開始」「終了」）、No が続く表の途中の空の行、見出しの列より右の備考も表にする' {
    $file = Join-Path (New-TestDir) '表.xlsx'
    New-SheetsBook $file ([ordered]@{
            'WBS' = @(@('No', 'タスク', '予定', '', '実績', ''), @('', '', '開始', '終了', '開始', '終了'), @('1', '設計', '10/1', '10/5', '10/1', '10/6'), @('2', '製造', '10/6', '10/20', '', ''))
            '項目' = @(@('No', '項目名', '桁数', ''), @('1', '受注番号', '12', ''), @('', '', '', ''), @('2', '受注日', '8', '※要確認'), @('3', '顧客名', '40', ''), @('', '', '', ''), @('※ 桁数はバイト数', '', '', ''))
            '帳票' = @(@('画面設計書', 'システム名', '受注管理', '', '', '作成者', '山田'), @('', '画面ID', 'SCR001', '画面名', '受注一覧', '更新者', '山田'))
        })
    $md = Invoke-Markdown @{ Path = $file }
    Assert-Match $md '(?m)^## WBS\n\n\| No \| タスク \| 予定 開始 \| 予定 終了 \| 実績 開始 \| 実績 終了 \|\n\| --- \| --- \| --- \| --- \| --- \| --- \|\n\| 1 \| 設計 \| 10/1 \| 10/5 \| 10/1 \| 10/6 \|\n\| 2 \| 製造 \| 10/6 \| 10/20 \|  \|  \|\n' '2 行の見出し'
    Assert-Match $md '(?m)^\| No \| 項目名 \| 桁数 \|  \|\n\| --- \| --- \| --- \| --- \|\n\| 1 \| 受注番号 \| 12 \|  \|\n\| 2 \| 受注日 \| 8 \| ※要確認 \|\n\| 3 \| 顧客名 \| 40 \|  \|\n\n- ※ 桁数はバイト数$' '空の行と右の備考。表の下の注記は表に入れない'
    # 帳票の上の欄（2 行目の値が見出しの列の間にある）は、表にしない
    Assert-Match $md '(?m)^- 画面設計書 \| システム名 \| 受注管理 \| 作成者 \| 山田\n- 画面ID \| SCR001 \| 画面名 \| 受注一覧 \| 更新者 \| 山田$'
    $addr = Invoke-Markdown @{ Path = $file; CellAddress = $true }
    Assert-Match $addr '(?m)^\| 2 \| 1 \| 受注番号 \| 12 \|  \|\n\| 4 \| 2 \| 受注日 \| 8 \| ※要確認 \|$' '空の行をはさんだ行も、行番号はそのまま'
}

Test-Case 'Export-ExcelToMarkdown: 画像を名前・代替テキスト・大きさ・中身の目印で書く。-OutDir では .assets フォルダに画像を書き出してリンクし、無くなった画像は消す' {
    $dir = New-TestDir
    $file = Join-Path $dir '画面設計書.xlsx'
    New-SheetsBook $file ([ordered]@{ '画面イメージ' = @(@('■ 画面イメージ'), @('以上')) })
    Add-TestDrawing $file @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Screen'; Descr = '受注一覧の画面'; From = @(1, 2); To = @(6, 12) }, @{ Kind = 'pic'; Name = '図 2'; Png = 'Old'; From = @(7, 2); To = @(9, 4); Hidden = $true })
    $md = Invoke-Markdown @{ Path = $file }
    Assert-Match $md '(?m)^### 画像\n\n- 図 1: 受注一覧の画面（画像 400×300・[0-9a-f]{8}）\n- 図 2（画像 40×30・[0-9a-f]{8}・非表示）$'
    Assert-NotMatch $md '!\[' '-OutDir でなければ画像のファイルは書かない'
    Assert-Match (Invoke-Markdown @{ Path = $file; CellAddress = $true }) '(?m)^- 図 1（B3 付近）: 受注一覧の画面（画像 400×300・'

    # 画像を差し替えると、中身の目印が変わる（git diff に出る）
    $other = Join-Path $dir '差し替え.xlsx'
    New-SheetsBook $other ([ordered]@{ '画面イメージ' = @(@('■ 画面イメージ'), @('以上')) })
    Add-TestDrawing $other @(@{ Kind = 'pic'; Name = '図 1'; Png = 'Screen'; Descr = '受注一覧の画面'; From = @(1, 2); To = @(6, 12) }, @{ Kind = 'pic'; Name = '図 2'; Png = 'New'; From = @(7, 2); To = @(9, 4); Hidden = $true })
    $mark = { param($text) [regex]::Match($text, '図 2（画像 40×30・([0-9a-f]{8})').Groups[1].Value }
    Assert-True ((& $mark $md) -ne (& $mark (Invoke-Markdown @{ Path = $other }))) '差し替えた画像は目印が違う'

    $out = Join-Path $dir 'md'
    $null = & $MarkdownTool -Path $file -OutDir $out 6>$null
    $saved = [System.IO.File]::ReadAllText((Join-Path $out '画面設計書.md'))
    Assert-Match $saved '(?m)^- 図 1: 受注一覧の画面（画像 400×300・[0-9a-f]{8}）\n  !\[図 1\]\(<画面設計書\.assets/image1_1\.png>\)$'
    Assert-Equal @('image1_1.png', 'image1_2.png') @(Get-ChildItem -LiteralPath (Join-Path $out '画面設計書.assets') | ForEach-Object { $_.Name } | Sort-Object)
    # 画像を消したブックを変換し直すと、書き出した画像とフォルダも消す
    New-SheetsBook $file ([ordered]@{ '画面イメージ' = @(@('■ 画面イメージ'), @('以上')) })
    $null = & $MarkdownTool -Path $file -OutDir $out -Force 6>$null
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $out '画面設計書.assets'))) '.assets フォルダを消す'
    Assert-NotMatch ([System.IO.File]::ReadAllText((Join-Path $out '画面設計書.md'))) '### 画像'
}

Test-Case 'Export-ExcelToMarkdown: Get-Help でヘルプが表示される' {
    $h = Get-Help $MarkdownTool
    Assert-Match $h.Synopsis 'Markdown'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
