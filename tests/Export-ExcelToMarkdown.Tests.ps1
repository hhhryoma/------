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

Test-Case 'Export-ExcelToMarkdown: Get-Help でヘルプが表示される' {
    $h = Get-Help $MarkdownTool
    Assert-Match $h.Synopsis 'Markdown'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
