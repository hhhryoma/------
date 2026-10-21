#     Invoke-DesignReview.ps1 のテスト。設計書は docs\design-sample、観点表は同梱のものを使う

$Review = Join-Path $script:ToolsDir 'Invoke-DesignReview.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'
$brokenChecklist = Join-Path $script:Fixtures 'review/broken-checklist.xlsx'

function Get-Manifest {
    param([string] $OutDir)
    $json = [System.IO.File]::ReadAllText((Join-Path $OutDir 'requests.json'))
    return ($json | ConvertFrom-Json)
}

function Write-Answer {
    param([string] $OutDir, [string] $Id, [string] $Text)
    [System.IO.File]::WriteAllText((Join-Path (Join-Path $OutDir 'answers') "$Id.txt"), $Text, (New-Object System.Text.UTF8Encoding($true)))
}

# Copilot の回答の例（Markdown の表・コードブロック・タブ区切り）
$answerTable = @'
以下がレビュー結果です。

| 観点ID | 判定 | 場所 | 指摘内容 | 根拠 |
|---|---|---|---|---|
| C01 | OK | - | - | - |
| C02 | OK | - | - | - |
| C03 | NG | イベント定義!J9 | 「押下」と「クリック」が混在 | 用語の統一 |
| S01 | **NG** | 画面項目定義!M11 | 受注ステータスの型が未定義 | 入力項目には型が必要 |
| S02 | OK | - | - | - |
| S03 | OK | - | - | - |
| S04 | OK | - | - | - |
| S05 | OK | - | - | - |
| S06 | 要確認 | イベント定義!P8 | クリア時の動きが不明 | 異常時の記載 |
| S07 | OK | - | - | - |
| S08 | OK | - | - | - |
| S09 | OK | - | - | - |
| S10 | OK | - | - | - |
'@
$answerCode = @'
```
C01 | OK | - | - | -
C02 | NG | イベント定義!P11 | 連番の採番単位が不明 | 解釈が分かれる
S01 | OK | - | - | -
S04 | NG | 存在しないシート!A1 | 行追加ボタンの非活性条件 | テスト
S05 | ○ | - | - | -
```
'@
$answerTsv = "C01`tOK`t-`t-`t-`nC02`tOK`t-`t-`t-`nC03`tOK`t-`t-`t-`nT01`tOK`t-`t-`t-`nT02`tOK`t-`t-`t-`nT03`t要確認`tT_ORDER!AH16`t合計金額のデフォルトの意図`t不明`nT04`tOK`t-`t-`t-`nT05`tOK`t-`t-`t-`nT06`tNG`tM_CUSTOMER`t検索列にインデックスがない`t検索条件の列`n"

Test-Case 'Invoke-DesignReview: 同梱の観点表はチェックで問題なし' {
    $out = Get-HostText { & $Review -Path $designDir -CheckOnly }
    Assert-Match $out '有効 27 観点'
    Assert-Match $out '問題ありません'
}

Test-Case 'Invoke-DesignReview: 観点表の誤りを行番号付きで検出する' {
    $out = Get-HostText { & $Review -Path $designDir -Checklist $brokenChecklist -CheckOnly }
    Assert-Match $out 'エラー 4 件 / 警告 7 件'
    Assert-Match $out '\[エラー\] 3行目 s01 観点IDが重複しています（3行目、5行目）'
    Assert-Match $out '\[エラー\] 4行目 C03 観点が空欄です'
    Assert-Match $out '\[エラー\] 10行目 S 08 観点IDに空白'
    Assert-Match $out '\[警告\] 6行目 S02 判定基準が空欄です'
    Assert-Match $out '\[警告\] 7行目 S03 判定基準にあいまいな言葉があります（適切）'
    Assert-Match $out '\[警告\] 8行目 S04 重要度「最高」'
    Assert-Match $out '\[警告\] 9行目 S05 「有効」の値「はい」'
    Assert-Match $out '\[警告\] 11行目 S07 対象設計書「\*存在しない設計書\*」'
    Assert-Match $out '\[警告\] 12行目 S08 参照設計書「\*存在しない参照\*」'
    Assert-Match $out '\[警告\] 13行目 S09 観点・判定基準に「\|」'
    Assert-NotMatch $out '14行目' '無効（×）の行の重複はエラーにしない'
    Assert-NotMatch $out '必要に応じて' '（例: …）の中は対象外'
}

Test-Case 'Invoke-DesignReview: 観点表にエラーがあればプロンプトを作らずに止まる' {
    $dir = New-TestDir
    Assert-Throws { & $Review -Path $designDir -Checklist $brokenChecklist -OutDir $dir -PromptOnly 6>$null } '観点表にエラー'
    Assert-True (-not (Test-Path (Join-Path $dir 'requests.json'))) '依頼が作られています'
}

Test-Case 'Invoke-DesignReview: 共通の観点しか当たらない設計書を知らせる' {
    # ファイル名に種類が入っていない設計書（ID だけの名前）を混ぜる
    $dir = Copy-Fixture $designDir
    Copy-Item -LiteralPath (Join-Path $dir 'DD04_コード定義書.xlsx') -Destination (Join-Path $dir 'BA-03-02-0001.xlsx')
    $out = Get-HostText { & $Review -Path $dir -CheckOnly }
    Assert-Match $out '\[情報\] 設計書「BA-03-02-0001\.xlsx」には共通の観点（対象設計書 \*）しか当たりません'
    Assert-Match $out 'エラー 0 件 / 警告 0 件 / 情報 1 件'
}

Test-Case 'Invoke-DesignReview: -PromptOnly で上限文字数に収まる依頼を作る' {
    foreach ($max in 4000, 8000, 16000) {
        $dir = New-TestDir
        $null = & $Review -Path $designDir -OutDir $dir -PromptOnly -MaxChars $max 6>$null
        $m = Get-Manifest $dir
        $over = @($m.Requests | Where-Object { $_.Chars -gt $max })
        Assert-Equal 0 $over.Count "上限 $max を超える依頼"
        $files = @(Get-ChildItem -LiteralPath (Join-Path $dir 'prompts') -Filter 'R*.txt')
        Assert-Equal @($m.Requests).Count $files.Count 'プロンプトのファイル数'
        if ($max -eq 8000) { Assert-Equal 10 @($m.Requests).Count '8000 文字での依頼数' }
        if ($max -eq 16000) { Assert-Equal 8 @($m.Requests).Count '16000 文字なら 1 設計書 1 依頼' }
    }
}

Test-Case 'Invoke-DesignReview: プロンプトにセル番地付きの設計書と参照資料が入る' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null
    $p = [System.IO.File]::ReadAllText((Join-Path $dir 'prompts/R002.txt'))
    Assert-Match $p '【設計書】DD02_画面設計書_SCR001_受注一覧\.xlsx'
    Assert-Match $p '### シート: 画面項目定義'
    Assert-Match $p '\n6\| A:1 \| C:受注番号 \| I:orderNo'
    Assert-Match $p '【参照資料】DD05_メッセージ一覧\.xlsx'
    Assert-Match $p '1～3行目の帳票ヘッダは前のシートと同じため省略'
    # 改行は OS によらず LF（Windows で CR の分だけ文字数が増え、分割数が変わるのを防ぐ）
    Assert-True (-not $p.Contains("`r")) 'プロンプトに CR が含まれています'
}

Test-Case 'Invoke-DesignReview: -TextFormat Markdown で、表に行番号と列を付けた Markdown の設計書を渡す' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly -TextFormat Markdown 6>$null
    $p = [System.IO.File]::ReadAllText((Join-Path $dir 'prompts/R002.txt'))
    Assert-Match $p '【設計書】DD02_画面設計書_SCR001_受注一覧\.xlsx'
    Assert-Match $p '箇条書きの「K12: 必須」はセル K12 です。表は 1 列目が行番号で、見出しの「K:桁数」の K が列です'
    Assert-NotMatch $p '「12\| K:必須」'
    Assert-Match $p '\n## 画面項目定義\n\n（1～3行目の帳票ヘッダは前のシートと同じため省略）\n\n### ■ 画面項目定義\n'
    Assert-Match $p '\n\| 行 \| A:No \| C:項目名 \| I:項目ID \| P:区分 \|'
    Assert-Match $p '\n\| 6 \| 1 \| 受注番号 \| orderNo \| テキスト \| I \| 半角英数 \| 12 \|  \|  \| 半角英数（E0003） \|'
    Assert-Match $p '\n- Q22: 文書番号 \| X22: ODR-DD-02-01\n'
    $all = (Get-ChildItem -LiteralPath (Join-Path $dir 'prompts') -Filter 'R*.txt' | ForEach-Object { [System.IO.File]::ReadAllText($_.FullName) }) -join "`n"
    Assert-Match $all '【参照資料】DD05_メッセージ一覧\.xlsx（判定の参考。指摘の対象外）\n## 表紙\n'
    Assert-True (-not $p.Contains("`r")) 'プロンプトに CR が含まれています'

    # 長いシートを分けたときは、表の続きに表の見出しを付け直す。上限も守る
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly -TextFormat Markdown -MaxChars 4000 6>$null
    $m = Get-Manifest $dir
    Assert-Equal 0 @($m.Requests | Where-Object { $_.Chars -gt 4000 }).Count '上限 4000 を超える依頼'
    $continued = @(Get-ChildItem -LiteralPath (Join-Path $dir 'prompts') -Filter 'R*.txt' | ForEach-Object { [System.IO.File]::ReadAllText($_.FullName) } | Where-Object { $_.Contains('## 画面項目定義（続き）') })
    Assert-True ($continued.Count -gt 0) '表の途中で分けた依頼がない'
    foreach ($c in $continued) { Assert-Match $c '## 画面項目定義（続き）\n\| 行 \| A:No \| C:項目名 [^\n]*\|\n\| --- \|' }
}

Test-Case 'Invoke-DesignReview: 3 形式の回答を取り込んでレビュー結果を作る' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null
    Write-Answer $dir 'R002' $answerTable
    Write-Answer $dir 'R003' $answerCode
    Write-Answer $dir 'R007' $answerTsv
    $null = & $Review -OutDir $dir -ImportOnly 6>$null

    $report = Join-Path $dir 'レビュー結果.xlsx'
    $f = Read-XlsxSheet -Path $report -Index 1
    Assert-Equal @('指摘一覧', '観点消化状況', '依頼一覧') $f.Sheets
    Assert-Equal 8 $f.Rows.Count '見出し＋指摘 7 件'
    # 設計書ごとに重要度の高い順
    Assert-Equal @('S01', 'S06', 'C03', 'C02', 'S04', 'T03', 'T06') @(2..8 | ForEach-Object { $f.Rows[$_]['D'] })
    Assert-Equal 'NG' $f.Rows[2]['B'] '**NG** は NG'
    Assert-Equal 6 $f.Links.Count '存在しないシートの場所にはリンクを付けない'
    Assert-True (@($f.Links | Where-Object { $_.Location -eq "'M_CUSTOMER'!A1" }).Count -eq 1) 'シート名だけの場所は A1'
    Assert-Match $f.Rows[6]['P'] '場所のシート名・セル番地を設計書で確認'

    $s = Read-XlsxSheet -Path $report -Index 2
    $verdicts = @($s.Rows.Keys | Where-Object { $_ -gt 1 } | ForEach-Object { $s.Rows[$_]['F'] })
    Assert-Equal 5 @($verdicts | Where-Object { $_ -eq 'NG' }).Count 'NG'
    Assert-Equal 2 @($verdicts | Where-Object { $_ -eq '要確認' }).Count '要確認'
    Assert-Equal 20 @($verdicts | Where-Object { $_ -eq 'OK' }).Count 'OK（○ も OK）'
    Assert-Equal 41 @($verdicts | Where-Object { $_ -eq '未回答' }).Count '未回答'
}

Test-Case 'Invoke-DesignReview: 設計書・観点が変わった依頼の古い回答は退避する' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null
    Write-Answer $dir 'R001' "C01 | OK | - | - | -"
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null 3>$null
    Assert-True (Test-Path (Join-Path $dir 'answers/R001.txt')) '内容が同じなら回答はそのまま'
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly -SheetNameExclude '表紙' 6>$null 3>$null
    Assert-True (Test-Path (Join-Path $dir 'answers/R001.txt.old')) '内容が変われば退避'
    Assert-True (-not (Test-Path (Join-Path $dir 'answers/R001.txt'))) '古い回答は使わない'
}

Test-Case 'Invoke-DesignReview: 対話モードでやり直し・中断・再開ができる' {
    $dir = New-TestDir
    # クリップボードと入力をテスト用に差し替える（ツールからは関数が優先して見える）。
    # 関数内の $script: はツール側のスコープを指すので、状態はハッシュテーブルで持つ
    $state = @{ Clip = ''; Step = 0 }
    function Set-Clipboard { param($Value) $state.Clip = $Value }
    function Read-Host { param($Prompt) $state.Step++; if ($state.Step -eq 6) { return 'q' }; return '' }
    function Get-Clipboard {
        param([switch] $Raw)
        if ($state.Step -eq 1) { return $state.Clip }                       # 依頼のまま → やり直し
        if ($state.Step -eq 2) { return 'うまく回答できませんでした' }       # 読めない → やり直し
        $ids = [regex]::Matches($state.Clip, '(?m)^([A-Z]\d{2}) \| ') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
        return (($ids | ForEach-Object { "$_ | OK | - | - | -" }) -join "`n")
    }

    $out = Get-HostText { & $Review -Path $designDir -OutDir $dir }
    Assert-Match $out 'クリップボードの内容が依頼のまま'
    Assert-Match $out '読み取れませんでした'
    Assert-Match $out '中断しました'
    Assert-Equal 3 @(Get-ChildItem -LiteralPath (Join-Path $dir 'answers') -Filter 'R*.txt').Count '中断までに 3 件回答'

    $state.Step = 100
    $out = Get-HostText { & $Review -Path $designDir -OutDir $dir }
    Assert-Match $out '未回答の依頼 7 件'
    Assert-Match $out '依頼 10 件中 10 件回答済み'
}

Test-Case 'Invoke-DesignReview: Get-Help でヘルプが表示される' {
    $h = Get-Help $Review
    Assert-Match $h.Synopsis 'AI'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
