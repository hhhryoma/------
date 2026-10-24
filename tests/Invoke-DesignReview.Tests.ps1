#     Invoke-DesignReview.ps1 のテスト。設計書は docs\design-sample、観点表は同梱のものを使う

$Review = Join-Path $script:ToolsDir 'Invoke-DesignReview.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'
$brokenChecklist = Join-Path $script:Fixtures 'review/broken-checklist.xlsx'
$consistencyDir = Join-Path $script:Fixtures 'consistency'

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
S08 | NG | 存在しないシート!A1 | 使用テーブルの記載 | テスト
S05 | ○ | - | - | -
```
'@
$answerTsv = "C01`tOK`t-`t-`t-`nC02`tOK`t-`t-`t-`nC03`tOK`t-`t-`t-`nT01`tOK`t-`t-`t-`nT02`tOK`t-`t-`t-`nT03`t要確認`tT_ORDER!AH16`t合計金額のデフォルトの意図`t不明`nT04`tOK`t-`t-`t-`nT05`tOK`t-`t-`t-`nT06`tNG`tM_CUSTOMER`t検索列にインデックスがない`t検索条件の列`n"

Test-Case 'Invoke-DesignReview: 同梱の観点表はチェックで問題なし' {
    $out = Get-HostText { & $Review -Path $designDir -CheckOnly }
    Assert-Match $out '有効 28 観点'
    Assert-Match $out '確かめ方: AI 14 観点・ツール 7 観点・両方 7 観点'
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
    Assert-Match $p '\n観点ID \| 判定 \| 場所 \| 引用 \| 指摘内容 \| 根拠\n'
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
    Assert-Equal @('指摘一覧', '観点消化状況', '依頼一覧', '観点別の採用率') $f.Sheets
    Assert-Equal 11 $f.Rows.Count '見出し＋指摘 10 件（AI 7 件・ツール 3 件）'
    # 設計書ごとに重要度の高い順。ツールの指摘（同梱の観点表では初期値の空欄が 3 か所）も入る
    Assert-Equal @('S01', 'S06', 'C03', 'C02', 'S08', 'S03', 'S03', 'S03', 'T03', 'T06') @(2..11 | ForEach-Object { $f.Rows[$_]['E'] })
    Assert-Equal @('AI', 'AI', 'AI', 'AI', 'AI', 'ツール（初期値）', 'ツール（初期値）', 'ツール（初期値）', 'AI', 'AI') @(2..11 | ForEach-Object { $f.Rows[$_]['C'] })
    Assert-Equal 'NG' $f.Rows[2]['B'] '**NG** は NG'
    Assert-Equal 9 $f.Links.Count '存在しないシートの場所にはリンクを付けない'
    Assert-True (@($f.Links | Where-Object { $_.Location -eq "'M_CUSTOMER'!A1" }).Count -eq 1) 'シート名だけの場所は A1'
    Assert-Match $f.Rows[6]['S'] '場所のシート名・セル番地を設計書で確認'
    Assert-Equal '引用なし' $f.Rows[2]['K'] '引用の列の無い回答（5 列）も読む'

    $s = Read-XlsxSheet -Path $report -Index 2
    $verdicts = @($s.Rows.Keys | Where-Object { $_ -gt 1 } | ForEach-Object { $s.Rows[$_]['F'] })
    Assert-Equal 7 @($verdicts | Where-Object { $_ -eq 'NG' }).Count 'NG'
    Assert-Equal 2 @($verdicts | Where-Object { $_ -eq '要確認' }).Count '要確認'
    Assert-Equal 34 @($verdicts | Where-Object { $_ -eq 'OK' }).Count 'OK（○ も OK。ツールだけの観点も）'
    Assert-Equal 33 @($verdicts | Where-Object { $_ -eq '未回答' }).Count '未回答（両方の観点は、AI の回答が無ければ未回答）'
}

Test-Case 'Invoke-DesignReview: 確かめ方がツールの観点は AI に依頼せず、ツールで見つけた指摘を判定元「ツール（確認の名前）」で入れる' {
    $dir = Copy-Fixture $designDir
    $scr1 = Join-Path $dir 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $scr2 = Join-Path $dir 'DD02_画面設計書_SCR002_受注登録.xlsx'
    # 表紙の版数を改訂履歴と違えて、消し忘れの TBD と長音の無い「ユーザ」を書く
    Set-CellValues $scr1 '表紙' @{ X23 = '1.1' }
    Set-CellValues $scr1 '画面概要' @{ I9 = '受注担当者（ログインユーザ）。TBD' }
    # 型・桁数の空欄、メッセージID の無い入力チェック、イベントの無いボタン、定義されていないメッセージ（希望納品日の初期値は「なし」にする）
    Set-CellValues $scr2 '画面項目定義' @{ X7 = $null; AB12 = $null; AK18 = '必須（E0001）、範囲（1〜999）'; AF10 = 'なし'; C21 = 'CSV出力'; AK12 = '桁数（E0099）' }
    # 処理概要の E0008 をメッセージ欄から消す
    Set-CellValues $scr2 'イベント定義' @{ AU11 = 'Q0001, E0001, E0002, E0003, E0004, E0005, E0007, E0009, E0010, I0001' }
    # 主キーと、CHAR の桁数・NUMERIC の小数を消す
    Set-CellValues (Join-Path $dir 'DD03_テーブル定義書.xlsx') 'T_ORDER' @{ AC13 = $null; X15 = $null; AA20 = $null }

    $out = Join-Path $dir 'review'
    $null = & $Review -Path $dir -OutDir $out -PromptOnly 6>$null 3>$null
    $prompt = [System.IO.File]::ReadAllText((Join-Path $out 'prompts/R002.txt'))
    Assert-NotMatch $prompt '\nS02 \|' 'ツールだけの観点は AI に依頼しない'
    Assert-Match $prompt '\nC01 \| 表紙と改訂履歴が記入されている \| [^\n]*（表紙・改訂履歴・ファイル名の版数と更新日の食い違いは、ツールで確かめるので指摘しなくてよい）\n'

    # 依頼の一覧に設計書の一覧があるので、-ImportOnly だけでもツールで確かめる
    $text = Get-HostText { & $Review -OutDir $out -ImportOnly }
    Assert-Match $text 'ツールの確認: 指摘 14 件'
    $f = Read-XlsxSheet -Path (Join-Path $out 'レビュー結果.xlsx') -Index 1
    Assert-Equal 'No|判定|判定元|重要度|観点ID' ((@('A', 'B', 'C', 'D', 'E') | ForEach-Object { $f.Rows[1][$_] }) -join '|')
    $scr1Name = 'DD02_画面設計書_SCR001_受注一覧.xlsx'
    $scr2Name = 'DD02_画面設計書_SCR002_受注登録.xlsx'
    Assert-Equal @(
        "ツール（残っている文言）|C02|$scr1Name|画面概要!I9",
        "ツール（版数）|C01|$scr1Name|表紙!X23",
        "ツール（表記ゆれ）|C03|$scr1Name|画面概要!I9",
        "ツール（参照先なし）|C04|$scr2Name|画面項目定義!AK12",
        "ツール（入力項目の型・桁数）|S01|$scr2Name|画面項目定義!X7",
        "ツール（入力項目の型・桁数）|S01|$scr2Name|画面項目定義!AB12",
        "ツール（ボタンのイベント）|S04|$scr2Name|画面項目定義!C21",
        "ツール（入力チェックのメッセージ）|S02|$scr2Name|画面項目定義!AK18",
        "ツール（メッセージ欄）|S07|$scr2Name|イベント定義!AU11",
        'ツール（初期値）|S03|DD02_画面設計書_SCR003_受注詳細.xlsx|画面項目定義!AF12',
        'ツール（初期値）|S03|DD02_画面設計書_SCR003_受注詳細.xlsx|画面項目定義!AF13',
        'ツール（主キー）|T01|DD03_テーブル定義書.xlsx|T_ORDER!AC12',
        'ツール（カラムの型・桁数）|T02|DD03_テーブル定義書.xlsx|T_ORDER!X15',
        'ツール（カラムの型・桁数）|T02|DD03_テーブル定義書.xlsx|T_ORDER!AA20') @(2..15 | ForEach-Object { $r = $f.Rows[$_]; '{0}|{1}|{2}|{3}' -f $r['C'], $r['E'], $r['H'], $r['I'] })
    $details = @(2..15 | ForEach-Object { $f.Rows[$_]['L'] })
    foreach ($expected in @(
            '表紙の版数（1.1）と、改訂履歴の最後の版数（1.0）が違います',
            '「ログインユーザ」は「ログインユーザー」の表記ゆれです（長音（ー）。このファイルで 1 か所）',
            'メッセージ「E0099」が定義されていません',
            '入力項目「顧客コード」の型が空欄です',
            '入力項目「備考」（型: 文字）の桁数が空欄です',
            'ボタン「CSV出力」のイベントが、イベント定義にありません',
            '入力項目「数量」の入力チェック「範囲(1〜999)」に、エラーのときのメッセージIDがありません',
            'イベント「登録」の処理概要にあるメッセージ E0008 が、メッセージ欄にありません',
            'テーブル T_ORDER に主キーがありません（PK の欄がどのカラムも空欄です）',
            'T_ORDER.CUSTOMER_CD（CHAR）の桁数が空欄です',
            'T_ORDER.TOTAL_AMOUNT（NUMERIC）の小数の桁数が空欄です')) {
        Assert-True ($details -contains $expected) ("指摘がありません: {0}`n実際: {1}" -f $expected, ($details -join "`n"))
    }
    Assert-Equal '機械のチェック（版数）' $f.Rows[3]['M']
    Assert-Equal 'ツール' $f.Rows[3]['N'] '依頼No'
    Assert-Equal '' ([string]$f.Rows[3]['K']) 'ツールの指摘は引用を確かめない'
    Assert-True (@($f.Links | Where-Object { $_.Location -eq "'T_ORDER'!AA20" }).Count -eq 1) 'ツールの指摘の場所へのリンク'

    $s = Read-XlsxSheet -Path (Join-Path $out 'レビュー結果.xlsx') -Index 2
    $rows = @($s.Rows.Keys | Where-Object { $_ -gt 1 } | ForEach-Object { $r = $s.Rows[$_]; '{0}|{1}|{2}|{3}|{4}' -f $r['A'], $r['B'], $r['F'], $r['G'], $r['H'] })
    Assert-True ($rows -contains "$scr2Name|S02|NG|1|ツール") 'ツールだけの観点は、ツールの指摘で NG'
    Assert-True ($rows -contains "$scr2Name|S04|NG|1|ツール")
    Assert-True ($rows -contains 'DD05_メッセージ一覧.xlsx|C04|OK|0|ツール') 'ツールの指摘が無ければ OK'
    Assert-True ($rows -contains "$scr1Name|C01|NG|1|R002, ツール") '両方の観点は、AI が未回答でもツールの指摘で NG'
    Assert-True ($rows -contains "$scr2Name|C01|未回答|0|R003, ツール") '両方の観点は、ツールの指摘が無くても AI が未回答なら未回答'
}

Test-Case 'Invoke-DesignReview: ツールの確認に（ ）で種類を付けると、その種類の食い違いだけを入れる。ツールだけの観点表は依頼を作らない' {
    $dir = New-TestDir
    $csv = Join-Path $dir '観点表.csv'
    @('観点ID,分類,観点,判定基準,重要度,対象設計書,有効,確かめ方,ツールの確認',
        'X01,整合性,メッセージの参照,,高,*,○,ツール,参照先なし（メッセージ）',
        'X02,画面一覧,画面一覧との違い,,中,*画面一覧*,○,ツール,一覧との違い（画面）') | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Join-Path $dir 'review'
    $text = Get-HostText { & $Review -Path $consistencyDir -Checklist $csv -OutDir $out -PromptOnly }
    Assert-Match $text 'AI に依頼する観点はありません'
    $null = & $Review -Checklist $csv -OutDir $out -ImportOnly 6>$null 3>$null
    $f = Read-XlsxSheet -Path (Join-Path $out 'レビュー結果.xlsx') -Index 1
    # テーブル・カラム・コードの参照先なしと、テーブル一覧との違いは入れない
    Assert-Equal @(
        'ツール（一覧との違い）|X02|DD01_画面一覧.xlsx|画面一覧!C8',
        'ツール（一覧との違い）|X02|DD01_画面一覧.xlsx|画面一覧!C9',
        'ツール（参照先なし）|X01|DD02_画面設計書_SCR001_受注一覧.xlsx|画面項目定義!AK9') @(2..($f.Rows.Count) | ForEach-Object { $r = $f.Rows[$_]; '{0}|{1}|{2}|{3}' -f $r['C'], $r['E'], $r['H'], $r['I'] })
}

Test-Case 'Invoke-DesignReview: 観点表の確かめ方・ツールの確認の誤りを知らせる（ツールだけの観点は判定基準を調べない）' {
    $dir = New-TestDir
    $csv = Join-Path $dir '観点表.csv'
    @('観点ID,分類,観点,判定基準,重要度,対象設計書,有効,確かめ方,ツールの確認',
        'V01,画面項目,ツールの確認の無い観点,ツールで確かめる,中,*画面設計書*,○,ツール,',
        'V02,画面項目,名前の誤り,ツールで確かめる,中,*画面設計書*,○,ツール,初期値、存在しない確認',
        'V03,画面項目,種類を付けられない確認,ツールで確かめる,中,*画面設計書*,○,ツール,版数（画面）',
        'V04,画面項目,種類の誤り,ツールで確かめる,中,*画面設計書*,○,両方,参照先なし（存在しない種類）',
        'V05,画面項目,確かめ方の誤り,桁数が書かれている,中,*画面設計書*,○,tool,初期値',
        'V06,画面項目,使わないツールの確認,桁数が書かれている,中,*画面設計書*,○,AI,初期値',
        'V07,画面項目,ツールだけの観点,適切に設定されている,中,*画面設計書*,○,ツール,初期値',
        'V08,画面項目,種類を付けた確認,桁数が書かれている,中,*画面設計書*,○,ＡＩ,',
        'V09,画面項目,種類を付けた確認,桁数が書かれている,中,*画面設計書*,○,,参照先なし（メッセージ）') | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Get-HostText { & $Review -Checklist $csv -CheckOnly }
    Assert-Match $out '確かめ方: AI 2 観点・ツール 4 観点・両方 3 観点' 'ツールの確認があって確かめ方が空欄・誤りなら両方'
    Assert-Match $out 'エラー 4 件 / 警告 2 件'
    Assert-Match $out '\[エラー\] 2行目 V01 確かめ方が「ツール」ですが、「ツールの確認」が空欄です'
    Assert-Match $out '\[エラー\] 3行目 V02 ツールの確認「存在しない確認」はありません（使える確認: 版数、残っている文言、'
    Assert-Match $out '\[エラー\] 4行目 V03 ツールの確認「版数（画面）」には、（ ）で種類を付けられません'
    Assert-Match $out '\[エラー\] 5行目 V04 ツールの確認「参照先なし（存在しない種類）」の種類「存在しない種類」は、整合性チェックの参照のルールにありません（種類: テーブル、カラム、'
    Assert-Match $out '\[警告\] 6行目 V05 「確かめ方」の値「tool」は AI・ツール・両方 のどれかにしてください（両方として扱います）'
    Assert-Match $out '\[警告\] 7行目 V06 確かめ方が「AI」のため、「ツールの確認」は使いません'
    Assert-NotMatch $out 'V07' 'ツールだけの観点の判定基準は AI に渡さないので、あいまいな言葉を調べない'
    Assert-Throws { & $Review -Path $designDir -Checklist $csv -OutDir (New-TestDir) -PromptOnly 6>$null 3>$null } '観点表にエラー'
}

Test-Case 'Invoke-DesignReview: 回答の引用を設計書のセルと突き合わせ、場所が違えば引用のあるセルに直し、見つからない引用を知らせる' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null
    # R002 は SCR001（受注一覧）
    Write-Answer $dir 'R002' @'
| 観点ID | 判定 | 場所 | 引用 | 指摘内容 | 根拠 |
|---|---|---|---|---|---|
| C01 | OK | - | - | - | - |
| C02 | NG | イベント定義!P11 | 採番単位が不明 | 解釈が分かれる |
| C03 | NG | イベント定義!J9 | 「新規登録ボタン押下」 | 「押下」と「クリック」が混在 | 用語の統一 |
| S01 | NG | 画面項目定義!AB11 | （空欄） | 受注ステータスの桁数が空欄 | 入力項目には桁数が必要 |
| S05 | NG | 画面概要!I8 | 受注を削除する | 削除の確認が無い | 更新系 |
| S06 | 要確認 | イベント定義!P7 | 検索条件を初期表示の状態に戻し | クリア時の動きが不明 | 異常時の記載 |
| S08 | NG | 存在しないシート!A1 | 前回と同じ検索条件で | 使用テーブルの記載が無い | テスト |
| S10 | NG | 画面項目定義!AK6 | （空欄） | メッセージが無い | テスト |
'@
    $text = Get-HostText { & $Review -OutDir $dir -ImportOnly }
    Assert-Match $text 'AI の指摘の引用: 引用なし 1・一致 2・見つからない 2・場所を直した 2'
    $f = Read-XlsxSheet -Path (Join-Path $dir 'レビュー結果.xlsx') -Index 1
    Assert-Equal '場所|引用|引用の確認|指摘内容' ((@('I', 'J', 'K', 'L') | ForEach-Object { $f.Rows[1][$_] }) -join '|')
    $rows = @{}
    foreach ($r in @($f.Rows.Keys | Where-Object { $_ -gt 1 })) { if ($f.Rows[$r]['C'] -eq 'AI') { $rows[$f.Rows[$r]['E']] = $f.Rows[$r] } }
    $show = { param($id) '{0}|{1}|{2}|{3}' -f $rows[$id]['I'], $rows[$id]['J'], $rows[$id]['K'], $rows[$id]['S'] }
    Assert-Equal 'イベント定義!J9|新規登録ボタン押下|一致|' (& $show 'C03') '「」は外して比べる'
    Assert-Equal '画面項目定義!AB11|（空欄）|一致|' (& $show 'S01') '空欄のセルの指摘'
    Assert-Equal 'イベント定義!P8|検索条件を初期表示の状態に戻し|場所を直した|AI の場所: イベント定義!P7（引用のあるセルに直しました）' (& $show 'S06') '近くの行'
    Assert-Equal 'イベント定義!P11|前回と同じ検索条件で|場所を直した|AI の場所: 存在しないシート!A1（引用のあるセルに直しました）' (& $show 'S08') '無いシート'
    Assert-Equal '画面概要!I8|受注を削除する|見つからない|引用が設計書に見つかりません（AI の誤った指摘かもしれません）' (& $show 'S05')
    Assert-Equal '画面項目定義!AK6|（空欄）|見つからない|場所のセルは空欄ではありません（半角英数（E0003））' (& $show 'S10')
    Assert-Equal 'イベント定義!P11||引用なし|' (& $show 'C02') '引用の列の無い行'
    Assert-True (@($f.Links | Where-Object { $_.Location -eq "'イベント定義'!P8" }).Count -eq 1) '直した場所へのリンク'

    # 前の形の依頼（依頼の一覧に回答の形が無い）は、引用の列を読まない
    $manifestPath = Join-Path $dir 'requests.json'
    $m = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $m.PSObject.Properties.Remove('AnswerFormat')
    [System.IO.File]::WriteAllText($manifestPath, ($m | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($true)))
    $null = & $Review -OutDir $dir -ImportOnly 6>$null 3>$null
    $f = Read-XlsxSheet -Path (Join-Path $dir 'レビュー結果.xlsx') -Index 1
    $c03 = @($f.Rows.Keys | Where-Object { $_ -gt 1 -and $f.Rows[$_]['E'] -eq 'C03' })[0]
    Assert-Equal '「新規登録ボタン押下」||' ('{0}|{1}|{2}' -f $f.Rows[$c03]['L'], $f.Rows[$c03]['J'], $f.Rows[$c03]['K']) '4 列目を指摘内容として読む'
}

Test-Case 'Invoke-DesignReview: -Baseline で、前の版から変わった行だけを AI に渡し、変わっていない設計書は依頼しない' {
    $dir = New-TestDir
    $old = Join-Path $dir 'v1.0'
    $new = Join-Path $dir 'v1.1'
    New-Item -ItemType Directory -Path $old, $new | Out-Null
    Copy-Item -Path (Join-Path $designDir '*.xlsx') -Destination $old
    Copy-Item -Path (Join-Path $designDir '*.xlsx') -Destination $new
    $scr2 = Join-Path $new 'DD02_画面設計書_SCR002_受注登録.xlsx'
    # 版数だけの変更（表紙）はレビューしない
    Set-CellValues $scr2 '表紙' @{ X23 = '1.1' }
    # 桁数を変え、初期値を消し、行を足す。処理概要を変える
    Set-CellValues $scr2 '画面項目定義' @{ AB12 = '400'; AF11 = $null; A25 = '20'; C25 = 'ポイント利用'; I25 = 'usePoint'; P25 = 'チェックボックス'; U25 = 'I'; X25 = '－' }
    Set-CellValues $scr2 'イベント定義' @{ P12 = '入力内容を破棄せずに、受注一覧画面へ遷移する。' }
    $out = Join-Path $dir 'review'
    $text = Get-HostText { & $Review -Path $new -Baseline $old -OutDir $out -PromptOnly }
    Assert-Match $text '前の版と比べた設計書: 変更あり 1・変更なし 7'
    $m = Get-Manifest $out
    Assert-Equal @('DD02_画面設計書_SCR002_受注登録.xlsx') @($m.Requests | ForEach-Object { $_.Relative } | Select-Object -Unique) '変わっていない設計書は依頼しない'
    $p = [System.IO.File]::ReadAllText((Join-Path $out 'prompts/R001.txt'))
    Assert-Match $p '【設計書】DD02_画面設計書_SCR002_受注登録\.xlsx（前の版から変わった部分だけです。★の付いた行'
    Assert-Match $p '\n### シート: 画面項目定義（前の版から変わった行だけ）\n4\| A:■ 画面項目定義\n5\| A:No \| C:項目名 [^\n]*\n★11\| ' '節の見出しと表の見出しの行も入れる'
    Assert-Match $p '\n★11\| A:6 \| C:配送時間帯 [^\n]*\| AB:2 \| AF:（削除。旧: 指定なし） \| AS:' '消した値'
    Assert-Match $p '\n★12\| A:7 \| C:備考 [^\n]*\| AB:400（旧: 500） \|' '変えた値'
    Assert-Match $p '\n★25\| A:20 \| C:ポイント利用 \| I:usePoint' '足した行'
    Assert-Match $p '\n★12\| A:EV07 [^\n]*（旧: 入力内容を破棄し、受注一覧画面へ遷移する。）'
    # 設計書の部分（参照資料より前）
    $design = $p.Substring(0, $p.IndexOf('【参照資料】'))
    Assert-NotMatch $design '### シート: 表紙' '版数だけの変更'
    Assert-NotMatch $design '\n6\| A:1 \| C:受注日' '変わっていない行'
    Assert-Match $p '【参照資料】DD05_メッセージ一覧\.xlsx' '参照資料は全体'

    # 変わっていない設計書の AI の観点は「変更なし」。ツールの確認は、変わっていない設計書もする
    $null = & $Review -OutDir $out -ImportOnly 6>$null 3>$null
    $s = Read-XlsxSheet -Path (Join-Path $out 'レビュー結果.xlsx') -Index 2
    $rows = @($s.Rows.Keys | Where-Object { $_ -gt 1 } | ForEach-Object { $r = $s.Rows[$_]; '{0}|{1}|{2}|{3}' -f $r['A'], $r['B'], $r['F'], $r['H'] })
    Assert-True ($rows -contains 'DD03_テーブル定義書.xlsx|T03|変更なし|変更なし') 'AI だけの観点'
    Assert-True ($rows -contains 'DD03_テーブル定義書.xlsx|T01|OK|ツール') 'ツールだけの観点'
    Assert-True ($rows -contains 'DD01_画面一覧.xlsx|L01|OK|ツール, 変更なし') '両方の観点は、ツールの判定'
    Assert-True ($rows -contains 'DD02_画面設計書_SCR002_受注登録.xlsx|S05|未回答|R001') '変わった設計書は依頼する'
    Assert-Throws { & $Review -OutDir $out -ImportOnly -Baseline $old 6>$null } '依頼を作るとき'
}

Test-Case 'Invoke-DesignReview: Word の設計書（.docx）を段落・表の行の形で渡し、「本文!段落 n」の場所の指摘を取り込む' {
    $dir = New-TestDir
    $design = Join-Path $dir '設計書'
    New-WordDoc (Join-Path $design 'BD01_基本設計書_受注管理.docx') @('1. 概要', '受注を登録する画面である。受注番号は TBD とする。', @(@('項目', '桁数'), @('受注番号', '10'), @('受注日', '8')), '以上')
    $csv = Join-Path $dir '観点表.csv'
    @('観点ID,分類,観点,判定基準,重要度,対象設計書,有効,確かめ方,ツールの確認',
        'W01,記述,目的が書かれている,画面を使う目的が書かれている,中,*基本設計書*,○,,',
        'W02,記述,桁数の単位が分かる,桁数が文字数かバイト数か書かれている,高,*基本設計書*,○,,',
        'W03,記述,未確定事項が残っていない,TBD などが残っていない,高,*基本設計書*,○,ツール,残っている文言') | Set-Content -LiteralPath $csv -Encoding UTF8
    $out = Join-Path $dir 'review'
    $null = & $Review -Path $design -Checklist $csv -OutDir $out -PromptOnly 6>$null 3>$null
    $p = [System.IO.File]::ReadAllText((Join-Path $out 'prompts/R001.txt'))
    Assert-Match $p '\n- Word の設計書は、各行が「段落 12\| 文」「表 1 の 3 行目\| セル \| セル」の形です。'
    Assert-Match $p '【設計書】BD01_基本設計書_受注管理\.docx\n### シート: 本文\n段落 1\| 1\. 概要\n段落 2\| 受注を登録する画面である。受注番号は TBD とする。\n'
    Assert-Match $p '\n表 1 の 2 行目\| 受注番号 \| 10\n'

    Write-Answer $out 'R001' @'
W01 | NG | 本文!段落2 | 受注を登録する画面である | 使う目的が書かれていない | 背景が無い
W02 | NG | 本文!段落 1 | 受注番号 | 桁数の単位が分からない | 文字数かバイト数か
'@
    $null = & $Review -OutDir $out -Checklist $csv -ImportOnly 6>$null 3>$null
    $f = Read-XlsxSheet -Path (Join-Path $out 'レビュー結果.xlsx') -Index 1
    Assert-Equal @(
        'AI|W02|本文!段落 2|受注番号|場所を直した',
        'ツール（残っている文言）|W03|本文!段落 2||',
        'AI|W01|本文!段落 2|受注を登録する画面である|一致') @(2..4 | ForEach-Object { $r = $f.Rows[$_]; '{0}|{1}|{2}|{3}|{4}' -f $r['C'], $r['E'], $r['I'], $r['J'], $r['K'] }) '「段落2」は「段落 2」にそろえる'
    Assert-Equal 3 $f.Links.Count
    Assert-Equal 0 @($f.Links | Where-Object { $_.Location }).Count 'Word の文書へのリンクは開くだけ'

    # 前の版があっても、Word の文書は全体を渡す
    $old = Join-Path $dir '前の版'
    New-WordDoc (Join-Path $old 'BD01_基本設計書_受注管理.docx') @('1. 概要', '受注を登録する画面である。', '以上')
    $text = Get-HostText { & $Review -Path $design -Checklist $csv -OutDir (Join-Path $dir 'review2') -PromptOnly -Baseline $old }
    Assert-Match $text '前の版と比べた設計書: Word（全体） 1'
}

Test-Case 'Invoke-DesignReview: 作り直しても指摘一覧の対応要否などを引き継ぎ、対応要否から観点ごとの AI の指摘の採用率を出す' {
    $dir = New-TestDir
    $null = & $Review -Path $designDir -OutDir $dir -PromptOnly 6>$null 3>$null
    Write-Answer $dir 'R002' @'
S01 | NG | 画面項目定義!AB11 | （空欄） | 受注ステータスの桁数が空欄 | 入力項目には桁数が必要
S06 | NG | イベント定義!P7 | 1. 入力チェックを行う | 0 件のときの動きが無い | 異常時の記載
S06 | 要確認 | イベント定義!P8 | 検索条件を初期表示の状態に戻し | クリア後の一覧の扱い | 異常時の記載
C03 | NG | イベント定義!J9 | 新規登録ボタン押下 | 「押下」と「クリック」が混在 | 用語の統一
'@
    $report = Join-Path $dir 'レビュー結果.xlsx'
    $null = & $Review -OutDir $dir -ImportOnly 6>$null 3>$null
    $f = Read-XlsxSheet -Path $report -Index 1
    Assert-Equal @('S01', 'S06', 'S06', 'C03', 'S03') @(2..6 | ForEach-Object { $f.Rows[$_]['E'] })
    # 人が対応要否などを書く（O: 対応要否・P: 対応状況・Q: 対応者・R: 対応内容）
    Set-CellValues $report '指摘一覧' @{ O2 = '要'; P2 = '完了'; Q2 = '山田'; R2 = "型と桁数を書いた`n（v1.1）"; O3 = '不要'; O4 = '不要'; O6 = '要' }

    # 回答を足して作り直しても、同じ指摘の値は引き継ぐ
    Write-Answer $dir 'R003' "C02 | NG | イベント定義!P11 | 受注番号を採番する | 連番の採番単位が不明 | 解釈が分かれる"
    $text = Get-HostText { & $Review -OutDir $dir -ImportOnly }
    Assert-Match $text '前のレビュー結果の対応要否・対応状況・対応者・対応内容を 4 件引き継ぎました'
    Assert-Match $text 'AI の指摘の採用率（対応要否から）: 要 1・不要 2（33%）'
    Assert-Match $text '不要が多い観点: S06 0/2'
    $f = Read-XlsxSheet -Path $report -Index 1
    $rows = @(2..($f.Rows.Count) | ForEach-Object { $r = $f.Rows[$_]; '{0}|{1}|{2}|{3}|{4}|{5}' -f $r['C'], $r['E'], $r['O'], $r['P'], $r['Q'], ([string]$r['R']).Replace("`n", '/') })
    Assert-True ($rows -contains 'AI|S01|要|完了|山田|型と桁数を書いた/（v1.1）') ($rows -join "`n")
    Assert-Equal 2 @($rows | Where-Object { $_ -like 'AI|S06|不要|*' }).Count
    Assert-True ($rows -contains 'ツール（初期値）|S03|要|||') 'ツールの指摘も引き継ぐ'
    Assert-True ($rows -contains 'AI|C02||||') '新しい指摘は空欄'

    $a = Read-XlsxSheet -Path $report -Index 4
    Assert-Equal '観点ID|AI の指摘|要|不要|未記入|採用率|ツールの指摘' ((@('A', 'D', 'E', 'F', 'G', 'H', 'I') | ForEach-Object { $a.Rows[1][$_] }) -join '|')
    $byId = @{}
    foreach ($r in @($a.Rows.Keys | Where-Object { $_ -gt 1 })) { $byId[$a.Rows[$r]['A']] = $a.Rows[$r] }
    Assert-Equal @('C02', 'C03', 'S01', 'S03', 'S06') @($a.Rows.Keys | Where-Object { $_ -gt 1 } | Sort-Object | ForEach-Object { $a.Rows[$_]['A'] }) '観点表の順'
    Assert-Equal '1|1|0|0|1|0' ('{0}|{1}|{2}|{3}|{4}|{5}' -f $byId['S01']['D'], $byId['S01']['E'], $byId['S01']['F'], $byId['S01']['G'], $byId['S01']['H'], $byId['S01']['I'])
    Assert-Equal '2|0|2|0|0' ('{0}|{1}|{2}|{3}|{4}' -f $byId['S06']['D'], $byId['S06']['E'], $byId['S06']['F'], $byId['S06']['G'], $byId['S06']['H'])
    Assert-Match $byId['S06']['J'] '^不要が多い観点です'
    Assert-Equal '0|3' ('{0}|{1}' -f $byId['S03']['D'], $byId['S03']['I']) 'ツールの指摘は採用率に数えない'
    Assert-Match $byId['C02']['J'] '対応要否（要・不要）を書くと'
}

Test-Case 'Invoke-DesignReview: -Recheck で、前のレビューの NG・要確認の指摘が直ったかを確かめる（AI の指摘は AI に、ツールの指摘はツールで）' {
    $dir = Copy-Fixture $designDir
    $first = Join-Path $dir '1回目'
    $null = & $Review -Path $dir -OutDir $first -PromptOnly 6>$null 3>$null
    Write-Answer $first 'R002' $answerTable
    $null = & $Review -OutDir $first -ImportOnly 6>$null 3>$null
    $report = Join-Path $first 'レビュー結果.xlsx'
    $f = Read-XlsxSheet -Path $report -Index 1
    Assert-Equal @('S01', 'S06', 'C03', 'S03', 'S03', 'S03') @(2..7 | ForEach-Object { $f.Rows[$_]['E'] })
    # C03（No 3）は対応しない。SCR002 の希望納品日の初期値を直す
    Set-CellValues $report '指摘一覧' @{ O4 = '不要' }
    Set-CellValues (Join-Path $dir 'DD02_画面設計書_SCR002_受注登録.xlsx') '画面項目定義' @{ AF10 = 'なし' }

    $second = Join-Path $dir '再確認'
    $text = Get-HostText { & $Review -Path $dir -OutDir $second -Recheck $first -PromptOnly }
    Assert-Match $text '前のレビューの指摘: 5 件（AI 2 件・ツール 3 件。今の設計書が見つからない指摘 0 件）' '対応要否が不要の指摘は除く'
    $m = Get-Manifest $second
    Assert-Equal 1 @($m.Requests).Count 'AI の指摘のある設計書だけ'
    Assert-Equal '再確認|1,2' ('{0}|{1}' -f $m.Requests[0].Kind, (@($m.Requests[0].FindingNos) -join ',')) '依頼の一覧の指摘No は配列のまま書く（Windows PowerShell 5.1 の ConvertTo-Json でも）'
    $p = [System.IO.File]::ReadAllText((Join-Path $second 'prompts/R001.txt'))
    Assert-Match $p '【前回の指摘】\n指摘No \| 観点 \| 場所（前の版） \| 指摘内容\n1 \| 入力項目の定義に漏れがない \| 画面項目定義!M11 \| 受注ステータスの型が未定義\n2 \| 処理手順と異常時の動きが書かれている \| イベント定義!P8 \| クリア時の動きが不明\n'
    Assert-Match $p '\n### シート: 画面項目定義\n' '指摘のあったシート'
    Assert-NotMatch $p '### シート: 画面概要' '指摘の無いシートは渡さない'

    Write-Answer $second 'R001' @'
1 | 解消 | - | - | 型が書かれた
2 | 未解消 | イベント定義!P7 | 検索条件を初期表示の状態に戻し | クリア後の扱いがまだ書かれていない
'@
    $text = Get-HostText { & $Review -Path $dir -OutDir $second -ImportOnly }
    Assert-Match $text '再確認: 解消 2・未解消 3'
    $r = Read-XlsxSheet -Path (Join-Path $second 'レビュー結果.xlsx') -Index 1
    Assert-Equal @('再確認', '依頼一覧') $r.Sheets
    Assert-Equal @(
        '1|AI|S01|解消||R001',
        '2|AI|S06|未解消|イベント定義!P8|R001',
        '4|ツール（初期値）|S03|解消||ツール',
        '5|ツール（初期値）|S03|未解消|画面項目定義!AF12|ツール',
        '6|ツール（初期値）|S03|未解消|画面項目定義!AF13|ツール') @(2..6 | ForEach-Object { $x = $r.Rows[$_]; '{0}|{1}|{2}|{3}|{4}|{5}' -f $x['A'], $x['B'], $x['C'], $x['H'], $x['I'], $x['L'] })
    Assert-Match $r.Rows[3]['K'] '引用のあるセルに直しました' '残っている場所の引用も確かめる'
    Assert-True (@($r.Links | Where-Object { $_.Location -eq "'イベント定義'!P8" }).Count -eq 1) '残っている場所へのリンク'
    Assert-Throws { & $Review -Path $dir -OutDir $second -Recheck $first -Baseline $first 6>$null } '同時に指定できません'
    Assert-Throws { & $Review -Path $dir -OutDir $first -Recheck $report 6>$null } '別のフォルダを指定してください'
    Assert-Equal @('指摘一覧', '観点消化状況', '依頼一覧', '観点別の採用率') (Read-XlsxSheet -Path $report -Index 1).Sheets '前のレビュー結果は上書きしない'
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

Test-Case 'Invoke-DesignReview: 対話モード（-AnswerMode Enter）でやり直し・中断・再開ができる' {
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

    $out = Get-HostText { & $Review -Path $designDir -OutDir $dir -AnswerMode Enter }
    Assert-Match $out 'クリップボードの内容が依頼のまま'
    Assert-Match $out '読み取れませんでした'
    Assert-Match $out '中断しました'
    Assert-Equal 3 @(Get-ChildItem -LiteralPath (Join-Path $dir 'answers') -Filter 'R*.txt').Count '中断までに 3 件回答'

    $state.Step = 100
    $out = Get-HostText { & $Review -Path $designDir -OutDir $dir -AnswerMode Enter }
    Assert-Match $out '未回答の依頼 7 件'
    Assert-Match $out '依頼 10 件中 10 件回答済み'
}

Test-Case 'Invoke-DesignReview: -AnswerMode Watch では、回答をコピーすると Enter を押さなくても取り込んで次の依頼に進む' {
    $dir = New-TestDir
    # クリップボードと待ち時間をテスト用に差し替える。依頼ごとに、2 回は依頼のまま（改行が CRLF になったものも）、
    # 3 回目は回答でない文、4 回目に回答を返す
    $state = @{ Clip = ''; Polls = 0; Sleeps = 0 }
    function Set-Clipboard { param($Value) $state.Clip = $Value; $state.Polls = 0 }
    function Start-Sleep { param($Milliseconds) $state.Sleeps++ }
    function Get-Clipboard {
        param([switch] $Raw)
        $state.Polls++
        if ($state.Polls -gt 10) { throw '回答を取り込まずに見張り続けています' }
        if ($state.Polls -eq 1) { return $state.Clip }
        if ($state.Polls -eq 2) { return $state.Clip.Replace("`n", "`r`n") }
        if ($state.Polls -eq 3) { return 'うまく回答できませんでした' }
        $ids = [regex]::Matches($state.Clip, '(?m)^([A-Z]\d{2}) \| ') | ForEach-Object { $_.Groups[1].Value } | Select-Object -Unique
        return (($ids | ForEach-Object { "$_ | OK | - | - | - | -" }) -join "`n")
    }
    $out = Get-HostText { & $Review -Path $designDir -OutDir $dir -AnswerMode Watch }
    Assert-Match $out '回答をコピー（回答の下のコピー ボタン）すると、自動で取り込んで次の依頼に進みます'
    Assert-Match $out 'コピーした内容から「観点ID \| 判定 \| \.\.\.」の行を読み取れませんでした'
    Assert-Equal 10 @(Get-ChildItem -LiteralPath (Join-Path $dir 'answers') -Filter 'R*.txt').Count 'すべて取り込む'
    Assert-Match $out '依頼 10 件中 10 件回答済み'
    Assert-Equal 30 $state.Sleeps '取り込むまでは待ってから見直す'
}

Test-Case 'Invoke-DesignReview: Get-Help でヘルプが表示される' {
    $h = Get-Help $Review
    Assert-Match $h.Synopsis 'AI'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
