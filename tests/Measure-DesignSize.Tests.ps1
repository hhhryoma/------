#     Measure-DesignSize.ps1 のテスト。設計書は docs\design-sample と tests\fixtures\size（make_fixtures.py で生成）

$MeasureSize = Join-Path $script:ToolsDir 'Measure-DesignSize.ps1'
$sizeDir = Join-Path $script:Fixtures 'size'
$sampleDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Invoke-Measure {
    param([hashtable] $Options)
    return , @(& $MeasureSize @Options 6>$null 3>$null)
}

Test-Case 'Measure-DesignSize: 詳細設計サンプルの画面・テーブル・IF・メッセージ・コードを数える' {
    $r = Invoke-Measure @{ Path = $sampleDir }
    $got = @($r | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Kind, $_.Name, $_.Measures, $_.Level })
    Assert-Equal (@(
        '画面|SCR001 受注一覧|項目数 17・イベント数 6|普通',
        '画面|SCR002 受注登録|項目数 19・イベント数 7|普通',
        '画面|SCR003 受注詳細|項目数 21・イベント数 4|普通',
        'テーブル|M_CUSTOMER 顧客マスタ|カラム数 14|簡易',
        'テーブル|M_PRODUCT 商品マスタ|カラム数 10|簡易',
        'テーブル|T_ORDER 受注|カラム数 14|簡易',
        'テーブル|T_ORDER_DETAIL 受注明細|カラム数 11|簡易',
        'IF|IF001 出荷指示データ|項目数 12|簡易',
        'メッセージ|DD05_メッセージ一覧.xlsx（メッセージ一覧）|件数 16|',
        'コード|DD04_コード定義書.xlsx（コード一覧）|件数 15|') | Sort-Object) ($got | Sort-Object)
    $s = @(& $MeasureSize -Path $sampleDir -Summary 6>$null 3>$null)[0]
    Assert-Equal 20.1 $s.TotalDays '15 + 2 + 2 + 16×0.05 + 15×0.02'
    Assert-Equal 140 $s.FpIndicative '35 × テーブル 4'
    Assert-Equal 61 $s.FpEstimated '7×4 + 4×EI 2 + 5×EO 1 + 4×EQ 5'
}

Test-Case 'Measure-DesignSize: イベントの使用テーブルで外部入力・外部照会を分け、IF の向き・帳票・他システムのテーブルを FP に入れる' {
    $r = Invoke-Measure @{ Path = $sizeDir; ExternalTables = '^M_USER' }
    $s01 = @($r | Where-Object Name -eq 'S01 ログイン')[0]
    Assert-Equal '項目数 3・イベント数 3' $s01.Measures
    Assert-Equal '1|1|0' ('{0}|{1}|{2}' -f $s01.EI, $s01.EQ, $s01.EO) 'EV02 は登録あり、EV03 は参照だけ、EV01 はテーブルなし'
    Assert-Equal 1 (@($r | Where-Object Kind -eq 'IF')[0].EI) '受信の IF は外部入力'
    Assert-Equal 1 (@($r | Where-Object Kind -eq '帳票')[0].EO) '帳票は外部出力'
    Assert-Equal '普通' (@($r | Where-Object Name -like 'M_USER*')[0].Level) 'カラム 25 は普通'
    $s = @(& $MeasureSize -Path $sizeDir -ExternalTables '^M_USER' -Summary 6>$null 3>$null)[0]
    Assert-Equal '1|1|2|1|1' ('{0}|{1}|{2}|{3}|{4}' -f $s.ILF, $s.EIF, $s.EI, $s.EO, $s.EQ)
    Assert-Equal 50 $s.FpIndicative
    Assert-Equal 29 $s.FpEstimated
}

Test-Case 'Measure-DesignSize: 見出しが見つからない設計書と、0 件の表を警告する' {
    $w = $null
    $null = & $MeasureSize -Path $sizeDir -WarningVariable w -WarningAction SilentlyContinue 6>$null
    $text = ($w | ForEach-Object { $_.Message }) -join "`n"
    Assert-Match $text '画面設計書_S02_壊れ\.xlsx: 「画面」の見出し（項目名 / 項目ID）が見つからない'
    Assert-Match $text '画面設計書_S02_壊れ\.xlsx: 画面「S02 壊れた画面」のイベント数が 0 件'
}

Test-Case 'Measure-DesignSize: 係数表で難易度と工数を変えられる（「件」の行は件数に掛ける）' {
    $dir = New-TestDir
    $coef = Join-Path $dir '係数.csv'
    [System.IO.File]::WriteAllText($coef, "種類,難易度,基準,上限,工数,単位`r`n画面,小,項目数,5,1,`r`n画面,大,項目数,,9,`r`nテーブル,,カラム数,,0.1,件`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $r = Invoke-Measure @{ Path = $sizeDir; Coefficients = $coef }
    Assert-Equal '小|1' ('{0}|{1}' -f @($r | Where-Object Name -eq 'S01 ログイン')[0].Level, @($r | Where-Object Name -eq 'S01 ログイン')[0].Days)
    Assert-Equal 2.5 ([math]::Round(@($r | Where-Object Name -like 'M_USER*')[0].Days, 2)) 'カラム 25 × 0.1'
    Assert-True ($null -eq @($r | Where-Object Kind -eq 'IF')[0].Days) '係数表に無い種類は工数なし'
}

Test-Case 'Measure-DesignSize: -OutXlsx で集計・規模一覧・係数表を保存する' {
    $out = Join-Path (New-TestDir) 'size.xlsx'
    $null = Invoke-Measure @{ Path = $sampleDir; OutXlsx = $out }
    $sum = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('集計', '規模一覧', '係数表') $sum.Sheets
    $labels = @($sum.Rows.Keys | ForEach-Object { $sum.Rows[$_]['A'] })
    Assert-True ($labels -contains 'FP（NESMA 試算法）') 'FP の行'
    $detail = Read-XlsxSheet -Path $out -Index 2
    Assert-True (@($detail.Links | Where-Object { $_.Location -like "'T_ORDER'!A*" }).Count -eq 1) '設計書の表へのリンク'
}

Test-Case 'Measure-DesignSize: Get-Help でヘルプが表示される' {
    $h = Get-Help $MeasureSize
    Assert-Match $h.Synopsis '規模'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
