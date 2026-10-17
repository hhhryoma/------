#     Test-ExcelTerms.ps1 と Replace-Excel.ps1 -ListPath のテスト。フィクスチャは tests\fixtures\terms（make_fixtures.py で生成）

$TermsTool = Join-Path $script:ToolsDir 'Test-ExcelTerms.ps1'
$ReplaceTool = Join-Path $script:ToolsDir 'Replace-Excel.ps1'

function Invoke-Terms {
    param([hashtable] $Options)
    return , @(& $TermsTool @Options 6>$null 3>$null)
}

function Get-VariantKeys {
    param($Results)
    return @($Results | ForEach-Object { '{0}→{1}|{2}|{3}' -f $_.Variant, $_.Preferred, $_.Kind, $_.Count } | Sort-Object)
}

Test-Case 'Test-ExcelTerms: 用語集の揺れと、長音・全角半角の揺れを見つける' {
    $dir = Copy-Fixture 'terms'
    $r = Invoke-Terms @{ Path = $dir; Glossary = (Join-Path $dir '用語集.csv') }
    Assert-Equal (@(
        'ＩＤ→ID|全角・半角|1',
        'インタフェース→インターフェース|長音（ー）|1',
        'サーバ→サーバー|長音（ー）|1',
        'ユーザ→ユーザー|用語集|1',
        'ﾕｰｻﾞｰ→ユーザー|全角・半角|1',
        '出来る→できる|用語集|1') | Sort-Object) (Get-VariantKeys $r)
    # 「ユーザー」の中の「ユーザ」は数えない。推奨の表記の件数も出す
    Assert-Equal 2 (@($r | Where-Object Variant -eq 'ユーザ')[0].PreferredCount)
    Assert-Equal 'a.xlsx 機能!A2' (@($r | Where-Object Variant -eq 'ユーザ')[0].Examples)
    # 大文字・小文字だけの違い（Id / ID）は揺れにしない
    Assert-Equal 0 @($r | Where-Object { $_.Variant -ceq 'Id' -or $_.Preferred -ceq 'Id' }).Count
}

Test-Case 'Test-ExcelTerms: -NoAuto では用語集だけで確かめる。用語集の正しい表記は自動の検出より優先する' {
    $dir = Copy-Fixture 'terms'
    $r = Invoke-Terms @{ Path = $dir; Glossary = (Join-Path $dir '用語集.csv'); NoAuto = $true }
    Assert-Equal (@('ユーザ→ユーザー|用語集|1', '出来る→できる|用語集|1') | Sort-Object) (Get-VariantKeys $r)

    # 用語集で「サーバ」を正しい表記にすると、自動の検出も「サーバ」に合わせる
    $g = Join-Path $dir 'g2.csv'
    [System.IO.File]::WriteAllText($g, "正しい表記,揺れ`r`nサーバ,サーバー`r`n", (New-Object System.Text.UTF8Encoding($true)))
    $r = Invoke-Terms @{ Path = $dir; Glossary = $g }
    Assert-True (@($r | Where-Object { $_.Variant -eq 'サーバー' -and $_.Preferred -eq 'サーバ' }).Count -eq 1) 'サーバー→サーバ'
    Assert-Equal 0 @($r | Where-Object { $_.Variant -eq 'サーバ' }).Count '逆向きの指摘をしない'
}

Test-Case 'Test-ExcelTerms: -OutCsv の置換リストを Replace-Excel -ListPath に渡すと、揺れがなくなる' {
    $dir = Copy-Fixture 'terms'
    $csv = Join-Path (New-TestDir) 'list.csv'
    $null = Invoke-Terms @{ Path = $dir; Glossary = (Join-Path $dir '用語集.csv'); OutCsv = $csv }
    $rows = @(Import-Csv -LiteralPath $csv -Encoding UTF8)
    Assert-Equal 6 $rows.Count
    Assert-Equal '置換前|置換後|種類|件数|例' (@($rows[0].PSObject.Properties | ForEach-Object { $_.Name }) -join '|')

    $replaced = @(& $ReplaceTool -ListPath $csv -Path $dir -Force 6>$null 3>$null)
    Assert-Equal 6 @($replaced | Where-Object Result -eq '置換').Count
    $after = Invoke-Terms @{ Path = $dir; Glossary = (Join-Path $dir '用語集.csv') }
    Assert-Equal 0 $after.Count ((Get-VariantKeys $after) -join ', ')
    # 置換後の文字の一部（「ユーザー管理」の「ユーザ」）は置換しない
    $texts = @(& (Join-Path $script:ToolsDir 'Search-Excel.ps1') -Pattern 'ユーザ' -Path $dir 6>$null | ForEach-Object { $_.Text } | Sort-Object)
    Assert-Equal (@('ユーザーIDを入力', 'ユーザー一覧', 'ユーザー管理', 'ユーザー登録') | Sort-Object) $texts
}

Test-Case 'Replace-Excel: -ListPath は Shift_JIS の CSV と正規表現の行も扱え、何度実行しても同じ結果になる' {
    $dir = Copy-Fixture 'terms'
    $csv = Join-Path (New-TestDir) 'list.csv'
    $sjis = [System.Text.Encoding]::GetEncoding(932)
    [System.IO.File]::WriteAllText($csv, "置換前,置換後,正規表現`r`nユーザ,ユーザー,`r`n(変更|削除)出来る,`$1できる,○`r`n", $sjis)
    $r = @(& $ReplaceTool -ListPath $csv -Path $dir -Force 6>$null 3>$null)
    Assert-Equal (@('a.xlsx|機能!A2|ユーザ登録 → ユーザー登録', 'b.xlsx|画面!A4|変更出来る → 変更できる') | Sort-Object) @($r | ForEach-Object { '{0}|{1}|{2}' -f $_.File, $_.Location, $_.Change } | Sort-Object)
    # 2 回目は置換する箇所がない（「ユーザー」の中の「ユーザ」を置換して「ユーザーー」にしない）
    $again = @(& $ReplaceTool -ListPath $csv -Path $dir -Force 6>$null 3>$null)
    Assert-Equal 0 $again.Count
    $single = @(& $ReplaceTool -Pattern 'ユーザ' -Replacement 'ユーザー' -Path $dir -SimpleMatch -Force 6>$null 3>$null)
    Assert-Equal 0 $single.Count '-Pattern でも同じ'
}

Test-Case 'Replace-Excel: -ListPath と -Pattern の指定の誤りを知らせる' {
    $dir = Copy-Fixture 'terms'
    Assert-Throws { & $ReplaceTool -Pattern 'a' -ListPath (Join-Path $dir '用語集.csv') -Path $dir -Force 6>$null } '同時に指定できません'
    Assert-Throws { & $ReplaceTool -ListPath (Join-Path $dir '用語集.csv') -Path $dir -Force 6>$null } '「置換前」「置換後」の列'
    Assert-Throws { & $ReplaceTool -Pattern 'a' -Path $dir -Force 6>$null } '-Replacement'
    Assert-Throws { & $ReplaceTool -Path $dir -Force 6>$null } '-ListPath'
}

Test-Case 'Test-ExcelTerms: -OutXlsx で出現箇所へのリンク付きの一覧を保存する' {
    $dir = Copy-Fixture 'terms'
    $out = Join-Path (New-TestDir) 'terms.xlsx'
    $null = Invoke-Terms @{ Path = $dir; OutXlsx = $out }
    $summary = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal '推奨' $summary.Rows[1]['A']
    Assert-Equal 6 $summary.Rows.Count '見出しと、用語集なしで見つかる 5 種類（ユーザ も自動で見つかる）'
    $detail = Read-XlsxSheet -Path $out -Index 2
    Assert-True (@($detail.Links | Where-Object { $_.Location -eq "'機能'!A4" }).Count -eq 1) 'ＩＤ の出現箇所へのリンク'
}

Test-Case 'Test-ExcelTerms: Get-Help でヘルプが表示される' {
    $h = Get-Help $TermsTool
    Assert-Match $h.Synopsis '表記ゆれ'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
