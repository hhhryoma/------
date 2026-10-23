#     New-TestSpecDraft.ps1 のテスト（詳細設計サンプルの画面設計書から、テスト仕様書の下書きを作る）

$TestDraft = Join-Path $script:ToolsDir 'New-TestSpecDraft.ps1'
$designSample = Join-Path $script:RepoRoot 'docs/design-sample'

Test-Case 'New-TestSpecDraft: 入力項目の必須・桁数・型・入力チェック・コード値と、イベントのケースを、メッセージの文言付きで作る' {
    $out = Join-Path (New-TestDir) 'UT_下書き.xlsx'
    $cases = @(& $TestDraft $designSample -OutXlsx $out 6>$null)
    $scr2 = @($cases | Where-Object { $_.Function -eq 'SCR002 受注登録' })
    $find = { param($content) @($scr2 | Where-Object { $_.Content -eq $content })[0] }

    Assert-Equal 'E0001「受注日は必須入力です。」 が表示され、処理が進まない' (& $find '受注日: 未入力（必須）').Expected
    Assert-Equal '境界|8 桁' ('{0}|{1}' -f (& $find '顧客コード: 最大の 8 桁').Kind, (& $find '顧客コード: 最大の 8 桁').InputValue)
    Assert-Match (& $find '顧客コード: 最大を超える 9 桁').Expected '^E0002「顧客コードは8文字以内で入力してください。」'
    Assert-Equal 'あいう（全角）' (& $find '顧客コード: 半角英数').InputValue
    # 存在のチェックは、項目のイベント（フォーカスアウト）で確かめる
    Assert-Match (& $find '顧客コード: 存在').Steps '顧客コードのフォーカスアウト$'
    Assert-Equal '0|E0009「数量は1以上の数値を入力してください。」 が表示され、処理が進まない' ('{0}|{1}' -f (& $find '数量: 1以上').InputValue, (& $find '数量: 1以上').Expected)
    Assert-Match (& $find '希望納品日: 受注日以降').Expected '希望納品日は受注日以降の日付'
    Assert-Equal @('銀行振込', 'クレジットカード', '代金引換') @($scr2 | Where-Object { $_.Content -like '支払方法: 「*' } | ForEach-Object { $_.InputValue })
    Assert-Match (& $find '初期表示（EV01）').Expected '初期値: 受注日: システム日付、支払方法: 銀行振込'
    Assert-Match (& $find '登録（EV06）: 正しい値で実行する').Expected "遷移先: SCR003$"
    Assert-True ($null -ne (& $find '登録（EV06）: 確認で「キャンセル」を選ぶ')) '確認メッセージのキャンセル'
    # 入力項目のケースで確かめていないエラーだけをイベントのケースにする（パラメータは {0} のまま）
    Assert-Equal @('登録（EV06）: E0008 になる場合') @($scr2 | Where-Object { $_.Content -match 'になる場合' } | ForEach-Object { $_.Content })
    Assert-Match (& $find '登録（EV06）: E0008 になる場合').Expected '\{0\}'
    Assert-True (@($scr2 | Where-Object { $_.Content -like '顧客名*' }).Count -eq 0) '出力だけの項目にはケースを作らない'
    Assert-True (@($scr2 | Where-Object { $_.Content -like '登録:*' }).Count -eq 0) 'ボタンにはケースを作らない'

    # 画面ごとのシートと一覧。既存のツールで読めて、抜けが無い
    $book = Read-XlsxSheet -Path $out -Index 1
    Assert-Equal @('一覧', 'SCR001', 'SCR002', 'SCR003') $book.Sheets
    $spec = @(& (Join-Path $script:ToolsDir 'Test-TestSpec.ps1') $out -DesignPath $designSample 6>$null 3>$null)
    Assert-Equal 0 $spec.Count (@($spec | ForEach-Object { $_.Detail }) -join ' / ')

    Assert-Throws { & $TestDraft $designSample -OutXlsx $out 6>$null } '-Force を付けたときだけ上書き'
    $only = @(& $TestDraft $designSample -OutXlsx $out -Screen 'SCR00[13]' -Force 6>$null)
    Assert-Equal @('SCR001 受注一覧', 'SCR003 受注詳細') @($only | ForEach-Object { $_.Function } | Select-Object -Unique)
}

Test-Case 'New-TestSpecDraft: 画面設計書が無いときと、拡張子の誤りを知らせる' {
    $dir = New-TestDir
    Assert-Throws { & $TestDraft (Join-Path $designSample 'DD05_メッセージ一覧.xlsx') -OutXlsx (Join-Path $dir 'a.xlsx') 6>$null } '画面設計書'
    Assert-Throws { & $TestDraft $designSample -OutXlsx (Join-Path $dir 'a.csv') 6>$null } '拡張子は .xlsx'
}

Test-Case 'New-TestSpecDraft: Get-Help でヘルプが表示される' {
    $h = Get-Help $TestDraft
    Assert-Match $h.Synopsis 'テスト仕様書の下書き'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
