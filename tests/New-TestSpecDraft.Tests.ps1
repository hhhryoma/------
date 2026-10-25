#     New-TestSpecDraft.ps1 のテスト（詳細設計サンプルの画面設計書から、テスト仕様書の下書きを作る）

$TestDraft = Join-Path $script:ToolsDir 'New-TestSpecDraft.ps1'
$designSample = Join-Path $script:RepoRoot 'docs/design-sample'

function New-DraftDesign {
    #     画面 SCR010（顧客登録）の画面設計書（画面項目定義と、登録のイベント）とメッセージ一覧を作る（テスト用）。
    #     Items: 画面項目定義の行（No・項目名・入出力・型・桁数・必須・入力チェック）の配列、EventMessages: 登録のイベントのメッセージ欄、Messages: メッセージ一覧の行（ID・文言）の配列
    param([string] $Dir, [object[]] $Items, [string] $EventMessages, [object[]] $Messages)
    $hdr = @('画面ID', 'SCR010', '画面名', '顧客登録', '', '', '', '')
    New-SheetsBook (Join-Path $Dir '画面設計書_SCR010.xlsx') ([ordered]@{
            '画面項目定義' = @($hdr, @(), @('No', '項目名', '入出力', '型', '桁数', '必須', '入力チェック')) + $Items
            'イベント定義' = @($hdr, @(), @('イベントID', 'イベント名', 'トリガー', '処理概要', 'メッセージ'), @('EV02', '登録', '登録ボタン押下', '入力チェックを行い、顧客を登録する。', $EventMessages))
        })
    New-SheetsBook (Join-Path $Dir 'メッセージ一覧.xlsx') ([ordered]@{ 'メッセージ一覧' = @(, @('メッセージID', 'メッセージ')) + $Messages })
}

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

Test-Case 'New-TestSpecDraft: 入力チェック欄はセル内の改行でも区切り、「1,000以上」のような数の中の「,」では区切らない' {
    $dir = New-TestDir
    New-DraftDesign $dir @(
        @('1', '顧客コード', 'I', '半角英数', '8', '○', "必須（E0001）`n半角英数（E0003）`n存在（E0005）"),
        @('2', '与信限度額', 'I', '数値', '9', '', '1,000以上（E0009）、9,999,999以下（E0011）')) 'E0001, E0003, E0005, E0009, E0011' @(
        @('E0001', '{0}は必須入力です。'), @('E0003', '{0}は半角英数字で入力してください。'), @('E0005', '指定された{0}は存在しません。'),
        @('E0009', '{0}は{1}以上の値を入力してください。'), @('E0011', '{0}は{1}以下の値を入力してください。'))
    $cases = @(& $TestDraft $dir -OutXlsx (Join-Path $dir 'UT.xlsx') 6>$null)
    $find = { param($content) @($cases | Where-Object { $_.Content -eq $content })[0] }

    Assert-Equal 'E0001「顧客コードは必須入力です。」 が表示され、処理が進まない' (& $find '顧客コード: 未入力（必須）').Expected
    Assert-Match (& $find '顧客コード: 半角英数').Expected '^E0003「顧客コードは半角英数字で入力してください。」'
    Assert-Match (& $find '顧客コード: 存在').Expected '^E0005「指定された顧客コードは存在しません。」'
    Assert-Equal 0 @($cases | Where-Object { $_.Content -match 'になる場合' }).Count '入力チェックで確かめたエラーは、イベントのケースにしない'
    # 「1,000以上」には 999、「9,999,999以下」には 10000000 を入れる（{1} は設計書の書き方のまま）
    Assert-Equal @(
        '与信限度額: 1,000以上|999|E0009「与信限度額は1,000以上の値を入力してください。」 が表示され、処理が進まない',
        '与信限度額: 9,999,999以下|10000000|E0011「与信限度額は9,999,999以下の値を入力してください。」 が表示され、処理が進まない') @($cases | Where-Object { $_.Viewpoint -eq '入力チェック' -and $_.Content -like '与信限度額*' } | ForEach-Object { '{0}|{1}|{2}' -f $_.Content, $_.InputValue, $_.Expected })
}

Test-Case 'New-TestSpecDraft: 処理概要・遷移先・初期値のセル内の改行は改行で書き、メッセージと初期値の文言は設計書のまま（全角の「？」「－」）にする' {
    $out = Join-Path (New-TestDir) 'UT_下書き.xlsx'
    $cases = @(& $TestDraft $designSample -OutXlsx $out 6>$null)
    Assert-Equal 0 @($cases | Where-Object { "$($_.Content)$($_.Steps)$($_.Expected)" -match '↵' }).Count '↵ のまま書かない'
    $initial = @($cases | Where-Object { $_.Function -eq 'SCR001 受注一覧' -and $_.Content -eq '初期表示（EV01）' })[0]
    Assert-Match $initial.Expected "システム日付を設定する。`n2\. 受注ステータスの選択肢"
    Assert-Match $initial.Expected '初期値: 受注日（From）: システム日付－30日、'
    Assert-Match @($cases | Where-Object { $_.Content -eq '登録（EV06）: 確認で「キャンセル」を選ぶ' })[0].Steps 'Q0001「登録します。よろしいですか？」 で「キャンセル」を選ぶ$'
    # 書き出したセルも改行で区切る
    $sheet = Read-XlsxSheet -Path $out -Index 2
    Assert-Match $sheet.Rows[2]['H'] "設定する。`n2\. "
}

Test-Case 'New-TestSpecDraft: MSG-E001・MSGE001 のような接頭辞のあるメッセージ ID も、期待結果の文言・確認のキャンセル・エラーのケースにする' {
    foreach ($p in @('MSG-', 'MSG')) {
        $dir = New-TestDir
        New-DraftDesign $dir @(, @('1', '顧客コード', 'I', '半角英数', '8', '○', ('必須（{0}E001）、半角英数（{0}E003）' -f $p))) ('{0}Q001, {0}E001, {0}E003, {0}E008' -f $p) @(
            @("${p}E001", '{0}は必須入力です。'), @("${p}E002", '{0}は{1}文字以内で入力してください。'), @("${p}E003", '｛０｝は半角英数字で入力してください。'),
            @("${p}E008", '在庫が不足しています。'), @("${p}Q001", '登録します。よろしいですか？'))
        $cases = @(& $TestDraft $dir -OutXlsx (Join-Path $dir 'UT.xlsx') 6>$null)
        $find = { param($content) @($cases | Where-Object { $_.Content -eq $content })[0] }

        Assert-Equal "${p}E001「顧客コードは必須入力です。」 が表示され、処理が進まない" (& $find '顧客コード: 未入力（必須）').Expected $p
        Assert-Match (& $find '顧客コード: 最大を超える 9 桁').Expected ('^{0}E002「顧客コードは8文字以内で入力してください。」' -f $p) $p
        # 全角で書いた ｛０｝ も項目名にする
        Assert-Equal "${p}E003「顧客コードは半角英数字で入力してください。」 が表示され、処理が進まない" (& $find '顧客コード: 半角英数').Expected $p
        Assert-Match (& $find '登録（EV02）: 確認で「キャンセル」を選ぶ').Steps ('{0}Q001「登録します。よろしいですか？」 で「キャンセル」' -f $p) $p
        Assert-Equal @("登録（EV02）: ${p}E008 になる場合") @($cases | Where-Object { $_.Content -match 'になる場合' } | ForEach-Object { $_.Content }) $p
    }
}

Test-Case 'New-TestSpecDraft: 必須の印（●・◎・✓）と、桁数の「全角20」「40バイト」「7,2」を読み、読めない桁数は知らせる' {
    $dir = New-TestDir
    New-DraftDesign $dir @(
        @('1', '顧客コード', 'I', '半角英数', '8', '●', ''),
        @('2', '顧客名', 'I', '全角', '全角20', '◎', ''),
        @('3', '顧客名カナ', 'I', '全角カナ', '40バイト', '✓', ''),
        @('4', '単価', 'I', '数値', '7,2', '', ''),
        @('5', '電話番号', 'I', '半角数字', '10-11', '', ''),
        @('6', '備考', 'I', '文字', '可変', '', ''),
        @('7', '生年月日', 'I', '日付', 'yyyy/MM/dd', '', '')) 'E0001' @(@('E0001', '{0}は必須入力です。'), @('E0002', '{0}は{1}文字以内で入力してください。'))
    $cases = @(& $TestDraft $dir -OutXlsx (Join-Path $dir 'UT.xlsx') -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)

    Assert-Equal @('顧客コード: 未入力（必須）', '顧客名: 未入力（必須）', '顧客名カナ: 未入力（必須）') @($cases | Where-Object { $_.Viewpoint -eq '必須' } | ForEach-Object { $_.Content })
    # 「7,2」（全体と小数の桁数）は全体の 7 桁
    Assert-Equal @('顧客コード: 最大の 8 桁', '顧客コード: 最大を超える 9 桁', '顧客名: 最大の 20 文字', '顧客名: 最大を超える 21 文字',
        '顧客名カナ: 最大の 40 文字', '顧客名カナ: 最大を超える 41 文字', '単価: 最大の 7 桁', '単価: 最大を超える 8 桁') @($cases | Where-Object { $_.Viewpoint -eq '桁数' } | ForEach-Object { $_.Content })
    # 読めない桁数（範囲・可変）は知らせる。日付の書式は境界のケースを作らないので知らせない
    $messages = @($warnings | ForEach-Object { $_.Message })
    Assert-Equal 2 $messages.Count ($messages -join ' / ')
    Assert-Match $messages[0] '「電話番号」の桁数「10-11」を読めない'
    Assert-Match $messages[1] '「備考」の桁数「可変」を読めない'
}

Test-Case 'New-TestSpecDraft: -Screen に合う画面が無ければ設計書の画面 ID を挙げて止め、一部の ID だけが合わなければ知らせる' {
    $out = Join-Path (New-TestDir) 'UT.xlsx'
    Assert-Throws { & $TestDraft $designSample -OutXlsx $out -Screen 'SCR999' 6>$null } '-Screen に合う画面がありません: SCR999（設計書の画面: SCR001・SCR002・SCR003）'
    $cases = @(& $TestDraft $designSample -OutXlsx $out -Screen 'SCR002', 'SCR999' -WarningVariable warnings -WarningAction SilentlyContinue 6>$null)
    Assert-Equal @('SCR002 受注登録') @($cases | ForEach-Object { $_.Function } | Select-Object -Unique)
    Assert-Equal @('-Screen の SCR999 に合う画面がありません') @($warnings | ForEach-Object { $_.Message })
}

Test-Case 'New-TestSpecDraft: 画面設計書の関連要件（要件 ID）を、その画面のケースの要件ID の列に書き、要件からテストまでたどれるようにする' {
    $dir = New-TestDir
    $design = Join-Path $dir 'design'
    New-Item -ItemType Directory -Path $design | Out-Null
    $hdr = { param($id, $name) @('画面ID', $id, '画面名', $name, '', '', '') }
    foreach ($s in @(@('SCR010', '顧客登録', 'REQ-010、REQ-011'), @('SCR011', '顧客一覧', ''))) {
        $overview = @((& $hdr $s[0] $s[1]), @('画面概要', '顧客を扱う'))
        if ($s[2]) { $overview += , @('関連要件', $s[2]) }
        New-SheetsBook (Join-Path $design ('画面設計書_{0}.xlsx' -f $s[0])) ([ordered]@{
                '画面概要'     = $overview
                '画面項目定義' = @((& $hdr $s[0] $s[1]), @(), @('No', '項目名', '入出力', '型', '桁数', '必須', '入力チェック'), @('1', '顧客名', 'I', '全角', '20', '○', ''))
            })
    }
    $out = Join-Path $dir 'UT_下書き.xlsx'
    $cases = @(& $TestDraft $design -OutXlsx $out 6>$null)
    Assert-Equal @('SCR010|REQ-010、REQ-011', 'SCR011|') @($cases | Group-Object { $_.Id.Split('-')[0] } | ForEach-Object { '{0}|{1}' -f $_.Name, (@($_.Group | ForEach-Object { $_.Requirement } | Select-Object -Unique) -join '/') }) '画面ごとの要件 ID を、その画面のすべてのケースに'
    $sheet = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal '設計書|要件ID' ('{0}|{1}' -f $sheet.Rows[1]['M'], $sheet.Rows[1]['N'])
    Assert-Equal 'REQ-010、REQ-011' $sheet.Rows[2]['N']
    # Test-Traceability: 下書きのテスト仕様書で、画面設計書に書いた要件にテストがある
    @('要件ID,要件名', 'REQ-010,顧客を登録する', 'REQ-011,顧客名を必須にする', 'REQ-012,顧客を削除する') | Set-Content -LiteralPath (Join-Path $dir '要件一覧.csv') -Encoding UTF8
    $gaps = @(& (Join-Path $script:ToolsDir 'Test-Traceability.ps1') -DesignPath $design -TestPath $out -Requirement (Join-Path $dir '要件一覧.csv') 6>$null)
    Assert-Equal @('REQ-012|設計・テストなし') @($gaps | Where-Object Kind -eq '要件' | ForEach-Object { '{0}|{1}' -f $_.Id, $_.Judgement })
    # 要件 ID の無い設計書からは、要件ID の列を作らない
    $null = & $TestDraft $designSample -Screen SCR001 -OutXlsx (Join-Path $dir 'UT_SCR001.xlsx') 6>$null
    Assert-Equal '設計書|' ('{0}|{1}' -f (Read-XlsxSheet -Path (Join-Path $dir 'UT_SCR001.xlsx') -Index 2).Rows[1]['M'], (Read-XlsxSheet -Path (Join-Path $dir 'UT_SCR001.xlsx') -Index 2).Rows[1]['N'])
}

Test-Case 'New-TestSpecDraft: Get-Help でヘルプが表示される' {
    $h = Get-Help $TestDraft
    Assert-Match $h.Synopsis 'テスト仕様書の下書き'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}

Test-Case 'New-TestSpecDraft: 版を上げたあとの古い版の画面設計書は読まない（ケースを二重にしない）' {
    $dir = Copy-Fixture (Join-Path $designSample '.')
    $scr = Join-Path $dir 'DD02_画面設計書_SCR003_受注詳細_v1.0.xlsx'
    Move-Item -LiteralPath (Join-Path $dir 'DD02_画面設計書_SCR003_受注詳細.xlsx') -Destination $scr
    Copy-Item -LiteralPath $scr -Destination (Join-Path $dir 'DD02_画面設計書_SCR003_受注詳細_v1.1.xlsx')
    $expected = @(& $TestDraft $designSample -OutXlsx (Join-Path (New-TestDir) 'a.xlsx') 6>$null 3>$null)
    $w = $null
    $cases = @(& $TestDraft $dir -OutXlsx (Join-Path (New-TestDir) 'b.xlsx') -WarningVariable w -WarningAction SilentlyContinue 6>$null)
    Assert-Equal $expected.Count $cases.Count
    Assert-Match (($w | ForEach-Object { $_.Message }) -join "`n") '古い版 1 件は読みませんでした'
}
