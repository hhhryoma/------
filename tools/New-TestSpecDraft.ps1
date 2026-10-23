[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $DesignPath,

    [Parameter(Mandatory = $true)]
    [string] $OutXlsx,

    [string[]] $Screen,

    [switch] $Force,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

$OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ((Test-Path -LiteralPath $OutXlsx) -and -not $Force) { throw "同じ名前のファイルがあります（書き足したテスト仕様書を上書きしないよう、-Force を付けたときだけ上書きします）: $OutXlsx" }

Update-TypeData -TypeName 'TestSpecDraft.Case' -DefaultDisplayPropertySet Id, Viewpoint, Kind, Content, Expected -Force -WhatIf:$false

# ---------------------------------------------------------------- 設計書を読む


$searchArgs = @{ Pattern = '\S'; Path = $DesignPath; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$design = Read-ScreenDesign (ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null))
$screens = @($design.Items | ForEach-Object { $_.Screen }) + @($design.Events | ForEach-Object { $_.Screen }) | Select-Object -Unique
if ($Screen) { $screens = @($screens | Where-Object { $s = $_; @($Screen | Where-Object { $s -like $_ }).Count -gt 0 }) }
if (@($screens).Count -eq 0) { throw '画面設計書（画面ID と、項目名・桁数の画面項目定義、またはイベント定義）が見つかりません' }

# ---------------------------------------------------------------- ケースを作る


$selectKinds = '選択|ラジオ|チェック|プルダウン|リスト|コンボ'

function Format-Message {
    #     メッセージ ID と文言（{0} は項目名、{1} はチェックの値（桁数・「1以上」の 1・「受注日以降」の受注日））。
    #     メッセージ一覧に無い ID は ID だけ。項目名・値が無ければ {0}・{1} のまま
    param([string] $Id, [string] $ItemName, $Param)
    if (-not $Id) { return 'エラーメッセージ' }
    $m = $design.Messages[$Id.ToUpperInvariant()]
    if (-not $m) { return $Id }
    $text = $m.Text
    if ($ItemName) { $text = $text.Replace('{0}', $ItemName) }
    if ($null -ne $Param -and "$Param" -ne '') { $text = $text.Replace('{1}', [string]$Param) }
    return '{0}「{1}」' -f $Id, $text
}

function Find-MessageId {
    #     入力チェックの欄に ID が無いとき、メッセージ一覧の文言から探す（必須・桁数・半角英数 など）
    param([string] $Pattern)
    foreach ($id in ($design.Messages.Keys | Sort-Object)) { if ($id -like 'E*' -and $design.Messages[$id].Text -match $Pattern) { return $id } }
    return $null
}

function Split-Checks {
    #     入力チェックの欄（「必須（E0001）、半角英数（E0003）」）を、名前とメッセージ ID の組にする
    param([string] $Text)
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($part in ($Text -split '[、,，;；]')) {
        $p = $part.Trim()
        if (-not $p) { continue }
        $m = [regex]::Match($p, '^(?<name>.+?)\s*[（(]\s*(?<ids>[A-Za-z]+\d+(?:\s*[／/・,]\s*[A-Za-z]+\d+)*)\s*[）)]$')
        if ($m.Success) { $out.Add([pscustomobject]@{ Name = $m.Groups['name'].Value.Trim(); Id = ($m.Groups['ids'].Value -split '[／/・,]')[0].Trim().ToUpperInvariant() }) }
        else { $out.Add([pscustomobject]@{ Name = $p; Id = $null }) }
    }
    return , $out
}

function Get-BadValue {
    #     型のチェックに引っかかる値と、その説明
    param([string] $Name, $Item)
    $n = "$Name $($Item.Type)"
    if ($n -match '日付') { return '2026/02/30（存在しない日付）' }
    if ($n -match '時刻') { return '25:00（存在しない時刻）' }
    if ($n -match '半角数字|数字|数値|金額') { return 'abc（数字以外）' }
    if ($n -match '半角英数|英数') { return 'あいう（全角）' }
    if ($n -match 'カナ') { return 'あいう（ひらがな）' }
    if ($n -match '全角') { return 'ABC（半角）' }
    if ($n -match 'メール') { return 'abc（@ の無い値）' }
    if ($n -match '電話') { return 'abc-defg（数字以外）' }
    return '形式に合わない値'
}

$cases = New-Object System.Collections.Generic.List[object]
$sheets = New-Object System.Collections.Generic.List[object]
foreach ($scr in $screens) {
    $items = @($design.Items | Where-Object { $_.Screen -eq $scr })
    $events = @($design.Events | Where-Object { $_.Screen -eq $scr })
    $name = @(@($items | ForEach-Object { $_.ScreenName }) + @($events | ForEach-Object { $_.ScreenName }) | Where-Object { $_ })[0]
    $function = ('{0} {1}' -f $scr, $name).Trim()
    $list = New-Object System.Collections.Generic.List[object]
    $usedIds = @{}
    $add = {
        param([string] $Viewpoint, [string] $Kind, [string] $Content, [string] $Steps, [string] $InputValue, [string] $Expected, $Source)
        $c = [pscustomobject]@{
            Id = '{0}-{1:000}' -f $scr, ($list.Count + 1); Function = $function; Viewpoint = $Viewpoint; Kind = $Kind; Content = $Content
            Steps = $Steps; InputValue = $InputValue; Expected = $Expected; Source = $(if ($Source) { '{0} {1} {2} 行目' -f $Source.File, $Source.Sheet, $Source.Row } else { '' })
        }
        $c.PSObject.TypeNames.Insert(0, 'TestSpecDraft.Case')
        $list.Add($c)
    }

    # 入力チェックを行うイベント（「入力チェック」と書かれた処理）と、項目ごとのイベント（「顧客コード入力」など）
    $checkEvent = @($events | Where-Object { $_.Summary -match '入力チェック|チェックを行う' })[0]
    $checkStep = '登録する'
    if ($checkEvent) { $checkStep = $(if ($checkEvent.Trigger) { $checkEvent.Trigger } else { $checkEvent.Name }) }

    # ---- 初期表示
    foreach ($ev in @($events | Where-Object { $_.Name -match '初期表示' -or $_.Trigger -match '表示時' })) {
        $initial = @($items | Where-Object { $_.Initial } | ForEach-Object { '{0}: {1}' -f $_.Item, $_.Initial })
        $expected = $ev.Summary
        if ($initial.Count -gt 0) { $expected = ($expected + "`n初期値: " + ($initial -join '、')).Trim() }
        & $add 'イベント' '正常' ('{0}（{1}）' -f $ev.Name, $ev.Id) '画面を表示する' '' $expected $ev
    }

    # ---- 入力項目
    foreach ($it in @($items | Where-Object { $_.Input -and $_.Kind -notmatch 'ボタン|リンク' })) {
        $checks = Split-Checks $it.Checks
        $codes = $design.Codes[(ConvertTo-Label $it.Item)]
        $isSelect = $codes -or $it.Kind -match $selectKinds
        $isText = -not $isSelect -and $it.Type -notmatch '日付|時刻' -and $null -ne $it.Digits
        # 項目ごとのイベント（フォーカスアウトなど）があれば、そのトリガーで確かめる
        $itemEvent = @($events | Where-Object { $_.Name -match [regex]::Escape($it.Item) -or $_.Trigger -match [regex]::Escape($it.Item) })[0]
        $stepOf = { param($value, $useItemEvent) $trigger = $checkStep; if ($useItemEvent -and $itemEvent -and $itemEvent.Trigger) { $trigger = $itemEvent.Trigger }; '{0}に{1}、{2}' -f $it.Item, $value, $trigger }
        $handled = @{}

        # 必須
        $req = @($checks | Where-Object { $_.Name -match '必須' })[0]
        if ($it.Required -or $req) {
            $id = $(if ($req -and $req.Id) { $req.Id } else { Find-MessageId '必須' })
            if ($id) { $usedIds[$id] = $true }
            & $add '必須' '異常' ('{0}: 未入力（必須）' -f $it.Item) ('{0}を未入力にして、{1}' -f $it.Item, $checkStep) '（空欄）' ('{0} が表示され、処理が進まない' -f (Format-Message $id $it.Item $it.Digits)) $it
            if ($req) { $handled[$req.Name] = $true }
        }
        # 桁数の境界
        if ($isText) {
            $len = @($checks | Where-Object { $_.Name -match '桁|文字数|長さ' })[0]
            $id = $(if ($len -and $len.Id) { $len.Id } else { Find-MessageId '文字以内|桁以内|桁数' })
            if ($id) { $usedIds[$id] = $true }
            $unit = $(if ($it.Type -match '数') { '桁' } else { '文字' })
            & $add '桁数' '境界' ('{0}: 最大の {1} {2}' -f $it.Item, $it.Digits, $unit) (& $stepOf ('{0} {1}を入力し' -f $it.Digits, $unit) $false) ('{0} {1}' -f $it.Digits, $unit) 'エラーにならず、入力した値のまま処理される' $it
            & $add '桁数' '境界' ('{0}: 最大を超える {1} {2}' -f $it.Item, ($it.Digits + 1), $unit) (& $stepOf ('{0} {1}を入力し' -f ($it.Digits + 1), $unit) $false) ('{0} {1}' -f ($it.Digits + 1), $unit) ('{0} が表示されるか、{1} {2}目が入力できない' -f (Format-Message $id $it.Item $it.Digits), ($it.Digits + 1), $unit) $it
            if ($len) { $handled[$len.Name] = $true }
        }
        # 型（入力チェックに書かれた形式、無ければ型から）
        $typeChecks = @($checks | Where-Object { $_.Name -match '半角|全角|英数|数字|数値|日付|時刻|形式|カナ|メール|電話' })
        if ($typeChecks.Count -eq 0 -and -not $isSelect -and $it.Type -match '日付|時刻|半角|数値|数字|英数|カナ') { $typeChecks = @([pscustomobject]@{ Name = $it.Type; Id = $null }) }
        foreach ($tc in $typeChecks) {
            $bad = Get-BadValue $tc.Name $it
            $id = $tc.Id
            if (-not $id) { $id = Find-MessageId ([regex]::Escape(($tc.Name -replace '形式', ''))) }
            if ($id) { $usedIds[$id] = $true }
            & $add '型' '異常' ('{0}: {1}' -f $it.Item, $tc.Name) (& $stepOf ('{0}を入力し' -f $bad) $false) $bad ('{0} が表示され、処理が進まない' -f (Format-Message $id $it.Item $null)) $it
            $handled[$tc.Name] = $true
        }
        # そのほかの入力チェック（存在・範囲・日付の前後 など）
        foreach ($ck in @($checks | Where-Object { -not $handled.ContainsKey($_.Name) })) {
            $value = '「{0}」を満たさない値' -f $ck.Name
            $param = $null
            $useItemEvent = $false
            if ($ck.Name -match '存在') { $value = '登録されていない値'; $useItemEvent = $true }
            elseif ($ck.Name -match '(\d+)\s*以上') { $param = $Matches[1]; $value = [string]([int]$Matches[1] - 1) }
            elseif ($ck.Name -match '(\d+)\s*以下') { $param = $Matches[1]; $value = [string]([int]$Matches[1] + 1) }
            elseif ($ck.Name -match '(.+)以降') { $param = $Matches[1]; $value = '{0}より前の日付' -f $Matches[1] }
            elseif ($ck.Name -match '(.+)以前') { $param = $Matches[1]; $value = '{0}より後の日付' -f $Matches[1] }
            if ($ck.Id) { $usedIds[$ck.Id] = $true }
            & $add '入力チェック' '異常' ('{0}: {1}' -f $it.Item, $ck.Name) (& $stepOf ('{0}を入力し' -f $value) $useItemEvent) $value ('{0} が表示され、処理が進まない' -f (Format-Message $ck.Id $it.Item $param)) $it
        }
        # 選ぶ項目: コード値ごと（コード定義書に同じ名前のコード区分があるとき）
        if ($codes) {
            foreach ($v in $codes) {
                & $add 'コード値' '正常' ('{0}: 「{1}」を選ぶ' -f $it.Item, $v) ('{0}で「{1}」を選び、{2}' -f $it.Item, $v, $checkStep) $v ('「{0}」で処理される' -f $v) $it
            }
        }
        elseif ($it.Kind -match 'チェック') {
            $note = $(if ($it.Note) { '（{0}）' -f $it.Note } else { '' })
            & $add 'チェック' '正常' ('{0}: チェックする' -f $it.Item) ('{0}をチェックして、{1}' -f $it.Item, $checkStep) 'チェックあり' ('チェックしたときの動きになる{0}' -f $note) $it
            & $add 'チェック' '正常' ('{0}: チェックしない' -f $it.Item) ('{0}をチェックせずに、{1}' -f $it.Item, $checkStep) 'チェックなし' 'チェックしないときの動きになる' $it
        }
    }

    # ---- イベント（初期表示以外）
    foreach ($ev in @($events | Where-Object { -not ($_.Name -match '初期表示' -or $_.Trigger -match '表示時') })) {
        $trigger = $(if ($ev.Trigger) { $ev.Trigger } else { $ev.Name })
        $expected = $ev.Summary
        if ($ev.Next) { $expected = ($expected + "`n遷移先: " + $ev.Next).Trim() }
        & $add 'イベント' '正常' ('{0}（{1}）: 正しい値で実行する' -f $ev.Name, $ev.Id) $trigger '' $expected $ev
        $ids = @([regex]::Matches($ev.Messages, '[A-Za-z]+\d+') | ForEach-Object { $_.Value.ToUpperInvariant() } | Select-Object -Unique)
        # 確認メッセージ: キャンセルしたとき
        foreach ($q in @($ids | Where-Object { $_ -like 'Q*' })) {
            & $add 'イベント' '異常' ('{0}（{1}）: 確認で「キャンセル」を選ぶ' -f $ev.Name, $ev.Id) ('{0}し、{1} で「キャンセル」を選ぶ' -f $trigger, (Format-Message $q '' $null)) '' '処理を中断し、画面は操作する前のまま' $ev
        }
        # エラーメッセージのうち、入力項目のケースで確かめていないもの（在庫不足・排他など）。なる条件は設計書から書き足す
        foreach ($e in @($ids | Where-Object { $_ -like 'E*' -and -not $usedIds.ContainsKey($_) })) {
            $usedIds[$e] = $true
            & $add 'イベント' '異常' ('{0}（{1}）: {2} になる場合' -f $ev.Name, $ev.Id, $e) ('{0} になる条件（処理概要から書き足す）で、{1}' -f $e, $trigger) '' ('{0} が表示され、処理が進まない' -f (Format-Message $e '' $null)) $ev
        }
    }
    foreach ($c in $list) { $cases.Add($c) }
    $sheets.Add([pscustomobject]@{ Screen = $scr; Function = $function; Cases = $list })
}

# ---------------------------------------------------------------- 保存


$columns = @(
    @{ Header = 'テストID'; Width = 13 }, @{ Header = '機能'; Width = 18 }, @{ Header = '観点'; Width = 10 }, @{ Header = '区分'; Width = 7 },
    @{ Header = '確認内容'; Width = 36 }, @{ Header = '手順'; Width = 40 }, @{ Header = '入力値'; Width = 18 }, @{ Header = '期待結果'; Width = 50 },
    @{ Header = '予定日'; Width = 11 }, @{ Header = '実施日'; Width = 11 }, @{ Header = '結果'; Width = 7 }, @{ Header = '障害番号'; Width = 10 }, @{ Header = '設計書'; Width = 40 })
$bookSheets = New-Object System.Collections.Generic.List[object]
$summary = New-Object System.Collections.Generic.List[object]
foreach ($s in $sheets) {
    $rows = @($s.Cases | ForEach-Object { @{ Values = @($_.Id, $_.Function, $_.Viewpoint, $_.Kind, $_.Content, $_.Steps, $_.InputValue, $_.Expected, '', '', '', '', $_.Source) } })
    $bookSheets.Add(@{ Name = $s.Screen; Columns = $columns; Rows = $rows })
    $count = { param($k) @($s.Cases | Where-Object { $_.Kind -eq $k }).Count }
    $summary.Add(@{ Values = @($s.Function, [double]$s.Cases.Count, [double](& $count '正常'), [double](& $count '異常'), [double](& $count '境界')) })
}
$summarySheet = @{ Name = '一覧'; Columns = @(@{ Header = '機能'; Width = 30 }, @{ Header = 'ケース'; Width = 8 }, @{ Header = '正常'; Width = 8 }, @{ Header = '異常'; Width = 8 }, @{ Header = '境界'; Width = 8 }); Rows = $summary.ToArray() }
Export-Workbook -FilePath $OutXlsx -Sheets (@($summarySheet) + $bookSheets.ToArray())

$cases
Write-Host ('テスト仕様書の下書きを作りました: {0} 画面・{1} ケース → {2}' -f $sheets.Count, $cases.Count, $OutXlsx)
Write-Host '下書きです。「（処理概要から書き足す）」の条件や、画面固有の確認を書き足してから使ってください。Test-TestSpec.ps1 -DesignPath で抜けを確かめられます。'

# .SYNOPSIS
#     画面設計書（画面項目定義・イベント定義）から、テスト仕様書の下書き（単体テストのケース）を作ります（Excel 不要）。
#
# .DESCRIPTION
#     画面設計書の画面項目定義の必須・型・桁数・入力チェックと、コード定義書のコード値、イベント定義の処理概要・メッセージ・遷移先から、
#     画面ごとのテストケースを Excel（画面 ID ごとのシート）に書き出します。期待結果には、メッセージ一覧の文言（{0} に項目名）を入れます。
#
#       ・入力項目: 必須の未入力、桁数の境界（最大・最大＋1）、型の不正な値（日付・数字・半角英数 など）、
#         入力チェック欄のそのほかのチェック（存在・「1以上」・「受注日以降」など）、コード値ごとの選択、チェックボックスのあり・なし
#       ・イベント: 初期表示（初期値）、各イベントの正常な実行（処理概要・遷移先）、確認メッセージのキャンセル、
#         入力項目のケースで確かめていないエラーメッセージ（在庫不足など。なる条件は書き足す）
#
#     列は テストID・機能・観点・区分（正常・異常・境界）・確認内容・手順・入力値・期待結果・予定日・実施日・結果・障害番号・設計書 で、
#     Test-TestSpec.ps1・Test-Traceability.ps1・Measure-TestQuality.ps1 でそのまま読めます。
#
# .PARAMETER DesignPath
#     画面設計書と、コード定義書・メッセージ一覧のファイルかフォルダ（いっしょに渡すと、コード値とメッセージの文言を使います）。
#
# .PARAMETER OutXlsx
#     書き出すテスト仕様書の下書き（.xlsx）。同じ名前のファイルがあるときは -Force を付けたときだけ上書きします。
#
# .PARAMETER Screen
#     下書きを作る画面 ID（ワイルドカード可。複数可）。既定はすべての画面。
#
# .PARAMETER Force
#     -OutXlsx のファイルがあっても上書きします。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読めません）。
#
# .EXAMPLE
#     .\New-TestSpecDraft.ps1 D:\設計書 -OutXlsx D:\テスト\UT_テスト仕様書_下書き.xlsx
#
# .EXAMPLE
#     # 受注登録の画面だけ
#     .\New-TestSpecDraft.ps1 D:\設計書 -Screen SCR002 -OutXlsx D:\テスト\UT_SCR002.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに Search-Excel.ps1 と ExcelTools.psm1 が必要です。
