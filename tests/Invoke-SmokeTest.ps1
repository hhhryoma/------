[CmdletBinding()]
param(
    [string] $OutDir,

    [switch] $NoOpen
)

$ErrorActionPreference = 'Stop'

$repo = Split-Path -Path $PSScriptRoot -Parent
$tools = Join-Path $repo 'tools'
$design = Join-Path $repo 'docs/design-sample'
$fixtures = Join-Path $PSScriptRoot 'fixtures'
Import-Module (Join-Path $tools 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $OutDir) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if (-not $docs) { $docs = [System.IO.Path]::GetTempPath() }
    $OutDir = Join-Path $docs ('Excel ツールの動作確認_{0}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
$data = Join-Path $OutDir '資料'
New-Item -ItemType Directory -Force -Path $data | Out-Null
# Excel が要るツール（PDF にする）は、Excel の入った Windows だけで実行する
$hasExcel = [System.Environment]::OSVersion.Platform -eq [System.PlatformID]::Win32NT -and $null -ne [type]::GetTypeFromProgID('Excel.Application')

# ---------------------------------------------------------------- 確かめるための資料（サンプルの設計書に合わせた指摘・テスト・障害・WBS）


function New-Book {
    param([string] $FilePath, [string] $Sheet, [string[]] $Header, $Rows)
    Export-Workbook -FilePath $FilePath -Sheets @(@{ Name = $Sheet; Columns = @($Header | ForEach-Object { @{ Header = $_; Width = 14 } }); Rows = @($Rows | ForEach-Object { @{ Values = $_ } }) })
}

function New-WordDocx {
    #     Word で開ける最小の文書を作る。Body: 文字は段落、配列の配列は罫線付きの表（行ごとのセル）
    #     Comment は最初の段落へのコメント、Inserted は「以上」の前に挿入した変更履歴（どちらも消し忘れの例）
    param([string] $FilePath, [object[]] $Body, [string] $Comment, [string] $Inserted)
    $w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    $run = { param($t) '<w:r><w:t xml:space="preserve">' + [System.Security.SecurityElement]::Escape($t) + '</w:t></w:r>' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in $Body) {
        if ($b -is [string]) {
            if ($Inserted -and $b -eq '以上') { [void]$sb.Append('<w:p><w:ins w:id="1" w:author="山田" w:date="2026-10-20T10:00:00Z">' + (& $run $Inserted) + '</w:ins></w:p>') }
            if ($Comment -and $sb.Length -eq 0) { [void]$sb.Append('<w:p><w:commentRangeStart w:id="0"/>' + (& $run $b) + '<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>') }
            else { [void]$sb.Append('<w:p>' + (& $run $b) + '</w:p>') }
            continue
        }
        $cols = ($b | ForEach-Object { @($_).Count } | Measure-Object -Maximum).Maximum
        $borders = (@('top', 'left', 'bottom', 'right', 'insideH', 'insideV') | ForEach-Object { '<w:{0} w:val="single" w:sz="4" w:space="0" w:color="auto"/>' -f $_ }) -join ''
        [void]$sb.Append('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>' + $borders + '</w:tblBorders></w:tblPr><w:tblGrid>' + ('<w:gridCol w:w="2400"/>' * $cols) + '</w:tblGrid>')
        foreach ($row in $b) { [void]$sb.Append('<w:tr>' + ((@($row | ForEach-Object { '<w:tc><w:p>' + (& $run $_) + '</w:p></w:tc>' })) -join '') + '</w:tr>') }
        [void]$sb.Append('</w:tbl>')
    }
    $parts = [ordered]@{}
    $types = '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    $rels = ''
    if ($Comment) {
        $types += '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>'
        $rels = '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>'
        $parts['word/comments.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments ' + $w + '><w:comment w:id="0" w:author="佐藤" w:date="2026-10-20T10:00:00Z" w:initials="S"><w:p>' + (& $run $Comment) + '</w:p></w:comment></w:comments>'
    }
    $parts['[Content_Types].xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + $types + '</Types>'
    $parts['_rels/.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    $parts['word/_rels/document.xml.rels'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rels + '</Relationships>'
    $parts['word/document.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ' + $w + '><w:body>' + $sb.ToString() + '</w:body></w:document>'
    $fs = [System.IO.File]::Create($FilePath)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($k in $parts.Keys) {
            $sw = New-Object System.IO.StreamWriter($zip.CreateEntry($k).Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($parts[$k]); $sw.Dispose()
        }
        $zip.Dispose()
    }
    finally { $fs.Dispose() }
}

$scr1 = 'DD02_画面設計書_SCR001_受注一覧'
$scr2 = 'DD02_画面設計書_SCR002_受注登録'
New-Item -ItemType Directory -Force -Path (Join-Path $data 'レビュー'), (Join-Path $data 'テスト') | Out-Null
New-Book (Join-Path $data 'レビュー/指摘管理表.xlsx') '指摘' @('No', '指摘日', '対象成果物', '指摘内容', '指摘区分', '工程', '対応期限', '対応状況') @(
    , @('1', '2026/10/01', $scr1, '受注ステータスの桁数が未記入', '漏れ', '詳細設計レビュー', '2026/10/10', '完了')
    , @('2', '2026/10/01', $scr1, 'イベントが足りない', '漏れ', '詳細設計レビュー', '2026/10/05', '対応中')
    , @('3', '2026/10/02', $scr2, '支払方法の初期値が書かれていない', '漏れ', '詳細設計レビュー', '2026/10/30', '未対応')
    , @('4', '2026/10/02', 'DD03_テーブル定義書', '型が違う', '誤り', '詳細設計レビュー', '2026/10/30', '完了'))
@('設計書,レビュー工数', "$scr1,2", "$scr2,1.5", 'DD03_テーブル定義書,1') | Set-Content -LiteralPath (Join-Path $data 'レビュー記録.csv') -Encoding UTF8
New-Book (Join-Path $data 'AIのレビュー結果.xlsx') '指摘一覧' @('No', '判定', '観点ID', '設計書', '場所', '指摘内容') @(
    , @('1', 'NG', 'S01', "$scr1.xlsx", '画面項目定義!AB11', '受注ステータスの桁数が無い')
    , @('2', 'NG', 'S02', "$scr1.xlsx", '画面項目定義!AK7', '数量のエラー時のメッセージが無い')
    , @('3', 'NG', 'S02', "$scr1.xlsx", '画面項目定義!AK8', '単価のエラー時のメッセージが無い'))

# テスト仕様書（60 ケース、10/01〜10/12 に 1 日 5 件）と、収束していく障害管理表
$cases = New-Object System.Collections.Generic.List[object]
for ($i = 1; $i -le 60; $i++) {
    $function = ''
    if ($i -eq 1) { $function = 'SCR001' } elseif ($i -eq 31) { $function = 'SCR002' } elseif ($i -eq 51) { $function = 'SCR003' }
    $day = (Get-Date '2026-10-01').AddDays([math]::Floor(($i - 1) / 5)).ToString('yyyy/MM/dd')
    $kind = $(if ($i % 4 -eq 0) { '異常' } else { '正常' })
    $text = "確認 $i"; $expected = "結果 $i が表示される"; $result = 'OK'; $bug = ''
    switch ($i) {
        2 { $text = '受注番号に 13 桁を入れて検索する'; $expected = 'E0002 が表示される' }
        32 { $text = '受注日を未入力で登録する'; $expected = 'E0001「受注日は必須入力です。」が表示される' }
        36 { $text = '受注日に存在しない日付を入れて登録する'; $expected = 'E0004「日付が正しくありません。」が表示される' }
        40 { $expected = '正しく登録されること' }
        44 { $result = 'NG'; $bug = 'B003' }
    }
    if ($i -gt 56) { $result = '' ; $day2 = '' } else { $day2 = $day }
    $cases.Add(@([string]$i, $function, $kind, $text, '画面を開く', $expected, $day, $day2, $result, $bug))
}
New-Book (Join-Path $data 'テスト/UT_受注.xlsx') 'テストケース' @('No', '機能', '区分', '確認内容', '手順', '期待結果', '予定日', '実施日', '結果', '障害番号') $cases
# エビデンス（SCR001 の No.1〜3 の画面の代わりの画像。No.3 は No.1 と同じ画像＝使い回し）
$pngA = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAYElEQVR4nO2XMQpAIQxDE+n9r1wH4QviomALv8kidXk0pEXp7shQS6ECsHGQDEMOj22pX+vrMM1qgcNkm7ujhN9Gsp7VAv8fvBunkKVdz2qBBRZY4GvNXR35pgfAcn+nDgTjEEBUFCIpAAAAAElFTkSuQmCC'
$pngB = 'iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAAbUlEQVR4nO2XwQoAIQhEZxb//5ftEHQIYSlaXdK5RF0eYzoUVRURekKoAKQvJN2QvcYy7b/WcBhW6gK7SYyzpQ7fbcl8pS7w/WBrnFxC2wJv6XX4Jzv57rjABb4PfCxAVuPuB4493/QAmO7v1ACh5BRCnP+rmgAAAABJRU5ErkJggg=='
New-Item -ItemType Directory -Force -Path (Join-Path $data 'エビデンス/SCR001') | Out-Null
foreach ($e in @(@('1', $pngA), @('2', $pngB), @('3', $pngA))) { [System.IO.File]::WriteAllBytes((Join-Path $data ('エビデンス/SCR001/{0}.png' -f $e[0])), [Convert]::FromBase64String($e[1])) }
$daily = @(1, 2, 2, 1, 1, 0, 1, 0, 0, 0)
$bugs = New-Object System.Collections.Generic.List[object]
$n = 0
for ($d = 0; $d -lt $daily.Count; $d++) {
    for ($j = 0; $j -lt $daily[$d]; $j++) {
        $n++
        $found = (Get-Date '2026-10-01').AddDays($d)
        $function = $(if ($n -le 5) { 'SCR001' } else { 'SCR002' })
        $status = $(if ($n -le 6) { '完了' } else { '対応中' })
        $closed = $(if ($n -le 6) { $found.AddDays(2).ToString('yyyy/MM/dd') } else { '' })
        $bugs.Add(@(('B{0:000}' -f $n), $function, "障害 $n", @('詳細設計', '製造')[$n % 2], '単体テスト', $found.ToString('yyyy/MM/dd'), $closed, $status))
    }
}
New-Book (Join-Path $data '障害管理表.xlsx') '障害' @('障害番号', '機能', '障害内容', '混入工程', '摘出工程', '起票日', '完了日', '状態') $bugs
@('要件ID,要件名', 'REQ-001,受注を検索する', 'REQ-002,受注を登録する') | Set-Content -LiteralPath (Join-Path $data '要件一覧.csv') -Encoding UTF8

New-Book (Join-Path $data 'WBS.xlsx') 'WBS' @('WBS', 'タスク名', '担当者', '予定', '', '実績', '', '進捗率', '予定工数', '実績工数') @(
    , @('', '', '', '開始', '終了', '開始', '終了', '', '', '')
    , @('1', '詳細設計', '', '2026/10/01', '2026/10/16', '', '', '', '', '')
    , @('1.1', '画面設計', '山田', '2026/10/01', '2026/10/07', '2026/10/01', '2026/10/09', '100%', '5', '7')
    , @('1.2', 'テーブル設計', '佐藤', '2026/10/05', '2026/10/09', '2026/10/05', '', '60%', '5', '4')
    , @('1.3', 'IF 設計', '佐藤', '2026/10/12', '2026/10/16', '', '', '0%', '4', ''))
@('定例会議 議事録', '日時：2026/10/20', '', '■ 決定事項', '・受注日の初期値は当日とする', '', '■ 宿題', '・画面設計書の受注日の桁数を直す（山田、10/23まで）', '・テーブル定義書のレビュー日程を決める　担当：佐藤　期限：10/27', '・お客様に締め日の扱いを確認する（来週中）') |
    Set-Content -LiteralPath (Join-Path $data '20261020_定例議事録.txt') -Encoding UTF8
New-Book (Join-Path $data '工数実績.xlsx') '10月' @('日付', '氏名', 'WBS', '作業内容', '時間') @(
    , @('2026/10/01', '山田', '1.1', '画面設計', '8')
    , @('2026/10/02', '山田', '1.1', '', '7:30')
    , @('2026/10/05', '佐藤', '1.2', 'テーブル設計', '8')
    , @('2026/10/06', '佐藤', '1.2', '', '8')
    , @('', '', '', '定例会議', '1'))

# 書き換えるツール用の写し（元のサンプルは変えない）
$copies = Join-Path $data '書き換え用'
New-Item -ItemType Directory -Force -Path $copies | Out-Null
Copy-Item -LiteralPath (Join-Path $fixtures 'deliverable/bad.xlsx') -Destination (Join-Path $copies '表示を整える.xlsx')
Copy-Item -LiteralPath (Join-Path $fixtures 'replace/table.xlsx') -Destination (Join-Path $copies '置換.xlsx')
# 設計書の新版（SCR002 の支払方法の初期値と E0004 の文言を変えたもの。再テストするケースを選ぶ）
$designV2 = Join-Path $data '設計書_新版'
New-Item -ItemType Directory -Force -Path $designV2 | Out-Null
Get-ChildItem -LiteralPath $design -Filter '*.xlsx' | Copy-Item -Destination $designV2
$null = & (Join-Path $tools 'Replace-Excel.ps1') -Pattern '銀行振込' -Replacement 'クレジットカード' -SimpleMatch -Path (Join-Path $designV2 "$scr2.xlsx") -SheetNamePattern '画面項目定義' -Force 6>$null
$null = & (Join-Path $tools 'Replace-Excel.ps1') -Pattern '日付の形式が正しくありません' -Replacement '日付の形式が違います' -SimpleMatch -Path (Join-Path $designV2 'DD05_メッセージ一覧.xlsx') -Force 6>$null
# 版数の付いた設計書（版を続けて比べる）。Word は 3 版で、v1.2 は変更履歴とコメントを消し忘れたもの（提出前チェックで見つける）
$versions = Join-Path $data '版'
New-Item -ItemType Directory -Force -Path $versions | Out-Null
$wordV11 = @('1. 概要', '受注を登録・更新する画面である。受注番号は自動で採番する。', @(@('項目', '桁数', '必須'), @('受注番号', '12', '○'), @('受注日', '8', '○'), @('顧客', '20', '')), '2. 補足', '取消は当日中のみ可能とする。', '以上')
New-WordDocx (Join-Path $versions '基本設計書_v1.0.docx') @('1. 概要', '受注を登録する画面である。受注番号は自動で採番する。', @(@('項目', '桁数', '必須'), @('受注番号', '10', '○'), @('受注日', '8', '○')), '2. 補足', '取消は当日中のみ可能とする。', '以上')
New-WordDocx (Join-Path $versions '基本設計書_v1.1.docx') $wordV11
New-WordDocx (Join-Path $versions '基本設計書_v1.2.docx') @($wordV11 | ForEach-Object { if ($_ -eq '取消は当日中のみ可能とする。') { '取消は翌日まで可能とする。' } else { , $_ } }) -Comment 'ここは仮です' -Inserted '締め日の扱いは TODO'
Copy-Item -LiteralPath (Join-Path $design "$scr2.xlsx") -Destination (Join-Path $versions "${scr2}_v1.0.xlsx")
Copy-Item -LiteralPath (Join-Path $designV2 "$scr2.xlsx") -Destination (Join-Path $versions "${scr2}_v1.1.xlsx")

# ソース（設計書の ID と突き合わせる。E0099 はメッセージ一覧に無い ID）
$source = Join-Path $data 'ソース'
New-Item -ItemType Directory -Force -Path $source | Out-Null
@('public class OrderList {', '    void search() { error("E0001"); error("E0099"); }', '    String sql = "SELECT * FROM T_ORDER";', '}') | Set-Content -LiteralPath (Join-Path $source 'OrderList.java') -Encoding UTF8

# ---------------------------------------------------------------- ツールを順に実行する


$steps = New-Object System.Collections.Generic.List[object]
function Invoke-Step {
    #     ツールを 1 つ実行し、結果のファイルと「Excel で開いて見るところ」を確認リストに足す。失敗しても次へ進む
    param([string] $Name, [string] $File, [string] $Check, [scriptblock] $Run, [switch] $NeedsExcel)
    $path = $(if ($File) { Join-Path $OutDir $File } else { '' })
    $status = '成功'; $message = ''
    if ($NeedsExcel -and -not $hasExcel) { $status = 'スキップ'; $message = 'Excel が必要です（Excel の入った Windows で実行してください）' }
    else {
        Write-Host ('  {0} …' -f $Name)
        try {
            $null = & $Run $path 3>$null 6>$null
            if ($path -and -not (Test-Path -LiteralPath $path)) { $status = '失敗'; $message = '結果のファイルができていません' }
        }
        catch { $status = '失敗'; $message = $_.Exception.Message }
    }
    $steps.Add([pscustomobject]@{ No = $steps.Count + 1; Name = $Name; File = $File; Path = $path; Check = $Check; Status = $status; Message = $message })
}

$t = { param($name) Join-Path $tools $name }
Write-Host "動作確認の結果を作っています: $OutDir"
Invoke-Step 'Search-Excel（検索）' '01_検索.xlsx' 'リンクをクリックすると、該当のブック・シート・セルが開く。Word（基本設計書_v1.2.docx）の行は「本文!段落 n」で、リンクで文書が開く' { param($o) & (& $t 'Search-Excel.ps1') -Pattern 'TODO' -Path (Join-Path $fixtures 'grep'), $versions -IncludeWord -NoExcel -OutXlsx $o }
Invoke-Step 'Test-ExcelDeliverable（提出前チェック）' '02_提出前チェック.xlsx' '指摘一覧のリンクで、該当のブックのセルが開く（ファイル別のシートもある）。Word（基本設計書_v1.2.docx）の変更履歴・コメント・残っている文言（TODO）の行があり、リンクで文書が開く。Word で開いても同じ変更履歴とコメントが見える' { param($o) & (& $t 'Test-ExcelDeliverable.ps1') -Path (Join-Path $fixtures 'deliverable/bad.xlsx'), (Join-Path $fixtures 'print'), (Join-Path $versions '基本設計書_v1.2.docx') -OutXlsx $o }
Invoke-Step 'New-DocumentRegister（成果物一覧）' '02b_成果物一覧.xlsx' 'ファイルのリンクで文書が開く。文書名・版数・最新の改訂が読める。古い版のシートに 基本設計書 の v1.0・v1.1 と SCR002 の v1.0 がある' { param($o) & (& $t 'New-DocumentRegister.ps1') -Path $versions -OutXlsx $o }
Invoke-Step 'Test-ExcelTerms（表記ゆれ）' '03_表記ゆれ.xlsx' '表記ゆれ・出現箇所のシートが読める（文字化け・列幅）' { param($o) & (& $t 'Test-ExcelTerms.ps1') -Path $design -OutXlsx $o }
Invoke-Step 'Compare-Excel（差分）' '04_差分.xlsx' '差分一覧のリンクが開く。04_差分_色付け フォルダの写しで、変更のあったセルに色が付いている' { param($o) & (& $t 'Compare-Excel.ps1') -Old (Join-Path $fixtures 'compare/v1') -New (Join-Path $fixtures 'compare/v2') -OutXlsx $o -OutMarkedDir (Join-Path $OutDir '04_差分_色付け') }
Invoke-Step 'Compare-Excel（差分をブラウザーで見る）' '04b_差分.html' 'ブラウザーで開き、変わったセルに色が付く・マウスを重ねると変更前後が出る・左右に並べるとスクロールが連動する・一覧のクリックで該当セルへ移る' { param($o) & (& $t 'Compare-Excel.ps1') -Old $design -New $designV2 -OutHtml $o }
Invoke-Step 'Compare-Excel（版を続けて比べる・Word）' '04c_版の差分.html' '一覧に 基本設計書（v1.0 → v1.1）・（v1.1 → v1.2）と SCR002（v1.0 → v1.1）が版の順に並ぶ。Word の段落は折り返して表示され、表の行の追加で左右の行がずれない。検索欄（/ キー）で「受注」を探すと、ファイルをまたいで次へ進む' { param($o) & (& $t 'Compare-Excel.ps1') -Versions $versions -OutHtml $o }
Invoke-Step 'Test-DesignConsistency（設計書の整合性）' '05_整合性.xlsx' '指摘一覧のリンクで設計書の該当セルが開く。CRUD のシートが読める' { param($o) & (& $t 'Test-DesignConsistency.ps1') -Path (Join-Path $fixtures 'consistency') -OutXlsx $o }
Invoke-Step 'Test-Traceability（網羅）' '06_網羅.xlsx' '抜け・網羅・要件のシートと、場所のリンク' { param($o) & (& $t 'Test-Traceability.ps1') -DesignPath $design -TestPath (Join-Path $data 'テスト') -Requirement (Join-Path $data '要件一覧.csv') -OutXlsx $o }
Invoke-Step 'Select-RegressionTests（再テストするケース）' '06b_再テスト.xlsx' '選んだ理由の列と、テスト仕様書の場所のリンク（該当セルが開く）' { param($o) & (& $t 'Select-RegressionTests.ps1') -Old $design -New $designV2 -TestPath (Join-Path $data 'テスト') -OutXlsx $o }
Invoke-Step 'Test-CodeTrace（設計書とソース）' '06c_設計とソース.xlsx' '指摘のリンクで、メッセージ一覧の該当行と OrderList.java が開く。対応表の E0001 の行にソースの場所がある' { param($o) & (& $t 'Test-CodeTrace.ps1') -DesignPath $design -SourcePath $source -OutXlsx $o }
Invoke-Step 'New-TestData（テストデータ）' '09b_テストデータ.xlsx' 'テーブルごとのシートに標準・最大・最小・NULL の行がある。09b_テストデータ フォルダの CSV を Excel で開いても文字化けしない' { param($o) & (& $t 'New-TestData.ps1') -Path $design -OutDir (Join-Path $OutDir '09b_テストデータ') -BaseDate '2026-04-01' -OutXlsx $o }
Invoke-Step 'Measure-DesignSize（規模）' '07_規模.xlsx' '集計・規模一覧・係数表のシート' { param($o) & (& $t 'Measure-DesignSize.ps1') -Path $design -OutXlsx $o }
Invoke-Step 'Test-ExcelEstimate（見積書の検算）' '08_見積書の検算.xlsx' '検算結果のリンクで、見積書の該当セルが開く' { param($o) & (& $t 'Test-ExcelEstimate.ps1') -Path (Join-Path $fixtures 'estimate/見積書.xlsx') -OutXlsx $o }
Invoke-Step 'Export-DesignDiagram（画面遷移図・ER 図）' '09d_画面遷移図とER図.xlsx' '画面遷移図・ER図のシートで、四角を動かすと矢印が付いてくる。両向きの矢印のラベルに → と ← がある。09d_画面遷移図とER図.md を VS Code などで開くと Mermaid の図になる' { param($o) & (& $t 'Export-DesignDiagram.ps1') -Path $design -OutXlsx $o -OutMarkdown ($o -replace '\.xlsx$', '.md') }
Invoke-Step 'Export-TableDdl（DDL）' '09_DDL.xlsx' 'DB との差異・定義書の指摘・テーブルのシート（09_DDL.sql も作る）' { param($o) & (& $t 'Export-TableDdl.ps1') -Path $design -OutFile (Join-Path $OutDir '09_DDL.sql') -CompareCatalog (Join-Path $fixtures 'ddl/postgresql_columns.csv') -OutXlsx $o }
Invoke-Step 'New-WbsDraft（WBS の下書き）' '09c_WBSの下書き.xlsx' 'WBS のシートの予定開始・予定終了が日付のセルで、Excel で並べ替えられる。担当者別のシートに山田・佐藤の工数が並ぶ' { param($o) & (& $t 'New-WbsDraft.ps1') -Path $design -Start '2026-11-02' -Members '山田', '佐藤' -OutXlsx $o }
Invoke-Step 'Measure-Progress（進捗）' '10_進捗.xlsx' 'タスク・担当者別・週次のシート' { param($o) & (& $t 'Measure-Progress.ps1') -Path (Join-Path $data 'WBS.xlsx') -AsOf '2026-10-14' -OutXlsx $o }
Invoke-Step 'New-WeeklyReport（週次報告）' '10b_週次報告/週次報告_20261020.xlsx' '担当者別のシートのリンクで WBS・指摘管理表の該当セルが開く。同じフォルダの週次報告_20261020.txt をメールに貼って読める' { param($o) & (& $t 'New-WeeklyReport.ps1') -Wbs (Join-Path $data 'WBS.xlsx') -Findings (Join-Path $data 'レビュー') -BugList (Join-Path $data '障害管理表.xlsx') -AsOf '2026-10-20' -OutDir (Split-Path -Parent $o) }
Invoke-Step 'Measure-Effort（工数実績）' '10c_工数.xlsx' 'タスクのシートの予定・実績・判定と、WBS へのリンク。担当者×週・WBS に無い作業のシート' { param($o) & (& $t 'Measure-Effort.ps1') -Path (Join-Path $data '工数実績.xlsx') -Wbs (Join-Path $data 'WBS.xlsx') -AsOf '2026-10-14' -OutXlsx $o }
Invoke-Step 'Export-ActionItems（議事録の宿題）' '10d_議事録の宿題.xlsx' '宿題のシートの件名・担当・期限・メモ（担当なしなど）と、決定事項のシート' { param($o) & (& $t 'Export-ActionItems.ps1') -Path (Join-Path $data '20261020_定例議事録.txt') -OutXlsx $o }
Invoke-Step 'Measure-ReviewQuality（レビュー）' '11_レビュー.xlsx' 'グラフ（2 つ）が表示される（ゾーン分析・推移のシート）' { param($o) & (& $t 'Measure-ReviewQuality.ps1') -Path (Join-Path $data 'レビュー') -DesignPath $design -ReviewLog (Join-Path $data 'レビュー記録.csv') -AsOf '2026-10-20' -OutXlsx $o }
Invoke-Step 'Measure-TestQuality（テスト）' '12_テスト.xlsx' 'グラフ（3 つ）が表示される（推移・機能別などのシート）' { param($o) & (& $t 'Measure-TestQuality.ps1') -Path (Join-Path $data 'テスト') -BugList (Join-Path $data '障害管理表.xlsx') -DesignPath $design -MinBugs 5 -AsOf '2026-10-12' -OutXlsx $o }
Invoke-Step 'Test-TestSpec（テスト仕様書）' '13_テスト仕様書.xlsx' 'リンクでテスト仕様書の該当セルが開く（必須・型・コード値・メッセージの指摘）' { param($o) & (& $t 'Test-TestSpec.ps1') -Path (Join-Path $data 'テスト') -DesignPath $design -OutXlsx $o }
Invoke-Step 'Test-TestEvidence（エビデンス）' '13c_エビデンス.xlsx' 'ケース別のエビデンスの場所のリンクで画像が開く。No.1 と No.3 に同じ画像の使い回しの警告がある。エビデンスの無いケースのリンクで、テスト仕様書の該当セルが開く' { param($o) & (& $t 'Test-TestEvidence.ps1') -Path (Join-Path $data 'テスト') -EvidencePath (Join-Path $data 'エビデンス') -OutXlsx $o }
Invoke-Step 'New-TestSpecDraft（テスト仕様書の下書き）' '13b_テスト仕様書の下書き.xlsx' '一覧と画面ごと（SCR001〜SCR003）のシート。期待結果にメッセージの文言が入り、そのまま書き足せる（列幅・折り返し）' { param($o) & (& $t 'New-TestSpecDraft.ps1') $design -OutXlsx $o }
Invoke-Step 'Measure-DefectRemoval（工程別の摘出）' '14_工程別の摘出.xlsx' '工程別のシートのパレート図が表示される' { param($o) & (& $t 'Measure-DefectRemoval.ps1') -BugList (Join-Path $data '障害管理表.xlsx') -Findings (Join-Path $data 'レビュー') -OutXlsx $o }
Invoke-Step 'Measure-ReviewEffectiveness（レビューの効き目）' '15_レビューの効き目.xlsx' '設計書別のシートのグラフが表示される' { param($o) & (& $t 'Measure-ReviewEffectiveness.ps1') -Findings (Join-Path $data 'レビュー') -BugList (Join-Path $data '障害管理表.xlsx') -DesignPath $design -AsOf '2026-10-20' -OutXlsx $o }
Invoke-Step 'Compare-ReviewFindings（AI と人の指摘）' '16_AIと人の指摘.xlsx' '観点表への反映案のシートの B〜G 列が観点表と同じ並び' { param($o) & (& $t 'Compare-ReviewFindings.ps1') (Join-Path $data 'AIのレビュー結果.xlsx') (Join-Path $data 'レビュー') -OutXlsx $o }

# 品質評価: 2 回作って、前回との比較・推移と、定性分析の書き込みまで確かめる
$reportDir = Join-Path $OutDir '17_品質評価'
$reportArgs = @{ OutDir = $reportDir; DesignPath = $design; Findings = (Join-Path $data 'レビュー'); AiReview = (Join-Path $data 'AIのレビュー結果.xlsx'); ReviewLog = (Join-Path $data 'レビュー記録.csv'); TestPath = (Join-Path $data 'テスト'); BugList = (Join-Path $data '障害管理表.xlsx'); Requirement = (Join-Path $data '要件一覧.csv'); HistoryCsv = (Join-Path $reportDir '実績.csv'); Project = '動作確認' }
Invoke-Step 'New-QualityReport（品質評価、1 回目）' '17_品質評価/品質評価_20261006.xlsx' '（2 回目のブックを見る）' { param($o) & (& $t 'New-QualityReport.ps1') @reportArgs -AsOf '2026-10-06' }
Invoke-Step 'Invoke-QualityAnalysis（1 回目の定性分析を書き込む）' '17_品質評価/品質評価_20261006.xlsx' '（2 回目の依頼に「前回の定性分析と対策」が入る）' {
    param($o)
    $null = & (& $t 'Invoke-QualityAnalysis.ps1') $o -PromptOnly
    $answer = '| No | 定性分析（原因） | 対策 |' + "`r`n" + '| 1 | （動作確認用の回答）レビュー時間が短かった可能性がある | 設計リーダーが再レビューする（今週中） |'
    [System.IO.File]::WriteAllText(($o -replace '\.xlsx$', '_定性分析の回答.txt'), $answer, (New-Object System.Text.UTF8Encoding($true)))
    & (& $t 'Invoke-QualityAnalysis.ps1') $o -ImportOnly
}
Invoke-Step 'New-QualityReport（品質評価、2 回目）' '17_品質評価/品質評価_20261020.xlsx' '前回（2026/10/06）・増減の列、推移・要対応のシート、詳細の列のリンクで各ツールの結果が開く' { param($o) & (& $t 'New-QualityReport.ps1') @reportArgs -AsOf '2026-10-20' }
Invoke-Step 'Invoke-QualityAnalysis（依頼を作る）' '17_品質評価/品質評価_20261020_定性分析の依頼.txt' 'この依頼を Copilot に貼り付けて、「No | 定性分析（原因） | 対策」の表で答えが返るか（手順は下の「Copilot」）' { param($o) & (& $t 'Invoke-QualityAnalysis.ps1') (Join-Path $reportDir '品質評価_20261020.xlsx') -PromptOnly }
Invoke-Step 'Invoke-DesignReview（AI レビューの依頼）' '18_AIレビュー/prompts/R001.txt' 'この依頼を Copilot に貼り付けて、観点ごとの判定が返るか（手順は下の「Copilot」）' { param($o) & (& $t 'Invoke-DesignReview.ps1') -Path (Join-Path $design "$scr2.xlsx") -OutDir (Join-Path $OutDir '18_AIレビュー') -PromptOnly }

# 書き換えるツール（写しに対して）
Invoke-Step 'Reset-ExcelView（表示を整える）' '資料/書き換え用/表示を整える.xlsx' 'Excel で開いたとき、先頭シートの A1 が倍率 100% で表示される' { param($o) & (& $t 'Reset-ExcelView.ps1') -Path (Join-Path $copies '表示を整える.xlsx') -Force }
Invoke-Step 'Replace-Excel（置換）' '19_置換.xlsx' '置換した箇所の一覧と、書き換えた写し（資料/書き換え用/置換.xlsx）が Excel で開ける' { param($o) & (& $t 'Replace-Excel.ps1') -Pattern 'TODO' -Replacement '対応済み' -SimpleMatch -Path (Join-Path $copies '置換.xlsx') -Force -OutXlsx $o }
Invoke-Step 'Export-ExcelToPdf（PDF）' '20_PDF' 'PDF の向き・はみ出し' { param($o) & (& $t 'Export-ExcelToPdf.ps1') -Path (Join-Path $design "$scr1.xlsx") -OutputRoot $o } -NeedsExcel

# ---------------------------------------------------------------- 確認リスト


Write-Host ''
$failed = @($steps | Where-Object Status -eq '失敗')
$listPath = Join-Path $OutDir '00_確認リスト.xlsx'
$rows = New-Object System.Collections.Generic.List[object]
foreach ($s in $steps) {
    $row = @{ Values = @([double]$s.No, $s.Name, $s.Status, $s.File, $s.Check, '', $s.Message) }
    if ($s.Path -and (Test-Path -LiteralPath $s.Path -PathType Leaf)) { $row.Link = @{ Column = 4; Path = $s.Path; Location = '' } }
    $rows.Add($row)
}
$manual = @(
    @('Copilot（品質評価の定性分析）', ('PowerShell で .\tools\Invoke-QualityAnalysis.ps1 "{0}" -Force を実行 → 依頼が Copilot に貼り付けられる → 回答をコピーして Enter → ブックの定性分析・対策の列に入る' -f (Join-Path $reportDir '品質評価_20261020.xlsx'))),
    @('Copilot（AI レビュー）', ('PowerShell で .\tools\Invoke-DesignReview.ps1 "{0}" -OutDir "{1}" を実行 → 依頼を Copilot に貼り付け、回答をコピーすると Enter を押さなくても次の依頼に進む → レビュー結果.xlsx ができる' -f (Join-Path $design "$scr2.xlsx"), (Join-Path $OutDir '18_AIレビュー'))),
    @('「送る」', '.\tools\Install-SendTo.ps1 を実行し、エクスプローラーで 17_品質評価 のフォルダを右クリック →「送る」→「Excel 品質評価の定性分析」'),
    @('画面（Start-ExcelTools）', ('tools\Start-ExcelTools.cmd をダブルクリック → 左の「提出前チェック」を選ぶ → 「{0}」のフォルダを画面にドラッグ＆ドロップ → 「実行」→ 記録の欄に進み具合が出て、結果の Excel が開く。「表示を整える」は「試しに実行」で書き換える箇所が出る。閉じて開き直すと、前の入力が出る' -f $versions)),
    @('画面の見た目と操作（Start-ExcelTools）', ('① 表示の倍率（設定 → ディスプレイ → 拡大縮小）が 125%・150% の PC で、文字がにじまず、入力欄やボタンが切れていない（にじむ・切れるときは、倍率と画面の写しを書き留める） ② 左上の「やりたいことを探す」に「差分」と入れて Enter → 差分の処理が選ばれ、入力欄に移る（実行はしない） ③「PDF にする」の「Excel のフォルダ」に「{0}」、「PDF の保存先のフォルダ」に「{1}」を入れて実行し、途中で画面を閉じる →「中止して閉じますか」で はい → 少し待って閉じ、タスク マネージャーに EXCEL.EXE が残らない ④ 名前に半角の [ ] を含む Excel（例: 設計書[旧].xlsx）を「提出前チェック」に入れて実行 →「見つかりません」にならない' -f $design, (Join-Path $OutDir '20_PDF_中止'))),
    @('差分の画面で画像を比べる（Compare-Excel）', '画面レイアウトのスクリーンショットを貼った設計書を写し、写しのほうのスクリーンショットを撮り直して貼り替え（重ねた丸数字も 1 つ動かす）→ 2 つを「送る」→「差分」で比べる → 差分の一覧に「画像 変更」と「図形 移動」が出る → 画像の見本を押すと、旧と新が並び、違うところが枠で囲まれる（撮り直しただけで絵が同じなら「見た目は同じ」）。Edge と Chrome で確かめる'),
    @('配布（Install-ExcelTools）', 'tools フォルダを共有フォルダ（または別のフォルダ）に写し、その中の Install-ExcelTools.cmd をダブルクリック →「送る」に項目ができる → 共有フォルダの README.md を書き換え、PowerShell で & "$env:LOCALAPPDATA\ExcelTools\tools\Install-ExcelTools.ps1" -Update →「更新」と出る → -Uninstall で「送る」の項目が消える'),
    @('SharePoint の版と比べる', 'OneDrive で同期している SharePoint のフォルダで、版が 2 つ以上ある Excel を右クリック →「送る」→「Excel SharePoint の版と比べる」→ バージョン履歴の番号（1.0 など）を入力 → その版と現在の版の差分の画面が開く（Excel で取り出せたか、ブラウザーでダウンロードしたかも書き留める）'))
foreach ($m in $manual) { $rows.Add(@{ Values = @('', $m[0], '手で確かめる', '', $m[1], '', '') }) }
Export-Workbook -FilePath $listPath -Sheets @(@{
        Name = '確認リスト'
        Columns = @(@{ Header = 'No'; Width = 5 }, @{ Header = 'ツール'; Width = 34 }, @{ Header = '実行'; Width = 9 }, @{ Header = '結果のファイル'; Width = 40 }, @{ Header = 'Excel で開いて見るところ'; Width = 70 }, @{ Header = '確認（OK / NG）'; Width = 12 }, @{ Header = 'エラー'; Width = 50 })
        Rows = $rows
    })

Write-Host (($steps | Format-Table No, Name, Status, Message -AutoSize -Wrap | Out-String).TrimEnd())
Write-Host ('実行 {0} 件: 成功 {1}・失敗 {2}・スキップ {3}' -f $steps.Count, @($steps | Where-Object Status -eq '成功').Count, $failed.Count, @($steps | Where-Object Status -eq 'スキップ').Count)
Write-Host "確認リスト: $listPath"
Write-Host '確認リストの「結果のファイル」のリンクから順に開き、「Excel で開いて見るところ」を確かめて、確認の列に OK / NG を書いてください。'
if (-not $NoOpen) { Invoke-Item -LiteralPath $listPath }
if ($failed.Count -gt 0) { throw ('失敗したツールがあります: {0}' -f ((@($failed | ForEach-Object { $_.Name })) -join '、')) }

# .SYNOPSIS
#     サンプルの設計書と、その場で作る指摘管理表・テスト仕様書・障害管理表・WBS で全ツールを実行し、結果の Excel と確認リストを作ります。
#
# .DESCRIPTION
#     自動テスト（Invoke-Tests.ps1）では確かめられない、Excel で開いたときの見た目（グラフ・リンク・色・表示位置）と、
#     Copilot とのやり取りを、Windows の手元でまとめて確かめるためのものです。
#
#     結果は「ドキュメント\Excel ツールの動作確認_日時」に保存し、最後に「00_確認リスト.xlsx」を開きます。
#     確認リストのリンクから結果のファイルを開き、「Excel で開いて見るところ」を確かめてください。
#     Copilot・「送る」は、確認リストの最後の行の手順で確かめます。サンプルの設計書は書き換えません。
#
# .PARAMETER OutDir
#     結果を保存するフォルダ。省略時は「ドキュメント\Excel ツールの動作確認_日時」。
#
# .PARAMETER NoOpen
#     終わったときに確認リストを開きません。
#
# .EXAMPLE
#     .\tests\Invoke-SmokeTest.ps1
#
# .NOTES
#     PDF にする（Export-ExcelToPdf.ps1）は Excel が必要なので、Excel が無ければスキップします。
