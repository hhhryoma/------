[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('List', 'Check', 'ResetView', 'Compare', 'SpVersion', 'Search', 'Terms', 'Consistency', 'TestDraft', 'Size', 'Estimate', 'Quality', 'Analysis', 'Weekly', 'Actions', 'Pdf')]
    [string] $Action,

    # 「送る」は選んだファイルを引数の最後に並べるため、名前の無い引数はすべてここに入れる（PositionalBinding = $false）
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]] $Path,

    [string] $OutDir,

    [string] $Text,

    [switch] $NoOpen,

    [switch] $NoPause
)

$ErrorActionPreference = 'Stop'

#     「送る」メニューの項目。Label はショートカットの名前（Install-SendTo.ps1 が使う）、File は結果のファイル名。
#     Literal はツールが -Path をワイルドカードとして扱わない（そのまま渡す）もの
$actions = [ordered]@{
    Check       = @{ Label = '提出前チェック'; File = '提出前チェック'; Tool = 'Test-ExcelDeliverable.ps1'; Detail = '体裁・消し忘れ・情報漏れ・印刷のはみ出しなどを確かめ、結果の Excel を開きます' }
    ResetView   = @{ Label = '表示を整える'; Tool = 'Reset-ExcelView.ps1'; Detail = '開いたときに先頭シートの A1 が倍率 100% で表示される状態に直します（上書きします）' }
    Compare     = @{ Label = '差分（2 つ選ぶ）'; File = '差分'; Tool = 'Compare-Excel.ps1'; Literal = $true; Detail = '2 つのブック・Word 文書（またはフォルダ）を比べ、差分をシートの方眼のまま色付けした画面（ブラウザー）を開きます（一覧の Excel も保存）。版数（ファイル名）か更新日時の古い方を旧にします。3 つ以上か、版数の付いたファイルを置いたフォルダを 1 つ選ぶと、版を続けて比べます' }
    SpVersion   = @{ Label = 'SharePoint の版と比べる'; Tool = 'Compare-SharePointVersion.ps1'; Literal = $true; Detail = 'OneDrive で同期している SharePoint のファイルを 1 つ選び、バージョン履歴の番号（3.0 など）を入力すると、その版と現在の版の差分の画面（ブラウザー）を開きます（版は Excel で取り出します）' }
    Search      = @{ Label = '検索'; File = '検索'; Tool = 'Search-Excel.ps1'; Detail = '入力した文字を、値・数式・コメント・図形などから探し、結果の Excel を開きます' }
    Terms       = @{ Label = '表記ゆれ'; File = '表記ゆれ'; Tool = 'Test-ExcelTerms.ps1'; Detail = '表記ゆれ（ユーザ／ユーザー など）を見つけ、結果の Excel を開きます' }
    Consistency = @{ Label = '設計書の整合性チェック'; File = '整合性チェック'; Tool = 'Test-DesignConsistency.ps1'; Detail = '設計書どうしの参照の食い違いを見つけ、結果の Excel を開きます' }
    TestDraft   = @{ Label = 'テスト仕様書の下書き'; File = 'テスト仕様書の下書き'; Tool = 'New-TestSpecDraft.ps1'; Literal = $true; Detail = '画面設計書（とコード定義書・メッセージ一覧）のフォルダかファイルを選ぶと、必須・桁数・型・コード値・イベントのテストケースの下書きを作り、Excel を開きます' }
    Size        = @{ Label = '規模の見積もり'; File = '規模'; Tool = 'Measure-DesignSize.ps1'; Detail = '設計書の規模を数えて工数の目安を出し、結果の Excel を開きます' }
    Estimate    = @{ Label = '見積書の検算'; File = '見積書の検算'; Tool = 'Test-ExcelEstimate.ps1'; Detail = '見積書の合計の範囲漏れ・手入力の合計などを見つけ、結果の Excel を開きます' }
    Quality     = @{ Label = '品質評価'; Tool = 'New-QualityReport.ps1'; Literal = $true; Detail = '案件のフォルダ（品質評価の設定.csv を置いたもの）を選ぶと、品質の各ツールを実行して品質評価のブックを作り、開きます' }
    Analysis    = @{ Label = '品質評価の定性分析'; Tool = 'Invoke-QualityAnalysis.ps1'; Literal = $true; Detail = '品質評価のブック（またはそのフォルダ）を選ぶと、× と △ の指標の定性分析・対策の下書きを Copilot とクリップボードでやり取りして書き込み、開きます' }
    Weekly      = @{ Label = '週次報告'; Tool = 'New-WeeklyReport.ps1'; Literal = $true; Detail = '案件のフォルダ（週次報告の設定.csv を置いたもの）を選ぶと、WBS・指摘管理表・障害管理表から週次報告（進捗・指摘・障害と担当者ごとのやること）を作って開き、報告の文章をクリップボードに写します' }
    Actions     = @{ Label = '議事録の宿題'; File = '議事録の宿題'; Tool = 'Export-ActionItems.ps1'; Detail = '議事録（Word・Excel・テキスト）かそのフォルダを選ぶと、宿題（担当・期限付き）と決定事項を抜き出し、課題管理表の形の Excel を開きます' }
    Pdf         = @{ Label = 'PDF にする'; Tool = 'Export-ExcelToPdf.ps1'; Literal = $true; Detail = 'Excel を PDF にして、元のファイルと同じフォルダに保存します（Excel が必要）' }
}

if ($Action -eq 'List') {
    foreach ($key in $actions.Keys) {
        [pscustomobject]@{ Action = $key; Label = $actions[$key].Label; Detail = $actions[$key].Detail }
    }
    return
}

# ---------------------------------------------------------------- ヘルパー


function ConvertFrom-SendToArgument {
    #     「送る」から渡されたパスを直す。ドライブの直下（"D:\"）は、コマンドラインの解釈で D:" になって届く
    param([string] $Value)
    $v = $Value.Trim().Trim('"')
    if ($v -match '^[A-Za-z]:$') { $v += '\' }
    return $v
}

function Get-OldNew {
    #     2 つのパスを旧・新の順に並べる。両方に版数があり違えば版数で（ExcelTools.psm1 の Get-NameVersion）、そうでなければ更新日時で決める
    param([string[]] $Pair)
    Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
    $a = Get-Item -LiteralPath $Pair[0]
    $b = Get-Item -LiteralPath $Pair[1]
    if ($a.PSIsContainer -ne $b.PSIsContainer) { throw 'ファイルどうし、またはフォルダどうしを選んでください' }
    $va = Get-NameVersion $a.Name
    $vb = Get-NameVersion $b.Name
    if ($null -ne $va -and $null -ne $vb -and $va -ne $vb) { $swap = $va -gt $vb }
    else { $swap = $a.LastWriteTime -gt $b.LastWriteTime }
    if ($swap) { return , @($b.FullName, $a.FullName) }
    return , @($a.FullName, $b.FullName)
}

function Get-ResultPath {
    #     結果のファイル名（名前_日時.xlsx）。同じ名前があれば連番を付ける
    param([string] $Dir, [string] $Name)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $candidate = Join-Path $Dir ('{0}_{1}.xlsx' -f $Name, $stamp)
    for ($n = 2; Test-Path -LiteralPath $candidate; $n++) { $candidate = Join-Path $Dir ('{0}_{1}({2}).xlsx' -f $Name, $stamp, $n) }
    return $candidate
}

# ---------------------------------------------------------------- 実行


$def = $actions[$Action]
try { $Host.UI.RawUI.WindowTitle = 'Excel ツール - ' + $def.Label } catch { }

$failure = $null
try {
    $targets = @($Path | Where-Object { $_ -and $_.Trim() } | ForEach-Object { ConvertFrom-SendToArgument $_ })
    if ($targets.Count -eq 0) { throw 'ファイルまたはフォルダを選んで、右クリックの「送る」から実行してください' }
    foreach ($t in $targets) { if (-not (Test-Path -LiteralPath $t)) { throw "見つかりません: $t" } }
    $targets = @($targets | ForEach-Object { (Resolve-Path -LiteralPath $_).ProviderPath })

    Write-Host ('{0}（{1} 件）' -f $def.Label, $targets.Count)
    foreach ($t in $targets) { Write-Host "  $t" }
    Write-Host ''

    $tool = Join-Path $PSScriptRoot $def.Tool
    # ほとんどのツールは -Path にワイルドカードを使えるため、[ ] を含む名前がそのまま通るようにする
    $toolPaths = $targets
    if (-not $def.Literal) { $toolPaths = @($targets | ForEach-Object { [System.Management.Automation.WildcardPattern]::Escape($_) }) }

    $out = $null
    if ($def.File) {
        if (-not $OutDir) {
            $docs = [Environment]::GetFolderPath('MyDocuments')
            if (-not $docs) { $docs = [System.IO.Path]::GetTempPath() }
            $OutDir = Join-Path $docs 'Excel ツールの結果'
        }
        $OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
        New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
        $out = Get-ResultPath $OutDir $def.File
    }

    switch ($Action) {
        'Compare' {
            # 差分はシートの方眼のまま色付けした HTML で開く（一覧の Excel も同じ名前で保存する）
            $html = [System.IO.Path]::ChangeExtension($out, '.html')
            if ($targets.Count -ge 3 -or ($targets.Count -eq 1 -and (Test-Path -LiteralPath $targets[0] -PathType Container))) {
                # 3 つ以上か、版数の付いたファイルを置いたフォルダ 1 つなら、版を続けて比べる（v1.0 → v1.1 → v1.2）
                Write-Host '版を続けて比べます（版数の順。版数が無ければ選んだ順）'
                Write-Host ''
                $null = & $tool -Versions $targets -OutXlsx $out -OutHtml $html
            }
            else {
                if ($targets.Count -ne 2) { throw "比べるファイル（またはフォルダ）を 2 つ選んでください（$($targets.Count) 件選ばれています。版を続けて比べるときは 3 つ以上か、版数の付いたファイルを置いたフォルダを 1 つ）" }
                $pair = Get-OldNew $targets
                Write-Host "旧: $($pair[0])"
                Write-Host "新: $($pair[1])"
                Write-Host ''
                $null = & $tool -Old $pair[0] -New $pair[1] -OutXlsx $out -OutHtml $html
            }
            Write-Host "差分の一覧（Excel）: $out"
            $out = $html
        }
        'SpVersion' {
            if ($targets.Count -ne 1 -or (Test-Path -LiteralPath $targets[0] -PathType Container)) { throw "OneDrive で同期している SharePoint のファイルを 1 つ選んでください（$($targets.Count) 件選ばれています）" }
            # 版の番号は、-Text が無ければツールが入力を求める。差分の画面はツールが開く
            $spArgs = @{ Path = $targets[0]; NoOpen = [bool]$NoOpen }
            if ($Text) { $spArgs['Version'] = $Text }
            if ($OutDir) { $spArgs['OutDir'] = $OutDir }
            $null = & $tool @spArgs
        }
        'TestDraft' {
            $null = & $tool -DesignPath $targets -OutXlsx $out
        }
        'Search' {
            if (-not $Text) { $Text = Read-Host '探す文字を入力してください' }
            if (-not $Text) { throw '探す文字が入力されていません' }
            $null = & $tool -Pattern $Text -SimpleMatch -Path $toolPaths -OutXlsx $out
        }
        'ResetView' {
            $results = @(& $tool -Path $toolPaths)
            $changed = @($results | Where-Object { $_.Result -ne '変更なし' })
            if ($changed.Count -gt 0) { $changed | Format-Table File, Result, Detail -AutoSize -Wrap | Out-Host }
        }
        'Pdf' {
            foreach ($t in $toolPaths) { & $tool -Path $t | Out-Host }
        }
        'Quality' {
            if ($targets.Count -ne 1) { throw "案件のフォルダ（品質評価の設定.csv を置いたフォルダ）か、設定ファイルを 1 つ選んでください（$($targets.Count) 件選ばれています）" }
            Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
            $settings = Read-QualityReportConfig $targets[0]
            $reportArgs = @{ Config = $targets[0] }
            $dir = $settings.OutDir
            if ($OutDir) { $reportArgs['OutDir'] = $OutDir; $dir = $OutDir }
            $title = '品質評価'
            if ($settings.Contains('Title')) { $title = $settings.Title }
            $null = & $tool @reportArgs
            $out = Join-Path $dir ('{0}_{1}.xlsx' -f $title, (Get-Date).ToString('yyyyMMdd'))
        }
        'Weekly' {
            if ($targets.Count -ne 1) { throw "案件のフォルダ（週次報告の設定.csv を置いたフォルダ）か、設定ファイルを 1 つ選んでください（$($targets.Count) 件選ばれています）" }
            Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
            $settings = Read-WeeklyReportConfig $targets[0]
            $reportArgs = @{ Config = $targets[0] }
            $dir = $settings.OutDir
            if ($OutDir) { $reportArgs['OutDir'] = $OutDir; $dir = $OutDir }
            $title = '週次報告'
            if ($settings.Contains('Title')) { $title = $settings.Title }
            $null = & $tool @reportArgs
            $name = '{0}_{1}' -f $title, (Get-Date).ToString('yyyyMMdd')
            $out = Join-Path $dir ($name + '.xlsx')
            # 報告の文章は、メールやチャットにそのまま貼れるようにクリップボードに写す
            $textFile = Join-Path $dir ($name + '.txt')
            if (-not $NoOpen -and (Test-Path -LiteralPath $textFile)) {
                try {
                    Set-Clipboard -Value ([System.IO.File]::ReadAllText($textFile))
                    Write-Host ''
                    Write-Host '報告の文章をクリップボードに写しました（メールやチャットに貼り付けられます）'
                }
                catch { Write-Host ('報告の文章をクリップボードに写せませんでした: ' + $_.Exception.Message) }
            }
        }
        'Analysis' {
            if ($targets.Count -ne 1) { throw "品質評価のブックか、それを置いたフォルダを 1 つ選んでください（$($targets.Count) 件選ばれています）" }
            & $tool -Path $targets[0] | Format-Table No, Judgement, Name, Cause -AutoSize -Wrap | Out-Host
            # 書き込んだブック（フォルダなら、その中で一番新しい xlsx。書き込むと更新日時が新しくなる）
            $out = $targets[0]
            if (Test-Path -LiteralPath $out -PathType Container) {
                $out = @(Get-ChildItem -LiteralPath $targets[0] -Filter '*.xlsx' -File | Where-Object { $_.Name -notlike '~$*' } | Sort-Object LastWriteTime -Descending | ForEach-Object { $_.FullName })[0]
            }
        }
        default {
            $null = & $tool -Path $toolPaths -OutXlsx $out
        }
    }

    if ($out) {
        if (-not (Test-Path -LiteralPath $out)) { throw "結果のファイルを保存できませんでした: $out" }
        Write-Host ''
        Write-Host "結果: $out"
        if (-not $NoOpen) { Invoke-Item -LiteralPath $out }
    }
}
catch {
    $failure = $_
    Write-Host ''
    Write-Host ('エラー: ' + $_.Exception.Message) -ForegroundColor Red
}
finally {
    # Install-ExcelTools.ps1 で入れたツールなら、配布元が更新されていないかを確かめる（12 時間に 1 回まで。次の実行から新しい版になる）
    $installDir = Split-Path -Parent $PSScriptRoot
    if (Test-Path -LiteralPath (Join-Path $installDir 'install.json')) {
        try {
            $u = & (Join-Path $PSScriptRoot 'Install-ExcelTools.ps1') -InstallDir $installDir -CheckOnly 6>$null
            if ($u -and $u.Result -eq '更新') { Write-Host ''; Write-Host ('Excel ツールを新しい版に更新しました（次の実行から使います）: {0} ファイル' -f $u.Copied) }
        }
        catch { Write-Host ''; Write-Host ('Excel ツールの更新を確かめられませんでした: ' + $_.Exception.Message) }
    }
    if (-not $NoPause) {
        Write-Host ''
        $null = Read-Host 'Enter キーを押すと閉じます'
    }
}
if ($failure) { throw $failure }

# .SYNOPSIS
#     エクスプローラーの右クリック →「送る」から、選んだファイル・フォルダに各ツールを実行します（Excel 不要）。
#
# .DESCRIPTION
#     Install-SendTo.ps1 で「送る」メニューに登録したショートカットから呼ばれます。PowerShell に慣れていなくても、
#     ファイルやフォルダを選んで「送る」を選ぶだけでツールを使えます。
#
#       ・結果のある処理（提出前チェック・差分・検索など）は、結果を Excel のファイルに保存して開きます
#         （既定の保存先は「ドキュメント\Excel ツールの結果」）
#       ・終わったら Enter キーを押すまでウィンドウを閉じません（エラーのときも、内容を読めるように）
#       ・差分は 2 つ選びます。ファイル名の版数（_v1.1 など）、無ければ更新日時の古い方を旧にします
#       ・検索は、探す文字を入力します（正規表現ではなく、入力した文字そのものを探します）
#
#     使える処理は -Action List で確かめられます。
#
# .PARAMETER Action
#     実行する処理。Check（提出前チェック）・ResetView（表示を整える）・Compare（差分）・SpVersion（SharePoint の版と比べる）・Search（検索）・
#     Terms（表記ゆれ）・Consistency（設計書の整合性チェック）・TestDraft（テスト仕様書の下書き）・Size（規模の見積もり）・Estimate（見積書の検算）・
#     Quality（品質評価。案件のフォルダの「品質評価の設定.csv」を読む）・Analysis（品質評価の定性分析の下書き）・
#     Weekly（週次報告。案件のフォルダの「週次報告の設定.csv」を読む）・Actions（議事録の宿題）・Pdf（PDF にする）。
#     List は処理の一覧を返します。
#
# .PARAMETER Path
#     対象のファイル・フォルダ（「送る」から渡されます）。
#
# .PARAMETER OutDir
#     結果のファイルを保存するフォルダ。既定は「ドキュメント\Excel ツールの結果」（品質評価・週次報告は設定ファイルの出力先）。
#
# .PARAMETER Text
#     検索（Search）で探す文字、SharePoint の版と比べる（SpVersion）ときの旧の版の番号。指定しなければ入力を求めます。
#
# .PARAMETER NoOpen
#     結果のファイルを開きません（週次報告の文章もクリップボードに写しません）。
#
# .PARAMETER NoPause
#     終わったときに Enter キーを待ちません。
#
# .EXAMPLE
#     # 「送る」と同じことを PowerShell から行う
#     .\Invoke-SendTo.ps1 -Action Check D:\提出物
#
# .EXAMPLE
#     # 使える処理の一覧
#     .\Invoke-SendTo.ps1 -Action List
#
# .NOTES
#     要件: PowerShell 5.1 以降。PDF にする（Pdf）だけは Excel が必要です。同じフォルダに各ツールと ExcelTools.psm1 が必要です。
