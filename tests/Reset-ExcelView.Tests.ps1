#     Reset-ExcelView.ps1 のテスト。フィクスチャは tests\fixtures\deliverable と tests\fixtures\view

$ResetView = Join-Path $script:ToolsDir 'Reset-ExcelView.ps1'
$CheckTool = Join-Path $script:ToolsDir 'Test-ExcelDeliverable.ps1'

function Invoke-Reset {
    param([hashtable] $Options)
    return , @(& $ResetView @Options 6>$null 3>$null)
}

function Get-ViewIssues {
    param([string] $Path, [hashtable] $Extra = @{})
    return , @(& $CheckTool -Path $Path -Check View @Extra 6>$null)
}

function Read-ZipTexts {
    # xlsx の XML パーツを名前 → 文字列で読む（比較用）
    param([string] $Path)
    $result = @{}
    $fs = [System.IO.File]::OpenRead($Path)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        foreach ($e in $zip.Entries) {
            $sr = New-Object System.IO.StreamReader($e.Open())
            try { $result[$e.FullName] = $sr.ReadToEnd() } finally { $sr.Dispose() }
        }
    }
    finally { $fs.Dispose() }
    return $result
}

function Remove-ViewParts {
    # 表示に関する部分（<sheetViews> と <workbookView>）を除いた文字列
    param([string] $Xml)
    $x = [regex]::Replace($Xml, '<sheetViews>.*?</sheetViews>', '', 'Singleline')
    return [regex]::Replace($x, '<workbookView\b[^>]*/>', '')
}

Test-Case 'Reset-ExcelView: -WhatIf では書き換えず、直す予定だけを表示する' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $hash = (Get-FileHash -LiteralPath $bad).Hash
    $r = Invoke-Reset @{ Path = $bad; WhatIf = $true }
    Assert-Equal '修正予定' $r[0].Result
    Assert-Match $r[0].Detail '作業グループ'
    Assert-Equal $hash (Get-FileHash -LiteralPath $bad).Hash 'ファイルが変わっています'
}

Test-Case 'Reset-ExcelView: 体裁の指摘がすべて解消される' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    Assert-True ((Get-ViewIssues $bad).Count -gt 0) '前提: 直す前は指摘あり'
    $r = Invoke-Reset @{ Path = $bad }
    Assert-Equal '修正' $r[0].Result
    $after = Get-ViewIssues $bad
    Assert-Equal @('表示モード') @($after | ForEach-Object { $_.Item }) '-NormalView なしなら表示モードだけ残る'

    $null = Invoke-Reset @{ Path = $bad; NormalView = $true }
    Assert-Equal 0 (Get-ViewIssues $bad).Count '-NormalView で表示モードも直る'
}

Test-Case 'Reset-ExcelView: 表示に関する部分以外は 1 文字も変えない' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $before = Read-ZipTexts $bad
    $null = Invoke-Reset @{ Path = $bad; NormalView = $true }
    $after = Read-ZipTexts $bad
    Assert-Equal (@($before.Keys | Sort-Object) -join ',') (@($after.Keys | Sort-Object) -join ',') 'パーツの一覧'
    $changed = @()
    foreach ($name in $before.Keys) {
        if ($before[$name] -ceq $after[$name]) { continue }
        $changed += $name
        Assert-Equal (Remove-ViewParts $before[$name]) (Remove-ViewParts $after[$name]) "$name の表示以外の部分"
        # 属性を消すときに空の属性（name=""）を書いていない
        Assert-NotMatch $after[$name] '\s\w+=""' "$name に空の属性"
    }
    Assert-Equal @('xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml') @($changed | Sort-Object) '書き換えたパーツ'
}

Test-Case 'Reset-ExcelView: 直す必要のないファイルは更新日時も含めて触らない' {
    $dir = Copy-Fixture 'deliverable'
    $good = Get-Item -LiteralPath (Join-Path $dir 'good.xlsx')
    $hash = (Get-FileHash -LiteralPath $good.FullName).Hash
    $time = $good.LastWriteTimeUtc
    $r = Invoke-Reset @{ Path = $good.FullName }
    Assert-Equal '変更なし' $r[0].Result
    Assert-Equal $hash (Get-FileHash -LiteralPath $good.FullName).Hash
    Assert-Equal $time (Get-Item -LiteralPath $good.FullName).LastWriteTimeUtc
}

Test-Case 'Reset-ExcelView: -KeepTimestamp で更新日時を保ち、-BackupDir に元のファイルを残す' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $old = [datetime]::new(2026, 1, 2, 3, 4, 5, [System.DateTimeKind]::Utc)
    [System.IO.File]::SetLastWriteTimeUtc($bad, $old)
    $hash = (Get-FileHash -LiteralPath $bad).Hash
    $backup = Join-Path (New-TestDir) 'backup'
    $null = Invoke-Reset @{ Path = $dir; KeepTimestamp = $true; BackupDir = $backup }
    Assert-Equal $old (Get-Item -LiteralPath $bad).LastWriteTimeUtc '更新日時'
    Assert-True ((Get-FileHash -LiteralPath $bad).Hash -ne $hash) '中身は直っている'
    Assert-Equal $hash (Get-FileHash -LiteralPath (Join-Path $backup 'bad.xlsx')).Hash 'バックアップは元のまま'
    Assert-Equal @('bad.xlsx', '書き換える前のファイル.txt') @(Get-ChildItem -LiteralPath $backup -File | ForEach-Object { $_.Name } | Sort-Object) '直したファイルだけをバックアップ（と印のファイル）'
}

Test-Case 'Reset-ExcelView: 見出しを固定したシートはスクロール領域の左上に戻す' {
    $dir = Copy-Fixture 'view'
    $file = Join-Path $dir 'frozen.xlsx'
    Assert-Equal 2 (Get-ViewIssues $file).Count '前提: 選択セルと表示位置の指摘'
    $null = Invoke-Reset @{ Path = $file }
    Assert-Equal 0 (Get-ViewIssues $file).Count
    $xml = (Read-ZipTexts $file)['xl/worksheets/sheet1.xml']
    Assert-Match $xml '<pane [^>]*topLeftCell="B2"' '固定位置の直後'
    Assert-Match $xml '<selection pane="bottomRight" activeCell="B2" sqref="B2"/>' 'カーソルのある枠の選択'
    Assert-Match $xml '<selection pane="topRight" activeCell="D1" sqref="D1"/>' 'ほかの枠の選択はそのまま'
}

Test-Case 'Reset-ExcelView: -Zoom で倍率をそろえられる' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $null = Invoke-Reset @{ Path = $bad; Zoom = 85; NormalView = $true }
    Assert-Equal 0 (Get-ViewIssues $bad @{ Zoom = 85 }).Count
    Assert-Match ((Read-ZipTexts $bad)['xl/worksheets/sheet2.xml']) 'zoomScale="85"'
}

Test-Case 'Reset-ExcelView: パスワード付きはスキップし、xls は対象外' {
    $dir = Copy-Fixture 'deliverable'
    $r = Invoke-Reset @{ Path = $dir }
    $byFile = @{}
    foreach ($x in $r) { $byFile[$x.File] = $x.Result }
    Assert-Equal 'スキップ' $byFile['encrypted.xlsx']
    Assert-True (-not $byFile.ContainsKey('legacy.xls')) 'xls は対象外'
    Assert-Equal '変更なし' $byFile['macro.xlsm']
}

Test-Case 'Reset-ExcelView: 開いている（ロックされた）ファイルは失敗として知らせる' -WindowsOnly {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $lock = [System.IO.File]::Open($bad, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
    try {
        $r = Invoke-Reset @{ Path = $bad }
        Assert-Equal '失敗' $r[0].Result
        Assert-Match $r[0].Detail '開いている'
    }
    finally { $lock.Dispose() }
}

Test-Case 'Reset-ExcelView: 読み取り専用のファイルは書き換えない' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Get-Item -LiteralPath (Join-Path $dir 'bad.xlsx')
    $bad.IsReadOnly = $true
    try {
        $hash = (Get-FileHash -LiteralPath $bad.FullName).Hash
        $r = Invoke-Reset @{ Path = $bad.FullName }
        Assert-Equal 'スキップ' $r[0].Result
        Assert-Match $r[0].Detail '読み取り専用'
        Assert-Equal $hash (Get-FileHash -LiteralPath $bad.FullName).Hash
    }
    finally { (Get-Item -LiteralPath $bad.FullName).IsReadOnly = $false }
}

Test-Case 'Reset-ExcelView: インターネットから取得したファイルの印（Zone.Identifier）を消さない' -WindowsOnly {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    Set-Content -LiteralPath $bad -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"
    $r = Invoke-Reset @{ Path = $bad }
    Assert-Equal '修正' $r[0].Result
    $zone = Get-Content -LiteralPath $bad -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue
    Assert-Match ($zone -join "`n") 'ZoneId=3' 'Zone.Identifier が残っていない'
}

Test-Case 'Reset-ExcelView: フォルダを複数指定したときは、-BackupDir に指定したフォルダごとに分けて残す' {
    $dir = New-TestDir
    $places = @('提出物/第1回', '提出物/第2回', '再提出/第1回')
    $hashes = New-Object System.Collections.Generic.List[string]
    foreach ($p in $places) {
        $file = Join-Path $dir "$p/見積書.xlsx"
        New-Item -ItemType Directory -Force -Path (Split-Path $file -Parent) | Out-Null
        Copy-Item -LiteralPath (Join-Path $script:Fixtures 'deliverable/bad.xlsx') -Destination $file
        Set-CellValues $file '詳細' @{ A1 = $p }
        $hashes.Add((Get-FileHash -LiteralPath $file).Hash)
    }
    $backup = Join-Path $dir 'bk'
    # 3 つ目はファイルで指定する（そのフォルダの名前「第1回」は 1 つ目と同じ）
    $r = Invoke-Reset @{ Path = @((Join-Path $dir '提出物/第1回'), (Join-Path $dir '提出物/第2回'), (Join-Path $dir '再提出/第1回/見積書.xlsx')); BackupDir = $backup }
    Assert-Equal @('修正', '修正', '修正') @($r | ForEach-Object { $_.Result })
    Assert-Equal 3 @(Get-ChildItem -LiteralPath $backup -File -Recurse -Filter '*.xlsx').Count '直したファイルをすべて残す'
    Assert-Equal $hashes.ToArray() @(@('第1回', '第2回', '第1回 (2)') | ForEach-Object { (Get-FileHash -LiteralPath (Join-Path $backup "$_/見積書.xlsx")).Hash }) '指定したフォルダの名前のフォルダに、元のファイルを残す'
}

Test-Case 'Reset-ExcelView: -BackupDir に前に残したファイルは上書きせず、番号を付けて残す' {
    $dir = Copy-Fixture 'deliverable'
    $bad = Join-Path $dir 'bad.xlsx'
    $original = (Get-FileHash -LiteralPath $bad).Hash
    $backup = Join-Path (New-TestDir) 'backup'
    $null = Invoke-Reset @{ Path = $bad; BackupDir = $backup }
    $fixed = (Get-FileHash -LiteralPath $bad).Hash
    # 2 回目（-NormalView で表示モードも直す）
    $text = Get-HostText { $null = & $ResetView -Path $bad -BackupDir $backup -NormalView }
    Assert-Equal $original (Get-FileHash -LiteralPath (Join-Path $backup 'bad.xlsx')).Hash '1 回目に残した元のファイル'
    Assert-Equal $fixed (Get-FileHash -LiteralPath (Join-Path $backup 'bad (2).xlsx')).Hash '2 回目に書き換える前のファイル'
    Assert-Match $text '番号を付けて残し.*bad \(2\)\.xlsx'
}

Test-Case 'Reset-ExcelView: スクロールした状態でウィンドウ枠を固定したシートは、固定した行・列を変えずにスクロール領域の左上に戻す' {
    # 固定した部分の左上は sheetView の topLeftCell。固定していない向き（列だけを固定したときの行など）は、スクロールの位置
    $views = [ordered]@{
        # 10〜11 行目（表の見出し）を固定し、スクロール領域を下へスクロールした
        '行'       = '<sheetView tabSelected="1" topLeftCell="A10" workbookViewId="0"><pane ySplit="2" topLeftCell="A30" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A35" sqref="A35"/></sheetView>'
        # 右へスクロールして J〜K 列を固定した（倍率だけを直す）
        '列'       = '<sheetView topLeftCell="J1" zoomScale="85" workbookViewId="0"><pane xSplit="2" topLeftCell="L1" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="L1" sqref="L1"/></sheetView>'
        # A〜B 列を固定して下へスクロールした
        '列・下へ' = '<sheetView topLeftCell="A50" workbookViewId="0"><pane xSplit="2" topLeftCell="C50" activePane="topRight" state="frozen"/><selection pane="topRight" activeCell="C60" sqref="C60"/></sheetView>'
        # 10〜11 行目を固定して右へスクロールした
        '行・右へ' = '<sheetView topLeftCell="D10" workbookViewId="0"><pane ySplit="2" topLeftCell="D12" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="D12" sqref="D12"/></sheetView>'
    }
    $book = Join-Path (New-TestDir) '画面一覧.xlsx'
    $sheets = [ordered]@{}
    foreach ($name in $views.Keys) { $sheets[$name] = @(@('No', '項目'), @('1', 'x')) }
    New-SheetsBook $book $sheets
    $i = 0
    foreach ($view in $views.Values) {
        $i++
        Edit-ZipText $book "xl/worksheets/sheet$i.xml" { param($t) [regex]::Replace($t, '<sheetViews>.*?</sheetViews>', '<sheetViews>' + $view + '</sheetViews>') }
    }
    $r = Invoke-Reset @{ Path = $book }
    Assert-Equal '修正' $r[0].Result
    Assert-Equal 0 (Get-ViewIssues $book).Count
    $xml = Read-ZipTexts $book
    $tags = { param($n) (@('sheetView', 'pane', 'selection') | ForEach-Object { [regex]::Match($xml["xl/worksheets/sheet$n.xml"], "<$_\b[^>]*>").Value }) -join ' ' }
    Assert-Match (& $tags 1) '^<sheetView [^>]*topLeftCell="A10"[^>]*> <pane [^>]*topLeftCell="A12"[^>]*> <selection pane="bottomLeft" activeCell="A12" sqref="A12"/>$' '固定した 10〜11 行目はそのまま'
    Assert-Match (& $tags 2) '^<sheetView [^>]*topLeftCell="J1"[^>]*> <pane [^>]*topLeftCell="L1"[^>]*> <selection pane="topRight" activeCell="L1" sqref="L1"/>$' '固定した J〜K 列はそのまま'
    Assert-NotMatch (& $tags 2) 'zoomScale'
    Assert-Match (& $tags 3) '^<sheetView [^>]*> <pane [^>]*topLeftCell="C1"[^>]*> <selection pane="topRight" activeCell="C1" sqref="C1"/>$' '列だけを固定したシートは 1 行目に戻す'
    Assert-NotMatch (& $tags 3) '<sheetView [^>]*topLeftCell'
    Assert-Match (& $tags 4) '^<sheetView [^>]*topLeftCell="A10"[^>]*> <pane [^>]*topLeftCell="A12"[^>]*> <selection pane="bottomLeft" activeCell="A12" sqref="A12"/>$' '行だけを固定したシートは A 列に戻す'
}

Test-Case 'Reset-ExcelView: Get-Help でヘルプが表示される' {
    $h = Get-Help $ResetView
    Assert-Match $h.Synopsis 'A1'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
