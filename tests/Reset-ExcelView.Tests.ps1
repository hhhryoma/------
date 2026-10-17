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
    Assert-Equal 1 @(Get-ChildItem -LiteralPath $backup -File).Count '直したファイルだけをバックアップ'
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

Test-Case 'Reset-ExcelView: Get-Help でヘルプが表示される' {
    $h = Get-Help $ResetView
    Assert-Match $h.Synopsis 'A1'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -notin @('WhatIf', 'Confirm') -and -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
