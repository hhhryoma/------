#     Export-ExcelToPdf.ps1 のテスト。PDF 出力は Excel が必要なため、Excel を使わないところ（-CountOnly・-WhatIf、
#     -SkipExisting ですべてスキップするとき）だけを確かめる

$ExportPdf = Join-Path $script:ToolsDir 'Export-ExcelToPdf.ps1'
$pdfDir = Join-Path $script:Fixtures 'pdf'
$pdfBook = Join-Path $script:Fixtures 'grep/sample.xlsx'

function Invoke-CountOnly {
    param([string] $Path = $pdfDir, [switch] $IgnoreSmallPages)
    $csv = Join-Path (New-TestDir) 'count.csv'
    $null = & $ExportPdf -Path $Path -CountOnly -IgnoreSmallPages:$IgnoreSmallPages -LogPath $csv 6>$null
    $script:LastExit = $LASTEXITCODE
    $byName = @{}
    foreach ($r in (Import-Csv -LiteralPath $csv -Encoding UTF8)) { $byName[[System.IO.Path]::GetFileName($r.ファイル)] = $r }
    return $byName
}

function Invoke-PdfExport {
    #     PDF を出力し、-LogPath の CSV の行を元のファイル名ごとに返す（画面の表示は $script:PdfHost に入れる）
    param([hashtable] $Options)
    $csv = Join-Path (New-TestDir) 'result.csv'
    $script:PdfHost = Get-HostText { & $ExportPdf @Options -LogPath $csv }
    $byName = @{}
    foreach ($r in (Import-Csv -LiteralPath $csv -Encoding UTF8)) { $byName[[System.IO.Path]::GetFileName($r.ファイル)] = $r }
    return $byName
}

function New-UpdatedPdf {
    #     PDF に追記保存を 1 回足した PDF を作る。Acrobat で承認印・注釈・電子署名を足して保存したときと同じく、
    #     書き換えたオブジェクトを同じ番号で末尾に足し、その分の xref と、前の xref を /Prev で指す trailer を足す。
    #     Objects: オブジェクト番号 → 中身（<< … >>）
    param([string] $Source, [string] $Path, [hashtable] $Objects)
    $latin1 = [System.Text.Encoding]::GetEncoding(28591)
    $pdf = $latin1.GetString([System.IO.File]::ReadAllBytes($Source))
    $prev = [regex]::Match($pdf, 'startxref\s+(\d+)\s+%%EOF\s*$').Groups[1].Value
    $size = [int][regex]::Match($pdf, '/Size\s+(\d+)').Groups[1].Value
    $sb = New-Object System.Text.StringBuilder($pdf)
    $xref = New-Object System.Text.StringBuilder("xref`n0 1`n0000000000 65535 f `n")
    foreach ($n in @($Objects.Keys | Sort-Object)) {
        [void]$xref.Append(("{0} 1`n{1:D10} 00000 n `n" -f $n, $sb.Length))
        [void]$sb.Append(("{0} 0 obj`n{1}`nendobj`n" -f $n, $Objects[$n]))
        $size = [math]::Max($size, $n + 1)
    }
    $start = $sb.Length
    [void]$sb.Append($xref.ToString())
    [void]$sb.Append(("trailer`n<< /Size {0} /Root 1 0 R /Prev {1} >>`nstartxref`n{2}`n%%EOF`n" -f $size, $prev, $start))
    [System.IO.File]::WriteAllBytes($Path, $latin1.GetBytes($sb.ToString()))
}

Test-Case 'Export-ExcelToPdf: -CountOnly で PDF のページ数を数える' {
    $r = Invoke-CountOnly
    Assert-Equal '3' $r['a4x3.pdf'].ページ数
    Assert-Equal '3' $r['mixed.pdf'].ページ数
    Assert-Equal '成功' $r['a4x3.pdf'].結果
}

Test-Case 'Export-ExcelToPdf: 壊れた PDF は失敗として記録し、終了コード 1' {
    $r = Invoke-CountOnly
    Assert-Equal '失敗' $r['broken.pdf'].結果
    Assert-Equal 1 $script:LastExit
}

Test-Case 'Export-ExcelToPdf: -IgnoreSmallPages で極端に小さいページを集計から除く' {
    $r = Invoke-CountOnly -IgnoreSmallPages
    Assert-Equal '2' $r['mixed.pdf'].ページ数
    Assert-Equal '1' $r['mixed.pdf'].除外ページ
    Assert-Equal '3' $r['a4x3.pdf'].ページ数 '同じ大きさのページは除かない'
}

Test-Case 'Export-ExcelToPdf: 承認印・注釈などを追記保存した PDF は、末尾に足された同じページを重ねて数えず、最後に保存した大きさで判定する' {
    $dir = New-TestDir
    # 3 ページの PDF の 1 ページ目に承認印（スタンプの注釈）を押して保存したもの
    New-UpdatedPdf (Join-Path $pdfDir 'a4x3.pdf') (Join-Path $dir 'stamped.pdf') @{
        3 = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Annots [6 0 R] >>'
        6 = '<< /Type /Annot /Subtype /Stamp /Rect [400 700 500 800] /Name /Approved >>'
    }
    # 小さいページ（3 ページ目）を A4 に直して保存したもの
    New-UpdatedPdf (Join-Path $pdfDir 'mixed.pdf') (Join-Path $dir 'resized.pdf') @{ 5 = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>' }
    $r = Invoke-CountOnly -Path $dir -IgnoreSmallPages
    Assert-Equal '3|0' ('{0}|{1}' -f $r['stamped.pdf'].ページ数, $r['stamped.pdf'].除外ページ)
    Assert-Equal '3|0' ('{0}|{1}' -f $r['resized.pdf'].ページ数, $r['resized.pdf'].除外ページ) '直す前の小さいページを数えた'
}

Test-Case 'Export-ExcelToPdf: -WhatIf では Excel を起動せず、対象のファイルと出力する PDF の名前を一覧にする' {
    $src = New-TestDir
    New-Item -ItemType Directory -Path (Join-Path $src 'sub') | Out-Null
    Copy-Item -LiteralPath $pdfBook -Destination (Join-Path $src '基本設計書.xlsx')
    Copy-Item -LiteralPath $pdfBook -Destination (Join-Path $src 'sub/画面設計書.xlsm')
    $out = Join-Path (New-TestDir) 'pdf'
    $text = Get-HostText { & $ExportPdf -Path $src -OutputRoot $out -WhatIf }
    Assert-Match $text '(?m)^基本設計書\.xlsx\s+基本設計書\.pdf\s+-\s+WhatIf'
    Assert-Match $text '(?m)^sub[\\/]画面設計書\.xlsm\s+画面設計書\.pdf\s+-\s+WhatIf'
    Assert-True (-not (Test-Path -LiteralPath $out)) '-WhatIf で出力先のフォルダを作った'
    $text = Get-HostText { & $ExportPdf -Path $src -PerSheet -WhatIf }
    Assert-Match $text '(?m)^基本設計書\.xlsx\s+基本設計書_\*\.pdf\s+-\s+WhatIf' 'シート単位の PDF の名前は、ブックを開くまで分からない'
}

Test-Case 'Export-ExcelToPdf: 拡張子だけが違う同じ名前のブックは、PDF の名前に拡張子も含めて別々にし、-SkipExisting も別々に判定する' {
    $src = New-TestDir
    foreach ($name in @('設計書.xls', '設計書.xlsx', '一覧.xlsx')) { Copy-Item -LiteralPath $pdfBook -Destination (Join-Path $src $name) }
    $text = Get-HostText { & $ExportPdf -Path $src -WhatIf }
    Assert-Match $text '(?m)^設計書\.xls\s+設計書\.xls\.pdf\s'
    Assert-Match $text '(?m)^設計書\.xlsx\s+設計書\.xlsx\.pdf\s'
    Assert-Match $text '(?m)^一覧\.xlsx\s+一覧\.pdf\s' '重ならない名前はそのまま'
    $text = Get-HostText { & $ExportPdf -Path $src -PerSheet -WhatIf }
    Assert-Match $text '(?m)^設計書\.xlsx\s+設計書\.xlsx_\*\.pdf\s' 'シート単位'

    # どの PDF も Excel より新しければ、Excel を起動せずにすべてスキップし、それぞれの PDF のページ数を数える。
    # 前に出力した 設計書.pdf は、どちらのブックの PDF とも見なさない
    foreach ($f in @(Get-ChildItem -LiteralPath $src -File)) { $f.LastWriteTime = (Get-Date).AddMinutes(-10) }
    foreach ($name in @('設計書.pdf', '設計書.xlsx.pdf', '一覧.pdf')) { Copy-Item -LiteralPath (Join-Path $pdfDir 'a4x3.pdf') -Destination (Join-Path $src $name) }
    # 設計書.xls.pdf は、3 ページの PDF に 1 ページ足して保存したもの
    New-UpdatedPdf (Join-Path $pdfDir 'a4x3.pdf') (Join-Path $src '設計書.xls.pdf') @{
        2 = '<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R 6 0 R] /Count 4 >>'
        6 = '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>'
    }
    foreach ($f in @(Get-ChildItem -LiteralPath $src -Filter '*.pdf')) { $f.LastWriteTime = Get-Date }
    $r = Invoke-PdfExport @{ Path = $src; SkipExisting = $true }
    Assert-Equal 'スキップ|スキップ|スキップ' (@('設計書.xls', '設計書.xlsx', '一覧.xlsx' | ForEach-Object { $r[$_].結果 }) -join '|')
    Assert-Equal (Join-Path $src '設計書.xls.pdf') $r['設計書.xls'].出力先
    Assert-Equal '4|3' ('{0}|{1}' -f $r['設計書.xls'].ページ数, $r['設計書.xlsx'].ページ数)
    Assert-Match $script:PdfHost '合計ページ数 10'
}

Test-Case 'Export-ExcelToPdf: -Path に Excel のファイルを指定すると、そのファイルだけを同じフォルダ（-OutputRoot の直下）に出力する（「送る」の PDF にする）' {
    $src = New-TestDir
    Copy-Item -LiteralPath $pdfBook -Destination (Join-Path $src '基本設計書.xlsx')
    Copy-Item -LiteralPath $pdfBook -Destination (Join-Path $src '詳細設計書.xlsx')
    $text = Get-HostText { & $ExportPdf -Path (Join-Path $src '基本設計書.xlsx') -WhatIf }
    Assert-Match $text '対象 1 ファイル'
    Assert-Match $text '(?m)^基本設計書\.xlsx\s+基本設計書\.pdf\s+-\s+WhatIf'
    # -OutputRoot の直下の PDF が新しければ、Excel を起動せずにスキップする
    $out = New-TestDir
    (Get-Item -LiteralPath (Join-Path $src '基本設計書.xlsx')).LastWriteTime = (Get-Date).AddMinutes(-10)
    Copy-Item -LiteralPath (Join-Path $pdfDir 'a4x3.pdf') -Destination (Join-Path $out '基本設計書.pdf')
    (Get-Item -LiteralPath (Join-Path $out '基本設計書.pdf')).LastWriteTime = Get-Date
    $r = Invoke-PdfExport @{ Path = (Join-Path $src '基本設計書.xlsx'); OutputRoot = $out; SkipExisting = $true }
    Assert-Equal '基本設計書.xlsx' (@($r.Keys) -join ', ')
    Assert-Equal ('スキップ|3|' + (Join-Path $out '基本設計書.pdf')) ('{0}|{1}|{2}' -f $r['基本設計書.xlsx'].結果, $r['基本設計書.xlsx'].ページ数, $r['基本設計書.xlsx'].出力先)
    # -CountOnly では PDF のファイルを指定できる
    $r = Invoke-CountOnly -Path (Join-Path $pdfDir 'a4x3.pdf')
    Assert-Equal 'a4x3.pdf|3' ('{0}|{1}' -f (@($r.Keys) -join ', '), $r['a4x3.pdf'].ページ数)
}

Test-Case 'Export-ExcelToPdf: Get-Help でヘルプが表示される' {
    $h = Get-Help $ExportPdf
    Assert-Match $h.Synopsis 'PDF'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -eq 'Extensions' -and $_.description }).Count -eq 1) '-Extensions の説明'
}
