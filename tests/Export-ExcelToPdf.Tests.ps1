#     Export-ExcelToPdf.ps1 のテスト。PDF 出力は Excel が必要なため、Excel を使わない -CountOnly だけを確かめる

$ExportPdf = Join-Path $script:ToolsDir 'Export-ExcelToPdf.ps1'
$pdfDir = Join-Path $script:Fixtures 'pdf'

function Invoke-CountOnly {
    param([switch] $IgnoreSmallPages)
    $csv = Join-Path (New-TestDir) 'count.csv'
    $null = & $ExportPdf -Path $pdfDir -CountOnly -IgnoreSmallPages:$IgnoreSmallPages -LogPath $csv 6>$null
    $script:LastExit = $LASTEXITCODE
    $byName = @{}
    foreach ($r in (Import-Csv -LiteralPath $csv -Encoding UTF8)) { $byName[[System.IO.Path]::GetFileName($r.ファイル)] = $r }
    return $byName
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

Test-Case 'Export-ExcelToPdf: Get-Help でヘルプが表示される' {
    $h = Get-Help $ExportPdf
    Assert-Match $h.Synopsis 'PDF'
    Assert-True (@($h.parameters.parameter | Where-Object { $_.name -eq 'Extensions' -and $_.description }).Count -eq 1) '-Extensions の説明'
}
