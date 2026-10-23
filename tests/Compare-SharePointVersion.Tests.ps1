#     Compare-SharePointVersion.ps1 のテスト。SharePoint・Excel の代わりに、-Fetcher で URL ごとのブックを写す

$SpVersion = Join-Path $script:ToolsDir 'Compare-SharePointVersion.ps1'

function New-SpBook {
    param([string] $Path, [string] $Name)
    Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    Export-Workbook -FilePath $Path -Sheets @(@{
        Name = 'Data'
        Columns = @(@{ Header = 'ID'; Width = 10 }, @{ Header = '名前'; Width = 20 })
        Rows = @(@{ Values = @('1', $Name) }, @{ Values = @('2', '顧客') })
    })
}

function New-SpFetcher {
    #     URL の _vti_history の番号（現在の版は current）ごとのブックを写す。取り出した URL を $Log に残す
    param([hashtable] $Books, [System.Collections.Generic.List[string]] $Log)
    return {
        param($Url, $Dest)
        $Log.Add($Url)
        $m = [regex]::Match($Url, '/_vti_history/(\d+)/')
        $key = 'current'
        if ($m.Success) { $key = $m.Groups[1].Value }
        if (-not $Books.ContainsKey($key)) { throw "その版はありません（$key）" }
        Copy-Item -LiteralPath $Books[$key] -Destination $Dest
    }.GetNewClosure()
}

Test-Case 'Compare-SharePointVersion: URL の版（_vti_history/番号）と現在の版を取り出して比べ、差分の HTML を保存する' {
    $dir = New-TestDir
    New-SpBook (Join-Path $dir 'src/v3.xlsx') '受注'
    New-SpBook (Join-Path $dir 'src/now.xlsx') '受注明細'
    $log = New-Object System.Collections.Generic.List[string]
    $fetch = New-SpFetcher @{ '1536' = (Join-Path $dir 'src/v3.xlsx'); 'current' = (Join-Path $dir 'src/now.xlsx') } $log
    $out = Join-Path $dir 'out'
    $url = 'https://contoso.sharepoint.com/sites/juchu/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03.xlsx?web=1'
    $diffs = @(& $SpVersion $url 3.0 -OutDir $out -NoOpen -Fetcher $fetch 6>$null)
    Assert-Equal @(
        'https://contoso.sharepoint.com/sites/juchu/_vti_history/1536/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03.xlsx',
        'https://contoso.sharepoint.com/sites/juchu/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03.xlsx') @($log)
    Assert-Equal 'Data!B2|値|変更' (@($diffs | ForEach-Object { '{0}|{1}|{2}' -f $_.Location, $_.Kind, $_.Change }) -join ' / ')
    Assert-True (Test-Path -LiteralPath (Join-Path $out 'DD03_v3.0-現在.html')) '差分の HTML'
    Assert-True (Test-Path -LiteralPath (Join-Path $out 'DD03_v3.0.xlsx')) '取り出した版を残す'

    # 番号の付いた版は取り出し直さない。版の指定は 3・v3.0 などでもよく、新を指定すると版どうしで比べる。旧と新が逆なら入れ替える
    $log.Clear()
    New-SpBook (Join-Path $dir 'src/v5.xlsx') '受注'
    $fetch = New-SpFetcher @{ '1536' = (Join-Path $dir 'src/v3.xlsx'); '2561' = (Join-Path $dir 'src/v5.xlsx') } $log
    $text = Get-HostText { $script:d2 = @(& $SpVersion 'https://contoso.sharepoint.com/:x:/r/sites/juchu/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03.xlsx?d=w1&csf=1' 'ｖ５．１' -NewVersion 3 -OutDir $out -NoOpen -Fetcher $fetch) }
    Assert-Equal @('https://contoso.sharepoint.com/sites/juchu/_vti_history/2561/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03.xlsx') @($log) 'v3.0 は前に取り出したもの'
    Assert-Match $text '旧: v3\.0'
    Assert-Match $text '新: v5\.1'
    Assert-Equal 0 @($script:d2).Count
}

Test-Case 'Compare-SharePointVersion: OneDrive で同期しているファイルは、SharePoint の URL に直し、現在の版はそのファイルを使う' {
    $dir = New-TestDir
    $mount = Join-Path $dir '会社 - 受注システム - ドキュメント'
    $local = Join-Path (Join-Path $mount '設計') 'DD03 テーブル定義書.xlsx'
    New-SpBook $local '受注明細'
    New-SpBook (Join-Path $dir 'src/v2.xlsx') '受注'
    $log = New-Object System.Collections.Generic.List[string]
    $fetch = New-SpFetcher @{ '1024' = (Join-Path $dir 'src/v2.xlsx') } $log
    $roots = @(
        [pscustomobject]@{ MountPoint = $dir; UrlNamespace = 'https://contoso-my.sharepoint.com/personal/yamada/Documents/'; WebUrl = $null },
        [pscustomobject]@{ MountPoint = $mount; UrlNamespace = 'https://contoso.sharepoint.com/sites/juchu/Shared Documents/'; WebUrl = 'https://contoso.sharepoint.com/sites/juchu' })
    $diffs = @(& $SpVersion $local 2 -OutDir (Join-Path $dir 'out') -NoOpen -Fetcher $fetch -SyncRoots $roots 6>$null)
    Assert-Equal @('https://contoso.sharepoint.com/sites/juchu/_vti_history/1024/Shared%20Documents/%E8%A8%AD%E8%A8%88/DD03%20%E3%83%86%E3%83%BC%E3%83%96%E3%83%AB%E5%AE%9A%E7%BE%A9%E6%9B%B8.xlsx') @($log) '一番深い同期フォルダの URL'
    Assert-Equal 1 $diffs.Count
    Assert-True (Test-Path -LiteralPath (Join-Path $dir 'out/DD03 テーブル定義書_v2.0-現在.html'))

    # サブサイトは -SiteUrl で指定する
    $log.Clear()
    $fetch = New-SpFetcher @{ '1024' = (Join-Path $dir 'src/v2.xlsx'); 'current' = $local } $log
    $null = & $SpVersion 'https://contoso.sharepoint.com/sites/juchu/sub/Docs/a.xlsx' 2 -SiteUrl 'https://contoso.sharepoint.com/sites/juchu/sub' -OutDir (Join-Path $dir 'out2') -NoOpen -Fetcher $fetch 6>$null
    Assert-Equal 'https://contoso.sharepoint.com/sites/juchu/sub/_vti_history/1024/Docs/a.xlsx' $log[0]
}

Test-Case 'Compare-SharePointVersion: 使えない URL・同期していないファイル・版の番号・取り出せない版を知らせる' {
    $dir = New-TestDir
    $fetch = New-SpFetcher @{} (New-Object System.Collections.Generic.List[string])
    $o = @{ OutDir = (Join-Path $dir 'out'); NoOpen = $true; Fetcher = $fetch; SyncRoots = @() }
    Assert-Throws { & $SpVersion 'https://contoso.sharepoint.com/:x:/s/juchu/EabcDEF?e=1' 3 @o 6>$null } '共有リンク'
    Assert-Throws { & $SpVersion 'https://contoso.sharepoint.com/sites/juchu/_layouts/15/Doc.aspx?sourcedoc=%7B1%7D' 3 @o 6>$null } 'パスのコピー'
    New-SpBook (Join-Path $dir 'a.xlsx') 'x'
    Assert-Throws { & $SpVersion (Join-Path $dir 'a.xlsx') 3 @o 6>$null } '同期している SharePoint のフォルダのファイルではありません'
    $url = 'https://contoso.sharepoint.com/sites/juchu/Shared%20Documents/a.xlsx'
    Assert-Throws { & $SpVersion $url '三' @o 6>$null } 'バージョン履歴'
    Assert-Throws { & $SpVersion $url '現在' @o 6>$null } 'どちらも現在の版'
    Assert-Throws { & $SpVersion $url 9.0 @o 6>$null } 'v9\.0 を取り出せませんでした: その版はありません'
}

Test-Case 'Compare-SharePointVersion: Get-Help でヘルプが表示される' {
    $h = Get-Help $SpVersion
    Assert-Match $h.Synopsis 'SharePoint'
    # テスト用の隠したパラメーター（-Fetcher・-SyncRoots）を除いた、説明のあるパラメーター
    Assert-Equal 11 @($h.parameters.parameter | Where-Object { $_.description }).Count
}
