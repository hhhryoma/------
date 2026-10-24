#     ExcelTools.psm1 の共通の部品のテスト（ツールのテストで確かめにくいもの）

Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking

Test-Case 'ExcelTools: Group-InOrder は、まとまりを最初に出てきた順に返す（PowerShell の版によらない）' {
    $items = @(
        [pscustomobject]@{ Process = '結合テスト'; No = 1 },
        [pscustomobject]@{ Process = '単体テスト'; No = 2 },
        [pscustomobject]@{ Process = '結合テスト'; No = 3 },
        [pscustomobject]@{ Process = 'abc'; No = 4 },
        [pscustomobject]@{ Process = 'ABC'; No = 5 })
    $groups = @($items | Group-InOrder Process)
    Assert-Equal @('結合テスト|2|1,3|0', '単体テスト|1|2|1', 'abc|2|4,5|2') @($groups | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Name, $_.Count, (@($_.Group | ForEach-Object { $_.No }) -join ','), $_.Index }) 'Group-Object と同じく大文字小文字は区別しない'
    $byBlock = @($items | Group-InOrder { $_.No % 2 })
    Assert-Equal @('1|3', '0|2') @($byBlock | ForEach-Object { '{0}|{1}' -f $_.Name, $_.Count }) 'スクリプトブロックで値を決める'
    $rows = @(@{ Values = @('a', 'OK') }, @{ Values = @('b', 'NG') }, @{ Values = @('c', 'OK') })
    Assert-Equal @('OK 2', 'NG 1') @($rows | Group-InOrder { $_.Values[1] } | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) '配列を持つ要素も 1 つとして扱う'
    Assert-Equal 0 @(@() | Group-InOrder Process).Count
}
