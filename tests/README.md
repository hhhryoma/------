# テスト

`tools` 以下のスクリプトのテストです。外部モジュール（Pester など）は使いません。

```powershell
.\tests\Invoke-Tests.ps1                    # すべて実行
.\tests\Invoke-Tests.ps1 -Filter 'Search'   # 名前で絞り込み（正規表現）
```

失敗が 1 件でもあれば終了コード 1 で終わります。

GitHub Actions（`.github/workflows/tests.yml`）で、push と PR のたびに次の 3 環境で実行します。

| 環境 | 目的 |
| --- | --- |
| Windows PowerShell 5.1（Windows） | 社内 PC の標準環境での互換性 |
| PowerShell 7（Windows） | 新しい PowerShell での互換性 |
| PowerShell 7（Linux） | OS に依存する書き方が混ざっていないか |

GitHub Actions の Windows 環境には Excel が無いため、Excel を使う機能（.xls / .xlsb の読み取り、PDF 出力、`-Open`）は
「Excel が無いときに正しくエラーになること」だけを確かめます。

## Windows の手元での動作確認

自動テストでは、Excel で開いたときの見た目（グラフ・リンク・色・表示位置）と、Copilot とのやり取りは確かめられません。
ツールを入れ替えたときなどは、Windows の手元で次を実行してください。

```powershell
.\tests\Invoke-SmokeTest.ps1
```

- サンプルの設計書（`docs\design-sample`）と、その場で作る指摘管理表・テスト仕様書・障害管理表・WBS で全ツールを実行します
- 結果は「ドキュメント\Excel ツールの動作確認_日時」に保存し、最後に `00_確認リスト.xlsx` を開きます
- 確認リストのリンクから結果のファイルを順に開き、「Excel で開いて見るところ」を確かめて、確認の列に OK / NG を書きます
- Copilot（定性分析の下書き・AI レビュー）と「送る」は、確認リストの最後の行の手順で確かめます
- サンプルの設計書は書き換えません（書き換えるツールは写しに対して実行します）。PDF にする（Excel が必要）は Excel が無ければスキップします
- 自動テスト（`Invoke-SmokeTest.Tests.ps1`）でも、全ツールが動いて確認リストができることだけは確かめています

## ファイル

| ファイル | 内容 |
| --- | --- |
| `Invoke-Tests.ps1` | テストの実行と、`Test-Case` / `Assert-*` などの補助関数 |
| `Invoke-SmokeTest.ps1` | Windows の手元での動作確認（全ツールを実行し、結果のファイルと確認リストを作る） |
| `*.Tests.ps1` | ツールごとのテスト |
| `fixtures/` | テスト用のファイル。`fixtures/make_fixtures.py` で生成（openpyxl と LibreOffice が必要。生成物はリポジトリに含めているので通常は実行不要） |

## テストの書き方

```powershell
Test-Case 'ツール名: 確かめること' {
    $dir = Copy-Fixture 'grep'                   # フィクスチャを一時フォルダへコピー（後で自動削除）
    $hits = & $SearchExcel -Pattern 'TODO' -Path $dir 6>$null
    Assert-Equal 13 $hits.Count '件数'
}
```

テストファイルは **UTF-8（BOM 付き）・CRLF** で保存してください（Windows PowerShell 5.1 が日本語を正しく読むため）。
