# Excel 業務ツール

Excel の設計書・成果物を扱う作業を効率化するツール集です。どれも PowerShell のスクリプトで、多くは **Excel がなくても動きます**。

## ツール

| ツール | できること | Excel |
| --- | --- | --- |
| [Search-Excel.ps1](tools/README.md#search-excelps1) | フォルダ配下の Excel を横断検索（値・数式・コメント・図形・シート名）。結果から該当セルを開ける | 不要 ※1 |
| [Test-ExcelDeliverable.ps1](tools/README.md#test-exceldeliverableps1) | 提出前チェック（体裁・消し忘れ・情報漏れ・壊れた箇所） | 不要 |
| [Reset-ExcelView.ps1](tools/README.md#reset-excelviewps1) | 開いたときに先頭シートの A1 が表示される状態に一括で直す | 不要 |
| [Compare-Excel.ps1](tools/README.md#compare-excelps1) | 2 つのブック・フォルダのセル単位の差分（行の挿入・削除を考慮） | 不要 ※1 |
| [Invoke-DesignReview.ps1](tools/README.md#invoke-designreviewps1) | 設計書を観点表に沿って Microsoft 365 Copilot にレビューさせ、指摘一覧にまとめる | 不要 |
| [Export-ExcelToPdf.ps1](tools/README.md#export-exceltopdfps1) | フォルダ配下の Excel を PDF に一括出力し、ページ数を数える | 必要 ※2 |

※1 .xls / .xlsb を扱う場合のみ Excel が必要です。
※2 PDF のページ数を数えるだけ（`-CountOnly`）なら不要です。

使い方の詳細は [tools/README.md](tools/README.md) を見てください。`tools` フォルダは**まとめて**配置してください（共通部品の `ExcelTools.psm1` を各スクリプトが使います）。

```powershell
.\tools\Search-Excel.ps1 'TODO' D:\Docs
.\tools\Test-ExcelDeliverable.ps1 D:\提出物 -OutXlsx D:\提出前チェック.xlsx
```

実行ポリシーで止まる場合は `powershell -ExecutionPolicy Bypass -File .\tools\Search-Excel.ps1 ...` のように実行してください。

## フォルダ構成

| フォルダ | 内容 |
| --- | --- |
| `tools/` | ツール本体と共通部品（`ExcelTools.psm1`）、設計レビューの観点表と Copilot プロンプト集（`review-checklists/`） |
| `tests/` | ツールのテスト（[tests/README.md](tests/README.md)）。GitHub Actions で Windows PowerShell 5.1・PowerShell 7 の両方で実行 |
| `docs/design-sample/` | 設計ツールの検証用の詳細設計サンプル（架空の受注管理システム） |
| `docs/schedule-recovery/` | スケジュール遅延リカバリ計画のテンプレートと記入例 |
| `vba-files/`・`Book1.xlsm` | Excel マクロ（XVBA で管理）。「体裁整える」マクロは `Reset-ExcelView.ps1` で置き換え可能 |

## 動作環境

- Windows PowerShell 5.1 以降（PowerShell 7 でも動作）
- Excel を使う機能はデスクトップ版 Excel が必要
