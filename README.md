# Excel 業務ツール

Excel の設計書・成果物を扱う作業を効率化するツール集です。どれも PowerShell のスクリプトで、多くは **Excel がなくても動きます**。

## ツール

| ツール | できること | Excel |
| --- | --- | --- |
| [Search-Excel.ps1](tools/README.md#search-excelps1) | フォルダ配下の Excel を横断検索（値・数式・コメント・図形・シート名・ヘッダー／フッター・名前の定義・入力規則）。結果から該当セルを開ける | 不要 ※1 |
| [Replace-Excel.ps1](tools/README.md#replace-excelps1) | 複数の Excel の文字を一括置換（値・コメント・図形・ヘッダー／フッター・入力規則のリスト、数式も可）。置換リストでまとめて置換もできる。確認してから書き換え、部分ごとの書式は保つ | 不要 |
| [Test-ExcelTerms.ps1](tools/README.md#test-exceltermsps1) | 表記ゆれ（ユーザ／ユーザー、ＩＤ／ID など）を用語集と自動の検出で見つけ、置換リストを作る | 不要 ※1 |
| [Test-ExcelDeliverable.ps1](tools/README.md#test-exceldeliverableps1) | 提出前チェック（体裁・消し忘れ・情報漏れ・壊れた箇所・印刷のはみ出し・表紙と改訂履歴の版数の食い違い）。プロパティの個人名・ヘッダーのファイルの場所・表示の状態は `-Fix` で直せる | 不要 |
| [Reset-ExcelView.ps1](tools/README.md#reset-excelviewps1) | 開いたときに先頭シートの A1 が表示される状態に一括で直す | 不要 |
| [Compare-Excel.ps1](tools/README.md#compare-excelps1) | 2 つのブック・フォルダのセル単位の差分（行・列の挿入や削除・移動・シート名変更を考慮。書式や長い文字列の変わった部分も示し、新のブックに印を付けた写しや、改訂履歴の記載漏れの確認・下書きもできる） | 不要 ※1 |
| [Test-DesignConsistency.ps1](tools/README.md#test-designconsistencyps1) | 設計書どうしの参照の食い違い（定義されていないテーブル・カラム・メッセージ・コード・画面、重複した定義、画面一覧との過不足、画面遷移の食い違い、桁数・必須の違い）。CRUD 図と影響調査（使っている場所の一覧）も作る | 不要 ※1 |
| [Test-Traceability.ps1](tools/README.md#test-traceabilityps1) | 要件一覧 → 設計書 → テスト仕様書のつながりと、画面・イベント・メッセージ・IF にテストがあるか（テストの抜け）を全件確かめる | 不要 ※1 |
| [Export-TableDdl.ps1](tools/README.md#export-tableddlps1) | テーブル定義書から DDL（Oracle・PostgreSQL・SQL Server・MySQL）を作り、DB の列の一覧との食い違い（型・桁数・NULL）も確かめる | 不要 ※1 |
| [New-DesignDocument.ps1](tools/README.md#new-designdocumentps1) | 一覧から雛形の差し込み欄を埋めた設計書を一括で作る（1 件 1 シートの設計書には、雛形のシートからシートを足す）。版上げ（表紙の版数・更新日、改訂履歴の行、ファイル名の版数）もし、前の版との差分から改訂内容を書ける | 不要 |
| [Measure-Progress.ps1](tools/README.md#measure-progressps1) | WBS から遅延タスク（営業日）と PV／EV／SPI を出し、スケジュール遅延リカバリ計画のテンプレートに書き込む | 不要 ※1 |
| [Measure-DesignSize.ps1](tools/README.md#measure-designsizeps1) | 設計書から画面・テーブル・IF などの規模を数え、係数表で工数の目安と FP（NESMA）を出す。前の版との変更の規模（追加見積もりの根拠）と、実績の工数から係数表も作れる | 不要 ※1 |
| [Test-ExcelEstimate.ps1](tools/README.md#test-excelestimateps1) | 見積書の検算（合計の範囲漏れ・手入力の合計・周りと違う数式・文字列の数値・換算の日数や単価の食い違い） | 不要 |
| [Export-ExcelToMarkdown.ps1](tools/README.md#export-exceltomarkdownps1) | Excel を Markdown にする（表は Markdown の表に）。`git diff` で Excel の変更を文字の差分で見られる | 不要 ※1 |
| [Measure-ReviewQuality.ps1](tools/README.md#measure-reviewqualityps1) | 指摘管理表を集計し、設計書ごとの指摘件数・指摘密度・未完了・期限切れ、基準値から外れた設計書、起票・完了の推移を出す。レビュー記録の工数から、レビュー工数密度 × 指摘密度のゾーン分析もする | 不要 ※1 |
| [Measure-TestQuality.ps1](tools/README.md#measure-testqualityps1) | テスト仕様書と障害管理表から、機能ごとのテスト密度・バグ密度とゾーン分析、テストの消化とバグの収束（信頼度成長曲線）、再テスト漏れなどの食い違いを出す | 不要 ※1 |
| [Test-TestSpec.ps1](tools/README.md#test-testspecps1) | テスト仕様書の期待結果の空欄・あいまいさ、重複したケース、異常系・必須・型・境界値・コード値のケースの不足を見つける | 不要 ※1 |
| [Measure-DefectRemoval.ps1](tools/README.md#measure-defectremovalps1) | バグとレビューの指摘の混入工程 × 摘出工程、工程ごとの摘出率・流出、原因の多い順を出す | 不要 ※1 |
| [Measure-ReviewEffectiveness.ps1](tools/README.md#measure-revieweffectivenessps1) | 設計書ごとにレビューの指摘と設計で入ったバグを突き合わせ、レビューで見逃した設計書と設計工程の摘出率を出す | 不要 ※1 |
| [Compare-ReviewFindings.ps1](tools/README.md#compare-reviewfindingsps1) | AI のレビューと人のレビューの指摘を突き合わせ、一致・AI だけ・人だけと、観点ごとの一致率を出す | 不要 ※1 |
| [Measure-CodeSize.ps1](tools/README.md#measure-codesizeps1) | ソースコードのステップ数（KLOC）を機能ごとに数え、前の版との新規・修正・削除も出す | 不要 |
| [New-QualityLimits.ps1](tools/README.md#new-qualitylimitsps1) | 過去のプロジェクトの品質の実績から、社内の基準値（下限・上限）の CSV を作る | 不要 |
| [New-QualityReport.ps1](tools/README.md#new-qualityreportps1) | レビュー・テスト・工程・網羅・整合性の各ツールをまとめて実行し、工程完了判定に使う品質評価のブック（指標・実績・判定 ○△×・要対応）を作る | 不要 ※1 |
| [Invoke-DesignReview.ps1](tools/README.md#invoke-designreviewps1) | 設計書を観点表に沿って Microsoft 365 Copilot にレビューさせ、指摘一覧にまとめる | 不要 |
| [Export-ExcelToPdf.ps1](tools/README.md#export-exceltopdfps1) | フォルダ配下の Excel を PDF に一括出力し、ページ数を数える | 必要 ※2 |
| [Install-SendTo.ps1](tools/README.md#install-sendtops1) | 右クリックの「送る」に各ツールを登録し、PowerShell を使わずにファイルを選ぶだけで実行できるようにする | 不要 ※3 |

※1 .xls / .xlsb を扱う場合のみ Excel が必要です。
※2 PDF のページ数を数えるだけ（`-CountOnly`）なら不要です。
※3 「PDF にする」だけは Excel が必要です。

使い方の詳細は [tools/README.md](tools/README.md) を見てください。`tools` フォルダは**まとめて**配置してください（共通部品の `ExcelTools.psm1` を各スクリプトが使います）。

```powershell
.\tools\Search-Excel.ps1 'TODO' D:\Docs
.\tools\Test-ExcelDeliverable.ps1 D:\提出物 -OutXlsx D:\提出前チェック.xlsx
```

実行ポリシーで止まる場合は `powershell -ExecutionPolicy Bypass -File .\tools\Search-Excel.ps1 ...` のように実行してください。

PowerShell に慣れていない人には、`.\tools\Install-SendTo.ps1` を 1 回実行してもらうと、エクスプローラーでファイルやフォルダを選んで右クリック →「送る」→「Excel 提出前チェック」などから使えます（結果の Excel が開きます）。

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
