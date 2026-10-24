# Excel 業務ツール

Excel の設計書・成果物を扱う作業を効率化するツール集です。どれも PowerShell のスクリプトで、多くは **Excel がなくても動きます**。

**どの工程でどのツールを使うかは [工程別ガイド](docs/工程別ガイド.md) にまとめています。**

## 画面から使う（PowerShell を使わない）

`tools` フォルダの **`Start-ExcelTools.cmd` をダブルクリック**すると、ツールの画面が開きます。Install-ExcelTools で入れた PC では、**スタートメニューの「Excel ツール」**から開けます（右クリックの「送る」→「Excel ツールの画面で開く」なら、選んだファイルを入れた状態で開きます）。

1. 左の一覧から、やりたいこと（提出前チェック・2 つを比べる・文字を探す など）を選ぶ
2. 右の入力欄に、ファイルやフォルダを入れる（「ファイル…」「フォルダ…」のボタンか、エクスプローラーからドラッグ＆ドロップ）
3. 「実行」を押す → 下の記録の欄に進み具合が出て、終わると結果の Excel（差分は画面）が開く

- 結果の保存先を空欄にすると、「ドキュメント\Excel ツールの結果」に「名前_日時.xlsx」で保存します
- ファイルを書き換える処理（置換・表示を整える・版を上げる）は、最初は「試しに実行」で書き換える箇所を確かめられます
- 入力は保存され、次に開いたときに出ます。各ツールのすべての設定は「詳細な設定も表示する」で入力できます
- 詳しくは [Start-ExcelTools.ps1](tools/README.md#start-exceltoolsps1画面) を見てください

## ツール

担当者向け・管理者向けは、主に使う人の目安です。どのツールも誰が使ってもかまいません。

### 担当者向け

設計書・テスト仕様書を作る人、直す人、レビューする人が、毎日の作業で使うツールです。多くは右クリックの「送る」から使えます。

#### 設計書を作る・直す

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [New-DesignDocument.ps1](tools/README.md#new-designdocumentps1) | 一覧から雛形の差し込み欄を埋めた設計書を一括で作る（1 件 1 シートの設計書には、雛形のシートからシートを足す）。版上げ（表紙の版数・更新日、改訂履歴の行、ファイル名の版数）もし、前の版との差分から改訂内容を書ける | 不要 | － |
| [Search-Excel.ps1](tools/README.md#search-excelps1) | フォルダ配下の Excel を横断検索（値・数式・コメント・図形・シート名・ヘッダー／フッター・名前の定義・入力規則）。Word の文書も探せる。結果から該当セルを開ける | 不要 ※1 | 検索 |
| [Replace-Excel.ps1](tools/README.md#replace-excelps1) | 複数の Excel の文字を一括置換（値・コメント・図形・ヘッダー／フッター・入力規則のリスト、数式も可）。置換リストでまとめて置換もできる。確認してから書き換え、部分ごとの書式は保つ | 不要 | － |
| [Test-ExcelTerms.ps1](tools/README.md#test-exceltermsps1) | 表記ゆれ（ユーザ／ユーザー、ＩＤ／ID など）を用語集と自動の検出で見つけ、置換リストを作る。Word の文書も調べられる | 不要 ※1 | 表記ゆれ |
| [Test-DesignConsistency.ps1](tools/README.md#test-designconsistencyps1) | 設計書どうしの参照の食い違い（定義されていないテーブル・カラム・メッセージ・コード・画面、重複した定義、画面一覧との過不足、画面遷移の食い違い、桁数・必須の違い）。CRUD 図と影響調査（使っている場所の一覧）も作る | 不要 ※1 | 設計書の整合性チェック |
| [Test-ScreenNumbers.ps1](tools/README.md#test-screennumbersps1) | 画面レイアウトのスクリーンショットに重ねた丸数字（①・No.5 など）を画面項目定義の No と突き合わせ、無い番号・重なった番号を見つける。前の版と比べて、項目を足して No がずれたのに直していない丸数字も見つける | 不要 ※1 | － |
| [Test-Screenshots.ps1](tools/README.md#test-screenshotsps1) | 設計書に貼った画面のスクリーンショットが、元のワイヤーフレーム（Excel）より古くなっていないかを確かめる（貼り替えたあとに記録し、元の印刷範囲の中が変わると知らせる。設計書は変更しない） | 不要 | － |
| [Export-TableDdl.ps1](tools/README.md#export-tableddlps1) | テーブル定義書から DDL（Oracle・PostgreSQL・SQL Server・MySQL）を作り、DB の列の一覧との食い違い（型・桁数・NULL）も確かめる | 不要 ※1 | － |

#### 変わったところを見る

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [Compare-Excel.ps1](tools/README.md#compare-excelps1) | 2 つのブック・フォルダ（サブフォルダも含む）のセル単位の差分（行・列の挿入や削除・移動・シート名変更を考慮。書式や長い文字列の変わった部分、図形の移動と画像（スクリーンショット）の貼り替えも示し、新のブックに印を付けた写し、シートの方眼のまま色付けしてブラウザーで見る HTML、改訂履歴の記載漏れの確認・下書きもできる。Word（.docx）も段落・表の行で比べる） | 不要 ※1 | 差分（2 つ選ぶ） |
| [Compare-SharePointVersion.ps1](tools/README.md#compare-sharepointversionps1) | SharePoint（Microsoft 365）のファイルの版（バージョン履歴の 3.0 など）を取り出し、現在の版や別の版と比べて差分の画面を開く | 必要 ※4 | SharePoint の版と比べる |
| [Export-ExcelToMarkdown.ps1](tools/README.md#export-exceltomarkdownps1) | Excel を Markdown にする（表は Markdown の表に）。`git diff` で Excel の変更を文字の差分で見られる | 不要 ※1 | － |

#### レビュー・提出

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [Invoke-DesignReview.ps1](tools/README.md#invoke-designreviewps1) | Excel・Word の設計書を観点表に沿って Microsoft 365 Copilot にレビューさせ（数える・突き合わせるだけの観点はツールで確かめ）、指摘一覧にまとめる | 不要 | － |
| [Test-ExcelDeliverable.ps1](tools/README.md#test-exceldeliverableps1) | 提出前チェック（体裁・消し忘れ・情報漏れ・壊れた箇所・印刷のはみ出し・表紙と改訂履歴の版数の食い違い・スクリーンショットの縦横比の崩れや拡大しすぎ）。Word（.docx）の変更履歴・コメント・隠し文字の残りも。プロパティの個人名・ヘッダーのファイルの場所・表示の状態は `-Fix` で直せる | 不要 | 提出前チェック |
| [Reset-ExcelView.ps1](tools/README.md#reset-excelviewps1) | 開いたときに先頭シートの A1 が表示される状態に一括で直す | 不要 | 表示を整える |
| [Export-ExcelToPdf.ps1](tools/README.md#export-exceltopdfps1) | フォルダ配下の Excel を PDF に一括出力し、ページ数を数える | 必要 ※2 | PDF にする |

#### テスト

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [New-TestSpecDraft.ps1](tools/README.md#new-testspecdraftps1) | 画面設計書の必須・桁数・型・入力チェック・コード値・イベントから、テスト仕様書の下書き（メッセージの文言付きの期待結果）を作る | 不要 ※1 | テスト仕様書の下書き |
| [Test-TestSpec.ps1](tools/README.md#test-testspecps1) | テスト仕様書の期待結果の空欄・あいまいさ、重複したケース、異常系・必須・型・境界値・コード値のケースの不足を見つける | 不要 ※1 | － |
| [Test-Traceability.ps1](tools/README.md#test-traceabilityps1) | 要件一覧 → 設計書 → テスト仕様書のつながりと、画面・イベント・メッセージ・IF にテストがあるか（テストの抜け）を全件確かめる | 不要 ※1 | － |
| [Select-RegressionTests.ps1](tools/README.md#select-regressiontestsps1) | 設計書の旧版と新版の差分から、再テストするテストケースを選ぶ（改修の回帰テストの範囲決め） | 不要 ※1 | － |

### 管理者向け

PL・PM・品質管理の担当が、見積もり・進捗・品質を数字で見るときに使うツールです。指摘管理表・テスト仕様書・障害管理表・WBS などを集計します。

#### 見積もり・進捗

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [Measure-DesignSize.ps1](tools/README.md#measure-designsizeps1) | 設計書から画面・テーブル・IF などの規模を数え、係数表で工数の目安と FP（NESMA）を出す。前の版との変更の規模（追加見積もりの根拠）と、実績の工数から係数表も作れる | 不要 ※1 | 規模の見積もり |
| [Test-ExcelEstimate.ps1](tools/README.md#test-excelestimateps1) | 見積書の検算（合計の範囲漏れ・手入力の合計・周りと違う数式・文字列の数値・換算の日数や単価の食い違い） | 不要 | 見積書の検算 |
| [Measure-Progress.ps1](tools/README.md#measure-progressps1) | WBS から遅延タスク（営業日）と PV／EV／SPI を出し、スケジュール遅延リカバリ計画のテンプレートに書き込む | 不要 ※1 | － |
| [Measure-Effort.ps1](tools/README.md#measure-effortps1) | 日報・工数実績表の工数を、タスク・担当者・週ごとに集計し、WBS の予定工数と比べて予定超過・超過の見込みのタスクと CPI を出す（WBS に無い作業も） | 不要 ※1 | － |
| [Export-ActionItems.ps1](tools/README.md#export-actionitemsps1) | 議事録（Word・Excel・テキスト）から宿題（担当・期限付き）と決定事項を抜き出し、課題管理表の形の一覧にする（週次報告で期限切れを追える） | 不要 ※1 | 議事録の宿題 |
| [New-WeeklyReport.ps1](tools/README.md#new-weeklyreportps1) | WBS・指摘管理表・障害管理表・課題（QA）管理表・工数実績から週次報告（今週の完了・遅延・来週の予定、指摘・障害・課題の起票・完了・期限切れ、今週の工数と予定超過のタスク、担当者ごとのやること）を作り、メールに貼れる文章・担当者ごとのリマインドの文章と Excel に保存する。毎週の数字を残して前回と比べる | 不要 ※1 | 週次報告 |

#### 品質の集計・分析

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [Measure-ReviewQuality.ps1](tools/README.md#measure-reviewqualityps1) | 指摘管理表を集計し、設計書ごとの指摘件数・指摘密度・未完了・期限切れ、基準値から外れた設計書、起票・完了の推移を出す。レビュー記録の工数から、レビュー工数密度 × 指摘密度のゾーン分析もする | 不要 ※1 | － |
| [Measure-TestQuality.ps1](tools/README.md#measure-testqualityps1) | テスト仕様書と障害管理表から、機能ごとのテスト密度・バグ密度とゾーン分析、テストの消化とバグの収束（信頼度成長曲線）、再テスト漏れなどの食い違いを出す | 不要 ※1 | － |
| [Measure-DefectRemoval.ps1](tools/README.md#measure-defectremovalps1) | バグとレビューの指摘の混入工程 × 摘出工程、工程ごとの摘出率・流出、原因の多い順を出す | 不要 ※1 | － |
| [Measure-ReviewEffectiveness.ps1](tools/README.md#measure-revieweffectivenessps1) | 設計書ごとにレビューの指摘と設計で入ったバグを突き合わせ、レビューで見逃した設計書と設計工程の摘出率を出す | 不要 ※1 | － |
| [Compare-ReviewFindings.ps1](tools/README.md#compare-reviewfindingsps1) | AI のレビューと人のレビューの指摘を突き合わせ、一致・AI だけ・人だけと、観点ごとの一致率、観点表への反映案を出す | 不要 ※1 | － |
| [Measure-CodeSize.ps1](tools/README.md#measure-codesizeps1) | ソースコードのステップ数（KLOC）を機能ごとに数え、前の版との新規・修正・削除も出す | 不要 | － |

#### 品質会議・基準値

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [New-QualityReport.ps1](tools/README.md#new-qualityreportps1) | レビュー・テスト・工程・網羅・整合性の各ツールをまとめて実行し、工程完了判定に使う品質評価のブック（指標・実績・判定 ○△×・要対応）を作る | 不要 ※1 | 品質評価 |
| [Invoke-QualityAnalysis.ps1](tools/README.md#invoke-qualityanalysisps1) | 品質評価のブックの × と △ の指標について、定性分析（原因）と対策の下書きを Microsoft 365 Copilot に書かせ、ブックに書き込む | 不要 | 品質評価の定性分析 |
| [New-QualityLimits.ps1](tools/README.md#new-qualitylimitsps1) | 過去のプロジェクトの品質の実績から、社内の基準値（下限・上限）の CSV を作る | 不要 | － |

### 導入

最初に 1 回だけ使います。チームに配るときは Install-ExcelTools を使います。

| ツール | できること | Excel | 「送る」 |
| --- | --- | --- | --- |
| [Install-ExcelTools.ps1](tools/README.md#install-exceltoolsps1) | 共有フォルダに置いたツールを、担当者の PC に 1 回のダブルクリック（`Install-ExcelTools.cmd`）で入れ、「送る」に登録する。共有フォルダを更新すると、担当者の PC も自動で新しい版になる | 不要 | － |
| [Install-SendTo.ps1](tools/README.md#install-sendtops1) | 右クリックの「送る」に各ツールを登録し、PowerShell を使わずにファイルを選ぶだけで実行できるようにする | 不要 ※3 | － |
| [Start-ExcelTools.ps1](tools/README.md#start-exceltoolsps1画面) | ツールの画面。やりたいことを選び、ファイルを入れて「実行」を押すだけで各ツールを使える（`Start-ExcelTools.cmd` をダブルクリック） | 不要 | － |

※1 .xls / .xlsb を扱う場合のみ Excel が必要です。
※2 PDF のページ数を数えるだけ（`-CountOnly`）なら不要です。
※3 「PDF にする」だけは Excel が必要です。
※4 版を Excel で取り出します。Excel が無いか、Excel で開けないときは、ブラウザーでダウンロードします。

使い方の詳細は [tools/README.md](tools/README.md) を見てください。`tools` フォルダは**まとめて**配置してください（共通部品の `ExcelTools.psm1` を各スクリプトが使います）。

```powershell
.\tools\Search-Excel.ps1 'TODO' D:\Docs
.\tools\Test-ExcelDeliverable.ps1 D:\提出物 -OutXlsx D:\提出前チェック.xlsx
```

実行ポリシーで止まる場合は `powershell -ExecutionPolicy Bypass -File .\tools\Search-Excel.ps1 ...` のように実行してください。

**チームに配るとき（管理者）**は、`tools` フォルダを共有フォルダに置き、担当者にその中の `Install-ExcelTools.cmd` をダブルクリックしてもらいます。ツールが担当者の PC に入り、スタートメニューの「Excel ツール」（画面）と「送る」に登録されます。共有フォルダの `tools` を入れ替えれば、担当者の PC も自動で新しい版になります（[Install-ExcelTools.ps1](tools/README.md#install-exceltoolsps1)）。

自分だけで使うときは、`.\tools\Install-SendTo.ps1` を 1 回実行すると、エクスプローラーでファイルやフォルダを選んで右クリック →「送る」→「Excel 提出前チェック」などから使えます（結果の Excel が開きます）。

## フォルダ構成

| フォルダ | 内容 |
| --- | --- |
| `tools/` | ツール本体と共通部品（`ExcelTools.psm1`）、設計レビューの観点表と Copilot プロンプト集（`review-checklists/`） |
| `tests/` | ツールのテスト（[tests/README.md](tests/README.md)）。GitHub Actions で Windows PowerShell 5.1・PowerShell 7 の両方で実行 |
| `docs/工程別ガイド.md` | 設計・レビュー・提出・見積もり・テスト・改修・品質会議の工程ごとに、使うツールと順番 |
| `docs/design-sample/` | 設計ツールの検証用の詳細設計サンプル（架空の受注管理システム） |
| `docs/schedule-recovery/` | スケジュール遅延リカバリ計画のテンプレートと記入例 |
| `vba-files/`・`Book1.xlsm` | Excel マクロ（XVBA で管理）。「体裁整える」マクロは `Reset-ExcelView.ps1` で置き換え可能 |

## 動作環境

- Windows PowerShell 5.1 以降（PowerShell 7 でも動作）
- Excel を使う機能はデスクトップ版 Excel が必要
