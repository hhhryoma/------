# tools

| スクリプト | 内容 |
| --- | --- |
| [Search-Excel.ps1](#search-excelps1) | フォルダ配下の Excel を横断して文字列を検索する（Excel 版 grep） |
| [Export-ExcelToPdf.ps1](#export-exceltopdfps1) | フォルダ配下の Excel を PDF に一括出力し、ページ数を数える |

## Search-Excel.ps1

フォルダ配下の Excel ファイルをまとめて検索し、ヒットした **ファイル・シート・セル番地・内容** を一覧表示します。

検索対象は **セルの値 / 数式 / コメント（メモ） / 図形・テキストボックスの文字 / シート名** です。

### 要件

- PowerShell 5.1 以降（PowerShell 7 でも動作）
- **.xlsx / .xlsm / .xltx / .xltm は Excel 不要。** ファイル内部の XML を直接読むため、Excel を起動せず高速に動作します
- .xls / .xlsb を検索する場合のみデスクトップ版 Excel が必要です（見つかったときだけ Excel を起動して読みます）

### 使い方

```powershell
# カレントフォルダ配下（サブフォルダ含む）を検索
.\Search-Excel.ps1 '顧客ID'

# フォルダを指定
.\Search-Excel.ps1 '顧客ID' D:\Docs

# 正規表現（既定）。大文字小文字は区別しない
.\Search-Excel.ps1 'TODO|要確認' D:\Docs

# 記号を含む文字列をそのまま探す（正規表現として解釈しない）
.\Search-Excel.ps1 '(株)山田' D:\Docs -SimpleMatch

# 全角・半角を区別しない（「ＡＢＣ」と「ABC」、「ｶﾅ」と「カナ」が一致）
.\Search-Excel.ps1 'ABC' D:\Docs -IgnoreWidth

# 数式だけを検索して CSV に保存
.\Search-Excel.ps1 'VLOOKUP' D:\Docs -Target Formula -OutCsv D:\vlookup.csv

# 結果をウィンドウで絞り込みながら見る
.\Search-Excel.ps1 'TODO' D:\Docs | Out-GridView

# 結果一覧から選んだセルを Excel で開く
.\Search-Excel.ps1 'TODO' D:\Docs -Open
```

実行ポリシーで止まる場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\Search-Excel.ps1 'TODO' D:\Docs
```

### 実行結果の表示

```
File                      Location            Kind     Text
----                      --------            ----     ----
顧客管理\顧客一覧.xlsx    顧客一覧!B4         値       複数行↵の文字列 TODO
顧客管理\顧客一覧.xlsx    顧客一覧!E5         数式     =VLOOKUP(A2,顧客一覧!A:B,2,FALSE)
顧客管理\顧客一覧.xlsx    顧客一覧!A2         コメント TODO: 住所が古い
顧客管理\顧客一覧.xlsx    顧客一覧!G3         図形     注意：TODO 図形内↵2行目
設計書.xlsm               作業用              シート名 作業用TODO

2 ファイル中 2 ファイルで 5 件ヒット（0.4 秒）
```

- `Location` は `シート名!セル番地`。図形は図形の左上が乗っているセル、シート名のヒットはシート名のみ
- セル内の改行は `↵` で表示します
- 長い内容は画面幅で切れます。全文を見るには `| Format-Table -Wrap`、`| Out-GridView`、または `-OutCsv` を使ってください

結果は PowerShell のオブジェクトとして出力されるので、そのまま加工できます。

```powershell
$hits = .\Search-Excel.ps1 'TODO' D:\Docs
$hits | Where-Object Kind -eq '値' | Select-Object File, Sheet, Cell, Text
$hits | Group-Object File | Sort-Object Count -Descending   # ファイルごとの件数
```

| プロパティ | 内容 |
| --- | --- |
| `File` | 検索フォルダからの相対パス |
| `Location` | `シート名!セル番地` |
| `Sheet` / `Cell` | シート名 / セル番地（個別） |
| `Kind` | `値` / `数式` / `コメント` / `図形` / `シート名` |
| `Text` | ヒットした内容（全文） |
| `Name` | コメントの作成者、または図形の名前 |
| `Path` | ファイルのフルパス |

`-OutCsv` の CSV は UTF-8（BOM 付き）で、Excel でそのまま開けます。列名は日本語（ファイル / シート / セル / 種類 / 内容 / 名前 / フルパス）です。

### 該当セルを Excel で開く

`-Open` を付けると、検索後に結果が一覧ウィンドウで表示されます。行を選んで **OK** を押すと、そのファイルを Excel で開き、該当セルへジャンプして選択します。

```powershell
.\Search-Excel.ps1 'TODO' D:\Docs -Open
```

- ウィンドウ上部のフィルター欄に文字を入れると、結果をさらに絞り込めます
- Ctrl / Shift で複数行を選ぶと、すべて開きます（最後に選んだ行が前面になります）
- 図形のヒットは、図形の左上のセルへスクロールしてから図形を選択します。シート名のヒットはそのシートを表示します
- **起動中の Excel があればそれを使い**、既に開いているブックは開き直さずに切り替えます
- 読み取り専用ではなく、ダブルクリックで開いたときと同じく通常どおり開きます。マクロの扱いもセキュリティセンターの設定に従います
- 非表示シートのヒットは、ブックだけを開いて警告を出します（非表示のシートは選択できないため）
- Windows 専用です（`Out-GridView` を使うため）

### 主なパラメーター

| パラメーター | 説明 |
| --- | --- |
| `-Pattern` | 検索文字列（必須・1 番目の引数）。既定は .NET 正規表現 |
| `-Path` | 検索するフォルダまたはファイル（2 番目の引数、複数・ワイルドカード可）。省略時はカレントフォルダ |
| `-SimpleMatch` | 正規表現ではなく文字列そのままで検索 |
| `-CaseSensitive` | 大文字小文字を区別する |
| `-IgnoreWidth` | 全角半角を区別しない |
| `-Target` | 検索対象を限定。`Value` `Formula` `Comment` `Shape` `SheetName` から複数指定（既定はすべて） |
| `-NoRecurse` | サブフォルダを検索しない |
| `-SkipHiddenSheets` | 非表示シートを検索しない（既定では非表示シートも検索します） |
| `-SheetNamePattern` | 検索するシート名の正規表現 |
| `-SheetNameExclude` | 除外するシート名の正規表現（`-SheetNamePattern` より優先） |
| `-NoExcel` | Excel を起動しない（.xls / .xlsb は対象外になる） |
| `-Open` | 結果を一覧ウィンドウに表示し、選んだ行のセルを Excel で開く |
| `-OutCsv` | 結果を CSV に保存 |
| `-Extensions` | 対象拡張子（既定 `.xlsx .xlsm .xltx .xltm .xls .xlsb`） |

### 値の検索について

セルの「値」は、画面表示ではなくファイルに保存されている値で検索します。

| セルの内容 | 検索される文字列 |
| --- | --- |
| 文字列 | そのまま（ふりがなは含まない） |
| 数値 | 有効桁 15 桁の表記。`0.1+0.2` の結果は `0.3`、`1234567` は `1234567`（桁区切りや `¥` などの表示形式は付かない） |
| 日付・時刻（日付の表示形式のセル） | `2026/10/16`、`13:00:00`、`2026/01/02 09:30:00` の形式。和暦表示のセルも西暦で検索 |
| 論理値 | `TRUE` / `FALSE` |
| 数式のセル | 値（最後に計算された結果）は「値」、式は「数式」として別々に検索 |

### 設計上の考慮点

- xlsx 系はファイルを ZIP として開き XML を逐次読み込みするため、Excel を起動せず、大きなシートでもメモリを圧迫しません
- **Excel で開いているファイルも検索できます**（共有モードで読み取り、ファイルは変更しません）
- 拡張子ではなくファイルの中身で形式を判定します。中身が xlsx の「.xls」は Excel なしで読みます
- .xls / .xlsb 用に Excel を起動する場合は、マクロを強制無効化し読み取り専用で開きます。自分が起動した Excel プロセスだけを後始末します
- Excel の一時ファイル（`~$` で始まるファイル）は対象外
- 1 ファイル読めなくても全体は止めず、最後に読めなかったファイルと理由をまとめて表示します

### 既知の制約・注意

- **パスワード付きのファイルは検索できません**（「読めなかったファイル」に表示されます）
- 共有数式（数式をコピーして入力した範囲）は、xlsx では先頭セルにしか式が保存されていないため、数式の検索は先頭セルだけがヒットします。値の検索は全セル対象です
- グラフ内の文字、ヘッダー／フッター、フォームコントロール（ActiveX）、名前定義、入力規則のリストは検索しません
- .xls / .xlsb（Excel 経由）で読む場合、数式は英語版の関数名で検索されます。xlsx は保存されたとおり（英語の関数名）です
- 行数の多いシートでは 1 ファイル数秒かかることがあります（目安: 6 万セルで約 2 秒）

## Export-ExcelToPdf.ps1

フォルダ配下の Excel ファイル（.xls / .xlsx / .xlsm / .xlsb）を PDF に一括出力します。
処理後に、ファイルごとの出力先PDFと**ページ数**の一覧を表示します。

### 要件

- Windows + **デスクトップ版 Excel**（COM 自動化を使用。Microsoft 365 のクラウド版のみでは動作しません）
- PowerShell 5.1 以降（PowerShell 7 でも動作）
- **対話的なデスクトップセッション**での実行

### 使い方

```powershell
# 元ファイルと同じ場所に PDF を出力
.\Export-ExcelToPdf.ps1 -Path D:\Docs

# 出力先を分け、フォルダ構成を再現し、結果をCSVログに残す
.\Export-ExcelToPdf.ps1 -Path D:\Docs -OutputRoot D:\Pdf -LogPath D:\Pdf\result.csv

# シート単位で出力し、横1ページに収める
.\Export-ExcelToPdf.ps1 -Path D:\Docs -PerSheet -FitToWidth

# 実行せずに対象だけ確認
.\Export-ExcelToPdf.ps1 -Path D:\Docs -WhatIf
```

実行ポリシーで止まる場合:

```powershell
powershell -ExecutionPolicy Bypass -File .\Export-ExcelToPdf.ps1 -Path D:\Docs
```

### 主なパラメーター

| パラメーター | 説明 |
| --- | --- |
| `-Path` | 走査するフォルダ（必須） |
| `-OutputRoot` | 出力先ルート。省略時は元ファイルと同じ場所 |
| `-NoRecurse` | サブフォルダを走査しない |
| `-PerSheet` | シート単位で出力（`ファイル名_シート名.pdf`） |
| `-SkipExisting` | 既存 PDF が元ファイルより新しければスキップ（差分実行） |
| `-FitToWidth` | 全シートを横1ページに収める（処理は遅くなる） |
| `-IncludeHiddenSheets` | 非表示シートも出力（`-PerSheet` 時のみ） |
| `-SheetNamePattern` | 対象とするシート名の正規表現（一致したシートだけを出力・カウント） |
| `-SheetNameExclude` | 除外するシート名の正規表現（`-SheetNamePattern` より優先） |
| `-CountOnly` | PDF を出力せず、既存 PDF のページ数を数えるだけ（Excel を起動しない） |
| `-IgnoreSmallPages` | 用紙サイズが極端に小さいページをページ数の集計から除外 |
| `-SmallPageRatio` | 小ページ判定のしきい値。最大ページ面積に対する比率（既定 0.2） |
| `-OpenMode` | `Workbooks.Open` の呼び出し形式（`Auto` / `Full` / `Simple`）。既定は自動判定 |
| `-LogPath` | 結果（ページ数を含む）を CSV（UTF-8 BOM 付き）で保存 |
| `-WhatIf` | 実行せず対象を表示 |

### シート名で対象を絞る

`-SheetNamePattern` / `-SheetNameExclude` に .NET 正規表現を渡すと、一致するシートだけを出力・カウントします。**大文字小文字は区別しません。**

```powershell
# 「画面」で始まるシートだけ
.\Export-ExcelToPdf.ps1 -Path D:\Docs -SheetNamePattern '^画面'

# 表紙・改訂履歴・目次を除いたページ数を数える
.\Export-ExcelToPdf.ps1 -Path D:\Docs -SheetNameExclude '表紙|改訂履歴|目次'

# 併用（除外が優先）
.\Export-ExcelToPdf.ps1 -Path D:\Docs -SheetNamePattern '^(画面|帳票)' -SheetNameExclude '作業用'
```

| シート名の例 | `-SheetNamePattern '^(画面\|帳票)'` | `-SheetNameExclude '表紙\|改訂履歴\|目次'` |
| --- | --- | --- |
| 表紙 | 除外 | 除外 |
| 改訂履歴 | 除外 | 除外 |
| 画面遷移図 | **対象** | **対象** |
| 帳票一覧 | **対象** | **対象** |
| Sheet1 | 除外 | **対象** |

動作の仕組み:

- **シート単位出力（`-PerSheet`）** … 一致しないシートを単純にスキップします
- **ブック単位出力（既定）** … 一致しないシートを**一時的に非表示**にしてから出力します。非表示シートは `ExportAsFixedFormat` の対象外になるため、これだけで出力範囲を限定できます。読み取り専用で開いており保存しないので、**元の Excel ファイルは変更されません**
- 1 枚も一致しないブックは、そのファイルごとスキップして「条件に一致する印刷対象シートなし」と記録します（Excel は全シートを非表示にできないため、非表示化は行いません）
- 正規表現は**実行前に検証**します。不正なパターンは処理を始める前にエラーで止まります
- **`-CountOnly` では使えません**（PDF にシート名の情報が残らないため）。指定すると警告が出て無視されます

### ページ数を数えるだけのモード

`-CountOnly` を付けると、**PDF を出力せず**、`-Path` 配下にある既存の PDF のページ数を数えるだけになります。

```powershell
.\Export-ExcelToPdf.ps1 -Path D:\Pdf -CountOnly
.\Export-ExcelToPdf.ps1 -Path D:\Pdf -CountOnly -IgnoreSmallPages -LogPath D:\count.csv
```

- 対象は Excel ファイルではなく **`*.pdf`** です（サブフォルダも走査、`-NoRecurse` で抑止）
- **Excel を一切起動しません。** Excel が入っていない端末でも動作し、大量のファイルでも高速です
- `-IgnoreSmallPages` / `-SmallPageRatio` / `-LogPath` / `-NoRecurse` はそのまま使えます
- `-OutputRoot` `-PerSheet` `-FitToWidth` `-SkipExisting` `-IncludeHiddenSheets` は無視され、指定すると警告が出ます

```
対象 12 PDF / 走査元 D:\Pdf

==== ページ数一覧 ========================================================

ファイル                       ページ 結果 詳細
--------                       ------ ---- ----
画面遷移図\BA-03-02-0001.pdf        4 成功
画面遷移図\BA-03-04-0001.pdf        2 成功

成功 12 / スキップ 0 / 失敗 0
合計ページ数 48
```

### 実行結果の表示

処理の最後に一覧とサマリを出力します。

```
==== 出力ファイル一覧 ====================================================

ファイル                  出力PDF                 ページ 結果       詳細
--------                  -------                 ------ ----       ----
売上報告_202608.xlsx      売上報告_202608.pdf          3 成功
見積書\A社_見積.xlsx      A社_見積.pdf                 1 成功
集計表.xlsx               集計表.pdf                   - スキップ   印刷対象なし（空ブック／全シート非表示）
旧台帳.xls                -                            - 失敗       パスワードが正しくありません

成功 2 / スキップ 1 / 失敗 1
合計ページ数 4
```

`-LogPath` を指定すると、同じ内容（元ファイルのフルパス・出力先・ページ数・結果・詳細・所要秒）が CSV に保存されます。

#### 小さいページを集計から除外する

シートごとに用紙設定が異なるブックでは、意図しない極端に小さいページが混ざることがあります。
`-IgnoreSmallPages` を付けると、そうしたページをページ数の集計から外します。

```powershell
.\Export-ExcelToPdf.ps1 -Path D:\Docs -IgnoreSmallPages -Verbose
```

- PDF の `/MediaBox`（ページの寸法）を読み、**そのPDF内で最大のページ面積の 20% 未満**のページを除外します
- しきい値は `-SmallPageRatio 0.3` のように変更できます（0.01〜1.0）
- ページ個別に `/MediaBox` が無い場合は、ページツリー側の値を継承して判定します
- **PDF 自体は変更しません。** 集計から外すだけなので、出力された PDF には該当ページが残ります
- `-Verbose` を付けると、除外されたページの番号と寸法（mm）が表示されます。しきい値の調整に使ってください

一覧には除外列が追加されます。

```
ファイル              出力PDF             ページ 除外 結果   詳細
--------              -------             ------ ---- ----   ----
伝票フォーム.xlsx     伝票フォーム.pdf         4    2 成功

成功 1 / スキップ 0 / 失敗 0
合計ページ数 4
小サイズとして除外したページ 2（最大ページ面積の 20% 未満）
```

なお、PDF がオブジェクトストリームで圧縮されていてページ個別の解析ができない場合は、
サイズ判定が適用できないため総ページ数をそのまま返します（`-Verbose` に記録されます）。

#### ページ数の取得方法

1. 生成された PDF のバイト列から `/Type /Page` の出現数を数えます（外部ツール不要）
2. 取れない場合のみ、Excel の `PageSetup.Pages.Count` にページ割りを計算させます（低速なためフォールバック扱い）
3. どちらでも取得できなければ `-` と表示します

`-SkipExisting` でスキップしたファイルも、既存 PDF を読んでページ数を表示します。
なお 2. のフォールバックはプリンタードライバーが必要で、`-FitToWidth` と同様に処理時間が伸びます。

### 設計上の考慮点

- マクロは `AutomationSecurity = ForceDisable` で強制無効化。xlsm を安全に処理できる
- 元ファイルは読み取り専用で開き、保存しない
- ユーザーが開いている Excel は操作せず専用インスタンスを使用。起動前後の PID 差分を取り、**自分が起動したプロセスだけ**を後始末する
- パスワード保護ファイルはダミーパスワードを渡すことで、入力待ちで停止せずエラーとして記録される
- 空ブック・全シート非表示のブックは事前判定してスキップ（`ExportAsFixedFormat` の例外を回避）
- `-FitToWidth` 時は `PrintCommunication = $false` で PageSetup 変更を高速化
- 1ファイルの失敗で全体を止めず、結果を CSV に残す

### 既知の制約・注意

- **タスクスケジューラで「ユーザーがログオンしているかどうかに関わらず実行する」にすると失敗します。** Excel COM は対話的セッションを前提とするためです。回避する場合は `C:\Windows\System32\config\systemprofile\Desktop`（32bit Excel なら `SysWOW64` 側）を作成する必要がありますが、Microsoft 非推奨の構成です。定期実行するならログオンユーザーのタスクとして動かしてください
- **プリンタードライバーが1つも無い環境では PageSetup 操作が失敗することがあります。** Microsoft Print to PDF を有効にしておくと安定します
- 破損ファイル・非常に大きいブックで Excel が応答しなくなった場合、本スクリプトは待ち続けます。タイムアウトが必要なら1ファイルずつ別プロセス（`Start-Job` 等）で処理してください
- 260文字を超える長いパスは Excel 側が失敗することがあります
- 1ファイルあたり数秒かかります。数千ファイルでは数時間規模になるため `-SkipExisting` での差分実行を前提にしてください

### トラブルシューティング

#### 「Workbooks クラスの Open プロパティを取得できません」

Excel が `Workbooks.Open` の呼び出しを拒否したときの汎用エラーです。原因は主に次の5つ。

1. **省略可能引数が通らない** — PowerShell 7 系や一部の Excel バージョンでは `[Type]::Missing` での引数省略が拒否されます。本スクリプトは起動時に自作の空ブックで呼び出し形式を判定し、通らない環境では自動的に簡易形式へ切り替えます。それでも出る場合は `-OpenMode Simple` を明示してください
2. **OneDrive / SharePoint 同期フォルダ** — Excel がクラウド URL で開こうとして失敗します。「ファイルオンデマンド」でダウンロードされていない（雲アイコンの）ファイルも同様。ローカルにコピーしてから実行してください
3. **パスが 218 文字超** — Excel の制限です（OS の 260 文字より手前）。本スクリプトは事前に検出してメッセージを出します
4. **拡張子と実体の不一致** — 基幹システムが出力した「.xls」が実際は HTML や CSV というケース。Excel で手動で開くと警告が出ます
5. **Excel が応答待ち状態** — 前回の異常終了後の「ドキュメントの回復」ウィンドウやアドインのエラーダイアログが裏で開いていると、以降すべての Open が失敗します。`EXCEL.EXE` を全て終了してから再実行してください

切り分けには、呼び出し形式を変えながら同じファイルを開いてみるのが確実です。

```powershell
$f = 'D:\Docs\失敗したファイル.xlsx'
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
foreach ($n in 1,3,7,15) {
    try {
        switch ($n) {
            1  { $wb = $xl.Workbooks.Open($f) }
            3  { $wb = $xl.Workbooks.Open($f,0,$true) }
            7  { $wb = $xl.Workbooks.Open($f,0,$true,[Type]::Missing,'x','x',$true) }
            15 { $wb = $xl.Workbooks.Open($f,0,$true,[Type]::Missing,'x','x',$true,[Type]::Missing,[Type]::Missing,$false,$false,[Type]::Missing,$false,$true,[Type]::Missing) }
        }
        "$n 引数: OK"; $wb.Close($false)
    } catch { "$n 引数: NG  $($_.Exception.Message)" }
}
$xl.Quit()
```

`15 引数だけ NG` なら原因 1、`全部 NG` なら原因 2〜5 です。

### Excel が無い環境の代替

LibreOffice がある環境なら、COM を使わずに変換できます（書式の再現度は Excel に劣ります）。

```powershell
Get-ChildItem D:\Docs -Recurse -Include *.xlsx,*.xls,*.xlsm |
  ForEach-Object {
    & "C:\Program Files\LibreOffice\program\soffice.com" `
      --headless --convert-to pdf --outdir D:\Pdf $_.FullName
  }
```
