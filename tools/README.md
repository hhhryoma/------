# Export-ExcelToPdf.ps1

フォルダ配下の Excel ファイル（.xls / .xlsx / .xlsm / .xlsb）を PDF に一括出力します。
処理後に、ファイルごとの出力先PDFと**ページ数**の一覧を表示します。

## 要件

- Windows + **デスクトップ版 Excel**（COM 自動化を使用。Microsoft 365 のクラウド版のみでは動作しません）
- PowerShell 5.1 以降（PowerShell 7 でも動作）
- **対話的なデスクトップセッション**での実行

## 使い方

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

## 主なパラメーター

| パラメーター | 説明 |
| --- | --- |
| `-Path` | 走査するフォルダ（必須） |
| `-OutputRoot` | 出力先ルート。省略時は元ファイルと同じ場所 |
| `-NoRecurse` | サブフォルダを走査しない |
| `-PerSheet` | シート単位で出力（`ファイル名_シート名.pdf`） |
| `-SkipExisting` | 既存 PDF が元ファイルより新しければスキップ（差分実行） |
| `-FitToWidth` | 全シートを横1ページに収める（処理は遅くなる） |
| `-IncludeHiddenSheets` | 非表示シートも出力（`-PerSheet` 時のみ） |
| `-LogPath` | 結果（ページ数を含む）を CSV（UTF-8 BOM 付き）で保存 |
| `-WhatIf` | 実行せず対象を表示 |

## 実行結果の表示

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

### ページ数の取得方法

1. 生成された PDF のバイト列から `/Type /Page` の出現数を数えます（外部ツール不要）
2. 取れない場合のみ、Excel の `PageSetup.Pages.Count` にページ割りを計算させます（低速なためフォールバック扱い）
3. どちらでも取得できなければ `-` と表示します

`-SkipExisting` でスキップしたファイルも、既存 PDF を読んでページ数を表示します。
なお 2. のフォールバックはプリンタードライバーが必要で、`-FitToWidth` と同様に処理時間が伸びます。

## 設計上の考慮点

- マクロは `AutomationSecurity = ForceDisable` で強制無効化。xlsm を安全に処理できる
- 元ファイルは読み取り専用で開き、保存しない
- ユーザーが開いている Excel は操作せず専用インスタンスを使用。起動前後の PID 差分を取り、**自分が起動したプロセスだけ**を後始末する
- パスワード保護ファイルはダミーパスワードを渡すことで、入力待ちで停止せずエラーとして記録される
- 空ブック・全シート非表示のブックは事前判定してスキップ（`ExportAsFixedFormat` の例外を回避）
- `-FitToWidth` 時は `PrintCommunication = $false` で PageSetup 変更を高速化
- 1ファイルの失敗で全体を止めず、結果を CSV に残す

## 既知の制約・注意

- **タスクスケジューラで「ユーザーがログオンしているかどうかに関わらず実行する」にすると失敗します。** Excel COM は対話的セッションを前提とするためです。回避する場合は `C:\Windows\System32\config\systemprofile\Desktop`（32bit Excel なら `SysWOW64` 側）を作成する必要がありますが、Microsoft 非推奨の構成です。定期実行するならログオンユーザーのタスクとして動かしてください
- **プリンタードライバーが1つも無い環境では PageSetup 操作が失敗することがあります。** Microsoft Print to PDF を有効にしておくと安定します
- 破損ファイル・非常に大きいブックで Excel が応答しなくなった場合、本スクリプトは待ち続けます。タイムアウトが必要なら1ファイルずつ別プロセス（`Start-Job` 等）で処理してください
- 260文字を超える長いパスは Excel 側が失敗することがあります
- 1ファイルあたり数秒かかります。数千ファイルでは数時間規模になるため `-SkipExisting` での差分実行を前提にしてください

## Excel が無い環境の代替

LibreOffice がある環境なら、COM を使わずに変換できます（書式の再現度は Excel に劣ります）。

```powershell
Get-ChildItem D:\Docs -Recurse -Include *.xlsx,*.xls,*.xlsm |
  ForEach-Object {
    & "C:\Program Files\LibreOffice\program\soffice.com" `
      --headless --convert-to pdf --outdir D:\Pdf $_.FullName
  }
```
