#     Excel ツールの画面（Start-ExcelTools.ps1）に並べる処理の一覧。
#     画面の入力欄は、各ツールの引数とヘルプから作る。ここには、画面の名前・分類・説明と、最初から見せる入力欄（Main）の順番・名前だけを書く。
#     Main に無い引数は「詳細な設定」に出る。
#
#       Id           処理の名前（設定の保存・-Tool での指定に使う）
#       Label        画面に出す名前（やりたいこと）
#       Group        分類（Groups の Name）
#       Tool         ツールのファイル名
#       Set          引数のセット（ツールに複数あるとき）
#       Summary      説明
#       Main         最初から見せる入力欄（引数の名前。この順に並べる）
#       Labels       入力欄の名前（省略時はヘルプの最初の文）
#       Kinds        入力欄の種類（省略時は引数の型とヘルプから決める）。Text・File・Folder・Path・SaveFile・SaveFolder・Number・Date・Switch・Choice
#       Required     画面では必ず入れてもらう引数（ツールでは省略できても、既定がカレントフォルダなどで画面に合わないもの）
#       Hide         画面に出さない引数
#       Defaults     入力欄の最初の値
#       Fixed        いつも付ける引数
#       Verbatim     入力のまま渡す文字の引数（探す文字・置き換える文字。前後の空白・引用符を除かず、空白だけの入力も渡す）
#       AllowEmpty   空欄なら空の文字を渡す引数（置き換えた後の文字。空欄なら文字を消す）
#       Output       空欄なら「ドキュメント\Excel ツールの結果\OutputName_日時.拡張子」を入れる引数（結果のファイル）
#       Open         終わったら開く引数（ファイルかフォルダ。先に書いたものから、あるものを開く）。省略時は Output
#       WhatIf       「試しに実行（ファイルを書き換えない）」を出す（最初は選んだ状態）
#       Confirm      実際に書き換える前に出す確認の文。ConfirmWhen があれば、そのスイッチを選んだときだけ
#       ConfirmArguments  確認で「はい」を選んだときに付ける引数（ツールの確認を省くため）
#       Console      PowerShell の画面で実行する（Copilot とのやり取りなど、途中で入力が要るもの）
#       Keywords     探すときの言葉

@{
    Groups = @(
        @{ Name = '設計書を作る・直す'; For = '担当者' }
        @{ Name = '変わったところを見る'; For = '担当者' }
        @{ Name = 'レビュー・提出'; For = '担当者' }
        @{ Name = 'テスト'; For = '担当者' }
        @{ Name = '見積もり・進捗'; For = '管理者' }
        @{ Name = '品質の集計・分析'; For = '管理者' }
        @{ Name = '品質会議・基準値'; For = '管理者' }
    )

    Tools = @(
        # ---------------------------------------------------------------- 設計書を作る・直す
        @{
            Id = 'Search'; Label = '文字を探す（検索）'; Group = '設計書を作る・直す'; Tool = 'Search-Excel.ps1'
            Summary = 'フォルダの中の Excel・Word を横断して、セルの値・数式・コメント・図形などから文字を探します。結果の Excel の場所をクリックすると、そのセルが開きます。'
            Main = @('Pattern', 'Path', 'SimpleMatch', 'IncludeWord', 'OutXlsx')
            Labels = @{ Pattern = '探す文字'; Path = '探すファイル・フォルダ'; SimpleMatch = '入力した文字そのものを探す（正規表現として扱わない）'; IncludeWord = 'Word の文書も探す'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Hide = @('Open'); Defaults = @{ SimpleMatch = $true; IncludeWord = $true }; Verbatim = @('Pattern')
            Output = @('OutXlsx'); OutputName = '検索'; Keywords = 'grep 検索 探す'
        }
        @{
            Id = 'Replace'; Label = '文字を置き換える（置換）'; Group = '設計書を作る・直す'; Tool = 'Replace-Excel.ps1'
            Summary = '複数の Excel の文字をまとめて置き換えます（値・コメント・図形・ヘッダー／フッターなど）。最初は「試しに実行」で、置き換える箇所を確かめてください。'
            Main = @('Pattern', 'Replacement', 'Path', 'SimpleMatch', 'ListPath', 'BackupDir', 'OutXlsx')
            Labels = @{ Pattern = '置き換える文字'; Replacement = '置き換えた後の文字（空欄なら消す）'; Path = '対象のファイル・フォルダ'; SimpleMatch = '入力した文字そのものを置き換える（正規表現として扱わない）'; ListPath = '置換リスト（CSV。まとめて置き換えるとき）'; BackupDir = '書き換える前のファイルの保存先'; OutXlsx = '置き換えた箇所の一覧の保存先' }
            Required = @('Path'); Hide = @('Force'); Defaults = @{ SimpleMatch = $true }; Verbatim = @('Pattern', 'Replacement'); AllowEmpty = @('Replacement')
            Output = @('OutXlsx'); OutputName = '置換'
            WhatIf = $true; Confirm = 'ファイルの文字を置き換えます（元に戻せません）。よろしいですか？'; ConfirmArguments = @{ Force = $true }
            Keywords = '置換 一括'
        }
        @{
            Id = 'Terms'; Label = '表記ゆれを見つける'; Group = '設計書を作る・直す'; Tool = 'Test-ExcelTerms.ps1'
            Summary = 'ユーザ／ユーザー、ＩＤ／ID のような表記ゆれを、用語集と自動の検出で見つけます。置換リスト（CSV）を作ると、「文字を置き換える」でまとめて直せます。'
            Main = @('Path', 'Glossary', 'IncludeWord', 'OutXlsx', 'OutCsv')
            Labels = @{ Path = '調べるファイル・フォルダ'; Glossary = '用語集（CSV。無くてもよい）'; IncludeWord = 'Word の文書も調べる'; OutXlsx = '結果の保存先'; OutCsv = '置換リストの保存先（CSV）' }
            Required = @('Path'); Defaults = @{ IncludeWord = $true }
            Output = @('OutXlsx'); OutputName = '表記ゆれ'; Keywords = '用語 表記'
        }
        @{
            Id = 'Consistency'; Label = '設計書の整合性チェック'; Group = '設計書を作る・直す'; Tool = 'Test-DesignConsistency.ps1'
            Summary = '設計書どうしの参照の食い違い（定義されていないテーブル・カラム・メッセージ・コード・画面など）を見つけます。CRUD 図も作ります。影響調査では、ID を使っている場所の一覧を作ります。'
            Main = @('Path', 'Where', 'OutXlsx')
            Labels = @{ Path = '設計書のフォルダ'; Where = '影響調査する ID（カンマ区切り。空欄ならチェックだけ）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '整合性チェック'; Keywords = 'CRUD 影響調査 参照'
        }
        @{
            Id = 'ScreenNumbers'; Label = '画面の丸数字を確かめる'; Group = '設計書を作る・直す'; Tool = 'Test-ScreenNumbers.ps1'
            Summary = '画面レイアウトのスクリーンショットに重ねた丸数字（①・No.5 など）を、画面項目定義の No と突き合わせ、無い番号・重なった番号を見つけます。前の版を指定すると、項目を足して No がずれたのに直していない丸数字も見つけます。'
            Main = @('Path', 'Baseline', 'OutXlsx')
            Labels = @{ Path = '画面設計書のフォルダ'; Baseline = '前の版の画面設計書（番号のずれを見るとき）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '画面の番号'; Keywords = '丸数字 番号 画面項目 スクリーンショット ワイヤーフレーム'
        }
        @{
            Id = 'Screenshots'; Label = '古いスクリーンショットを見つける'; Group = '設計書を作る・直す'; Tool = 'Test-Screenshots.ps1'
            Summary = '設計書に貼った画面のスクリーンショットが、元のワイヤーフレーム（Excel）より古くなっていないかを確かめます。貼り替えたあとに「貼り替えた画面を記録する」を選んで実行すると、いまの状態を記録します（設計書は変更しません）。'
            Main = @('Path', 'Wireframe', 'Record', 'OutXlsx')
            Labels = @{ Path = '設計書のフォルダ'; Wireframe = '元のワイヤーフレームのフォルダ'; Record = '貼り替えた画面を記録する'; OutXlsx = '結果の保存先' }
            Kinds = @{ RecordFile = 'SaveFile' }; Extensions = @{ RecordFile = '.csv' }
            Required = @('Path', 'Wireframe'); Output = @('OutXlsx'); OutputName = 'スクリーンショットの状態'; Keywords = 'スクリーンショット 貼り替え ワイヤーフレーム 古い 記録'
        }
        @{
            Id = 'NewDoc'; Label = '雛形から設計書を作る'; Group = '設計書を作る・直す'; Tool = 'New-DesignDocument.ps1'; Set = 'Create'
            Summary = '一覧（Excel・CSV）の行ごとに、雛形の {{名前}} の欄を埋めた設計書を作ります。'
            Main = @('Template', 'List', 'FileName', 'OutDir')
            Labels = @{ Template = '雛形（xlsx・xlsm）'; List = '一覧（Excel・CSV）'; FileName = '作るファイルの名前（例: DD02_画面設計書_{{画面ID}}.xlsx）'; OutDir = '保存先のフォルダ' }
            Required = @('OutDir'); Open = @('OutDir'); WhatIf = $true; Keywords = '雛形 テンプレート 作成'
        }
        @{
            Id = 'AddSheet'; Label = '雛形のシートからシートを足す'; Group = '設計書を作る・直す'; Tool = 'New-DesignDocument.ps1'; Set = 'AddSheet'
            Summary = '一覧の行ごとに、ブックの雛形のシートをコピーしてシートを足します（1 件 1 シートの設計書）。ブックを上書きします。'
            Main = @('Book', 'TemplateSheet', 'List', 'SheetName')
            Labels = @{ Book = 'シートを足すブック'; TemplateSheet = '雛形のシートの名前'; List = '一覧（Excel・CSV）'; SheetName = '足すシートの名前（例: {{テーブルID}}）' }
            Open = @('Book'); WhatIf = $true; Confirm = 'ブックにシートを足して上書きします。よろしいですか？'; Keywords = '雛形 シート'
        }
        @{
            Id = 'Bump'; Label = '版を上げる'; Group = '設計書を作る・直す'; Tool = 'New-DesignDocument.ps1'; Set = 'Bump'
            Summary = '設計書の表紙の版数・更新日、改訂履歴の行、ファイル名の版数を書き換えます。前の版を指定すると、差分から改訂内容を書きます。'
            Main = @('Path', 'CompareWith', 'Description', 'Version')
            Labels = @{ Path = '版を上げる設計書'; CompareWith = '前の版（ファイル・フォルダ。差分から改訂内容を書く）'; Description = '改訂内容（前の版を指定しないとき）'; Version = '新しい版数（空欄なら 1 つ上げる）' }
            Fixed = @{ Bump = $true }; Hide = @('Bump'); WhatIf = $true; Confirm = '設計書の表紙・改訂履歴・ファイル名を書き換えます。よろしいですか？'; Keywords = '版数 改訂履歴'
        }
        @{
            Id = 'Ddl'; Label = 'テーブル定義書から DDL を作る'; Group = '設計書を作る・直す'; Tool = 'Export-TableDdl.ps1'
            Summary = 'テーブル定義書から CREATE TABLE の DDL を作ります。DB の列の一覧（CSV）と比べて、型・桁数・NULL の食い違いも確かめられます。'
            Main = @('Path', 'Database', 'OutFile', 'CompareCatalog', 'OutXlsx')
            Labels = @{ Path = 'テーブル定義書のフォルダ・ファイル'; Database = 'DB の種類'; OutFile = 'DDL の保存先'; CompareCatalog = 'DB の列の一覧（CSV。比べるとき）'; OutXlsx = 'チェック結果の保存先' }
            Kinds = @{ OutFile = 'SaveFile' }; Extensions = @{ OutFile = '.sql' }
            Required = @('Path'); Output = @('OutFile', 'OutXlsx'); OutputName = 'DDL'; Open = @('OutXlsx'); Keywords = 'SQL テーブル'
        }
        @{
            Id = 'Diagram'; Label = '画面遷移図・ER 図を作る'; Group = '設計書を作る・直す'; Tool = 'Export-DesignDiagram.ps1'
            Summary = 'イベント定義の遷移先から画面遷移図を、テーブル定義書の外部キーから ER 図を作ります。Excel の図形（四角と矢印）と Mermaid（Markdown）で保存し、遷移先・参照先の書き間違いも知らせます。'
            Main = @('Path', 'Kind', 'OutXlsx', 'OutMarkdown')
            Labels = @{ Path = '設計書のフォルダ'; Kind = '作る図（選ばなければ両方）'; OutXlsx = 'Excel の図の保存先'; OutMarkdown = 'Mermaid（Markdown）の保存先' }
            Extensions = @{ OutMarkdown = '.md' }
            Required = @('Path'); Output = @('OutXlsx', 'OutMarkdown'); OutputName = '画面遷移図とER図'; Open = @('OutXlsx'); Keywords = '画面遷移図 ER図 図 Mermaid 外部キー 遷移'
        }

        # ---------------------------------------------------------------- 変わったところを見る
        @{
            Id = 'Compare'; Label = '2 つを比べる（差分）'; Group = '変わったところを見る'; Tool = 'Compare-Excel.ps1'; Set = 'Pair'
            Summary = '2 つのブック・Word 文書（またはフォルダ）を比べ、差分をシートの方眼のまま色付けした画面（ブラウザー）を開きます。画面レイアウトのスクリーンショットの貼り替えや、重ねた図形の移動も出します。差分の一覧の Excel も保存します。'
            Main = @('Old', 'New', 'OutHtml', 'OutXlsx')
            Labels = @{ Old = '旧（変更前）のファイル・フォルダ'; New = '新（変更後）のファイル・フォルダ'; OutHtml = '差分の画面（HTML）の保存先'; OutXlsx = '差分の一覧の保存先' }
            Output = @('OutHtml', 'OutXlsx'); OutputName = '差分'; Keywords = '差分 比較 diff'
        }
        @{
            Id = 'CompareVersions'; Label = '版を続けて比べる'; Group = '変わったところを見る'; Tool = 'Compare-Excel.ps1'; Set = 'Versions'
            Summary = 'v1.0・v1.1・v1.2 のように版のあるファイル（またはフォルダ）を、版の順に続けて比べます。版数の付いたファイルを置いたフォルダを 1 つ選んでもかまいません。'
            Main = @('Versions', 'OutHtml', 'OutXlsx')
            Labels = @{ Versions = '版のファイル・フォルダ（3 つ以上か、版のファイルを置いたフォルダ 1 つ）'; OutHtml = '差分の画面（HTML）の保存先'; OutXlsx = '差分の一覧の保存先' }
            Kinds = @{ Versions = 'Path' }; Output = @('OutHtml', 'OutXlsx'); OutputName = '版の差分'; Keywords = '差分 比較 版'
        }
        @{
            Id = 'SpVersion'; Label = 'SharePoint の版と比べる'; Group = '変わったところを見る'; Tool = 'Compare-SharePointVersion.ps1'
            Summary = 'SharePoint（Microsoft 365）のファイルの版（バージョン履歴の 3.0 など）を取り出し、現在の版と比べて差分の画面を開きます。'
            Main = @('Path', 'Version', 'NewVersion')
            Labels = @{ Path = '比べるファイル（同期しているファイルか URL）'; Version = '旧の版の番号（例: 3.0）'; NewVersion = '新の版の番号（空欄なら現在の版）' }
            Kinds = @{ Path = 'File' }; Required = @('Version'); Keywords = 'SharePoint バージョン履歴 OneDrive'
        }
        @{
            Id = 'Markdown'; Label = 'Markdown（文字）にする'; Group = '変わったところを見る'; Tool = 'Export-ExcelToMarkdown.ps1'
            Summary = 'Excel を Markdown（表は Markdown の表）にします。AI に読ませたり、git で文字の差分を見たりするときに使います。'
            Main = @('Path', 'OutDir', 'CellAddress')
            Labels = @{ Path = 'Excel のファイル・フォルダ'; OutDir = 'Markdown の保存先のフォルダ'; CellAddress = 'セル番地を付ける（AI に場所を答えさせるとき）' }
            Required = @('Path', 'OutDir'); Hide = @('TextConv'); Open = @('OutDir'); Keywords = 'Markdown テキスト git'
        }

        # ---------------------------------------------------------------- レビュー・提出
        @{
            Id = 'Check'; Label = '提出前チェック'; Group = 'レビュー・提出'; Tool = 'Test-ExcelDeliverable.ps1'
            Summary = '体裁・消し忘れ・情報漏れ・印刷のはみ出し・版数の食い違い・スクリーンショットの縦横比や拡大を確かめます（Word は変更履歴・コメント・隠し文字の残りも）。「直す」を選ぶと、機械で直せる指摘（プロパティの個人名・表示の状態など）を直します。'
            Main = @('Path', 'Fix', 'BackupDir', 'OutXlsx')
            Labels = @{ Path = 'チェックするファイル・フォルダ'; Fix = '機械で直せる指摘を直す（ファイルを書き換える）'; BackupDir = '書き換える前のファイルの保存先'; OutXlsx = '結果の保存先' }
            Output = @('OutXlsx'); OutputName = '提出前チェック'
            Confirm = '機械で直せる指摘を直します（ファイルを書き換えます）。よろしいですか？'; ConfirmWhen = 'Fix'; Keywords = '提出 納品 チェック 個人情報'
        }
        @{
            Id = 'ResetView'; Label = '表示を整える（A1・倍率 100%）'; Group = 'レビュー・提出'; Tool = 'Reset-ExcelView.ps1'
            Summary = '開いたときに、先頭のシートの A1 が倍率 100% で表示される状態にまとめて直します（ファイルを書き換えます）。'
            Main = @('Path', 'Zoom', 'NormalView', 'BackupDir')
            Labels = @{ Path = '直すファイル・フォルダ'; Zoom = '倍率（%）'; NormalView = '改ページプレビューのシートを標準の表示にする'; BackupDir = '書き換える前のファイルの保存先' }
            WhatIf = $true; Confirm = 'ファイルの表示の状態を書き換えます。よろしいですか？'; Keywords = '体裁 A1 倍率'
        }
        @{
            Id = 'Review'; Label = 'AI（Copilot）にレビューさせる'; Group = 'レビュー・提出'; Tool = 'Invoke-DesignReview.ps1'
            Summary = '設計書（Excel・Word）を観点表に沿って Microsoft 365 Copilot にレビューさせ、指摘一覧にまとめます。前の版を入れると変わった部分だけを、前のレビュー結果を入れると指摘が直ったかを確かめます。PowerShell の画面が開くので、案内に沿って Copilot とやり取りしてください。'
            Main = @('Path', 'Checklist', 'OutDir', 'Baseline', 'Recheck')
            Labels = @{ Path = 'レビューする設計書'; Checklist = '観点表（空欄なら詳細設計の観点）'; OutDir = '結果の保存先のフォルダ'; Baseline = '前の版の設計書（変わった部分だけをレビューするとき）'; Recheck = '前のレビュー結果（指摘が直ったかを確かめるとき。保存先は別のフォルダに）' }
            Required = @('Path', 'OutDir'); Open = @('OutDir'); Console = $true; Keywords = 'AI Copilot レビュー 前の版 再確認 Word'
        }
        @{
            Id = 'Pdf'; Label = 'PDF にする'; Group = 'レビュー・提出'; Tool = 'Export-ExcelToPdf.ps1'
            Summary = 'フォルダの中の Excel を PDF にします（Excel が必要です）。保存先を空欄にすると、元のファイルと同じ場所に保存します。'
            Main = @('Path', 'OutputRoot', 'PerSheet', 'FitToWidth')
            Labels = @{ Path = 'Excel のフォルダ'; OutputRoot = 'PDF の保存先のフォルダ（空欄なら元と同じ場所）'; PerSheet = 'シートごとに PDF にする'; FitToWidth = '横 1 ページに収める' }
            Open = @('OutputRoot', 'Path'); Keywords = 'PDF 印刷'
        }
        @{
            Id = 'Register'; Label = '成果物一覧を作る'; Group = 'レビュー・提出'; Tool = 'New-DocumentRegister.ps1'
            Summary = 'フォルダの成果物の表紙と改訂履歴から、文書番号・文書名・版数・更新日・最新の改訂の一覧を作ります。古い版・版数の食い違いと、成果物の予定を入れると未提出も見つけます。'
            Main = @('Path', 'Planned', 'OutXlsx')
            Labels = @{ Path = '成果物のフォルダ'; Planned = '成果物の予定（Excel・CSV。突き合わせるとき）'; OutXlsx = '成果物一覧の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '成果物一覧'; Keywords = '納品物 納品 成果物 一覧 台帳 版数 未提出'
        }

        # ---------------------------------------------------------------- テスト
        @{
            Id = 'TestDraft'; Label = 'テスト仕様書の下書きを作る'; Group = 'テスト'; Tool = 'New-TestSpecDraft.ps1'
            Summary = '画面設計書の必須・桁数・型・入力チェック・コード値・イベントから、テスト仕様書の下書き（メッセージの文言付きの期待結果）を作ります。'
            Main = @('DesignPath', 'Screen', 'OutXlsx')
            Labels = @{ DesignPath = '画面設計書（とコード定義書・メッセージ一覧）のフォルダ'; Screen = '画面 ID（カンマ区切り。空欄ならすべて）'; OutXlsx = '下書きの保存先' }
            Output = @('OutXlsx'); OutputName = 'テスト仕様書の下書き'; Keywords = 'テストケース 単体テスト'
        }
        @{
            Id = 'TestSpec'; Label = 'テスト仕様書を確かめる'; Group = 'テスト'; Tool = 'Test-TestSpec.ps1'
            Summary = 'テスト仕様書の期待結果の空欄・あいまいさ、重複したケース、異常系・必須・型・境界値・コード値のケースの不足を見つけます。'
            Main = @('Path', 'DesignPath', 'OutXlsx')
            Labels = @{ Path = 'テスト仕様書のファイル・フォルダ'; DesignPath = '設計書のフォルダ（不足のケースを見るとき）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = 'テスト仕様書のチェック'; Keywords = 'テストケース レビュー'
        }
        @{
            Id = 'Evidence'; Label = 'テストのエビデンスを確かめる'; Group = 'テスト'; Tool = 'Test-TestEvidence.ps1'
            Summary = 'テスト仕様書の結果とエビデンス（画面のハードコピーなど）を突き合わせ、エビデンスの抜け・空・同じ画像の使い回しと、結果の書き忘れを見つけます。'
            Main = @('Path', 'EvidencePath', 'OutXlsx')
            Labels = @{ Path = 'テスト仕様書のファイル・フォルダ'; EvidencePath = 'エビデンスのフォルダ（空欄ならテスト仕様書と同じ場所）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = 'エビデンスのチェック'; Keywords = 'エビデンス 証跡 ハードコピー キャプチャ 使い回し'
        }
        @{
            Id = 'TestData'; Label = 'テストデータを作る'; Group = 'テスト'; Tool = 'New-TestData.ps1'
            Summary = 'テーブル定義書から、テーブルごとの CSV と INSERT 文のテストデータを作ります。主キー・外部キー・コード値を考え、桁数の最大・最小・NULL の行も作ります。'
            Main = @('Path', 'OutDir', 'Database', 'Table', 'OutXlsx')
            Labels = @{ Path = 'テーブル定義書（とコード定義書）のフォルダ・ファイル'; OutDir = 'テストデータの保存先のフォルダ'; Database = 'INSERT 文の DB の種類'; Table = '作るテーブル（空欄ならすべて。T_ORDER* のようにも書ける）'; OutXlsx = '一覧の保存先' }
            Required = @('Path', 'OutDir'); Output = @('OutXlsx'); OutputName = 'テストデータの一覧'; Open = @('OutDir'); Keywords = 'テストデータ INSERT CSV SQL 境界値'
        }
        @{
            Id = 'Trace'; Label = '要件・設計・テストのつながりを見る'; Group = 'テスト'; Tool = 'Test-Traceability.ps1'
            Summary = '要件一覧 → 設計書 → テスト仕様書のつながりと、画面・イベント・メッセージ・IF にテストがあるか（テストの抜け）を確かめます。'
            Main = @('DesignPath', 'TestPath', 'Requirement', 'OutXlsx')
            Labels = @{ DesignPath = '設計書のフォルダ'; TestPath = 'テスト仕様書のフォルダ'; Requirement = '要件一覧（無くてもよい）'; OutXlsx = '結果の保存先' }
            Output = @('OutXlsx'); OutputName = '網羅'; Keywords = 'トレーサビリティ 網羅 抜け'
        }
        @{
            Id = 'CodeTrace'; Label = '設計書とソースの ID を突き合わせる'; Group = 'テスト'; Tool = 'Test-CodeTrace.ps1'
            Summary = '設計書のメッセージ・画面・テーブルの ID がソースコードのどこにあるかを探し、ソースで使っていない ID と、ソースで使っているのに設計書に無い ID（書き間違い・設計書の直し忘れ）を見つけます。'
            Main = @('DesignPath', 'SourcePath', 'OutXlsx')
            Labels = @{ DesignPath = '設計書のフォルダ'; SourcePath = 'ソースコードのフォルダ'; OutXlsx = '結果の保存先' }
            Required = @('DesignPath', 'SourcePath'); Output = @('OutXlsx'); OutputName = '設計とソースの突き合わせ'; Keywords = 'ソース 実装 製造 メッセージ ID 未使用 トレーサビリティ'
        }
        @{
            Id = 'Regression'; Label = '再テストするケースを選ぶ'; Group = 'テスト'; Tool = 'Select-RegressionTests.ps1'
            Summary = '設計書の旧版と新版の差分から、再テストするテストケースを選びます（改修の回帰テストの範囲決め）。'
            Main = @('Old', 'New', 'TestPath', 'OutXlsx')
            Labels = @{ Old = '旧版の設計書'; New = '新版の設計書'; TestPath = 'テスト仕様書のフォルダ'; OutXlsx = '結果の保存先' }
            Output = @('OutXlsx'); OutputName = '再テスト'; Keywords = '回帰テスト 改修 デグレ'
        }
        @{
            Id = 'Similar'; Label = '障害の横展開の候補を探す'; Group = 'テスト'; Tool = 'Find-SimilarDefects.ps1'
            Summary = '障害の箇所と同じ作り（型・入力チェック・処理・カラムの型など）の箇所を設計書から集め、横展開の候補の一覧を作ります。Copilot に絞り込ませる依頼も作り、回答を入れると AI の判定を一覧に入れます。'
            Main = @('BugList', 'DesignPath', 'Id', 'OutXlsx', 'OutPrompt', 'Answer')
            Labels = @{ BugList = '障害管理表のファイル・フォルダ'; DesignPath = '設計書のフォルダ'; Id = '障害番号（カンマ区切り。空欄なら横展開が不要・済み以外のすべて）'; OutXlsx = '結果の保存先'; OutPrompt = 'Copilot への依頼の保存先'; Answer = 'Copilot の回答を保存したテキスト（AI の判定を入れるとき）' }
            Kinds = @{ Answer = 'File' }; Extensions = @{ OutPrompt = '.txt' }
            Required = @('BugList', 'DesignPath'); Output = @('OutXlsx', 'OutPrompt'); OutputName = '横展開の候補'; Keywords = '横展開 水平展開 類似 障害 バグ 不具合 Copilot AI'
        }

        # ---------------------------------------------------------------- 見積もり・進捗
        @{
            Id = 'Size'; Label = '規模を見積もる'; Group = '見積もり・進捗'; Tool = 'Measure-DesignSize.ps1'
            Summary = '設計書から画面・テーブル・IF などの規模を数え、係数表で工数の目安と FP（NESMA）を出します。前の版を指定すると、変更の規模も出します。'
            Main = @('Path', 'Baseline', 'OutXlsx')
            Labels = @{ Path = '設計書のフォルダ'; Baseline = '前の版の設計書（変更の規模を見るとき）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '規模'; Keywords = '見積もり FP 工数'
        }
        @{
            Id = 'Estimate'; Label = '見積書を検算する'; Group = '見積もり・進捗'; Tool = 'Test-ExcelEstimate.ps1'
            Summary = '見積書の合計の範囲漏れ・手入力の合計・周りと違う数式・文字列の数値・換算の日数や単価の食い違いを見つけます。'
            Main = @('Path', 'OutXlsx')
            Labels = @{ Path = '見積書のファイル・フォルダ'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '見積書の検算'; Keywords = '見積書 検算 合計'
        }
        @{
            Id = 'WbsDraft'; Label = 'WBS の下書きを作る'; Group = '見積もり・進捗'; Tool = 'New-WbsDraft.ps1'
            Summary = '設計書の規模の見積もりの工数を工程（詳細設計・レビュー・製造・単体テスト）に配分し、担当者と予定日（土日・休日を除く）を決めた WBS の下書きを作ります。進捗・工数・週次報告のツールがそのまま読めます。'
            Main = @('Path', 'Start', 'Members', 'Holidays', 'OutXlsx')
            Labels = @{ Path = '設計書のフォルダ'; Start = '開始日'; Members = '担当者（カンマ区切り。空欄なら担当者を決めない）'; Holidays = '休日の CSV（無くてもよい）'; OutXlsx = 'WBS の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = 'WBS'; Keywords = 'WBS 計画 スケジュール 日程 工数 見積もり 担当'
        }
        @{
            Id = 'Progress'; Label = '進捗を見る（遅延・EVM）'; Group = '見積もり・進捗'; Tool = 'Measure-Progress.ps1'
            Summary = 'WBS から遅れているタスク（営業日）と PV・EV・SPI を出します。'
            Main = @('Path', 'AsOf', 'Holidays', 'OutXlsx')
            Labels = @{ Path = 'WBS のファイル・フォルダ'; AsOf = '基準日（空欄なら今日）'; Holidays = '休日の CSV（無くてもよい）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '進捗'; Keywords = 'WBS 遅延 EVM SPI'
        }
        @{
            Id = 'Effort'; Label = '工数の実績を集計する'; Group = '見積もり・進捗'; Tool = 'Measure-Effort.ps1'
            Summary = '日報・工数実績表の工数をタスク・担当者・週ごとに集計し、WBS の予定工数と比べて予定超過のタスクと CPI を出します。'
            Main = @('Path', 'Wbs', 'AsOf', 'OutXlsx')
            Labels = @{ Path = '工数実績のファイル・フォルダ'; Wbs = 'WBS（予定工数と比べるとき）'; AsOf = '集計する日（空欄なら今日）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '工数'; Keywords = '工数 日報 CPI'
        }
        @{
            Id = 'Actions'; Label = '議事録から宿題を抜き出す'; Group = '見積もり・進捗'; Tool = 'Export-ActionItems.ps1'
            Summary = '議事録（Word・Excel・テキスト）から宿題（担当・期限付き）と決定事項を抜き出し、課題管理表の形の一覧にします。'
            Main = @('Path', 'OutXlsx', 'OutCsv')
            Labels = @{ Path = '議事録のファイル・フォルダ'; OutXlsx = '結果の保存先'; OutCsv = '課題管理表の CSV の保存先（週次報告に使うとき）' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = '議事録の宿題'; Keywords = '議事録 宿題 課題'
        }
        @{
            Id = 'Weekly'; Label = '週次報告を作る'; Group = '見積もり・進捗'; Tool = 'New-WeeklyReport.ps1'
            Summary = 'WBS・指摘管理表・障害管理表・課題管理表・工数実績から週次報告（進捗・指摘・障害と担当者ごとのやること）を作り、メールに貼れる文章と Excel に保存します。'
            Main = @('Config', 'Wbs', 'Findings', 'BugList', 'AsOf', 'OutDir')
            Labels = @{ Config = '週次報告の設定（CSV か、それを置いた案件のフォルダ）'; Wbs = 'WBS'; Findings = '指摘管理表'; BugList = '障害管理表'; AsOf = '報告の基準日（空欄なら今日）'; OutDir = '報告の保存先のフォルダ' }
            Open = @('OutDir'); Keywords = '週報 報告'
        }

        # ---------------------------------------------------------------- 品質の集計・分析
        @{
            Id = 'ReviewQuality'; Label = 'レビューの指摘を集計する'; Group = '品質の集計・分析'; Tool = 'Measure-ReviewQuality.ps1'
            Summary = '指摘管理表を集計し、設計書ごとの指摘件数・指摘密度・未完了・期限切れと、基準値から外れた設計書を出します。'
            Main = @('Path', 'DesignPath', 'ReviewLog', 'OutXlsx')
            Labels = @{ Path = '指摘管理表のファイル・フォルダ'; DesignPath = '設計書のフォルダ（指摘密度の分母）'; ReviewLog = 'レビュー記録（ゾーン分析をするとき）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = 'レビューの集計'; Keywords = '指摘 密度 ゾーン分析'
        }
        @{
            Id = 'TestQuality'; Label = 'テストの品質を見る（密度・収束）'; Group = '品質の集計・分析'; Tool = 'Measure-TestQuality.ps1'
            Summary = 'テスト仕様書と障害管理表から、機能ごとのテスト密度・バグ密度、テストの消化とバグの収束（信頼度成長曲線）を出します。'
            Main = @('Path', 'BugList', 'DesignPath', 'OutXlsx')
            Labels = @{ Path = 'テスト仕様書のファイル・フォルダ'; BugList = '障害管理表'; DesignPath = '設計書のフォルダ（密度の分母）'; OutXlsx = '結果の保存先' }
            Required = @('Path'); Output = @('OutXlsx'); OutputName = 'テストの品質'; Keywords = 'バグ 収束 密度'
        }
        @{
            Id = 'DefectRemoval'; Label = '工程ごとの摘出を見る'; Group = '品質の集計・分析'; Tool = 'Measure-DefectRemoval.ps1'
            Summary = 'バグとレビューの指摘の混入工程 × 摘出工程と、工程ごとの摘出率・流出を出します。'
            Main = @('BugList', 'Findings', 'OutXlsx')
            Labels = @{ BugList = '障害管理表'; Findings = '指摘管理表'; OutXlsx = '結果の保存先' }
            Required = @('BugList'); Output = @('OutXlsx'); OutputName = '工程別の摘出'; Keywords = '摘出率 流出 工程'
        }
        @{
            Id = 'ReviewEffect'; Label = 'レビューの効き目を見る'; Group = '品質の集計・分析'; Tool = 'Measure-ReviewEffectiveness.ps1'
            Summary = '設計書ごとにレビューの指摘と設計で入ったバグを突き合わせ、レビューで見逃した設計書と設計工程の摘出率を出します。'
            Main = @('Findings', 'BugList', 'DesignPath', 'OutXlsx')
            Labels = @{ Findings = '指摘管理表'; BugList = '障害管理表'; DesignPath = '設計書のフォルダ'; OutXlsx = '結果の保存先' }
            Kinds = @{ ReviewLog = 'Path' }; Output = @('OutXlsx'); OutputName = 'レビューの効き目'; Keywords = 'レビュー 見逃し'
        }
        @{
            Id = 'AiVsHuman'; Label = 'AI と人の指摘を突き合わせる'; Group = '品質の集計・分析'; Tool = 'Compare-ReviewFindings.ps1'
            Summary = 'AI のレビューと人のレビューの指摘を突き合わせ、一致・AI だけ・人だけと、観点ごとの一致率、観点表への反映案を出します。'
            Main = @('AiPath', 'HumanPath', 'OutXlsx')
            Labels = @{ AiPath = 'AI のレビュー結果'; HumanPath = '人の指摘管理表'; OutXlsx = '結果の保存先' }
            Output = @('OutXlsx'); OutputName = 'AIと人の指摘'; Keywords = 'AI Copilot 一致率'
        }
        @{
            Id = 'CodeSize'; Label = 'ソースコードの規模を数える'; Group = '品質の集計・分析'; Tool = 'Measure-CodeSize.ps1'
            Summary = 'ソースコードのステップ数（KLOC）を機能ごとに数えます。前の版を指定すると、新規・修正・削除の行数も出します。'
            Main = @('Path', 'Baseline', 'OutXlsx')
            Labels = @{ Path = 'ソースコードのフォルダ'; Baseline = '前の版のソースコードのフォルダ'; OutXlsx = '結果の保存先' }
            Output = @('OutXlsx'); OutputName = 'ソースコードの規模'; Keywords = 'KLOC ステップ数'
        }

        # ---------------------------------------------------------------- 品質会議・基準値
        @{
            Id = 'Quality'; Label = '品質評価を作る'; Group = '品質会議・基準値'; Tool = 'New-QualityReport.ps1'
            Summary = 'レビュー・テスト・工程・網羅・整合性の各ツールをまとめて実行し、工程完了判定に使う品質評価のブック（指標・実績・判定 ○△×）を作ります。'
            Main = @('Config', 'OutDir', 'AsOf')
            Labels = @{ Config = '品質評価の設定（CSV か、それを置いた案件のフォルダ）'; OutDir = '品質評価の保存先のフォルダ'; AsOf = '評価の時点（空欄なら今日）' }
            Open = @('OutDir'); Keywords = '品質評価 工程完了判定'
        }
        @{
            Id = 'Analysis'; Label = '品質評価の定性分析を書く（Copilot）'; Group = '品質会議・基準値'; Tool = 'Invoke-QualityAnalysis.ps1'
            Summary = '品質評価のブックの × と △ の指標について、定性分析（原因）と対策の下書きを Copilot に書かせて、ブックに書き込みます。PowerShell の画面が開くので、案内に沿ってやり取りしてください。'
            Main = @('Path')
            Labels = @{ Path = '品質評価のブック（またはそのフォルダ）' }
            Kinds = @{ Path = 'Path' }; Open = @('Path'); Console = $true; Keywords = 'Copilot 定性分析'
        }
        @{
            Id = 'Limits'; Label = '品質の基準値を作る'; Group = '品質会議・基準値'; Tool = 'New-QualityLimits.ps1'
            Summary = '過去のプロジェクトの品質の実績から、社内の基準値（下限・上限）の CSV を作ります。'
            Main = @('Path', 'Method', 'OutCsv')
            Labels = @{ Path = '実績の CSV のファイル・フォルダ'; Method = '求め方'; OutCsv = '基準値の CSV の保存先' }
            Output = @('OutCsv'); OutputName = '基準値'; Keywords = '基準値 実績'
        }
    )
}
