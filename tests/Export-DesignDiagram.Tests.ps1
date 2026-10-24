#     Export-DesignDiagram.ps1 のテスト。設計書は docs\design-sample（画面 3・遷移 6、テーブル 4・外部キー 3）

Import-Module (Join-Path $script:ToolsDir 'ExcelTools.psm1') -Force -DisableNameChecking

$DiagramTool = Join-Path $script:ToolsDir 'Export-DesignDiagram.ps1'
$designDir = Join-Path $script:RepoRoot 'docs/design-sample'

function Read-DrawingXml {
    #     xlsx の n 番目のシートの描画（図形）の XML
    param([string] $Path, [int] $Index)
    $pkg = Open-ZipPackage $Path
    try {
        $sheet = @((Get-WorkbookSheetParts $pkg).Sheets.Values)[$Index - 1]
        $drawing = @((Get-PartRelationships -Package $pkg -PartName $sheet) | Where-Object { $_.Kind -eq 'drawing' })[0].Target
        return (Read-PartXml -Package $pkg -PartName $drawing)
    }
    finally { $pkg.Zip.Dispose() }
}

Test-Case 'Export-DesignDiagram: イベント定義の遷移先から画面遷移を、テーブル定義書の外部キーからテーブルの関連を読む' {
    $edges = @(& $DiagramTool $designDir 6>$null)
    Assert-Equal @(
        '画面遷移|SCR001|SCR002|新規登録',
        '画面遷移|SCR001|SCR003|受注番号リンク',
        '画面遷移|SCR002|SCR003|登録',
        '画面遷移|SCR002|SCR001|戻る',
        '画面遷移|SCR003|SCR001|初期表示',
        '画面遷移|SCR003|SCR001|戻る',
        '外部キー|T_ORDER|M_CUSTOMER|CUSTOMER_CD',
        '外部キー|T_ORDER_DETAIL|T_ORDER|ORDER_NO',
        '外部キー|T_ORDER_DETAIL|M_PRODUCT|PRODUCT_CD') @($edges | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.Kind, $_.From, $_.To, $_.Label })
    Assert-Equal '受注一覧|受注登録|EV04 新規登録' ('{0}|{1}|{2}' -f $edges[0].FromName, $edges[0].ToName, $edges[0].Note)
    $text = Get-HostText { & $DiagramTool $designDir | Out-Null }
    Assert-Match $text '画面遷移図: 画面 3・遷移 6、ER 図: テーブル 4・外部キー 3'
}

Test-Case 'Export-DesignDiagram: -OutMarkdown に Mermaid の画面遷移図（flowchart）と ER 図（erDiagram）を保存する' {
    $dir = New-TestDir
    $md = Join-Path $dir '図.md'
    $null = & $DiagramTool $designDir -OutMarkdown $md 6>$null
    $text = [System.IO.File]::ReadAllText($md)
    Assert-Match $text '(?m)^flowchart LR$'
    Assert-Match $text '(?m)^    SCR001\["SCR001<br/>受注一覧"\]$'
    Assert-Match $text '(?m)^    SCR001 -->\|"新規登録"\| SCR002$'
    Assert-Match $text '(?m)^erDiagram$'
    Assert-Match $text '(?m)^    M_CUSTOMER \|\|--o\{ T_ORDER : "CUSTOMER_CD"$'
    Assert-Match $text '(?m)^        CHAR ORDER_NO PK, FK "受注番号"$' '主キーで外部キー'
    Assert-True ($text -notmatch 'ORDER_DATE') '-Columns Key（既定）は主キーと外部キーだけ'
    Assert-Match $text '\| T_ORDER_DETAIL \| 受注明細 \| T_ORDER（ORDER_NO）・M_PRODUCT（PRODUCT_CD） \|'

    $all = Join-Path $dir 'ER.md'
    $null = & $DiagramTool $designDir -Kind ER -Columns All -OutMarkdown $all 6>$null
    $text = [System.IO.File]::ReadAllText($all)
    Assert-Match $text '(?m)^        DATE ORDER_DATE "受注日"$' '-Columns All'
    Assert-True ($text -notmatch 'flowchart') '-Kind ER は ER 図だけ'
}

Test-Case 'Export-DesignDiagram: -OutXlsx に、四角と矢印（両向きは 1 本）の図形の図と一覧を保存する' {
    $dir = New-TestDir
    $out = Join-Path $dir '図.xlsx'
    $null = & $DiagramTool $designDir -OutXlsx $out 6>$null
    $book = Read-XlsxSheet -Path $out -Index 2
    Assert-Equal @('画面遷移図', '画面遷移', 'ER図', '外部キー') $book.Sheets
    Assert-Equal 'SCR001|受注一覧|EV04 新規登録|SCR002|受注登録' ('{0}|{1}|{2}|{3}|{4}' -f $book.Rows[2]['A'], $book.Rows[2]['B'], $book.Rows[2]['C'], $book.Rows[2]['D'], $book.Rows[2]['E'])
    Assert-True (@($book.Links | Where-Object { $_.Target -like '*DD02_*SCR001*.xlsx' -and $_.Location -like "'イベント定義'!A*" }).Count -gt 0) 'イベント定義の行へのリンク'

    $flow = Read-DrawingXml $out 1
    $ns = New-Object System.Xml.XmlNamespaceManager($flow.NameTable)
    $ns.AddNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $boxes = @($flow.SelectNodes('//xdr:sp[not(xdr:nvSpPr/xdr:cNvSpPr/@txBox)]', $ns) | ForEach-Object { $_.SelectSingleNode('xdr:nvSpPr/xdr:cNvPr', $ns).GetAttribute('name') })
    Assert-Equal @('SCR001', 'SCR002', 'SCR003') $boxes
    $arrows = @($flow.SelectNodes('//xdr:cxnSp', $ns))
    Assert-Equal 3 $arrows.Count '6 つの遷移を、画面の組ごとに 1 本（両向きは両端に矢じり）'
    $both = @($arrows | Where-Object { $_.SelectSingleNode('xdr:spPr/a:ln/a:headEnd', $ns) })
    Assert-Equal 2 $both.Count 'SCR001 と SCR002・SCR003 の間は両向き'
    $labels = @($flow.SelectNodes('//xdr:sp[xdr:nvSpPr/xdr:cNvSpPr/@txBox="1"]', $ns) | ForEach-Object { (@($_.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText })) -join '/' })
    Assert-True ($labels -contains '→ 受注番号リンク/← 初期表示・戻る') ($labels -join ' | ')
    $connect = $arrows[0].SelectSingleNode('xdr:nvCxnSpPr/xdr:cNvCxnSpPr', $ns)
    Assert-True ($null -ne $connect.SelectSingleNode('a:stCxn', $ns) -and $null -ne $connect.SelectSingleNode('a:endCxn', $ns)) '矢印は四角に付ける（四角を動かすと付いてくる）'

    $er = Read-DrawingXml $out 3
    $texts = @($er.SelectNodes('//xdr:sp', $ns) | ForEach-Object { (@($_.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText })) -join '/' })
    Assert-True ($texts -contains 'T_ORDER_DETAIL/受注明細/PK,FK ORDER_NO/PK LINE_NO/FK PRODUCT_CD') ($texts -join ' | ')
    Assert-Equal 3 @($er.SelectNodes('//xdr:cxnSp', $ns)).Count
}

Test-Case 'Export-DesignDiagram: 画面一覧に無い遷移先・テーブル定義書に無い参照先を警告し、図には点線の四角で入れる' {
    $dir = New-TestDir
    $design = Join-Path $dir '設計書'
    New-Item -ItemType Directory -Path $design | Out-Null
    Get-ChildItem -LiteralPath $designDir -Filter '*.xlsx' | Where-Object { $_.Name -match '^DD0[123]_' } | Copy-Item -Destination $design
    $scr2 = @(Get-ChildItem -LiteralPath $design -Filter '*SCR002*.xlsx')[0].FullName
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -Pattern '^SCR003$' -Replacement 'SCR009' -Path $scr2 -SheetNamePattern 'イベント定義' -Force 6>$null
    $table = @(Get-ChildItem -LiteralPath $design -Filter 'DD03_*.xlsx')[0].FullName
    $null = & (Join-Path $script:ToolsDir 'Replace-Excel.ps1') -Pattern '^M_PRODUCT\.PRODUCT_CD$' -Replacement 'M_ITEM.PRODUCT_CD' -Path $table -SheetNamePattern 'T_ORDER_DETAIL' -Force 6>$null
    $md = Join-Path $dir '図.md'
    $warnings = @()
    $null = & $DiagramTool $design -OutMarkdown $md -WarningVariable warnings -WarningAction SilentlyContinue 6>$null
    Assert-Equal @('SCR002 の登録（EV06）の遷移先「SCR009」が、画面一覧にありません', 'T_ORDER_DETAIL.PRODUCT_CD の外部キーの参照先「M_ITEM」が、テーブル定義書にありません') @($warnings | ForEach-Object { [string]$_ })
    $text = [System.IO.File]::ReadAllText($md)
    Assert-Match $text '(?m)^    class SCR009 missing$'
    Assert-Match $text '> - SCR002 の登録（EV06）の遷移先「SCR009」が、画面一覧にありません'
    Assert-Match $text '\| M_ITEM \| （テーブル定義書に無い） \|'
}

Test-Case 'Export-DesignDiagram: Get-Help でヘルプが表示される' {
    $h = Get-Help $DiagramTool
    Assert-Match $h.Synopsis '画面遷移図'
    Assert-True (@($h.parameters.parameter | Where-Object { -not $_.description }).Count -eq 0) '説明のないパラメーター'
}
