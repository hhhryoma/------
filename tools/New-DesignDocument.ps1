[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Create')]
param(
    [Parameter(ParameterSetName = 'Create', Mandatory = $true)]
    [string] $Template,

    [Parameter(ParameterSetName = 'Create', Mandatory = $true)]
    [Parameter(ParameterSetName = 'AddSheet', Mandatory = $true)]
    [string] $List,

    [Parameter(ParameterSetName = 'Create')]
    [Parameter(ParameterSetName = 'AddSheet')]
    [string] $ListSheet,

    [Parameter(ParameterSetName = 'Create')]
    [string] $FileName,

    [Parameter(ParameterSetName = 'Create')]
    [string] $OutDir = '.',

    [Parameter(ParameterSetName = 'AddSheet', Mandatory = $true)]
    [string] $Book,

    [Parameter(ParameterSetName = 'AddSheet', Mandatory = $true)]
    [string] $TemplateSheet,

    [Parameter(ParameterSetName = 'AddSheet', Mandatory = $true)]
    [string] $SheetName,

    [Parameter(ParameterSetName = 'Bump', Mandatory = $true)]
    [switch] $Bump,

    [Parameter(ParameterSetName = 'Bump', Mandatory = $true, Position = 0)]
    [string[]] $Path,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Description,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $CompareWith,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Section = '全体',

    [Parameter(ParameterSetName = 'Bump')]
    [string] $Version,

    [Parameter(ParameterSetName = 'Bump')]
    [string] $HistorySheet = '改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録',

    [Parameter(ParameterSetName = 'Bump')]
    [switch] $NoRecurse,

    [string] $Author = [System.Environment]::UserName,

    [datetime] $Date = (Get-Date).Date
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$replaceScript = Join-Path $PSScriptRoot 'Replace-Excel.ps1'
$compareScript = Join-Path $PSScriptRoot 'Compare-Excel.ps1'

Update-TypeData -TypeName 'DesignDocument.Result' -DefaultDisplayPropertySet File, Status, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
$dateText = $Date.ToString('yyyy/MM/dd', $invariant)

function New-Result {
    param([string] $File, [string] $Status, [string] $Detail, [string] $FilePath)
    $r = [pscustomobject]@{ File = $File; Status = $Status; Detail = $Detail; Path = $FilePath }
    $r.PSObject.TypeNames.Insert(0, 'DesignDocument.Result')
    return $r
}

# ================================================================ 雛形から作る（-Template / -List）


function Expand-Placeholder {
    #     ファイル名・シート名の {{名前}} を値に置き換える（値のセル内の改行は空白にする）
    param([string] $Text, [hashtable] $Values)
    return [regex]::Replace($Text, '\{\{([^{}]+)\}\}', [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $k = $m.Groups[1].Value.Trim()
        if ($Values.ContainsKey($k)) { return ([string]$Values[$k] -replace '\s*[\r\n]+\s*', ' ') }
        return $m.Value
    })
}

function Get-Placeholders {
    #     差し込み欄（{{名前}}）の名前の一覧
    param([string[]] $Texts)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($t in $Texts) {
        foreach ($m in [regex]::Matches([string]$t, '\{\{([^{}]+)\}\}')) { $n = $m.Groups[1].Value.Trim(); if (-not $names.Contains($n)) { $names.Add($n) } }
    }
    return , $names
}

function Read-ListRows {
    #     一覧（CSV、または画面一覧などの Excel）の行。返り値: @{ Values = 名前 → 値; Where = 一覧の場所 }
    param([string] $ListPath, [string] $Sheet, $Names)
    $rows = New-Object System.Collections.Generic.List[object]
    if ([System.IO.Path]::GetExtension($ListPath).ToLowerInvariant() -eq '.csv') {
        foreach ($r in (Read-CsvFile $ListPath $Names.ToArray() '一覧')) {
            $v = @{}
            foreach ($n in $Names) { $v[$n] = ([string]$r.$n).Trim() }
            $rows.Add([pscustomobject]@{ Values = $v; Where = '' })
        }
        return , $rows
    }
    $searchArgs = @{ Pattern = '\S'; Path = [System.Management.Automation.WildcardPattern]::Escape($ListPath); Target = 'Value' }
    if ($Sheet) { $searchArgs['SheetNamePattern'] = $Sheet }
    $books = ConvertTo-DesignBooks @(& $searchScript @searchArgs 6>$null)
    $labels = @($Names | ForEach-Object { ConvertTo-Label $_ })
    foreach ($b in $books.Values) {
        foreach ($sn in $b.Sheets.Keys) {
            $grid = $b.Sheets[$sn]
            foreach ($t in (Find-DesignTables $grid $labels $labels[0])) {
                foreach ($dr in $t.Rows) {
                    # セル内の改行は、Search-Excel では ↵ なので、改行に戻す
                    $v = @{}
                    for ($i = 0; $i -lt $Names.Count; $i++) { $v[$Names[$i]] = ([string]$dr.Cells[$t.ColumnOf[$labels[$i]]]).Replace('↵', "`n").Trim() }
                    $rows.Add([pscustomobject]@{ Values = $v; Where = ('{0}!{1}{2}' -f $sn, (ConvertTo-ColumnName $t.ColumnOf[$labels[0]]), $dr.Row) })
                }
            }
        }
    }
    if ($rows.Count -eq 0) { throw ("一覧に、見出し（{0}）がそろった表が見つかりません: {1}" -f ($Names -join '・'), $ListPath) }
    return , $rows
}

function Invoke-Fill {
    #     差し込み欄を値に置き換える（Replace-Excel.ps1。セル・図形・コメント・ヘッダー／フッター。セルの書式や部分ごとの書式は保たれる）。
    #     {{ 画面名 }} のように括弧の内側に空白があってもよい。SheetPattern を指定すると、そのシートだけを置き換える。
    #     返り値: 置き換えられなかったとき（Replace-Excel.ps1 の結果がスキップ・失敗）はその理由、置き換えたら空
    param([string] $FilePath, $Names, [hashtable] $Values, [string] $SheetPattern)
    $csv = Join-Path ([System.IO.Path]::GetTempPath()) ('placeholders-' + [guid]::NewGuid().ToString('N') + '.csv')
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('置換前,置換後,正規表現')
    # 正規表現で置き換えるので、値の $ は $$ にする
    foreach ($n in $Names) { $lines.Add(('"{0}","{1}",○' -f ('\{\{\s*' + [regex]::Escape($n) + '\s*\}\}').Replace('"', '""'), ([string]$Values[$n]).Replace('$', '$$').Replace('"', '""'))) }
    [System.IO.File]::WriteAllLines($csv, $lines.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
    $replaceArgs = @{ ListPath = $csv; Path = [System.Management.Automation.WildcardPattern]::Escape($FilePath); Force = $true; Confirm = $false; WhatIf = $false }
    if ($SheetPattern) { $replaceArgs['SheetNamePattern'] = $SheetPattern }
    try { $results = @(& $replaceScript @replaceArgs 6>$null) }
    finally { Remove-Item -LiteralPath $csv -Force -ErrorAction SilentlyContinue }
    return (@($results | Where-Object { $_.Result -eq 'スキップ' -or $_.Result -eq '失敗' } | ForEach-Object { $_.Detail } | Select-Object -Unique) -join ' / ')
}

$builtin = @{ '作成日' = $dateText; '日付' = $dateText; '作成者' = $Author }

if ($PSCmdlet.ParameterSetName -eq 'Create') {
    $Template = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Template)
    $List = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($List)
    $OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
    if (-not (Test-Path -LiteralPath $Template)) { throw "雛形が見つかりません: $Template" }
    if (@('.xlsx', '.xlsm') -notcontains [System.IO.Path]::GetExtension($Template).ToLowerInvariant()) { throw "雛形は .xlsx / .xlsm にしてください: $Template" }
    if (-not (Test-Path -LiteralPath $List)) { throw "一覧が見つかりません: $List" }
    if (-not $FileName) { $FileName = [System.IO.Path]::GetFileName($Template) }

    # 雛形とファイル名の差し込み欄
    $tplHits = @(& $searchScript -Pattern '\{\{[^{}]+\}\}' -Path ([System.Management.Automation.WildcardPattern]::Escape($Template)) -Target Value, Comment, Shape, HeaderFooter -NoExcel 6>$null)
    $inTemplate = Get-Placeholders @($tplHits | ForEach-Object { $_.Text })
    $inName = Get-Placeholders @($FileName)
    if ($inName.Count -eq 0) { throw 'ファイル名に {{画面ID}} のような差し込み欄を入れてください（-FileName）。同じ名前のファイルになってしまうためです' }
    $listNames = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($inName) + @($inTemplate)) { if (-not $builtin.ContainsKey($n) -and -not $listNames.Contains($n)) { $listNames.Add($n) } }
    $rows = Read-ListRows $List $ListSheet $listNames

    if (-not $WhatIfPreference -and -not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
    $invalid = [regex]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())
    $created = 0
    foreach ($row in $rows) {
        $values = @{}
        foreach ($k in $builtin.Keys) { $values[$k] = $builtin[$k] }
        foreach ($k in $row.Values.Keys) { $values[$k] = $row.Values[$k] }
        $empty = @($inName | Where-Object { -not $values[$_] })
        if ($empty.Count -gt 0) {
            New-Result '' 'スキップ' ('ファイル名の差し込み欄（{0}）の値が空です（{1}）' -f ($empty -join '・'), $row.Where) ''
            continue
        }
        $outName = [regex]::Replace((Expand-Placeholder $FileName $values), "[$invalid]", '_')
        $target = Join-Path $OutDir $outName
        if (Test-Path -LiteralPath $target) { New-Result $outName '既存' '同じ名前のファイルがあるため作りません' $target; continue }
        if (-not $PSCmdlet.ShouldProcess($outName, '雛形から作成')) { continue }

        # 雛形が読み取り専用でも、作った設計書は書き換えられるようにする（Replace-Excel.ps1 は読み取り専用のファイルを書き換えない）
        (Copy-Item -LiteralPath $Template -Destination $target -PassThru).IsReadOnly = $false
        $problem = ''
        if ($inTemplate.Count -gt 0) { $problem = Invoke-Fill $target $inTemplate $values }
        if ($problem) { New-Result $outName '失敗' ('差し込み欄を置き換えられませんでした（{0}）' -f $problem) $target; continue }
        $created++
        New-Result $outName '作成' '' $target
    }
    Write-Host ''
    Write-Host ('{0} ファイルを作りました（一覧の {1} 行）: {2}' -f $created, $rows.Count, $OutDir)
    return
}

# ================================================================ 雛形のシートからシートを足す（-Book / -TemplateSheet）


# シートと一緒にはコピーできないもの（リレーションの種類 → 表示の名前）
$uncopyableKinds = @{
    drawing = '図形・グラフ'; vmlDrawing = 'コメント・フォームのボタン'; comments = 'コメント'; threadedComment = 'コメント'; table = 'テーブル'
    pivotTable = 'ピボットテーブル'; oleObject = '埋め込みオブジェクト'; control = 'コントロール'; slicer = 'スライサー'; timeline = 'タイムライン'
}

function Add-SheetCopies {
    #     雛形のシートの写しを、ブックの末尾に足す（セル・書式・結合・印刷の設定・入力規則・条件付き書式・ハイパーリンクをそのまま写す）。
    #     プリンターごとの設定（printerSettings）は写さない（用紙・向き・拡大縮小などの印刷の設定はシートに残る）。
    #     雛形のシートだけに付いた名前（印刷範囲など）は、新しいシートにも付ける
    param([string] $FilePath, [string] $Template, [string[]] $Names)

    $changes = @{}
    $newParts = [ordered]@{}
    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        $read = {
            param([string] $Part)
            $e = Get-PackageEntry -Package $pkg -PartName $Part
            if ($null -eq $e) { return $null }
            $sr = New-Object System.IO.StreamReader($e.Open(), (New-Object System.Text.UTF8Encoding($false)), $true)
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        }
        $wbPart = @((Get-PartRelationships -Package $pkg -PartName '') | Where-Object { $_.Kind -eq 'officeDocument' })[0].Target
        $wbDir = ''
        if ($wbPart.Contains('/')) { $wbDir = $wbPart.Substring(0, $wbPart.LastIndexOf('/') + 1) }
        $wbXml = & $read $wbPart
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        $sheetEls = @($wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']"))
        $tplIndex = -1
        for ($i = 0; $i -lt $sheetEls.Count; $i++) { if ($sheetEls[$i].GetAttribute('name') -eq $Template) { $tplIndex = $i; break } }
        if ($tplIndex -lt 0) { throw "雛形のシート「$Template」が見つかりません" }
        $tplRid = Get-RelationshipIdAttribute $sheetEls[$tplIndex]
        $relsPart = '{0}_rels/{1}.rels' -f $wbDir, $wbPart.Substring($wbDir.Length)
        $relsXml = & $read $relsPart
        $relsDoc = Read-PartXml -Package $pkg -PartName $relsPart
        $tplRel = @($relsDoc.SelectNodes("//*[local-name()='Relationship']") | Where-Object { $_.GetAttribute('Id') -eq $tplRid })[0]
        $tplPart = Resolve-PartName -SourcePart $wbPart -TargetPath $tplRel.GetAttribute('Target')
        if ($tplRel.GetAttribute('Type') -notlike '*/worksheet') { throw "「$Template」はワークシートではないため、コピーできません" }

        # 雛形のシートのリレーション（コピーできないものがあれば中止する）
        $tplSheetXml = & $read $tplPart
        $slash = $tplPart.LastIndexOf('/')
        $tplRelsPart = $tplPart.Substring(0, $slash) + '/_rels/' + $tplPart.Substring($slash + 1) + '.rels'
        $tplRelsXml = & $read $tplRelsPart
        $dropIds = New-Object System.Collections.Generic.List[string]
        if ($tplRelsXml) {
            foreach ($m in [regex]::Matches($tplRelsXml, '<(?:\w+:)?Relationship\b[^>]*>')) {
                $type = [regex]::Match($m.Value, '\sType="([^"]*)"').Groups[1].Value
                $kind = $type.Substring($type.LastIndexOf('/') + 1)
                if ($uncopyableKinds.ContainsKey($kind)) { throw ("雛形のシート「{0}」に{1}があるため、コピーできません（{1}の無いシートを雛形にしてください）" -f $Template, $uncopyableKinds[$kind]) }
                if ($kind -eq 'printerSettings') { $dropIds.Add([regex]::Match($m.Value, '\sId="([^"]*)"').Groups[1].Value) }
            }
        }
        # 写しのシートの XML: 選択の状態（作業グループ）を外し、プリンターごとの設定への参照を消す
        $copyXml = [regex]::Replace($tplSheetXml, '\stabSelected="(?:1|true)"', '')
        foreach ($id in $dropIds) { $copyXml = [regex]::Replace($copyXml, ('\s(?:\w+:)?id="{0}"' -f [regex]::Escape($id)), '') }
        $copyRelsXml = $null
        if ($tplRelsXml) {
            $copyRelsXml = $tplRelsXml
            foreach ($id in $dropIds) { $copyRelsXml = [regex]::Replace($copyRelsXml, ('<(?:\w+:)?Relationship\b[^>]*\sId="{0}"[^>]*/>' -f [regex]::Escape($id)), '') }
            if ($copyRelsXml -notmatch '<(?:\w+:)?Relationship\b') { $copyRelsXml = $null }
        }

        $wpx = [regex]::Match($wbXml, '<(\w+:)?workbook\b').Groups[1].Value
        $rootTag = [regex]::Match($wbXml, '<(?:\w+:)?workbook\b[^>]*>').Value
        $rPrefix = [regex]::Match($rootTag, 'xmlns:(\w+)="[^"]*/relationships"').Groups[1].Value
        $maxId = 0
        foreach ($m in [regex]::Matches($wbXml, '\ssheetId="(\d+)"')) { $maxId = [math]::Max($maxId, [int]$m.Groups[1].Value) }
        $ids = @{}
        foreach ($m in [regex]::Matches($relsXml, '\sId="([^"]*)"')) { $ids[$m.Groups[1].Value] = $true }
        $wsType = $tplRel.GetAttribute('Type')
        $ct = & $read '[Content_Types].xml'
        $sheetEls2 = New-Object System.Text.StringBuilder
        $relEls = New-Object System.Text.StringBuilder
        $ctEls = New-Object System.Text.StringBuilder
        $nameEls = New-Object System.Text.StringBuilder
        $tplNames = @($wbDoc.SelectNodes("//*[local-name()='definedNames']/*[local-name()='definedName']") | Where-Object { $_.GetAttribute('localSheetId') -eq [string]$tplIndex })
        $quotedTpl = "'" + $Template.Replace("'", "''") + "'"
        $n = 1
        for ($k = 0; $k -lt $Names.Count; $k++) {
            while ($pkg.Entries.ContainsKey(('{0}worksheets/sheet{1}.xml' -f $wbDir, $n).ToLowerInvariant()) -or $newParts.Contains(('{0}worksheets/sheet{1}.xml' -f $wbDir, $n))) { $n++ }
            $part = '{0}worksheets/sheet{1}.xml' -f $wbDir, $n
            $newParts[$part] = $copyXml
            if ($copyRelsXml) { $newParts[('{0}worksheets/_rels/sheet{1}.xml.rels' -f $wbDir, $n)] = $copyRelsXml }
            $r = 1
            while ($ids.ContainsKey("rId$r")) { $r++ }
            $ids["rId$r"] = $true
            $rid = "rId$r"
            [void]$relEls.Append(('<Relationship Id="{0}" Type="{1}" Target="{2}"/>' -f $rid, $wsType, $part.Substring($wbDir.Length)))
            [void]$ctEls.Append(('<Override PartName="/{0}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $part))
            $name = ConvertTo-XmlText $Names[$k]
            # 雛形のシートが非表示でも、足したシートは表示する
            if ($rPrefix) { [void]$sheetEls2.Append(('<{0}sheet name="{1}" sheetId="{2}" {3}:id="{4}"/>' -f $wpx, $name, ($maxId + 1 + $k), $rPrefix, $rid)) }
            else { [void]$sheetEls2.Append(('<{0}sheet name="{1}" sheetId="{2}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{3}"/>' -f $wpx, $name, ($maxId + 1 + $k), $rid)) }
            # 印刷範囲・印刷タイトルなど、雛形のシートだけに付いた名前
            $quotedNew = "'" + $Names[$k].Replace("'", "''") + "'"
            foreach ($dn in $tplNames) {
                $formula = $dn.InnerText.Replace($quotedTpl + '!', $quotedNew + '!')
                $formula = [regex]::Replace($formula, ('(?<![\w''.])' + [regex]::Escape($Template) + '!'), [System.Text.RegularExpressions.MatchEvaluator] { param($m) $quotedNew + '!' })
                $attrs = New-Object System.Text.StringBuilder
                foreach ($a in $dn.Attributes) {
                    $v = $a.Value
                    if ($a.LocalName -eq 'localSheetId') { $v = [string]($sheetEls.Count + $k) }
                    [void]$attrs.Append((' {0}="{1}"' -f $a.Name, (ConvertTo-XmlText $v)))
                }
                [void]$nameEls.Append(('<{0}definedName{1}>{2}</{0}definedName>' -f $wpx, $attrs.ToString(), (ConvertTo-XmlText $formula)))
            }
        }
        $at = $wbXml.IndexOf(('</{0}sheets>' -f $wpx))
        $wbXml = $wbXml.Substring(0, $at) + $sheetEls2.ToString() + $wbXml.Substring($at)
        if ($nameEls.Length -gt 0) {
            $end = $wbXml.IndexOf(('</{0}definedNames>' -f $wpx))
            $wbXml = $wbXml.Substring(0, $end) + $nameEls.ToString() + $wbXml.Substring($end)
        }
        # 雛形の数式を、足したシートでも計算させる
        $changes[$wbPart.ToLowerInvariant()] = Set-FullCalcOnLoad $wbXml
        $at = $relsXml.LastIndexOf('</Relationships>')
        $changes[$relsPart.ToLowerInvariant()] = $relsXml.Substring(0, $at) + $relEls.ToString() + $relsXml.Substring($at)
        $at = $ct.LastIndexOf('</Types>')
        $changes['[content_types].xml'] = $ct.Substring(0, $at) + $ctEls.ToString() + $ct.Substring($at)
    }
    finally { $pkg.Zip.Dispose() }
    Save-ZipPackage -FilePath $FilePath -Changes $changes -NewParts $newParts
}

if ($PSCmdlet.ParameterSetName -eq 'AddSheet') {
    $Book = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Book)
    $List = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($List)
    if (-not (Test-Path -LiteralPath $Book)) { throw "ブックが見つかりません: $Book" }
    if (@('.xlsx', '.xlsm') -notcontains [System.IO.Path]::GetExtension($Book).ToLowerInvariant()) { throw "ブックは .xlsx / .xlsm にしてください: $Book" }
    if (-not (Test-Path -LiteralPath $List)) { throw "一覧が見つかりません: $List" }

    $pkg = Open-ZipPackage -FilePath $Book
    try { $existing = @((Get-WorkbookSheetParts $pkg).Sheets.Keys) } finally { $pkg.Zip.Dispose() }
    if ($existing -notcontains $TemplateSheet) { throw "雛形のシート「$TemplateSheet」が見つかりません（ブックのシート: $($existing -join '・')）" }
    # 大文字・小文字の違う指定（template と Template）でも雛形のシートを参照する名前（印刷範囲など）を付け替えられるよう、ブックのシート名にする
    $TemplateSheet = @($existing | Where-Object { $_ -eq $TemplateSheet })[0]

    $tplHits = @(& $searchScript -Pattern '\{\{[^{}]+\}\}' -Path ([System.Management.Automation.WildcardPattern]::Escape($Book)) -Target Value, Comment, Shape, HeaderFooter -SheetNamePattern ('^' + [regex]::Escape($TemplateSheet) + '$') -NoExcel 6>$null)
    $inTemplate = Get-Placeholders @($tplHits | ForEach-Object { $_.Text })
    $inName = Get-Placeholders @($SheetName)
    if ($inName.Count -eq 0) { throw 'シート名に {{テーブルID}} のような差し込み欄を入れてください（-SheetName）。同じ名前のシートになってしまうためです' }
    $listNames = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($inName) + @($inTemplate)) { if (-not $builtin.ContainsKey($n) -and -not $listNames.Contains($n)) { $listNames.Add($n) } }
    $rows = Read-ListRows $List $ListSheet $listNames

    $planned = New-Object System.Collections.Generic.List[object]
    $taken = @{}
    foreach ($e in $existing) { $taken[$e.ToUpperInvariant()] = $true }
    foreach ($row in $rows) {
        $values = @{}
        foreach ($k in $builtin.Keys) { $values[$k] = $builtin[$k] }
        foreach ($k in $row.Values.Keys) { $values[$k] = $row.Values[$k] }
        $empty = @($inName | Where-Object { -not $values[$_] })
        if ($empty.Count -gt 0) { New-Result '' 'スキップ' ('シート名の差し込み欄（{0}）の値が空です（{1}）' -f ($empty -join '・'), $row.Where) $Book; continue }
        # シート名に使えない文字は _ にし、31 文字までにする（Excel の決まり）
        $name = [regex]::Replace((Expand-Placeholder $SheetName $values), '[\[\]:*?/\\]', '_').Trim("'")
        if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
        if ($taken.ContainsKey($name.ToUpperInvariant())) { New-Result $name '既存' '同じ名前のシートがあるため作りません' $Book; continue }
        $taken[$name.ToUpperInvariant()] = $true
        $planned.Add([pscustomobject]@{ Name = $name; Values = $values })
    }
    if ($planned.Count -gt 0 -and $PSCmdlet.ShouldProcess([System.IO.Path]::GetFileName($Book), ('雛形のシート「{0}」から {1} シートを足す（{2}）' -f $TemplateSheet, $planned.Count, (@($planned | ForEach-Object { $_.Name }) -join '・')))) {
        Add-SheetCopies $Book $TemplateSheet @($planned | ForEach-Object { $_.Name })
        foreach ($p in $planned) {
            $problem = ''
            if ($inTemplate.Count -gt 0) { $problem = Invoke-Fill $Book $inTemplate $p.Values ('^' + [regex]::Escape($p.Name) + '$') }
            if ($problem) { New-Result $p.Name '失敗' ('雛形のシート「{0}」から作りましたが、差し込み欄を置き換えられませんでした（{1}）' -f $TemplateSheet, $problem) $Book; continue }
            New-Result $p.Name '作成' ('雛形のシート「{0}」から作りました' -f $TemplateSheet) $Book
        }
        Write-Host ''
        Write-Host ('{0} シートを足しました: {1}' -f $planned.Count, $Book)
    }
    return
}

# ================================================================ 版を上げる（-Bump）


function ConvertTo-VersionText {
    #     版数の文字（「第1.1版」「Ver.1.10」など）の数の部分（「1.1」「1.10」。全角は半角にし、桁はそのまま）。無ければ $null
    param([string] $Text)
    if (-not $Text) { return $null }
    $m = [regex]::Match($Text.Normalize($nfkc), '[0-9]+(?:\.[0-9]+)*')
    if (-not $m.Success) { return $null }
    return $m.Value
}

function Get-NextVersion {
    #     最後の数を 1 つ上げる（1.0 → 1.1、1.9 → 1.10）。最後の数の桁数は保つ（1.01 → 1.02、1.09 → 1.10）
    param([string] $Current)
    $parts = @($Current.Split('.'))
    if ($parts.Count -eq 1) { $parts += '0' }
    $last = $parts[$parts.Count - 1]
    $parts[$parts.Count - 1] = ([int64]$last + 1).ToString('D' + $last.Length, $invariant)
    return ($parts -join '.')
}

function Set-VersionText {
    #     「第1.0版」の数の部分だけを新しい版数にする（ほかの文字はそのまま。全角の数なら全角で書く）
    param([string] $Text, [string] $NewVersion)
    $m = [regex]::Match($Text, '[0-9０-９]+(?:[.．][0-9０-９]+)*')
    if (-not $m.Success) { return $NewVersion }
    if ($m.Value -match '[０-９]') { $NewVersion = -join @($NewVersion.ToCharArray() | ForEach-Object { [char]([int]$_ + 0xFEE0) }) }
    return $Text.Substring(0, $m.Index) + $NewVersion + $Text.Substring($m.Index + $m.Length)
}

function Read-DecimalStyles {
    #     セルの書式の番号 → 表示形式の小数点以下の桁数（「0.00」「#,##0.0」「0.0"版"」のように桁数を決めた表示形式のものだけ）
    param($Package, [string] $WorkbookPart)
    $digits = @{}
    $part = @((Get-PartRelationships -Package $Package -PartName $WorkbookPart) | Where-Object { $_.Kind -eq 'styles' } | ForEach-Object { $_.Target })[0]
    $doc = $null
    if ($part) { $doc = Read-PartXml -Package $Package -PartName $part }
    if (-not $doc) { return $digits }
    $codes = @{ '2' = '0.00'; '4' = '#,##0.00' }
    foreach ($nf in $doc.SelectNodes("//*[local-name()='numFmts']/*[local-name()='numFmt']")) { $codes[$nf.GetAttribute('numFmtId')] = $nf.GetAttribute('formatCode') }
    $i = 0
    foreach ($xf in $doc.SelectNodes("//*[local-name()='cellXfs']/*[local-name()='xf']")) {
        $m = [regex]::Match([string]$codes[$xf.GetAttribute('numFmtId')], '^(?:"[^"]*"|[#,0])*0\.(0+)(?:"[^"]*")*(?:;|$)')
        if ($m.Success) { $digits[[string]$i] = $m.Groups[1].Length }
        $i++
    }
    return $digits
}

function Format-ShownNumber {
    #     数値のセルの表示の文字（表示形式が小数点以下の桁数を決めたもの（Read-DecimalStyles）ならその桁数、ほかは R の形）
    param([double] $Number, $Cell, [hashtable] $Decimals)
    $s = $Cell.GetAttribute('s')
    if (-not $s) { $s = '0' }
    if ($Decimals.ContainsKey($s)) { return $Number.ToString('F' + $Decimals[$s], $invariant) }
    return $Number.ToString('R', $invariant)
}

function Get-ShownVersion {
    #     版数のセルの文字。数値のセルは表示の形で読む（Search-Excel.ps1 は表示形式によらずに読むため、0.00 の表示形式で 1.10 と表示した版数が 1.1 になる）
    param($Doc, [int] $Row, [int] $Col, [string] $Text, [hashtable] $Decimals)
    $rowEl = Get-SheetRowElement $Doc $Row
    if (-not $rowEl) { return $Text }
    $c = Get-SheetCellElement $Doc $rowEl $Row $Col
    $num = 0.0
    if ((Test-NumericCellElement $c) -and [double]::TryParse($Text, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$num)) { return Format-ShownNumber $num $c $Decimals }
    return $Text
}

function Set-CoverValue {
    #     表紙のセルを書き換え、書き換えた場所の名前を返す。数式のセル（改訂履歴を参照する LOOKUP など）は数式を残し、
    #     計算結果（Excel が計算し直すまでの表示）だけを書き換える（名前に「（数式の結果）」を付ける）
    param($Doc, $Cell, $Value, [string] $What)
    if (@(Get-XmlChildElements $Cell 'f').Count -eq 0) { Set-SheetCellValue $Doc $Cell $Value; return $What }
    foreach ($child in @($Cell.ChildNodes)) { if ($child.LocalName -ne 'f') { [void]$Cell.RemoveChild($child) } }
    $v = $Doc.CreateElement($Cell.Prefix, 'v', $Cell.NamespaceURI)
    if ($Value -is [string]) { $Cell.SetAttribute('t', 'str'); $v.InnerText = $Value }
    else { $Cell.RemoveAttribute('t'); $v.InnerText = ([double]$Value).ToString('R', $invariant) }
    [void]$Cell.AppendChild($v)
    return $What + '（数式の結果）'
}

$versionLabel = '^(版数|版|改訂版数|バージョン|Version|Ver\.?|Rev\.?|改版番号)$'
$coverDateLabel = '^(最終更新日|更新日|改訂日|改版日|最終改訂日)$'
$histDateLabel = '^(日付|改訂日|改版日|更新日|年月日|変更日)$'
$sectionLabel = '^(改訂箇所|改版箇所|変更箇所|改訂対象|対象|箇所|変更対象)$'
$descLabel = '^(改訂内容|改版内容|変更内容|内容|概要|改訂概要|変更概要)$'
$authorLabel = '^(改訂者|改版者|変更者|更新者|担当者|担当|作成者)$'

function Get-DateValue {
    #     日付のセルに書く値（もとが数値なら Excel のシリアル値、文字ならもとの書き方に合わせた文字）
    param($Cell, [string] $OldText)
    if (Test-NumericCellElement $Cell) { return ConvertTo-ExcelDate $Date }
    $t = ([string]$OldText).Normalize($nfkc)
    if ($t -match '\d{4}年') { return $Date.ToString('yyyy年M月d日', $invariant) }
    if ($t -match '\d{4}-\d') { return $Date.ToString('yyyy-MM-dd', $invariant) }
    return $dateText
}

function Find-LabelPairs {
    #     見出し（正規表現）に一致するセルと、その右の値のセルの組
    param($Grid, [string] $Pattern)
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($row in $Grid.Rows) {
        $cols = $Grid.Cols[$row]
        for ($i = 0; $i -lt $cols.Count - 1; $i++) {
            if ($Grid.Labels[$row][$cols[$i]] -match $Pattern) { $list.Add([pscustomobject]@{ Row = $row; Col = $cols[$i + 1]; Text = $Grid.Cells[$row][$cols[$i + 1]] }) }
        }
    }
    return , $list
}

if (-not $Description -and -not $CompareWith) { throw '-Bump では、改訂内容（-Description）か、比べる前の版（-CompareWith）を指定してください。' }

# 改訂履歴に一度に足す行の上限（超えた分は「ほか n 件」の 1 行にまとめる）
$maxHistoryRows = 20

function Get-HistoryDrafts {
    #     前の版と比べ、改訂履歴に書く行（改訂箇所・改訂内容）の下書きを求める（Compare-Excel.ps1 -OutHistory）
    param([string] $OldPath, [string] $NewPath)
    $csv = Join-Path ([System.IO.Path]::GetTempPath()) ('history-' + [guid]::NewGuid().ToString('N') + '.csv')
    try {
        $null = & $compareScript -Old $OldPath -New $NewPath -OutHistory $csv -HistorySheet $HistorySheet -NoExcel -WarningAction SilentlyContinue 6>$null
        if (-not (Test-Path -LiteralPath $csv)) { return @() }
        return @(Read-TextFile $csv | ConvertFrom-Csv | ForEach-Object { [pscustomobject]@{ Section = $_.'改訂箇所'; Description = $_.'改訂内容' } })
    }
    finally { Remove-Item -LiteralPath $csv -Force -ErrorAction SilentlyContinue }
}

$files = New-Object System.Collections.Generic.List[object]
foreach ($p in $Path) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $items) {
        if ($item.PSIsContainer) { $found = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) } else { $found = @($item) }
        foreach ($f in $found) { if (@('.xlsx', '.xlsm') -contains $f.Extension.ToLowerInvariant() -and -not $f.Name.StartsWith('~$')) { $files.Add($f) } }
    }
}

# 前の版: ファイルならそのファイル（版を上げるファイルが 1 つのとき）、フォルダなら版数を除いた名前が同じファイル
$oldByName = $null
if ($CompareWith) {
    $cw = Get-Item -LiteralPath $CompareWith -ErrorAction SilentlyContinue
    if (-not $cw) { throw "-CompareWith が見つかりません: $CompareWith" }
    if ($cw.PSIsContainer) {
        $oldByName = @{}
        foreach ($o in (Get-ChildItem -LiteralPath $cw.FullName -File -Recurse | Where-Object { @('.xlsx', '.xlsm') -contains $_.Extension.ToLowerInvariant() -and -not $_.Name.StartsWith('~$') })) {
            $k = Get-VersionlessName $o.Name
            if ($oldByName.ContainsKey($k)) { $oldByName[$k] = $null } else { $oldByName[$k] = $o.FullName }
        }
    }
    elseif ($files.Count -gt 1) { throw '-CompareWith にファイルを指定できるのは、版を上げるファイルが 1 つのときだけです（複数のときは前の版のフォルダを指定してください）。' }
}

foreach ($f in $files) {
    $oldPath = $null
    if ($CompareWith) {
        if ($oldByName) { $oldPath = $oldByName[(Get-VersionlessName $f.Name)] } else { $oldPath = $cw.FullName }
        if (-not $oldPath) { New-Result $f.Name 'スキップ' ('前の版のファイルが -CompareWith のフォルダに見つかりません（版数を除いた名前が同じファイルが 1 つだけあるものを使います）') $f.FullName; continue }
    }
    $hits = @(& $searchScript -Pattern '\S' -Path ([System.Management.Automation.WildcardPattern]::Escape($f.FullName)) -Target Value -NoExcel 6>$null)
    $books = ConvertTo-DesignBooks $hits
    if ($books.Count -eq 0) { New-Result $f.Name 'スキップ' '読めませんでした' $f.FullName; continue }
    $design = @($books.Values)[0]

    $pkg = Open-ZipPackage $f.FullName
    try {
        $wbParts = Get-WorkbookSheetParts $pkg
        $sheetParts = $wbParts.Sheets
        $coverName = @($sheetParts.Keys | Where-Object { $_ -match '^\s*表紙\s*$' })[0]
        if (-not $coverName) { $coverName = @($sheetParts.Keys)[0] }
        $histName = @($sheetParts.Keys | Where-Object { $_ -match $HistorySheet })[0]

        # ---- 表紙の版数・更新日
        $coverVersion = $null; $coverDate = $null
        if ($design.Sheets.Contains($coverName)) {
            $grid = $design.Sheets[$coverName]
            $pairs = Find-LabelPairs $grid $versionLabel
            foreach ($pair in $pairs) { if (ConvertTo-VersionText $pair.Text) { $coverVersion = $pair; break } }
            if (-not $coverVersion) {
                foreach ($row in $grid.Rows) { foreach ($col in $grid.Cols[$row]) {
                    if (-not $coverVersion -and $grid.Labels[$row][$col] -match '^(第\d+(\.\d+)*版|(Ver|Version|Rev)\.?\d+(\.\d+)*)$') { $coverVersion = [pscustomobject]@{ Row = $row; Col = $col; Text = $grid.Cells[$row][$col] } }
                } }
            }
            $pairs = Find-LabelPairs $grid $coverDateLabel
            if ($pairs.Count -gt 0) { $coverDate = $pairs[0] }
        }

        # ---- 改訂履歴: 版数と日付の見出しがある行のうち、下に版数の行がいちばん多いもの（Test-ExcelDeliverable.ps1 と同じ）
        $hist = $null
        if ($histName -and $design.Sheets.Contains($histName)) {
            $grid = $design.Sheets[$histName]
            foreach ($row in $grid.Rows) {
                $cols = @{}
                foreach ($col in $grid.Cols[$row]) {
                    $l = $grid.Labels[$row][$col]
                    if (-not $cols['Version'] -and $l -match $versionLabel) { $cols['Version'] = $col }
                    elseif (-not $cols['Date'] -and $l -match $histDateLabel) { $cols['Date'] = $col }
                    elseif (-not $cols['Section'] -and $l -match $sectionLabel) { $cols['Section'] = $col }
                    elseif (-not $cols['Description'] -and $l -match $descLabel) { $cols['Description'] = $col }
                    elseif (-not $cols['Author'] -and $l -match $authorLabel) { $cols['Author'] = $col }
                }
                if (-not $cols['Version'] -or -not $cols['Date']) { continue }
                $entries = @($grid.Rows | Where-Object { $_ -gt $row -and $grid.Cells[$_].ContainsKey($cols['Version']) -and (ConvertTo-VersionText $grid.Cells[$_][$cols['Version']]) })
                if (-not $hist -or $entries.Count -gt $hist.Entries.Count) { $hist = [pscustomobject]@{ HeaderRow = $row; Cols = $cols; Entries = $entries } }
            }
        }

        # 表紙と改訂履歴のシートの XML（ここに書き換える）。数値の版数のセルは、表示の形（0.00 の表示形式なら 1.10）で読む
        $docs = @{}
        foreach ($name in @($coverName, $histName)) {
            if ($name -and -not $docs.ContainsKey($name)) { $docs[$name] = Read-PartXmlForEdit $pkg $sheetParts[$name] }
        }
        $decimals = Read-DecimalStyles $pkg $wbParts.WorkbookPart
        if ($coverVersion) { $coverVersion.Text = Get-ShownVersion $docs[$coverName] $coverVersion.Row $coverVersion.Col $coverVersion.Text $decimals }
        $lastVersion = $null
        if ($hist -and $hist.Entries.Count -gt 0) {
            $entryRow = $hist.Entries[-1]
            $lastVersion = Get-ShownVersion $docs[$histName] $entryRow $hist.Cols['Version'] $design.Sheets[$histName].Cells[$entryRow][$hist.Cols['Version']] $decimals
        }

        $current = $null
        if ($lastVersion) { $current = ConvertTo-VersionText $lastVersion }
        elseif ($coverVersion) { $current = ConvertTo-VersionText $coverVersion.Text }
        if (-not $current -and -not $Version) { New-Result $f.Name 'スキップ' '表紙・改訂履歴に版数が見つかりません' $f.FullName; continue }
        $newVersion = $Version
        if (-not $newVersion) { $newVersion = Get-NextVersion $current }
        $newVersion = ConvertTo-VersionText $newVersion

        # 改訂履歴に書く行（改訂箇所・改訂内容）。-Description の行と、前の版と比べた下書き
        $entries = New-Object System.Collections.Generic.List[object]
        if ($Description) { $entries.Add([pscustomobject]@{ Section = $Section; Description = $Description }) }
        if ($oldPath) {
            $drafts = @(Get-HistoryDrafts $oldPath $f.FullName)
            if ($drafts.Count -gt $maxHistoryRows - $entries.Count) {
                $keep = $maxHistoryRows - $entries.Count - 1
                foreach ($d in $drafts[0..($keep - 1)]) { $entries.Add($d) }
                $entries.Add([pscustomobject]@{ Section = '全体'; Description = ('ほか {0} 件の変更（Compare-Excel.ps1 の差分を参照）' -f ($drafts.Count - $keep)) })
            }
            else { foreach ($d in $drafts) { $entries.Add($d) } }
            if ($entries.Count -eq 0) { New-Result $f.Name 'スキップ' ('前の版（{0}）と内容が同じです（改訂内容を -Description で指定してください）' -f [System.IO.Path]::GetFileName($oldPath)) $f.FullName; continue }
        }

        # 改訂履歴に足す行（最後に書いてある行の下。改訂履歴の列が空いていなければ足さない）
        $newRow = 0
        $last = 0
        if ($hist) {
            $grid = $design.Sheets[$histName]
            $last = $hist.HeaderRow
            foreach ($r in $grid.Rows) {
                if ($r -le $last) { continue }
                foreach ($col in $hist.Cols.Values) { if ($grid.Cells[$r].ContainsKey($col)) { $last = $r; break } }
            }
            $newRow = $last + 1
            $busy = @($newRow..($newRow + $entries.Count - 1) | Where-Object { $row = $_; $grid.Cells.ContainsKey($row) -and @($hist.Cols.Values | Where-Object { $grid.Cells[$row].ContainsKey($_) }).Count -gt 0 })
            if ($busy.Count -gt 0) { New-Result $f.Name 'スキップ' ('改訂履歴の {0} 行目が空いていないため、行を足せません' -f $busy[0]) $f.FullName; continue }
        }

        # 保存先: ファイル名に版数（_v1.0・第1.0版 など。「第2期」「第3章」は版数にしない）があれば新しい版数の名前で別に保存し、
        # 無ければ上書きする。版数が 2 つあれば後ろのもの
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
        $fms = [regex]::Matches($baseName, '(?i)(?<![a-z])v(?:er)?\.?\s*(\d+(?:\.\d+)*)|第\s*(\d+(?:\.\d+)*)\s*版')
        $outPath = $null
        if ($fms.Count -gt 0) {
            $g = $fms[$fms.Count - 1].Groups[1]
            if (-not $g.Success) { $g = $fms[$fms.Count - 1].Groups[2] }
            $outPath = Join-Path $f.DirectoryName ($baseName.Substring(0, $g.Index) + $newVersion + $baseName.Substring($g.Index + $g.Length) + $f.Extension)
            if (Test-Path -LiteralPath $outPath) { New-Result $f.Name 'スキップ' ('新しい版のファイルがすでにあります: {0}' -f [System.IO.Path]::GetFileName($outPath)) $f.FullName; continue }
        }
        $what = '版数 {0} → {1}' -f $current, $newVersion
        $shown = $f.Name
        if ($outPath) { $shown = '{0} → {1}' -f $f.Name, [System.IO.Path]::GetFileName($outPath) }
        if (-not $PSCmdlet.ShouldProcess($shown, $what)) { continue }

        $changes = @{}
        $done = New-Object System.Collections.Generic.List[string]
        if ($coverVersion) {
            $doc = $docs[$coverName]
            $c = Get-SheetCellElement $doc (Get-SheetRowElement $doc $coverVersion.Row -Create) $coverVersion.Row $coverVersion.Col -Create
            $value = Set-VersionText $coverVersion.Text $newVersion
            # 数値のセルには、表示が変わらない版数だけを数値で書く（1.9 の次の 1.10 を数値で書くと 1.1 と表示されるので、文字で書く）
            $num = 0.0
            if ((Test-NumericCellElement $c) -and [double]::TryParse($value, [System.Globalization.NumberStyles]::Float, $invariant, [ref]$num) -and (Format-ShownNumber $num $c $decimals) -eq $value) { $value = $num }
            $done.Add((Set-CoverValue $doc $c $value '表紙の版数'))
        }
        if ($coverDate) {
            $doc = $docs[$coverName]
            $c = Get-SheetCellElement $doc (Get-SheetRowElement $doc $coverDate.Row -Create) $coverDate.Row $coverDate.Col -Create
            $done.Add((Set-CoverValue $doc $c (Get-DateValue $c $coverDate.Text) '表紙の更新日'))
        }
        if ($hist) {
            $doc = $docs[$histName]
            $grid = $design.Sheets[$histName]
            # 版数・日付の書き方は最後の版の行に、セルの書式（罫線など）と結合は最後に書いてある行に合わせる
            $srcRow = $hist.HeaderRow
            if ($hist.Entries.Count -gt 0) { $srcRow = $hist.Entries[-1] }
            $srcEl = Get-SheetRowElement $doc $srcRow
            $styleEl = $null
            if ($last -gt $hist.HeaderRow) { $styleEl = Get-SheetRowElement $doc $last }
            $versionText = $newVersion
            if ($lastVersion) { $versionText = Set-VersionText $lastVersion $newVersion }
            for ($i = 0; $i -lt $entries.Count; $i++) {
                $row = $newRow + $i
                $rowEl = Get-SheetRowElement $doc $row -Create -TemplateRow $styleEl
                if ($styleEl) {
                    foreach ($sc in (Get-XmlChildElements $styleEl 'c')) {
                        $col = Get-ColumnIndex $sc.GetAttribute('r')
                        $null = Get-SheetCellElement $doc $rowEl $row $col -Create -Style $sc.GetAttribute('s')
                    }
                    Copy-SheetRowMerges $doc $last $row
                }
                # 版数・日付・改訂者は最初の行だけに書く（続きの行は改訂箇所・改訂内容だけ）
                $set = [ordered]@{ Section = $entries[$i].Section; Description = $entries[$i].Description }
                if ($i -eq 0) { $set = [ordered]@{ Version = $versionText; Date = $null; Section = $entries[$i].Section; Description = $entries[$i].Description; Author = $Author } }
                foreach ($k in $set.Keys) {
                    if (-not $hist.Cols[$k]) { continue }
                    $col = $hist.Cols[$k]
                    $c = Get-SheetCellElement $doc $rowEl $row $col -Create
                    $value = $set[$k]
                    if ($k -eq 'Date') {
                        $old = $null
                        if ($srcEl) { $old = Get-SheetCellElement $doc $srcEl $srcRow $col }
                        $oldText = ''
                        if ($grid.Cells.ContainsKey($srcRow)) { $oldText = $grid.Cells[$srcRow][$col] }
                        if (Test-NumericCellElement $old) { $value = Get-DateValue $old $oldText } else { $value = Get-DateValue $null $oldText }
                    }
                    if ($null -ne $value) { Set-SheetCellValue $doc $c $value }
                }
            }
            $lastNew = $newRow + $entries.Count - 1
            Update-SheetDimension $doc $lastNew
            if ($lastNew -gt $newRow) { $done.Add(('改訂履歴の {0}〜{1} 行目' -f $newRow, $lastNew)) } else { $done.Add(('改訂履歴の {0} 行目' -f $newRow)) }
        }
        foreach ($name in $docs.Keys) { $changes[$sheetParts[$name].ToLowerInvariant()] = ConvertTo-XmlString $docs[$name] }
    }
    finally { $pkg.Zip.Dispose() }

    if ($outPath) { Save-ZipPackage -FilePath $f.FullName -Changes $changes -OutPath $outPath; $resultPath = $outPath }
    else { Save-ZipPackage -FilePath $f.FullName -Changes $changes; $resultPath = $f.FullName }
    New-Result ([System.IO.Path]::GetFileName($resultPath)) '版上げ' ('{0}（{1}）' -f $what, ($done -join '・')) $resultPath
}

# .SYNOPSIS
#     設計書を雛形から一括で作り（-Template / -List）、雛形のシートからシートを足し（-Book / -TemplateSheet）、
#     または設計書の版を上げます（-Bump）。Excel 不要です。
#
# .DESCRIPTION
#     雛形から作る: 雛形の xlsx に {{画面ID}} {{画面名}} のような差し込み欄を書いておくと、一覧（画面一覧の Excel や CSV）の
#     行ごとに雛形をコピーし、差し込み欄を一覧の値に置き換えます（セル・図形・コメント・ヘッダー／フッター。Replace-Excel.ps1 を使います）。
#     {{作成日}}（{{日付}}）と {{作成者}} は -Date・-Author の値になります。同じ名前のファイルがあれば作りません。
#
#     シートを足す（-Book / -TemplateSheet）: テーブル定義書のように 1 つのブックに 1 件 1 シートで書く設計書で、
#     一覧（テーブル一覧など）の行ごとに雛形のシートをコピーしてブックの末尾に足し、そのシートの差し込み欄を置き換えます。
#     図形・コメント・テーブルなどのあるシートは雛形にできません。同じ名前のシートがあれば作りません。
#
#     版を上げる（-Bump）: 表紙の版数と更新日を新しい版数・日付にし、改訂履歴の最後の行の下に
#     版数・改訂日・改訂箇所・改訂内容・改訂者の行を足します（最後の行の書式と結合を写します）。
#     表紙の版数・更新日が数式のセルなら、数式は残し、計算結果だけを新しい版数・日付にします。
#     -CompareWith で前の版を渡すと、前の版との差分から改訂箇所・改訂内容の行を作ります。
#     ファイル名に版数（_v1.0・第1.0版 など）があれば新しい版数の名前で別のファイルに保存し、無ければ上書きします。
#     表紙と改訂履歴の見つけ方は Test-ExcelDeliverable.ps1 の版数のチェックと同じです。
#
# .PARAMETER Template
#     雛形の xlsx / xlsm。差し込み欄は {{名前}} の形で書きます。
#
# .PARAMETER List
#     一覧の Excel または CSV。差し込み欄と同じ名前の見出しの列を読みます（Excel は見出しがそろった表を探します）。
#
# .PARAMETER ListSheet
#     一覧の Excel のうち、読むシートの名前の正規表現。
#
# .PARAMETER FileName
#     作るファイルの名前。例: 'DD02_画面設計書_{{画面ID}}_{{画面名}}.xlsx'。既定は雛形のファイル名（差し込み欄が必要です）。
#
# .PARAMETER OutDir
#     作ったファイルを保存するフォルダ。既定はカレントフォルダ。
#
# .PARAMETER Book
#     シートを足すブック（xlsx / xlsm）。ブックを上書きします。
#
# .PARAMETER TemplateSheet
#     コピーする雛形のシートの名前（非表示のシートでもかまいません。足したシートは表示します）。
#     セル・書式・結合・印刷の設定・入力規則・条件付き書式をそのまま写し、印刷範囲などの名前も付けます。
#     プリンターごとの設定は写しません。図形・コメント・テーブル・ピボットテーブルのあるシートはコピーできません。
#
# .PARAMETER SheetName
#     足すシートの名前。例: '{{テーブルID}}'。シート名に使えない文字は _ にし、31 文字までにします。
#
# .PARAMETER Bump
#     版を上げます（-Path の設計書の表紙と改訂履歴を書き換えます）。
#
# .PARAMETER Path
#     版を上げる設計書のファイルまたはフォルダ（複数可・ワイルドカード可）。
#
# .PARAMETER Description
#     改訂履歴の改訂内容。-CompareWith を指定しないときは必須です。
#
# .PARAMETER CompareWith
#     前の版のファイル、または前の版のフォルダ。前の版と比べ（Compare-Excel.ps1 -OutHistory）、変更した行ごとに
#     「合計金額（TOTAL_AMOUNT）: 桁数を 9 → 11 に変更」のような改訂箇所・改訂内容の行を改訂履歴に足します。
#     版数・改訂日・改訂者は最初の行だけに書きます。-Description も指定すると、その行を最初に書きます。
#     フォルダを指定すると、版数を除いた名前が同じファイル（DD03_テーブル定義書_v1.0.xlsx と _v1.1.xlsx など）と比べます。
#     一度に足す行は 20 行までで、超えた分は「ほか n 件の変更」の 1 行にまとめます。
#
# .PARAMETER Section
#     改訂履歴の改訂箇所。既定は「全体」。
#
# .PARAMETER Version
#     新しい版数（例: 2.0）。既定は最後の数を 1 つ上げた版数（1.0 → 1.1。最後の数の桁数は保ち、1.01 → 1.02）。
#
# .PARAMETER HistorySheet
#     改訂履歴のシート名の正規表現。既定は「改訂履歴|変更履歴|改版履歴|更新履歴|改訂記録」。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER Author
#     {{作成者}} と改訂者に書く名前。既定はログオンしているユーザー名。
#
# .PARAMETER Date
#     {{作成日}} と改訂日・表紙の更新日に書く日付。既定は今日。
#
# .EXAMPLE
#     # 画面一覧の行ごとに、雛形から画面設計書を作る
#     .\New-DesignDocument.ps1 -Template .\画面設計書_雛形.xlsx -List D:\設計書\DD01_画面一覧.xlsx -ListSheet '^画面一覧$' -FileName 'DD02_画面設計書_{{画面ID}}_{{画面名}}.xlsx' -OutDir D:\設計書 -Author '山田 太郎'
#
# .EXAMPLE
#     # テーブル一覧の行ごとに、雛形のシートからテーブル定義のシートを足す
#     .\New-DesignDocument.ps1 -Book D:\設計書\DD03_テーブル定義書.xlsx -TemplateSheet '雛形' -List D:\設計書\DD03_テーブル定義書.xlsx -ListSheet '^テーブル一覧$' -SheetName '{{テーブルID}}'
#
# .EXAMPLE
#     # 版を上げる（まず -WhatIf で確かめる）
#     .\New-DesignDocument.ps1 -Bump D:\設計書\DD03_テーブル定義書_v1.1.xlsx -Description '配送時間帯を追加' -Section 'T_ORDER' -Author '佐藤 花子' -WhatIf
#
# .EXAMPLE
#     # 前の版（納品フォルダ）との差分から、改訂履歴の改訂箇所・改訂内容を書く
#     .\New-DesignDocument.ps1 -Bump D:\設計書 -CompareWith D:\納品\v1.0 -Author '佐藤 花子'
#
# .NOTES
#     要件: PowerShell 5.1 以降。Excel は不要です。同じフォルダに ExcelTools.psm1・Search-Excel.ps1・Replace-Excel.ps1 が必要です。
#     ヘッダー／フッター・シート名の差し込み欄は置き換えません。
