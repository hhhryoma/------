[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Position = 0)]
    [string] $Pattern,

    [Parameter(Position = 1)]
    [AllowEmptyString()]
    [string] $Replacement,

    [string] $ListPath,

    [Parameter(Position = 2)]
    [string[]] $Path = @('.'),

    [switch] $SimpleMatch,
    [switch] $CaseSensitive,

    [ValidateSet('Value', 'Formula', 'Comment', 'Shape', 'HeaderFooter', 'Validation')]
    [string[]] $Target = @('Value', 'Comment', 'Shape', 'HeaderFooter', 'Validation'),

    [switch] $NoRecurse,
    [switch] $SkipHiddenSheets,

    [string] $SheetNamePattern,
    [string] $SheetNameExclude,

    [string] $BackupDir,

    [switch] $Force,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml
Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

$wantValue   = $Target -contains 'Value'
$wantFormula = $Target -contains 'Formula'
$wantComment = $Target -contains 'Comment'
$wantShape   = $Target -contains 'Shape'
$wantHeaderFooter = $Target -contains 'HeaderFooter'
$wantValidation   = $Target -contains 'Validation'

$regexOptions = [System.Text.RegularExpressions.RegexOptions]::CultureInvariant
if (-not $CaseSensitive) { $regexOptions = $regexOptions -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }

function New-ReplaceRule {
    #     置換の規則。Protect は置換後の文字（$1 などを使わない場合）で、その一部として一致した箇所は置換しない
    #     （「ユーザ」→「ユーザー」を「ユーザー」に当てて「ユーザーー」にしない。何度実行しても同じ結果になる）
    param([string] $Find, [string] $To, [bool] $Literal)
    $text = $Find
    if ($Literal) { $text = [regex]::Escape($Find) }
    try { $rx = New-Object System.Text.RegularExpressions.Regex($text, $regexOptions) }
    catch { throw "検索パターンが正規表現として不正です: $Find （文字列として探すなら -SimpleMatch を付けてください） 詳細: $($_.Exception.InnerException.Message)" }
    $protect = $null
    if ($To -and ($Literal -or $To.IndexOf('$') -lt 0)) {
        $protect = New-Object System.Text.RegularExpressions.Regex([regex]::Escape($To), $regexOptions)
    }
    return [pscustomobject]@{ Find = $Find; Rx = $rx; To = $To; Literal = $Literal; Protect = $protect }
}

# 置換の規則: -Pattern / -Replacement の 1 つ、または -ListPath の CSV（置換前・置換後・正規表現）の行
$script:Rules = New-Object System.Collections.Generic.List[object]
if ($ListPath) {
    if ($PSBoundParameters.ContainsKey('Pattern')) { throw '-Pattern と -ListPath は同時に指定できません。' }
    $ListPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ListPath)
    if (-not (Test-Path -LiteralPath $ListPath)) { throw "置換リストが見つかりません: $ListPath" }
    $line = 1
    foreach ($row in @(Read-TextFile $ListPath | ConvertFrom-Csv)) {
        $line++
        $names = @($row.PSObject.Properties | ForEach-Object { $_.Name })
        if ($names -notcontains '置換前' -or $names -notcontains '置換後') { throw "置換リストには「置換前」「置換後」の列が必要です: $ListPath" }
        if (-not $row.'置換前') { continue }
        $isRegex = $names -contains '正規表現' -and @('○', '〇', '1', 'true', 'yes', 'はい') -contains ([string]$row.'正規表現').Trim().ToLowerInvariant()
        try { $script:Rules.Add((New-ReplaceRule $row.'置換前' ([string]$row.'置換後') (-not $isRegex))) }
        catch { throw ("置換リストの {0} 行目: {1}" -f $line, $_.Exception.Message) }
    }
    if ($script:Rules.Count -eq 0) { throw "置換リストに置換の規則がありません: $ListPath" }
}
else {
    if (-not $Pattern) { throw '置換する文字（-Pattern）と置換後の文字（-Replacement）、または置換リスト（-ListPath）を指定してください。' }
    if (-not $PSBoundParameters.ContainsKey('Replacement')) { throw '置換後の文字（-Replacement）を指定してください（削除するなら空文字 '''' を指定）。' }
    $script:Rules.Add((New-ReplaceRule $Pattern $Replacement ([bool]$SimpleMatch)))
}
foreach ($p in @($SheetNamePattern, $SheetNameExclude)) {
    if ($p) {
        try { [void][regex]::new($p) }
        catch { throw "シート名の正規表現が不正です: $p 詳細: $($_.Exception.InnerException.Message)" }
    }
}

if ($BackupDir) { $BackupDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($BackupDir) }
if ($OutXlsx) {
    $OutXlsx = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutXlsx)
    if ([System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
}

Update-TypeData -TypeName 'ReplaceExcel.Item' -DefaultDisplayPropertySet File, Location, Kind, Result, Change -Force -WhatIf:$false

$xmlNs = 'http://www.w3.org/XML/1998/namespace'

# ---------------------------------------------------------------- 文字の置換


function Invoke-TextReplace {
    #     文字の部分（<t> 要素）の並びに、置換の規則を順に当てる。
    #     返り値: 置換したら @(置換前, 置換後)、一致しなければ $null
    param($TextNodes, [switch] $PreserveSpace)

    $nodes = @($TextNodes)
    if ($nodes.Count -eq 0) { return $null }
    $before = -join @($nodes | ForEach-Object { $_.InnerText })
    $changed = $false
    foreach ($rule in $script:Rules) {
        if (Invoke-RuleReplace $nodes $rule -PreserveSpace:$PreserveSpace) { $changed = $true }
    }
    if (-not $changed) { return $null }
    $after = -join @($nodes | ForEach-Object { $_.InnerText })
    if ($after -ceq $before) { return $null }
    return @($before, $after)
}

function Invoke-StringReplace {
    #     文字列に置換の規則を当てる。返り値: 置換したら @(置換前, 置換後)、一致しなければ $null
    param([string] $Text)
    return Invoke-TextReplace @([pscustomobject]@{ InnerText = $Text })
}

function Test-AnyRuleMatch {
    param([string] $Text)
    foreach ($rule in $script:Rules) { if ($rule.Rx.IsMatch($Text)) { return $true } }
    return $false
}

function Invoke-RuleReplace {
    #     文字の部分（<t> 要素）の並びをつなげた文字列で 1 つの規則の置換をし、置換後の文字を元の部分に配り直す。
    #     置換した文字は一致した箇所の先頭がある部分に入れるので、ほかの文字の書式（部分ごとの太字・色など）は変わらない。
    #     空文字に一致した箇所と、置換後の文字の一部として一致した箇所は置換しない。返り値: 置換したか
    param($TextNodes, $Rule, [switch] $PreserveSpace)

    $nodes = @($TextNodes)
    $texts = @($nodes | ForEach-Object { $_.InnerText })
    $full = -join $texts
    $protected = @()
    if ($Rule.Protect) { $protected = @($Rule.Protect.Matches($full)) }
    $matches_ = @($Rule.Rx.Matches($full) | Where-Object {
        $m = $_
        $inside = $false
        foreach ($p in $protected) { if ($m.Index -ge $p.Index -and $m.Index + $m.Length -le $p.Index + $p.Length) { $inside = $true; break } }
        $m.Length -gt 0 -and -not $inside
    })
    if ($matches_.Count -eq 0) { return $false }

    $starts = New-Object 'int[]' ($nodes.Count + 1)
    for ($k = 0; $k -lt $nodes.Count; $k++) { $starts[$k + 1] = $starts[$k] + $texts[$k].Length }
    $out = @($nodes | ForEach-Object { New-Object System.Text.StringBuilder })
    $copy = {
        param([int] $From, [int] $To)
        for ($k = 0; $k -lt $nodes.Count; $k++) {
            $a = [math]::Max($From, $starts[$k]); $b = [math]::Min($To, $starts[$k + 1])
            if ($b -gt $a) { [void]$out[$k].Append($full.Substring($a, $b - $a)) }
        }
    }
    $pos = 0
    foreach ($m in $matches_) {
        & $copy $pos $m.Index
        $rep = $Rule.To
        if (-not $Rule.Literal) { $rep = $m.Result($Rule.To) }
        $owner = $nodes.Count - 1
        for ($k = 0; $k -lt $nodes.Count; $k++) { if ($m.Index -ge $starts[$k] -and $m.Index -lt $starts[$k + 1]) { $owner = $k; break } }
        [void]$out[$owner].Append($rep)
        $pos = $m.Index + $m.Length
    }
    & $copy $pos $full.Length

    $after = -join @($out | ForEach-Object { $_.ToString() })
    if ($after -ceq $full) { return $false }
    for ($k = 0; $k -lt $nodes.Count; $k++) {
        $t = $out[$k].ToString()
        if ($t -ceq $texts[$k]) { continue }
        $nodes[$k].InnerText = $t
        # 先頭・末尾の空白や改行を保つ
        if ($PreserveSpace -and $t -ne $t.Trim()) { [void]$nodes[$k].SetAttribute('space', $xmlNs, 'preserve') }
    }
    return $true
}

function Invoke-RichTextReplace {
    #     <si> / <is> / コメントの <text> の中の文字を置換する。文字が変わったらふりがな（<rPh>）は位置が合わなくなるので消す
    param([System.Xml.XmlElement] $Element)

    $nodes = @($Element.SelectNodes(".//*[local-name()='t'][not(ancestor::*[local-name()='rPh'])]"))
    $res = Invoke-TextReplace $nodes -PreserveSpace
    if ($null -eq $res) { return $null }
    foreach ($ph in @($Element.SelectNodes("*[local-name()='rPh']"))) { [void]$Element.RemoveChild($ph) }
    return $res
}

function Get-SheetCells {
    #     シートのセルを、番地・行・列と合わせて順に返す（r 属性を省略するツールで作られたファイル向けに位置を数える）
    param([System.Xml.XmlDocument] $Doc)
    $list = New-Object System.Collections.Generic.List[object]
    $rowNo = 0
    foreach ($row in $Doc.SelectNodes("//*[local-name()='sheetData']/*[local-name()='row']")) {
        $r = $row.GetAttribute('r')
        if ($r) { $rowNo = [int]$r } else { $rowNo++ }
        $colNo = 0
        foreach ($c in $row.SelectNodes("*[local-name()='c']")) {
            $ref = $c.GetAttribute('r')
            if ($ref) { $colNo = Get-ColumnIndex $ref }
            else { $colNo++; $ref = (ConvertTo-ColumnName $colNo) + $rowNo }
            $list.Add([pscustomobject]@{ Ref = $ref; Row = $rowNo; Col = $colNo; Cell = $c })
        }
    }
    return , $list
}

function Update-SharedFormulas {
    #     共有数式（Excel が数式をコピーしたときの保存形式。先頭のセルにだけ数式があり、写しのセルはそれを位置に合わせてずらして使う）を置換する。
    #     写しのセルも、それぞれの位置の数式で置換する。置換した後も「先頭の数式をずらしたもの」になっていれば先頭の数式だけを書き換え、
    #     一部のセルだけが変わるときなどは、各セルに数式を書いて共有をやめる。
    #     返り値: 共有数式のセル番地 → 置換したら @(置換前, 置換後)、変わらなければ $null
    param($Cells)

    $groups = [ordered]@{}
    foreach ($x in $Cells) {
        $f = $x.Cell.SelectSingleNode("*[local-name()='f']")
        if (-not $f -or $f.GetAttribute('t') -ne 'shared' -or -not $f.GetAttribute('si')) { continue }
        $si = $f.GetAttribute('si')
        if (-not $groups.Contains($si)) { $groups[$si] = New-Object System.Collections.Generic.List[object] }
        $groups[$si].Add([pscustomobject]@{ Ref = $x.Ref; Row = $x.Row; Col = $x.Col; F = $f; Text = $f.InnerText; New = $null })
    }

    $result = @{}
    foreach ($members in $groups.Values) {
        foreach ($m in $members) { $result[$m.Ref] = $null }
        $master = @($members | Where-Object { $_.Text })[0]
        # 先頭の数式が無い（壊れた）ものは変えない
        if (-not $master) { continue }
        $changed = $false
        foreach ($m in $members) {
            if ($m -ne $master) { $m.Text = Move-FormulaReference $master.Text ($m.Row - $master.Row) ($m.Col - $master.Col) }
            $m.New = $m.Text
            $res = Invoke-StringReplace $m.Text
            if ($res) { $m.New = $res[1]; $result[$m.Ref] = $res; $changed = $true }
        }
        if (-not $changed) { continue }
        $keep = $true
        foreach ($m in $members) {
            if ((Move-FormulaReference $master.New ($m.Row - $master.Row) ($m.Col - $master.Col)) -cne $m.New) { $keep = $false; break }
        }
        if ($keep) { $master.F.InnerText = $master.New; continue }
        foreach ($m in $members) {
            $m.F.InnerText = $m.New
            foreach ($a in @('t', 'si', 'ref')) { $m.F.RemoveAttribute($a) }
        }
    }
    return $result
}

function Update-HeaderFooter {
    #     ヘッダー・フッターの文字を置換する（書式（&"フォント" など）と、ページ番号などの項目はそのまま。項目をまたぐ一致は置換しない）。
    #     返り値: 置換した位置ごとの @{ Name; Before; After }
    param([System.Xml.XmlDocument] $Doc, [string] $FileName, [string] $SheetName)

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($hf in @($Doc.SelectNodes("//*[local-name()='headerFooter']/*"))) {
        $label = Get-HeaderFooterName $hf.LocalName
        if (-not $label -or -not $hf.InnerText) { continue }
        $tokens = Split-HeaderFooterText $hf.InnerText
        $before = Get-HeaderFooterSections $tokens
        # 位置の指定・項目で区切った、文字の並びごとに置換する
        $runs = New-Object System.Collections.Generic.List[object]
        $run = New-Object System.Collections.Generic.List[object]
        foreach ($t in $tokens) {
            if ($t.Kind -eq 'Text') { $run.Add($t) }
            elseif ($t.Kind -eq 'Section' -or $t.Kind -eq 'Field') {
                if ($run.Count -gt 0) { $runs.Add($run); $run = New-Object System.Collections.Generic.List[object] }
            }
        }
        if ($run.Count -gt 0) { $runs.Add($run) }
        $changed = $false
        foreach ($r in $runs) { if (Invoke-TextReplace $r.ToArray()) { $changed = $true } }
        if (-not $changed) { continue }
        $text = Join-HeaderFooterText $tokens
        if ($text.Length -gt 255) {
            # ヘッダー・フッターは 255 文字まで（超えると Excel が開くときに修復する）
            Write-Warning ("{0}: シート「{1}」の{2}は、置換すると 255 文字を超えるため置換しません。" -f $FileName, $SheetName, $label)
            continue
        }
        $after = Get-HeaderFooterSections $tokens
        foreach ($key in @('L', 'C', 'R')) {
            $b = [string](@($before | Where-Object { $_.Key -eq $key }) | ForEach-Object { $_.Text })
            $a = [string](@($after | Where-Object { $_.Key -eq $key }) | ForEach-Object { $_.Text })
            if ($a -cne $b) {
                $section = @{ L = '左'; C = '中央'; R = '右' }[$key]
                $result.Add([pscustomobject]@{ Name = ('{0}（{1}）' -f $label, $section); Before = $b; After = $a })
            }
        }
        $hf.InnerText = $text
    }
    return , $result
}

function Update-DataValidations {
    #     入力規則のリスト（"未着手,完了" のように直接書いたもの）と、入力時・エラーのメッセージを置換する。
    #     ほかのシートの範囲を参照するリストや条件の数式は、-Target Formula のときだけ置換する。
    #     返り値: 置換した箇所ごとの @{ Cell; Name; Before; After }
    param([System.Xml.XmlDocument] $Doc, [string] $FileName, [string] $SheetName)

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($dv in @($Doc.SelectNodes("//*[local-name()='dataValidation']"))) {
        $range = Get-ValidationRange $dv
        $first = ($range -split '[\s:]')[0]
        foreach ($v in (Get-ValidationTexts $dv)) {
            $isFormula = -not $v.IsList -and -not $v.Attribute
            if (($isFormula -and -not $wantFormula) -or (-not $isFormula -and -not $wantValidation)) { continue }
            $text = $v.Text
            if ($isFormula) { $text = $text.Substring(1) }
            $res = Invoke-StringReplace $text
            if (-not $res) { continue }
            $new = $res[1]
            if ($v.IsList) {
                # 直接書いたリストは 255 文字まで（超えると Excel が開くときに修復する）
                if ($new.Length -gt 255) {
                    Write-Warning ("{0}: 入力規則のリスト（{1}!{2}）は、置換すると 255 文字を超えるため置換しません。" -f $FileName, $SheetName, $range)
                    continue
                }
                $v.Node.InnerText = '"' + $new.Replace('"', '""') + '"'
            }
            elseif ($v.Attribute) { $v.Node.SetAttribute($v.Attribute, $new) }
            else { $v.Node.InnerText = $new }
            $before = $res[0]; $after = $new
            if ($isFormula) { $before = '=' + $before; $after = '=' + $after }
            $result.Add([pscustomobject]@{ Cell = $first; Name = ('{0}（{1}）' -f $v.Name, $range); Before = $before; After = $after })
        }
    }
    return , $result
}

function Get-TableHeaderCells {
    #     シートのテーブル（リスト）の見出しのセルを返す。見出しの文字はテーブルの列名と一致している必要があるため置換しない
    param($Package, [string] $SheetPart)
    $cells = @{}
    $rels = Get-PartRelationships -Package $Package -PartName $SheetPart
    foreach ($rel in @($rels | Where-Object { $_.Kind -eq 'table' })) {
        $doc = Read-PartXml -Package $Package -PartName $rel.Target
        if ($null -eq $doc) { continue }
        $table = $doc.DocumentElement
        if ($table.GetAttribute('headerRowCount') -eq '0') { continue }
        $m = [regex]::Match($table.GetAttribute('ref'), '^([A-Z]+)(\d+):([A-Z]+)\d+$')
        if (-not $m.Success) { continue }
        for ($c = (Get-ColumnIndex $m.Groups[1].Value); $c -le (Get-ColumnIndex $m.Groups[3].Value); $c++) { $cells[(ConvertTo-ColumnName $c) + $m.Groups[2].Value] = $true }
    }
    return $cells
}

function ConvertTo-OneLine {
    #     表示用に、OOXML のエスケープ（_x000D_ など）を戻し、改行を ↵ にした 1 行の文字列にする
    param([string] $Text)
    if ($null -eq $Text) { return '' }
    if ($Text.IndexOf('_x') -ge 0) {
        $Text = [regex]::Replace($Text, '_x([0-9A-Fa-f]{4})_', { param($m) [string][char][Convert]::ToInt32($m.Groups[1].Value, 16) })
        $Text = [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F]', '')
    }
    return $Text.Replace("`r`n", '↵').Replace("`n", '↵').Replace("`r", '↵')
}

function New-ReplaceItem {
    #     置換した（する）箇所の記録
    param([string] $Sheet, [string] $Cell, [string] $Kind, [string] $Before, [string] $After, [string] $Name)
    $location = $Sheet
    if ($Cell) { $location = '{0}!{1}' -f $Sheet, $Cell }
    return [pscustomobject]@{ Sheet = $Sheet; Cell = $Cell; Location = $location; Kind = $Kind; Before = (ConvertTo-OneLine $Before); After = (ConvertTo-OneLine $After); Name = $Name }
}

# ---------------------------------------------------------------- 1 ファイルの置換の計画


function Get-ReplacePlan {
    #     ファイルを読み、置換する箇所と、書き換えたパーツの内容を求める（ファイルはまだ書き換えない）
    param([string] $FilePath)

    $items = New-Object System.Collections.Generic.List[object]
    $changes = @{}
    $pkg = Open-ZipPackage -FilePath $FilePath
    try {
        $wbPart = 'xl/workbook.xml'
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName '')) { if ($rel.Kind -eq 'officeDocument') { $wbPart = $rel.Target; break } }
        $wbDoc = Read-PartXml -Package $pkg -PartName $wbPart
        if ($null -eq $wbDoc) { throw 'Excel ブックではありません（workbook.xml が見つかりません）' }
        $relById = @{}
        $sstPart = $null
        foreach ($rel in (Get-PartRelationships -Package $pkg -PartName $wbPart)) {
            $relById[$rel.Id] = $rel
            if ($rel.Kind -eq 'sharedStrings') { $sstPart = $rel.Target }
        }

        # 対象のシート（-SheetNamePattern などで外したシートの文字は変えない）
        $sheets = New-Object System.Collections.Generic.List[object]
        foreach ($s in $wbDoc.SelectNodes("//*[local-name()='sheets']/*[local-name()='sheet']")) {
            $rel = $relById[(Get-RelationshipIdAttribute $s)]
            if ($null -eq $rel -or ($rel.Kind -ne 'worksheet' -and $rel.Kind -ne 'macrosheet')) { continue }
            $name = $s.GetAttribute('name')
            $state = $s.GetAttribute('state')
            $included = $true
            if ($SkipHiddenSheets -and ($state -eq 'hidden' -or $state -eq 'veryHidden')) { $included = $false }
            if ($SheetNamePattern -and $name -notmatch $SheetNamePattern) { $included = $false }
            if ($SheetNameExclude -and $name -match $SheetNameExclude) { $included = $false }
            $headers = @{}
            if ($included) { $headers = Get-TableHeaderCells $pkg $rel.Target }
            $sheets.Add([pscustomobject]@{ Name = $name; Part = $rel.Target; Included = $included; Headers = $headers })
        }
        $skippedHeaders = New-Object System.Collections.Generic.List[string]

        # ---- 値（共有文字列）: 置換した文字列を使っているセルを探す
        $repoint = @{}    # シートのパーツ → (セル番地 → 新しい共有文字列の番号)
        $sstDoc = $null
        $sstChanged = @{}
        if ($wantValue -and $sstPart) {
            $sstDoc = Read-PartXmlForEdit -Package $pkg -PartName $sstPart
            $siList = @($sstDoc.DocumentElement.SelectNodes("*[local-name()='si']"))
            for ($i = 0; $i -lt $siList.Count; $i++) {
                $orig = $siList[$i].CloneNode($true)
                $res = Invoke-RichTextReplace $siList[$i]
                if ($res) { $sstChanged[$i] = [pscustomobject]@{ Before = $res[0]; After = $res[1]; Orig = $orig; New = $siList[$i]; In = New-Object System.Collections.Generic.List[object]; Out = 0 } }
            }
            if ($sstChanged.Count -gt 0) {
                foreach ($sh in $sheets) {
                    $reader = New-PartXmlReader -Package $pkg -PartName $sh.Part
                    if ($null -eq $reader) { continue }
                    try {
                        $ref = $null; $type = $null; $idx = 0; $rowNo = 0; $colNo = 0
                        while ($reader.Read()) {
                            if ($reader.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
                            if ($reader.LocalName -eq 'row') {
                                $r = $reader.GetAttribute('r')
                                if ($r) { $rowNo = [int]$r } else { $rowNo++ }
                                $colNo = 0
                            }
                            elseif ($reader.LocalName -eq 'c') {
                                $ref = $reader.GetAttribute('r'); $type = $reader.GetAttribute('t')
                                if ($ref) { $colNo = Get-ColumnIndex $ref } else { $colNo++; $ref = (ConvertTo-ColumnName $colNo) + $rowNo }
                            }
                            elseif ($reader.LocalName -eq 'v' -and $type -eq 's') {
                                if ([int]::TryParse($reader.ReadElementContentAsString(), [ref]$idx) -and $sstChanged.ContainsKey($idx)) {
                                    if ($sh.Included -and $ref -and -not $sh.Headers.ContainsKey($ref)) { $sstChanged[$idx].In.Add(@($sh, $ref)) }
                                    else {
                                        $sstChanged[$idx].Out++
                                        if ($sh.Included -and $ref) { $skippedHeaders.Add(('{0}!{1}' -f $sh.Name, $ref)) }
                                    }
                                }
                            }
                        }
                    }
                    finally { $reader.Dispose() }
                }
                $root = $sstDoc.DocumentElement
                $appended = 0
                foreach ($i in @($sstChanged.Keys | Sort-Object)) {
                    $c = $sstChanged[$i]
                    if ($c.In.Count -eq 0) { [void]$root.ReplaceChild($c.Orig, $c.New); continue }
                    if ($c.Out -gt 0) {
                        # 対象外のシートでも使っている文字列は元のまま残し、置換後の文字列を末尾に足して対象のセルだけ付け替える
                        [void]$root.ReplaceChild($c.Orig, $c.New)
                        [void]$root.AppendChild($c.New)
                        $newIndex = $siList.Count + $appended
                        $appended++
                        foreach ($x in $c.In) {
                            if (-not $repoint.ContainsKey($x[0].Part)) { $repoint[$x[0].Part] = @{} }
                            $repoint[$x[0].Part][$x[1]] = $newIndex
                        }
                    }
                    foreach ($x in $c.In) { $items.Add((New-ReplaceItem $x[0].Name $x[1] '値' $c.Before $c.After)) }
                }
                if (@($sstChanged.Values | Where-Object { $_.In.Count -gt 0 }).Count -gt 0) {
                    if ($appended -gt 0 -and $root.HasAttribute('uniqueCount')) { $root.SetAttribute('uniqueCount', [string]([int]$root.GetAttribute('uniqueCount') + $appended)) }
                    $changes[$sstPart.ToLowerInvariant()] = ConvertTo-XmlString $sstDoc
                }
            }
        }

        # ---- シートごと: セル内の文字列・共有文字列の付け替え・数式、コメント、図形
        $formulaChanged = $false
        foreach ($sh in $sheets) {
            if (-not $sh.Included) { continue }
            $entry = Get-PackageEntry -Package $pkg -PartName $sh.Part
            if ($null -eq $entry) { continue }
            $sr = New-Object System.IO.StreamReader($entry.Open())
            try { $raw = $sr.ReadToEnd() } finally { $sr.Dispose() }
            $needSheet = $repoint.ContainsKey($sh.Part) -or ($wantValue -and $raw.Contains('inlineStr')) -or ($wantFormula -and $raw -match '<(\w+:)?f[ >]') -or
                ($wantHeaderFooter -and $raw -match '<(\w+:)?(odd|even|first)(Header|Footer)>') -or (($wantValidation -or $wantFormula) -and $raw -match '<(\w+:)?dataValidation[ >]')
            if ($needSheet) {
                $doc = Read-PartXmlForEdit -Package $pkg -PartName $sh.Part
                $map = $repoint[$sh.Part]
                $edited = $false
                $cells = Get-SheetCells $doc
                $shared = @{}
                if ($wantFormula) { $shared = Update-SharedFormulas $cells }
                foreach ($x in $cells) {
                    $c = $x.Cell; $ref = $x.Ref
                    if ($map -and $map.ContainsKey($ref)) {
                        $v = $c.SelectSingleNode("*[local-name()='v']")
                        $v.InnerText = [string]$map[$ref]
                        $edited = $true
                    }
                    if ($wantValue -and $c.GetAttribute('t') -eq 'inlineStr') {
                        $is = $c.SelectSingleNode("*[local-name()='is']")
                        if ($is -and $sh.Headers.ContainsKey($ref)) {
                            if (Test-AnyRuleMatch $is.InnerText) { $skippedHeaders.Add(('{0}!{1}' -f $sh.Name, $ref)) }
                        }
                        elseif ($is) {
                            $res = Invoke-RichTextReplace $is
                            if ($res) { $items.Add((New-ReplaceItem $sh.Name $ref '値' $res[0] $res[1])); $edited = $true }
                        }
                    }
                    if ($wantFormula) {
                        $res = $null
                        if ($shared.ContainsKey($ref)) { $res = $shared[$ref] }
                        else {
                            $f = $c.SelectSingleNode("*[local-name()='f']")
                            if ($f -and $f.InnerText) { $res = Invoke-TextReplace @($f) }
                        }
                        if ($res) { $items.Add((New-ReplaceItem $sh.Name $ref '数式' ('=' + $res[0]) ('=' + $res[1]))); $edited = $true; $formulaChanged = $true }
                    }
                }
                if ($wantHeaderFooter) {
                    foreach ($h in (Update-HeaderFooter $doc ([System.IO.Path]::GetFileName($FilePath)) $sh.Name)) { $items.Add((New-ReplaceItem $sh.Name $null 'ヘッダー・フッター' $h.Before $h.After $h.Name)); $edited = $true }
                }
                if ($wantValidation -or $wantFormula) {
                    foreach ($v in (Update-DataValidations $doc ([System.IO.Path]::GetFileName($FilePath)) $sh.Name)) { $items.Add((New-ReplaceItem $sh.Name $v.Cell '入力規則' $v.Before $v.After $v.Name)); $edited = $true }
                }
                if ($edited) { $changes[$sh.Part.ToLowerInvariant()] = ConvertTo-XmlString $doc }
            }

            if (-not $wantComment -and -not $wantShape) { continue }
            $rels = Get-PartRelationships -Package $pkg -PartName $sh.Part
            if ($wantComment) {
                # スレッド形式のコメントは本文を置換する。互換用のメモにも同じ文字があるので、メモも置換して記録は 1 回だけにする
                $threaded = @{}
                foreach ($rel in @($rels | Where-Object { $_.Kind -eq 'threadedComment' })) {
                    $doc = Read-PartXmlForEdit -Package $pkg -PartName $rel.Target
                    if ($null -eq $doc) { continue }
                    $edited = $false
                    foreach ($tc in $doc.SelectNodes("//*[local-name()='threadedComment']")) {
                        $text = $tc.SelectSingleNode("*[local-name()='text']")
                        if ($null -eq $text) { continue }
                        $res = Invoke-TextReplace @($text)
                        if ($res) { $items.Add((New-ReplaceItem $sh.Name $tc.GetAttribute('ref') 'コメント' $res[0] $res[1])); $edited = $true }
                        $threaded[$tc.GetAttribute('ref')] = $true
                    }
                    if ($edited) { $changes[$rel.Target.ToLowerInvariant()] = ConvertTo-XmlString $doc }
                }
                foreach ($rel in @($rels | Where-Object { $_.Kind -eq 'comments' })) {
                    $doc = Read-PartXmlForEdit -Package $pkg -PartName $rel.Target
                    if ($null -eq $doc) { continue }
                    $edited = $false
                    foreach ($cm in $doc.SelectNodes("//*[local-name()='commentList']/*[local-name()='comment']")) {
                        $text = $cm.SelectSingleNode("*[local-name()='text']")
                        if ($null -eq $text) { continue }
                        $res = Invoke-RichTextReplace $text
                        if (-not $res) { continue }
                        $edited = $true
                        if (-not $threaded.ContainsKey($cm.GetAttribute('ref'))) { $items.Add((New-ReplaceItem $sh.Name $cm.GetAttribute('ref') 'コメント' $res[0] $res[1])) }
                    }
                    if ($edited) { $changes[$rel.Target.ToLowerInvariant()] = ConvertTo-XmlString $doc }
                }
            }
            if ($wantShape) {
                foreach ($rel in @($rels | Where-Object { $_.Kind -eq 'drawing' })) {
                    $doc = Read-PartXmlForEdit -Package $pkg -PartName $rel.Target
                    if ($null -eq $doc) { continue }
                    $edited = $false
                    foreach ($sp in $doc.SelectNodes("//*[local-name()='sp']")) {
                        $paras = @($sp.SelectNodes("*[local-name()='txBody']/*[local-name()='p']"))
                        if ($paras.Count -eq 0) { continue }
                        $before = (@($paras | ForEach-Object { $_.InnerText }) -join "`n").TrimEnd("`n")
                        $hit = $false
                        foreach ($p in $paras) {
                            # 段落の中の文字（<a:r> と <a:fld> の <a:t>）をつなげて置換する。改行（<a:br>）をまたぐ一致は置換しない
                            if (Invoke-TextReplace @($p.SelectNodes("*[local-name()='r' or local-name()='fld']/*[local-name()='t']"))) { $hit = $true }
                        }
                        if (-not $hit) { continue }
                        $edited = $true
                        # 互換用の代替表示（mc:Fallback）は同じ図形の写しなので、置換はするが記録はしない
                        if ($sp.SelectSingleNode("ancestor::*[local-name()='Fallback']")) { continue }
                        $after = (@($paras | ForEach-Object { $_.InnerText }) -join "`n").TrimEnd("`n")
                        $shapeName = $null
                        $nv = $sp.SelectSingleNode("*[local-name()='nvSpPr']/*[local-name()='cNvPr']")
                        if ($nv) { $shapeName = $nv.GetAttribute('name') }
                        $cell = $null
                        $from = $sp.SelectSingleNode("ancestor::*[local-name()='twoCellAnchor' or local-name()='oneCellAnchor'][1]/*[local-name()='from']")
                        if ($from) {
                            $col = $from.SelectSingleNode("*[local-name()='col']"); $row = $from.SelectSingleNode("*[local-name()='row']")
                            if ($col -and $row) { $cell = (ConvertTo-ColumnName ([int]$col.InnerText + 1)) + ([int]$row.InnerText + 1) }
                        }
                        $items.Add((New-ReplaceItem $sh.Name $cell '図形' $before $after $shapeName))
                    }
                    if ($edited) { $changes[$rel.Target.ToLowerInvariant()] = ConvertTo-XmlString $doc }
                }
            }
        }

        # ---- 数式を変えたら、開いたときに再計算させる（保存されている計算結果は古いため）
        if ($formulaChanged) {
            $entry = Get-PackageEntry -Package $pkg -PartName $wbPart
            $sr = New-Object System.IO.StreamReader($entry.Open())
            try { $wbXml = $sr.ReadToEnd() } finally { $sr.Dispose() }
            $changes[$wbPart.ToLowerInvariant()] = Set-FullCalcOnLoad $wbXml
        }
    }
    finally { $pkg.Zip.Dispose() }

    if ($skippedHeaders.Count -gt 0) {
        Write-Warning ("{0}: テーブルの見出しのセルは、テーブルの列名と一致している必要があるため置換しません: {1}" -f [System.IO.Path]::GetFileName($FilePath), (($skippedHeaders | Select-Object -Unique) -join ', '))
    }
    return [pscustomobject]@{ Items = $items; Changes = $changes }
}

function Test-FileLocked {
    # Excel で開いているファイルは書き換えられないので、先に確かめる
    param([string] $FilePath)
    try {
        $fs = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Dispose()
        return $false
    }
    catch { return $true }
}

function Get-FileSignature {
    # 先頭 4 バイトで実体を判定する（パスワード付きのファイルは中身が暗号化された OLE 形式になる）
    param([string] $FilePath)
    $fs = New-Object System.IO.FileStream($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, ([System.IO.FileShare]::ReadWrite -bor [System.IO.FileShare]::Delete))
    try {
        $buf = New-Object byte[] 4
        if ($fs.Read($buf, 0, 4) -lt 4) { return 'Unknown' }
        if ($buf[0] -eq 0x50 -and $buf[1] -eq 0x4B) { return 'Zip' }
        if ($buf[0] -eq 0xD0 -and $buf[1] -eq 0xCF) { return 'Ole' }
        return 'Unknown'
    }
    finally { $fs.Dispose() }
}

# ---------------------------------------------------------------- 実行


$files = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($p in $Path) {
    $found = @(Get-PathItem $p)
    if ($found.Count -eq 0) { Write-Warning "見つかりません: $p"; continue }
    foreach ($item in $found) {
        if ($item.PSIsContainer) {
            $root = $item.FullName.TrimEnd('\', '/')
            $list = Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue
        }
        else {
            $root = [System.IO.Path]::GetDirectoryName($item.FullName)
            $list = @($item)
        }
        foreach ($f in $list) {
            if (@('.xlsx', '.xlsm', '.xltx', '.xltm') -notcontains $f.Extension.ToLowerInvariant()) { continue }
            if ($f.Name.StartsWith('~$')) { continue }
            if ($BackupDir -and $f.FullName.StartsWith($BackupDir, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($OutXlsx -and [string]::Equals($f.FullName, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            if ($seen.ContainsKey($f.FullName)) { continue }
            $seen[$f.FullName] = $true
            $files.Add([pscustomobject]@{ Info = $f; Relative = $f.FullName.Substring($root.Length).TrimStart('\', '/') })
        }
    }
}
if ($files.Count -eq 0) { Write-Warning '対象の Excel ファイル（.xlsx / .xlsm）がありません。'; return }

# ---- まず全ファイルを読んで、置換する箇所を求める（ファイルはまだ変えない）
$results = New-Object System.Collections.Generic.List[object]
$plans = New-Object System.Collections.Generic.List[object]
$newResult = {
    param($Entry, $Item, [string] $Result, [string] $Detail)
    $r = [pscustomobject]@{
        File = $Entry.Relative; Location = ''; Sheet = ''; Cell = ''; Kind = ''; Before = ''; After = ''; Name = ''
        Result = $Result; Detail = $Detail; Change = $Detail; Path = $Entry.Info.FullName
    }
    if ($Item) {
        foreach ($k in @('Location', 'Sheet', 'Cell', 'Kind', 'Before', 'After', 'Name')) { $r.$k = $Item.$k }
        $short = { param($t) if ($t.Length -gt 40) { $t.Substring(0, 40) + '…' } else { $t } }
        $r.Change = '{0} → {1}' -f (& $short $Item.Before), (& $short $Item.After)
    }
    $r.PSObject.TypeNames.Insert(0, 'ReplaceExcel.Item')
    return $r
}
$index = 0
foreach ($entry in $files) {
    $index++
    Write-Progress -Activity '置換する箇所を探しています' -Status ("{0}/{1}  {2}" -f $index, $files.Count, $entry.Relative) -PercentComplete ([int](100 * $index / $files.Count))
    try {
        if ((Get-FileSignature $entry.Info.FullName) -ne 'Zip') { $results.Add((& $newResult $entry $null 'スキップ' 'パスワードで保護されているか、Excel ブックではないため読めません')); continue }
        $plan = Get-ReplacePlan $entry.Info.FullName
        if ($plan.Items.Count -eq 0) { continue }
        if ($entry.Info.IsReadOnly) {
            # 確定版などで意図的に読み取り専用にしていることが多いので、書き換えない
            $results.Add((& $newResult $entry $null 'スキップ' ('読み取り専用のファイルです（一致した箇所: {0} か所）' -f $plan.Items.Count)))
            continue
        }
        $plans.Add([pscustomobject]@{ Entry = $entry; Plan = $plan })
    }
    catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        $msg = $ex.Message
        if ($ex -is [System.IO.InvalidDataException]) { $msg = "ファイルが壊れているか Excel 形式ではありません（$msg）" }
        $results.Add((& $newResult $entry $null '失敗' $msg))
        Write-Warning ("{0}: {1}" -f $entry.Relative, $msg)
    }
}
Write-Progress -Activity '置換する箇所を探しています' -Completed

$total = 0
foreach ($p in $plans) { $total += $p.Plan.Items.Count }
Write-Host ''
if ($total -eq 0) {
    Write-Host ("{0} ファイルを調べました。置換する箇所はありません。" -f $files.Count)
    $results
    return
}

# ---- 置換する箇所を見せてから、まとめて確認する
if ($ListPath) { Write-Host ("置換リスト: {0}（{1} 件）" -f $ListPath, $script:Rules.Count) }
Write-Host ("{0} ファイル・{1} か所を置換します。" -f $plans.Count, $total)
$shown = 0
foreach ($p in $plans) {
    foreach ($it in $p.Plan.Items) {
        if ($shown -ge 20) { break }
        Write-Host ("  {0}  {1}  {2}: {3} → {4}" -f $p.Entry.Relative, $it.Location, $it.Kind, $it.Before, $it.After)
        $shown++
    }
}
if ($total -gt $shown) { Write-Host ("  …ほか {0} か所" -f ($total - $shown)) }

$go = $true
if (-not $WhatIfPreference -and -not $Force) {
    $go = $PSCmdlet.ShouldContinue(('{0} ファイル・{1} か所を置換します。よろしいですか？' -f $plans.Count, $total), '置換の確認')
}

foreach ($p in $plans) {
    $entry = $p.Entry
    $result = '置換予定'
    $detail = ''
    if (-not $go) { $result = '中止' }
    elseif ($PSCmdlet.ShouldProcess($entry.Relative, ('{0} か所を置換' -f $p.Plan.Items.Count))) {
        try {
            if (Test-FileLocked $entry.Info.FullName) { throw 'ファイルを開いているため書き換えられません（Excel で閉じてから実行してください）' }
            if ($BackupDir) {
                $dest = Join-Path $BackupDir $entry.Relative
                New-Item -ItemType Directory -Force -Path (Split-Path $dest -Parent) | Out-Null
                Copy-Item -LiteralPath $entry.Info.FullName -Destination $dest -Force
            }
            Save-ZipPackage -FilePath $entry.Info.FullName -Changes $p.Plan.Changes
            $result = '置換'
        }
        catch {
            $ex = $_.Exception
            while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
            $result = '失敗'
            $detail = $ex.Message
            Write-Warning ("{0}: {1}" -f $entry.Relative, $detail)
        }
    }
    foreach ($it in $p.Plan.Items) {
        $r = & $newResult $entry $it $result $detail
        $results.Add($r)
    }
}

$results

if ($OutXlsx) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $row = @{ Values = @($r.File, $r.Location, $r.Kind, $r.Result, $r.Before, $r.After, $r.Detail) }
        if ($r.Sheet) {
            $loc = ''
            if ($r.Cell) { $loc = "'{0}'!{1}" -f $r.Sheet.Replace("'", "''"), $r.Cell }
            $row.Link = @{ Column = 1; Path = $r.Path; Location = $loc }
        }
        $rows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(@{
            Name = '置換結果'
            Columns = @(@{ Header = 'ファイル'; Width = 36 }, @{ Header = '場所'; Width = 20 }, @{ Header = '種類'; Width = 9 }, @{ Header = '結果'; Width = 9 },
                @{ Header = '置換前'; Width = 45 }, @{ Header = '置換後'; Width = 45 }, @{ Header = '詳細'; Width = 30 })
            Rows = $rows
        })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$counts = $results | Group-Object Result | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }
Write-Host ("{0} ファイルを調べました: {1}" -f $files.Count, ($counts -join ' / '))
if ($BackupDir -and @($results | Where-Object Result -eq '置換').Count -gt 0) { Write-Host "置換前のファイル: $BackupDir" }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     複数の Excel ファイルの文字を一括で置換します（Excel 不要）。
#
# .DESCRIPTION
#     Search-Excel.ps1 と同じ指定で、セルの値・コメント・図形の文字・ヘッダー／フッター・入力規則のリスト
#     （-Target Formula で数式も）を置換します。
#     システム名・版数・旧称の書き換えなどに使えます。
#
#     まず全ファイルを調べて置換する箇所を表示し、確認してから書き換えます（-WhatIf なら表示だけ、-Force なら確認なし）。
#     文字列の一部だけに付けた書式（赤字・太字など）は、置換していない文字ではそのまま保ちます。
#     置換後の文字の一部として一致した箇所は置換しません（「ユーザ」→「ユーザー」を何度実行しても「ユーザーー」にならない）。
#     -ListPath で、CSV の置換リストに沿って複数の置換をまとめて行えます。
#     書き換えるのは置換した文字を含むパーツだけで、ほかの部分はそのままです。
#
#     読み取り専用のファイル、パスワード付きのファイルは書き換えません。Excel で開いているファイルは失敗として知らせます。
#
# .PARAMETER Pattern
#     置換する文字（正規表現）。-SimpleMatch を付けると、そのままの文字として扱います。
#
# .PARAMETER Replacement
#     置換後の文字。正規表現のときは $1 などで一致した部分を使えます（-SimpleMatch のときはそのままの文字）。空文字なら削除します。
#
# .PARAMETER ListPath
#     置換リストの CSV。複数の置換をまとめて行います（-Pattern / -Replacement の代わり）。
#     「置換前」「置換後」の列が必要で、「正規表現」の列が ○ / 1 / true の行は正規表現として扱います（UTF-8 / Shift_JIS）。
#     Test-ExcelTerms.ps1 -OutCsv で作った置換リストをそのまま使えます。
#
# .PARAMETER Path
#     対象のファイルまたはフォルダ（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER SimpleMatch
#     Pattern を正規表現ではなく、そのままの文字として扱います。
#
# .PARAMETER CaseSensitive
#     英字の大文字と小文字を区別します。
#
# .PARAMETER Target
#     置換する対象。Value（セルの文字列）/ Formula（数式）/ Comment（コメント）/ Shape（図形の文字）
#     / HeaderFooter（ヘッダー・フッター）/ Validation（入力規則のリストとメッセージ）。
#     既定は Formula 以外のすべて。数式を置換すると、開いたときに再計算されるようにします。
#     共有数式（コピーした数式）は、写しのセルもそれぞれの位置の数式で置換します。一部のセルだけが変わるときは、
#     各セルに数式を書いて共有をやめます。入力規則のうち、ほかのシートの範囲を参照するリストや条件の数式は
#     Formula を指定したときに置換します。ヘッダー・フッターの書式（フォントなど）とページ番号などの項目はそのままです。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER SkipHiddenSheets
#     非表示のシートは置換しません。
#
# .PARAMETER SheetNamePattern
#     この正規表現に一致する名前のシートだけを置換します。
#
# .PARAMETER SheetNameExclude
#     この正規表現に一致する名前のシートは置換しません。
#
# .PARAMETER BackupDir
#     書き換える前のファイルを、このフォルダに同じ相対パスで残します。
#
# .PARAMETER Force
#     確認せずに置換します。
#
# .PARAMETER OutXlsx
#     置換した（する）箇所の一覧を、該当セルへのリンク付きで xlsx に保存します。
#
# .EXAMPLE
#     # まず -WhatIf で置換する箇所を確かめる
#     .\Replace-Excel.ps1 '受注管理システム' '販売管理システム' D:\設計書 -SimpleMatch -WhatIf
#
# .EXAMPLE
#     # 確認してから置換し、元のファイルを残す
#     .\Replace-Excel.ps1 '受注管理システム' '販売管理システム' D:\設計書 -SimpleMatch -BackupDir D:\置換前
#
# .EXAMPLE
#     # 置換リスト（Test-ExcelTerms.ps1 -OutCsv で作ったものなど）でまとめて置換する
#     .\Replace-Excel.ps1 -ListPath D:\置換リスト.csv -Path D:\設計書 -WhatIf
#
# .EXAMPLE
#     # 正規表現: 「第1.0版」→「第1.1版」のように版数だけを書き換える（表紙シートだけ）
#     .\Replace-Excel.ps1 '第(\d+)\.0版' '第$1.1版' D:\設計書 -SheetNamePattern '^表紙$'
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm（.xltx / .xltm）が対象で、.xls / .xlsb は置換できません。
#     数値・日付のセル、シート名、ヘッダー／フッター、グラフの文字は置換しません。改行をまたぐ図形の文字も置換しません。
#     同じフォルダに ExcelTools.psm1 が必要です。
