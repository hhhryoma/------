[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string[]] $DesignPath,

    [Parameter(Mandatory = $true, Position = 1)]
    [string[]] $SourcePath,

    [ValidateSet('Message', 'Screen', 'Table', 'Column')]
    [string[]] $Kind = @('Message', 'Screen', 'Table'),

    [string[]] $Exclude = @('.git', '.svn', 'node_modules', 'bin', 'obj', 'target', 'build', 'dist', 'out', 'vendor', '.vs', '.idea'),

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'ddl/テーブル定義書の見出し.csv' }
foreach ($name in @('Headers', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

Update-TypeData -TypeName 'CodeTrace.Issue' -DefaultDisplayPropertySet Level, Item, Kind, Id, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC
$kindNames = @{ Message = 'メッセージ'; Screen = '画面'; Table = 'テーブル'; Column = 'カラム' }
# メッセージのリソース・設定のファイル（ソースと同じく、ID がどこかに書いてあれば実装にあるとみなす）
$resourceLanguage = @{ Name = '設定'; Ext = @('.properties', '.yml', '.yaml', '.ini', '.conf', '.cfg', '.toml', '.json', '.resx'); Line = @('#', ';'); Block = $null }
# SQL のテーブル名とみなさない名前
$sqlIgnore = @('DUAL', 'SELECT', 'SET', 'VALUES', 'WHERE', 'LATERAL', 'UNNEST', 'TABLE', 'ONLY', 'IF', 'EXISTS', 'NOWAIT', 'WAIT', 'OF', 'SKIP', 'ON')

# ---------------------------------------------------------------- 設計書の ID を読む


$searchArgs = @{ Pattern = '\S'; Path = $DesignPath; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$books = ConvertTo-DesignBooks $hits

# 設計書の ID: 種類 → ID（大文字）→ @{ Kind; Id; Name; File; Path; Sheet; Row; Hits（ソースの場所の一覧）}
$defined = @{}
foreach ($k in $kindNames.Keys) { $defined[$k] = [ordered]@{} }
function Add-Defined {
    param([string] $Kind, [string] $Id, [string] $Name, $Book, [string] $Sheet, [int] $Row)
    $id = ([string]$Id).Normalize($nfkc).Trim()
    if (-not $id) { return }
    $key = $id.ToUpperInvariant()
    if ($defined[$Kind].Contains($key)) { return }
    $defined[$Kind][$key] = [pscustomobject]@{ Kind = $Kind; Id = $id; Name = ([string]$Name).Trim(); File = $Book.File; Path = $Book.Path; Sheet = $Sheet; Row = $Row; Hits = (New-Object System.Collections.Generic.List[object]) }
}

if ($Kind -contains 'Message') {
    $design = Read-ScreenDesign $books
    foreach ($id in @($design.Messages.Keys | Sort-Object)) {
        $m = $design.Messages[$id]
        $book = [pscustomobject]@{ File = $m.File; Path = (@($books.Values | Where-Object { $_.File -eq $m.File })[0]).Path }
        Add-Defined 'Message' $id $m.Text $book $m.Sheet $m.Row
    }
}
if ($Kind -contains 'Screen') {
    $idPattern = '^[A-Za-z][A-Za-z0-9_\-]*\d[A-Za-z0-9_\-]*$'
    foreach ($book in $books.Values) {
        foreach ($sheetName in $book.Sheets.Keys) {
            $grid = $book.Sheets[$sheetName]
            # 画面一覧（画面ID の見出しの次に画面名の見出しがある表）
            foreach ($t in (Find-DesignTables $grid @('画面ID', '画面名') '画面ID')) {
                $labels = @($grid.Cols[$t.HeaderRow] | ForEach-Object { $grid.Labels[$t.HeaderRow][$_] })
                $at = [array]::IndexOf($labels, '画面ID')
                if ($at -lt 0 -or $at + 1 -ge $labels.Count -or $labels[$at + 1] -ne '画面名') { continue }
                foreach ($dr in $t.Rows) {
                    $id = ([string]$dr.Cells[$t.ColumnOf['画面ID']]).Trim()
                    if ($id -match $idPattern) { Add-Defined 'Screen' $id ([string]$dr.Cells[$t.ColumnOf['画面名']]) $book $sheetName $dr.Row }
                }
            }
            # 画面設計書のヘッダ（「画面ID」の右の値）
            $c = Find-LabelCell $grid '画面ID'
            if ($c -and $c.Value -match $idPattern) { Add-Defined 'Screen' $c.Value ([string](Find-LabelValue $grid '画面名')) $book $sheetName $c.Row }
        }
    }
}
# テーブル定義書のカラム名（SQL の FROM の後ろの名前でも、カラム名ならテーブルとみなさない。EXTRACT(YEAR FROM ORDER_DATE) など）
$columnNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
if ($Kind -contains 'Table' -or $Kind -contains 'Column') {
    $tables = Read-TableDesign -Books $books -Candidates (Read-TableHeaderCandidates $Headers)
    foreach ($t in $tables) {
        foreach ($c in $t.Columns) { [void]$columnNames.Add($c.Name) }
        $book = [pscustomobject]@{ File = $t.File; Path = $t.Place.Path }
        if ($Kind -contains 'Table') { Add-Defined 'Table' $t.Id $t.Name $book $t.Sheet ([int]([regex]::Match($t.Place.Cell, '\d+$').Value)) }
        if ($Kind -contains 'Column') {
            foreach ($c in $t.Columns) { Add-Defined 'Column' ('{0}.{1}' -f $t.Id, $c.Name) $c.Logical $book $t.Sheet ([int]([regex]::Match($c.Place.Cell, '\d+$').Value)) }
        }
    }
}
foreach ($k in $Kind) { if ($defined[$k].Count -eq 0) { Write-Warning ('設計書に{0}の ID が見つかりません: {1}' -f $kindNames[$k], ($DesignPath -join ', ')) } }

# ID の形（E0001 → 英字 E と数字 4 桁、MSG-E001 → MSG-E と数字 3 桁）。ソースにある同じ形の ID で、設計書に無いものを見つけるため
function Get-IdShapes {
    param($Ids)
    $shapes = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($id in $Ids) {
        $m = [regex]::Match($id, '^([A-Za-z]+(?:[_\-][A-Za-z]+)*[_\-]?)(\d+)$')
        if ($m.Success) { [void]$shapes.Add(('^{0}\d{{{1}}}$' -f [regex]::Escape($m.Groups[1].Value.ToUpperInvariant()), $m.Groups[2].Value.Length)) }
    }
    if ($shapes.Count -eq 0) { return $null }
    return New-Object System.Text.RegularExpressions.Regex(('(?:{0})' -f (@($shapes) -join '|')))
}
$shapeOf = @{}
foreach ($k in @('Message', 'Screen')) { if ($Kind -contains $k) { $shapeOf[$k] = Get-IdShapes @($defined[$k].Values | ForEach-Object { $_.Id }) } }

# カラムは、ORM の名前（orderNo・order_no）でも見つける（大文字にして _ を除いた形で比べる）
$columnKeys = @{}
foreach ($c in $defined['Column'].Values) {
    $bare = ($c.Id.Split('.')[-1] -replace '_', '').ToUpperInvariant()
    if (-not $columnKeys.ContainsKey($bare)) { $columnKeys[$bare] = New-Object System.Collections.Generic.List[object] }
    $columnKeys[$bare].Add($c)
}
$tableKeys = @{}
foreach ($t in $defined['Table'].Values) { $tableKeys[($t.Id -replace '_', '').ToUpperInvariant()] = $t }

# ---------------------------------------------------------------- ソースを読む


$files = New-Object System.Collections.Generic.List[object]
$seen = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$excludeSet = @($Exclude | ForEach-Object { $_.ToLowerInvariant() })
foreach ($p in $SourcePath) {
    $items = @(Get-PathItem $p)
    if ($items.Count -eq 0) { Write-Warning "ソースが見つかりません: $p"; continue }
    foreach ($item in $items) {
        $root = $item.FullName
        $list = @($item)
        if ($item.PSIsContainer) { $list = @(Get-ChildItem -LiteralPath $item.FullName -File -Recurse:(-not $NoRecurse) -ErrorAction SilentlyContinue) }
        else { $root = $item.DirectoryName }
        foreach ($f in $list) {
            $language = Get-SourceLanguage $f.Extension
            if (-not $language -and $resourceLanguage.Ext -contains $f.Extension.ToLowerInvariant()) { $language = $resourceLanguage }
            if (-not $language) { continue }
            $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
            $segments = @($rel -split '[\\/]')
            if ($segments.Count -gt 1 -and @($segments[0..($segments.Count - 2)] | Where-Object { $excludeSet -contains $_.ToLowerInvariant() }).Count -gt 0) { continue }
            if ($seen.Add($f.FullName)) { $files.Add([pscustomobject]@{ Info = $f; Relative = $rel; Language = $language }) }
        }
    }
}
# 相対パスの順（[Array]::Sort に PowerShell の配列を渡すと写しが並べ替わるだけなので、List の Sort で並べる）
$files.Sort([System.Comparison[object]] { param($a, $b) [string]::Compare($a.Relative, $b.Relative, [System.StringComparison]::OrdinalIgnoreCase) })

$tokenRx = New-Object System.Text.RegularExpressions.Regex('[A-Za-z0-9_]+(?:-[A-Za-z0-9_]+)*')
# SQL のテーブル（FROM・JOIN・UPDATE・TABLE と、INSERT・MERGE などの INTO の後ろの名前。.sql と MyBatis などの .xml のほかは、大文字の SQL だけ）。
# SELECT … INTO・FETCH … INTO の後ろ（PL/SQL の変数）と、FOR UPDATE・ON UPDATE CASCADE・ON DUPLICATE KEY UPDATE・FOR NO KEY UPDATE の後ろは見ない。
# INSERT の後ろのヒント（/*+ APPEND */）は飛ばす
$sqlUpdateInto = '(?<!\b(?:FOR|ON|KEY)\s+)UPDATE|(?<=\b(?:INSERT|MERGE|REPLACE|IGNORE)\s*(?:/\*.*?\*/\s*)?)INTO'
$sqlName = '\s+(?:[A-Za-z_][A-Za-z0-9_$#]*\.)?([A-Za-z_][A-Za-z0-9_$#]*)'
$sqlRxAny = New-Object System.Text.RegularExpressions.Regex(('\b(?:FROM|JOIN|TABLE|{0}){1}' -f $sqlUpdateInto, $sqlName), [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
$sqlRxUpper = New-Object System.Text.RegularExpressions.Regex(('\b(?:FROM|JOIN|{0}){1}' -f $sqlUpdateInto, $sqlName))
# .xml のタグ（MyBatis の <update id="…"> の update を SQL にしない。CDATA の中の SQL は見る。次の行に続くタグは行の終わりまで）
$xmlTagRx = New-Object System.Text.RegularExpressions.Regex('<!\[CDATA\[|\]\]>|<[/?!]?[A-Za-z_][^>]*>?')
$undefined = @{}   # 種類 → ID（大文字）→ @{ Id; Hits }
foreach ($k in $kindNames.Keys) { $undefined[$k] = [ordered]@{} }
function Add-Undefined {
    param([string] $Kind, [string] $Id, $Hit)
    $key = $Id.ToUpperInvariant()
    if (-not $undefined[$Kind].Contains($key)) { $undefined[$Kind][$key] = [pscustomobject]@{ Kind = $Kind; Id = $Id; Hits = (New-Object System.Collections.Generic.List[object]) } }
    $undefined[$Kind][$key].Hits.Add($Hit)
}

$readErrors = 0
foreach ($entry in $files) {
    try { $source = Measure-SourceLines $entry.Info.FullName $entry.Language -CodeText }
    catch { $readErrors++; Write-Warning ("ソースを読めません: {0} : {1}" -f $entry.Info.FullName, $_.Exception.Message); continue }
    $extension = $entry.Info.Extension.ToLowerInvariant()
    $isSql = @('.sql', '.pls', '.pks', '.pkb') -contains $extension
    # .xml は、MyBatis・iBATIS の SQL のファイル（mapper・sqlMap）だけ小文字の SQL も見る（ほかの .xml の英文の from を拾わない）
    if ($extension -eq '.xml') { $isSql = @($source.CodeLines | Where-Object { $_.Text -match '<(mapper|sqlMap)\b' }).Count -gt 0 }
    # 行の後ろのコメント（// …・/* … */・-- … など）の ID も見ない（コメントに残した古い ID を「使っている」に数えない）
    $hits = New-Object System.Collections.Generic.List[object]
    foreach ($line in $source.CodeLines) {
        $hit = [pscustomobject]@{ File = $entry.Info.Name; Relative = $entry.Relative; Path = $entry.Info.FullName; Line = $line.No; Text = $line.Text }
        $hits.Add($hit)
        $found = @{}
        foreach ($m in $tokenRx.Matches($line.CodeText)) {
            $token = $m.Value
            $upper = $token.ToUpperInvariant()
            if ($found.ContainsKey($upper)) { continue }
            $found[$upper] = $true
            foreach ($k in @('Message', 'Screen', 'Table')) {
                if ($Kind -notcontains $k) { continue }
                if ($defined[$k].Contains($upper)) { $defined[$k][$upper].Hits.Add($hit); continue }
                if ($shapeOf[$k] -and $shapeOf[$k].IsMatch($upper)) { Add-Undefined $k $token $hit }
            }
            if ($Kind -contains 'Table' -and -not $defined['Table'].Contains($upper)) {
                $bare = ($upper -replace '_', '')
                if ($tableKeys.ContainsKey($bare)) { $tableKeys[$bare].Hits.Add($hit) }
            }
            if ($Kind -contains 'Column') {
                $bare = ($upper -replace '_', '')
                if ($columnKeys.ContainsKey($bare)) { foreach ($c in $columnKeys[$bare]) { $c.Hits.Add($hit) } }
            }
        }
    }
    if ($Kind -contains 'Table' -and $hits.Count -gt 0) {
        # .sql と MyBatis などの .xml は、コメントの行を除いた行をつないで見る（FROM と次の行のテーブル名、INSERT と次の行の INTO など）。
        # ほかのファイル（Java の文字列の SQL・設定ファイルなど）は行をまたがない（行の間に、空白でない「|」を挟む）
        $rx = $sqlRxUpper
        $separator = "`n|"
        if ($isSql) { $rx = $sqlRxAny; $separator = "`n" }
        $starts = New-Object 'System.Collections.Generic.List[int]'
        $joined = New-Object System.Text.StringBuilder
        foreach ($line in $source.CodeLines) {
            $starts.Add($joined.Length)
            $sqlText = $line.CodeText
            if ($extension -eq '.xml') { $sqlText = $xmlTagRx.Replace($sqlText, ' ') }
            [void]$joined.Append($sqlText).Append($separator)
        }
        $at = 0
        foreach ($m in $rx.Matches($joined.ToString())) {
            # テーブル名のある行
            while ($at + 1 -lt $starts.Count -and $starts[$at + 1] -le $m.Groups[1].Index) { $at++ }
            $name = $m.Groups[1].Value
            $upper = $name.ToUpperInvariant()
            if ($defined['Table'].Contains($upper) -or $sqlIgnore -contains $upper -or $columnNames.Contains($upper)) { continue }
            Add-Undefined 'Table' $name $hits[$at]
        }
    }
}

# ---------------------------------------------------------------- 確かめる


$issues = New-Object System.Collections.Generic.List[object]
function Get-HitText {
    #     ソースの場所（「Foo.java 12 行目 ほか 2 か所」）
    param($Hits)
    $all = @($Hits | ForEach-Object { $_ })
    if ($all.Count -eq 0) { return '' }
    $text = '{0} {1} 行目' -f $all[0].Relative, $all[0].Line
    if ($all.Count -gt 1) { $text += (' ほか {0} か所' -f ($all.Count - 1)) }
    return $text
}

foreach ($k in @('Message', 'Screen', 'Table', 'Column')) {
    if ($Kind -notcontains $k) { continue }
    foreach ($d in $defined[$k].Values) {
        if ($d.Hits.Count -gt 0) { continue }
        $level = '警告'
        if ($k -eq 'Column') { $level = '情報' }
        $name = ''
        if ($d.Name) { $name = '「{0}」' -f $d.Name }
        $i = [pscustomobject]@{ Level = $level; Item = '実装に無い'; Kind = $kindNames[$k]; Id = $d.Id; Detail = ('{0} {1}{2}がソースに見つかりません（使っていないか、ID の書き間違い）' -f $kindNames[$k], $d.Id, $name); File = $d.File; Location = ('{0} {1} 行目' -f $d.Sheet, $d.Row); Path = $d.Path; Sheet = $d.Sheet; Row = $d.Row; Seq = $issues.Count }
        $i.PSObject.TypeNames.Insert(0, 'CodeTrace.Issue')
        $issues.Add($i)
    }
    foreach ($u in $undefined[$k].Values) {
        $what = @{ Message = 'メッセージ ID'; Screen = '画面 ID'; Table = ' SQL のテーブル'; Column = 'カラム' }[$k]
        $where = @{ Message = 'メッセージ一覧'; Screen = '画面一覧・画面設計書'; Table = 'テーブル定義書'; Column = 'テーブル定義書' }[$k]
        $first = $u.Hits[0]
        $i = [pscustomobject]@{ Level = '警告'; Item = '設計書に無い'; Kind = $kindNames[$k]; Id = $u.Id; Detail = ('ソースの{0} {1}（{2}）が、{3}にありません' -f $what, $u.Id, (Get-HitText $u.Hits), $where); File = $first.Relative; Location = ('{0} 行目' -f $first.Line); Path = $first.Path; Sheet = ''; Row = $first.Line; Seq = $issues.Count }
        $i.PSObject.TypeNames.Insert(0, 'CodeTrace.Issue')
        $issues.Add($i)
    }
}

# ---------------------------------------------------------------- 結果


$sorted = @($issues | Sort-Object @{ e = { @{ 'エラー' = 0; '警告' = 1; '情報' = 2 }[$_.Level] } }, @{ e = { @('実装に無い', '設計書に無い').IndexOf($_.Item) } }, Seq)
$sorted

if ($OutXlsx) {
    $issueRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $sorted) {
        $row = @{ Values = @($i.Level, $i.Item, $i.Kind, $i.Id, $i.Detail, $i.File, $i.Location) }
        if ($i.Sheet) { $row.Link = @{ Column = 5; Path = $i.Path; Location = ("'{0}'!A{1}" -f $i.Sheet.Replace("'", "''"), $i.Row) } }
        else { $row.Link = @{ Column = 5; Path = $i.Path } }
        $issueRows.Add($row)
    }
    $traceRows = New-Object System.Collections.Generic.List[object]
    foreach ($k in @('Message', 'Screen', 'Table', 'Column')) {
        if ($Kind -notcontains $k) { continue }
        foreach ($d in $defined[$k].Values) {
            $filesUsed = @($d.Hits | ForEach-Object { $_.Relative } | Select-Object -Unique)
            $row = @{ Values = @($kindNames[$k], $d.Id, $d.Name, ('{0} {1} {2} 行目' -f $d.File, $d.Sheet, $d.Row), [double]$d.Hits.Count, [double]$filesUsed.Count, (Get-HitText $d.Hits)) }
            if ($d.Hits.Count -gt 0) { $row.Link = @{ Column = 6; Path = $d.Hits[0].Path } }
            $traceRows.Add($row)
        }
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = '指摘'; Columns = @(@{ Header = '区分'; Width = 7 }, @{ Header = '項目'; Width = 12 }, @{ Header = '種類'; Width = 10 }, @{ Header = 'ID'; Width = 16 }, @{ Header = '内容'; Width = 80 }, @{ Header = 'ファイル'; Width = 36 }, @{ Header = '場所'; Width = 18 }); Rows = $issueRows },
            @{ Name = '対応表'; Columns = @(@{ Header = '種類'; Width = 10 }, @{ Header = 'ID'; Width = 22 }, @{ Header = '名前'; Width = 30 }, @{ Header = '設計書'; Width = 40 }, @{ Header = 'ソースの箇所'; Width = 10 }, @{ Header = 'ファイル'; Width = 8 }, @{ Header = '最初の場所'; Width = 50 }); Rows = $traceRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$parts = @(foreach ($k in @('Message', 'Screen', 'Table', 'Column')) {
        if ($Kind -notcontains $k) { continue }
        $all = $defined[$k].Count
        $used = @($defined[$k].Values | Where-Object { $_.Hits.Count -gt 0 }).Count
        '{0} {1}/{2}' -f $kindNames[$k], $used, $all
    })
Write-Host ('ソース {0} ファイルで、設計書の ID が見つかったもの: {1}' -f $files.Count, ($parts -join '・'))
$notInDesign = @(foreach ($k in @('Message', 'Screen', 'Table')) { if ($undefined[$k].Count -gt 0) { '{0} {1}' -f $kindNames[$k], $undefined[$k].Count } })
if ($notInDesign.Count -gt 0) { Write-Host ('設計書に無い ID: {0}' -f ($notInDesign -join '・')) }
Write-Host ('警告 {0} 件 / 情報 {1} 件' -f @($issues | Where-Object Level -eq '警告').Count, @($issues | Where-Object Level -eq '情報').Count)
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     設計書の ID（メッセージ・画面・テーブル・カラム）とソースコードを突き合わせ、実装に無い ID と、設計書に無い ID を見つけます（Excel 不要）。
#
# .DESCRIPTION
#     設計書から次の ID を読み、ソースコード（とメッセージのリソースなどの設定ファイル）のどこに書いてあるかを探します。
#
#       メッセージ: メッセージ一覧のメッセージ ID（E0001・MSG-E001 など）
#       画面: 画面一覧の画面 ID と、画面設計書のヘッダの画面 ID
#       テーブル・カラム: テーブル定義書のテーブル ID とカラムの物理名（Export-TableDdl.ps1 と同じ見出しで読む）
#
#     次を知らせます。
#
#       実装に無い（警告）: 設計書の ID がソースのどこにも無い（使っていないか、ID の書き間違い）。カラムは情報
#       設計書に無い（警告）: ソースに、設計書の ID と同じ形（E0001 なら英字 E と数字 4 桁）の ID があるが、設計書に無い。
#                         テーブルは、SQL の FROM・JOIN・UPDATE と、INSERT・MERGE などの INTO の後ろの名前
#                         （行をまたいでも見る。.sql と MyBatis などの .xml のほかは、大文字の SQL だけ）
#
#     ID は単語として探します（大文字小文字は区別しない）。テーブル・カラムは、_ を除いた名前（orderNo・TOrder など）でも見つけます。
#     コメント（//・/* */・--・# など。言語ごとの書き方は Measure-CodeSize.ps1 と同じ）は、コメントだけの行も、行の後ろのコメントも見ません。
#
# .PARAMETER DesignPath
#     設計書のフォルダまたはファイル（メッセージ一覧・画面一覧・画面設計書・テーブル定義書。複数可）。
#
# .PARAMETER SourcePath
#     ソースのフォルダまたはファイル（複数可）。Measure-CodeSize.ps1 と同じ拡張子のソースと、.properties・.yml・.json などの設定ファイルを読みます。
#
# .PARAMETER Kind
#     突き合わせる ID の種類（Message / Screen / Table / Column）。既定は Message・Screen・Table。
#
# .PARAMETER Exclude
#     見ないフォルダの名前（既定は .git・node_modules・bin・obj・target・build・dist・out・vendor・.vs・.idea）。
#
# .PARAMETER Headers
#     テーブル定義書の見出しの候補の CSV。既定は tools\ddl\テーブル定義書の見出し.csv（Export-TableDdl.ps1 と同じ）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません（設計書・ソースとも）。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb の設計書は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（指摘と、ID ごとのソースの箇所の対応表）。
#
# .EXAMPLE
#     .\Test-CodeTrace.ps1 D:\設計書 D:\src -OutXlsx D:\設計とソースの突き合わせ.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     ID が文字として書いてあるかを見るだけで、正しく使っているかは見ません。定数や設定から組み立てる ID は見つけられません。
