[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [ValidateSet('Oracle', 'PostgreSQL', 'SQLServer', 'MySQL')]
    [string] $Database = 'Oracle',

    [string[]] $Table,

    [switch] $IncludeDrop,

    [string] $OutFile,

    [ValidateSet('UTF8', 'UTF8BOM', 'ShiftJIS')]
    [string] $Encoding = 'UTF8',

    [string] $CompareCatalog,

    [string] $OutXlsx,

    [string] $TypeMap,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'

if (-not $TypeMap) { $TypeMap = Join-Path $PSScriptRoot 'ddl/型の対応.csv' }
if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'ddl/テーブル定義書の見出し.csv' }
foreach ($name in @('TypeMap', 'Headers', 'OutFile', 'CompareCatalog', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($CompareCatalog -and -not (Test-Path -LiteralPath $CompareCatalog)) { throw "-CompareCatalog の CSV が見つかりません: $CompareCatalog" }

Update-TypeData -TypeName 'TableDdl.Difference' -DefaultDisplayPropertySet Table, Column, Item, Detail -Force -WhatIf:$false

$nfkc = [System.Text.NormalizationForm]::FormKC

# 名前の長さの上限（バイト）。Oracle は 12.2 から 128、それより前は 30
$nameLimit = @{ Oracle = 128; PostgreSQL = 63; SQLServer = 128; MySQL = 64 }[$Database]

# 日時・時刻の秒の小数部の桁数の上限
$fractionLimit = @{ Oracle = 9; PostgreSQL = 6; SQLServer = 7; MySQL = 6 }[$Database]

# どの DB でも予約語になっている（または一部の DB で予約語の）よく使われる語
$reserved = New-Object 'System.Collections.Generic.HashSet[string]'
foreach ($w in @('ACCESS', 'ADD', 'ALL', 'ALTER', 'AND', 'ANY', 'AS', 'ASC', 'AUDIT', 'BETWEEN', 'BY', 'CASE', 'CHAR', 'CHECK', 'COLUMN', 'COMMENT',
        'CONSTRAINT', 'CREATE', 'CURRENT', 'DATE', 'DECIMAL', 'DEFAULT', 'DELETE', 'DESC', 'DISTINCT', 'DROP', 'ELSE', 'END', 'EXISTS', 'FILE', 'FLOAT',
        'FOR', 'FOREIGN', 'FROM', 'GRANT', 'GROUP', 'HAVING', 'IN', 'INDEX', 'INSERT', 'INTEGER', 'INTO', 'IS', 'JOIN', 'KEY', 'LEVEL', 'LIKE', 'LIMIT',
        'LOCK', 'MODE', 'NOT', 'NULL', 'NUMBER', 'OF', 'OFFSET', 'ON', 'OPTION', 'OR', 'ORDER', 'PRIMARY', 'PUBLIC', 'RANGE', 'REFERENCES', 'RESOURCE',
        'ROW', 'ROWID', 'ROWNUM', 'ROWS', 'SELECT', 'SESSION', 'SET', 'SIZE', 'SYSDATE', 'TABLE', 'THEN', 'TIME', 'TIMESTAMP', 'TO', 'TRIGGER', 'UID',
        'UNION', 'UNIQUE', 'UPDATE', 'USER', 'VALUES', 'VARCHAR', 'VARCHAR2', 'VIEW', 'WHEN', 'WHERE', 'WITH')) { [void]$reserved.Add($w) }

# ---------------------------------------------------------------- 設定（見出しの候補・型の対応）


$candidates = Read-TableHeaderCandidates $Headers

function ConvertTo-TypeKey {
    param([string] $Text)
    return ($Text.Normalize($nfkc) -replace '\s', '').ToUpperInvariant()
}

$typeTemplates = @{}
$typeRows = Read-CsvFile $TypeMap @('定義書の型', $Database) '型の対応'
foreach ($r in $typeRows) {
    $k = ConvertTo-TypeKey ([string]$r.'定義書の型')
    if ($k -and ([string]$r.$Database).Trim()) { $typeTemplates[$k] = ([string]$r.$Database).Trim() }
}

# ---------------------------------------------------------------- テーブル定義書を読む


# 同じ文書の古い版・ツールの出力・書き換える前のファイルと、非表示のシート・雛形のシート（差し込み欄のまま）は読まない（知らせる）
$designFiles = Select-DesignFiles -Path $Path -NoRecurse:$NoRecurse -Exclude @($OutXlsx)
$searchArgs = @{ Pattern = '\S'; Path = $designFiles.Paths; Target = 'Value'; SkipHiddenSheets = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = Select-DesignHits @(& $searchScript @searchArgs 6>$null) $designFiles
$books = ConvertTo-DesignBooks $hits

$script:Findings = New-Object System.Collections.Generic.List[object]

function Add-Finding {
    #     テーブル定義書の気になる点。警告として表示し、-OutXlsx にも載せる
    param([string] $Message, $Place)
    $where = ''
    if ($Place) { $where = '{0} {1}!{2}: ' -f $Place.File, $Place.Sheet, $Place.Cell }
    Write-Warning ($where + $Message)
    $script:Findings.Add([pscustomobject]@{ Message = $Message; Place = $Place })
}

$tables = Read-TableDesign -Books $books -Candidates $candidates -OnFinding { param($Message, $Place) Add-Finding $Message $Place }
$tableById = @{}
foreach ($t in $tables) { $tableById[$t.Id.ToUpperInvariant()] = $t }

# -Table で絞り込む（ワイルドカード）
if ($Table) {
    $filtered = New-Object System.Collections.Generic.List[object]
    foreach ($t in $tables) { foreach ($p in $Table) { if ($t.Id -like $p) { $filtered.Add($t); break } } }
    $tables = $filtered
}
if ($tables.Count -eq 0) { Write-Warning 'テーブル定義（物理名とデータ型の見出しのある表）が見つかりませんでした'; return }

# ---------------------------------------------------------------- 型と名前の確かめ


function Expand-TypeTemplate {
    #     型の対応の書き方: {n} は桁数、{s} は小数。[…] は中の値が無ければ省く。値の無い (…) は省く
    param([string] $Template, [string] $N, [string] $S)
    $t = [regex]::Replace($Template, '\[([^\[\]]*)\]', [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        $inner = $m.Groups[1].Value
        if (($inner.Contains('{n}') -and -not $N) -or ($inner.Contains('{s}') -and -not $S)) { return '' }
        return $inner
    })
    $missing = ($t.Contains('{n}') -and -not $N) -or ($t.Contains('{s}') -and -not $S)
    $t = [regex]::Replace($t, '\s*\([^()]*\)', [System.Text.RegularExpressions.MatchEvaluator] {
        param($m)
        if (($m.Value.Contains('{n}') -and -not $N) -or ($m.Value.Contains('{s}') -and -not $S)) { return '' }
        return $m.Value
    })
    return [pscustomobject]@{ Sql = $t.Replace('{n}', $N).Replace('{s}', $S); Missing = $missing }
}

function Get-TypeFamily {
    #     型の種類（DB ごとの名前の違いをそろえる）
    param([string] $Type)
    $b = (([string]$Type) -replace '\(.*$', '').Trim().ToLowerInvariant()
    switch -Regex ($b) {
        '^(varchar2?|nvarchar2?|character varying|national character varying)$' { return '可変長文字' }
        '^(char|nchar|character|bpchar|national character)$' { return '固定長文字' }
        '^(number|numeric|decimal|dec|integer|int|int2|int4|int8|smallint|bigint|tinyint|mediumint)$' { return '数値' }
        '^date$' { return '日付' }
        '^(timestamp|datetime|datetime2|smalldatetime|datetimeoffset)' { return '日時' }
        '^time' { return '時刻' }
        '^(text|clob|nclob|ntext|longtext|mediumtext|tinytext|long)$' { return '長い文字' }
        '^(blob|bytea|varbinary|binary|longblob|mediumblob|image|raw|long raw)$' { return 'バイナリ' }
        '^(boolean|bool|bit)$' { return '真偽' }
    }
    return $b
}

function Test-NameLength {
    #     名前の長さ（DB の上限と、Oracle 12.1 以前の 30 バイト）。定義書の名前と、作る制約名（PK_・FK_）に使う
    param([string] $Name, [string] $What, $Place)
    $bytes = [System.Text.Encoding]::UTF8.GetByteCount($Name)
    if ($bytes -gt $nameLimit) { Add-Finding ('{0}「{1}」が {2} の名前の上限（{3} バイト）を超えています' -f $What, $Name, $Database, $nameLimit) $Place }
    elseif ($Database -eq 'Oracle' -and $bytes -gt 30) { Add-Finding ('{0}「{1}」が 30 バイトを超えています（Oracle 12.1 以前では使えません）' -f $What, $Name) $Place }
}

function Test-Name {
    param([string] $Name, [string] $What, $Place)
    if ($Name -notmatch '^[A-Za-z][A-Za-z0-9_$#]*$') { Add-Finding ('{0}「{1}」に英数字と _ 以外の文字が含まれています' -f $What, $Name) $Place; return }
    if ($reserved.Contains($Name.ToUpperInvariant())) { Add-Finding ('{0}「{1}」は SQL の予約語です' -f $What, $Name) $Place }
    Test-NameLength $Name $What $Place
}

# Oracle・PostgreSQL のインデックス名はスキーマの中で 1 つ（SQL Server・MySQL はテーブルごと）。インデックス名 → テーブル
$indexTables = @{}
foreach ($t in $tables) {
    Test-Name $t.Id 'テーブル名' $t.Place
    foreach ($c in $t.Columns) {
        Test-Name $c.Name 'カラム名' $c.Place
        $k = ConvertTo-TypeKey $c.BaseType
        if ($typeTemplates.ContainsKey($k)) {
            $x = Expand-TypeTemplate $typeTemplates[$k] $c.Length $c.Scale
            $family = Get-TypeFamily $x.Sql
            if ($c.Length -and -not $typeTemplates[$k].Contains('{n}')) {
                # 整数・真偽・日付の桁数（INTEGER の 10 など）は型で大きさが決まるので、使わないことを知らせない
                if (@('数値', '真偽', '日付') -notcontains $family) { Add-Finding ('{0}.{1} の桁数 {2} は、データ型 {3}（{4} の {5}）では使いません' -f $t.Id, $c.Name, $c.Length, $c.BaseType, $Database, $x.Sql) $c.Place }
            }
            elseif ($c.Length -and @('日時', '時刻') -contains $family -and [int]$c.Length -gt $fractionLimit) {
                # 日時・時刻の桁数は秒の小数部の桁数。DB の上限を超える桁数（「yyyy/MM/dd HH:mm:ss」の 19 など）は使わない
                Add-Finding ('{0}.{1} の桁数 {2} は、{3} の秒の小数部の上限（{4} 桁）を超えるため使いません' -f $t.Id, $c.Name, $c.Length, $Database, $fractionLimit) $c.Place
                $x = Expand-TypeTemplate $typeTemplates[$k] '' $c.Scale
            }
            $c.SqlType = $x.Sql
            if ($x.Missing) { Add-Finding ('{0}.{1} のデータ型 {2} に桁数がありません' -f $t.Id, $c.Name, $c.BaseType) $c.Place }
        }
        else {
            $c.SqlType = $c.BaseType
            if ($c.Length -and $c.Scale) { $c.SqlType = '{0}({1},{2})' -f $c.BaseType, $c.Length, $c.Scale } elseif ($c.Length) { $c.SqlType = '{0}({1})' -f $c.BaseType, $c.Length }
            Add-Finding ('{0}.{1} のデータ型「{2}」は型の対応（{3}）にないため、そのまま書きます' -f $t.Id, $c.Name, $c.BaseType, [System.IO.Path]::GetFileName($TypeMap)) $c.Place
        }
    }
    if (@($t.Columns | Where-Object { $null -ne $_.PkOrder }).Count -eq 0) { Add-Finding ('テーブル {0} に主キーがありません' -f $t.Id) $t.Place }
    else { Test-NameLength ('PK_' + $t.Id) '主キーの制約名' $t.Place }
    $names = @{}
    foreach ($c in $t.Columns) { $names[$c.Name.ToUpperInvariant()] = $true }
    foreach ($ix in $t.Indexes) {
        Test-Name $ix.Name 'インデックス名' $ix.Place
        if ($Database -eq 'Oracle' -or $Database -eq 'PostgreSQL') {
            $ixKey = $ix.Name.ToUpperInvariant()
            if ($indexTables.ContainsKey($ixKey)) { Add-Finding ('インデックス名「{0}」はテーブル {1} のインデックスにも使っています（{2} ではスキーマの中で同じインデックス名を使えません）' -f $ix.Name, $indexTables[$ixKey], $Database) $ix.Place }
            else { $indexTables[$ixKey] = $t.Id }
        }
        foreach ($ic in $ix.Columns) {
            $bare = ($ic -replace '\s+(ASC|DESC)$', '')
            if (-not $names.ContainsKey($bare.ToUpperInvariant())) { Add-Finding ('インデックス {0} のカラム {1} がテーブル {2} にありません' -f $ix.Name, $bare, $t.Id) $ix.Place }
        }
    }
}

function Get-ColumnSetKey {
    #     カラムの組み合わせを比べる形（並び順の ASC・DESC を除き、大文字にして名前の順に並べる）
    param([string[]] $Names)
    return (@($Names | ForEach-Object { ($_ -replace '\s+(ASC|DESC)$', '').ToUpperInvariant() } | Sort-Object) -join ',')
}

# 外部キー: 参照先のテーブルごとにまとめる（同じテーブルの同じカラムを 2 回参照するときは分ける）
$allById = @{}
foreach ($t in $tables) { $allById[$t.Id.ToUpperInvariant()] = $t }
foreach ($t in $tableById.Values) { if (-not $allById.ContainsKey($t.Id.ToUpperInvariant())) { $allById[$t.Id.ToUpperInvariant()] = $t } }
$foreignKeys = New-Object System.Collections.Generic.List[object]
foreach ($t in $tables) {
    $groups = [ordered]@{}
    foreach ($c in $t.Columns) {
        if (-not $c.FkTable) { continue }
        $parent = $allById[$c.FkTable.ToUpperInvariant()]
        if (-not $parent) { Add-Finding ('{0}.{1} の外部キーの参照先テーブル {2} がテーブル定義書にありません' -f $t.Id, $c.Name, $c.FkTable) $c.Place; continue }
        $pc = @($parent.Columns | Where-Object { $_.Name -eq $c.FkColumn })[0]
        if (-not $pc) { Add-Finding ('{0}.{1} の外部キーの参照先 {2}.{3} がテーブル定義書にありません' -f $t.Id, $c.Name, $parent.Id, $c.FkColumn) $c.Place; continue }
        if ($pc.SqlType -and $c.SqlType -and $pc.SqlType -ne $c.SqlType) {
            Add-Finding ('{0}.{1}（{2}）と外部キーの参照先 {3}.{4}（{5}）の型・桁数が違います' -f $t.Id, $c.Name, $c.SqlType, $parent.Id, $pc.Name, $pc.SqlType) $c.Place
        }
        $gk = $parent.Id.ToUpperInvariant()
        while ($groups.Contains($gk) -and @($groups[$gk].ParentColumns) -contains $pc.Name) { $gk += '+' }
        if (-not $groups.Contains($gk)) { $groups[$gk] = [pscustomobject]@{ Parent = $parent.Id; Columns = (New-Object System.Collections.Generic.List[string]); ParentColumns = (New-Object System.Collections.Generic.List[string]); Place = $c.Place } }
        $groups[$gk].Columns.Add($c.Name)
        $groups[$gk].ParentColumns.Add($pc.Name)
    }
    $used = @{}
    foreach ($g in $groups.Values) {
        $fkName = 'FK_{0}_{1}' -f $t.Id, $g.Parent
        $i = 1
        while ($used.ContainsKey($fkName.ToUpperInvariant())) { $i++; $fkName = 'FK_{0}_{1}_{2}' -f $t.Id, $g.Parent, $i }
        $used[$fkName.ToUpperInvariant()] = $true
        Test-NameLength $fkName '外部キーの制約名' $t.Place
        # 参照先は、親の主キーかユニークインデックスのカラムのすべて（主キーの一部だけなどでは、DB で外部キーを作れない）
        $parent = $allById[$g.Parent.ToUpperInvariant()]
        $want = Get-ColumnSetKey $g.ParentColumns
        $isKey = (Get-ColumnSetKey @($parent.Columns | Where-Object { $null -ne $_.PkOrder } | ForEach-Object { $_.Name })) -eq $want
        foreach ($ix in $parent.Indexes) { if ($ix.Unique -and (Get-ColumnSetKey $ix.Columns) -eq $want) { $isKey = $true } }
        if (-not $isKey) { Add-Finding ('{0} ({1}) の外部キーの参照先 {2} ({3}) が、{2} の主キーでもユニークインデックスでもありません' -f $t.Id, ($g.Columns -join ', '), $g.Parent, ($g.ParentColumns -join ', ')) $g.Place }
        $foreignKeys.Add([pscustomobject]@{ Name = $fkName; Table = $t.Id; Columns = $g.Columns.ToArray(); Parent = $g.Parent; ParentColumns = $g.ParentColumns.ToArray() })
    }
}

# 参照先のテーブルが先に来る順（循環していればテーブル定義書の順）
$ordered = New-Object System.Collections.Generic.List[object]
$done = @{}
$pending = New-Object System.Collections.Generic.List[object]
foreach ($t in $tables) { $pending.Add($t) }
while ($pending.Count -gt 0) {
    $next = $null
    foreach ($t in $pending) {
        $ready = $true
        foreach ($fk in $foreignKeys) {
            if ($fk.Table -ne $t.Id -or $fk.Parent -eq $t.Id) { continue }
            if (@($pending | Where-Object { $_.Id -eq $fk.Parent }).Count -gt 0) { $ready = $false; break }
        }
        if ($ready) { $next = $t; break }
    }
    if (-not $next) { $next = $pending[0] }
    $ordered.Add($next)
    [void]$pending.Remove($next)
}

# ---------------------------------------------------------------- DDL


function ConvertTo-SqlString {
    #     コメントの文字列。セル内の改行（↵）は空白にして 1 行にする。MySQL は \ もエスケープの文字なので \\ にする
    param([string] $Text)
    $s = ([string]$Text -replace '\s*↵\s*', ' ').Trim().Replace("'", "''")
    if ($Database -eq 'MySQL') { $s = $s.Replace('\', '\\') }
    return "'" + $s + "'"
}

$lines = New-Object System.Collections.Generic.List[string]
$sources = @($ordered | ForEach-Object { $_.File } | Select-Object -Unique)
$lines.Add(('-- テーブル定義書から作成（Export-TableDdl.ps1、対象 DB: {0}）' -f $Database))
$lines.Add(('-- 元のファイル: {0}' -f ($sources -join '、')))
$lines.Add('')
if ($IncludeDrop) {
    $lines.Add('-- 削除（参照しているテーブルから）')
    for ($i = $ordered.Count - 1; $i -ge 0; $i--) {
        $id = $ordered[$i].Id
        switch ($Database) {
            'Oracle' { $lines.Add("DROP TABLE $id CASCADE CONSTRAINTS;") }
            'PostgreSQL' { $lines.Add("DROP TABLE IF EXISTS $id CASCADE;") }
            default { $lines.Add("DROP TABLE IF EXISTS $id;") }
        }
    }
    $lines.Add('')
}
foreach ($t in $ordered) {
    $title = $t.Id
    if ($t.Name) { $title = '{0} {1}' -f $t.Id, $t.Name }
    $lines.Add(('-- {0}（{1} {2}）' -f $title, $t.File, $t.Sheet))
    $lines.Add(('CREATE TABLE {0} (' -f $t.Id))
    $nameWidth = ($t.Columns | ForEach-Object { $_.Name.Length } | Measure-Object -Maximum).Maximum
    $typeWidth = ($t.Columns | ForEach-Object { $_.SqlType.Length } | Measure-Object -Maximum).Maximum
    $items = New-Object System.Collections.Generic.List[string]
    foreach ($c in $t.Columns) {
        $line = '    {0} {1}' -f $c.Name.PadRight($nameWidth), $c.SqlType.PadRight($typeWidth)
        if ($c.Default) { $line += ' DEFAULT ' + $c.Default }
        if ($c.NotNull) { $line += ' NOT NULL' }
        if ($Database -eq 'MySQL' -and $c.Logical) { $line += ' COMMENT ' + (ConvertTo-SqlString $c.Logical) }
        $items.Add($line.TrimEnd())
    }
    $pk = @($t.Columns | Where-Object { $null -ne $_.PkOrder } | Sort-Object PkOrder | ForEach-Object { $_.Name })
    if ($pk.Count -gt 0) { $items.Add(('    CONSTRAINT PK_{0} PRIMARY KEY ({1})' -f $t.Id, ($pk -join ', '))) }
    for ($i = 0; $i -lt $items.Count; $i++) {
        $suffix = ','
        if ($i -eq $items.Count - 1) { $suffix = '' }
        $lines.Add($items[$i] + $suffix)
    }
    if ($Database -eq 'MySQL' -and $t.Name) { $lines.Add(') COMMENT = ' + (ConvertTo-SqlString $t.Name) + ';') } else { $lines.Add(');') }
    # コメント（論理名）
    switch ($Database) {
        { $_ -eq 'Oracle' -or $_ -eq 'PostgreSQL' } {
            if ($t.Name) { $lines.Add(('COMMENT ON TABLE {0} IS {1};' -f $t.Id, (ConvertTo-SqlString $t.Name))) }
            foreach ($c in $t.Columns) { if ($c.Logical) { $lines.Add(('COMMENT ON COLUMN {0}.{1} IS {2};' -f $t.Id, $c.Name, (ConvertTo-SqlString $c.Logical))) } }
        }
        'SQLServer' {
            $base = "EXEC sys.sp_addextendedproperty @name = N'MS_Description', @value = N{0}, @level0type = N'SCHEMA', @level0name = N'dbo', @level1type = N'TABLE', @level1name = N'{1}'"
            if ($t.Name) { $lines.Add(($base -f (ConvertTo-SqlString $t.Name), $t.Id) + ';') }
            foreach ($c in $t.Columns) {
                if ($c.Logical) { $lines.Add(($base -f (ConvertTo-SqlString $c.Logical), $t.Id) + (", @level2type = N'COLUMN', @level2name = N'{0}';" -f $c.Name)) }
            }
        }
    }
    foreach ($ix in $t.Indexes) {
        $bare = @($ix.Columns | ForEach-Object { ($_ -replace '\s+(ASC|DESC)$', '').ToUpperInvariant() })
        $pkUpper = @($pk | ForEach-Object { $_.ToUpperInvariant() })
        if (($bare -join ',') -eq ($pkUpper -join ',')) { $lines.Add(('-- {0} は主キーと同じカラムのため作りません' -f $ix.Name)); continue }
        $unique = ''
        if ($ix.Unique) { $unique = 'UNIQUE ' }
        $lines.Add(('CREATE {0}INDEX {1} ON {2} ({3});' -f $unique, $ix.Name, $t.Id, ($ix.Columns -join ', ')))
    }
    $lines.Add('')
}
if ($foreignKeys.Count -gt 0) {
    $lines.Add('-- 外部キー')
    foreach ($t in $ordered) {
        foreach ($fk in @($foreignKeys | Where-Object { $_.Table -eq $t.Id })) {
            $lines.Add(('ALTER TABLE {0} ADD CONSTRAINT {1} FOREIGN KEY ({2}) REFERENCES {3} ({4});' -f $fk.Table, $fk.Name, ($fk.Columns -join ', '), $fk.Parent, ($fk.ParentColumns -join ', ')))
        }
    }
    $lines.Add('')
}
$ddl = ($lines -join "`r`n")

if ($OutFile) {
    switch ($Encoding) {
        'UTF8' { $enc = New-Object System.Text.UTF8Encoding($false) }
        'UTF8BOM' { $enc = New-Object System.Text.UTF8Encoding($true) }
        'ShiftJIS' { $enc = [System.Text.Encoding]::GetEncoding(932) }
    }
    [System.IO.File]::WriteAllText($OutFile, $ddl, $enc)
}
elseif (-not $CompareCatalog) { $ddl }

# ---------------------------------------------------------------- DB のカタログと比べる


function Get-CatalogValue {
    #     カタログの CSV の値（列名の候補のうち最初にあるもの）
    param($Row, [string[]] $Names)
    foreach ($n in $Names) { $p = $Row.PSObject.Properties[$n]; if ($p) { return ([string]$p.Value).Trim() } }
    return ''
}

$differences = New-Object System.Collections.Generic.List[object]
if ($CompareCatalog) {
    $catalog = Read-CsvFile $CompareCatalog @() 'カタログの CSV'
    if ($catalog.Count -gt 0) {
        $names = @($catalog[0].PSObject.Properties | ForEach-Object { $_.Name.ToUpperInvariant() })
        foreach ($need in @('TABLE_NAME', 'COLUMN_NAME', 'DATA_TYPE')) { if ($names -notcontains $need) { throw "カタログの CSV に「$need」の列がありません: $CompareCatalog" } }
    }
    $db = [ordered]@{}
    foreach ($r in $catalog) {
        $tn = (Get-CatalogValue $r @('TABLE_NAME')).ToUpperInvariant()
        if (-not $tn) { continue }
        if (-not $db.Contains($tn)) { $db[$tn] = [ordered]@{} }
        $db[$tn][(Get-CatalogValue $r @('COLUMN_NAME')).ToUpperInvariant()] = $r
    }
    $docTables = @{}
    foreach ($t in $ordered) { $docTables[$t.Id.ToUpperInvariant()] = $true }

    function Add-Difference {
        param([string] $TableName, [string] $ColumnName, [string] $Item, [string] $Doc, [string] $Db, [string] $Detail, $Place)
        $d = [pscustomobject]@{ Table = $TableName; Column = $ColumnName; Item = $Item; Definition = $Doc; Database = $Db; Detail = $Detail; File = ''; Location = ''; Path = ''; Sheet = ''; Cell = '' }
        if ($Place) { $d.File = $Place.File; $d.Path = $Place.Path; $d.Sheet = $Place.Sheet; $d.Cell = $Place.Cell; $d.Location = '{0}!{1}' -f $Place.Sheet, $Place.Cell }
        $d.PSObject.TypeNames.Insert(0, 'TableDdl.Difference')
        $differences.Add($d)
    }

    foreach ($t in $ordered) {
        $tk = $t.Id.ToUpperInvariant()
        if (-not $db.Contains($tk)) { Add-Difference $t.Id '' 'テーブルなし' $t.Id '' ('DB にテーブル {0} がありません' -f $t.Id) $t.Place; continue }
        $dbCols = $db[$tk]
        foreach ($c in $t.Columns) {
            $ck = $c.Name.ToUpperInvariant()
            if (-not $dbCols.Contains($ck)) { Add-Difference $t.Id $c.Name 'カラムなし' $c.SqlType '' ('DB にカラム {0}.{1} がありません' -f $t.Id, $c.Name) $c.Place; continue }
            $r = $dbCols[$ck]
            $dbType = Get-CatalogValue $r @('DATA_TYPE')
            $famDoc = Get-TypeFamily $c.SqlType
            $famDb = Get-TypeFamily $dbType
            # 型の対応に桁数（{n}）・小数（{s}）の無い型（INTEGER など）は、その値を DB と比べない。型の対応に無い型は、桁数・小数をそのまま書いている
            $template = $typeTemplates[(ConvertTo-TypeKey $c.BaseType)]
            $usesLength = (-not $template) -or $template.Contains('{n}')
            $usesScale = (-not $template) -or $template.Contains('{s}')
            if ($famDoc -ne $famDb) {
                Add-Difference $t.Id $c.Name '型' $c.SqlType $dbType ('{0}.{1} の型が違います（定義書: {2} / DB: {3}）' -f $t.Id, $c.Name, $c.SqlType, $dbType) $c.Place
            }
            elseif ($c.Length -and $usesLength) {
                if ($famDoc -eq '可変長文字' -or $famDoc -eq '固定長文字') {
                    $len = Get-CatalogValue $r @('CHARACTER_MAXIMUM_LENGTH', 'CHAR_LENGTH', 'CHAR_COL_DECL_LENGTH')
                    if ($len -match '^\d+$' -and [int]$len -ne [int]$c.Length) {
                        Add-Difference $t.Id $c.Name '桁数' $c.Length $len ('{0}.{1} の桁数が違います（定義書: {2} / DB: {3}）' -f $t.Id, $c.Name, $c.Length, $len) $c.Place
                    }
                }
                elseif ($famDoc -eq '数値') {
                    $prec = Get-CatalogValue $r @('NUMERIC_PRECISION', 'DATA_PRECISION')
                    $scale = Get-CatalogValue $r @('NUMERIC_SCALE', 'DATA_SCALE')
                    if (-not $scale) { $scale = '0' }
                    $docScale = $c.Scale
                    if (-not $docScale) { $docScale = '0' }
                    if ($prec -match '^\d+$' -and ([int]$prec -ne [int]$c.Length -or ($usesScale -and [int]$scale -ne [int]$docScale))) {
                        Add-Difference $t.Id $c.Name '桁数' ('{0},{1}' -f $c.Length, $docScale) ('{0},{1}' -f $prec, $scale) ('{0}.{1} の桁数が違います（定義書: {2},{3} / DB: {4},{5}）' -f $t.Id, $c.Name, $c.Length, $docScale, $prec, $scale) $c.Place
                    }
                }
            }
            $nullable = (Get-CatalogValue $r @('IS_NULLABLE', 'NULLABLE')).ToUpperInvariant()
            if ($nullable) {
                $dbNotNull = @('NO', 'N', 'FALSE', '0') -contains $nullable
                if ($dbNotNull -ne $c.NotNull) {
                    $docText = 'NULL 可'; if ($c.NotNull) { $docText = 'NOT NULL' }
                    $dbText = 'NULL 可'; if ($dbNotNull) { $dbText = 'NOT NULL' }
                    Add-Difference $t.Id $c.Name 'NULL' $docText $dbText ('{0}.{1} の NULL の可否が違います（定義書: {2} / DB: {3}）' -f $t.Id, $c.Name, $docText, $dbText) $c.Place
                }
            }
        }
        $docCols = @{}
        foreach ($c in $t.Columns) { $docCols[$c.Name.ToUpperInvariant()] = $true }
        foreach ($ck in $dbCols.Keys) {
            if (-not $docCols.ContainsKey($ck)) {
                $colName = Get-CatalogValue $dbCols[$ck] @('COLUMN_NAME')
                Add-Difference $t.Id $colName '定義書にないカラム' '' (Get-CatalogValue $dbCols[$ck] @('DATA_TYPE')) ('テーブル定義書の {0} にカラム {1} がありません' -f $t.Id, $colName) $t.Place
            }
        }
    }
    if (-not $Table) {
        foreach ($tk in $db.Keys) {
            if (-not $docTables.ContainsKey($tk)) {
                $tn = Get-CatalogValue @($db[$tk].Values)[0] @('TABLE_NAME')
                Add-Difference $tn '' '定義書にないテーブル' '' $tn ('テーブル定義書にテーブル {0} がありません' -f $tn) $null
            }
        }
    }
    $differences
}

# ---------------------------------------------------------------- 結果


if ($OutXlsx) {
    $sheets = New-Object System.Collections.Generic.List[object]
    if ($CompareCatalog) {
        $rows = New-Object System.Collections.Generic.List[object]
        foreach ($d in $differences) {
            $row = @{ Values = @($d.Table, $d.Column, $d.Item, $d.Definition, $d.Database, $d.Detail, $d.File, $d.Location) }
            if ($d.Sheet) { $row.Link = @{ Column = 7; Path = $d.Path; Location = ("'{0}'!{1}" -f $d.Sheet.Replace("'", "''"), $d.Cell) } }
            $rows.Add($row)
        }
        $sheets.Add(@{ Name = 'DB との差異'; Columns = @(@{ Header = 'テーブル'; Width = 20 }, @{ Header = 'カラム'; Width = 20 }, @{ Header = '項目'; Width = 14 }, @{ Header = '定義書'; Width = 18 }, @{ Header = 'DB'; Width = 18 }, @{ Header = '内容'; Width = 60 }, @{ Header = 'ファイル'; Width = 30 }, @{ Header = '場所'; Width = 20 }); Rows = $rows })
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $script:Findings) {
        $file = ''; $loc = ''
        if ($f.Place) { $file = $f.Place.File; $loc = '{0}!{1}' -f $f.Place.Sheet, $f.Place.Cell }
        $row = @{ Values = @($f.Message, $file, $loc) }
        if ($f.Place) { $row.Link = @{ Column = 2; Path = $f.Place.Path; Location = ("'{0}'!{1}" -f $f.Place.Sheet.Replace("'", "''"), $f.Place.Cell) } }
        $rows.Add($row)
    }
    $sheets.Add(@{ Name = '定義書の指摘'; Columns = @(@{ Header = '内容'; Width = 80 }, @{ Header = 'ファイル'; Width = 30 }, @{ Header = '場所'; Width = 20 }); Rows = $rows })
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $ordered) {
        $pkText = @($t.Columns | Where-Object { $null -ne $_.PkOrder } | Sort-Object PkOrder | ForEach-Object { $_.Name }) -join ', '
        $rows.Add(@{ Values = @($t.Id, $t.Name, [double]$t.Columns.Count, $pkText, [double]$t.Indexes.Count, [double]@($foreignKeys | Where-Object { $_.Table -eq $t.Id }).Count, $t.File, ('{0}!{1}' -f $t.Place.Sheet, $t.Place.Cell)); Link = @{ Column = 0; Path = $t.Place.Path; Location = ("'{0}'!{1}" -f $t.Place.Sheet.Replace("'", "''"), $t.Place.Cell) } })
    }
    $sheets.Add(@{ Name = 'テーブル'; Columns = @(@{ Header = 'テーブル'; Width = 20 }, @{ Header = '名前'; Width = 20 }, @{ Header = 'カラム'; Width = 8 }, @{ Header = '主キー'; Width = 30 }, @{ Header = 'インデックス'; Width = 10 }, @{ Header = '外部キー'; Width = 8 }, @{ Header = 'ファイル'; Width = 30 }, @{ Header = '場所'; Width = 20 }); Rows = $rows })
    try { Export-Workbook -FilePath $OutXlsx -Sheets $sheets.ToArray() }
    catch { Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message); $OutXlsx = $null }
}

Write-Host ''
$colCount = ($ordered | ForEach-Object { $_.Columns.Count } | Measure-Object -Sum).Sum
Write-Host ('テーブル {0}・カラム {1}・インデックス {2}・外部キー {3}（対象 DB: {4}、定義書の指摘 {5} 件）' -f $ordered.Count, $colCount, ($ordered | ForEach-Object { $_.Indexes.Count } | Measure-Object -Sum).Sum, $foreignKeys.Count, $Database, $script:Findings.Count)
if ($OutFile) { Write-Host "DDL を保存しました: $OutFile" }
if ($CompareCatalog) { Write-Host ('DB との差異 {0} 件' -f $differences.Count) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テーブル定義書（Excel）から DDL（CREATE TABLE など）を作り、DB のカタログとの食い違いも確かめます（Excel 不要）。
#
# .DESCRIPTION
#     テーブル定義書のシートごとに、物理名とデータ型の見出しのある表をカラム定義として読み、次の DDL を作ります。
#
#       CREATE TABLE    データ型（型の対応で DB の型に変換）・DEFAULT・NOT NULL・主キー
#       コメント        テーブル名・論理名（Oracle / PostgreSQL は COMMENT ON、SQL Server は拡張プロパティ、MySQL は COMMENT）
#       CREATE INDEX    インデックス定義の表（主キーと同じカラムのものは作らない）
#       外部キー        外部キー参照先（テーブルID.カラム）から ALTER TABLE … ADD CONSTRAINT
#
#     テーブルは参照先が先に来る順に並べます。あわせて、名前の文字・長さ（作る制約名の長さも）・予約語、未知の型、桁数の抜けと使わない桁数、
#     主キーの無いテーブル、外部キーの参照先の有無・型の違い・主キーかユニークインデックスか、インデックスのカラムの有無と名前の重なりを確かめて警告します。
#
#     -CompareCatalog に DB のカタログ（列の一覧）の CSV を渡すと、DDL の代わりに定義書との差異を出力します。
#
# .PARAMETER Path
#     テーブル定義書のフォルダまたはファイル（複数可・ワイルドカード可）。既定はカレントフォルダ。
#     同じ文書の古い版・ツールが書き出したファイル・書き換える前のファイル（-BackupDir）と、非表示のシート・雛形のシート（ID が {{テーブルID}} などのまま）は読みません（読まなかったものは警告で知らせます）。
#
# .PARAMETER Database
#     対象の DB（Oracle / PostgreSQL / SQLServer / MySQL）。既定は Oracle。
#
# .PARAMETER Table
#     DDL を作るテーブルのテーブル ID（複数可・ワイルドカード可）。例: 'T_ORDER*'
#
# .PARAMETER IncludeDrop
#     先頭に DROP TABLE を付けます（参照しているテーブルから削除する順）。
#
# .PARAMETER OutFile
#     DDL を保存するファイル。指定しないときは DDL を出力します。
#
# .PARAMETER Encoding
#     -OutFile の文字コード（UTF8 / UTF8BOM / ShiftJIS）。既定は UTF8（BOM なし）。
#
# .PARAMETER CompareCatalog
#     DB のカタログの CSV（TABLE_NAME・COLUMN_NAME・DATA_TYPE・CHARACTER_MAXIMUM_LENGTH・NUMERIC_PRECISION・NUMERIC_SCALE・IS_NULLABLE）。
#     テーブル・カラムの有無、型・桁数・NULL の可否を比べた差異を出力します。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（DB との差異・定義書の指摘・テーブルの一覧。場所をクリックすると定義書の該当セルが開きます）。
#
# .PARAMETER TypeMap
#     型の対応の CSV（定義書の型と、DB ごとの型）。既定は tools\ddl\型の対応.csv。
#
# .PARAMETER Headers
#     テーブル定義書の見出しの CSV（項目と、見出しの候補）。既定は tools\ddl\テーブル定義書の見出し.csv。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\Export-TableDdl.ps1 D:\設計書\テーブル定義書.xlsx -Database PostgreSQL -OutFile D:\create_tables.sql
#
# .EXAMPLE
#     # DB の列の一覧（カタログ）と比べる
#     .\Export-TableDdl.ps1 D:\設計書 -Database Oracle -CompareCatalog D:\columns.csv -OutXlsx D:\DB差異.xlsx
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     作った DDL は、実行する前に内容を確かめてください。
