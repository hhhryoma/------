[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [Parameter(Mandatory = $true)]
    [string] $OutDir,

    [string[]] $Table,

    [ValidateRange(1, 1000)]
    [int] $Rows = 3,

    [ValidateSet('Oracle', 'PostgreSQL', 'SQLServer', 'MySQL')]
    [string] $Database = 'Oracle',

    [datetime] $BaseDate = (Get-Date).Date,

    [switch] $NoBoundary,

    [switch] $IncludeDelete,

    [ValidateSet('UTF8', 'UTF8BOM', 'ShiftJIS')]
    [string] $Encoding = 'UTF8BOM',

    [string] $OutXlsx,

    [string] $Headers,

    [switch] $NoRecurse,

    [switch] $NoExcel
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking

if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'ddl/テーブル定義書の見出し.csv' }
foreach ($name in @('Headers', 'OutDir', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }

$nfkc = [System.Text.NormalizationForm]::FormKC
# 日付・日時は、Windows の地域の設定（和暦など）によらず西暦で書く
$invariant = [System.Globalization.CultureInfo]::InvariantCulture
switch ($Encoding) {
    'UTF8' { $enc = New-Object System.Text.UTF8Encoding($false) }
    'UTF8BOM' { $enc = New-Object System.Text.UTF8Encoding($true) }
    'ShiftJIS' { $enc = [System.Text.Encoding]::GetEncoding(932) }
}
$candidates = Read-TableHeaderCandidates $Headers

# ---------------------------------------------------------------- テーブル定義書・コード定義書を読む


$searchArgs = @{ Pattern = '\S'; Path = $Path; Target = 'Value' }
if ($NoRecurse) { $searchArgs['NoRecurse'] = $true }
if ($NoExcel) { $searchArgs['NoExcel'] = $true }
$hits = @(& (Join-Path $PSScriptRoot 'Search-Excel.ps1') @searchArgs 6>$null | Where-Object { -not ($OutXlsx -and [string]::Equals($_.Path, $OutXlsx, [System.StringComparison]::OrdinalIgnoreCase)) })
$books = ConvertTo-DesignBooks $hits
$tables = Read-TableDesign -Books $books -Candidates $candidates -OnFinding { param($Message, $Place) Write-Warning ('{0} {1}!{2}: {3}' -f $Place.File, $Place.Sheet, $Place.Cell, $Message) }
if ($tables.Count -eq 0) { Write-Warning 'テーブル定義（物理名とデータ型の見出しのある表）が見つかりませんでした'; return }
$tableById = @{}
foreach ($t in $tables) { $tableById[$t.Id.ToUpperInvariant()] = $t }

# コード定義書（コード区分名とコード値の見出しのある表）: コード区分名・コード区分ID（比べる形）→ コード値の一覧（表の順）。
# 区分名を最初の行だけに書いた表は、上の行の区分のコード値とみなす
$codes = @{}
foreach ($book in $books.Values) {
    foreach ($sheetName in $book.Sheets.Keys) {
        foreach ($ct in (Find-DesignTables $book.Sheets[$sheetName] @('コード区分名', 'コード値') 'コード値')) {
            $lastName = ''; $lastId = ''
            foreach ($dr in $ct.Rows) {
                $codeName = ([string]$dr.Cells[$ct.ColumnOf['コード区分名']]).Trim()
                if ($codeName) { $lastName = $codeName } else { $codeName = $lastName }
                $codeId = ''
                if ($ct.ColumnOf.ContainsKey('コード区分ID')) {
                    $codeId = ([string]$dr.Cells[$ct.ColumnOf['コード区分ID']]).Trim()
                    if ($codeId) { $lastId = $codeId } else { $codeId = $lastId }
                }
                $value = ([string]$dr.Cells[$ct.ColumnOf['コード値']]).Trim()
                if (-not $codeName -or -not $value) { continue }
                foreach ($k in @($codeName, $codeId)) {
                    $key = (ConvertTo-Label $k).ToUpperInvariant()
                    if (-not $key) { continue }
                    if (-not $codes.ContainsKey($key)) { $codes[$key] = New-Object System.Collections.Generic.List[string] }
                    if (-not $codes[$key].Contains($value)) { $codes[$key].Add($value) }
                }
            }
        }
    }
}

# -Table で絞り込む（ワイルドカード）。外部キーの参照先のテーブルも作る（無いと外部キーの制約で登録できない）
$selected = New-Object System.Collections.Generic.List[object]
$addedParents = New-Object System.Collections.Generic.List[string]
if ($Table) {
    foreach ($t in $tables) { foreach ($p in $Table) { if ($t.Id -like $p) { $selected.Add($t); break } } }
    if ($selected.Count -eq 0) { Write-Warning ('-Table に合うテーブルがありません: {0}' -f ($Table -join ', ')); return }
    for ($i = 0; $i -lt $selected.Count; $i++) {
        foreach ($c in $selected[$i].Columns) {
            if (-not $c.FkTable) { continue }
            $parent = $tableById[$c.FkTable.ToUpperInvariant()]
            if ($parent -and -not $selected.Contains($parent)) { $selected.Add($parent); $addedParents.Add($parent.Id) }
        }
    }
}
else { foreach ($t in $tables) { $selected.Add($t) } }

# 参照先のテーブルから順に作る（同じ順に登録できる）。自分を参照する外部キーは見ない
$ordered = New-Object System.Collections.Generic.List[object]
$pending = New-Object System.Collections.Generic.List[object]
foreach ($t in $selected) { $pending.Add($t) }
while ($pending.Count -gt 0) {
    $next = $null
    foreach ($t in $pending) {
        $ready = $true
        foreach ($c in $t.Columns) {
            if (-not $c.FkTable -or [string]::Equals($c.FkTable, $t.Id, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
            $parent = $tableById[$c.FkTable.ToUpperInvariant()]
            if ($parent -and $pending.Contains($parent)) { $ready = $false; break }
        }
        if ($ready) { $next = $t; break }
    }
    if (-not $next) {
        $next = $pending[0]
        Write-Warning ('外部キーが循環しているため、{0} から作ります（外部キーの制約を外してから登録してください）' -f $next.Id)
    }
    $ordered.Add($next)
    [void]$pending.Remove($next)
}

# ---------------------------------------------------------------- 値を作る


function Get-ValueKind {
    #     データ型の種類（Char: 固定長の文字、Text: 可変長の文字、Decimal・Integer・Float: 数値、Date・Timestamp・Time、Boolean、Binary）
    param($Column)
    $t = ([string]$Column.BaseType).Normalize($nfkc).ToUpperInvariant() -replace '\s', ''
    switch -Regex ($t) {
        '^(CHAR|NCHAR|CHARACTER|BPCHAR|固定長文字列?)$' { return 'Char' }
        '^(VARCHAR2?|NVARCHAR2?|CHARACTERVARYING|STRING|文字列?|可変長文字列?|TEXT|NTEXT|N?CLOB|LONGTEXT|MEDIUMTEXT|TINYTEXT)$' { return 'Text' }
        '^(NUMBER|NUMERIC|DECIMAL|DEC|数値)$' { return 'Decimal' }
        '^(INTEGER|INT|INT2|INT4|INT8|SMALLINT|BIGINT|TINYINT|MEDIUMINT|整数)$' { return 'Integer' }
        '^(FLOAT|DOUBLE|DOUBLEPRECISION|REAL|BINARY_FLOAT|BINARY_DOUBLE)$' { return 'Float' }
        '^(DATE|日付)$' { return 'Date' }
        '^(TIMESTAMP|DATETIME2?|SMALLDATETIME|日時)' { return 'Timestamp' }
        '^(TIME|時刻)$' { return 'Time' }
        '^(BOOLEAN|BOOL|BIT)$' { return 'Boolean' }
        '^(BLOB|BYTEA|BINARY|VARBINARY|RAW|LONGRAW|IMAGE|LONGBLOB)$' { return 'Binary' }
    }
    return 'Text'
}

function Get-CodeValues {
    #     コード値の一覧（論理名がコード区分名と同じカラムか、備考にコード区分 ID・区分名を書いたカラム）。無ければ $null
    param($Column)
    $logical = (ConvertTo-Label $Column.Logical).ToUpperInvariant()
    if ($logical -and $codes.ContainsKey($logical)) { return , $codes[$logical].ToArray() }
    $note = (ConvertTo-Label $Column.Note).ToUpperInvariant()
    if (-not $note) { return $null }
    # 長い区分名から（「顧客区分」と「区分」のように、ほかの区分名を含む名前があるため）
    foreach ($k in @($codes.Keys | Sort-Object @{ e = { $_.Length }; Descending = $true }, @{ e = { $_ } })) {
        if ($k.Length -ge 2 -and $note.Contains($k)) { return , $codes[$k].ToArray() }
    }
    return $null
}

function Get-IntegerMax {
    param($Column)
    switch -Regex (([string]$Column.BaseType).ToUpperInvariant()) {
        '^SMALLINT|^INT2' { return '32767' }
        '^BIGINT|^INT8' { return '9223372036854775807' }
        '^TINYINT' { return '127' }
    }
    return '2147483647'
}

function Get-FitNumber {
    #     番号 N を、桁に入る 1 から Capacity までにする（入らない番号は 1 から数え直す。NUMBER(4,3) の 10 行目は 1）
    param([long] $N, [long] $Capacity)
    if ($Capacity -le 0 -or $N -le $Capacity) { return $N }
    return (($N - 1) % $Capacity) + 1
}

function Get-DigitsCapacity {
    #     Digits 桁に入るいちばん大きい数（9・99・999…）
    param([int] $Digits)
    return [long]([math]::Pow(10, [math]::Min($Digits, 18))) - 1
}

function Get-NewValue {
    #     カラムの値（参照先のテーブルを見ない値）。Kind: Normal（標準）・Max（最大）・Min（最小）・Null（NULL の行）。Seq: 標準の行の番号。
    #     番号が桁に入らないとき（NUMBER(4,3) の 10 行目・CHAR(1) の主キーの 10 行目など）は、1 から数え直す
    param($Column, [string] $Kind, [int] $Seq)
    $vk = Get-ValueKind $Column
    $len = 0
    if ([string]$Column.Length -match '^\d+$') { $len = [int]$Column.Length }
    $scale = 0
    if ([string]$Column.Scale -match '^\d+$') { $scale = [int]$Column.Scale }
    $isPk = $null -ne $Column.PkOrder
    if ($Kind -eq 'Null' -and -not $Column.NotNull) { return $null }
    # NULL の行の必須のカラムは、標準の次の番号の値（主キー・一意のカラムが重ならないように）
    $n = $Seq
    if ($Kind -eq 'Null') { $n = $Rows + 1 }
    if (-not $isPk) {
        $codeList = Get-CodeValues $Column
        if ($codeList) {
            switch ($Kind) { 'Max' { return $codeList[$codeList.Count - 1] } 'Min' { return $codeList[0] } 'Null' { return $codeList[0] } default { return $codeList[($n - 1) % $codeList.Count] } }
        }
        if ($Column.Logical -match 'フラグ$' -or $Column.Name -match '_(FLG|FLAG)$') {
            switch ($Kind) { 'Max' { return '1' } 'Min' { return '0' } 'Null' { return '0' } default { return [string](($n - 1) % 2) } }
        }
    }
    switch ($vk) {
        'Decimal' {
            # 整数部の無い型（NUMBER(3,3) など）は 0.xxx にする。最大は 0.999（小数が桁数より大きい NUMBER(2,4) は 0.0099）、ほかは番号を小数の下の桁に（0.001・0.002…）
            $intDigits = 10
            if ($len -gt 0) { $intDigits = [math]::Max(0, $len - $scale) }
            $fraction = ''
            if ($scale -gt 0) { $fraction = '.' + ('9' * $scale) }
            if ($Kind -eq 'Max') { if ($intDigits -eq 0) { return '0.' + ('9' * $len).PadLeft($scale, '0') }; return ('9' * $intDigits) + $fraction }
            if ($Kind -eq 'Min') { return '0' }
            if ($intDigits -eq 0) { return '0.' + ([string](Get-FitNumber $n (Get-DigitsCapacity $len))).PadLeft($scale, '0') }
            $v = Get-FitNumber $n (Get-DigitsCapacity $intDigits)
            if (-not $isPk -and ([string]($v * 100)).Length -le $intDigits) { $v = $v * 100 }
            if ($scale -gt 0 -and -not $isPk) { return '{0}.{1}' -f $v, '5'.PadRight($scale, '0') }
            return [string]$v
        }
        'Integer' { if ($Kind -eq 'Max') { return Get-IntegerMax $Column }; if ($Kind -eq 'Min') { return '0' }; return [string](Get-FitNumber $n ([long](Get-IntegerMax $Column))) }
        'Float' { if ($Kind -eq 'Max') { return '99999999.99' }; if ($Kind -eq 'Min') { return '0' }; return '{0}.5' -f $n }
        'Date' { if ($Kind -eq 'Max') { return '9999-12-31' }; if ($Kind -eq 'Min') { return '1900-01-01' }; return $BaseDate.AddDays($n - 1).ToString('yyyy-MM-dd', $invariant) }
        'Timestamp' { if ($Kind -eq 'Max') { return '9999-12-31 23:59:59' }; if ($Kind -eq 'Min') { return '1900-01-01 00:00:00' }; return $BaseDate.AddDays($n - 1).AddHours(9).ToString('yyyy-MM-dd HH:mm:ss', $invariant) }
        'Time' { if ($Kind -eq 'Max') { return '23:59:59' }; if ($Kind -eq 'Min') { return '00:00:00' }; return $BaseDate.AddHours(9).AddMinutes($n - 1).ToString('HH:mm:ss', $invariant) }
        'Boolean' { if ($Kind -eq 'Max') { return '1' }; if ($Kind -eq 'Min') { return '0' }; return [string]($n % 2) }
        'Binary' { return $null }
    }
    # 文字: 主キーは番号（固定長は桁数まで 0 で埋める）、ほかは論理名と番号。最大は桁数いっぱい（固定長は 9、可変長は全角の「あ」）
    $width = $len
    if ($width -le 0) { $width = 100 }
    if ($isPk) {
        if ($Kind -eq 'Max') { if ($vk -eq 'Char') { return '9' * $width } else { return 'Z' * $width } }
        if ($Kind -eq 'Min') { if ($vk -eq 'Char') { return '0' * $width } else { return '0' } }
        $digits = [string](Get-FitNumber $n (Get-DigitsCapacity $width))
        if ($vk -eq 'Char' -and $len -gt 0) { return $digits.PadLeft($len, '0') }
        return $digits.PadLeft([math]::Min($width, 6), '0')
    }
    if ($Kind -eq 'Max') { if ($vk -eq 'Char') { return '9' * $width } else { return 'あ' * $width } }
    if ($Kind -eq 'Min') { if ($vk -eq 'Char') { return '0' * $width } else { return 'あ' } }
    $base = $Column.Logical
    if (-not $base) { $base = $Column.Name }
    $text = '{0}{1}' -f $base, $n
    if ($len -gt 0 -and $text.Length -gt $len) {
        $digits = [string]$n
        if ($digits.Length -le $len) { $text = $digits.PadLeft($len, '0') } else { $text = $text.Substring(0, $len) }
    }
    return $text
}

$generated = @{}    # テーブル ID（大文字）→ 行の一覧（@{ Pattern; Kind; Seq; Values = カラム名（大文字）→ 値（NULL は $null） }）
$notes = @{}        # テーブル ID（大文字）→ 注意の一覧
$codeColumns = New-Object System.Collections.Generic.List[string]
foreach ($t in $ordered) {
    $tableNotes = New-Object System.Collections.Generic.List[string]
    $kinds = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -le $Rows; $i++) { $kinds.Add([pscustomobject]@{ Kind = 'Normal'; Seq = $i; Pattern = "標準 $i" }) }
    if (-not $NoBoundary) {
        $kinds.Add([pscustomobject]@{ Kind = 'Max'; Seq = $Rows + 1; Pattern = '最大' })
        $kinds.Add([pscustomobject]@{ Kind = 'Min'; Seq = $Rows + 2; Pattern = '最小' })
        $kinds.Add([pscustomobject]@{ Kind = 'Null'; Seq = $Rows + 3; Pattern = 'NULL' })
    }
    foreach ($c in $t.Columns) {
        if (-not ($null -ne $c.PkOrder) -and (Get-CodeValues $c)) { $codeColumns.Add(('{0}.{1}' -f $t.Id, $c.Name)) }
        if ((Get-ValueKind $c) -eq 'Binary' -and $c.NotNull) { $tableNotes.Add(('{0} はバイナリのため値を作りません（必須のため、登録の前に値を入れてください）' -f $c.Name)) }
        if ($c.FkTable -and -not [string]::Equals($c.FkTable, $t.Id, [System.StringComparison]::OrdinalIgnoreCase)) {
            $parent = $tableById[$c.FkTable.ToUpperInvariant()]
            if (-not $parent -or -not $generated.ContainsKey($parent.Id.ToUpperInvariant())) { $tableNotes.Add(('{0} の参照先 {1}.{2} のデータを作っていないため、値は参照先にあるとは限りません' -f $c.Name, $c.FkTable, $c.FkColumn)) }
        }
    }
    $rowsOfTable = New-Object System.Collections.Generic.List[object]
    foreach ($k in $kinds) {
        $values = [ordered]@{}
        # 参照先のテーブルを見ない値を先に作り、外部キーはあとで参照先の行の値にする
        foreach ($c in $t.Columns) { $values[$c.Name.ToUpperInvariant()] = Get-NewValue $c $k.Kind $k.Seq }
        foreach ($c in $t.Columns) {
            if (-not $c.FkTable) { continue }
            $key = $c.Name.ToUpperInvariant()
            if ($k.Kind -eq 'Null' -and -not $c.NotNull -and $null -eq $c.PkOrder) { $values[$key] = $null; continue }
            if ([string]::Equals($c.FkTable, $t.Id, [System.StringComparison]::OrdinalIgnoreCase)) {
                # 自分を参照する外部キー（親の行など）: 必須でなければ空、必須なら同じ行
                if ($c.NotNull -and $values.Contains($c.FkColumn.ToUpperInvariant())) { $values[$key] = $values[$c.FkColumn.ToUpperInvariant()] } else { $values[$key] = $null }
                continue
            }
            $parentRows = $null
            if ($generated.ContainsKey($c.FkTable.ToUpperInvariant())) { $parentRows = $generated[$c.FkTable.ToUpperInvariant()] }
            if (-not $parentRows) { continue }
            # 標準の行は参照先の標準の行を順に、最大・最小・NULL の行は参照先の同じ種類の行を参照する
            $normal = @($parentRows | Where-Object { $_.Kind -eq 'Normal' })
            $target = $normal[($k.Seq - 1) % $normal.Count]
            if ($k.Kind -ne 'Normal') {
                $same = @($parentRows | Where-Object { $_.Kind -eq $k.Kind })
                if ($same.Count -gt 0) { $target = $same[0] }
            }
            $fk = $c.FkColumn.ToUpperInvariant()
            if ($target.Values.Contains($fk)) { $values[$key] = $target.Values[$fk] }
        }
        $rowsOfTable.Add([pscustomobject]@{ Pattern = $k.Pattern; Kind = $k.Kind; Seq = $k.Seq; Values = $values })
    }
    # 主キーが重なる行（桁数の少ない主キーで、番号を 1 から数え直したものなど）
    $pkNames = @($t.Columns | Where-Object { $null -ne $_.PkOrder } | ForEach-Object { $_.Name.ToUpperInvariant() })
    if ($pkNames.Count -gt 0) {
        $firstOf = New-Object 'System.Collections.Generic.Dictionary[string,string]'
        $same = New-Object System.Collections.Generic.List[string]
        foreach ($r in $rowsOfTable) {
            $key = (@($pkNames | ForEach-Object { [string]$r.Values[$_] })) -join "`t"
            if ($firstOf.ContainsKey($key)) { $same.Add(('{0} と {1}' -f $firstOf[$key], $r.Pattern)) } else { $firstOf[$key] = $r.Pattern }
        }
        if ($same.Count -gt 0) {
            $list = (@($same | Select-Object -First 3)) -join '・'
            if ($same.Count -gt 3) { $list += (' ほか {0} 行' -f ($same.Count - 3)) }
            $tableNotes.Add(('主キー（{0}）が重なる行があります（{1}）。主キーの桁数に入らない番号は 1 から数え直すため、-Rows を減らすか、主キーの桁数を確かめてください' -f ($pkNames -join '・'), $list))
        }
    }
    $generated[$t.Id.ToUpperInvariant()] = $rowsOfTable
    $notes[$t.Id.ToUpperInvariant()] = $tableNotes
}

# ---------------------------------------------------------------- 書き出す


function ConvertTo-SqlLiteral {
    #     値を SQL のリテラルにする（日付・日時は DB ごとの書き方。文字は ' を重ね、SQL Server は N'…'、MySQL は \ も重ねる）
    param($Value, [string] $Kind)
    if ($null -eq $Value) { return 'NULL' }
    switch ($Kind) {
        'Decimal' { return $Value }
        'Integer' { return $Value }
        'Float' { return $Value }
        'Boolean' { if ($Database -eq 'PostgreSQL') { if ($Value -eq '1') { return 'TRUE' } else { return 'FALSE' } }; return $Value }
        'Date' { if (@('Oracle', 'PostgreSQL') -contains $Database) { return "DATE '$Value'" }; return "'$Value'" }
        'Timestamp' { if (@('Oracle', 'PostgreSQL') -contains $Database) { return "TIMESTAMP '$Value'" }; return "'$Value'" }
        'Time' {
            # Oracle には時刻だけの型が無い（型の対応では TIMESTAMP）ので、基準日の日時にする
            if ($Database -eq 'Oracle') { return "TIMESTAMP '{0} {1}'" -f $BaseDate.ToString('yyyy-MM-dd', $invariant), $Value }
            if ($Database -eq 'PostgreSQL') { return "TIME '$Value'" }
            return "'$Value'"
        }
    }
    $text = ([string]$Value).Replace("'", "''")
    if ($Database -eq 'MySQL') { $text = $text.Replace('\', '\\') }
    if ($Database -eq 'SQLServer') { return "N'$text'" }
    return "'$text'"
}

function ConvertTo-CsvField {
    param($Value)
    if ($null -eq $Value) { return '' }
    $v = [string]$Value
    if ($v -match '[",\r\n]') { return '"' + $v.Replace('"', '""') + '"' }
    return $v
}

if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
$sql = New-Object System.Text.StringBuilder
[void]$sql.AppendLine(('-- テストデータ（New-TestData.ps1・{0}）。参照先のテーブルから順に登録します' -f $Database))
if ($IncludeDelete) {
    [void]$sql.AppendLine('')
    for ($i = $ordered.Count - 1; $i -ge 0; $i--) { [void]$sql.AppendLine(('DELETE FROM {0};' -f $ordered[$i].Id)) }
}
$results = New-Object System.Collections.Generic.List[object]
$sheets = New-Object System.Collections.Generic.List[object]
$summaryRows = New-Object System.Collections.Generic.List[object]
foreach ($t in $ordered) {
    $rowsOfTable = $generated[$t.Id.ToUpperInvariant()]
    $names = @($t.Columns | ForEach-Object { $_.Name })
    $valueKinds = @($t.Columns | ForEach-Object { Get-ValueKind $_ })
    # CSV（1 行目はカラムの物理名。NULL は空）
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add((@($names | ForEach-Object { ConvertTo-CsvField $_ })) -join ',')
    foreach ($r in $rowsOfTable) { $lines.Add((@($t.Columns | ForEach-Object { ConvertTo-CsvField $r.Values[$_.Name.ToUpperInvariant()] })) -join ',') }
    $csvPath = Join-Path $OutDir ('{0}.csv' -f $t.Id)
    [System.IO.File]::WriteAllText($csvPath, (($lines -join "`r`n") + "`r`n"), $enc)
    # INSERT 文
    [void]$sql.AppendLine('')
    $title = $t.Id
    if ($t.Name) { $title = '{0}（{1}）' -f $t.Id, $t.Name }
    [void]$sql.AppendLine(('-- {0}' -f $title))
    foreach ($r in $rowsOfTable) {
        $literals = @(for ($i = 0; $i -lt $t.Columns.Count; $i++) { ConvertTo-SqlLiteral $r.Values[$t.Columns[$i].Name.ToUpperInvariant()] $valueKinds[$i] })
        [void]$sql.AppendLine(('INSERT INTO {0} ({1}) VALUES ({2}); -- {3}' -f $t.Id, ($names -join ', '), ($literals -join ', '), $r.Pattern))
    }
    $tableNotes = $notes[$t.Id.ToUpperInvariant()]
    foreach ($n in $tableNotes) { Write-Warning ('{0}: {1}' -f $t.Id, $n) }
    $parents = @($t.Columns | Where-Object { $_.FkTable } | ForEach-Object { $_.FkTable } | Select-Object -Unique)
    $results.Add([pscustomobject]@{ Table = $t.Id; Name = $t.Name; Rows = $rowsOfTable.Count; Csv = $csvPath })
    $summaryRows.Add(@{ Values = @($t.Id, $t.Name, [double]$rowsOfTable.Count, ($parents -join '・'), ($tableNotes -join ' / ')); Link = @{ Column = 0; Path = $csvPath } })
    if ($OutXlsx) {
        $columns = @(@{ Header = 'パターン'; Width = 9 }) + @($t.Columns | ForEach-Object { @{ Header = $_.Name; Width = [math]::Min(40, [math]::Max(10, $_.Name.Length + 2)) } })
        $dataRows = @($rowsOfTable | ForEach-Object { $r = $_; @{ Values = @(@($r.Pattern) + @($t.Columns | ForEach-Object { $v = $r.Values[$_.Name.ToUpperInvariant()]; if ($null -eq $v) { '' } else { [string]$v } })) } })
        $sheetName = $t.Id
        if ($sheetName.Length -gt 31) { $sheetName = $sheetName.Substring(0, 31) }
        $sheets.Add(@{ Name = $sheetName; Columns = $columns; Rows = $dataRows })
    }
}
$sqlPath = Join-Path $OutDir 'テストデータ.sql'
[System.IO.File]::WriteAllText($sqlPath, $sql.ToString().Replace("`r`n", "`n").Replace("`n", "`r`n"), $enc)

if ($OutXlsx) {
    try {
        $all = @(@{ Name = '一覧'; Columns = @(@{ Header = 'テーブル'; Width = 20 }, @{ Header = 'テーブル名'; Width = 20 }, @{ Header = '行'; Width = 6 }, @{ Header = '参照先'; Width = 24 }, @{ Header = '注意'; Width = 70 }); Rows = $summaryRows }) + $sheets.ToArray()
        Export-Workbook -FilePath $OutXlsx -Sheets $all
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

$results

$normalCount = $Rows * $ordered.Count
$total = 0
foreach ($r in $results) { $total += $r.Rows }
Write-Host ''
Write-Host ('テストデータを作りました: テーブル {0} 件・行 {1} 件（標準 {2}・境界値 {3}）: {4}' -f $ordered.Count, $total, $normalCount, ($total - $normalCount), $OutDir)
if ($addedParents.Count -gt 0) { Write-Host ('参照先のテーブルも作りました: {0}' -f ($addedParents -join '・')) }
if ($codeColumns.Count -gt 0) { Write-Host ('コード値から作ったカラム: {0}' -f ($codeColumns -join '・')) }
Write-Host ('INSERT 文: {0}' -f $sqlPath)
if ($OutXlsx) { Write-Host "一覧を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テーブル定義書から、テストデータ（テーブルごとの CSV と INSERT 文）を作ります。主キー・外部キー・コード値・桁数の境界を考えて作ります（Excel 不要）。
#
# .DESCRIPTION
#     テーブル定義書（Export-TableDdl.ps1 と同じ見出しで読む）のカラム定義から、テーブルごとに次の行を作ります。
#
#       標準（-Rows 件）: 主キーは番号、文字は「論理名＋番号」、数値・日付は番号と基準日から
#       最大: 文字は桁数いっぱい（固定長は 9、可変長は全角の「あ」）、数値は桁数の最大、日付は 9999-12-31
#       最小: 文字は 1 文字（固定長は桁数の 0）、数値は 0、日付は 1900-01-01
#       NULL: 必須でないカラムを空にした行
#
#     外部キーのカラムは、参照先のテーブルの行の値にします（参照先のテーブルから順に作り、-Table で選んだテーブルの参照先も作ります）。
#     コード定義書（コード区分名・コード値の見出しのある表）を -Path に入れると、論理名がコード区分名と同じカラムか、
#     備考にコード区分 ID・区分名を書いたカラムは、コード値から作ります。名前が「フラグ」で終わるカラムは 0・1 にします。
#
#     結果は -OutDir に、テーブルごとの CSV（テーブル ID.csv。1 行目はカラムの物理名）と、すべてのテーブルの INSERT 文（テストデータ.sql）を保存します。
#
# .PARAMETER Path
#     テーブル定義書（とコード定義書）のファイルまたはフォルダ（複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER OutDir
#     テストデータの保存先のフォルダ。テーブルごとの CSV と、INSERT 文（テストデータ.sql）を保存します。
#
# .PARAMETER Table
#     作るテーブルの ID（ワイルドカード可・複数可）。外部キーの参照先のテーブルも作ります。省略するとすべてのテーブル。
#
# .PARAMETER Rows
#     テーブルごとの標準の行の数（既定 3）。
#
# .PARAMETER Database
#     INSERT 文の日付・日時・文字の書き方（Oracle / PostgreSQL / SQLServer / MySQL）。既定は Oracle。
#
# .PARAMETER BaseDate
#     標準の行の日付の基準日（1 行目がこの日、2 行目が次の日…）。既定は今日。
#
# .PARAMETER NoBoundary
#     最大・最小・NULL の行を作りません（標準の行だけ）。
#
# .PARAMETER IncludeDelete
#     INSERT 文の前に、すべてのテーブルの DELETE 文（参照する側のテーブルから）を書きます。
#
# .PARAMETER Encoding
#     CSV と INSERT 文の文字コード（UTF8 / UTF8BOM / ShiftJIS）。既定は UTF8BOM（Excel でそのまま開ける）。
#
# .PARAMETER OutXlsx
#     作ったデータの一覧を xlsx に保存します（一覧と、テーブルごとのシート。行のパターン付き）。
#
# .PARAMETER Headers
#     テーブル定義書の見出しの候補の CSV。既定は tools\ddl\テーブル定義書の見出し.csv（Export-TableDdl.ps1 と同じ）。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .EXAMPLE
#     .\New-TestData.ps1 D:\設計書 -OutDir D:\テストデータ -Database PostgreSQL -Table T_ORDER*
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1 と Search-Excel.ps1 が必要です。
#     作るのは、テーブル定義（型・桁数・主キー・外部キー・必須）とコード値から機械的に決まる値です。業務のルール（チェック制約・項目の関係）は見ません。
