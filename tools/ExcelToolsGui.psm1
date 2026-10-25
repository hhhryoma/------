#     Excel ツールの画面（Start-ExcelTools.ps1）の部品。
#     ツールの一覧（gui\ツールの一覧.psd1）を読み、各ツールの引数とヘルプから入力欄を作り、入力から引数を組み立てて、
#     別の実行空間（runspace）でツールを動かす。前半は画面（Windows Forms）を使わないので、どの OS でもテストできる

$script:CommonParameters = @([System.Management.Automation.PSCmdlet]::CommonParameters) + @([System.Management.Automation.PSCmdlet]::OptionalCommonParameters)
$script:IsWindowsOs = ($PSVersionTable.PSVersion.Major -lt 6) -or $IsWindows
$script:FieldCache = @{}

# ---------------------------------------------------------------- ツールの一覧


function Get-GuiCatalog {
    #     ツールの一覧を読む。戻り値: Groups（Name・For）と Tools（各処理。省略した項目は空の値にそろえる）
    param([string] $Path = (Join-Path (Join-Path $PSScriptRoot 'gui') 'ツールの一覧.psd1'))
    $data = Import-PowerShellDataFile -LiteralPath $Path
    $toolsDir = $PSScriptRoot
    $groups = @($data.Groups | ForEach-Object { [pscustomobject]@{ Name = $_.Name; For = $_.For } })
    $groupNames = @($groups | ForEach-Object { $_.Name })
    $ids = @{}
    $tools = foreach ($t in $data.Tools) {
        if (-not $t.Id -or -not $t.Label -or -not $t.Tool) { throw "ツールの一覧に Id・Label・Tool の無い行があります: $($t | Out-String)" }
        if ($ids.ContainsKey($t.Id)) { throw "ツールの一覧の Id が重なっています: $($t.Id)" }
        $ids[$t.Id] = $true
        if ($groupNames -notcontains $t.Group) { throw "ツールの一覧の Group が Groups にありません: $($t.Id)（$($t.Group)）" }
        $get = { param($Name, $Empty) if ($t.ContainsKey($Name) -and $null -ne $t[$Name]) { $t[$Name] } else { $Empty } }
        $output = @(& $get 'Output' @())
        $open = @(& $get 'Open' @())
        if ($open.Count -eq 0) { $open = $output }
        [pscustomobject]@{
            Id = $t.Id; Label = $t.Label; Group = $t.Group; For = (@($groups | Where-Object Name -eq $t.Group))[0].For
            Tool = $t.Tool; ToolPath = (Join-Path $toolsDir $t.Tool); Set = [string](& $get 'Set' '')
            Summary = [string](& $get 'Summary' ''); Keywords = [string](& $get 'Keywords' '')
            Main = @(& $get 'Main' @()); Labels = (& $get 'Labels' @{}); Kinds = (& $get 'Kinds' @{}); Extensions = (& $get 'Extensions' @{})
            Required = @(& $get 'Required' @()); Hide = @(& $get 'Hide' @()); Defaults = (& $get 'Defaults' @{}); Fixed = (& $get 'Fixed' @{})
            Verbatim = @(& $get 'Verbatim' @()); AllowEmpty = @(& $get 'AllowEmpty' @())
            Output = $output; OutputName = [string](& $get 'OutputName' $t.Label); Open = $open
            WhatIf = [bool](& $get 'WhatIf' $false); Confirm = [string](& $get 'Confirm' ''); ConfirmWhen = [string](& $get 'ConfirmWhen' '')
            ConfirmArguments = (& $get 'ConfirmArguments' @{}); Console = [bool](& $get 'Console' $false)
        }
    }
    return [pscustomobject]@{ Groups = $groups; Tools = @($tools) }
}

function Find-GuiTool {
    #     処理を言葉で探す（名前・説明・キーワード・ツールのファイル名。空白で区切った言葉のすべてを含むもの）
    param($Catalog, [string] $Text)
    $words = @($Text -split '\s+' | Where-Object { $_ })
    foreach ($t in $Catalog.Tools) {
        $hay = ('{0} {1} {2} {3} {4}' -f $t.Label, $t.Summary, $t.Keywords, $t.Tool, $t.Group).ToLowerInvariant()
        if (@($words | Where-Object { $hay.IndexOf($_.ToLowerInvariant()) -lt 0 }).Count -eq 0) { $t }
    }
}

# ---------------------------------------------------------------- 入力欄


function Get-GuiFirstSentence {
    #     ヘルプの最初の文（括弧の外の最初の「。」まで）
    param([string] $Help)
    $depth = 0
    for ($i = 0; $i -lt $Help.Length; $i++) {
        $c = $Help[$i]
        if ($c -eq [char]'（' -or $c -eq [char]'(') { $depth++ }
        elseif (($c -eq [char]'）' -or $c -eq [char]')') -and $depth -gt 0) { $depth-- }
        elseif ($c -eq [char]'。' -and $depth -eq 0) { return $Help.Substring(0, $i).Trim() }
    }
    return $Help.Trim()
}

function ConvertTo-GuiDate {
    #     日付の文字を日付にする（Windows の地域の設定が和暦でも、2026/10/23 は西暦として読む）。読めなければ $null
    param([string] $Text)
    $culture = [System.Globalization.CultureInfo]::CurrentCulture.Clone()
    try { $culture.DateTimeFormat.Calendar = New-Object System.Globalization.GregorianCalendar }
    catch { $culture = [System.Globalization.CultureInfo]::InvariantCulture }
    $d = [datetime]::MinValue
    if ([datetime]::TryParse($Text, $culture, [System.Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Format-GuiTime {
    #     時刻を「09:05:03」の形にする（地域の設定で時刻の区切りを変えていても「:」）
    param([datetime] $Time)
    return $Time.ToString('HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-GuiLabel {
    #     ヘルプの最初の文から、入力欄の名前を作る（括弧の中を除き、長ければ最後の「、」のあと）
    param([string] $Help, [string] $Name)
    $text = Get-GuiFirstSentence $Help
    $mark = [string][char]1
    do { $before = $text; $text = [regex]::Replace($text, '（[^（）]*）|\([^()]*\)', $mark) } while ($text -ne $before)
    #     括弧を除いたところは、英数字のとなりなら空白にする（CSV（UTF-8・BOM 付き）に → CSV に）
    $text = [regex]::Replace($text, '(?<=[!-~])\x01+|\x01+(?=[!-~])', ' ').Replace($mark, '')
    $text = [regex]::Replace($text, '\s+', ' ').Trim()
    if (-not $text) { return $Name }
    if ($text.Length -gt 30 -and $text.LastIndexOf('、') -gt 0) {
        $tail = $text.Substring($text.LastIndexOf('、') + 1).Trim()
        if ($tail.Length -ge 8) { $text = $tail }
    }
    if ($text.Length -gt 34) { $text = $text.Substring(0, 33) + '…' }
    return $text
}

function Get-GuiFieldKind {
    #     入力欄の種類を、引数の名前・型・ヘルプの最初の文から決める
    param([string] $Name, [type] $Type, [string] $Help, [bool] $HasChoices)
    $element = $Type
    if ($Type.IsArray) { $element = $Type.GetElementType() }
    if ($Type -eq [System.Management.Automation.SwitchParameter]) { return 'Switch' }
    if ($HasChoices) { return 'Choice' }
    if ($element -eq [datetime]) { return 'Date' }
    if (@([int], [long], [double], [decimal], [single]) -contains $element) { return 'Number' }
    $first = Get-GuiFirstSentence $Help
    $plain = $first
    do { $before = $plain; $plain = [regex]::Replace($plain, '（[^（）]*）|\([^()]*\)', '') } while ($plain -ne $before)
    $folder = $first -match 'フォルダ'
    $file = $first -match 'ファイル|ブック|CSV|csv|xlsx|xlsm|テンプレート|雛形'
    if ($Name -match '^Out' -or ($first -match '保存します' -and ($file -or $folder))) {
        if ($Name -match '(Dir|Root)$' -or ($folder -and -not $file)) { return 'SaveFolder' }
        return 'SaveFile'
    }
    if ($Name -match '(Dir|Root)$') { return 'SaveFolder' }
    # 名前・番号・正規表現などは文字（「ファイル名」「シートの名前」もファイルを選ぶ欄にしない）
    if ($plain -match '名前|ファイル名|番号|正規表現|URL|タイトル') { return 'Text' }
    if ($file -and $folder) { return 'Path' }
    if ($folder) { return 'Folder' }
    if ($file) { return 'File' }
    return 'Text'
}

function Get-GuiExtension {
    #     保存するファイルの拡張子（引数の名前・ヘルプから。わからなければ空）
    param([string] $Name, [string] $Help)
    if ($Name -match 'Xlsx$') { return '.xlsx' }
    if ($Name -match 'Html$') { return '.html' }
    if ($Name -match 'Csv$') { return '.csv' }
    $first = ($Help -split '。')[0]
    if ($first -match 'CSV') { return '.csv' }
    if ($first -match 'xlsx') { return '.xlsx' }
    return ''
}

function Get-GuiFields {
    #     処理の入力欄の一覧。ツールの引数（Get-Command）とヘルプ（Get-Help）から作り、ツールの一覧の指定で上書きする。
    #     Main は最初から見せる欄（ツールの一覧の Main の順）、そのほかは「詳細な設定」の欄（ツールの引数の順）
    param($Entry)
    if ($script:FieldCache.ContainsKey($Entry.Id)) { return $script:FieldCache[$Entry.Id] }
    $cmd = Get-Command -Name $Entry.ToolPath -ErrorAction Stop
    $help = Get-Help -Name $Entry.ToolPath -ErrorAction SilentlyContinue
    $helpText = @{}
    foreach ($p in @($help.parameters.parameter)) {
        if ($p -and $p.name) { $helpText[$p.name] = ((@($p.description | ForEach-Object { $_.Text })) -join "`n").Trim() }
    }
    $fields = New-Object System.Collections.Generic.List[object]
    $order = 0
    foreach ($p in $cmd.Parameters.Values) {
        $order++
        if ($script:CommonParameters -contains $p.Name) { continue }
        if ($Entry.Hide -contains $p.Name -or $Entry.Fixed.ContainsKey($p.Name)) { continue }
        # 引数のセットの違うもの・テスト用の引数（ヘルプの無いもの・スクリプトブロック）は出さない
        $setName = $null
        if ($Entry.Set) {
            if ($p.ParameterSets.ContainsKey($Entry.Set)) { $setName = $Entry.Set }
            elseif ($p.ParameterSets.ContainsKey('__AllParameterSets')) { $setName = '__AllParameterSets' }
            else { continue }
        }
        else { $setName = @($p.ParameterSets.Keys)[0] }
        $text = [string]$helpText[$p.Name]
        if (-not $text) { continue }
        if (@([scriptblock], [object], [object[]]) -contains $p.ParameterType) { continue }
        $choices = @($p.Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | ForEach-Object { $_.ValidValues })
        $kind = Get-GuiFieldKind -Name $p.Name -Type $p.ParameterType -Help $text -HasChoices ($choices.Count -gt 0)
        if ($Entry.Kinds.ContainsKey($p.Name)) { $kind = $Entry.Kinds[$p.Name] }
        $label = Get-GuiLabel -Help $text -Name $p.Name
        if ($Entry.Labels.ContainsKey($p.Name)) { $label = $Entry.Labels[$p.Name] }
        $ext = ''
        if ($kind -eq 'SaveFile') { $ext = Get-GuiExtension -Name $p.Name -Help $text }
        if ($Entry.Extensions.ContainsKey($p.Name)) { $ext = $Entry.Extensions[$p.Name] }
        $mainIndex = [array]::IndexOf([object[]]$Entry.Main, $p.Name)
        $default = $null
        if ($Entry.Defaults.ContainsKey($p.Name)) { $default = $Entry.Defaults[$p.Name] }
        $fields.Add([pscustomobject]@{
                Name = $p.Name; Label = $label; Help = $text; Kind = $kind; Multiple = $p.ParameterType.IsArray
                Choices = $choices; Required = ($p.ParameterSets[$setName].IsMandatory -or $Entry.Required -contains $p.Name)
                Main = ($mainIndex -ge 0); Order = $(if ($mainIndex -ge 0) { $mainIndex } else { 1000 + $order })
                Default = $default; Extension = $ext; Type = $p.ParameterType
                # ワイルドカードを使える引数には、[ ] を含む名前がそのまま通るよう、パスを逃がして渡す
                Wildcard = ($text -match 'ワイルドカード')
            })
    }
    $missing = @($Entry.Main | Where-Object { $n = $_; @($fields | Where-Object Name -eq $n).Count -eq 0 })
    if ($missing.Count -gt 0) { throw ('ツールの一覧の Main に、{0} の引数に無い名前があります: {1}' -f $Entry.Tool, ($missing -join ', ')) }
    $sorted = @($fields | Sort-Object Order)
    $script:FieldCache[$Entry.Id] = $sorted
    return $sorted
}

# ---------------------------------------------------------------- 引数を組み立てる


function Test-GuiEmpty {
    param($Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [bool]) { return -not $Value }
    if ($Value -is [string]) { return -not $Value.Trim() }
    if ($Value -is [array]) { return @($Value | Where-Object { -not (Test-GuiEmpty $_) }).Count -eq 0 }
    return $false
}

function ConvertTo-GuiNumber {
    #     数の文字を数にする（全角も読む。8,000 のような 3 桁ごとのカンマも読む）。読めなければ $null
    param([string] $Text, [type] $Type)
    $styles = [System.Globalization.NumberStyles]::Float
    $n = 0.0
    $t = $Text.Trim().Normalize([System.Text.NormalizationForm]::FormKC)
    # 3 桁ごとの区切りのカンマだけを除く（7,5 のようなカンマは数として読まない）
    if ($t -match '^[+-]?[0-9]{1,3}(,[0-9]{3})+(\.[0-9]+)?$') { $t = $t.Replace(',', '') }
    if (-not [double]::TryParse($t, $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { return $null }
    if (@([int], [long]) -contains $Type) { if ($n -ne [math]::Truncate($n)) { return $null }; return [Convert]::ChangeType($n, $Type) }
    return $n
}

function ConvertTo-GuiArguments {
    #     入力（引数の名前 → 値）から、ツールに渡す引数（スプラッティング用のハッシュテーブル）を作る。
    #     空の欄は渡さない。戻り値: Arguments と Errors（入力の誤りの文）・ErrorFields（誤りのある欄の引数の名前）
    param($Entry, $Fields, [hashtable] $Values)
    $arguments = [ordered]@{}
    $errors = New-Object System.Collections.Generic.List[string]
    $errorFields = New-Object System.Collections.Generic.List[string]
    $previous = $null
    $counted = 0
    foreach ($f in $Fields) {
        # 前の欄で誤りが増えていれば、その欄を覚える（欄の途中で continue するため、次の欄の頭で数える）
        if ($errors.Count -gt $counted) { $errorFields.Add($previous); $counted = $errors.Count }
        $previous = $f.Name
        $v = $Values[$f.Name]
        # 探す文字・置き換える文字は入力のまま（全角の空白を半角にする・" を探す、など）
        if ($Entry.Verbatim -contains $f.Name -and $v -is [string] -and $v.Length -gt 0) { $arguments[$f.Name] = $v; continue }
        if (Test-GuiEmpty $v) {
            if ($Entry.AllowEmpty -contains $f.Name) { $arguments[$f.Name] = ''; continue }
            if ($f.Required) { $errors.Add(('「{0}」を入れてください' -f $f.Label)) }
            continue
        }
        switch ($f.Kind) {
            'Switch' { $arguments[$f.Name] = $true }
            'Date' {
                $d = $v
                if ($d -isnot [datetime]) {
                    $parsed = ConvertTo-GuiDate ([string]$d)
                    if (-not $parsed) { $errors.Add(('「{0}」は日付で入れてください（例: 2026/10/23）' -f $f.Label)); continue }
                    $d = $parsed
                }
                if ($f.Type -eq [datetime]) { $arguments[$f.Name] = $d.Date } else { $arguments[$f.Name] = $d.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
            }
            'Number' {
                $element = $f.Type
                if ($f.Type.IsArray) { $element = $f.Type.GetElementType() }
                # 複数の数は、カンマ・読点・空白で区切って入れられる。1 つの数の欄は区切らずに読む（8,000 を 8 にしない）
                $texts = @($v | ForEach-Object { [string]$_ })
                if ($f.Multiple) { $texts = @($texts | ForEach-Object { $_ -split '[,、\s]+' } | Where-Object { $_ }) }
                $numbers = @($texts | ForEach-Object { ConvertTo-GuiNumber $_ $element })
                if (@($numbers | Where-Object { $null -eq $_ }).Count -gt 0) { $errors.Add(('「{0}」は数で入れてください' -f $f.Label)); continue }
                if ($f.Multiple) { $arguments[$f.Name] = $numbers } else { $arguments[$f.Name] = $numbers[0] }
            }
            default {
                $items = @($v | ForEach-Object { [string]$_ } | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ })
                # 複数の文字は、カンマ・読点・改行で区切って入れられる（パスは 1 行に 1 つ）
                if ($f.Multiple -and $f.Kind -eq 'Text') { $items = @($items | ForEach-Object { $_ -split '[,、\r\n]+' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                if ($f.Wildcard -and @('File', 'Folder', 'Path') -contains $f.Kind) { $items = @($items | ForEach-Object { [System.Management.Automation.WildcardPattern]::Escape($_) }) }
                if ($f.Multiple) { $arguments[$f.Name] = [string[]]$items } else { $arguments[$f.Name] = $items[0] }
            }
        }
    }
    if ($errors.Count -gt $counted) { $errorFields.Add($previous) }
    foreach ($k in $Entry.Fixed.Keys) { $arguments[$k] = $Entry.Fixed[$k] }
    return [pscustomobject]@{ Arguments = $arguments; Errors = $errors.ToArray(); ErrorFields = $errorFields.ToArray() }
}

function Get-GuiResultDir {
    #     結果の既定の保存先（「送る」と同じ「ドキュメント\Excel ツールの結果」）
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if (-not $docs) { $docs = [System.IO.Path]::GetTempPath() }
    return (Join-Path $docs 'Excel ツールの結果')
}

function Get-GuiDefaultOutputs {
    #     結果の欄（ツールの一覧の Output）のうち空のものに入れるパス（名前_日時.拡張子。同じ名前があれば連番）。戻り値: 引数の名前 → パス。
    #     AutoFilled（引数の名前 → 前にここで入れたパス）のままの欄も空とみなす（入れた名前のままもう一度実行しても、前の結果を上書きしない）
    param($Entry, $Fields, [hashtable] $Values, [string] $Dir = (Get-GuiResultDir), [datetime] $Now = (Get-Date), [hashtable] $AutoFilled = @{})
    $result = @{}
    $stamp = $Now.ToString('yyyyMMdd_HHmmss', [System.Globalization.CultureInfo]::InvariantCulture)
    foreach ($name in $Entry.Output) {
        $current = $Values[$name]
        $auto = $AutoFilled -and $AutoFilled.ContainsKey($name) -and -not (Test-GuiEmpty $current) -and ([string]$current).Trim() -eq [string]$AutoFilled[$name]
        if (-not $auto -and -not (Test-GuiEmpty $current)) { continue }
        $f = @($Fields | Where-Object Name -eq $name)[0]
        if (-not $f) { continue }
        $ext = $f.Extension
        if (-not $ext) { $ext = '.xlsx' }
        $candidate = Join-Path $Dir ('{0}_{1}{2}' -f $Entry.OutputName, $stamp, $ext)
        for ($n = 2; Test-Path -LiteralPath $candidate; $n++) { $candidate = Join-Path $Dir ('{0}_{1}({2}){3}' -f $Entry.OutputName, $stamp, $n, $ext) }
        $result[$name] = $candidate
    }
    return $result
}

function Find-GuiNewResult {
    #     フォルダに保存する処理（週次報告・品質評価など）の結果のファイル: フォルダの直下で、実行を始めてから書いたもののうち、
    #     Excel → HTML → PDF → CSV → 文字 の順、同じ種類なら新しいもの。無ければ $null
    param([string] $Dir, [datetime] $Since)
    if (-not (Test-Path -LiteralPath $Dir -PathType Container)) { return $null }
    $rank = @{ '.xlsx' = 1; '.xlsm' = 2; '.html' = 3; '.pdf' = 4; '.csv' = 5; '.txt' = 6; '.md' = 7 }
    $files = @(Get-ChildItem -LiteralPath $Dir -File | Where-Object { $rank.ContainsKey($_.Extension.ToLowerInvariant()) -and $_.LastWriteTime -ge $Since -and -not $_.Name.StartsWith('~$') })
    if ($files.Count -eq 0) { return $null }
    return @($files | Sort-Object { $rank[$_.Extension.ToLowerInvariant()] }, @{ Expression = { $_.LastWriteTime }; Descending = $true })[0].FullName
}

function Get-GuiOldNewOrder {
    #     2 つのパスを旧・新の順にする（「送る」の差分と同じ。ExcelTools.psm1 の Get-OldNewOrder。ファイルは版数か更新日時、
    #     フォルダは名前の版数・中のファイルの版数・中でいちばん新しいファイルの更新日時で決める）。
    #     Detail を付けると、パスの並び（Paths）と決め方（Rule・ByTime。決めなかったときは Rule が空）を返す
    param([string[]] $Paths, [switch] $Detail)
    $result = [pscustomobject]@{ Paths = @($Paths); Rule = ''; ByTime = $false }
    if (@($Paths).Count -eq 2) {
        Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -DisableNameChecking
        $a = Get-Item -LiteralPath $Paths[0] -ErrorAction SilentlyContinue
        $b = Get-Item -LiteralPath $Paths[1] -ErrorAction SilentlyContinue
        if ($a -and $b -and $a.PSIsContainer -eq $b.PSIsContainer) {
            $order = Get-OldNewOrder $a.FullName $b.FullName
            if ($order.Swapped) { $result.Paths = @($Paths[1], $Paths[0]) }
            $result.Rule = $order.Rule
            $result.ByTime = $order.ByTime
        }
    }
    if ($Detail) { return $result }
    return $result.Paths
}

function Get-GuiBrowseStart {
    #     ファイル・フォルダを選ぶ画面を開く場所（欄に入っているパスから）。戻り値: Dir（あるフォルダ。無ければ $null）・Name（ファイルの名前）。
    #     名前だけ・使えない文字を含むなど、フォルダのわからない入力でも止まらない
    param([string] $Current)
    $result = [pscustomobject]@{ Dir = $null; Name = '' }
    if (-not $Current -or -not $Current.Trim()) { return $result }
    $Current = $Current.Trim().Trim('"')
    try {
        if (Test-Path -LiteralPath $Current -PathType Container) { $result.Dir = $Current; return $result }
        $result.Name = Split-Path -Leaf $Current
        $parent = Split-Path -Parent $Current
        if ($parent -and (Test-Path -LiteralPath $parent -PathType Container)) { $result.Dir = $parent }
    }
    catch { }
    return $result
}

function ConvertTo-GuiLiteral {
    #     値を PowerShell の式の文字にする（コマンドの表示・PowerShell の画面での実行に使う）。日付は括弧で囲む（コマンドの引数でも日付になるように）。
    #     文字の中の引用符は 2 つ重ねる（' のほか、PowerShell が同じ引用符とみなす ‘ ’ ‚ ‛ も。Word などが Customer’s のように書く）
    param($Value)
    if ($null -eq $Value) { return '$null' }
    if ($Value -is [bool]) { if ($Value) { return '$true' } else { return '$false' } }
    if ($Value -is [datetime]) { return "([datetime]'{0}')" -f $Value.ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) { return ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $Value)) }
    if ($Value -is [array]) { return '@(' + ((@($Value | ForEach-Object { ConvertTo-GuiLiteral $_ })) -join ', ') + ')' }
    return "'" + [System.Management.Automation.Language.CodeGeneration]::EscapeSingleQuotedStringContent([string]$Value) + "'"
}

function Format-GuiCommandLine {
    #     同じことをする PowerShell のコマンド
    param([string] $ToolPath, $Arguments)
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add('& ' + (ConvertTo-GuiLiteral $ToolPath))
    foreach ($k in $Arguments.Keys) {
        $v = $Arguments[$k]
        if ($v -is [bool] -and $v) { $parts.Add('-' + $k); continue }
        if ($v -is [array]) { $parts.Add(('-{0} {1}' -f $k, ((@($v | ForEach-Object { ConvertTo-GuiLiteral $_ })) -join ', '))); continue }
        $parts.Add(('-{0} {1}' -f $k, (ConvertTo-GuiLiteral $v)))
    }
    return ($parts -join ' ')
}

# ---------------------------------------------------------------- 実行


function Start-GuiToolRun {
    #     ツールを別の実行空間で動かし始める（画面が固まらないように）。Write-Host・警告・エラー・進み具合は Receive-GuiToolRun で受け取る
    param([string] $ToolPath, $Arguments)
    $runspace = [runspacefactory]::CreateRunspace()
    # Excel を使うツールのため、Windows では STA にする
    if ($script:IsWindowsOs) { $runspace.ApartmentState = [System.Threading.ApartmentState]::STA }
    $runspace.ThreadOptions = [System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $runspace
    # Write-Host と警告は、出力に混ぜて受け取ってから Queue に分ける。実行空間の Streams で受け取ると、ツールが中で呼んだツールの
    # 出力を 6>$null・3>$null で捨てても届いてしまい、記録の欄に関係のない行（「1 ファイル中 1 ファイルで 607 件ヒット」など）が出る
    $queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[object]'
    $runner = {
        param($ToolPath, $Arguments, $Queue)
        $ErrorActionPreference = 'Stop'
        filter Split-HostRecord {
            if ($_ -is [System.Management.Automation.InformationRecord] -or $_ -is [System.Management.Automation.WarningRecord]) { $Queue.Enqueue($_) }
            else { $_ }
        }
        # 失敗を exit 1 で知らせるツール（PDF にする）のため、終了コードも受け取る
        $global:LASTEXITCODE = 0
        $result = @(& $ToolPath @Arguments 3>&1 6>&1 | Split-HostRecord)
        $code = $global:LASTEXITCODE
        $text = ''
        if ($result.Count -gt 0) { $text = ($result | Select-Object -First 200 | Format-Table -AutoSize -Wrap | Out-String -Width 220).Trim() }
        [pscustomobject]@{ Count = $result.Count; Text = $text; ExitCode = [int]$code }
    }
    $splat = @{}
    foreach ($k in $Arguments.Keys) { $splat[$k] = $Arguments[$k] }
    [void]$ps.AddScript($runner.ToString()).AddArgument($ToolPath).AddArgument($splat).AddArgument($queue)
    $handle = $ps.BeginInvoke()
    return [pscustomobject]@{
        PowerShell = $ps; Runspace = $runspace; Handle = $handle; Started = (Get-Date)
        Queue = $queue; ErrorIndex = 0; Stopping = $false; Done = $false
    }
}

function ConvertTo-GuiRunMessage {
    #     Write-Host・Write-Information・警告の記録を、記録の欄の 1 行にする
    param($Record)
    if ($Record -is [System.Management.Automation.WarningRecord]) {
        return [pscustomobject]@{ Kind = '警告'; Text = ('警告: ' + $Record.Message); Color = 'DarkYellow'; NoNewLine = $false }
    }
    $data = $Record.MessageData
    if ($data -is [System.Management.Automation.HostInformationMessage]) {
        $color = ''
        if ($null -ne $data.ForegroundColor) { $color = [string]$data.ForegroundColor }
        return [pscustomobject]@{ Kind = '情報'; Text = [string]$data.Message; Color = $color; NoNewLine = [bool]$data.NoNewLine }
    }
    return [pscustomobject]@{ Kind = '情報'; Text = [string]$data; Color = ''; NoNewLine = $false }
}

function Receive-GuiToolRun {
    #     実行中のツールの新しい出力を受け取る。戻り値: Messages（Text・Kind（情報・警告・エラー）・Color・NoNewLine）・Progress・Completed・
    #     Succeeded・Summary（出力の件数と表）・Error
    param($Run)
    $messages = New-Object System.Collections.Generic.List[object]
    $streams = $Run.PowerShell.Streams
    $record = $null
    while ($Run.Queue.TryDequeue([ref]$record)) { $messages.Add((ConvertTo-GuiRunMessage $record)) }
    while ($Run.ErrorIndex -lt $streams.Error.Count) {
        $messages.Add([pscustomobject]@{ Kind = 'エラー'; Text = ('エラー: ' + $streams.Error[$Run.ErrorIndex].Exception.Message); Color = 'Red'; NoNewLine = $false })
        $Run.ErrorIndex++
    }
    $progress = $null
    if ($streams.Progress.Count -gt 0) {
        $last = $streams.Progress[$streams.Progress.Count - 1]
        $streams.Progress.Clear()
        $progress = [pscustomobject]@{ Activity = $last.Activity; Status = $last.StatusDescription; Percent = $last.PercentComplete; Completed = ($last.RecordType -eq 'Completed') }
    }
    $state = [pscustomobject]@{ Messages = $messages.ToArray(); Progress = $progress; Completed = $false; Succeeded = $false; Stopped = $false; Summary = $null; Error = $null; Seconds = ((Get-Date) - $Run.Started).TotalSeconds }
    if ($Run.Done -or -not $Run.Handle.IsCompleted) { return $state }
    $Run.Done = $true
    $state.Completed = $true
    try {
        $out = @($Run.PowerShell.EndInvoke($Run.Handle))
        $state.Summary = @($out | Where-Object { $_ -and $_.PSObject.Properties['Count'] -and $_.PSObject.Properties['Text'] })[0]
        $state.Succeeded = $true
        if ($state.Summary -and $state.Summary.ExitCode -ne 0) {
            $state.Succeeded = $false
            $state.Error = ('ツールが失敗を知らせました（終了コード {0}）。記録の欄の「失敗」の行を確かめてください' -f $state.Summary.ExitCode)
        }
    }
    catch {
        $ex = $_.Exception
        while ($ex -is [System.Management.Automation.MethodInvocationException] -and $null -ne $ex.InnerException) { $ex = $ex.InnerException }
        if ($ex -is [System.Management.Automation.PipelineStoppedException] -or $Run.Stopping) { $state.Stopped = $true; $state.Error = '中止しました' }
        else {
            $err = $ex.Message
            if ($ex -is [System.Management.Automation.RuntimeException] -and $ex.ErrorRecord -and $ex.ErrorRecord.Exception) { $err = $ex.ErrorRecord.Exception.Message }
            $state.Error = $err
        }
    }
    finally {
        # 終わってから届いた出力も受け取る
        $rest = Receive-GuiToolRunTail $Run
        foreach ($m in $rest) { $messages.Add($m) }
        $state.Messages = $messages.ToArray()
        $Run.PowerShell.Dispose()
        $Run.Runspace.Dispose()
    }
    return $state
}

function Receive-GuiToolRunTail {
    param($Run)
    $streams = $Run.PowerShell.Streams
    $list = New-Object System.Collections.Generic.List[object]
    $record = $null
    while ($Run.Queue.TryDequeue([ref]$record)) {
        $m = ConvertTo-GuiRunMessage $record
        $m.NoNewLine = $false
        $list.Add($m)
    }
    while ($Run.ErrorIndex -lt $streams.Error.Count) { $list.Add([pscustomobject]@{ Kind = 'エラー'; Text = ('エラー: ' + $streams.Error[$Run.ErrorIndex].Exception.Message); Color = 'Red'; NoNewLine = $false }); $Run.ErrorIndex++ }
    return $list.ToArray()
}

function Stop-GuiToolRun {
    #     実行中のツールを止める（止まったら Receive-GuiToolRun の Stopped になる）。WaitSeconds があれば、止まるまでその秒数まで待つ
    #     （画面を閉じるとき。ツールの後片付け（開いた Excel を閉じるなど）を終えさせる）。戻り値: 止まったら $true
    param($Run, [int] $WaitSeconds = 0)
    if ($Run.Done) { return $true }
    if (-not $Run.Stopping) {
        $Run.Stopping = $true
        [void]$Run.PowerShell.BeginStop($null, $null)
    }
    if ($WaitSeconds -le 0) { return $Run.Handle.IsCompleted }
    return $Run.Handle.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($WaitSeconds))
}

function Wait-GuiToolRun {
    #     ツールが終わるまで待ち、受け取った出力をまとめて返す（テストと、画面を使わない実行のため）
    param($Run, [int] $TimeoutSeconds = 600)
    $all = New-Object System.Collections.Generic.List[object]
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($true) {
        $r = Receive-GuiToolRun $Run
        foreach ($m in $r.Messages) { $all.Add($m) }
        if ($r.Completed) { $r.Messages = $all.ToArray(); return $r }
        if ((Get-Date) -gt $deadline) { $null = Stop-GuiToolRun $Run; throw 'ツールが時間内に終わりませんでした' }
        Start-Sleep -Milliseconds 100
    }
}

function New-GuiConsoleScript {
    #     PowerShell の画面で実行するためのスクリプト（途中で入力が要るツール用）。引数を書き込んだ一時ファイルを作り、そのパスを返す
    param([string] $ToolPath, $Arguments, [string] $Title, [string] $Dir = (Join-Path ([System.IO.Path]::GetTempPath()) 'ExcelTools'))
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    # 前に作ったスクリプトのうち、1 日より古いものは消す（たまらないように）
    Get-ChildItem -LiteralPath $Dir -Filter '実行_*.ps1' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } | Remove-Item -Force -ErrorAction SilentlyContinue
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('$ErrorActionPreference = ''Stop''')
    $lines.Add('try { $Host.UI.RawUI.WindowTitle = ' + (ConvertTo-GuiLiteral ('Excel ツール - ' + $Title)) + ' } catch { }')
    $lines.Add('$arguments = @{}')
    foreach ($k in $Arguments.Keys) { $lines.Add(('$arguments[{0}] = {1}' -f (ConvertTo-GuiLiteral $k), (ConvertTo-GuiLiteral $Arguments[$k]))) }
    $lines.Add('try { & ' + (ConvertTo-GuiLiteral $ToolPath) + ' @arguments | Out-Host }')
    $lines.Add('catch { Write-Host ''''; Write-Host (''エラー: '' + $_.Exception.Message) -ForegroundColor Red }')
    $lines.Add('Write-Host ''''')
    $lines.Add('$null = Read-Host ''Enter キーを押すと閉じます''')
    $path = Join-Path $Dir ('実行_{0}.ps1' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
    [System.IO.File]::WriteAllText($path, (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    return $path
}

# ---------------------------------------------------------------- 入力の保存（次に開いたとき、前の入力を出す）


function Get-GuiSettingsPath {
    $base = $env:APPDATA
    if (-not $base) { $base = [Environment]::GetFolderPath('ApplicationData') }
    if (-not $base) { $base = [System.IO.Path]::GetTempPath() }
    return (Join-Path (Join-Path $base 'ExcelTools') '画面の設定.json')
}

function Read-GuiSettings {
    #     保存した入力を読む。戻り値: Tools（処理の Id → 引数の名前 → 値）・LastTool・Recent（最近実行した処理の Id）・OpenResult・ShowAdvanced・
    #     Window（画面の位置と大きさ「左,上,幅,高さ」。最大化なら末尾に「,max」）・Splitter（左の一覧の幅）
    param([string] $Path = (Get-GuiSettingsPath))
    $settings = [pscustomobject]@{ Tools = @{}; LastTool = ''; Recent = @(); OpenResult = $true; ShowAdvanced = $false; Window = ''; Splitter = 0 }
    if (-not (Test-Path -LiteralPath $Path)) { return $settings }
    try { $json = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8) | ConvertFrom-Json }
    catch { return $settings }
    if ($json.LastTool) { $settings.LastTool = [string]$json.LastTool }
    if ($json.Recent) { $settings.Recent = @($json.Recent | ForEach-Object { [string]$_ } | Where-Object { $_ }) }
    if ($json.Window -and [string]$json.Window -match '^-?\d+,-?\d+,\d+,\d+(,max)?$') { $settings.Window = [string]$json.Window }
    if ($json.Splitter) { $settings.Splitter = [int]$json.Splitter }
    if ($null -ne $json.OpenResult) { $settings.OpenResult = [bool]$json.OpenResult }
    if ($null -ne $json.ShowAdvanced) { $settings.ShowAdvanced = [bool]$json.ShowAdvanced }
    if ($json.Tools) {
        foreach ($tp in $json.Tools.PSObject.Properties) {
            $values = @{}
            foreach ($vp in $tp.Value.PSObject.Properties) {
                $v = $vp.Value
                if ($v -is [array]) { $v = [string[]]@($v | ForEach-Object { [string]$_ }) }
                $values[$vp.Name] = $v
            }
            $settings.Tools[$tp.Name] = $values
        }
    }
    return $settings
}

function Save-GuiSettings {
    #     入力を保存する（結果の保存先と日付は保存しない。次は新しい名前・今日にする）
    param($Settings, [string] $Path = (Get-GuiSettingsPath))
    $tools = [ordered]@{}
    foreach ($id in @($Settings.Tools.Keys | Sort-Object)) {
        $values = [ordered]@{}
        foreach ($k in @($Settings.Tools[$id].Keys | Sort-Object)) {
            $v = $Settings.Tools[$id][$k]
            if ($v -is [datetime] -or $null -eq $v) { continue }
            $values[$k] = $v
        }
        $tools[$id] = $values
    }
    $json = [ordered]@{ LastTool = $Settings.LastTool; Recent = @($Settings.Recent); Window = [string]$Settings.Window; Splitter = [int]$Settings.Splitter; OpenResult = [bool]$Settings.OpenResult; ShowAdvanced = [bool]$Settings.ShowAdvanced; Tools = $tools } | ConvertTo-Json -Depth 6
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Path) | Out-Null
    [System.IO.File]::WriteAllText($Path, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Add-GuiRecent {
    #     最近実行した処理の一覧の先頭に足す（同じものは前から除き、5 つまで）
    param($Settings, [string] $Id, [int] $Max = 5)
    $Settings.Recent = @(@($Id) + @($Settings.Recent | Where-Object { $_ -and $_ -ne $Id }) | Select-Object -First $Max)
}

function Get-GuiValuesToSave {
    #     保存する入力（結果の保存先・日付・試しに実行などを除く）
    param($Entry, $Fields, [hashtable] $Values)
    $keep = @{}
    foreach ($f in $Fields) {
        if ($Entry.Output -contains $f.Name -or $f.Kind -eq 'Date') { continue }
        $v = $Values[$f.Name]
        if ($null -eq $v) { continue }
        $keep[$f.Name] = $v
    }
    return $keep
}

# ---------------------------------------------------------------- 画面（Windows Forms。Windows でだけ使う）
#     画面の状態は $script:Gui に置く（イベントの処理はこのモジュールの中で動くので、$script: で同じものを見る）


function New-GuiWindow {
    #     画面を作る（表示はしない。Start-ExcelTools.ps1 が表示する）。Path は最初のパスの欄に入れるファイル・フォルダ（「送る」から）、
    #     Tool は最初に選ぶ処理の Id。Quiet はテスト用（メッセージの画面を出さず、確認には QuietAnswer で答え、結果を開かない）
    param([string[]] $Path, [string] $Tool, [string] $SettingsPath = (Get-GuiSettingsPath), [string] $CatalogPath, [switch] $Quiet)
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $catalogArgs = @{}
    if ($CatalogPath) { $catalogArgs['Path'] = $CatalogPath }
    $s = @{
        Catalog = (Get-GuiCatalog @catalogArgs); SettingsPath = $SettingsPath; Settings = (Read-GuiSettings $SettingsPath)
        InitialPaths = @($Path | Where-Object { $_ }); Bindings = (New-Object System.Collections.Generic.List[object]); WrapLabels = (New-Object System.Collections.Generic.List[object])
        Entry = $null; Run = $null; RunArguments = $null; ResultPath = $null; ResultDir = $null; LastState = $null; LastErrors = @(); LastConsoleScript = $null; AutoOutputs = @{}
        Quiet = [bool]$Quiet; QuietAnswer = $true; CanDrop = ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA)
    }
    $script:Gui = $s
    $font = [System.Drawing.SystemFonts]::MessageBoxFont

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Excel ツール'
    $form.Font = $font
    $form.Size = New-Object System.Drawing.Size(1120, 800)
    $form.MinimumSize = New-Object System.Drawing.Size(880, 620)
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $s.Form = $form
    Restore-GuiWindowBounds
    $s.ToolTip = New-Object System.Windows.Forms.ToolTip
    $s.ToolTip.AutoPopDelay = 20000

    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Dock = [System.Windows.Forms.DockStyle]::Fill
    $split.FixedPanel = [System.Windows.Forms.FixedPanel]::Panel1
    $form.Controls.Add($split)
    $s.Split = $split
    try { $split.SplitterDistance = 300 } catch { }

    # ---- 左: やりたいことの一覧
    $left = New-Object System.Windows.Forms.TableLayoutPanel
    $left.Dock = [System.Windows.Forms.DockStyle]::Fill
    $left.ColumnCount = 1
    $left.RowCount = 3
    [void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$left.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $split.Panel1.Controls.Add($left)
    $findLabel = New-Object System.Windows.Forms.Label
    $findLabel.Text = 'やりたいことを探す（例: 差分・提出）'
    $findLabel.AutoSize = $true
    $findLabel.Margin = New-Object System.Windows.Forms.Padding(6, 8, 6, 2)
    $left.Controls.Add($findLabel, 0, 0)
    $find = New-Object System.Windows.Forms.TextBox
    $find.Dock = [System.Windows.Forms.DockStyle]::Fill
    $find.Margin = New-Object System.Windows.Forms.Padding(6, 2, 6, 6)
    $left.Controls.Add($find, 0, 1)
    $s.Find = $find
    $tree = New-Object System.Windows.Forms.TreeView
    $tree.Dock = [System.Windows.Forms.DockStyle]::Fill
    $tree.HideSelection = $false
    $tree.ShowNodeToolTips = $true
    $tree.ItemHeight = [int]($font.Height * 1.6)
    $tree.Margin = New-Object System.Windows.Forms.Padding(6, 0, 6, 6)
    $left.Controls.Add($tree, 0, 2)
    $s.Tree = $tree

    # ---- 右: 選んだ処理の説明・入力欄・実行・記録
    $right = New-Object System.Windows.Forms.TableLayoutPanel
    $right.Dock = [System.Windows.Forms.DockStyle]::Fill
    $right.ColumnCount = 1
    $right.RowCount = 7
    $right.Padding = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
    foreach ($style in @('AutoSize', 'AutoSize', 'Percent', 'AutoSize', 'AutoSize', 'Absolute', 'AutoSize')) {
        if ($style -eq 'Percent') { [void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) }
        elseif ($style -eq 'Absolute') { [void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 190))) }
        else { [void]$right.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) }
    }
    $split.Panel2.Controls.Add($right)

    $title = New-Object System.Windows.Forms.Label
    $title.AutoSize = $true
    $title.Font = New-Object System.Drawing.Font($font.FontFamily, ($font.Size + 4), [System.Drawing.FontStyle]::Bold)
    $title.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 4)
    $right.Controls.Add($title, 0, 0)
    $s.Title = $title
    $summary = New-Object System.Windows.Forms.Label
    $summary.AutoSize = $true
    $summary.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
    $right.Controls.Add($summary, 0, 1)
    $s.Summary = $summary
    $s.WrapLabels.Add($summary)

    $host1 = New-Object System.Windows.Forms.Panel
    $host1.Dock = [System.Windows.Forms.DockStyle]::Fill
    $host1.AutoScroll = $true
    $host1.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $right.Controls.Add($host1, 0, 2)
    $s.FieldsHost = $host1
    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = [System.Windows.Forms.DockStyle]::Top
    $table.AutoSize = $true
    $table.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $table.ColumnCount = 2
    $table.Padding = New-Object System.Windows.Forms.Padding(4)
    [void]$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 240)))
    [void]$table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    $host1.Controls.Add($table)
    $s.Fields = $table

    $options = New-Object System.Windows.Forms.FlowLayoutPanel
    $options.AutoSize = $true
    $options.Dock = [System.Windows.Forms.DockStyle]::Fill
    $options.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)
    $right.Controls.Add($options, 0, 3)
    $s.ShowAdvanced = New-GuiCheckBox '詳細な設定も表示する' $s.Settings.ShowAdvanced
    $s.WhatIf = New-GuiCheckBox '試しに実行（ファイルを書き換えない）' $true
    $s.OpenResult = New-GuiCheckBox '終わったら結果を開く' $s.Settings.OpenResult
    $s.HelpButton = New-GuiButton 'ツールの説明'
    $s.CopyButton = New-GuiButton 'コマンドをコピー'
    foreach ($c in @($s.ShowAdvanced, $s.WhatIf, $s.OpenResult, $s.HelpButton, $s.CopyButton)) { $options.Controls.Add($c) }
    $s.ToolTip.SetToolTip($s.CopyButton, '同じことをする PowerShell のコマンドをクリップボードに写します（くり返し使うときや、ほかの人に伝えるときに）')

    $runRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $runRow.AutoSize = $true
    $runRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $right.Controls.Add($runRow, 0, 4)
    $s.RunButton = New-GuiButton '実行'
    $s.RunButton.Font = New-Object System.Drawing.Font($font.FontFamily, ($font.Size + 2), [System.Drawing.FontStyle]::Bold)
    $s.RunButton.MinimumSize = New-Object System.Drawing.Size(140, 36)
    $s.StopButton = New-GuiButton '中止'
    $s.StopButton.Enabled = $false
    $s.StopButton.MinimumSize = New-Object System.Drawing.Size(80, 36)
    $s.Progress = New-Object System.Windows.Forms.ProgressBar
    $s.Progress.Size = New-Object System.Drawing.Size(260, 22)
    $s.Progress.Margin = New-Object System.Windows.Forms.Padding(12, 10, 6, 0)
    $s.Status = New-Object System.Windows.Forms.Label
    $s.Status.AutoSize = $true
    $s.Status.Margin = New-Object System.Windows.Forms.Padding(6, 12, 0, 0)
    foreach ($c in @($s.RunButton, $s.StopButton, $s.Progress, $s.Status)) { $runRow.Controls.Add($c) }
    $form.AcceptButton = $s.RunButton

    $log = New-Object System.Windows.Forms.RichTextBox
    $log.Dock = [System.Windows.Forms.DockStyle]::Fill
    $log.ReadOnly = $true
    $log.BackColor = [System.Drawing.SystemColors]::Window
    $log.WordWrap = $false
    $log.Font = New-Object System.Drawing.Font([System.Drawing.FontFamily]::GenericMonospace, $font.Size)
    $log.DetectUrls = $false
    $right.Controls.Add($log, 0, 5)
    $s.Log = $log

    $resultRow = New-Object System.Windows.Forms.FlowLayoutPanel
    $resultRow.AutoSize = $true
    $resultRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $right.Controls.Add($resultRow, 0, 6)
    $s.OpenButton = New-GuiButton '結果を開く'
    $s.FolderButton = New-GuiButton '保存先のフォルダを開く'
    $s.OpenButton.Enabled = $false
    $s.FolderButton.Enabled = $false
    $s.ResultLabel = New-Object System.Windows.Forms.Label
    $s.ResultLabel.AutoSize = $true
    $s.ResultLabel.Margin = New-Object System.Windows.Forms.Padding(6, 8, 0, 0)
    foreach ($c in @($s.OpenButton, $s.FolderButton, $s.ResultLabel)) { $resultRow.Controls.Add($c) }

    $s.Timer = New-Object System.Windows.Forms.Timer
    $s.Timer.Interval = 200

    # ---- イベント（エラーは記録の欄とメッセージで知らせ、画面は続ける）
    $tree.Add_AfterSelect({ if ($_.Node -and $_.Node.Tag) { $id = $_.Node.Tag.Id; Invoke-GuiSafe { Select-GuiTool $id } } })
    $find.Add_TextChanged({ Invoke-GuiSafe { Update-GuiTree $script:Gui.Find.Text } })
    # 左（探す欄・一覧）では Enter で実行しない。探す欄の Enter は見つかった処理を選び、一覧の Enter は入力欄に移る
    foreach ($c in @($find, $tree)) {
        $c.Add_Enter({ $script:Gui.Form.AcceptButton = $null })
        $c.Add_Leave({ $script:Gui.Form.AcceptButton = $script:Gui.RunButton })
    }
    $find.Add_KeyDown({ $e = $_; Invoke-GuiSafe { Invoke-GuiFindKey $e } })
    $tree.Add_KeyDown({ if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) { $_.SuppressKeyPress = $true; Invoke-GuiSafe { Enter-GuiFirstField } } })
    $s.ShowAdvanced.Add_CheckedChanged({ Invoke-GuiSafe { Update-GuiAdvanced } })
    $s.RunButton.Add_Click({ Invoke-GuiSafe { Invoke-GuiRun } })
    $s.StopButton.Add_Click({ if ($script:Gui.Run) { $null = Stop-GuiToolRun $script:Gui.Run; Write-GuiLog '中止しています…' 'DarkYellow' } })
    $s.Timer.Add_Tick({ Update-GuiRun })
    $s.OpenButton.Add_Click({ if ($script:Gui.ResultPath) { Open-GuiPath $script:Gui.ResultPath } })
    $s.FolderButton.Add_Click({ if ($script:Gui.ResultDir) { Open-GuiPath $script:Gui.ResultDir } })
    $s.HelpButton.Add_Click({ Invoke-GuiSafe { Show-GuiToolHelp } })
    $s.CopyButton.Add_Click({ Invoke-GuiSafe { Copy-GuiCommand } })
    $host1.Add_Resize({ Update-GuiWrap })
    $form.Add_Load({ $d = 300; if ($script:Gui.Settings.Splitter -ge 150) { $d = $script:Gui.Settings.Splitter }; try { $script:Gui.Split.SplitterDistance = $d } catch { } })
    $form.Add_FormClosing({ Close-GuiWindow $_ })
    if ($s.CanDrop) {
        # ファイル・フォルダを画面のどこに落としても、空いているパスの欄に入れる
        foreach ($c in @($form, $host1, $table, $log, $tree)) { Enable-GuiDrop $c $null }
    }

    Write-GuiLog '左で、やりたいことを選んでください。ファイルやフォルダは、エクスプローラーからこの画面にドラッグ＆ドロップでも入れられます。' 'Gray'
    Update-GuiTree ''
    $first = $Tool
    if (-not $first) { $first = $s.Settings.LastTool }
    if (-not $first -or @($s.Catalog.Tools | Where-Object Id -eq $first).Count -eq 0) { $first = $s.Catalog.Tools[0].Id }
    Select-GuiTool $first
    Update-GuiWrap
    return $form
}

function Invoke-GuiSafe {
    #     画面の操作を動かし、エラーは記録の欄とメッセージで知らせる（画面は閉じない）
    param([scriptblock] $Action)
    try { & $Action }
    catch {
        Write-GuiLog ('エラー: ' + $_.Exception.Message) 'Red'
        [void](Show-GuiMessage ('エラー: ' + $_.Exception.Message) -Warning)
    }
}

function Get-GuiState {
    #     画面の状態（テスト用）
    return $script:Gui
}

function New-GuiButton {
    param([string] $Text)
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $true
    $b.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $b.Padding = New-Object System.Windows.Forms.Padding(6, 2, 6, 2)
    return $b
}

function New-GuiCheckBox {
    param([string] $Text, [bool] $Checked)
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    $c.Checked = $Checked
    $c.Margin = New-Object System.Windows.Forms.Padding(3, 6, 12, 3)
    return $c
}

function Update-GuiTree {
    #     やりたいことの一覧（担当者向け・管理者向け → 分類 → 処理）。Filter があれば、それを含む処理だけ
    param([string] $Filter)
    $s = $script:Gui
    $tree = $s.Tree
    $entries = @($s.Catalog.Tools)
    if ($Filter -and $Filter.Trim()) { $entries = @(Find-GuiTool $s.Catalog $Filter) }
    $tree.BeginUpdate()
    try {
        $tree.Nodes.Clear()
        # 最近実行したもの（絞り込んでいないとき）
        $recent = @($s.Settings.Recent | ForEach-Object { $id = $_; @($s.Catalog.Tools | Where-Object Id -eq $id)[0] } | Where-Object { $_ })
        if (-not ($Filter -and $Filter.Trim()) -and $recent.Count -gt 0) {
            $root = New-Object System.Windows.Forms.TreeNode('最近使ったもの')
            foreach ($t in $recent) {
                $n = New-Object System.Windows.Forms.TreeNode($t.Label)
                $n.Tag = $t
                $n.ToolTipText = $t.Summary
                [void]$root.Nodes.Add($n)
            }
            [void]$tree.Nodes.Add($root)
            $root.Expand()
        }
        foreach ($for in @('担当者', '管理者')) {
            $root = New-Object System.Windows.Forms.TreeNode(($for + '向け'))
            foreach ($g in @($s.Catalog.Groups | Where-Object For -eq $for)) {
                $items = @($entries | Where-Object Group -eq $g.Name)
                if ($items.Count -eq 0) { continue }
                $gn = New-Object System.Windows.Forms.TreeNode($g.Name)
                foreach ($t in $items) {
                    $n = New-Object System.Windows.Forms.TreeNode($t.Label)
                    $n.Tag = $t
                    $n.ToolTipText = $t.Summary
                    [void]$gn.Nodes.Add($n)
                }
                [void]$root.Nodes.Add($gn)
            }
            if ($root.Nodes.Count -eq 0) { continue }
            [void]$tree.Nodes.Add($root)
            if ($Filter -or $for -eq '担当者') { $root.ExpandAll() }
        }
    }
    finally { $tree.EndUpdate() }
    Select-GuiTreeNode
}

function Invoke-GuiFindKey {
    #     探す欄のキー: Enter は見つかった処理を選んで最初の入力欄に移り、↓ は一覧に移る
    param($KeyArgs)
    $s = $script:Gui
    if ($KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $KeyArgs.SuppressKeyPress = $true
        if (Select-GuiFound) { Enter-GuiFirstField }
    }
    elseif ($KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Down) {
        $KeyArgs.Handled = $true
        [void]$s.Tree.Focus()
    }
}

function Select-GuiFound {
    #     探す欄の言葉で見つかった処理を選ぶ（選んでいる処理が見つかったものにあればそのまま、無ければ一覧の最初のもの）。戻り値: 選んだら $true
    $s = $script:Gui
    $text = $s.Find.Text
    if (-not ($text -and $text.Trim())) { return [bool]$s.Entry }
    $found = @(Find-GuiTool $s.Catalog $text | ForEach-Object { $_.Id })
    if ($found.Count -eq 0) { return $false }
    if ($s.Entry -and $found -contains $s.Entry.Id) { return $true }
    foreach ($root in $s.Tree.Nodes) {
        foreach ($g in $root.Nodes) {
            foreach ($n in $g.Nodes) {
                if ($n.Tag -and $found -contains $n.Tag.Id) { Select-GuiTool $n.Tag.Id; return $true }
            }
        }
    }
    Select-GuiTool $found[0]
    return $true
}

function Enter-GuiFirstField {
    #     最初の入力欄に移る
    $s = $script:Gui
    $first = @($s.Bindings | Where-Object { $_.Field.Main })[0]
    if (-not $first) { $first = @($s.Bindings)[0] }
    if ($first) { [void]$first.Input.Focus() } else { [void]$s.RunButton.Focus() }
}

function Select-GuiTreeNode {
    #     一覧の中で、選んでいる処理の行を選んだ状態にする（選び直しの処理は起こさない）
    $s = $script:Gui
    if (-not $s.Entry) { return }
    # 「最近使ったもの」の行を選んだときは、その行のままにする
    $current = $s.Tree.SelectedNode
    if ($current -and $current.Tag -and $current.Tag.Id -eq $s.Entry.Id) { return }
    foreach ($root in $s.Tree.Nodes) {
        foreach ($g in $root.Nodes) {
            foreach ($n in $g.Nodes) {
                if ($n.Tag -and $n.Tag.Id -eq $s.Entry.Id) {
                    if ($s.Tree.SelectedNode -ne $n) { $s.Selecting = $true; try { $n.EnsureVisible(); $s.Tree.SelectedNode = $n } finally { $s.Selecting = $false } }
                    return
                }
            }
        }
    }
}

function Select-GuiTool {
    #     処理を選び、入力欄を作り直す（前の処理の入力は保存する）
    param([string] $Id)
    $s = $script:Gui
    if ($s.Selecting) { return }
    $entry = @($s.Catalog.Tools | Where-Object Id -eq $Id)[0]
    if (-not $entry) { throw "処理が見つかりません: $Id" }
    if ($s.Entry -and $s.Entry.Id -eq $entry.Id -and $s.Bindings.Count -gt 0) { return }
    if ($s.Entry) { Save-GuiCurrentValues }
    $s.Entry = $entry
    $s.Title.Text = $entry.Label
    $s.Summary.Text = $entry.Summary
    $fields = Get-GuiFields $entry
    $s.CurrentFields = $fields
    $table = $s.Fields
    $table.SuspendLayout()
    try {
        foreach ($c in @($table.Controls)) { $c.Dispose() }
        $table.Controls.Clear()
        $table.RowStyles.Clear()
        $table.RowCount = 0
        $s.Bindings.Clear()
        for ($i = $s.WrapLabels.Count - 1; $i -ge 1; $i--) { $s.WrapLabels.RemoveAt($i) }
        $advancedHeader = $null
        foreach ($f in $fields) {
            if (-not $f.Main -and -not $advancedHeader) {
                $advancedHeader = New-Object System.Windows.Forms.Label
                $advancedHeader.Text = '詳細な設定（空欄・選ばないままなら、ツールの既定のとおり）'
                $advancedHeader.AutoSize = $true
                $advancedHeader.Font = New-Object System.Drawing.Font($s.Form.Font, [System.Drawing.FontStyle]::Bold)
                $advancedHeader.Margin = New-Object System.Windows.Forms.Padding(3, 14, 3, 4)
                $row = Add-GuiTableRow
                $table.Controls.Add($advancedHeader, 0, $row)
                $table.SetColumnSpan($advancedHeader, 2)
                $s.AdvancedHeader = $advancedHeader
            }
            $s.Bindings.Add(@(New-GuiFieldControl $f)[-1])
        }
        if (-not $advancedHeader) { $s.AdvancedHeader = $null }
    }
    finally { $table.ResumeLayout() }

    # 入力: 保存した前の入力 → 既定 → 「送る」から渡されたパス（最初に実行するまで、どの処理を選んでも入れる）
    $saved = $s.Settings.Tools[$entry.Id]
    foreach ($b in $s.Bindings) {
        $v = $null
        if ($saved -and $saved.ContainsKey($b.Field.Name)) { $v = $saved[$b.Field.Name] }
        elseif ($null -ne $b.Field.Default) { $v = $b.Field.Default }
        if ($null -ne $v) { Set-GuiControlValue $b $v }
    }
    if ($s.InitialPaths.Count -gt 0) { Add-GuiPaths $s.InitialPaths -Replace }
    # 表示する前の画面では Visible を読んでも false になるので、出しているかは WhatIfShown に覚える
    $s.WhatIf.Visible = $entry.WhatIf
    $s.WhatIfShown = $entry.WhatIf
    $s.WhatIf.Checked = $true
    if ($entry.Console) { $s.RunButton.Text = 'PowerShell の画面で実行' } else { $s.RunButton.Text = '実行' }
    # 結果の欄は保存しないので、選び直したら画面が入れた保存先も忘れる
    $s.AutoOutputs = @{}
    $s.ResultPath = $null
    $s.ResultDir = $null
    $s.OpenButton.Enabled = $false
    $s.FolderButton.Enabled = $false
    $s.ResultLabel.Text = ''
    Update-GuiAdvanced
    Update-GuiWrap
    Select-GuiTreeNode
    $s.Settings.LastTool = $entry.Id
}

function Add-GuiTableRow {
    $t = $script:Gui.Fields
    $row = $t.RowCount
    $t.RowCount = $row + 1
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    return $row
}

function New-GuiFieldControl {
    #     1 つの入力欄（名前・入力・説明）を作って表に足す。戻り値: Field・Input・Rows（表示を切り替えるもの）
    param($Field)
    $s = $script:Gui
    $table = $s.Fields
    $rows = New-Object System.Collections.Generic.List[object]
    $binding = [pscustomobject]@{ Field = $Field; Input = $null; Rows = $rows }
    $row = Add-GuiTableRow

    $name = New-Object System.Windows.Forms.Label
    $name.AutoSize = $true
    $name.MaximumSize = New-Object System.Drawing.Size(232, 0)
    $name.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 0)
    $text = $Field.Label
    if ($Field.Required) { $text += ' *' }
    if ($Field.Kind -ne 'Switch') { $name.Text = $text }
    $table.Controls.Add($name, 0, $row)
    $rows.Add($name)
    $s.ToolTip.SetToolTip($name, ('{0}（-{1}）' -f $Field.Help, $Field.Name))

    $ctl = $null
    $panel = $null
    switch ($Field.Kind) {
        'Switch' {
            $ctl = New-GuiCheckBox $text $false
            $panel = $ctl
        }
        'Choice' {
            if ($Field.Multiple) {
                $ctl = New-Object System.Windows.Forms.CheckedListBox
                $ctl.CheckOnClick = $true
                foreach ($c in $Field.Choices) { [void]$ctl.Items.Add($c) }
                $ctl.Height = [Math]::Min(140, ($Field.Choices.Count * ($s.Form.Font.Height + 3)) + 8)
                $ctl.Width = 320
                $ctl.IntegralHeight = $false
            }
            else {
                $ctl = New-Object System.Windows.Forms.ComboBox
                $ctl.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
                if (-not $Field.Required) { [void]$ctl.Items.Add('') }
                foreach ($c in $Field.Choices) { [void]$ctl.Items.Add($c) }
                $ctl.Width = 240
                $ctl.SelectedIndex = 0
            }
            $panel = $ctl
        }
        'Date' {
            $ctl = New-Object System.Windows.Forms.DateTimePicker
            $ctl.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
            $ctl.ShowCheckBox = $true
            $ctl.Checked = $false
            $ctl.Width = 180
            $panel = $ctl
        }
        'Number' {
            $ctl = New-Object System.Windows.Forms.TextBox
            $ctl.Width = 160
            $panel = $ctl
        }
        default {
            $pathKind = @('File', 'Folder', 'Path', 'SaveFile', 'SaveFolder') -contains $Field.Kind
            $buttons = @()
            if ($Field.Kind -eq 'File') { $buttons = @('ファイルを選ぶ…') }
            elseif ($Field.Kind -eq 'Folder') { $buttons = @('フォルダを選ぶ…') }
            elseif ($Field.Kind -eq 'Path') { $buttons = @('ファイル…', 'フォルダ…') }
            elseif ($Field.Kind -eq 'SaveFile') { $buttons = @('保存先…') }
            elseif ($Field.Kind -eq 'SaveFolder') { $buttons = @('フォルダを選ぶ…') }
            $multiPath = $Field.Multiple -and @('File', 'Folder', 'Path') -contains $Field.Kind
            if ($multiPath) {
                $ctl = New-Object System.Windows.Forms.ListBox
                $ctl.Height = 68
                $ctl.IntegralHeight = $false
                $ctl.HorizontalScrollbar = $true
                $ctl.SelectionMode = [System.Windows.Forms.SelectionMode]::MultiExtended
                $ctl.Dock = [System.Windows.Forms.DockStyle]::Fill
                $buttons += '外す'
                $ctl.Add_KeyDown({ Invoke-GuiListKey $this $_ })
            }
            else {
                $ctl = New-Object System.Windows.Forms.TextBox
                $ctl.Dock = [System.Windows.Forms.DockStyle]::Fill
                # エクスプローラーでコピーしたファイル・フォルダも、Ctrl+V で入れられるようにする
                if ($pathKind) { $ctl.Add_KeyDown({ $box = $this; $e = $_; Invoke-GuiSafe { Invoke-GuiPathKey $box $e } }) }
            }
            $panel = New-Object System.Windows.Forms.TableLayoutPanel
            $panel.Dock = [System.Windows.Forms.DockStyle]::Fill
            $panel.AutoSize = $true
            $panel.Margin = New-Object System.Windows.Forms.Padding(0)
            $panel.ColumnCount = 2
            $panel.RowCount = 1
            [void]$panel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
            [void]$panel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
            $panel.Controls.Add($ctl, 0, 0)
            if ($buttons.Count -gt 0) {
                $flow = New-Object System.Windows.Forms.FlowLayoutPanel
                $flow.AutoSize = $true
                $flow.WrapContents = $false
                $flow.Margin = New-Object System.Windows.Forms.Padding(0)
                if ($multiPath) { $flow.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown }
                foreach ($bt in $buttons) {
                    $btn = New-GuiButton $bt
                    $btn.Tag = $binding
                    $btn.Margin = New-Object System.Windows.Forms.Padding(2, 1, 0, 1)
                    $btn.Add_Click({ $b = $this.Tag; $t = $this.Text; Invoke-GuiSafe { Invoke-GuiBrowse $b $t } })
                    $flow.Controls.Add($btn)
                }
                $panel.Controls.Add($flow, 1, 0)
            }
            if ($pathKind -and $s.CanDrop) { Enable-GuiDrop $ctl $binding }
        }
    }
    $ctl.Tag = $binding
    $binding.Input = $ctl
    if ($panel -eq $ctl) { $ctl.Anchor = [System.Windows.Forms.AnchorStyles]::Left }
    $ctl.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 0)
    $table.Controls.Add($panel, 1, $row)
    $rows.Add($panel)
    $s.ToolTip.SetToolTip($ctl, $Field.Help)

    # 説明（ヘルプ）。結果の欄は、空欄のときの保存先も書く
    $helpText = $Field.Help -replace '\s*\r?\n\s*', ' '
    if ($s.Entry.Output -contains $Field.Name) { $helpText = '空欄なら「ドキュメント\Excel ツールの結果」に保存します。' + $helpText }
    if ($helpText.Length -gt 220) { $helpText = $helpText.Substring(0, 219) + '…' }
    $hrow = Add-GuiTableRow
    $help = New-Object System.Windows.Forms.Label
    $help.Text = $helpText
    $help.AutoSize = $true
    $help.ForeColor = [System.Drawing.SystemColors]::GrayText
    $help.Margin = New-Object System.Windows.Forms.Padding(3, 1, 3, 6)
    $table.Controls.Add($help, 1, $hrow)
    $rows.Add($help)
    $s.WrapLabels.Add($help)
    return $binding
}

function Get-GuiControlValue {
    param($Binding)
    $f = $Binding.Field
    $c = $Binding.Input
    switch ($f.Kind) {
        'Switch' { return [bool]$c.Checked }
        'Date' { if ($c.Checked) { return $c.Value.Date } else { return $null } }
        'Choice' {
            if ($f.Multiple) { return [string[]]@($c.CheckedItems | ForEach-Object { [string]$_ }) }
            return [string]$c.SelectedItem
        }
        default {
            if ($c -is [System.Windows.Forms.ListBox]) { return [string[]]@($c.Items | ForEach-Object { [string]$_ }) }
            return $c.Text
        }
    }
}

function Set-GuiControlValue {
    param($Binding, $Value)
    $f = $Binding.Field
    $c = $Binding.Input
    switch ($f.Kind) {
        'Switch' { $c.Checked = [bool]$Value }
        'Date' {
            $d = $null
            if ($Value -and $Value -isnot [datetime]) { $d = ConvertTo-GuiDate ([string]$Value) }
            if ($Value -is [datetime]) { $c.Value = $Value; $c.Checked = $true }
            elseif ($d) { $c.Value = $d; $c.Checked = $true }
            else { $c.Checked = $false }
        }
        'Choice' {
            if ($f.Multiple) {
                $want = @($Value | ForEach-Object { [string]$_ })
                for ($i = 0; $i -lt $c.Items.Count; $i++) { $c.SetItemChecked($i, ($want -contains [string]$c.Items[$i])) }
            }
            else {
                $i = $c.Items.IndexOf([string]$Value)
                if ($i -ge 0) { $c.SelectedIndex = $i }
            }
        }
        default {
            if ($c -is [System.Windows.Forms.ListBox]) {
                $c.Items.Clear()
                foreach ($v in @($Value)) { if ($v -and -not $c.Items.Contains([string]$v)) { [void]$c.Items.Add([string]$v) } }
            }
            elseif ($Value -is [array]) { $c.Text = (@($Value | ForEach-Object { [string]$_ })) -join ', ' }
            else { $c.Text = [string]$Value }
        }
    }
}

function Set-GuiValue {
    #     入力欄に値を入れる（テスト用。引数の名前で指定）
    param([string] $Name, $Value)
    $b = @($script:Gui.Bindings | Where-Object { $_.Field.Name -eq $Name })[0]
    if (-not $b) { throw "入力欄がありません: $Name" }
    Set-GuiControlValue $b $Value
}

function Get-GuiValues {
    $values = @{}
    foreach ($b in $script:Gui.Bindings) { $values[$b.Field.Name] = Get-GuiControlValue $b }
    return $values
}

function Update-GuiAdvanced {
    $s = $script:Gui
    $show = $s.ShowAdvanced.Checked
    $s.Fields.SuspendLayout()
    try {
        foreach ($b in $s.Bindings) { if (-not $b.Field.Main) { foreach ($c in $b.Rows) { $c.Visible = $show } } }
        if ($s.AdvancedHeader) { $s.AdvancedHeader.Visible = $show }
    }
    finally { $s.Fields.ResumeLayout() }
    $s.Settings.ShowAdvanced = $show
    $s.AdvancedShown = $show
    # 隠した詳細な設定に入れてある欄の数を出す（見えない入力で結果が変わることに気づけるように。結果の保存先は数えない）
    $count = 0
    if (-not $show) { $count = @($s.Bindings | Where-Object { -not $_.Field.Main -and $s.Entry.Output -notcontains $_.Field.Name -and -not (Test-GuiEmpty (Get-GuiControlValue $_)) }).Count }
    $s.AdvancedCount = $count
    if ($count -gt 0) { $s.ShowAdvanced.Text = '詳細な設定も表示する（{0} 件入れています）' -f $count } else { $s.ShowAdvanced.Text = '詳細な設定も表示する' }
}

function Update-GuiWrap {
    #     説明の文を、右側の幅で折り返す
    $s = $script:Gui
    $w = [Math]::Max(200, $s.FieldsHost.ClientSize.Width - 290)
    foreach ($l in $s.WrapLabels) {
        if ($l -eq $s.Summary) { $l.MaximumSize = New-Object System.Drawing.Size(([Math]::Max(300, $s.FieldsHost.ClientSize.Width - 10)), 0) }
        else { $l.MaximumSize = New-Object System.Drawing.Size($w, 0) }
    }
}

function Enable-GuiDrop {
    #     ファイル・フォルダのドラッグ＆ドロップを受ける。Binding が無ければ、空いているパスの欄に入れる
    param($Control, $Binding)
    $Control.AllowDrop = $true
    if ($Binding) { $Control.Tag = $Binding }
    $Control.Add_DragEnter({ if ($_.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) { $_.Effect = [System.Windows.Forms.DragDropEffects]::Copy } })
    $Control.Add_DragDrop({
            $paths = @($_.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop))
            $target = $null
            if ($this.Tag -and $this.Tag.PSObject.Properties['Field']) { $target = $this.Tag }
            Add-GuiPaths $paths $target
        })
}

function Add-GuiPaths {
    #     パスを入力欄に入れる（複数の欄は足す）。欄を指定しなければ、最初の空いているパスの欄に入れ、1 つしか入らない欄が
    #     続くときは順に入れる（旧・新の 2 つの欄には、版数か更新日時の古い方を旧に）。Replace は空いていなくても先頭の欄から入れ直す（「送る」から）
    param([string[]] $Paths, $Binding, [switch] $Replace)
    $s = $script:Gui
    $Paths = @($Paths | Where-Object { $_ })
    if ($Paths.Count -eq 0) { return }
    if ($Binding) {
        if ($Binding.Input -is [System.Windows.Forms.ListBox]) {
            foreach ($p in $Paths) { if (-not $Binding.Input.Items.Contains($p)) { [void]$Binding.Input.Items.Add($p) } }
        }
        else { Set-GuiControlValue $Binding $Paths[0] }
        return
    }
    $candidates = @($s.Bindings | Where-Object { $_.Field.Main -and @('File', 'Folder', 'Path') -contains $_.Field.Kind })
    if (-not $Replace) {
        $empty = @($candidates | Where-Object { Test-GuiEmpty (Get-GuiControlValue $_) })
        if ($empty.Count -gt 0) { $candidates = @($candidates | Where-Object { $empty -contains $_ }) }
    }
    if ($candidates.Count -eq 0) { return }
    if ($candidates[0].Input -is [System.Windows.Forms.ListBox]) {
        if ($Replace) { $candidates[0].Input.Items.Clear() }
        Add-GuiPaths $Paths $candidates[0]
        return
    }
    # 1 つしか入らない欄が続くところへ、順に入れる
    $singles = New-Object System.Collections.Generic.List[object]
    foreach ($c in $candidates) { if ($c.Input -is [System.Windows.Forms.ListBox]) { break }; $singles.Add($c) }
    if ($Paths.Count -eq 2 -and $singles.Count -ge 2 -and $singles[0].Field.Name -eq 'Old' -and $singles[1].Field.Name -eq 'New') {
        $order = Get-GuiOldNewOrder $Paths -Detail
        $Paths = $order.Paths
        # どれで旧・新を決めたかを出す（更新日時で決めたときは、逆なら入れ替えるよう知らせる）
        if ($order.Rule -and $s.Status) {
            $s.Status.Text = '旧・新は{0}で決めました' -f $order.Rule
            if ($order.ByTime) { $s.Status.Text += '（版数で決められないため。逆のときは旧と新を入れ替えてください）' }
        }
    }
    for ($i = 0; $i -lt [Math]::Min($Paths.Count, $singles.Count); $i++) { Set-GuiControlValue $singles[$i] $Paths[$i] }
}

function Invoke-GuiListKey {
    #     パスの一覧で Delete キーは外す、Ctrl+V はエクスプローラーでコピーしたファイル（または 1 行 1 つのパスの文字）を入れる
    param($List, $KeyArgs)
    if ($KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
        foreach ($item in @($List.SelectedItems)) { $List.Items.Remove($item) }
    }
    elseif ($KeyArgs.Control -and $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::V) {
        $paths = @()
        if ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) { $paths = @([System.Windows.Forms.Clipboard]::GetFileDropList()) }
        elseif ([System.Windows.Forms.Clipboard]::ContainsText()) { $paths = @([System.Windows.Forms.Clipboard]::GetText() -split '\r?\n' | ForEach-Object { $_.Trim().Trim('"') } | Where-Object { $_ }) }
        Add-GuiPaths $paths $List.Tag
    }
}

function Invoke-GuiPathKey {
    #     1 つのパスの欄で Ctrl+V は、エクスプローラーでコピーしたファイル・フォルダを入れる（コピーしたのが文字なら、ふつうに貼り付ける）
    param($Box, $KeyArgs)
    if (-not ($KeyArgs.Control -and $KeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::V)) { return }
    if (-not [System.Windows.Forms.Clipboard]::ContainsFileDropList()) { return }
    $KeyArgs.SuppressKeyPress = $true
    Add-GuiPaths @([System.Windows.Forms.Clipboard]::GetFileDropList()) $Box.Tag
}

function Get-GuiFileFilter {
    param($Field)
    $help = $Field.Help
    $all = 'すべてのファイル (*.*)|*.*'
    if ($Field.Kind -eq 'SaveFile' -and $Field.Extension) { return ('{0} (*{1})|*{1}|{2}' -f $Field.Extension.TrimStart('.').ToUpperInvariant(), $Field.Extension, $all) }
    $first = Get-GuiFirstSentence $help
    if ($first -match 'CSV' -and $first -notmatch 'Excel|xlsx|ブック') { return ('CSV (*.csv)|*.csv|{0}' -f $all) }
    if ($first -match 'Excel|xlsx|xlsm|ブック|テンプレート|雛形') { return ('Excel・CSV (*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv)|*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv|{0}' -f $all) }
    return $all
}

function Invoke-GuiBrowse {
    #     「ファイルを選ぶ」「フォルダを選ぶ」「保存先」「外す」のボタン
    param($Binding, [string] $Action)
    $s = $script:Gui
    $f = $Binding.Field
    $c = $Binding.Input
    if ($Action -eq '外す') { foreach ($item in @($c.SelectedItems)) { $c.Items.Remove($item) }; return }
    $start = Get-GuiBrowseStart @(Get-GuiControlValue $Binding | Where-Object { $_ })[0]
    $startDir = $start.Dir
    $picked = @()
    if ($Action -match 'フォルダ') {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $f.Label
        $dlg.ShowNewFolderButton = $true
        if ($startDir) { $dlg.SelectedPath = $startDir }
        if ($dlg.ShowDialog($s.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $picked = @($dlg.SelectedPath) }
        $dlg.Dispose()
    }
    elseif ($Action -eq '保存先…') {
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = Get-GuiFileFilter $f
        $dlg.OverwritePrompt = $true
        if ($startDir) { $dlg.InitialDirectory = $startDir } else { $dlg.InitialDirectory = Get-GuiResultDir }
        if ($start.Name) { $dlg.FileName = $start.Name }
        elseif ($s.Entry.OutputName) { $dlg.FileName = $s.Entry.OutputName + $f.Extension }
        if ($dlg.ShowDialog($s.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $picked = @($dlg.FileName) }
        $dlg.Dispose()
    }
    else {
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = Get-GuiFileFilter $f
        $dlg.Multiselect = $f.Multiple
        $dlg.CheckFileExists = $false
        if ($startDir) { $dlg.InitialDirectory = $startDir }
        if ($dlg.ShowDialog($s.Form) -eq [System.Windows.Forms.DialogResult]::OK) { $picked = @($dlg.FileNames) }
        $dlg.Dispose()
    }
    if ($picked.Count -gt 0) { Add-GuiPaths $picked $Binding }
}

function Write-GuiLog {
    #     記録の欄に 1 行足す（色は PowerShell の色の名前）
    param([string] $Text, [string] $Color, [switch] $NoNewLine)
    $box = $script:Gui.Log
    $map = @{ Red = 'Firebrick'; DarkRed = 'DarkRed'; Yellow = 'DarkGoldenrod'; DarkYellow = 'DarkGoldenrod'; Green = 'ForestGreen'; DarkGreen = 'DarkGreen'; Cyan = 'Teal'; DarkCyan = 'Teal'; Gray = 'DimGray'; DarkGray = 'DimGray'; Blue = 'RoyalBlue'; Magenta = 'DarkMagenta' }
    $box.SelectionStart = $box.TextLength
    $box.SelectionLength = 0
    if ($Color -and $map.ContainsKey($Color)) { $box.SelectionColor = [System.Drawing.Color]::FromName($map[$Color]) } else { $box.SelectionColor = $box.ForeColor }
    if ($NoNewLine) { $box.AppendText($Text) } else { $box.AppendText($Text + "`n") }
    $box.SelectionColor = $box.ForeColor
    $box.ScrollToCaret()
}

function Show-GuiMessage {
    #     メッセージの画面（Quiet のときは出さず、確認には QuietAnswer で答える）。戻り値: はい なら $true
    param([string] $Text, [switch] $YesNo, [switch] $Warning)
    $s = $script:Gui
    if ($s.Quiet) { return $s.QuietAnswer }
    $buttons = [System.Windows.Forms.MessageBoxButtons]::OK
    if ($YesNo) { $buttons = [System.Windows.Forms.MessageBoxButtons]::YesNo }
    $icon = [System.Windows.Forms.MessageBoxIcon]::Information
    if ($Warning) { $icon = [System.Windows.Forms.MessageBoxIcon]::Warning }
    $r = [System.Windows.Forms.MessageBox]::Show($s.Form, $Text, 'Excel ツール', $buttons, $icon)
    return ($r -eq [System.Windows.Forms.DialogResult]::Yes -or $r -eq [System.Windows.Forms.DialogResult]::OK)
}

function Get-GuiRunArguments {
    #     入力から引数を作る（空の結果の欄と、前に画面が入れた保存先のままの欄には、新しい既定の保存先を入れる）。誤りがあれば知らせて $null
    param([switch] $NoDefaults)
    $s = $script:Gui
    $values = Get-GuiValues
    $defaults = @{}
    if (-not $NoDefaults) {
        $defaults = Get-GuiDefaultOutputs $s.Entry $s.CurrentFields $values -AutoFilled $s.AutoOutputs
        foreach ($k in $defaults.Keys) { $values[$k] = $defaults[$k] }
    }
    $built = ConvertTo-GuiArguments $s.Entry $s.CurrentFields $values
    $s.LastErrors = $built.Errors
    # 誤りのある欄の色を変え、最初の欄に移る（直した欄は元の色に戻す）
    $first = $null
    foreach ($b in $s.Bindings) {
        $bad = $built.ErrorFields -contains $b.Field.Name
        if ($b.Input -isnot [System.Windows.Forms.CheckBox]) {
            if ($bad) { $b.Input.BackColor = [System.Drawing.Color]::MistyRose } else { $b.Input.BackColor = [System.Drawing.SystemColors]::Window }
        }
        if ($bad -and -not $first) { $first = $b }
    }
    if ($built.Errors.Count -gt 0) {
        if ($first) { [void]$first.Input.Focus() }
        [void](Show-GuiMessage (($built.Errors -join "`n")) -Warning)
        return $null
    }
    # 空だった結果の欄に、既定の保存先を入れて見せる（入れたものは覚えておき、そのままもう一度実行したら新しい名前にする）
    foreach ($k in $defaults.Keys) {
        $b = @($s.Bindings | Where-Object { $_.Field.Name -eq $k })[0]
        if ($b) { Set-GuiControlValue $b $defaults[$k]; $s.AutoOutputs[$k] = $defaults[$k] }
    }
    return $built.Arguments
}

function Invoke-GuiRun {
    #     「実行」: 入力を確かめ、確認してから、ツールを動かし始める
    $s = $script:Gui
    if ($s.Run -or -not $s.Entry) { return }
    $entry = $s.Entry
    $arguments = Get-GuiRunArguments
    if ($null -eq $arguments) { return }
    $whatIf = $entry.WhatIf -and $s.WhatIf.Checked
    if ($whatIf) { $arguments['WhatIf'] = $true }
    elseif ($entry.Confirm -and (-not $entry.ConfirmWhen -or $arguments.Contains($entry.ConfirmWhen))) {
        if (-not (Show-GuiMessage $entry.Confirm -YesNo -Warning)) { Write-GuiLog '取りやめました。' 'Gray'; return }
        foreach ($k in $entry.ConfirmArguments.Keys) { $arguments[$k] = $entry.ConfirmArguments[$k] }
    }
    Add-GuiRecent $s.Settings $entry.Id
    Save-GuiCurrentValues
    $s.RecentChanged = $true
    foreach ($name in $entry.Output) {
        if ($arguments.Contains($name)) { $dir = Split-Path -Parent ([string]$arguments[$name]); if ($dir) { New-Item -ItemType Directory -Force -Path $dir | Out-Null } }
    }
    $s.RunArguments = $arguments
    $s.InitialPaths = @()
    $s.ResultPath = $null
    $s.ResultDir = $null
    $s.OpenButton.Enabled = $false
    $s.FolderButton.Enabled = $false
    $s.ResultLabel.Text = ''
    Write-GuiLog ''
    $mode = ''
    if ($whatIf) { $mode = '（試しに実行。ファイルは書き換えません）' }
    Write-GuiLog ('■ {0}{1}　{2}' -f $entry.Label, $mode, (Format-GuiTime (Get-Date))) 'Blue'
    Write-GuiLog (Format-GuiCommandLine $entry.ToolPath $arguments) 'Gray'
    if ($entry.Console) {
        $consoleScript = New-GuiConsoleScript $entry.ToolPath $arguments $entry.Label
        $s.LastConsoleScript = $consoleScript
        Write-GuiLog 'PowerShell の画面で実行しています。開いた画面の案内に沿って進めてください。' 'DarkYellow'
        if (-not $s.Quiet) {
            $exe = (Get-Process -Id $PID).Path
            Start-Process -FilePath $exe -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $consoleScript))
        }
        Set-GuiResult $entry $arguments
        $s.RecentChanged = $false
        Update-GuiTree $s.Find.Text
        return
    }
    # ファイルの更新日時の細かさの違いを見込んで、少し前からにする
    $s.RunStarted = (Get-Date).AddSeconds(-2)
    $s.LastProgressText = ''
    $s.Run = Start-GuiToolRun $entry.ToolPath $arguments
    Set-GuiRunning $true
    $s.Timer.Start()
}

function Set-GuiRunning {
    param([bool] $Running)
    $s = $script:Gui
    if (-not $Running -and $s.RecentChanged) { $s.RecentChanged = $false; Update-GuiTree $s.Find.Text }
    $s.RunButton.Enabled = -not $Running
    $s.StopButton.Enabled = $Running
    $s.Tree.Enabled = -not $Running
    $s.Find.Enabled = -not $Running
    $s.Fields.Enabled = -not $Running
    if ($Running) {
        $s.Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Marquee
        $s.Progress.MarqueeAnimationSpeed = 30
        $s.Status.Text = '実行しています…'
        $s.Form.Cursor = [System.Windows.Forms.Cursors]::AppStarting
    }
    else {
        $s.Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
        $s.Progress.Value = 0
        $s.Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Update-GuiRun {
    #     実行中のツールの出力を記録の欄に出す（タイマーから呼ぶ）
    $s = $script:Gui
    if (-not $s.Run) { $s.Timer.Stop(); return }
    try { $r = Receive-GuiToolRun $s.Run }
    catch {
        $s.Timer.Stop()
        $s.Run = $null
        Set-GuiRunning $false
        Write-GuiLog ('エラー: ' + $_.Exception.Message) 'Red'
        return
    }
    foreach ($m in $r.Messages) { Write-GuiLog $m.Text $m.Color -NoNewLine:$m.NoNewLine }
    if ($r.Progress -and -not $r.Completed) {
        if ($r.Progress.Percent -ge 0 -and -not $r.Progress.Completed) {
            $s.Progress.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous
            $s.Progress.Value = [Math]::Max(0, [Math]::Min(100, $r.Progress.Percent))
        }
        $s.LastProgressText = ('{0} {1}' -f $r.Progress.Activity, $r.Progress.Status).Trim()
        if ($r.Progress.Completed) { $s.LastProgressText = '' }
    }
    if (-not $r.Completed) {
        # 進み具合と、かかっている時間
        $text = $s.LastProgressText
        if (-not $text) { $text = '実行しています…' }
        $s.Status.Text = ('{0}（{1:N0} 秒）' -f $text, $r.Seconds)
    }
    if (-not $r.Completed) { return }
    $s.Timer.Stop()
    $s.Run = $null
    $s.LastState = $r
    Set-GuiRunning $false
    if ($r.Summary -and $r.Summary.Count -gt 0 -and $r.Summary.Text) {
        $more = ''
        if ($r.Summary.Count -gt 200) { $more = '（最初の 200 件）' }
        Write-GuiLog ('出力 {0} 件{1}:' -f $r.Summary.Count, $more) 'Gray'
        Write-GuiLog $r.Summary.Text
    }
    if ($r.Succeeded) {
        Write-GuiLog ('終わりました（{0:N1} 秒）' -f $r.Seconds) 'Green'
        $s.Status.Text = '終わりました'
        Set-GuiResult $s.Entry $s.RunArguments $s.RunStarted
        if ($s.ResultPath -and $s.OpenResult.Checked -and -not $s.Quiet) { Open-GuiPath $s.ResultPath }
    }
    else {
        Write-GuiLog $r.Error 'Red'
        $s.Status.Text = $r.Error
        if ($r.Summary) {
            # 最後まで動いて失敗を知らせたとき（一部が失敗）は、できた結果も開けるようにする
            Set-GuiResult $s.Entry $s.RunArguments $s.RunStarted
            [void](Show-GuiMessage $r.Error -Warning)
        }
        elseif (-not $r.Stopped) { [void](Show-GuiMessage ('エラーで止まりました: ' + $r.Error) -Warning) }
    }
}

# 設定ファイル（Config）で出力先を決めるツールと、その設定を読む関数（ExcelTools.psm1）。出力先の欄が空なら、設定ファイルの出力先を開く
$script:ConfigOutDirReaders = @{ 'New-QualityReport.ps1' = 'Read-QualityReportConfig'; 'New-WeeklyReport.ps1' = 'Read-WeeklyReportConfig' }

function Get-GuiResultTarget {
    #     開く結果のファイル・フォルダ。ツールの一覧の Open の順に、引数にあって、あるもの。フォルダなら、Since（実行を始めた時刻）より
    #     あとにそのフォルダに書いた結果のファイル。Open の引数が無く、設定ファイル（Config）だけで実行したとき（品質評価・週次報告）は、
    #     設定ファイルの出力先（「送る」と同じ）。戻り値: @{ Path（ファイル。無ければ $null）; Dir（フォルダ。無ければ $null） }
    param($Entry, $Arguments, $Since)
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($name in $Entry.Open) {
        if ($Arguments.Contains($name)) { $paths.Add([System.Management.Automation.WildcardPattern]::Unescape([string](@($Arguments[$name])[0]))) }
    }
    $reader = $script:ConfigOutDirReaders[[string]$Entry.Tool]
    if ($reader -and $Arguments.Contains('Config') -and -not @($Entry.Open | Where-Object { $Arguments.Contains($_) }).Count) {
        try {
            Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -DisableNameChecking
            $settings = & $reader ([string](@($Arguments['Config'])[0]))
            if ($settings.OutDir) { $paths.Add([string]$settings.OutDir) }
        }
        catch { }
    }
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p -PathType Leaf) { return [pscustomobject]@{ Path = $p; Dir = (Split-Path -Parent $p) } }
        if (Test-Path -LiteralPath $p -PathType Container) {
            $file = $null
            if ($Since) { $file = Find-GuiNewResult $p $Since }
            return [pscustomobject]@{ Path = $file; Dir = $p }
        }
    }
    return [pscustomobject]@{ Path = $null; Dir = $null }
}

function Set-GuiResult {
    #     結果のファイル・フォルダ（Get-GuiResultTarget）を「結果を開く」「保存先のフォルダを開く」にする
    param($Entry, $Arguments, $Since)
    $s = $script:Gui
    $target = Get-GuiResultTarget $Entry $Arguments $Since
    if ($target.Dir) { $s.ResultPath = $target.Path; $s.ResultDir = $target.Dir }
    $s.OpenButton.Enabled = [bool]$s.ResultPath
    $s.FolderButton.Enabled = [bool]$s.ResultDir
    if ($s.ResultPath) { $s.ResultLabel.Text = $s.ResultPath } elseif ($s.ResultDir) { $s.ResultLabel.Text = $s.ResultDir }
}

function Wait-GuiRun {
    #     実行が終わるまで待つ（テスト用。タイマーの代わりに出力を受け取る）
    param([int] $TimeoutSeconds = 600)
    $s = $script:Gui
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ($s.Run) {
        Update-GuiRun
        [System.Windows.Forms.Application]::DoEvents()
        if ((Get-Date) -gt $deadline) { $null = Stop-GuiToolRun $s.Run; throw 'ツールが時間内に終わりませんでした' }
        Start-Sleep -Milliseconds 50
    }
}

function Open-GuiPath {
    param([string] $Path)
    try { Invoke-Item -LiteralPath $Path } catch { [void](Show-GuiMessage ('開けませんでした: {0}' -f $_.Exception.Message) -Warning) }
}

function Save-GuiCurrentValues {
    #     いまの処理の入力を、次に開いたときのために保存する
    $s = $script:Gui
    if (-not $s.Entry) { return }
    $s.Settings.Tools[$s.Entry.Id] = Get-GuiValuesToSave $s.Entry $s.CurrentFields (Get-GuiValues)
    $s.Settings.LastTool = $s.Entry.Id
    $s.Settings.OpenResult = $s.OpenResult.Checked
    $s.Settings.ShowAdvanced = $s.ShowAdvanced.Checked
    try { Save-GuiSettings $s.Settings $s.SettingsPath } catch { }
}

function Copy-GuiCommand {
    #     同じことをする PowerShell のコマンドをクリップボードに写す（結果の欄が空なら既定の保存先を入れる）
    $s = $script:Gui
    $arguments = Get-GuiRunArguments
    if ($null -eq $arguments) { return }
    $text = Format-GuiCommandLine $s.Entry.ToolPath $arguments
    try { [System.Windows.Forms.Clipboard]::SetText($text); Write-GuiLog 'コマンドをクリップボードに写しました:' 'Gray' }
    catch { Write-GuiLog 'クリップボードに写せませんでした。次のコマンドを写してください:' 'DarkYellow' }
    Write-GuiLog $text
}

function Show-GuiToolHelp {
    #     ツールの説明（Get-Help）を別の画面で見せる
    $s = $script:Gui
    $text = (Get-Help -Name $s.Entry.ToolPath -Full | Out-String -Width 120).Trim()
    $f = New-Object System.Windows.Forms.Form
    $f.Text = 'ツールの説明 - ' + $s.Entry.Label + '（' + $s.Entry.Tool + '）'
    $f.Size = New-Object System.Drawing.Size(900, 700)
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $box = New-Object System.Windows.Forms.TextBox
    $box.Multiline = $true
    $box.ReadOnly = $true
    $box.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $box.WordWrap = $false
    $box.Dock = [System.Windows.Forms.DockStyle]::Fill
    $box.Font = New-Object System.Drawing.Font([System.Drawing.FontFamily]::GenericMonospace, $s.Form.Font.Size)
    $box.Text = $text -replace '\r?\n', "`r`n"
    $f.Controls.Add($box)
    if ($s.Quiet) { $f.Dispose(); return $text }
    [void]$f.ShowDialog($s.Form)
    $f.Dispose()
}

function Close-GuiWindow {
    #     閉じるとき: 実行中なら確かめて止め、入力を保存する
    param($ClosingArgs)
    $s = $script:Gui
    if ($s.Run) {
        if (-not (Show-GuiMessage '実行中です。中止して閉じますか？' -YesNo -Warning)) { $ClosingArgs.Cancel = $true; return }
        # ツールの後片付け（開いた Excel を閉じるなど）が終わるまで待つ（待たずに閉じると、Excel が残ることがある）
        $s.Status.Text = '中止しています…'
        $s.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $s.Form.Update()
        $null = Stop-GuiToolRun $s.Run -WaitSeconds 20
        $s.Run = $null
    }
    $s.Timer.Stop()
    Save-GuiWindowBounds
    Save-GuiCurrentValues
}

function Restore-GuiWindowBounds {
    #     前に閉じたときの画面の位置と大きさにする（どの画面にも入らないなら、真ん中に出す）
    $s = $script:Gui
    $m = [regex]::Match([string]$s.Settings.Window, '^(-?\d+),(-?\d+),(\d+),(\d+)(,max)?$')
    if (-not $m.Success) { return }
    $rect = New-Object System.Drawing.Rectangle([int]$m.Groups[1].Value, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value, [int]$m.Groups[4].Value)
    if ($rect.Width -lt $s.Form.MinimumSize.Width -or $rect.Height -lt $s.Form.MinimumSize.Height) { return }
    if (@([System.Windows.Forms.Screen]::AllScreens | Where-Object { $_.WorkingArea.IntersectsWith($rect) }).Count -eq 0) { return }
    $s.Form.StartPosition = [System.Windows.Forms.FormStartPosition]::Manual
    $s.Form.Bounds = $rect
    if ($m.Groups[5].Success) { $s.Form.WindowState = [System.Windows.Forms.FormWindowState]::Maximized }
}

function Save-GuiWindowBounds {
    #     閉じるときの画面の位置と大きさ（最大化・最小化しているときは、元の大きさ）と、左の一覧の幅を覚える
    $s = $script:Gui
    $f = $s.Form
    $b = $f.Bounds
    if ($f.WindowState -ne [System.Windows.Forms.FormWindowState]::Normal) { $b = $f.RestoreBounds }
    $text = '{0},{1},{2},{3}' -f $b.X, $b.Y, $b.Width, $b.Height
    if ($f.WindowState -eq [System.Windows.Forms.FormWindowState]::Maximized) { $text += ',max' }
    $s.Settings.Window = $text
    $s.Settings.Splitter = $s.Split.SplitterDistance
}

Export-ModuleMember -Function *
