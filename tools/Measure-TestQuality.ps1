[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [string[]] $Path = @('.'),

    [string[]] $BugList,

    [string[]] $DesignPath,

    [string] $SizeCsv,

    [string] $Limits,

    [string] $Headers,

    [ValidateRange(0.5, 1.0)]
    [double] $ConvergenceRate = 0.9,

    [ValidateRange(0, 100000)]
    [int] $MinBugs = 10,

    [ValidateRange(0.0, 1.0)]
    [double] $MinProgress = 0.5,

    [ValidateRange(1, 365)]
    [int] $RecentDays = 7,

    [datetime] $AsOf = (Get-Date).Date,

    [switch] $NoRecurse,

    [switch] $NoExcel,

    [string] $OutXlsx
)

$ErrorActionPreference = 'Stop'

Import-Module (Join-Path $PSScriptRoot 'ExcelTools.psm1') -Force -DisableNameChecking
$searchScript = Join-Path $PSScriptRoot 'Search-Excel.ps1'
$sizeScript = Join-Path $PSScriptRoot 'Measure-DesignSize.ps1'

if (-not $Limits) { $Limits = Join-Path $PSScriptRoot 'test-quality/基準値_サンプル.csv' }
if (-not $Headers) { $Headers = Join-Path $PSScriptRoot 'test-quality/見出し.csv' }
$zoneCsv = Join-Path $PSScriptRoot 'test-quality/ゾーン分析.csv'
foreach ($name in @('Limits', 'Headers', 'SizeCsv', 'OutXlsx')) {
    $v = Get-Variable -Name $name -ValueOnly
    if ($v) { Set-Variable -Name $name -Value $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($v) }
}
if ($OutXlsx -and [System.IO.Path]::GetExtension($OutXlsx) -ne '.xlsx') { throw "-OutXlsx の拡張子は .xlsx にしてください: $OutXlsx" }
if ($DesignPath -and $SizeCsv) { throw '-DesignPath と -SizeCsv は、どちらか一方を指定してください' }

Update-TypeData -TypeName 'TestQuality.Function' -DefaultDisplayPropertySet Function, Cases, Executed, NG, Bugs, OpenBugs, TestDensity, BugDensity, Zone -Force -WhatIf:$false

$invariant = [System.Globalization.CultureInfo]::InvariantCulture

# 障害の状態: 対応が終わった・障害として数えない
$closedPattern = '完了|済|クローズ|close|done|解決|対応不要'
$excludedPattern = '取下げ|取り下げ|取消|却下|対象外|重複|仕様どおり|仕様通り|障害ではない|非障害|バグではない'

# ---------------------------------------------------------------- 設定


$caseFields = @('機能', 'テストID', 'テスト内容', '結果', '予定日', '実施日', 'バグ番号', '工程')
$bugFields = @('機能', '障害ID', '障害内容', '起票日', '完了日', '状態', '重要度', '原因', '混入工程', '摘出工程')
$candidates = Read-HeaderCandidates $Headers ($caseFields + $bugFields)
foreach ($f in @('結果', '障害内容')) { if ($candidates[$f].Count -eq 0) { throw "見出しの候補に「$f」の行がありません: $Headers" } }

# 基準値。指標は「テスト密度」「バグ密度」
$limitRows = New-Object System.Collections.Generic.List[object]
foreach ($r in (Read-CsvFile $Limits @('指標', '分母', '下限', '上限') '基準値')) {
    $unit = ([string]$r.'分母').Trim()
    $metricText = ([string]$r.'指標').Trim()
    if (-not $unit -or -not $metricText) { continue }
    $metric = 'Test'
    if ($metricText -match 'バグ|障害|不具合') { $metric = 'Bug' }
    $process = ''
    if ($r.PSObject.Properties['工程']) { $process = ([string]$r.'工程').Trim() }
    $low = $null; $high = $null
    if (([string]$r.'下限').Trim()) { $low = [double]$r.'下限' }
    if (([string]$r.'上限').Trim()) { $high = [double]$r.'上限' }
    $limitRows.Add([pscustomobject]@{ Metric = $metric; Unit = $unit; Process = $process; Low = $low; High = $high })
}
$zones = Read-ZoneTable $zoneCsv 'テスト密度' 'バグ密度'

# ---------------------------------------------------------------- ヘルパー


function Get-Judgement {
    #     基準値（指標・分母・工程）と比べて「少ない」「基準内」「多い」。基準値が無ければ ''
    param([string] $Metric, [string] $Unit, [string] $Process, [double] $Value)
    $limit = @($limitRows | Where-Object { $_.Metric -eq $Metric -and $_.Unit -eq $Unit -and $Process -and $_.Process -eq $Process })[0]
    if (-not $limit) { $limit = @($limitRows | Where-Object { $_.Metric -eq $Metric -and $_.Unit -eq $Unit -and -not $_.Process })[0] }
    if (-not $limit) { return '' }
    if ($null -ne $limit.Low -and $Value -lt $limit.Low) { return '少ない' }
    if ($null -ne $limit.High -and $Value -gt $limit.High) { return '多い' }
    return '基準内'
}

function Get-GompertzFit {
    #     信頼度成長曲線（ゴンペルツ曲線 y = K × a^(b^x)）を当てはめる。
    #     x を 0〜1 に縮め、b = e^(-k) の k を探し、K と a は ln y = ln K + ln a × b^x の直線の当てはめで決める。
    #     戻り値: @{ K（潜在バグ数）; LnA; Rate（k）; XMax; RSquared; AtLowerBound（k が探す範囲の下限 = まだ増え続けている） }。点が 4 つ未満なら $null
    param([double[]] $X, [double[]] $Y)
    $n = $X.Count
    if ($n -lt 4) { return $null }
    $xMax = ($X | Measure-Object -Maximum).Maximum
    if ($xMax -le 0) { return $null }
    $xs = @($X | ForEach-Object { $_ / $xMax })
    $ln = @($Y | ForEach-Object { [math]::Log($_) })
    $evaluate = {
        param([double] $k)
        $u = @($xs | ForEach-Object { [math]::Exp(-$k * $_) })
        $mu = ($u | Measure-Object -Average).Average
        $ml = ($ln | Measure-Object -Average).Average
        $sxx = 0.0; $sxy = 0.0
        for ($i = 0; $i -lt $n; $i++) { $sxx += ($u[$i] - $mu) * ($u[$i] - $mu); $sxy += ($u[$i] - $mu) * ($ln[$i] - $ml) }
        if ($sxx -le 0) { return $null }
        $c1 = $sxy / $sxx
        if ($c1 -ge 0) { return $null }
        $c0 = $ml - $c1 * $mu
        $sse = 0.0
        for ($i = 0; $i -lt $n; $i++) { $p = [math]::Exp($c0 + $c1 * $u[$i]); $sse += ($Y[$i] - $p) * ($Y[$i] - $p) }
        return [pscustomobject]@{ K = [math]::Exp($c0); LnA = $c1; Rate = $k; Sse = $sse }
    }
    # 粗く探してから、いちばん良い点の両隣の間を黄金分割で詰める
    $lowK = 0.05; $highK = 40.0
    $grid = @(0..60 | ForEach-Object { $lowK * [math]::Pow($highK / $lowK, $_ / 60.0) })
    $best = $null; $bestIndex = -1
    for ($i = 0; $i -lt $grid.Count; $i++) {
        $f = & $evaluate $grid[$i]
        if ($f -and (-not $best -or $f.Sse -lt $best.Sse)) { $best = $f; $bestIndex = $i }
    }
    if (-not $best) { return $null }
    $a = $grid[[math]::Max(0, $bestIndex - 1)]; $b = $grid[[math]::Min($grid.Count - 1, $bestIndex + 1)]
    $phi = (1 + [math]::Sqrt(5)) / 2
    for ($iter = 0; $iter -lt 40; $iter++) {
        $c = $b - ($b - $a) / $phi; $d = $a + ($b - $a) / $phi
        $fc = & $evaluate $c; $fd = & $evaluate $d
        $sc = [double]::MaxValue; $sd = [double]::MaxValue
        if ($fc) { $sc = $fc.Sse }
        if ($fd) { $sd = $fd.Sse }
        if ($sc -lt $sd) { $b = $d } else { $a = $c }
        if ($fc -and $fc.Sse -lt $best.Sse) { $best = $fc }
        if ($fd -and $fd.Sse -lt $best.Sse) { $best = $fd }
    }
    $mean = ($Y | Measure-Object -Average).Average
    $sst = 0.0
    foreach ($value in $Y) { $sst += ($value - $mean) * ($value - $mean) }
    $r2 = 0.0
    if ($sst -gt 0) { $r2 = 1 - $best.Sse / $sst }
    return [pscustomobject]@{ K = $best.K; LnA = $best.LnA; Rate = $best.Rate; XMax = $xMax; RSquared = $r2; AtLowerBound = ($bestIndex -eq 0) }
}

function Get-GompertzValue {
    param($Fit, [double] $X)
    return $Fit.K * [math]::Exp($Fit.LnA * [math]::Exp(-$Fit.Rate * $X / $Fit.XMax))
}

# ---------------------------------------------------------------- テスト仕様書を読む


# 障害管理表を同じフォルダに置いても、テスト仕様書としては読まない
$bugFiles = $null
if ($BugList) { $bugFiles = Get-FilePathSet $BugList -NoRecurse:$NoRecurse }
$readArgs = @{ Candidates = $candidates; Exclude = $OutXlsx; NoRecurse = $NoRecurse; NoExcel = $NoExcel }
$caseRecords = Read-HeaderRecords @readArgs -Inputs (Get-InputFiles $Path 'テスト仕様書' -NoRecurse:$NoRecurse) -Fields $caseFields -Need @('結果', @('テストID', 'テスト内容')) -Key @('テストID', 'テスト内容') -What 'テスト仕様書' -Skip $bugFiles
$cases = ConvertTo-TestCases $caseRecords $AsOf
if ($cases.Count -eq 0) { Write-Warning 'テストケースが見つかりません（結果と、テストIDか確認内容の見出しがある表を探します）' }

# ---------------------------------------------------------------- 障害管理表を読む


$bugs = New-Object System.Collections.Generic.List[object]
$excludedBugs = 0
$bugListHasFunction = $false
if ($BugList) {
    $bugRecords = Read-HeaderRecords @readArgs -Inputs (Get-InputFiles $BugList '障害管理表' -NoRecurse:$NoRecurse) -Fields $bugFields -Need @('障害内容', @('起票日', '状態', '重要度', '障害ID')) -Key @('障害内容') -What '障害管理表'
    foreach ($rec in $bugRecords) {
        $v = $rec.Values
        $status = ([string]$v['状態']).Trim()
        if ($status -match $excludedPattern) { $excludedBugs++; continue }
        if ($rec.Labels['機能']) { $bugListHasFunction = $true }
        $closedOn = ConvertTo-DateValue $v['完了日'] $AsOf
        $bugs.Add([pscustomobject]@{
            Id = @(ConvertTo-BugIds $v['障害ID'])[0]; Function = ([string]$v['機能']).Trim(); Content = ([string]$v['障害内容']).Trim()
            Found = (ConvertTo-DateValue $v['起票日'] $AsOf); ClosedOn = $closedOn; Closed = [bool](($status -match $closedPattern) -or ((-not $status) -and $closedOn))
            Status = $status; Severity = ([string]$v['重要度']).Trim(); Cause = ([string]$v['原因']).Trim()
            Injected = ([string]$v['混入工程']).Trim(); Detected = ([string]$v['摘出工程']).Trim(); Place = $rec
        })
    }
    if ($bugs.Count -gt 0 -and -not $bugListHasFunction) { Write-Warning '障害管理表に機能の列が無いため、機能ごとのバグ密度は出せません（全体の件数・推移だけ出します）' }
}

# ---------------------------------------------------------------- 規模（分母）


$sizes = [ordered]@{}   # 機能の比較用の名前 → @{ Name; Size; Unit }
if ($SizeCsv) {
    $rows = Read-CsvFile $SizeCsv @() '規模の CSV'
    if ($rows.Count -gt 0) {
        $names = @($rows[0].PSObject.Properties | ForEach-Object { $_.Name })
        $nameCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(機能|機能名|機能ID|画面|画面ID|プログラム|モジュール|ファイル|ファイル名|名前|対象|設計書)$' })[0]
        $sizeCol = @($names | Where-Object { (ConvertTo-Label $_) -match '^(KLOC|KSTEP|LOC|ステップ数|ステップ|行数|規模|FP|サイズ|ページ数)$' })[0]
        if (-not $nameCol -or -not $sizeCol) { throw "規模の CSV には、機能（ファイル）の列と、規模（KLOC・ステップ数・FP など）の列が必要です: $SizeCsv" }
        foreach ($r in $rows) {
            $n = 0.0
            if (-not [double]::TryParse(([string]$r.$sizeCol).Trim(), [System.Globalization.NumberStyles]::Float, $invariant, [ref]$n)) { continue }
            $k = ConvertTo-NameKey $r.$nameCol
            if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = ([string]$r.$nameCol).Trim(); Size = 0.0; Unit = (ConvertTo-Label $sizeCol) } }
            $sizes[$k].Size += $n
        }
    }
}
elseif ($DesignPath) {
    $sizeArgs = @{ Path = $DesignPath }
    if ($NoRecurse) { $sizeArgs['NoRecurse'] = $true }
    if ($NoExcel) { $sizeArgs['NoExcel'] = $true }
    foreach ($item in @(& $sizeScript @sizeArgs 6>$null 3>$null)) {
        $k = ConvertTo-NameKey $item.File
        if (-not $sizes.Contains($k)) { $sizes[$k] = [pscustomobject]@{ Name = $item.File; Size = 0.0; Unit = '規模' } }
        $sizes[$k].Size += [double]$item.Size
    }
}

function Find-Key {
    #     名前に対応するキー（同じ名前、またはどちらかがもう一方を含む名前がただ 1 つ）
    param($Keys, [string] $Key)
    if (@($Keys) -contains $Key) { return $Key }
    $hit = @($Keys | Where-Object { $_.Contains($Key) -or $Key.Contains($_) })
    if ($hit.Count -eq 1) { return $hit[0] }
    return $null
}

# ---------------------------------------------------------------- 機能ごとの集計


function Format-Breakdown {
    #     「高 3・中 5」のような内訳（多い順）
    param($Items, [string] $Property)
    $groups = @($Items | Where-Object { $_.$Property } | Group-Object $Property | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)
    return (@($groups | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join '・')
}

$functions = [ordered]@{}
function Get-FunctionEntry {
    param([string] $Key, [string] $Name)
    if (-not $functions.Contains($Key)) {
        $functions[$Key] = [pscustomobject]@{ Name = $Name; SizeKey = (Find-Key @($sizes.Keys) $Key); Cases = (New-Object System.Collections.Generic.List[object]); Bugs = (New-Object System.Collections.Generic.List[object]) }
    }
    return $functions[$Key]
}
foreach ($c in $cases) { (Get-FunctionEntry $c.Key $c.Function).Cases.Add($c) }
foreach ($b in $bugs) {
    if (-not $bugListHasFunction) { continue }
    $name = $b.Function
    if (-not $name) { $name = '（機能なし）' }
    $key = Find-Key @($functions.Keys) (ConvertTo-NameKey $name)
    if (-not $key) { $key = ConvertTo-NameKey $name }
    (Get-FunctionEntry $key $name).Bugs.Add($b)
}
foreach ($k in $sizes.Keys) {
    if (@($functions.Values | Where-Object { $_.SizeKey -eq $k }).Count -eq 0) { $null = Get-FunctionEntry $k $sizes[$k].Name }
}

$results = New-Object System.Collections.Generic.List[object]
foreach ($k in $functions.Keys) {
    $fn = $functions[$k]
    $counted = @($fn.Cases | Where-Object { $_.Kind -ne '対象外' })
    $executed = @($counted | Where-Object { $_.Kind -eq 'OK' -or $_.Kind -eq 'NG' }).Count
    $size = $null; $unit = ''
    if ($fn.SizeKey) { $size = $sizes[$fn.SizeKey].Size; $unit = $sizes[$fn.SizeKey].Unit }
    # 基準値に使う工程は、ケースの多い工程（同じ数なら先に出てきた工程。Windows PowerShell 5.1 の Sort-Object は同じ値の順番を保たないため並べ替えない）
    $process = ''; $most = 0
    foreach ($g in @($fn.Cases | Where-Object { $_.Process } | Group-Object Process)) { if ($g.Count -gt $most) { $most = $g.Count; $process = $g.Name } }
    $testDensity = $null; $bugDensity = $null; $testJudgement = ''; $bugJudgement = ''; $zone = $null
    if ($null -ne $size -and $size -gt 0) {
        $testDensity = [math]::Round($counted.Count / $size, 3)
        $testJudgement = Get-Judgement 'Test' $unit $process $testDensity
        if ($bugListHasFunction) {
            $bugDensity = [math]::Round($fn.Bugs.Count / $size, 3)
            $bugJudgement = Get-Judgement 'Bug' $unit $process $bugDensity
        }
        if ($testJudgement -and $bugJudgement) { $zone = $zones[('{0}|{1}' -f $testJudgement, $bugJudgement)] }
    }
    $progress = $null
    if ($counted.Count -gt 0) { $progress = [math]::Round($executed / $counted.Count, 3) }
    $r = [pscustomobject]@{
        Function = $fn.Name; Cases = $counted.Count; Executed = $executed
        OK = @($counted | Where-Object Kind -eq 'OK').Count; NG = @($counted | Where-Object Kind -eq 'NG').Count
        Hold = @($counted | Where-Object Kind -eq '保留').Count; NotRun = @($counted | Where-Object Kind -eq '未実施').Count
        Skipped = @($fn.Cases | Where-Object Kind -eq '対象外').Count; Progress = $progress
        Bugs = $fn.Bugs.Count; OpenBugs = @($fn.Bugs | Where-Object { -not $_.Closed }).Count
        Size = $size; Unit = $unit; Process = $process
        TestDensity = $testDensity; TestJudgement = $testJudgement; BugDensity = $bugDensity; BugJudgement = $bugJudgement
        Zone = $(if ($zone) { $zone.Zone } else { '' }); ZoneView = $(if ($zone) { $zone.View } else { '' }); ZoneAction = $(if ($zone) { $zone.Action } else { '' })
        NeedsAction = [bool]($zone -and $zone.NeedsAction)
    }
    $r.PSObject.TypeNames.Insert(0, 'TestQuality.Function')
    $results.Add($r)
}
$results

# ---------------------------------------------------------------- 推移と信頼度成長曲線


$countedCases = @($cases | Where-Object { $_.Kind -ne '対象外' })
$dates = @(@($countedCases | ForEach-Object { $_.Planned; $_.Done }) + @($bugs | ForEach-Object { $_.Found; $_.ClosedOn }) | Where-Object { $_ } | Sort-Object -Unique)
# 日ごとの件数を数えてから累計にする（ケースが多くても日数 × ケース数の繰り返しにしない）
$perDay = @{}
function Add-DayCount {
    param($Day, [string] $Name)
    if (-not $Day) { return }
    if (-not $perDay.ContainsKey($Day)) { $perDay[$Day] = @{ Planned = 0; Executed = 0; Passed = 0; Found = 0; Closed = 0 } }
    $perDay[$Day][$Name]++
}
foreach ($c in $countedCases) {
    Add-DayCount $c.Planned 'Planned'
    if ($c.Kind -eq 'OK' -or $c.Kind -eq 'NG') { Add-DayCount $c.Done 'Executed' }
    if ($c.Kind -eq 'OK') { Add-DayCount $c.Done 'Passed' }
}
foreach ($b in $bugs) {
    Add-DayCount $b.Found 'Found'
    if ($b.Closed) { Add-DayCount $b.ClosedOn 'Closed' }
}
$trend = New-Object System.Collections.Generic.List[object]
$sum = @{ Planned = 0; Executed = 0; Passed = 0; Found = 0; Closed = 0 }
foreach ($day in $dates) {
    $count = $perDay[$day]
    foreach ($k in @($sum.Keys)) { $sum[$k] += $count[$k] }
    $trend.Add([pscustomobject]@{
        Date = $day; Planned = $sum.Planned; Executed = $sum.Executed; Passed = $sum.Passed
        Found = $count.Found; Closed = $count.Closed; TotalFound = $sum.Found; TotalClosed = $sum.Closed; Fitted = $null
    })
}

# 横軸はテストの実施数（実施日のあるケースがあれば）、無ければテストの日数
$useCases = @($trend | Where-Object { $_.Executed -gt 0 }).Count -gt 0
$axisName = 'テストの日数'
if ($useCases) { $axisName = '実施したケースの数' }
$points = [ordered]@{}
if ($trend.Count -gt 0) {
    $first = $trend[0].Date
    foreach ($t in $trend) {
        $x = ($t.Date - $first).TotalDays + 1
        if ($useCases) { $x = $t.Executed }
        if ($x -gt 0 -and $t.TotalFound -gt 0) { $points[[string]$x] = @([double]$x, [double]$t.TotalFound) }
    }
}
$fit = $null
if ($points.Count -ge 4) { $fit = Get-GompertzFit @($points.Values | ForEach-Object { $_[0] }) @($points.Values | ForEach-Object { $_[1] }) }
$totalFound = $bugs.Count
$growth = [pscustomobject]@{ Axis = $axisName; Points = $points.Count; Potential = $null; Rate = $null; Remaining = $null; RSquared = $null; Judgement = ''; Recent = $null; Before = $null }
if (-not $BugList) { $growth.Judgement = '' }
elseif (-not $fit) { $growth.Judgement = '判定できません（起票日のあるバグが少ないため、曲線を当てはめられません）' }
elseif ($fit.AtLowerBound -or $fit.K -gt $totalFound * 10) {
    $growth.Judgement = '収束していない（バグがまだ増え続けており、上限を推定できません）'
    $growth.RSquared = [math]::Round($fit.RSquared, 3)
}
else {
    $growth.Potential = [math]::Round($fit.K, 1)
    $growth.Rate = [math]::Round($totalFound / $fit.K, 3)
    $growth.Remaining = [math]::Round([math]::Max(0.0, $fit.K - $totalFound), 1)
    $growth.RSquared = [math]::Round($fit.RSquared, 3)
    if ($growth.Rate -ge $ConvergenceRate) { $growth.Judgement = '収束している' }
    else { $growth.Judgement = '収束していない' }
    foreach ($t in $trend) {
        $x = ($t.Date - $trend[0].Date).TotalDays + 1
        if ($useCases) { $x = $t.Executed }
        if ($x -gt 0) { $t.Fitted = [math]::Round((Get-GompertzValue $fit $x), 1) }
    }
}

# バグが少ないうち・テストの消化が進まないうちは曲線の形が定まらないため、収束を判定しない
$progressAll = 0.0
if ($countedCases.Count -gt 0) { $progressAll = @($countedCases | Where-Object { $_.Kind -eq 'OK' -or $_.Kind -eq 'NG' }).Count / $countedCases.Count }
if ($BugList -and ($totalFound -lt $MinBugs -or $progressAll -lt $MinProgress)) {
    $growth.Judgement = 'まだ判定できない（バグ {0} 件・消化率 {1:0}%。バグ {2} 件以上・消化率 {3:0}% 以上で判定します）' -f $totalFound, ($progressAll * 100), $MinBugs, ($MinProgress * 100)
}
# 直近の新規バグ（曲線とは別に、増え方が落ち着いてきたかを見る）
if ($BugList) {
    $recentFrom = $AsOf.AddDays(-$RecentDays)
    $growth.Recent = @($bugs | Where-Object { $_.Found -and $_.Found -gt $recentFrom -and $_.Found -le $AsOf }).Count
    $growth.Before = @($bugs | Where-Object { $_.Found -and $_.Found -gt $recentFrom.AddDays(-$RecentDays) -and $_.Found -le $recentFrom }).Count
}

# ---------------------------------------------------------------- 食い違い


$issues = New-Object System.Collections.Generic.List[object]
function Add-Issue {
    param([string] $Kind, $Case, [string] $Detail)
    $issues.Add([pscustomobject]@{ Kind = $Kind; Function = $Case.Function; Id = $Case.Id; Detail = $Detail; Place = $Case.Place })
}
$bugById = @{}
foreach ($b in $bugs) { if ($b.Id -and -not $bugById.ContainsKey($b.Id)) { $bugById[$b.Id] = $b } }
$bugListHasId = $bugById.Count -gt 0
foreach ($c in $countedCases) {
    if ($c.Kind -eq 'NG' -and $c.HasBugColumn -and $c.BugIds.Count -eq 0) { Add-Issue 'NG なのにバグ番号が無い' $c '結果が NG のケースに、バグ（障害）の番号を書いてください' }
    foreach ($id in $c.BugIds) {
        if ($bugListHasId -and -not $bugById.ContainsKey($id)) { Add-Issue '障害管理表に無いバグ番号' $c ("バグ番号 {0} が障害管理表にありません" -f $id) }
        elseif ($c.Kind -eq 'NG' -and $bugById.ContainsKey($id) -and $bugById[$id].Closed) { Add-Issue '再テストの結果が無い' $c ("バグ {0} は完了していますが、ケースの結果が NG のままです（再テストの結果を書いてください）" -f $id) }
    }
    if (($c.Kind -eq 'OK' -or $c.Kind -eq 'NG') -and $c.HasDoneColumn -and -not $c.Done) { Add-Issue '実施日が無い' $c '結果がありますが、実施日がありません' }
    if (($c.Kind -eq '未実施' -or $c.Kind -eq '保留') -and $c.Planned -and $c.Planned -lt $AsOf) { Add-Issue '予定日を過ぎても未実施' $c ("予定日 {0} を過ぎています（{1}）" -f $c.Planned.ToString('yyyy/MM/dd'), $c.Kind) }
}

# ---------------------------------------------------------------- 出力


if ($OutXlsx) {
    $executedAll = @($countedCases | Where-Object { $_.Kind -eq 'OK' -or $_.Kind -eq 'NG' }).Count
    $plannedByNow = @($countedCases | Where-Object { $_.Planned -and $_.Planned -le $AsOf }).Count
    $summary = New-Object System.Collections.Generic.List[object]
    $add = { param($name, $value) $summary.Add(@{ Values = @($name, $value) }) }
    & $add '時点' $AsOf.ToString('yyyy/MM/dd')
    & $add 'テストケース（対象外を除く）' ([double]$countedCases.Count)
    & $add '予定（時点まで）' ([double]$plannedByNow)
    & $add '実施（OK・NG）' ([double]$executedAll)
    & $add 'OK' ([double]@($countedCases | Where-Object Kind -eq 'OK').Count)
    & $add 'NG' ([double]@($countedCases | Where-Object Kind -eq 'NG').Count)
    & $add '保留' ([double]@($countedCases | Where-Object Kind -eq '保留').Count)
    & $add '未実施' ([double]@($countedCases | Where-Object Kind -eq '未実施').Count)
    if ($countedCases.Count -gt 0) { & $add '消化率' ([double][math]::Round($executedAll / $countedCases.Count, 3)) }
    if ($BugList) {
        & $add 'バグ' ([double]$bugs.Count)
        & $add '未完了のバグ' ([double]@($bugs | Where-Object { -not $_.Closed }).Count)
        & $add '信頼度成長曲線の横軸' $growth.Axis
        if ($null -ne $growth.Potential) {
            & $add '潜在バグ数（推定）' ([double]$growth.Potential)
            & $add '収束率（バグ ÷ 潜在バグ数）' ([double]$growth.Rate)
            & $add '残りのバグ（推定）' ([double]$growth.Remaining)
        }
        if ($null -ne $growth.RSquared) { & $add '当てはまり（決定係数）' ([double]$growth.RSquared) }
        & $add '収束の判定' ('{0}（基準: 収束率 {1}）' -f $growth.Judgement, $ConvergenceRate)
        & $add ('直近 {0} 日の新規バグ' -f $RecentDays) ([double]$growth.Recent)
        & $add ('その前の {0} 日の新規バグ' -f $RecentDays) ([double]$growth.Before)
    }
    & $add '食い違い' ([double]$issues.Count)

    $num = { param($v) if ($null -eq $v) { '' } else { [double]$v } }
    $funcRows = New-Object System.Collections.Generic.List[object]
    foreach ($r in $results) {
        $funcRows.Add(@{ Values = @($r.Function, [double]$r.Cases, [double]$r.Executed, [double]$r.OK, [double]$r.NG, [double]$r.Hold, [double]$r.NotRun, (& $num $r.Progress), [double]$r.Bugs, [double]$r.OpenBugs, (& $num $r.Size), $r.Unit, (& $num $r.TestDensity), $r.TestJudgement, (& $num $r.BugDensity), $r.BugJudgement, $r.Zone, $r.ZoneView, $r.ZoneAction) })
    }
    $zoneRows = New-Object System.Collections.Generic.List[object]
    foreach ($testLevel in @('多い', '基準内', '少ない')) {
        $values = @("テスト密度が$testLevel")
        foreach ($bugLevel in @('少ない', '基準内', '多い')) {
            $z = $zones[('{0}|{1}' -f $testLevel, $bugLevel)]
            $lines = New-Object System.Collections.Generic.List[string]
            if ($z) { $lines.Add('【' + $z.Zone + '】') }
            foreach ($r in $results) {
                if ($r.TestJudgement -eq $testLevel -and $r.BugJudgement -eq $bugLevel) { $lines.Add(('{0}（テスト {1} / バグ {2}）' -f $r.Function, $r.TestDensity, $r.BugDensity)) }
            }
            $values += ($lines -join "`n")
        }
        $zoneRows.Add(@{ Values = $values })
    }
    $trendRows = New-Object System.Collections.Generic.List[object]
    foreach ($t in $trend) {
        $trendRows.Add(@{ Values = @($t.Date.ToString('yyyy/MM/dd'), [double]$t.Planned, [double]$t.Executed, [double]$t.Passed, [double]$t.Found, [double]$t.Closed, [double]$t.TotalFound, [double]$t.TotalClosed, [double]($t.TotalFound - $t.TotalClosed), (& $num $t.Fitted)) })
    }
    $breakRows = New-Object System.Collections.Generic.List[object]
    foreach ($p in @(@('重要度', 'Severity'), @('原因', 'Cause'), @('混入工程', 'Injected'), @('摘出工程', 'Detected'), @('状態', 'Status'), @('機能', 'Function'))) {
        foreach ($g in @($bugs | Where-Object { $_.($p[1]) } | Group-Object $p[1] | Sort-Object @{ e = { $_.Count }; Descending = $true }, Name)) {
            $breakRows.Add(@{ Values = @($p[0], $g.Name, [double]$g.Count, [double]@($g.Group | Where-Object { -not $_.Closed }).Count) })
        }
    }
    $issueRows = New-Object System.Collections.Generic.List[object]
    foreach ($i in $issues) {
        $where = ('{0} {1} 行目' -f $i.Place.File, $i.Place.Row)
        if ($i.Place.Sheet) { $where = ('{0} {1}!{2}{3}' -f $i.Place.File, $i.Place.Sheet, (ConvertTo-ColumnName $i.Place.Column), $i.Place.Row) }
        $row = @{ Values = @($i.Kind, $i.Function, $i.Id, $i.Detail, $where) }
        if ($i.Place.Sheet) { $row.Link = @{ Column = 4; Path = $i.Place.Path; Location = ("'{0}'!{1}{2}" -f $i.Place.Sheet.Replace("'", "''"), (ConvertTo-ColumnName $i.Place.Column), $i.Place.Row) } }
        $issueRows.Add($row)
    }
    try {
        Export-Workbook -FilePath $OutXlsx -Sheets @(
            @{ Name = 'サマリー'; Columns = @(@{ Header = '項目'; Width = 30 }, @{ Header = '値'; Width = 50 }); Rows = $summary },
            @{ Name = '機能別'; Columns = @(@{ Header = '機能'; Width = 30 }, @{ Header = 'ケース'; Width = 7 }, @{ Header = '実施'; Width = 7 }, @{ Header = 'OK'; Width = 6 }, @{ Header = 'NG'; Width = 6 }, @{ Header = '保留'; Width = 6 }, @{ Header = '未実施'; Width = 7 }, @{ Header = '消化率'; Width = 7 }, @{ Header = 'バグ'; Width = 6 }, @{ Header = '未完了のバグ'; Width = 8 }, @{ Header = '規模'; Width = 8 }, @{ Header = '分母'; Width = 8 }, @{ Header = 'テスト密度'; Width = 9 }, @{ Header = 'テスト密度の判定'; Width = 9 }, @{ Header = 'バグ密度'; Width = 9 }, @{ Header = 'バグ密度の判定'; Width = 9 }, @{ Header = 'ゾーン'; Width = 18 }, @{ Header = '見方'; Width = 40 }, @{ Header = '対応の例'; Width = 40 }); Rows = $funcRows },
            @{ Name = 'ゾーン分析'; Columns = @(@{ Header = 'テスト密度 ＼ バグ密度'; Width = 18 }, @{ Header = 'バグ密度が少ない'; Width = 40 }, @{ Header = 'バグ密度が基準内'; Width = 40 }, @{ Header = 'バグ密度が多い'; Width = 40 }); Rows = $zoneRows },
            @{ Name = '推移'; Columns = @(@{ Header = '日付'; Width = 11 }, @{ Header = '予定（累計）'; Width = 10 }, @{ Header = '実施（累計）'; Width = 10 }, @{ Header = 'OK（累計）'; Width = 10 }, @{ Header = 'バグの発見'; Width = 9 }, @{ Header = 'バグの完了'; Width = 9 }, @{ Header = '発見の累計'; Width = 9 }, @{ Header = '完了の累計'; Width = 9 }, @{ Header = '未完了のバグ'; Width = 9 }, @{ Header = '信頼度成長曲線'; Width = 11 }); Rows = $trendRows },
            @{ Name = 'バグの内訳'; Columns = @(@{ Header = '項目'; Width = 10 }, @{ Header = '値'; Width = 24 }, @{ Header = 'バグ'; Width = 7 }, @{ Header = '未完了'; Width = 7 }); Rows = $breakRows },
            @{ Name = '食い違い'; Columns = @(@{ Header = '種類'; Width = 22 }, @{ Header = '機能'; Width = 24 }, @{ Header = 'テストID'; Width = 10 }, @{ Header = '内容'; Width = 60 }, @{ Header = '場所'; Width = 40 }); Rows = $issueRows })
    }
    catch {
        Write-Warning ("結果を保存できませんでした（同名のファイルを Excel で開いていませんか）: {0} : {1}" -f $OutXlsx, $_.Exception.Message)
        $OutXlsx = $null
    }
}

Write-Host ''
$executedAll = @($countedCases | Where-Object { $_.Kind -eq 'OK' -or $_.Kind -eq 'NG' }).Count
$rateText = ''
if ($countedCases.Count -gt 0) { $rateText = '消化率 {0:0.0}%' -f ($executedAll * 100.0 / $countedCases.Count) }
Write-Host ('テストケース {0} 件（実施 {1}・OK {2}・NG {3}・保留 {4}・未実施 {5}。対象外 {6}）{7}（{8} 時点）' -f $countedCases.Count, $executedAll, @($countedCases | Where-Object Kind -eq 'OK').Count, @($countedCases | Where-Object Kind -eq 'NG').Count, @($countedCases | Where-Object Kind -eq '保留').Count, @($countedCases | Where-Object Kind -eq '未実施').Count, @($cases | Where-Object Kind -eq '対象外').Count, $rateText, $AsOf.ToString('yyyy/MM/dd'))
if ($BugList) {
    Write-Host ('バグ {0} 件（未完了 {1} 件。取下げなどで数えなかったもの {2} 件）' -f $bugs.Count, @($bugs | Where-Object { -not $_.Closed }).Count, $excludedBugs)
    if ($null -ne $growth.Potential) {
        Write-Host ('信頼度成長曲線（横軸: {0}）: 潜在バグ数 {1} 件（推定）・収束率 {2:0.0}% → {3}（基準 {4:0}%）' -f $growth.Axis, $growth.Potential, ($growth.Rate * 100), $growth.Judgement, ($ConvergenceRate * 100))
    }
    else { Write-Host ('信頼度成長曲線: {0}' -f $growth.Judgement) }
    Write-Host ('直近 {0} 日の新規バグ: {1} 件（その前の {0} 日: {2} 件）' -f $RecentDays, $growth.Recent, $growth.Before)
}
$zoned = @($results | Where-Object { $_.Zone })
if ($zoned.Count -gt 0) { Write-Host ('ゾーン: {0}' -f (Format-Breakdown $zoned 'Zone')) }
foreach ($r in @($results | Where-Object NeedsAction)) { Write-Host ('  要対応: {0}（{1}）{2}' -f $r.Function, $r.Zone, $r.ZoneAction) }
if ($issues.Count -gt 0) { Write-Host ('食い違い: {0} 件（{1}）' -f $issues.Count, (Format-Breakdown $issues 'Kind')) }
if ($OutXlsx) { Write-Host "結果を保存しました: $OutXlsx" }

# .SYNOPSIS
#     テスト仕様書（テストケースの結果）と障害管理表を集計し、機能ごとのテスト密度・バグ密度と、バグの収束（信頼度成長曲線）を出します（Excel 不要）。
#
# .DESCRIPTION
#     テスト仕様書・障害管理表（Excel・CSV）から、結果・実施日・バグ番号・起票日・完了日などの列を見出しの文字で探して読み、次を出します。
#
#       機能別       テストケース数・実施・OK・NG・消化率、バグ数、規模あたりのテスト密度・バグ密度と基準値の判定
#       ゾーン分析   テスト密度 × バグ密度の 9 つのゾーン（テスト不足・品質に問題 など）と対応の例
#       推移         日ごとの予定・実施・OK の累計（テスト消化）、バグの発見・完了と未完了の残り（バグの収束）
#       信頼度成長曲線  バグの累計にゴンペルツ曲線を当てはめ、潜在バグ数（推定）と収束率から「収束している」かを判定
#       食い違い     NG なのにバグ番号が無い・障害管理表に無いバグ番号・バグは完了したのに NG のまま・実施日が無い・予定日を過ぎても未実施
#
#     テストケースの機能は、機能の列（空なら上の行の値）、無ければシート名（表のあるシートが複数のブック）かファイル名です。
#     テスト密度・バグ密度の分母は、-SizeCsv（KLOC・ステップ数・FP など）か、-DesignPath（Measure-DesignSize.ps1 の規模）で渡します。
#
# .PARAMETER Path
#     テスト仕様書のファイルまたはフォルダ（Excel・CSV。複数可・ワイルドカード可）。既定はカレントフォルダ。
#
# .PARAMETER BugList
#     障害管理表のファイルまたはフォルダ（Excel・CSV。複数可）。指定しなければテストの消化だけを出します。
#
# .PARAMETER DesignPath
#     設計書のフォルダ・ファイル。Measure-DesignSize.ps1 で数えた規模を分母にします（機能の名前と設計書のファイル名で対応付けます）。
#
# .PARAMETER SizeCsv
#     機能ごとの規模の CSV（機能の列と、KLOC・ステップ数・FP・規模などの列）。見出しの名前が分母の名前になります。
#
# .PARAMETER Limits
#     基準値の CSV（指標・分母・工程・下限・上限）。指標は「テスト密度」か「バグ密度」。既定は tools\test-quality\基準値_サンプル.csv（数値は例です）。
#
# .PARAMETER Headers
#     テスト仕様書・障害管理表の見出しの候補の CSV（項目と、見出しの候補）。既定は tools\test-quality\見出し.csv。
#
# .PARAMETER ConvergenceRate
#     収束とみなす収束率（バグの累計 ÷ 潜在バグ数の推定）。既定は 0.9。
#
# .PARAMETER MinBugs
#     収束を判定するのに要るバグの数（既定 10）。これより少ないうちは「まだ判定できない」とします。
#
# .PARAMETER MinProgress
#     収束を判定するのに要る消化率（既定 0.5）。テストの前半は「まだ判定できない」とします。
#
# .PARAMETER RecentDays
#     直近の新規バグを数える日数（既定 7）。直近とその前の同じ日数の新規バグを並べて出します。
#
# .PARAMETER AsOf
#     集計の時点（予定日を過ぎても未実施の判定などに使う）。既定は今日。
#
# .PARAMETER NoRecurse
#     フォルダを指定したとき、サブフォルダを見ません。
#
# .PARAMETER NoExcel
#     Excel を起動しません（.xls / .xlsb は読まなくなります）。
#
# .PARAMETER OutXlsx
#     結果を xlsx に保存します（サマリー・機能別・ゾーン分析・推移・バグの内訳・食い違い）。
#
# .EXAMPLE
#     .\Measure-TestQuality.ps1 D:\テスト\テスト仕様書 -BugList D:\テスト\障害管理表.xlsx -SizeCsv D:\規模.csv -OutXlsx D:\テスト品質.xlsx
#
# .EXAMPLE
#     # 設計書の規模（Measure-DesignSize.ps1）を分母にする
#     .\Measure-TestQuality.ps1 D:\テスト -BugList D:\テスト\障害管理表.xlsx -DesignPath D:\設計書
#
# .NOTES
#     要件: PowerShell 5.1 以降。.xlsx / .xlsm は Excel 不要。同じフォルダに ExcelTools.psm1・Search-Excel.ps1（-DesignPath では Measure-DesignSize.ps1 も）が必要です。
#     基準値のサンプルの数値は例です。社内の実績・標準値に合わせて直してください。
#     信頼度成長曲線は推定です。バグの起票日が少ない時期や、テストの途中で大きく増え続けている時期は、判定できないことがあります。
