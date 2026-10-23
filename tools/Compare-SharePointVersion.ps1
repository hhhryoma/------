[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string] $Path,

    [Parameter(Position = 1)]
    [string] $Version,

    [string] $NewVersion,

    [string] $SiteUrl,

    [ValidateSet('Auto', 'Excel', 'Browser')]
    [string] $Method = 'Auto',

    [string] $OutDir,

    [string] $OutHtml,

    [string] $OutXlsx,

    [switch] $CompareFormat,

    [switch] $CheckHistory,

    [switch] $NoOpen,

    [Parameter(DontShow = $true)]
    [scriptblock] $Fetcher,

    [Parameter(DontShow = $true)]
    [object[]] $SyncRoots
)

$ErrorActionPreference = 'Stop'

# -Fetcher・-SyncRoots はテスト用（版を取り出す処理 param($Url, $Dest) と、OneDrive で同期しているフォルダの一覧）

$compareScript = Join-Path $PSScriptRoot 'Compare-Excel.ps1'
$downloadWait = 180   # ブラウザーでダウンロードするとき、ファイルができるのを待つ秒数

# ---------------------------------------------------------------- SharePoint の場所


function Get-SyncRoot {
    #     OneDrive で同期している SharePoint のフォルダ（ローカルのフォルダ → SharePoint の URL）。
    #     HKCU\Software\SyncEngines\Providers\OneDrive の下のキーごとに MountPoint と UrlNamespace がある
    if ($null -ne $SyncRoots) { return , @($SyncRoots) }
    $key = 'HKCU:\Software\SyncEngines\Providers\OneDrive'
    $roots = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $key)) { return , $roots.ToArray() }
    foreach ($k in (Get-ChildItem -LiteralPath $key -ErrorAction SilentlyContinue)) {
        $p = Get-ItemProperty -LiteralPath $k.PSPath -ErrorAction SilentlyContinue
        if ($p -and $p.MountPoint -and $p.UrlNamespace) {
            $web = $null
            if ($p.PSObject.Properties['WebUrl']) { $web = $p.WebUrl }
            $roots.Add([pscustomobject]@{ MountPoint = $p.MountPoint; UrlNamespace = $p.UrlNamespace; WebUrl = $web })
        }
    }
    return , $roots.ToArray()
}

function ConvertTo-SharePointLocation {
    #     SharePoint のファイルの URL を、サイト（Web）の場所とその中のファイルの場所に分ける。
    #     返り値: Origin（https://会社.sharepoint.com）・Web（/sites/サイト）・File（Shared Documents/設計/a.xlsx）・Name
    param([string] $Url, [string] $Site)

    $uri = $null
    if (-not [Uri]::TryCreate($Url.Trim(), [UriKind]::Absolute, [ref]$uri) -or @('http', 'https') -notcontains $uri.Scheme) {
        throw "SharePoint の URL ではありません: $Url"
    }
    $origin = '{0}://{1}' -f $uri.Scheme, $uri.Authority
    $path = [Uri]::UnescapeDataString($uri.AbsolutePath)
    $guide = 'Excel でファイルを開き、「ファイル」→「情報」→「パスのコピー」で URL をコピーしてください'
    # 共有リンク（/:x:/s/…・/:x:/g/…）や Doc.aspx の URL からは、ファイルの場所が分からない
    if ($path -match '^/:[a-z]:/r(/.*)$') { $path = $Matches[1] }
    elseif ($path -match '^/:[a-z]:/') { throw "共有リンクからはファイルの場所が分かりません。$guide`: $Url" }
    if ($path -match '/_layouts/') { throw "ブラウザーで開いたときの URL からはファイルの場所が分かりません。$guide`: $Url" }
    $path = $path.TrimEnd('/')
    if ($path -notmatch '/[^/]+\.[A-Za-z0-9]+$') { throw "ファイルの URL ではありません（ファイル名がありません）。$guide`: $Url" }

    # サイトは、指定が無ければ /sites/名前・/teams/名前・/personal/名前（無ければルートのサイト）とみなす
    $web = ''
    if ($Site) {
        $su = $null
        if (-not [Uri]::TryCreate($Site.Trim(), [UriKind]::Absolute, [ref]$su)) { throw "-SiteUrl が URL ではありません: $Site" }
        $web = [Uri]::UnescapeDataString($su.AbsolutePath).TrimEnd('/')
        if ($web -and -not $path.StartsWith($web + '/', [StringComparison]::OrdinalIgnoreCase)) { throw "ファイルの URL が -SiteUrl のサイトの中にありません: $Url" }
    }
    else {
        $parts = $path.TrimStart('/').Split('/')
        if ($parts.Count -ge 3 -and @('sites', 'teams', 'personal') -contains $parts[0].ToLowerInvariant()) { $web = '/' + $parts[0] + '/' + $parts[1] }
    }
    $file = $path.Substring($web.Length).TrimStart('/')
    return [pscustomobject]@{ Origin = $origin; Web = $web; File = $file; Name = $file.Split('/')[-1] }
}

function Resolve-SharePointFile {
    #     URL か、OneDrive で同期しているフォルダのファイルから、SharePoint の場所を求める。LocalPath は同期しているファイル
    param([string] $Target, [string] $Site)

    if ($Target -match '^https?://') {
        $loc = ConvertTo-SharePointLocation $Target $Site
        $loc | Add-Member -NotePropertyName LocalPath -NotePropertyValue $null
        return $loc
    }
    $full = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Target)
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) { throw "ファイルが見つかりません: $full" }
    # 一番深い（長い）同期フォルダを使う
    $best = $null
    foreach ($r in (Get-SyncRoot)) {
        $mount = ([string]$r.MountPoint).TrimEnd('\', '/')
        if (-not $mount) { continue }
        if ($full.StartsWith($mount + '\', [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($mount + '/', [StringComparison]::OrdinalIgnoreCase)) {
            if (-not $best -or $mount.Length -gt $best.Mount.Length) { $best = [pscustomobject]@{ Mount = $mount; Root = $r } }
        }
    }
    if (-not $best) {
        throw "OneDrive で同期している SharePoint のフォルダのファイルではありません。SharePoint の URL を指定してください（Excel でファイルを開き、「ファイル」→「情報」→「パスのコピー」）: $full"
    }
    $rel = $full.Substring($best.Mount.Length).TrimStart('\', '/').Replace('\', '/')
    $url = ([string]$best.Root.UrlNamespace).TrimEnd('/') + '/' + $rel
    if (-not $Site -and $best.Root.WebUrl) { $Site = $best.Root.WebUrl }
    $loc = ConvertTo-SharePointLocation $url $Site
    $loc | Add-Member -NotePropertyName LocalPath -NotePropertyValue $full
    return $loc
}

function ConvertTo-VersionLabel {
    #     版の指定（3・3.0・v3.0・３．０）を「3.0」の形にする。現在の版（空・現在・current・latest）は $null
    param([string] $Text)
    if ($null -eq $Text) { return $null }
    $t = $Text.Normalize([System.Text.NormalizationForm]::FormKC).Trim()
    if (-not $t -or @('現在', '最新', 'current', 'latest') -contains $t.ToLowerInvariant()) { return $null }
    $m = [regex]::Match($t, '^(?:v(?:er)?\.?\s*)?(\d+)(?:\.(\d+))?$', 'IgnoreCase')
    if (-not $m.Success) { throw "版の番号は SharePoint の「バージョン履歴」の番号（3.0 など）で指定してください: $Text" }
    $minor = 0
    if ($m.Groups[2].Success) { $minor = [int]$m.Groups[2].Value }
    if ($minor -ge 512) { throw "版の小さい番号は 511 までです: $Text" }
    return '{0}.{1}' -f [int]$m.Groups[1].Value, $minor
}

function Get-VersionUrl {
    #     版の URL。SharePoint は版を「_vti_history/番号/ファイルの場所」に置き、番号は 大きい番号 × 512 + 小さい番号。
    #     $Label が $null なら現在の版（ファイルそのもの）
    param($Location, [string] $Label)
    $escape = { param($p) (@($p.Split('/') | Where-Object { $_ } | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/') }
    $web = & $escape $Location.Web
    $base = $Location.Origin + '/'
    if ($web) { $base += $web + '/' }
    if (-not $Label) { return $base + (& $escape $Location.File) }
    $x = $Label.Split('.')
    $id = [int]$x[0] * 512 + [int]$x[1]
    return '{0}_vti_history/{1}/{2}' -f $base, $id, (& $escape $Location.File)
}

# ---------------------------------------------------------------- 版を取り出す


function Save-VersionByExcel {
    #     Excel で URL のブックを読み取り専用で開き、写しを保存する（Office にサインインしているアカウントで SharePoint を読む）。
    #     取り出せなかったものは Error に理由を入れる
    param([object[]] $Items)

    $pidsBefore = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    try { $excel = New-Object -ComObject Excel.Application }
    catch {
        foreach ($it in $Items) { $it.Error = 'Excel を起動できません（デスクトップ版の Excel が必要です）' }
        return
    }
    $ownPids = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id } | Where-Object { $pidsBefore -notcontains $_ })
    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $excel.EnableEvents = $false
        try { $excel.AutomationSecurity = 3 } catch { }   # マクロを動かさない（msoAutomationSecurityForceDisable）
        foreach ($it in $Items) {
            Write-Host ("  Excel で取り出しています: {0}" -f $it.Label)
            $wb = $null
            try {
                $wb = $excel.Workbooks.Open($it.Url, 0, $true)
                $wb.SaveCopyAs($it.Dest)
                $it.Error = $null
            }
            catch { $it.Error = $_.Exception.Message }
            finally {
                if ($wb) {
                    try { $wb.Close($false) } catch { }
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
                }
            }
        }
    }
    finally {
        try { $excel.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        # 自分で起動した Excel が残っていれば終わらせる
        foreach ($id in $ownPids) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
    }
}

function Get-DownloadFolder {
    try {
        $p = (New-Object -ComObject Shell.Application).NameSpace('shell:Downloads').Self.Path
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    catch { }
    return Join-Path $env:USERPROFILE 'Downloads'
}

function Save-VersionByBrowser {
    #     既定のブラウザーで版の URL を開いてダウンロードさせ（ブラウザーのサインインで SharePoint を読む）、
    #     ダウンロードのフォルダにできたファイルを移す
    param($Item)

    $downloads = Get-DownloadFolder
    $ext = [System.IO.Path]::GetExtension($Item.Dest)
    $before = @{}
    foreach ($f in (Get-ChildItem -LiteralPath $downloads -File -ErrorAction SilentlyContinue)) { $before[$f.Name] = $f.LastWriteTime }
    $start = Get-Date
    Write-Host ("  ブラウザーでダウンロードします: {0}" -f $Item.Label)
    Write-Host ("  （ダウンロードが始まらないときは、ブラウザーの表示に従って「{0}」に保存してください。{1} 秒待ちます）" -f $downloads, $downloadWait)
    Start-Process $Item.BrowserUrl
    $deadline = $start.AddSeconds($downloadWait)
    $lastSize = -1
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 700
        $new = @(Get-ChildItem -LiteralPath $downloads -File -ErrorAction SilentlyContinue | Where-Object {
                $_.Extension -eq $ext -and (-not $before.ContainsKey($_.Name) -or $_.LastWriteTime -gt $before[$_.Name])
            } | Sort-Object LastWriteTime -Descending)
        if ($new.Count -eq 0) { continue }
        # 書き込みが終わる（大きさが変わらず、開ける）まで待つ
        $f = $new[0]
        if ($f.Length -ne $lastSize) { $lastSize = $f.Length; continue }
        try { $s = [System.IO.File]::Open($f.FullName, 'Open', 'Read', 'None'); $s.Dispose() } catch { continue }
        Move-Item -LiteralPath $f.FullName -Destination $Item.Dest -Force
        $Item.Error = $null
        return
    }
    $Item.Error = "{0} 秒待ちましたが、ダウンロードのフォルダにファイルができませんでした" -f $downloadWait
}

# ---------------------------------------------------------------- 実行


$loc = Resolve-SharePointFile $Path $SiteUrl
if (-not $PSBoundParameters.ContainsKey('Version')) {
    Write-Host ('比べる旧の版の番号を入力してください（SharePoint でファイルの「…」→「バージョン履歴」に出る 3.0 などの番号）')
    $Version = Read-Host '旧の版'
    if (-not $Version) { throw '旧の版が入力されていません' }
}
$oldLabel = ConvertTo-VersionLabel $Version
$newLabel = ConvertTo-VersionLabel $NewVersion
if (-not $oldLabel -and -not $newLabel) { throw '比べる 2 つの版が、どちらも現在の版です。-Version に旧の版の番号を指定してください' }
# 旧が新より後の版なら入れ替える（現在の版はいちばん新しい）
$order = { param($l) if (-not $l) { [double]::MaxValue } else { $x = $l.Split('.'); [int]$x[0] * 512 + [int]$x[1] } }
if ((& $order $oldLabel) -gt (& $order $newLabel)) { $tmp = $oldLabel; $oldLabel = $newLabel; $newLabel = $tmp }

if (-not $OutDir) {
    $docs = [Environment]::GetFolderPath('MyDocuments')
    if (-not $docs) { $docs = [System.IO.Path]::GetTempPath() }
    $OutDir = Join-Path (Join-Path $docs 'Excel ツールの結果') 'SharePoint の版'
}
$OutDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutDir)
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$base = [System.IO.Path]::GetFileNameWithoutExtension($loc.Name)
$ext = [System.IO.Path]::GetExtension($loc.Name)
$sides = foreach ($label in @($oldLabel, $newLabel)) {
    $text = '現在'
    if ($label) { $text = 'v' + $label }
    $url = Get-VersionUrl $loc $label
    $browserUrl = $url
    if (-not $label) { $browserUrl = $url + '?download=1' }   # 現在の版は、ブラウザーで開かずにダウンロードさせる
    [pscustomobject]@{ Label = $text; Version = $label; Url = $url; BrowserUrl = $browserUrl; Dest = (Join-Path $OutDir ('{0}_{1}{2}' -f $base, $text, $ext)); File = $null; Error = $null }
}

Write-Host ('SharePoint: {0}' -f (Get-VersionUrl $loc $null))
$pending = New-Object System.Collections.Generic.List[object]
foreach ($s in $sides) {
    if (-not $s.Version -and $loc.LocalPath) {
        # 現在の版は、同期しているファイルをそのまま使う
        $s.File = $loc.LocalPath
        Write-Host ('  {0}: 同期しているファイルを使います' -f $s.Label)
    }
    elseif ($s.Version -and (Test-Path -LiteralPath $s.Dest)) {
        # 番号の付いた版は変わらないので、前に取り出したものを使う
        $s.File = $s.Dest
        Write-Host ('  {0}: 前に取り出したものを使います: {1}' -f $s.Label, $s.Dest)
    }
    else { $pending.Add($s) }
}

if ($pending.Count -gt 0) {
    if ($Fetcher) {
        foreach ($s in $pending) {
            try { & $Fetcher $s.Url $s.Dest; $s.Error = $null } catch { $s.Error = $_.Exception.Message }
        }
    }
    else {
        if ($Method -ne 'Browser') { Save-VersionByExcel $pending.ToArray() }
        else { foreach ($s in $pending) { $s.Error = 'まだ取り出していません' } }
        if ($Method -ne 'Excel') {
            foreach ($s in $pending) {
                if (-not $s.Error) { continue }
                if ($Method -eq 'Auto') { Write-Host ('  Excel で取り出せませんでした（{0}）。ブラウザーで試します' -f $s.Error) }
                Save-VersionByBrowser $s
            }
        }
    }
    foreach ($s in $pending) {
        if ($s.Error -or -not (Test-Path -LiteralPath $s.Dest)) {
            $why = $s.Error
            if (-not $why) { $why = 'ファイルができませんでした' }
            throw ("{0} を取り出せませんでした: {1}`n  URL: {2}`n  版の番号が「バージョン履歴」にあるか、SharePoint を開けるアカウントで Office・ブラウザーにサインインしているかを確かめてください。" -f $s.Label, $why, $s.Url)
        }
        $s.File = $s.Dest
    }
}

if (-not $OutHtml) { $OutHtml = Join-Path $OutDir ('{0}_{1}-{2}.html' -f $base, $sides[0].Label, $sides[1].Label) }
$compareArgs = @{ Old = $sides[0].File; New = $sides[1].File; OutHtml = $OutHtml }
if ($OutXlsx) { $compareArgs['OutXlsx'] = $OutXlsx }
if ($CompareFormat) { $compareArgs['CompareFormat'] = $true }
if ($CheckHistory) { $compareArgs['CheckHistory'] = $true }
Write-Host ('旧: {0}（{1}）' -f $sides[0].Label, $sides[0].File)
Write-Host ('新: {0}（{1}）' -f $sides[1].Label, $sides[1].File)
& $compareScript @compareArgs
if (-not $NoOpen -and (Test-Path -LiteralPath $OutHtml)) { Invoke-Item -LiteralPath $OutHtml }

# .SYNOPSIS
#     SharePoint（Microsoft 365）のファイルの 2 つの版（バージョン履歴）を取り出して比べ、差分の画面を開きます。
#
# .DESCRIPTION
#     SharePoint の「バージョン履歴」の番号（3.0 など）を指定すると、その版と現在の版（または -NewVersion の版）を取り出し、
#     Compare-Excel.ps1 -OutHtml と同じ、差分をシートの方眼のまま色付けした画面（ブラウザー）を開きます。
#
#       ・ファイルは、SharePoint の URL か、OneDrive で同期しているフォルダのファイル（エクスプローラーのパス）で指定します。
#         同期しているファイルなら、現在の版はそのファイルをそのまま使います
#       ・版は Excel で開いて取り出します（Office にサインインしているアカウントで読むので、アプリの登録やモジュールは要りません）。
#         Excel で開けないときは、既定のブラウザーでダウンロードします（ブラウザーのサインインで読みます）
#       ・取り出した版は -OutDir（既定は「ドキュメント\Excel ツールの結果\SharePoint の版」）に「名前_v3.0.xlsx」の形で残り、
#         同じ番号の版は次から取り出し直しません
#
#     SharePoint の版の URL は「サイト/_vti_history/番号/ライブラリ/ファイル」で、番号は 大きい番号 × 512 + 小さい番号です（3.0 は 1536）。
#
# .PARAMETER Path
#     比べるファイル。SharePoint の URL（Excel の「ファイル」→「情報」→「パスのコピー」でコピーしたもの、またはブラウザーのアドレス）か、
#     OneDrive で同期しているフォルダのファイル。共有リンク（/:x:/s/… など）からはファイルの場所が分からないため使えません。
#
# .PARAMETER Version
#     旧の版の番号（SharePoint でファイルの「…」→「バージョン履歴」に出る番号。3.0・3・3.1 など）。指定しなければ入力を求めます。
#
# .PARAMETER NewVersion
#     新の版の番号。指定しなければ現在の版と比べます。旧より前の版を指定したときは、旧と新を入れ替えます。
#
# .PARAMETER SiteUrl
#     ファイルのあるサイトの URL。既定は URL の /sites/名前・/teams/名前（無ければルートのサイト）です。サブサイトのファイルのときに指定します。
#
# .PARAMETER Method
#     版の取り出し方。Auto（既定。Excel で開けなければブラウザー）・Excel・Browser。
#
# .PARAMETER OutDir
#     取り出した版と差分の HTML を保存するフォルダ。既定は「ドキュメント\Excel ツールの結果\SharePoint の版」。
#
# .PARAMETER OutHtml
#     差分の HTML のパス。既定は -OutDir の「名前_v3.0-現在.html」。
#
# .PARAMETER OutXlsx
#     差分の一覧を Excel にも保存します（Compare-Excel.ps1 の -OutXlsx）。
#
# .PARAMETER CompareFormat
#     書式（太字・文字色・塗りつぶしなど）の変更も比べます（Compare-Excel.ps1 の -CompareFormat）。
#
# .PARAMETER CheckHistory
#     改訂履歴に書かれていない変更を知らせます（Compare-Excel.ps1 の -CheckHistory）。
#
# .PARAMETER NoOpen
#     差分の画面を開きません。
#
# .EXAMPLE
#     # 同期しているフォルダの設計書を、版 3.0 と現在の版で比べる
#     .\Compare-SharePointVersion.ps1 "C:\Users\yamada\会社名\受注システム - ドキュメント\設計\DD03_テーブル定義書.xlsx" 3.0
#
# .EXAMPLE
#     # URL で指定し、版 2.0 と 5.0 を比べる
#     .\Compare-SharePointVersion.ps1 "https://contoso.sharepoint.com/sites/juchu/Shared Documents/設計/DD03_テーブル定義書.xlsx" 2.0 -NewVersion 5.0
#
# .NOTES
#     要件: Windows、PowerShell 5.1 以降。版を Excel で取り出すときはデスクトップ版の Excel が必要です（無ければブラウザーでダウンロードします）。
#     同じフォルダに Compare-Excel.ps1・Search-Excel.ps1・ExcelTools.psm1 と compare-excel フォルダが必要です。
