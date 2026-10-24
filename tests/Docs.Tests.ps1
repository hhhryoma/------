#     ドキュメント（README.md・tools\README.md・docs などの Markdown）のテスト。
#     リンク先のファイルと見出しがあること、ツールが README に載っていること、
#     ツールの引数がヘルプ（.PARAMETER）と tools\README.md のツールの節に書いてあることを確かめる

# README に載せない、ほかのツールから呼ばれるだけのスクリプト
$InternalTools = @('Invoke-SendTo.ps1')

function Read-Doc {
    param([string] $Path)
    return [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
}

function Get-DocFiles {
    #     リポジトリの Markdown（外から入れたモジュールと、テストのフィクスチャは除く）
    return @(Get-ChildItem -LiteralPath $script:RepoRoot -Filter '*.md' -Recurse -File | Where-Object {
            $rel = ConvertTo-Slash $_.FullName.Substring($script:RepoRoot.Length)
            $rel -notmatch '^/?(xvba_modules|node_modules|\.git|tests/fixtures)/'
        })
}

function ConvertTo-HeadingAnchor {
    #     GitHub と同じ見出しのアンカー。小文字にし、文字・数字・「-」「_」のほかの記号（「（」「・」「.」など）を除き、空白を「-」にする
    param([string] $Heading)
    $text = ($Heading -replace '\[([^\]]*)\]\([^)]*\)', '$1' -replace '<[^>]+>', '').Replace('`', '').Trim().ToLowerInvariant()
    $keep = @('NonSpacingMark', 'SpacingCombiningMark', 'EnclosingMark', 'LetterNumber', 'OtherNumber')
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $text.ToCharArray()) {
        if ($ch -eq ' ') { [void]$sb.Append('-') }
        elseif ($ch -eq '-' -or $ch -eq '_' -or [char]::IsLetterOrDigit($ch)) { [void]$sb.Append($ch) }
        elseif ($keep -contains [string][System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)) { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Get-DocLines {
    #     コードブロックの外の行（行番号付き）
    param([string] $Text)
    $result = New-Object System.Collections.Generic.List[object]
    $inCode = $false
    $no = 0
    foreach ($line in ($Text -split "\r?\n")) {
        $no++
        if ($line -match '^\s*```') { $inCode = -not $inCode; continue }
        if (-not $inCode) { $result.Add([pscustomobject]@{ No = $no; Text = $line }) }
    }
    return , $result
}

function Get-HeadingAnchors {
    #     見出しのアンカーの一覧（同じ見出しが続くときは GitHub と同じく -1・-2 を付ける）
    param([string] $Text)
    $anchors = @{}
    foreach ($line in (Get-DocLines $Text)) {
        $m = [regex]::Match($line.Text, '^#{1,6}\s+(.+?)\s*#*\s*$')
        if (-not $m.Success) { continue }
        $anchor = ConvertTo-HeadingAnchor $m.Groups[1].Value
        $key = $anchor
        $n = 0
        while ($anchors.ContainsKey($key)) { $n++; $key = '{0}-{1}' -f $anchor, $n }
        $anchors[$key] = $true
    }
    return $anchors
}

Test-Case 'Docs: Markdown のリンク先のファイルと見出しがある' {
    $anchorCache = @{}
    $broken = New-Object System.Collections.Generic.List[string]
    $count = 0
    foreach ($doc in (Get-DocFiles)) {
        $rel = ConvertTo-Slash $doc.FullName.Substring($script:RepoRoot.Length).TrimStart('\', '/')
        $text = Read-Doc $doc.FullName
        foreach ($line in (Get-DocLines $text)) {
            # `…` の中はリンクではない
            $plain = $line.Text -replace '`[^`]*`', ''
            foreach ($m in [regex]::Matches($plain, '\]\(([^)\s]+)(?:\s+"[^"]*")?\)')) {
                $link = $m.Groups[1].Value
                if ($link -match '^[A-Za-z][A-Za-z0-9+.\-]*:') { continue }
                $count++
                $path = $link; $fragment = ''
                $hash = $link.IndexOf('#')
                if ($hash -ge 0) { $path = $link.Substring(0, $hash); $fragment = [uri]::UnescapeDataString($link.Substring($hash + 1)) }
                $target = $doc.FullName
                if ($path) {
                    $target = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($doc.DirectoryName, [uri]::UnescapeDataString($path)))
                    if (-not (Test-Path -LiteralPath $target)) { $broken.Add(('{0}:{1}: {2}（ファイルが無い）' -f $rel, $line.No, $link)); continue }
                }
                if (-not $fragment -or $target -notmatch '\.md$') { continue }
                if (-not $anchorCache.ContainsKey($target)) { $anchorCache[$target] = Get-HeadingAnchors (Read-Doc $target) }
                if (-not $anchorCache[$target].ContainsKey($fragment.ToLowerInvariant())) { $broken.Add(('{0}:{1}: {2}（見出しが無い）' -f $rel, $line.No, $link)) }
            }
        }
    }
    Assert-True ($count -gt 100) "リンクを読めていない（$count 件）"
    Assert-Equal @() @($broken) 'リンク切れ'
}

Test-Case 'Docs: 見出しのアンカーは GitHub と同じ形にする' {
    Assert-Equal 'search-excelps1' (ConvertTo-HeadingAnchor 'Search-Excel.ps1')
    Assert-Equal 'start-exceltoolsps1画面' (ConvertTo-HeadingAnchor 'Start-ExcelTools.ps1（画面）')
    Assert-Equal '差分をブラウザーで見る-outhtml' (ConvertTo-HeadingAnchor '差分をブラウザーで見る（`-OutHtml`）')
    Assert-Equal 'ツールで確かめる観点確かめ方ツールの確認' (ConvertTo-HeadingAnchor 'ツールで確かめる観点（確かめ方・ツールの確認）')
    Assert-Equal '1-概要と-api' (ConvertTo-HeadingAnchor '1. 概要と [API](x.md)')
    $anchors = Get-HeadingAnchors (@('# A', '## 使い方', '```', '## 使い方（コードの中）', '```', '## 使い方') -join "`n")
    Assert-Equal @('a', '使い方', '使い方-1') @($anchors.Keys | Sort-Object) 'コードの中の # は見出しではない。同じ見出しには -1 を付ける'
}

Test-Case 'Docs: すべてのツールが README.md と tools\README.md に載っている' {
    $readme = Read-Doc (Join-Path $script:RepoRoot 'README.md')
    $toolsReadme = Read-Doc (Join-Path $script:ToolsDir 'README.md')
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($tool in (Get-ChildItem -LiteralPath $script:ToolsDir -Filter '*.ps1' -File)) {
        if ($InternalTools -contains $tool.Name) { continue }
        if (-not $readme.Contains("[$($tool.Name)](tools/README.md#")) { $missing.Add("README.md の一覧: $($tool.Name)") }
        if (-not $toolsReadme.Contains("[$($tool.Name)](#")) { $missing.Add("tools\README.md の一覧: $($tool.Name)") }
        if ($toolsReadme -notmatch ('(?m)^## ' + [regex]::Escape($tool.Name))) { $missing.Add("tools\README.md の節（## $($tool.Name)）: $($tool.Name)") }
    }
    Assert-Equal @() @($missing) 'README に無いツール'
}

Test-Case 'Docs: ツールの引数は、ヘルプ（.PARAMETER）と tools\README.md のツールの節に書いてある' {
    # tools\README.md を「## ツール名.ps1」の節に分ける
    $sections = @{}
    $current = $null
    foreach ($line in (Get-DocLines (Read-Doc (Join-Path $script:ToolsDir 'README.md')))) {
        $m = [regex]::Match($line.Text, '^## (\S+\.ps1)')
        if ($m.Success) { $current = $m.Groups[1].Value; $sections[$current] = New-Object System.Text.StringBuilder; continue }
        if ($line.Text -match '^## ') { $current = $null; continue }
        if ($current) { [void]$sections[$current].AppendLine($line.Text) }
    }
    $missing = New-Object System.Collections.Generic.List[string]
    $checked = 0
    foreach ($tool in (Get-ChildItem -LiteralPath $script:ToolsDir -Filter '*.ps1' -File)) {
        if ($InternalTools -contains $tool.Name) { continue }
        $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($tool.FullName, [ref]$null, [ref]$errors)
        if (-not $ast.ParamBlock) { continue }
        $source = Read-Doc $tool.FullName
        foreach ($p in $ast.ParamBlock.Parameters) {
            $name = $p.Name.VariablePath.UserPath
            $checked++
            # [Parameter(DontShow)] の引数（テスト用）は、ヘルプにも README にも書かない
            $hidden = @($p.Attributes | Where-Object { $_.TypeName.Name -eq 'Parameter' } | ForEach-Object { $_.NamedArguments } | Where-Object { $_.ArgumentName -eq 'DontShow' -and ($_.ExpressionOmitted -or $_.Argument.Extent.Text -eq '$true') })
            if ($hidden.Count -gt 0) { continue }
            if ($source -notmatch ('(?mi)^\s*#\s*\.PARAMETER\s+' + [regex]::Escape($name) + '\s*$')) { $missing.Add(('{0} -{1}: ヘルプの .PARAMETER が無い' -f $tool.Name, $name)); continue }
            $section = ''
            if ($sections.ContainsKey($tool.Name)) { $section = $sections[$tool.Name].ToString() }
            if ($section -notmatch ('(?i)`-' + [regex]::Escape($name) + '(?![A-Za-z0-9])')) { $missing.Add(('{0} -{1}: tools\README.md の節に無い' -f $tool.Name, $name)) }
        }
    }
    Assert-True ($checked -gt 300) "引数を読めていない（$checked 件）"
    Assert-Equal @() @($missing) '説明の無い引数（README の「主なパラメーター」の表などに `-名前` の形で書いてください）'
}
