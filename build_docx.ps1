<#
.SYNOPSIS
  Build a hardware architecture Word document from doc_spec.json.
.PARAMETER SpecPath
  Path to doc_spec.json
.PARAMETER OutPath
  Output .docx path
.PARAMETER AssetDir
  Directory containing figure PNGs referenced by figure.file
.NOTES
  Author: Fasih ud Din Farrukh
  Copyright (c) 2026 Altera Corporation. All rights reserved.
  Author is empty by default. Optional personal override: ../config.local.json
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SpecPath,
    [Parameter(Mandatory = $true)]
    [string]$OutPath,
    [Parameter(Mandatory = $true)]
    [string]$AssetDir
)

$ErrorActionPreference = "Stop"
$SkillRoot = Split-Path -Parent $PSScriptRoot
$wdAlignParagraphCenter = 1
$wdAlignParagraphRight = 2
$wdAlignParagraphLeft = 0
$wdCollapseEnd = 0
$wdHeaderFooterPrimary = 1
$wdFieldPage = 33
$wdFieldNumPages = 26
$wdLineStyleSingle = 1
$wdAutoFitWindow = 2
$wdStory = 6
$wdColorSteel = 0x64381F  # BGR of RGB(31,56,100)
$wdColorWhite = 0xFFFFFF
$wdColorCaption = 0x8B5E3C
$wdPageBreak = 7

function Get-LocalConfig {
    $path = Join-Path $SkillRoot "config.local.json"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $utf8 = New-Object System.Text.UTF8Encoding $false
    return ([System.IO.File]::ReadAllText($path, $utf8) | ConvertFrom-Json)
}

function Get-CopyrightLine($spec) {
    if ($spec.copyright -and [string]$spec.copyright -ne "") {
        return [string]$spec.copyright
    }
    $year = (Get-Date).Year.ToString()
    if ($spec.date) {
        $m = [regex]::Match([string]$spec.date, "(\d{4})")
        if ($m.Success) { $year = $m.Groups[1].Value }
    }
    return "Copyright (c) $year Altera Corporation. All rights reserved."
}

function Get-Authors($spec, $localCfg) {
    $fromSpec = @(Convert-ToStringArray $spec.authors | Where-Object { $_ -and $_.Trim() -ne "" })
    if ($fromSpec.Count -gt 0) { return $fromSpec }
    if ($null -ne $localCfg) {
        $fromLocal = @(Convert-ToStringArray $localCfg.authors | Where-Object { $_ -and $_.Trim() -ne "" })
        if ($fromLocal.Count -gt 0) { return $fromLocal }
    }
    return @()
}

function Get-Abs([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) {
        return [System.IO.Path]::GetFullPath($p)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $p))
}

function Convert-ToStringArray($value) {
    if ($null -eq $value) { return @() }
    if ($value -is [string]) { return @($value) }
    $out = @()
    foreach ($item in @($value)) { $out += [string]$item }
    return $out
}

function Set-RangeFont($range, $name, $size, $bold, $color) {
    $range.Font.Name = $name
    $range.Font.Size = $size
    $range.Font.Bold = [int]$bold
    if ($null -ne $color) { $range.Font.Color = $color }
}

function Add-Para($doc, [string]$text, $style = "Normal") {
    $sel = $doc.Application.Selection
    $end = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $end.Select()
    $sel.Style = $doc.Styles.Item($style)
    if ($style -eq "Normal") {
        $sel.ParagraphFormat.OutlineLevel = 10
    }
    $sel.TypeText([string]$text)
    $para = $sel.Paragraphs.Item(1)
    $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item("Normal")
    $sel.ParagraphFormat.OutlineLevel = 10
    return $para
}

function Add-Bookmark($doc, $para, [string]$id) {
    if ([string]::IsNullOrWhiteSpace($id)) { return }
    $safe = ($id -replace "[^A-Za-z0-9_]", "_")
    if ($doc.Bookmarks.Exists($safe)) {
        $doc.Bookmarks.Item($safe).Delete()
    }
    $bm = $para.Range.Duplicate
    if ($bm.End -gt $bm.Start) {
        $bm.End = $bm.End - 1
    }
    $null = $doc.Bookmarks.Add($safe, $bm)
}

function Add-Bullet($doc, [string]$text) {
    $clean = ([string]$text).Trim()
    $mojibake = ([char]0x00E2).ToString() + ([char]0x20AC).ToString() + ([char]0x00A2).ToString()
    if ($clean.StartsWith($mojibake)) { $clean = $clean.Substring($mojibake.Length).TrimStart() }
    $bullet = ([char]0x2022).ToString()
    if ($clean.StartsWith($bullet)) { $clean = $clean.Substring(1).TrimStart() }
    $sel = $doc.Application.Selection
    $end = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $end.Select()
    $sel.Style = $doc.Styles.Item("Normal")
    $sel.TypeText($clean)
    $sel.Range.ListFormat.ApplyBulletDefault() | Out-Null
    $para = $sel.Paragraphs.Item(1)
    $sel.TypeParagraph()
    $sel.Style = $doc.Styles.Item("Normal")
    try { $sel.Range.ListFormat.RemoveNumbers() | Out-Null } catch { }
    $sel.ParagraphFormat.OutlineLevel = 10
    return $para
}

function Add-TableFromSpec($doc, $headers, $rows) {
    $headerList = Convert-ToStringArray $headers
    $rowObjs = @($rows)
    $nRows = 1 + $rowObjs.Count
    $nCols = [Math]::Max(1, $headerList.Count)
    $p = $doc.Paragraphs.Add()
    $table = $doc.Tables.Add($p.Range, $nRows, $nCols)
    $table.Borders.Enable = $true
    $table.Borders.OutsideLineStyle = $wdLineStyleSingle
    $table.Borders.InsideLineStyle = $wdLineStyleSingle
    $table.AutoFitBehavior($wdAutoFitWindow) | Out-Null
    for ($c = 0; $c -lt $nCols; $c++) {
        $cell = $table.Cell(1, $c + 1)
        $cell.Range.Text = $headerList[$c]
        $cell.Shading.BackgroundPatternColor = $wdColorSteel
        $cell.Range.Font.Color = $wdColorWhite
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Size = 10
        $cell.Range.Font.Name = "Calibri"
        $cell.Range.ParagraphFormat.SpaceAfter = 0
    }
    for ($r = 0; $r -lt $rowObjs.Count; $r++) {
        $cols = Convert-ToStringArray $rowObjs[$r]
        for ($c = 0; $c -lt $nCols; $c++) {
            $val = ""
            if ($c -lt $cols.Count) { $val = $cols[$c] }
            $cell = $table.Cell($r + 2, $c + 1)
            $cell.Range.Text = $val
            $cell.Range.Font.Size = 10
            $cell.Range.Font.Name = "Calibri"
            $cell.Range.Font.Color = 0x000000
            $cell.Range.ParagraphFormat.SpaceAfter = 0
            if ($r % 2 -eq 1) {
                $cell.Shading.BackgroundPatternColor = 0xF3EDE6
            }
        }
    }
    $pEnd = $doc.Paragraphs.Add()
    $pEnd.Range.Text = ""
    $pEnd.Range.InsertParagraphAfter()
}

function Add-Equation($doc, [string]$math, $number, $caption) {
    if ([string]::IsNullOrWhiteSpace($math)) {
        throw "equation block missing math"
    }
    $p = $doc.Paragraphs.Add()
    $table = $doc.Tables.Add($p.Range, 1, 2)
    $table.Borders.Enable = $false
    $table.PreferredWidthType = 2
    $table.PreferredWidth = 100
    $left = $table.Cell(1, 1).Range
    $left.Text = $math.Trim()
    $left.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $null = $left.OMaths.Add($left)
    try {
        $om = $left.OMaths.Item(1)
        $om.BuildUp() | Out-Null
    }
    catch {
        Write-Warning "OMath BuildUp failed for: $math  ($($_.Exception.Message))"
    }
    $right = $table.Cell(1, 2).Range
    $right.ParagraphFormat.Alignment = $wdAlignParagraphRight
    if ($null -ne $number -and "$number" -ne "") {
        $right.Text = "($number)"
    }
    else {
        $right.Text = ""
    }
    $right.Font.Name = "Calibri"
    $right.Font.Size = 11
    if ($caption) {
        $cp = $doc.Paragraphs.Add()
        $cp.Range.Text = [string]$caption
        $cp.Range.Font.Italic = $true
        $cp.Range.Font.Size = 9
        $cp.Range.Font.Color = $wdColorSteel
        $cp.Range.InsertParagraphAfter()
    }
    $sp = $doc.Paragraphs.Add()
    $sp.Range.Text = ""
    $sp.Range.InsertParagraphAfter()
}

function Add-Figure($doc, [string]$assetDir, $block) {
    $file = [string]$block.file
    $full = [System.IO.Path]::GetFullPath((Join-Path $assetDir $file))
    if (-not (Test-Path -LiteralPath $full)) {
        throw "Figure not found: $full"
    }
    $widthIn = 6.5
    if ($block.widthIn) { $widthIn = [double]$block.widthIn }
    $p = $doc.Paragraphs.Add()
    $p.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $pic = $doc.InlineShapes.AddPicture($full, $false, $true, $p.Range)
    $pic.LockAspectRatio = -1
    $pic.Width = $widthIn * 72
    $p.Range.InsertParagraphAfter()
    $cap = "Figure"
    if ($block.caption) { $cap = [string]$block.caption }
    $cp = $doc.Paragraphs.Add()
    $cp.Range.Text = $cap
    $cp.Range.Font.Italic = $true
    $cp.Range.Font.Size = 9
    $cp.Range.Font.Color = $wdColorSteel
    $cp.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $cp.Range.InsertParagraphAfter()
}

$specFile = Get-Abs $SpecPath
$outFile = Get-Abs $OutPath
$assets = Get-Abs $AssetDir
if (-not (Test-Path -LiteralPath $specFile)) { throw "Spec not found: $specFile" }
if (-not (Test-Path -LiteralPath $assets)) { throw "AssetDir not found: $assets" }

$utf8 = New-Object System.Text.UTF8Encoding $false
$spec = [System.IO.File]::ReadAllText($specFile, $utf8) | ConvertFrom-Json
$localCfg = Get-LocalConfig
$copyrightLine = Get-CopyrightLine $spec
$authorList = Get-Authors $spec $localCfg
$defaultRevAuthor = if ($authorList.Count -gt 0) { $authorList[0] } else { "" }
$outDir = Split-Path -Parent $outFile
if (-not (Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}

$word = $null
$doc = $null
$startedWord = $false
try {
    $word = New-Object -ComObject Word.Application
    $startedWord = $true
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()

    $doc.PageSetup.TopMargin = 72
    $doc.PageSetup.BottomMargin = 72
    $doc.PageSetup.LeftMargin = 72
    $doc.PageSetup.RightMargin = 72

    $normal = $doc.Styles.Item("Normal")
    $normal.Font.Name = "Calibri"
    $normal.Font.Size = 11
    $normal.ParagraphFormat.SpaceAfter = 8
    $normal.ParagraphFormat.SpaceBefore = 0
    $normal.Font.Color = 0x2C1A12
    $normal.ParagraphFormat.OutlineLevel = 10

    foreach ($hn in @("Heading 1", "Heading 2", "Heading 3")) {
        $hs = $doc.Styles.Item($hn)
        $hs.Font.Name = "Calibri Light"
        $hs.Font.Bold = $true
        $hs.Font.Color = $wdColorSteel
        $hs.ParagraphFormat.SpaceBefore = 12
        $hs.ParagraphFormat.SpaceAfter = 6
    }
    $doc.Styles.Item("Heading 1").Font.Size = 18
    $doc.Styles.Item("Heading 2").Font.Size = 14
    $doc.Styles.Item("Heading 3").Font.Size = 12

    $sec = $doc.Sections.Item(1)
    $header = $sec.Headers.Item($wdHeaderFooterPrimary).Range
    $header.Text = [string]$spec.title
    $header.Font.Size = 9
    $header.Font.Color = $wdColorSteel
    $header.ParagraphFormat.Alignment = $wdAlignParagraphRight

    $footer = $sec.Footers.Item($wdHeaderFooterPrimary).Range
    $rev = "Rev"
    if ($spec.revision) { $rev = "Rev $($spec.revision)" }
    $footer.Text = "$copyrightLine  |  $rev  |  "
    $footer.Font.Size = 8
    $footer.Font.Color = $wdColorSteel
    $footer.Collapse($wdCollapseEnd)
    $null = $footer.Fields.Add($footer, $wdFieldPage)
    $footer.Collapse($wdCollapseEnd)
    $footer.InsertAfter(" / ")
    $footer.Collapse($wdCollapseEnd)
    $null = $footer.Fields.Add($footer, $wdFieldNumPages)

    # Title page
    $tp = $doc.Paragraphs.Item(1).Range
    $tp.Text = ""
    $titleP = $doc.Paragraphs.Item(1)
    $titleP.Range.Text = [string]$spec.title
    $titleP.Range.Font.Name = "Calibri Light"
    $titleP.Range.Font.Size = 28
    $titleP.Range.Font.Bold = $true
    $titleP.Range.Font.Color = $wdColorSteel
    $titleP.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    $titleP.Range.ParagraphFormat.SpaceBefore = 72
    $titleP.Range.InsertParagraphAfter()

    if ($spec.subtitle) {
        $sp = Add-Para $doc ([string]$spec.subtitle) "Normal"
        $sp.Range.Font.Size = 16
        $sp.Range.Font.Color = $wdColorSteel
        $sp.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
    }

    $meta = @()
    if ($authorList.Count -gt 0) {
        $meta += ("Author: " + ($authorList -join ", "))
    }
    if ($spec.date) { $meta += "Date: $($spec.date)" }
    if ($spec.revision) { $meta += "Revision: $($spec.revision)" }
    if ($spec.classification) { $meta += "Classification: $($spec.classification)" }
    $meta += $copyrightLine
    foreach ($line in $meta) {
        $mp = Add-Para $doc $line "Normal"
        $mp.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
        $mp.Range.Font.Size = 12
    }

    $doc.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    Add-Para $doc "Revision history" "Heading 1" | Out-Null
    $rhHeaders = @("Rev", "Date", "Author", "Description")
    $rhRows = @()
    if ($spec.revisionHistory) {
        foreach ($row in @($spec.revisionHistory)) {
            $rowAuthor = [string]$row.author
            if ([string]::IsNullOrWhiteSpace($rowAuthor)) { $rowAuthor = $defaultRevAuthor }
            $rhRows += , @([string]$row.rev, [string]$row.date, $rowAuthor, [string]$row.description)
        }
    }
    else {
        $rhRows += , @([string]$spec.revision, [string]$spec.date, $defaultRevAuthor, "Initial")
    }
    Add-TableFromSpec $doc $rhHeaders $rhRows

    $contentsP = Add-Para $doc "Contents" "Normal"
    $contentsP.Range.Font.Name = "Calibri Light"
    $contentsP.Range.Font.Size = 18
    $contentsP.Range.Font.Bold = $true
    $contentsP.Range.Font.Color = $wdColorSteel
    $contentsP.Range.ParagraphFormat.OutlineLevel = 10
    $tocPara = Add-Para $doc " " "Normal"
    if ($doc.Bookmarks.Exists("HWD_TOC_PLACEHOLDER")) {
        $doc.Bookmarks.Item("HWD_TOC_PLACEHOLDER").Delete()
    }
    $null = $doc.Bookmarks.Add("HWD_TOC_PLACEHOLDER", $tocPara.Range)
    $doc.Paragraphs.Add().Range.InsertBreak($wdPageBreak)

    foreach ($block in @($spec.sections)) {
        $type = [string]$block.type
        switch ($type) {
            "heading" {
                $level = [int]$block.level
                if ($level -lt 1 -or $level -gt 3) { throw "heading level must be 1-3" }
                $style = "Heading $level"
                $p = Add-Para $doc ([string]$block.text) $style
                if ($level -eq 1) { Add-Bookmark $doc $p ([string]$block.id) }
            }
            "para" {
                Add-Para $doc ([string]$block.text) "Normal" | Out-Null
            }
            "bullets" {
                foreach ($item in (Convert-ToStringArray $block.items)) {
                    Add-Bullet $doc $item | Out-Null
                }
            }
            "table" {
                if ($block.caption) {
                    $cp = Add-Para $doc ([string]$block.caption) "Normal"
                    $cp.Range.Font.Italic = $true
                    $cp.Range.Font.Size = 9
                    $cp.Range.Font.Color = $wdColorSteel
                }
                Add-TableFromSpec $doc $block.headers $block.rows
            }
            "figure" {
                Add-Figure $doc $assets $block
            }
            "equation" {
                Add-Equation $doc ([string]$block.math) $block.number $block.caption
            }
            "note" {
                $np = Add-Para $doc ("Note: " + [string]$block.text) "Normal"
                $np.Range.Font.Italic = $true
                $np.Range.Shading.BackgroundPatternColor = 0xD6EAF8
            }
            default {
                throw "Unknown section type: $type"
            }
        }
    }

    $headingText = @{}
    foreach ($block in @($spec.sections)) {
        if ([string]$block.type -eq "heading") {
            $headingText[[string]$block.text] = [int]$block.level
        }
    }
    $headingText["Revision history"] = 1
    foreach ($p in @($doc.Paragraphs)) {
        $t = ($p.Range.Text -replace "[\r\a]", "").Trim()
        if ($headingText.ContainsKey($t)) {
            $lvl = $headingText[$t]
            $p.Style = $doc.Styles.Item("Heading $lvl")
        }
        else {
            $cur = [string]$p.Style.NameLocal
            if ($cur -match '^Heading [123]$') {
                $p.Style = $doc.Styles.Item("Normal")
            }
            $p.Range.ParagraphFormat.OutlineLevel = 10
        }
    }
    if (-not $doc.Bookmarks.Exists("HWD_TOC_PLACEHOLDER")) {
        throw "Failed to insert contents list (placeholder bookmark missing)"
    }
    $tocBm = $doc.Bookmarks.Item("HWD_TOC_PLACEHOLDER")
    $tocRange = $tocBm.Range
    $tocRange.Text = ""
    $null = $doc.TablesOfContents.Add(
        $tocRange, $true, 1, 3, $false, "", $true, $true, "", $false, $true, $false)
    if ($doc.Bookmarks.Exists("HWD_TOC_PLACEHOLDER")) {
        $doc.Bookmarks.Item("HWD_TOC_PLACEHOLDER").Delete()
    }
    foreach ($toc in $doc.TablesOfContents) {
        $toc.Update() | Out-Null
    }

    $wdFormatXMLDocument = 12
    $savedPath = $outFile
    try {
        if (Test-Path -LiteralPath $outFile) {
            Remove-Item -LiteralPath $outFile -Force -ErrorAction Stop
        }
        $doc.SaveAs([ref]$outFile, [ref]$wdFormatXMLDocument)
    }
    catch {
        $savedPath = [System.IO.Path]::Combine(
            [System.IO.Path]::GetDirectoryName($outFile),
            ([System.IO.Path]::GetFileNameWithoutExtension($outFile) + "_rebuilt.docx"))
        if (Test-Path -LiteralPath $savedPath) {
            Remove-Item -LiteralPath $savedPath -Force
        }
        $doc.SaveAs([ref]$savedPath, [ref]$wdFormatXMLDocument)
        Write-Warning "Original file is locked. Wrote $savedPath instead. Close Word and replace the original."
    }
    Write-Host "Wrote $savedPath"
    Write-Host ("Figures embedded: {0}" -f $doc.InlineShapes.Count)
    Write-Host ("Equations (OMath): {0}" -f $doc.OMaths.Count)
    Write-Host ("TOC entries: {0}" -f $doc.TablesOfContents.Count)
}
finally {
    if ($null -ne $doc) {
        $doc.Close([ref]$false) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($startedWord -and $null -ne $word) {
        $word.Quit() | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
