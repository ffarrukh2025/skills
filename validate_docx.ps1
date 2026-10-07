<#
.SYNOPSIS
  Open a hardware architecture .docx in Word and check figure / equation counts.
.PARAMETER DocPath
  Path to the .docx
.PARAMETER ExpectedFigures
  Required InlineShape (picture) count
.PARAMETER ExpectedEquations
  Required OMath count
.PARAMETER SpecPath
  Optional doc_spec.json; if set, expected counts are derived from figure/equation blocks unless overridden
.NOTES
  Author: Fasih ud Din Farrukh
  Copyright (c) 2026 Altera Corporation. All rights reserved.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DocPath,
    [int]$ExpectedFigures = -1,
    [int]$ExpectedEquations = -1,
    [string]$SpecPath = ""
)

$ErrorActionPreference = "Stop"

function Get-Abs([string]$p) {
    if ([System.IO.Path]::IsPathRooted($p)) {
        return [System.IO.Path]::GetFullPath($p)
    }
    return [System.IO.Path]::GetFullPath((Join-Path (Get-Location).Path $p))
}

$docFile = Get-Abs $DocPath
if (-not (Test-Path -LiteralPath $docFile)) { throw "Document not found: $docFile" }

if ($SpecPath) {
    $specFile = Get-Abs $SpecPath
    $spec = [System.IO.File]::ReadAllText($specFile, (New-Object System.Text.UTF8Encoding $false)) | ConvertFrom-Json
    $figFromSpec = 0
    $eqFromSpec = 0
    foreach ($block in @($spec.sections)) {
        if ([string]$block.type -eq "figure") { $figFromSpec++ }
        if ([string]$block.type -eq "equation") { $eqFromSpec++ }
    }
    if ($ExpectedFigures -lt 0) { $ExpectedFigures = $figFromSpec }
    if ($ExpectedEquations -lt 0) { $ExpectedEquations = $eqFromSpec }
}

$word = $null
$doc = $null
$startedWord = $false
$failures = @()
try {
    $word = New-Object -ComObject Word.Application
    $startedWord = $true
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($docFile, $false, $true)

    $figCount = 0
    $wdInlineShapePicture = 3
    $wdInlineShapeLinkedPicture = 4
    foreach ($shape in @($doc.InlineShapes)) {
        if ($shape.Type -eq $wdInlineShapePicture -or $shape.Type -eq $wdInlineShapeLinkedPicture) {
            $figCount++
        }
    }

    $eqCount = [int]$doc.OMaths.Count
    $tblCount = [int]$doc.Tables.Count
    $tocCount = [int]$doc.TablesOfContents.Count
    $wdFieldTOC = 37
    foreach ($field in @($doc.Fields)) {
        if ($field.Type -eq $wdFieldTOC) { $tocCount++ }
    }
    $paraCount = [int]$doc.Paragraphs.Count
    $words = [int]$doc.ComputeStatistics(0)  # wdStatisticWords = 0

    Write-Host "Opened: $docFile"
    Write-Host "Words: $words"
    Write-Host "Paragraphs: $paraCount"
    Write-Host "Tables: $tblCount"
    Write-Host "TOC fields: $tocCount"
    Write-Host "Pictures (InlineShapes): $figCount"
    Write-Host "Equations (OMath): $eqCount"

    if ($tocCount -lt 1) { $failures += "Missing contents list (TOC field)" }
    if ($tblCount -lt 1) { $failures += "No tables found; interface/parameter tables are required" }
    if ($words -lt 80) { $failures += "Body text is too short to be an architecture document" }
    if ($ExpectedFigures -ge 0 -and $figCount -ne $ExpectedFigures) {
        $failures += "Figure count $figCount != expected $ExpectedFigures"
    }
    if ($ExpectedEquations -ge 0 -and $eqCount -ne $ExpectedEquations) {
        $failures += "Equation count $eqCount != expected $ExpectedEquations"
    }
    if ($ExpectedFigures -lt 0 -and $figCount -lt 1) {
        $failures += "No embedded diagrams"
    }
    if ($ExpectedEquations -lt 0 -and $eqCount -lt 1) {
        Write-Warning "No OMath equations found. Native editable equations are required when the module has math."
    }

    $body = $doc.Content.Text
    if ($body -like "*Error! Bookmark not defined*") {
        $failures += "Contents list has broken bookmarks (Error! Bookmark not defined)"
    }
    $mojibake = ([char]0x00E2).ToString() + ([char]0x20AC).ToString() + ([char]0x00A2).ToString()
    if ($body.Contains($mojibake)) {
        $failures += "Mojibake bullet prefix found (UTF-8 bullet read as ANSI)"
    }
    if ($tocCount -gt 0) {
        $tocText = $doc.TablesOfContents.Item(1).Range.Text
        if ([string]::IsNullOrWhiteSpace($tocText) -or ($tocText.Trim().Length -lt 8)) {
            $failures += "Contents list is empty"
        }
    }
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

if ($failures.Count -gt 0) {
    Write-Host "VALIDATION FAILED"
    foreach ($f in $failures) { Write-Host " - $f" }
    exit 1
}

Write-Host "VALIDATION PASSED"
exit 0
