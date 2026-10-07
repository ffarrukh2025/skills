# Visio-like GDI+ drawing helpers. Dot-source from a figure script. No Python.
# Author: Fasih ud Din Farrukh
# Copyright (c) 2026 Altera Corporation. All rights reserved.
Add-Type -AssemblyName System.Drawing

$script:VsInk = [System.Drawing.Color]::FromArgb(31, 56, 100)
$script:VsBg = [System.Drawing.Color]::FromArgb(247, 249, 252)
$script:VsTitle = [System.Drawing.Color]::FromArgb(31, 56, 100)
$script:VsRole = @{
    in   = @([System.Drawing.Color]::FromArgb(214, 234, 248), [System.Drawing.Color]::FromArgb(26, 82, 118), [System.Drawing.Color]::FromArgb(36, 113, 163))
    comb = @([System.Drawing.Color]::FromArgb(213, 245, 227), [System.Drawing.Color]::FromArgb(25, 111, 61), [System.Drawing.Color]::FromArgb(20, 90, 50))
    reg  = @([System.Drawing.Color]::FromArgb(252, 243, 207), [System.Drawing.Color]::FromArgb(183, 149, 11), [System.Drawing.Color]::FromArgb(146, 119, 9))
    mem  = @([System.Drawing.Color]::FromArgb(250, 219, 216), [System.Drawing.Color]::FromArgb(146, 43, 33), [System.Drawing.Color]::FromArgb(123, 36, 28))
    ctl  = @([System.Drawing.Color]::FromArgb(232, 218, 239), [System.Drawing.Color]::FromArgb(108, 52, 131), [System.Drawing.Color]::FromArgb(91, 44, 111))
    clk  = @([System.Drawing.Color]::FromArgb(245, 203, 167), [System.Drawing.Color]::FromArgb(175, 96, 26), [System.Drawing.Color]::FromArgb(147, 81, 22))
    err  = @([System.Drawing.Color]::FromArgb(245, 183, 177), [System.Drawing.Color]::FromArgb(192, 57, 43), [System.Drawing.Color]::FromArgb(146, 43, 33))
    csr  = @([System.Drawing.Color]::FromArgb(212, 230, 241), [System.Drawing.Color]::FromArgb(40, 116, 166), [System.Drawing.Color]::FromArgb(31, 97, 141))
    ext  = @([System.Drawing.Color]::FromArgb(244, 246, 247), [System.Drawing.Color]::FromArgb(86, 101, 115), [System.Drawing.Color]::FromArgb(52, 73, 94))
}

function Get-VsFont([single]$px, [System.Drawing.FontStyle]$style = [System.Drawing.FontStyle]::Regular) {
    return New-Object System.Drawing.Font "Calibri", $px, $style, ([System.Drawing.GraphicsUnit]::Pixel)
}

function New-VsRoundPath([single]$x, [single]$y, [single]$w, [single]$h, [single]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [Math]::Min([single](2 * $r), [Math]::Min($w, $h))
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc(($x + $w - $d), $y, $d, $d, 270, 90)
    $p.AddArc(($x + $w - $d), ($y + $h - $d), $d, $d, 0, 90)
    $p.AddArc($x, ($y + $h - $d), $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function New-VsCanvas([int]$w, [int]$h) {
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $bmp.SetResolution(200, 200)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear($script:VsBg)
    $script:VsG = $g
    $script:VsBmp = $bmp
    $script:VsW = $w
    $script:VsH = $h
}

function Add-VsTitle([string]$text) {
    $g = $script:VsG
    $bar = New-Object System.Drawing.SolidBrush $script:VsTitle
    $g.FillRectangle($bar, 48, 28, 10, 52)
    $sf = New-Object System.Drawing.StringFormat
    $sf.LineAlignment = "Center"
    $g.DrawString($text, (Get-VsFont 42 ([System.Drawing.FontStyle]::Bold)), $bar,
        (New-Object System.Drawing.RectangleF 70, 20, ($script:VsW - 120), 68), $sf)
}

function Add-VsBanner([single]$x, [single]$y, [single]$w, [single]$h, [string]$text, [string]$role = "clk") {
    $g = $script:VsG
    $fill = $script:VsRole[$role][0]
    $edge = $script:VsRole[$role][1]
    $path = New-VsRoundPath $x $y $w $h 12
    $g.FillPath((New-Object System.Drawing.SolidBrush $fill), $path)
    $g.DrawPath((New-Object System.Drawing.Pen $edge, 2.2), $path)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
    $g.DrawString($text, (Get-VsFont 22 ([System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush $edge),
        (New-Object System.Drawing.RectangleF ($x + 10), $y, ($w - 20), $h), $sf)
}

function Add-VsShadow($g, $path) {
    $m = New-Object System.Drawing.Drawing2D.Matrix
    $m.Translate(5, 6)
    $shadow = [System.Drawing.Drawing2D.GraphicsPath]$path.Clone()
    $shadow.Transform($m)
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(36, 20, 40, 70))), $shadow)
}

function Get-VsPort($box, [string]$side) {
    switch ($side) {
        "left" { return @{ x = $box.x; y = $box.y + $box.h / 2 } }
        "right" { return @{ x = $box.x + $box.w; y = $box.y + $box.h / 2 } }
        "top" { return @{ x = $box.x + $box.w / 2; y = $box.y } }
        "bottom" { return @{ x = $box.x + $box.w / 2; y = $box.y + $box.h } }
        default { return @{ x = $box.x + $box.w; y = $box.y + $box.h / 2 } }
    }
}

function Add-VsPortDot($g, $pt) {
    $g.FillEllipse((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)), ($pt.x - 4.5), ($pt.y - 4.5), 9, 9)
    $g.DrawEllipse((New-Object System.Drawing.Pen $script:VsInk, 1.6), ($pt.x - 4.5), ($pt.y - 4.5), 9, 9)
}

function Add-VsShape {
    param(
        [single]$x, [single]$y, [single]$w, [single]$h,
        [string]$header, [string]$body = "", [string]$role = "comb",
        [ValidateSet("process", "io", "store", "register", "decision", "terminator", "note")]
        [string]$kind = "process",
        [bool]$doubleBorder = $false,
        [string]$ports = "auto"
    )
    $g = $script:VsG
    $fill = $script:VsRole[$role][0]
    $edge = $script:VsRole[$role][1]
    $head = $script:VsRole[$role][2]
    $hdrH = if ($kind -eq "decision") { 0 } else { [Math]::Min(52, [single]($h * 0.34)) }
    if ($kind -eq "note") { $hdrH = 0 }

    if ($kind -eq "store") {
        $elH = 26
        $bodyTop = $y + $elH / 2
        $bodyH = $h - $elH
        $bodyPath = New-VsRoundPath $x $bodyTop $w ($bodyH - 4) 8
        Add-VsShadow $g $bodyPath
        $g.FillRectangle((New-Object System.Drawing.SolidBrush $fill), $x, $bodyTop, $w, ($bodyH - $elH / 2))
        $g.FillEllipse((New-Object System.Drawing.SolidBrush $fill), $x, ($y + $h - $elH), $w, $elH)
        $g.FillEllipse((New-Object System.Drawing.SolidBrush $fill), $x, $y, $w, $elH)
        $pen = New-Object System.Drawing.Pen $edge, 2.6
        $g.DrawLine($pen, $x, ($y + $elH / 2), $x, ($y + $h - $elH / 2))
        $g.DrawLine($pen, ($x + $w), ($y + $elH / 2), ($x + $w), ($y + $h - $elH / 2))
        $g.DrawEllipse($pen, $x, ($y + $h - $elH), $w, $elH)
        $g.FillEllipse((New-Object System.Drawing.SolidBrush $head), $x, $y, $w, $elH)
        $g.DrawEllipse($pen, $x, $y, $w, $elH)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
        $g.DrawString($header, (Get-VsFont 22 ([System.Drawing.FontStyle]::Bold)),
            (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)),
            (New-Object System.Drawing.RectangleF $x, ($y + 2), $w, ($elH - 2)), $sf)
        $g.DrawString($body, (Get-VsFont 22), (New-Object System.Drawing.SolidBrush $edge),
            (New-Object System.Drawing.RectangleF ($x + 14), ($y + $elH + 8), ($w - 28), ($h - $elH - 22)), $sf)
    }
    elseif ($kind -eq "decision") {
        $path = New-Object System.Drawing.Drawing2D.GraphicsPath
        $cx = $x + $w / 2; $cy = $y + $h / 2
        $pts = @(
            (New-Object System.Drawing.PointF $cx, $y),
            (New-Object System.Drawing.PointF ($x + $w), $cy),
            (New-Object System.Drawing.PointF $cx, ($y + $h)),
            (New-Object System.Drawing.PointF $x, $cy)
        )
        $path.AddPolygon($pts)
        Add-VsShadow $g $path
        $g.FillPath((New-Object System.Drawing.SolidBrush $fill), $path)
        $g.DrawPath((New-Object System.Drawing.Pen $edge, 2.8), $path)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
        $g.DrawString("$header`n$body", (Get-VsFont 22 ([System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush $edge),
            (New-Object System.Drawing.RectangleF ($x + $w * 0.18), ($y + $h * 0.22), ($w * 0.64), ($h * 0.56)), $sf)
    }
    else {
        $rad = if ($kind -eq "terminator" -or $kind -eq "register") { [single]($h / 2) } else { [single]14 }
        $path = New-VsRoundPath $x $y $w $h $rad
        Add-VsShadow $g $path
        $g.FillPath((New-Object System.Drawing.SolidBrush $fill), $path)
        if ($hdrH -gt 0 -and $kind -ne "terminator") {
            $clip = $g.Save()
            $g.SetClip($path)
            $g.FillRectangle((New-Object System.Drawing.SolidBrush $head), $x, $y, $w, $hdrH)
            $g.Restore($clip)
            $g.DrawLine((New-Object System.Drawing.Pen $edge, 1.4), ($x + 1), ($y + $hdrH), ($x + $w - 1), ($y + $hdrH))
        }
        $penW = if ($doubleBorder) { 5.0 } else { 2.6 }
        $g.DrawPath((New-Object System.Drawing.Pen $edge, $penW), $path)
        if ($doubleBorder) {
            $inner = New-VsRoundPath ($x + 7) ($y + 7) ($w - 14) ($h - 14) ([Math]::Max(6, $rad - 6))
            $g.DrawPath((New-Object System.Drawing.Pen $edge, 2.0), $inner)
        }
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
        if ($hdrH -gt 0 -and $kind -ne "terminator") {
            $g.DrawString($header, (Get-VsFont 24 ([System.Drawing.FontStyle]::Bold)),
                (New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)),
                (New-Object System.Drawing.RectangleF ($x + 8), ($y + 4), ($w - 16), ($hdrH - 6)), $sf)
            $g.DrawString($body, (Get-VsFont 22), (New-Object System.Drawing.SolidBrush $edge),
                (New-Object System.Drawing.RectangleF ($x + 14), ($y + $hdrH + 8), ($w - 28), ($h - $hdrH - 18)), $sf)
        }
        else {
            $g.DrawString("$header`n$body", (Get-VsFont 22 ([System.Drawing.FontStyle]::Bold)),
                (New-Object System.Drawing.SolidBrush $edge),
                (New-Object System.Drawing.RectangleF ($x + 12), ($y + 10), ($w - 24), ($h - 20)), $sf)
        }
    }

    $box = @{ x = $x; y = $y; w = $w; h = $h; role = $role; kind = $kind }
    if ($ports -eq "auto") {
        $ports = switch ($kind) {
            "note" { "none" }
            "decision" { "ltrb" }
            "store" { "lr" }
            default { "lr" }
        }
    }
    if ($ports -ne "none") {
        if ($ports -match "l") { Add-VsPortDot $g (Get-VsPort $box "left") }
        if ($ports -match "r") { Add-VsPortDot $g (Get-VsPort $box "right") }
        if ($ports -match "t") { Add-VsPortDot $g (Get-VsPort $box "top") }
        if ($ports -match "b") { Add-VsPortDot $g (Get-VsPort $box "bottom") }
    }
    return $box
}

function Add-VsContainer([single]$x, [single]$y, [single]$w, [single]$h, [string]$title) {
    $g = $script:VsG
    $path = New-VsRoundPath $x $y $w $h 16
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(255, 255, 255))), $path)
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(160, 31, 56, 100)), 2.0
    $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash
    $g.DrawPath($pen, $path)
    $g.FillRectangle((New-Object System.Drawing.SolidBrush $script:VsTitle), ($x + 18), ($y + 14), 8, 28)
    $g.DrawString($title, (Get-VsFont 22 ([System.Drawing.FontStyle]::Bold)), (New-Object System.Drawing.SolidBrush $script:VsTitle), ($x + 34), ($y + 12))
}

function Add-VsPt($list, $x, $y) {
    [void]$list.Add((New-Object System.Drawing.PointF $x, $y))
}

function Add-VsConnector {
    param($from, [string]$fromSide, $to, [string]$toSide, [string]$label = "", [switch]$back)
    $g = $script:VsG
    $a = Get-VsPort $from $fromSide
    $b = Get-VsPort $to $toSide
    $stub = 36
    $pts = New-Object System.Collections.Generic.List[System.Drawing.PointF]
    [void]$pts.Add((New-Object System.Drawing.PointF $a.x, $a.y))

    switch ("$fromSide->$toSide") {
        "right->left" {
            $mx = ($a.x + $b.x) / 2
            if ([Math]::Abs($a.y - $b.y) -lt 2) { Add-VsPt $pts $b.x $b.y }
            else { Add-VsPt $pts $mx $a.y; Add-VsPt $pts $mx $b.y; Add-VsPt $pts $b.x $b.y }
        }
        "left->right" {
            $mx = ($a.x + $b.x) / 2
            if ([Math]::Abs($a.y - $b.y) -lt 2) { Add-VsPt $pts $b.x $b.y }
            else { Add-VsPt $pts $mx $a.y; Add-VsPt $pts $mx $b.y; Add-VsPt $pts $b.x $b.y }
        }
        "bottom->bottom" {
            $yy = [Math]::Max($a.y, $b.y) + $stub
            Add-VsPt $pts $a.x $yy
            Add-VsPt $pts $b.x $yy
            Add-VsPt $pts $b.x $b.y
        }
        "top->top" {
            $yy = [Math]::Min($a.y, $b.y) - $stub
            Add-VsPt $pts $a.x $yy
            Add-VsPt $pts $b.x $yy
            Add-VsPt $pts $b.x $b.y
        }
        "bottom->top" {
            $my = ($a.y + $b.y) / 2
            if ([Math]::Abs($a.x - $b.x) -lt 2) { Add-VsPt $pts $b.x $b.y }
            else { Add-VsPt $pts $a.x $my; Add-VsPt $pts $b.x $my; Add-VsPt $pts $b.x $b.y }
        }
        "top->bottom" {
            $my = ($a.y + $b.y) / 2
            if ([Math]::Abs($a.x - $b.x) -lt 2) { Add-VsPt $pts $b.x $b.y }
            else { Add-VsPt $pts $a.x $my; Add-VsPt $pts $b.x $my; Add-VsPt $pts $b.x $b.y }
        }
        "bottom->left" { Add-VsPt $pts $a.x ($a.y + $stub); Add-VsPt $pts $b.x ($a.y + $stub); Add-VsPt $pts $b.x $b.y }
        "right->top" { Add-VsPt $pts ($a.x + $stub) $a.y; Add-VsPt $pts ($a.x + $stub) $b.y; Add-VsPt $pts $b.x $b.y }
        "right->bottom" { Add-VsPt $pts ($a.x + $stub) $a.y; Add-VsPt $pts ($a.x + $stub) $b.y; Add-VsPt $pts $b.x $b.y }
        "left->bottom" { Add-VsPt $pts ($a.x - $stub) $a.y; Add-VsPt $pts ($a.x - $stub) $b.y; Add-VsPt $pts $b.x $b.y }
        "left->top" { Add-VsPt $pts ($a.x - $stub) $a.y; Add-VsPt $pts ($a.x - $stub) $b.y; Add-VsPt $pts $b.x $b.y }
        "bottom->right" { Add-VsPt $pts $a.x ($a.y + $stub); Add-VsPt $pts $b.x ($a.y + $stub); Add-VsPt $pts $b.x $b.y }
        "top->left" { Add-VsPt $pts $a.x ($a.y - $stub); Add-VsPt $pts $b.x ($a.y - $stub); Add-VsPt $pts $b.x $b.y }
        "top->right" { Add-VsPt $pts $a.x ($a.y - $stub); Add-VsPt $pts $b.x ($a.y - $stub); Add-VsPt $pts $b.x $b.y }
        default {
            Add-VsPt $pts (($a.x + $b.x) / 2) $a.y
            Add-VsPt $pts (($a.x + $b.x) / 2) $b.y
            Add-VsPt $pts $b.x $b.y
        }
    }

    $pen = New-Object System.Drawing.Pen $script:VsInk, 2.8
    $cap = New-Object System.Drawing.Drawing2D.AdjustableArrowCap 4.4, 5.2, $true
    $pen.CustomEndCap = $cap
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    if ($back) { $pen.DashStyle = [System.Drawing.Drawing2D.DashStyle]::Dash }
    $g.DrawLines($pen, $pts.ToArray())

    if ($label) {
        $bestI = 0
        $bestLen = -1.0
        for ($i = 0; $i -lt ($pts.Count - 1); $i++) {
            $dx = $pts[$i + 1].X - $pts[$i].X
            $dy = $pts[$i + 1].Y - $pts[$i].Y
            $len = [Math]::Sqrt($dx * $dx + $dy * $dy)
            if ($len -gt $bestLen) { $bestLen = $len; $bestI = $i }
        }
        $lx = ($pts[$bestI].X + $pts[$bestI + 1].X) / 2
        $ly = ($pts[$bestI].Y + $pts[$bestI + 1].Y) / 2
        $font = Get-VsFont 18 ([System.Drawing.FontStyle]::Bold)
        $sz = $g.MeasureString($label, $font)
        $rw = $sz.Width + 18; $rh = $sz.Height + 4
        $rx = $lx - $rw / 2
        $ry = $ly - $rh - 10
        if ($ry -lt 8) { $ry = $ly + 10 }
        $labPath = New-VsRoundPath $rx $ry $rw $rh 8
        $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)), $labPath)
        $g.DrawPath((New-Object System.Drawing.Pen $script:VsInk, 1.3), $labPath)
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
        $g.DrawString($label, $font, (New-Object System.Drawing.SolidBrush $script:VsInk),
            (New-Object System.Drawing.RectangleF $rx, $ry, $rw, $rh), $sf)
    }
}

function Add-VsCaption([string]$text, [single]$x, [single]$y, [single]$w, [single]$h = 40, [bool]$bold = $false) {
    $style = if ($bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
    $script:VsG.DrawString($text, (Get-VsFont 22 $style), (New-Object System.Drawing.SolidBrush $script:VsInk),
        (New-Object System.Drawing.RectangleF $x, $y, $w, $h), $sf)
}

function Add-VsLegend([single]$x, [single]$y, $items) {
    $g = $script:VsG
    $n = @($items).Count
    $w = [Math]::Min(2080, 40 + $n * 310)
    $path = New-VsRoundPath $x $y $w 70 12
    $g.FillPath((New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)), $path)
    $g.DrawPath((New-Object System.Drawing.Pen $script:VsInk, 1.6), $path)
    $i = 0
    foreach ($it in $items) {
        $xx = $x + 24 + ($i * 300)
        $fill = $script:VsRole[$it[0]][0]
        $edge = $script:VsRole[$it[0]][1]
        $sw = New-VsRoundPath $xx ($y + 20) 36 28 7
        $g.FillPath((New-Object System.Drawing.SolidBrush $fill), $sw)
        $g.DrawPath((New-Object System.Drawing.Pen $edge, 2), $sw)
        $g.DrawString($it[1], (Get-VsFont 22), (New-Object System.Drawing.SolidBrush $script:VsInk), ($xx + 46), ($y + 22))
        $i++
    }
}

function Add-VsBitBar([single]$x, [single]$y, [single]$h, $fields) {
    # $fields = @( @(widthPx, "label", "role"), ... )
    $g = $script:VsG
    $xx = $x
    $pathOuter = New-VsRoundPath $x $y (($fields | ForEach-Object { $_[0] } | Measure-Object -Sum).Sum) $h 10
    Add-VsShadow $g $pathOuter
    $g.SetClip($pathOuter)
    foreach ($f in $fields) {
        $fw = [single]$f[0]; $lab = [string]$f[1]; $role = [string]$f[2]
        $g.FillRectangle((New-Object System.Drawing.SolidBrush $script:VsRole[$role][0]), $xx, $y, $fw, $h)
        $g.DrawLine((New-Object System.Drawing.Pen $script:VsRole[$role][1], 1.4), $xx, $y, $xx, ($y + $h))
        $sf = New-Object System.Drawing.StringFormat
        $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
        $g.DrawString($lab, (Get-VsFont 20 ([System.Drawing.FontStyle]::Bold)),
            (New-Object System.Drawing.SolidBrush $script:VsRole[$role][1]),
            (New-Object System.Drawing.RectangleF $xx, $y, $fw, $h), $sf)
        $xx += $fw
    }
    $g.ResetClip()
    $g.DrawPath((New-Object System.Drawing.Pen $script:VsInk, 2.4), $pathOuter)
}

function Save-VsCanvas([string]$path) {
    $dir = Split-Path -Parent $path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }
    $script:VsBmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    $script:VsG.Dispose()
    $script:VsBmp.Dispose()
    Write-Host "Wrote $path"
}
