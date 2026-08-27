param(
    [Parameter(Mandatory = $true)]
    [string]$OutFile
)

Add-Type -AssemblyName System.Drawing

$size = 256
$bmp = New-Object System.Drawing.Bitmap $size, $size
$g = [System.Drawing.Graphics]::FromImage($bmp)
$g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
$g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
$g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
$g.Clear([System.Drawing.Color]::White)

$navy = [System.Drawing.Color]::FromArgb(26, 54, 93)
$brush = New-Object System.Drawing.SolidBrush $navy
$font = New-Object System.Drawing.Font "Segoe UI Semibold", 42, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel

$text = "myslice"
$textSize = $g.MeasureString($text, $font)
$x = [single](($size - $textSize.Width) / 2)
$y = [single](($size - $textSize.Height) / 2 + 4)
$g.DrawString($text, $font, $brush, $x, $y)

# House mark inside the letter "c" (right side of the wordmark)
$houseX = [single]($x + $textSize.Width * 0.78)
$houseY = [single]($y + $textSize.Height * 0.22)
$houseW = [single]($textSize.Width * 0.11)
$houseH = [single]($textSize.Height * 0.42)
$roof = [System.Drawing.PointF[]]@(
    [System.Drawing.PointF]::new($houseX + $houseW / 2, $houseY),
    [System.Drawing.PointF]::new($houseX, $houseY + $houseH * 0.38),
    [System.Drawing.PointF]::new($houseX + $houseW, $houseY + $houseH * 0.38)
)
$g.FillPolygon($brush, $roof)
$body = New-Object System.Drawing.RectangleF ($houseX + $houseW * 0.12), ($houseY + $houseH * 0.32), ($houseW * 0.76), ($houseH * 0.55)
$g.FillRectangle($brush, $body)
$door = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
$doorRect = New-Object System.Drawing.RectangleF ($houseX + $houseW * 0.38), ($houseY + $houseH * 0.52), ($houseW * 0.24), ($houseH * 0.35)
$g.FillRectangle($door, $doorRect)

$dir = Split-Path -Parent $OutFile
if (-not (Test-Path $dir)) {
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
}
$bmp.Save($OutFile, [System.Drawing.Imaging.ImageFormat]::Png)

$g.Dispose()
$bmp.Dispose()
$brush.Dispose()
$door.Dispose()
$font.Dispose()
