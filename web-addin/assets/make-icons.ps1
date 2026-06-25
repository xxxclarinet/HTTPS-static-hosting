# 生成占位图标（16/32/80 px 的蓝底白字 S）。仅需运行一次。
Add-Type -AssemblyName System.Drawing
$dir = $PSScriptRoot
$sizes = @(16, 32, 80)
foreach ($s in $sizes) {
    $bmp = New-Object System.Drawing.Bitmap($s, $s)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'
    $g.TextRenderingHint = 'AntiAlias'
    $bg = [System.Drawing.Color]::FromArgb(15, 108, 189)
    $brush = New-Object System.Drawing.SolidBrush($bg)
    $g.FillRectangle($brush, 0, 0, $s, $s)
    $fontSize = [int]($s * 0.6)
    if ($fontSize -lt 6) { $fontSize = 6 }
    $font = New-Object System.Drawing.Font('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = 'Center'
    $sf.LineAlignment = 'Center'
    $rect = New-Object System.Drawing.RectangleF(0, 0, $s, $s)
    $g.DrawString('S', $font, [System.Drawing.Brushes]::White, $rect, $sf)
    $g.Dispose()
    $bmp.Save((Join-Path $dir "icon-$s.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Output "icon-$s.png created"
}
