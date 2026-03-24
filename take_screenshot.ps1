Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$bitmap = New-Object System.Drawing.Bitmap($screen.Width, $screen.Height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
$outPath = "C:\Users\agentcode01\Desktop\SalesMgr\screenshots\desktop_test.png"
$bitmap.Save($outPath)
Write-Host "Saved: $outPath"
$r = [System.Drawing.ColorTranslator]::FromOle([System.Drawing.Bitmap]::new($outPath).GetPixel(100,100).ToArgb())
Write-Host "Pixel at 100,100: R=$($r.R) G=$($r.G) B=$($r.B)"
