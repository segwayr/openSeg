#カーソル位置の画面座標と色をリアルタイムで教えてくれる
#座標操作や色分岐が必要なインターフェースがない、オブジェクトをキャッチできないようなアプリや業務を私用マクロ化する時など

Add-Type -AssemblyName System.Drawing, System.Windows.Forms
$bitmap = New-Object System.Drawing.Bitmap(1, 1)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$Timer = New-Object System.Windows.Forms.Timer
$dpiX = [Systen.Windows.Form.Screen]::PrimaryScreen.Bounds.Width / [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
$scalePercent = [Math]::Round($dpiX * 100)

$TimerTick= {
    $dpi = [Math]::Round([int]$dpiScaling.text.Substring(0, 3) * 0.01)
    $position = [System.Windows.Forms.Cursor]::Position
    $graphics.CopyFromScreen($position. X * $dpi, $position. Y * $dpi, 0, 0, $bitmap.Size)
    $pixel = $bitmap.GetPixel(0, 0)
    $decValue = "RGB({0}, {1}, {2})" -f $pixel.R, $pixel.G, $pixel.B
    $label.Text = "X:" + $position.X * $dpi + "Y:" + $position.Y * $dpi + "`n" + $decValue
}


$Timer.Add_Tick($TimerTick)
$Timer.Interval = 200
$Timer.Enabled = $TRUE
$Timer.Start()

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "color"
$Form.Size = New-Object System.Drawing.Size(260, 140)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point (10, 10)
$label.Size = New-Object System.Drawing.Size(160, 40)
$label.Font = New-Object System.Drawing.Font("ＭＳ ゴシック", 10)

#プルダウン
$pullLabel = New-Object System.Windows.Forms.Label
$pullLabel.Location = New-Object System.Drawing.Point(10,55)
$pullLabel.Size = New-Object System.Drawing.Size(70,20)
$pullLabel.Font = New-Object System.Drawing.Font("ＭＳ ゴシック", 8)
$pullLabel.text = "拡大/縮小"

$dpiScaling = New-Object System.Windows.Forms.Combobox
$dpiScaling.Location = New-Object System.Drawing.Point (10,75)
$dpiScaling.size = New-Object System.Drawing.Size(50,30)
$dpiScaling.DropDownStyle = "DropDown"
$dpiScaling.FlatStyle = "standard"
$dpiScaling.font = $Font
$dpiScaling.text = "$scalePercent%"

[void] $dpiScaling.Items.Add("100%")
[void] $dpiScaling.Items.Add("125%")
[void] $dpiScaling.Items.Add("150%")
[void] $dpiScaling.Items.Add("175%")

$Form.Controls.Add($label)
$Form.Controls.Add($pullLabel)
$Form.Controls.Add($dpiScaling)
$Form.ShowDialog()
