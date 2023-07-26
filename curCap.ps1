Add-Type -AssemblyName System.Drawing, System.Windows.Forms
$bitmap = New-Object System.Drawing.Bitmap(1, 1)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)
$Timer = New-Object System.Windows.Forms.Timer

$TimerTick= {
    $dpi = [int]$dpiScaling.text.Substring(0, 3) * 0.01
    $position = [System.Windows.Forms.Cursor]::Position
    $graphics.CopyFromScreen($position. X* $dpi, $position. Y * $dpi, 0, 0, $bitmap.Size)
    $pixel = $bitmap.GetPixel(0, 0)
    $decValue = "RGB([0], [1], [2])" -f $pixel.R, $pixel.G, $pixel.B
    $label.Text = "X:" + $position.X + "Y:" + $position. Y + $decValue
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

#ブルダウン
$pullLabel = New-Object System.Windows.Forms.Labele
$pullLabel.Location = New-Object System.Drawing.Point(10,55)
$pullLabel.Size = New-Object System.Drawing.Size(50,10)
$pullLabel.Font = New-Object System.Drawing.Font("ＭＳ ゴシック", 8)
$pullLabel.text = “画面倍率”

$dpiScaling = New-Object System.Windows.Forms.Combobox
$dpiScaling.Location = New-Object System.Drawing.Point (10,70)
$dpiScaling.size = New-Object System.Drawing.Size(50,30)
$dpiScaling.DropDownStyle = "DropDown"
$dpiScaling.FlatStyle= "standard"
$dpiScaling.font = $Font
$dpiScaling.text = "125%"

[void] $dpiScaling.Items.Add("100%")
[void] $dpiScaling.Items.Add("125%")
[void] $dpiScaling.Items.Add("150%")
[void] $dpiScaling.Items.Add("175%")

$Form.Controls.Add($label)
$Form.Controls.Add($pullLabel)
$Form.Controls.Add($dpiScaling)
$Form.ShowDialog()
