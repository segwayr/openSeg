
while ($true) {
	cls
	Add-Type -AssemblyName System.Windows.Forms
	Add-Type -AssemblyName System.Drawing

	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Over Rider"
	$form.Size = New-Object System.Drawing.Size(225, 250)
	$form.FormBorderStyle = "Fixed3D"


	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Location = New-Object System.Drawing.Point(20, 170)
	$OKButton.Size = New-Object System.Drawing.Size(75, 30)
	$OKButton.Text = "OK"
	$OKButton.DialogResult = "OK"


	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Location = New-Object System.Drawing.Point(110, 170)
	$CancelButton.Size = New-Object System.Drawing.Size(75, 30)
	$CancelButton.Text = "Cancel"
	$CancelButton.DialogResult = "Cancel"

	#日付
	$labelCal = New-Object System.Windows.Forms.Label
	$labelCal.Location = New-Object System.Drawing.Point(20, 10)
	$labelCal.Size = New-Object System.Drawing.Size(50, 20)
	$labelCal.Text = "日付"


	$calBox = New-Object System.Windows.Forms.DatetimePicker
	$calBox.Location = New-Object System.Drawing.Point(20, 30)
	$calBox.Size = New-Object System.Drawing.Size(110, 50)

	#時
	$labelHour = New-Object System.Windows.Forms.Label
	$labelHour.Location = New-Object System.Drawing.Point(130, 10)
	$labelHour.Size = New-Object System.Drawing.Size(50, 20)
	$labelHour.Text = "時"

	$hourBox = New-Object System.Windows.Forms.NumericUpDown
	$hourBox.Location = New-Object System.Drawing.Point(130, 30)
	$hourBox.Size = New-Object System.Drawing.Size(55, 50)
	$hourBox.text = 9
	$hourBox.TextAlign = "Right"
	$hourBox.UpDownAlign = "Right"
	$hourBox.Maximum = "23"
	$hourBox.Minimum = "0"
	$hourBox.InterceptArrowKeys = $True


	#分の最小値
	$labelMinuteMin = New-Object System.Windows.Forms.Label
	$labelMinuteMin.Location = New-Object System.Drawing.Point(20, 60)
	$labelMinuteMin.Size = New-Object System.Drawing.Size(65, 20)
	$labelMinuteMin.Text = "最小"

	$minuteMinBar = New-Object System.Windows.Forms.HScrollBar
	$minuteMinBar.Location = "85, 60"
	$minuteMinBar.size = "100, 15"
	$minuteMinBar.maximum = 59
	$minuteMinBar.minimum = 0
	$minuteMinBar.largechange = "1"
	$minuteMinBar.value = "20"
	$labelMinuteMin.text = "最小 " + $minuteMinBar.value + "分"
	$minuteMinBar.Add_ValueChanged({
		$labelMinuteMin.text = "最小 " + $minuteMinBar.value + "分"
		if ($minuteMaxBar.value -lt $minuteMinBar.value) {
			$minuteMaxBar.value = $minuteMinBar.value
		}
	})

	#分の最大値
	$labelMinuteMax = New-Object System.Windows.Forms.Label
	$labelMinuteMax. Location = New-Object System.Drawing.Point(20, 80)
	$labelMinuteMax.Size = New-Object System.Drawing.Size(75, 20)
	$labelMinuteMax.Text = "最大"


	$minuteMaxBar = New-Object System.Windows.Forms.HScrollBar
	$minuteMaxBar.Location = "85,80"
	$minuteMaxBar.size = "100, 15"
	$minuteMaxBar.maximum = 59
	$minuteMaxBar.minimum = 0
	$minuteMaxBar.largechange = "1"
	$minuteMaxBar.value = "59"
	$labelMinuteMax.text = "最大 " + $minuteMaxBar.value + "分"
	$minuteMaxBar.Add_ValueChanged({
		$labelMinuteMax.text = "最大 " + $minuteMaxBar.value + "分"
		if ($minuteMaxBar.value -lt $minuteMinBar.value) {
			$minuteMinBar.value = $minuteMaxBar.value
		}
	})

	#グループを作る    
	$radioGr = New-Object System.Windows.Forms.GroupBox
	$radioGr.Location = New-Object System.Drawing.Point(20, 100)
	$radioGr.size = New-Object System.Drawing.Size(165, 60)
	$radioGr.text = "設定"

	#グループの中のラジオボタンを作る
	$normaler = New-Object System.Windows.Forms.RadioButton
	$normaler.Location = New-Object System.Drawing.Point(20, 15)
	$normaler.size = New-Object System.Drawing.Size(140, 20)
	$normaler.Checked = $True
	$normaler.Text = "更新日のみ修正"

	$autoMaker = New-Object System.Windows.Forms.RadioButton
	$autoMaker.Location = New-Object System.Drawing.Point(20, 35)
	$autoMaker.size = New-Object System.Drawing.Size(140, 20)
	$autoMaker.Text = "作成日自動調整"
	$radioGr.Controls.AddRange(@($normaler, $autoMaker))

	#フォームの各アイテムを入れる       
	$Form.Controls.AddRange(@($radioGr))

	#フォームのロード
	$form.AcceptButton = $OKButton
	$form.CancelButton = $CancelButton
	$form.Controls.Add($OKButton)
	$form.Controls.Add($CancelButton)
	$form.Controls.Add($labelCal)
	$form.Controls.Add($calBox)
	$form.Controls.Add($labelHour)
	$form.Controls.Add($hourBox)
	$form.Controls.Add($minuteMinBar)
	$form.Controls.Add($labelMinuteMin)
	$form.Controls.Add($minuteMaxBar)
	$form.Controls.Add($labelMinuteMax)

	#イベントの戻り値等
	$result = $form.ShowDialog()
	if ($result -eq "Cancel") {
		exit
	} elseif ($hourBox.text -eq ""){
		#loop
	} elseif ($result -eq "OK") {
		break
	}
}

$yearer = ([Datetime]$calBox.text).ToString("yyyy")
$monther = ([Datetime]$calBox.text).ToString("MM")
$dater = ([Datetime]$calBox.text).ToString("dd")
$hourer = $hourBox.Text
#乱数は1が0なのでインクリメント
$max = $minuteMaxBar.value + 1
$min = $minuteMinBar.value + 1

#分数は最小～最大指定、値が同じ場合乱数発生はエラーが出るので回避して最小値を入れる
if ($min -eq $max) {
	$miniter = $min - 1
} else {
	$miniter = Get-Random -Maximum $max -Minimum $min
}

#秒数は0～59秒ランダム
$secer = Get-Random -Maximum 60 -Minimum 1

#作成日時を指定
$yearerMaker = $yearer
$montherMaker = $monther
$daterMaker = $dater

#計算後に分がマイナスになった場合60分加算させる
$hourerMaker = $hourer
$temp = Get-Random -Maximum 21 -Minimum 11
$miniterMaker = $miniter - $temp
$secerMaker = Get-Random -Maximum 60 -Minimum 1

#さらに0時下回った場合日付ごと-1する
if ($miniterMaker -lt 0) {
	$miniterMaker = 60 + $miniterMaker
	$hourerMaker = $hourerMaker - 1
	# X  0          ?   t    -1    
	if ($hourerMaker -lt 0) {
		$hourerMaker = 24 + $hourerMaker
		$yearerMaker = (([Datetime]$calBox.text).AddDays(-1)).ToString("yyyy")
		$montherMaker = (([Datetime]$calBox.text).AddDays(-1)).ToString("MM")
		$daterMaker = (([Datetime]$calBox.text).AddDays(-1)).ToString("dd")
	}
}


#なんかあったらエラーログ吐いてね
try {
		Set-ItemProperty $Args[0] -name LastWriteTime -value "$($yearer)/$($monther)/$($dater) $($hour):$($miniter):$($secer)"
		if ($autoMaker.Checked) {
			Set-ItemProperty $Args[0] -name CreationTime -value "$($yearerMaker)/$($montherMaker)/$($daterMaker) $($hourMaker):$($miniterMaker):$($secerMaker)"

		}
} catch {
	Write-output $error
	$error>>errorLog.txt
}
