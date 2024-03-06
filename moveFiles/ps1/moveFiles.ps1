[void][System.Reflection.Assembly]::Load("Microsoft.VisualBasic, Version=8.0.0.0, Culture=Neutral, PublicKeyToken=b03f5f7f11d50a3a")
Add-Type -Assembly System.Windows.Forms

$commonPath = "D:\Photos from "

$yymm = [Microsoft.VisualBasic.Interaction]::InputBox("yymm", "タイトル")

$targetPath = $commonPath + $yymm

$filenames = Get-ChildItem $targetPath -File  | Where-Object {$_.name -match "^IMG_.+\.jpg$"}

foreach ($filename in $filenames) {
	$originalFullfilename = $filename.fullname
	$originalFilename =  $filename.name
	$datePath = $originalFilename.substring(6,4)

	$datePath
	
	# 正規表現=(0000 - 2912 まで)
	if ($datePath -match '^[0-2][0-9][0-1][0-9]$') {
		$dateFullPath = $targetPath + "\" + $datePath
		$targetFullFilename = $dateFullPath + "\" + $originalFilename

		$originalFullfilename
		#$datePath
		$targetFullFilename

		#日付フォルダ(yymm)作成 / make a date folder 
		if(-not(Test-Path $dateFullPath)){
			Switch ([System.Windows.Forms.MessageBox]::Show($datePath,"","YesNoCancel")){
				{$_ -eq [System.Windows.Forms.DialogResult]::Yes}{
				New-Item ($dateFullPath) -ItemType Directory
				break
				}
				{$_ -eq [System.Windows.Forms.DialogResult]::Cancel}{
					exit
					break
				}				 
			}
		}

        # ファイル移動実行 / execute to move files
		Switch ([System.Windows.Forms.MessageBox]::Show($originalFullfilename + "`r`n" + $targetFullFilename,"","YesNoCancel")){
				{$_ -eq [System.Windows.Forms.DialogResult]::Yes}{
					Move-Item $originalFullfilename $targetFullFilename
					break
				}
				{$_ -eq [System.Windows.Forms.DialogResult]::Cancel}{
					exit
					break
				}		 
			}
	}
}
