Set-Location -Path  C:\
$array = Get-ChildItem -File | Where-Object {$_.Name -match '.txt$'}

foreach($a in $array) 
{
	$newFilename = ".\convFiles\" + $a.Name
	echo $newFilename

	#default = sjis
	Get-Content $a.Name -Encoding UTF8|Out-File -FilePath $newFilename -Encoding default

} 