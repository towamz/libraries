#参照サイト
#https://www.web-dev-qa-db-ja.com/ja/encoding/powershell%E3%82%92%E4%BD%BF%E7%94%A8%E3%81%97%E3%81%A6bom%E3%81%AA%E3%81%97%E3%81%A7%E3%83%95%E3%82%A1%E3%82%A4%E3%83%AB%E3%82%92utf8%E3%81%A7%E6%9B%B8%E3%81%8D%E8%BE%BC%E3%82%80/971230646/

Set-Location -Path  C:\sampleMacro\2209charcode
$array = Get-ChildItem -Path .\*.txt -File #-Include *.txt #| Where-Object {$_.Name -match '.txt$'}

$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False

foreach($a in $array) 
{
	$newFilename = "..\convFiles\" + $a.Name
	echo $newFilename

	$FileContent = Get-Content $a.Name -Encoding UTF8
	[System.IO.File]::WriteAllLines($newFilename, $FileContent, $Utf8NoBomEncoding)

} 