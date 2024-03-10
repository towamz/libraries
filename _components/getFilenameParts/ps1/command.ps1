$path = "c:\temp\test.txt"
$folderTest = "c:\testtest\subf\f"
$filenameTest = "test.png"

$drive         = Split-Path -Qualifier $path
$folderNoDrive = Split-Path -NoQualifier $path
$folder        = [System.IO.Path]::GetDirectoryName($path);
$filename      = [System.IO.Path]::GetFileName($path);
$filenameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($path);
$extension     = [System.IO.Path]::GetExtension($path);
 
Write-Host ($drive);
Write-Host ($folderNoDrive);
Write-Host ($folder);
Write-Host ($filename);
Write-Host ($filenameNoExt);
Write-Host ($extension);

$pathTest      = [System.IO.Path]::GetFullPath($filenameTest);
Write-Host ($pathTest);

$pathJoin      = Join-Path $folderTest $filenameTest
Write-Host ($pathJoin);

<# 
c:
\temp\test.txt
c:\temp
test.txt
test
.txt
C:\getFilenameParts\ps1\test.png
c:\testtest\subf\f\test.png
#>
