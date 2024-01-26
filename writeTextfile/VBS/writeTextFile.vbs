Dim objFso,objFile
Dim strFile

strFile = "C:\result.txt"

Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")

'ファイルが存在するときは開く/ Open the file if it exists
If objFso.FileExists(strFile) Then
    '上書き / overwrite
    'Set objFile = objFso.OpenTextFile(strFile, 2)
    '追記 / append
    Set objFile = objFso.OpenTextFile(strFile, 8)
'ファイルが存在しないときは作成する / Make the file if it does not exist
Else
    Set objFile = objFso.CreateTextFile(strFile)
End If

objFile.WriteLine("test")
objFile.Close

Set objFso = Nothing
Set objFile = Nothing
