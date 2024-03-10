'http://wsh.style-mods.net/ref_filesystemobject/index.htm
Option Explicit

'カレントディレクトリをこのファイルがあるディレクトリに変更
'change current directory to where this file is
Dim executeDirectory
With Wscript.createObject("Scripting.FileSystemObject")
	executeDirectory =  .getParentFolderName(WScript.ScriptFullName)
End With 

With Wscript.CreateObject("WScript.shell")
	.CurrentDirectory = executeDirectory
End With 


Dim fullfilename
Dim resultString
Dim FL

fullfilename = "C:\test\script\testname.txt"
resultString = fullfilename & vbcrlf
With CreateObject("Scripting.FileSystemObject")
    Set FL = .CreateTextFile(.GetTempName())

    resultString = resultString & vbcrlf &  .GetDriveName(fullfilename)
    resultString = resultString & vbcrlf &  .GetParentFolderName(fullfilename)
    resultString = resultString & vbcrlf &  .GetFileName(fullfilename)
    resultString = resultString & vbcrlf &  .GetBaseName(fullfilename)
    resultString = resultString & vbcrlf &  .GetExtensionName(fullfilename)

    resultString = resultString & vbcrlf &  .GetAbsolutePathName("absPath.txtx")
    resultString = resultString & vbcrlf &  .GetTempName()

    msgbox resultString

    FL.WriteLine(fullfilename)
    FL.WriteLine()
    FL.WriteLine(.GetDriveName(fullfilename))
    FL.WriteLine(.GetParentFolderName(fullfilename))
    FL.WriteLine(.GetFileName(fullfilename))
    FL.WriteLine(.GetBaseName(fullfilename))
    FL.WriteLine(.GetExtensionName(fullfilename))
    FL.WriteLine(.GetAbsolutePathName("absPath.txtx"))
    FL.WriteLine(.GetTempName())
    FL.Close
End with