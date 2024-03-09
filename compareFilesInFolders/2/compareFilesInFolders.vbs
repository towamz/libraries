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


'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272
Function Include(strFile)
	Dim TF

	With Wscript.CreateObject("Scripting.FileSystemObject")
		Set TF = .OpenTextFile(strFile)
	End With

	ExecuteGlobal TF.ReadAll()
	TF.Close

	Set TF = Nothing
End Function


Include("clsGetFilenames.vbs")	'クラスファイルの読み込み
Include("clsCompareFilesInFolders.vbs")	'クラスファイルの読み込み


Dim CFIF
Set CFIF = New clsCompareFilesInFolders

CFIF.setDirectory1 = "C:\compFiles\2\folder1"
CFIF.setDirectory2 = "C:\compFiles\2\folder2"
CFIF.setPattern = "\.png$"
CFIF.setFilenameResult1 = "C:\compFiles\2\result1Only.txt"
CFIF.setFilenameResult2 = "C:\compFiles\2\result2Only.txt"
CFIF.setFilenameResultBoth = "C:\compFiles\2\resultBoth.txt"


CFIF.compareFilesInFolders()

MsgBox "終了"