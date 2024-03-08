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


Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


'-----これより上をコピー-----
'-----copy above-----
'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□

Dim objGetFilenameParts

Set objGetFilenameParts = New clsGetFilenameParts

' クラスの関数を呼出し
objGetFilenameParts.setFullFilename = "C:\sjis.txt"


msgbox objGetFilenameParts.getFilename()


objGetFilenameParts.setDelimiter="/"
objGetFilenameParts.setFullFilename = "https://ja.wikipedia.org/wiki/kugiimoji.html"


msgbox objGetFilenameParts.getFilename()



