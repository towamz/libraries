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

'mainで下記のクラスファイルを読み込む

Include("clsGetFilenames.vbs")	'クラスファイルの読み込み
Include("clsRenameFilesTOU.vbs")	'クラスファイルの読み込み

Dim TOU 
Set TOU = New clsRenameFilesTOU


TOU.setRenameDigit=2
TOU.setDirectory="C:\東京通信大学\24-1\テクノロジーマーケティングⅠ\単位認定試験まとめ"
TOU.setPattern=".*.png"
'TOU.setSubLectureNumberEnd=10
TOU.renameFilesTOUMany()



