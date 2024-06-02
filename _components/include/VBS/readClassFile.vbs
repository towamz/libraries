Option Explicit

'カレントディレクトリをこのファイルがあるディレクトリに変更
'change current directory to where this file is
Dim FSO, SHL

Set FSO = Wscript.createObject("Scripting.FileSystemObject")
Set SHL = Wscript.CreateObject("WScript.shell")

SHL.CurrentDirectory = FSO.getParentFolderName(WScript.ScriptFullName)


'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272
Function Include(strFile)
	Dim OTF

	Set OTF = FSO.OpenTextFile(strFile)

	ExecuteGlobal OTF.ReadAll()
	OTF.Close

	Set OTF = Nothing
End Function


Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


'-----これより上をコピー-----
'-----copy above-----
'□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□□

Dim GFP

Set GFP = New clsGetFilenameParts

' クラスの関数を呼出し
GFP.setFullFilename = "C:\sjis.txt"


msgbox GFP.getFilename()


GFP.setDelimiter="/"
GFP.setFullFilename = "https://ja.wikipedia.org/wiki/kugiimoji.html"


msgbox GFP.getFilename()


