Option Explicit

'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272

Function Include(strFile)
	'strFile：読み込むvbsファイルパス
 
	Dim FSO, TF, strPath
	Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	Set TF = FSO.OpenTextFile(strFile)
	ExecuteGlobal TF.ReadAll()
	TF.Close
 
	Set TF = Nothing
	Set FSO = Nothing
 
End Function


Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


Dim objGetFilenameParts

Set objGetFilenameParts = New clsGetFilenameParts

' クラスの関数を呼出し
objGetFilenameParts.setFullFilename = "C:\sjis.txt"


msgbox objGetFilenameParts.getFilename()


objGetFilenameParts.setDelimiter="/"
objGetFilenameParts.setFullFilename = "https://ja.wikipedia.org/wiki/kugiimoji.html"


msgbox objGetFilenameParts.getFilename()



