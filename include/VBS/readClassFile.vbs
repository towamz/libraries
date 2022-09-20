Option Explicit

'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272

Function Include(strFile)
	'strFile：読み込むvbsファイルパス
 
	Dim objFso, objWsh, strPath
	Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	Set objWsh = objFso.OpenTextFile(strFile)
	ExecuteGlobal objWsh.ReadAll()
	objWsh.Close
 
	Set objWsh = Nothing
	Set objFso = Nothing
 
End Function


Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


Dim objGetFilenameParts

Set objGetFilenameParts = New clsGetFilenameParts

' クラスの関数を呼出し
objGetFilenameParts.setFullFilename = "C:\sjis.txt"


msgbox objGetFilenameParts.getFilenameNoExt()



