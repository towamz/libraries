Option Explicit

'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272

Function Include(strFile)
	'strFile：読み込むvbsファイルパス
 
	Dim FSO, objWsh, strPath
	Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	'msgbox strFile
	Set objWsh = FSO.OpenTextFile(strFile)
	ExecuteGlobal objWsh.ReadAll()
	objWsh.Close
 
	Set objWsh = Nothing
	Set FSO = Nothing
 
End Function


Include("clsGetFilenamesByFSO.vbs")	'クラスファイルの読み込み

Dim targetPath
Dim FG1,FG2
Set FG1 = New clsGetFilenamesByFSO
Set FG2 = New clsGetFilenamesByFSO

targetPath = InputBox("yymm")

' クラスの関数を呼出し
FG1.setDirectory = "D:\folder1\" & targetPath
FG2.setDirectory = "D:\folder2\" & targetPath
FG1.setPattern = ".+"
FG2.setPattern = ".+"


Dim FSO
Dim strFile1,strFile2
Dim T1,T2

strFile1 = "C:\result1.txt"
strFile2 = "C:\result2.txt"

Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")

Set T1 = FSO.OpenTextFile(strFile1, 2)
T1.WriteLine(Join(FG1.getFilenamesByFSO(),vbCrLf))
T1.Close

Set T2 = FSO.OpenTextFile(strFile2, 2)
T2.WriteLine(Join(FG2.getFilenamesByFSO(),vbCrLf))
T2.Close


msgbox "終了"
