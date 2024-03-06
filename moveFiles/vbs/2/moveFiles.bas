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

Dim clsPath
clsPath = "D:\photo\vbs\"

Include(clsPath & "clsGetFilenamesByFSO.vbs")	'クラスファイルの読み込み
Include(clsPath & "clsGetFilenameParts.vbs")	'クラスファイルの読み込み
Include(clsPath & "clsMoveFiles.vbs")	'クラスファイルの読み込み

Dim MF
Dim targetPath
targetPath = "C:\photo"

Set MF = New clsMoveFiles
MF.setDirectory = targetPath
'MF.setPattern = "^IMG_.+\.jpg$"
MF.setPattern = "^IMG_[0-9]{8}_[0-9]{6}\.jpg$"
MF.setMidPosition = 8
MF.moveFiles()
Set MF = Nothing

Set MF = New clsMoveFiles
MF.setDirectory = targetPath
'MF.setPattern = "^VID_.+\.mp4$"
MF.setPattern = "^VID_[0-9]{8}_[0-9]{6}\.mp4$"
MF.setMidPosition = 8
MF.moveFiles()
Set MF = Nothing

Set MF = New clsMoveFiles
MF.setDirectory = targetPath
MF.setPattern = "^IMG[0-9]{14}\.jpg$"
MF.setMidPosition = 7
MF.moveFiles()
Set MF = Nothing

Set MF = New clsMoveFiles
MF.setDirectory = targetPath
MF.setPattern = "^VID[0-9]{14}\.mp4$"
MF.setMidPosition = 7
MF.moveFiles()
Set MF = Nothing

Set MF = New clsMoveFiles
MF.setDirectory = targetPath
MF.setPattern = "\.json$"
MF.deleteFiles()
Set MF = Nothing

