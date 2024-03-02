Option Explicit

'外部ファイル読み込み参照サイト
'http://cloiwan.com/?p=272

Function Include(strFile)
	'strFile：読み込むvbsファイルパス
	Dim FSO, objWsh, strPath
	Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
	
	'外部ファイルの読み込み
	Set objWsh = FSO.OpenTextFile(strFile)
	ExecuteGlobal objWsh.ReadAll()
	objWsh.Close
 
	Set objWsh = Nothing
	Set FSO = Nothing
End Function


Include("clsGetFilenamesByFSO.vbs")	'クラスファイルの読み込み
Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


Dim FSO
Dim GF,GFP
Dim aryFilenames
Dim filename
Dim foldername
Dim targetFilename


Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
Set GF = New clsGetFilenamesByFSO
Set GFP = New clsGetFilenameParts

'操作対象フォルダとファイルパターンを設定
GF.setDirectory = "D:\picture"
GF.setPattern = "\.jpg$"

'実行前確認
If Inputbox("処理を続行する場合はexecuteを入力してください" & vbcrlf & _
			"処理対象フォルダ:" & GF.getDirectory & vbcrlf & _
			"処理対象パターン:" & GF.getPattern) <> "execute" Then
	MsgBox "処理を中断します",vbYesNo
	WScript.Quit
End If

'ファイル名取得
aryFilenames = GF.getFilenamesByFSO()

For Each filename In aryFilenames
	'GFP.setDelimiter="/"
	GFP.setFullFilename = filename
	
	'写真ファイル名からyymm部分を取り出す
	foldername = Mid(filename, GF.getDirectoryLen + 7, 4)

	'移動先フォルダが存在しない場合は作成する
	If Not FSO.FolderExists(GF.getDirectory & "\" & foldername) then
		If msgbox(foldername & "フォルダを作成しますか",vbYesNo) = vbYes Then
			FSO.CreateFolder(GF.getDirectory & "\" & foldername)
		Else
			MsgBox "処理を中断します"
			WScript.Quit
		End If
	End if
	
	targetFilename = GFP.getPath & foldername & "\" & GFP.getFilename

	'Select Case msgbox( filename & vbcrlf & targetFilename,vbYesNoCancel)
	'	Case vbYes
			Call FSO.MoveFile(filename,targetFilename)
	'	Case vbNo
	'		'ファイル移動をせずに次のファイルに移る
	'	Case vbCancel
	'		MsgBox "処理を中断します"
	'		WScript.Quit
	'End Select
Next 

Msgbox "処理が終了しました"
