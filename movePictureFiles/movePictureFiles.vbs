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


Include("clsGetFilenamesByFSO.vbs")	'クラスファイルの読み込み
Include("clsGetFilenameParts.vbs")	'クラスファイルの読み込み


Dim objFso
Dim objGetFilenames,objGetFilenameParts
Dim aryFilenames
Dim filename
Dim foldername
Dim targetFilename



Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")
Set objGetFilenames = New clsGetFilenamesByFSO
Set objGetFilenameParts = New clsGetFilenameParts

'操作対象フォルダとファイルパターンを設定
objGetFilenames.setDirectory = "C:\pictureFolder\"
objGetFilenames.setPattern = ".*"

'実行前確認
If Inputbox("処理を続行する場合はexecuteを入力してください" & vbcrlf & _
			"処理対象フォルダ:" & objGetFilenames.getDirectory & vbcrlf & _
			"処理対象パターン:" & objGetFilenames.getPattern) <> "execute" Then
	MsgBox "処理を中断します"
	WScript.Quit
End If

'ファイル名取得
aryFilenames = objGetFilenames.getFilenamesByFSO()


For Each filename In aryFilenames
	'objGetFilenameParts.setDelimiter="/"
	objGetFilenameParts.setFullFilename = filename
	
	'写真ファイル名からyymm部分を取り出す
	foldername = Mid(filename, objGetFilenames.getDirectoryLen + 6, 4)

	'移動先フォルダが存在しない場合は作成する
	If Not objFso.FolderExists(objGetFilenames.getDirectory & foldername) then
	
		If msgbox(foldername & "フォルダを作成しますか",vbYesNo) = vbYes Then
			objFSO.CreateFolder(objGetFilenames.getDirectory & foldername)
		Else
			MsgBox "処理を中断します"
			WScript.Quit
			
		End If

	End if
	
	
	targetFilename = objGetFilenameParts.getPath & foldername & "\" & objGetFilenameParts.getFilename

	'Select Case msgbox( filename & vbcrlf & targetFilename,vbYesNoCancel)
	'	Case vbYes
			Call objFSO.MoveFile(filename,targetFilename)
	'	Case vbNo
	'		'ファイル移動をせずに次のファイルに移る
	'	Case vbCancel
	'		MsgBox "処理を中断します"
	'		WScript.Quit
	'End Select

Next 

Msgbox "処理が終了しました"








