Option Explicit

'下記3つのクラスを読み込むこと
'Include("clsGetFilenamesByFSO.vbs")
'Include("clsGetFilenameParts.vbs")
'Include("clsMoveFiles.vbs")

Class clsMoveFiles
	Private FSO
	Private GF,GFP
	Private LNG_MidPosition
	private LNG_NoAlertNum
	private cnt

	Public Property Let setDirectory(argDirectory)
		GF.setDirectory = argDirectory
	End Property

	Public Property Let setPattern(argPattern)
		GF.setPattern = argPattern
	End Property

	Public Property Let setMidPosition(argMidPosition)
	    LNG_MidPosition = argMidPosition
	End Property

	Public Property Let setNoAlertNum(argNoAlertNum)
	    LNG_NoAlertNum = argNoAlertNum
	End Property


	Public Property Get getDirectory()
	    getDirectory = GF.getDirectory
	End Property

	Public Property Get getPattern()
	    getPattern = GF.getPattern
	End Property

	Public Property Get getMidPosition()
	    getMidPosition = LNG_MidPosition
	End Property

	Public Property Get getNoAlertNum()
	    getNoAlertNum = LNG_NoAlertNum
	End Property

	Public Property Get getCnt()
	    getCnt = cnt
	End Property


	'コンストラクタ
	Private Sub Class_Initialize()
		Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
		Set GF = New clsGetFilenamesByFSO
		Set GFP = New clsGetFilenameParts
		LNG_NoAlertNum = 10	'既定値10回
	End Sub

	Private Sub Class_Terminate()
		Set GFP = Nothing
		Set GF = Nothing
		Set FSO = Nothing
	End Sub

	Public Sub showSetting()
		msgbox  "処理対象フォルダ:" & GF.getDirectory & vbcrlf & _
				"処理対象パターン:" & GF.getPattern & vbcrlf & _
				"切り取り開始位置:" & LNG_MidPosition & vbcrlf & _
				"アラート停止回数:" & LNG_NoAlertNum
	End Sub	

	Public Sub moveFiles()
		Dim aryFilenames
		Dim filename
		Dim foldername
		Dim targetFilename
		
		'パラメータ確認
		If GF.getDirectory = "" OR GF.getPattern = "" OR LNG_MidPosition = ""  Then  
			err.raise 1001
		End If

		'実行前確認
		If Inputbox("処理を続行する場合はexecuteを入力してください" & vbcrlf & _
					"処理対象フォルダ:" & GF.getDirectory & vbcrlf & _
					"処理対象パターン:" & GF.getPattern & vbcrlf & _
					"切り取り開始位置:" & LNG_MidPosition & vbcrlf & _
					"アラート停止回数:" & LNG_NoAlertNum) <> "execute" Then
			MsgBox "処理を中断します",vbYesNo
			WScript.Quit
		End If

		'ファイル名取得
		aryFilenames = GF.getFilenamesByFSO()
		cnt = 0

		For Each filename In aryFilenames
			cnt = cnt + 1
			'GFP.setDelimiter="/"
			GFP.setFullFilename = filename
			
			'写真ファイル名からyymm部分を取り出す
			foldername = Mid(filename, GF.getDirectoryLen + LNG_MidPosition, 4)

			'移動先フォルダが存在しない場合は作成する
			If Not FSO.FolderExists(GF.getDirectory & "\" & foldername) then
				If msgbox(filename & vbcrlf & foldername & "フォルダを作成しますか",vbYesNo) = vbYes Then
					FSO.CreateFolder(GF.getDirectory & "\" & foldername)
				Else
					MsgBox "処理を中断します"
					WScript.Quit
				End If
			End if
			
			targetFilename = GFP.getPath & foldername & "\" & GFP.getFilename

			if cnt <= LNG_NoAlertNum then
				Select Case msgbox( filename & vbcrlf & targetFilename,vbYesNoCancel)
					Case vbYes
						Call FSO.MoveFile(filename,targetFilename)
					Case vbNo
						'ファイル移動をせずに次のファイルに移る
					Case vbCancel
						MsgBox "処理を中断します"
						WScript.Quit
				End Select

				if cnt = LNG_NoAlertNum then
					If Inputbox("これ以降はメッセージが表示されず自動実行されます。executeを入力してください") <> "execute" Then
						MsgBox "処理を中断します",vbYesNo
						WScript.Quit
					End If
				end if 
			Else
				Call FSO.MoveFile(filename,targetFilename)

				if cnt > 10000 Then
					cnt = LNG_NoAlertNum
				End If 
			End If
		Next 

		Msgbox "処理が終了しました"
	End Sub

	Public sub deleteFiles()
		Dim aryFilenames
		Dim filename

		'パラメータ確認
		If GF.getDirectory = "" OR GF.getPattern = "" Then  
			err.raise 1001
		End If

		'実行前確認
		If Inputbox("処理を続行する場合はexecuteを入力してください" & vbcrlf & _
					"処理対象フォルダ:" & GF.getDirectory & vbcrlf & _
					"処理対象パターン:" & GF.getPattern & vbcrlf & _
					"アラート停止回数:" & LNG_NoAlertNum) <> "execute" Then
			MsgBox "処理を中断します",vbYesNo
			WScript.Quit
		End If

		'ファイル名取得
		aryFilenames = GF.getFilenamesByFSO()
		cnt = 0


		For Each filename In aryFilenames
			cnt = cnt + 1
			If cnt <= LNG_NoAlertNum then
				Select Case msgbox(filename,vbYesNoCancel)
					Case vbYes
						Call FSO.DeleteFile(filename)
					'Case vbNo
						'ファイル移動をせずに次のファイルに移る
					Case vbCancel
						MsgBox "処理を中断します"
						WScript.Quit
				End Select

				if cnt = LNG_NoAlertNum then
					If Inputbox("これ以降はメッセージが表示されず自動実行されます。executeを入力してください") <> "execute" Then
						MsgBox "処理を中断します",vbYesNo
						WScript.Quit
					End If
				end if 			
			Else
				Call FSO.DeleteFile(filename)

				if cnt > 10000 Then
					cnt = LNG_NoAlertNum
				End If 
			End If		
		Next 

		Msgbox "処理が終了しました"

	End Sub
End Class