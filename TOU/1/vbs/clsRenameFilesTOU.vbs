Option Explicit

'mainで下記のクラスファイルを読み込む
'Include("clsGetFilenames.vbs")	'クラスファイルの読み込み
'Include("clsRenameFilesTOU.vbs")	'クラスファイルの読み込み

Class clsRenameFilesTOU

Private GF
Private FSO
Private APP
Private lectureNumberStart
Private lectureNumberEnd
Private lectureSubNumberEnd
Private lectureNumber
Private renameDigit


'検索ディレクトリを設定する
Public Property Let setDirectory(argDirectory)
	GF.setDirectory=argDirectory
End Property

'パターンを設定する
Public Property Let setPattern(argPattern)
	GF.setPattern=argPattern
End Property

'講義番号
Public Property Let setLectureNumber(argNumber)
	lectureNumber=argNumber
End Property

'講義開始番号
Public Property Let setLectureNumberStart(argNumber)
	lectureNumberStart=argNumber
End Property

'講義終了番号
Public Property Let setLectureNumberEnd(argNumber)
	lectureNumberEnd=argNumber
End Property

'桁数
Public Property Let setRenameDigit(argNumber)
	renameDigit=argNumber
End Property

Public Function getSettings()
	dim returnString
	returnString = returnString & "Directory:" & GF.getDirectory() & vbcrlf
	returnString = returnString & "Pattern:" & GF.getPattern() & vbcrlf
	returnString = returnString & "lectureNumber:" & lectureNumber & vbcrlf
	returnString = returnString & "lectureNumberStart:" & lectureNumberStart & vbcrlf
	returnString = returnString & "lectureNumberEnd:" & lectureNumberEnd & vbcrlf
	returnString = returnString & "lectureSubNumberEnd:" & lectureSubNumberEnd & vbcrlf
	returnString = returnString & "setRenameDigit:" & setRenameDigit & vbcrlf



	getSettings = returnString
End Function


Private Sub Class_Initialize()
	Set GF = New clsGetFilenames
	Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")
	Set APP = WScript.CreateObject("Shell.Application")

	lectureNumberStart=1
	lectureNumberEnd=8
	lectureSubNumberEnd=4
	renameDigit=2
End Sub

Private Sub Class_Terminate()
	Set GF = Nothing
	Set FSO = Nothing
	Set APP = Nothing
End Sub


Public Sub renameFilesTOUMany()
	dim currentLectureNumber 
	dim baseDirectory

	'現在設定されているディレクトリを保存
	baseDirectory=GF.getDirectory()

	for currentLectureNumber = lectureNumberStart to lectureNumberEnd
		'ディレクトリに講義番号のサブフォルダを付加する
		GF.setDirectory = FSO.BuildPath(baseDirectory, currentLectureNumber)
		lectureNumber = currentLectureNumber

		'debug用
		'msgbox getSettings()

		'実行する
		renameFilesTOU()
		renameFilesTOU2()
	Next

	'設定を元に戻す
	GF.setDirectory=baseDirectory

End Sub

Public Sub renameFilesTOU()
	Dim aryFilenames,filename,targetFilename,fullfilename
	Dim cnt

	Select Case msgbox(getSettings(),vbOkCancel)
		case vbOk
			APP.Explore GF.getDirectory()
		case vbCancel
			WScript.Quit
	End Select

	aryFilenames = GF.getFilenamesArray()

	'1秒待つ(wait for 1000 milli sec)
	WScript.Sleep 1000
	cnt=InputBox("開始数字>",,1)
	For Each filename In aryFilenames
		'フルパスに変更
		fullfilename = FSO.BuildPath(GF.getDirectory, filename)
		targetFilename = FSO.BuildPath(GF.getDirectory, getPaddingNumber(cnt) & ".png")

		Select Case msgbox(fullfilename & vbcrlf &targetFilename,vbYesNoCancel)
			case vbYes
				Call FSO.MoveFile(fullfilename,targetFilename)
				cnt=cnt+1
			case vbCancel
				WScript.Quit
		End Select
	Next 

End Sub

Public Sub renameFilesTOU2()
	Dim aryFilenames,filename,targetFilename,fullfilename
	Dim lectureSubNumber
	Dim flg

	'-----初期確認-----
	if not IsNumeric(lectureNumber) then
		msgbox "講義番号が入力されていません。中断します"
		WScript.Quit
	elseif lectureNumber = 0 then
		msgbox "講義番号が入力されていません。中断します"
		WScript.Quit
	end if
	'-----初期確認終了-----

	'1でリネームしたのでファイル名を再取得する
	aryFilenames = GF.getFilenamesArray()
	For Each filename In aryFilenames
		if IsNumeric(FSO.GetBaseName(filename)) then
			'フルパスに変更
			fullfilename = FSO.BuildPath(GF.getDirectory, filename)
			APP.ShellExecute(fullfilename)

			Do
				lectureSubNumber=InputBox(lectureNumber & "-?-" & filename)
				if lectureSubNumber = "exit" then
					WScript.Quit
				elseif Not IsNumeric(lectureSubNumber) then
					msgbox "数字以外が入力されました。中断します"
					flg=False
					'WScript.Quit
				elseif CLng(lectureSubNumber) < 1 Then
					msgbox "1-"& lectureSubNumberEnd & "を入力してください。中断します"
					flg=False
					'WScript.Quit
				elseif CLng(lectureSubNumberEnd) < CLng(lectureSubNumber) Then
					msgbox "1-"& lectureSubNumberEnd & "を入力してください。中断します"
					flg=False
					'WScript.Quit
				else
					flg=True
				end if

			Loop until flg

			targetFilename = FSO.BuildPath(GF.getDirectory, lectureNumber & "-" & lectureSubNumber & "-" &  filename)

			Select Case msgbox(fullfilename & vbcrlf &targetFilename,vbYesNoCancel)
				case vbYes
					Call FSO.MoveFile(fullfilename,targetFilename)
				case vbCancel
					WScript.Quit
			End Select
		else
			msgbox "対象外:" & filename
		End IF
	Next 
End Sub

Private Function getPaddingNumber(number)
	dim zeroString
	dim i

	'数字以外が入力されたらそのまま返す
	if not IsNumeric(number) then
		getPaddingNumber=number
		Exit Function
	End if

	'指定の桁数より大きいときはそのまま返す
	if Len(number)>renameDigit then
		getPaddingNumber=number
		Exit Function
	End if

	for i = 0 to renameDigit
		zeroString = zeroString & "0"
	next

	getPaddingNumber = right(zeroString & number,renameDigit)

End Function

End Class