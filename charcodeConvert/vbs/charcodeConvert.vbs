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
Include("clsGetFilenamesByDir.vbs")	'クラスファイルの読み込み

'変数宣言
Dim objInput,objOutput
Dim objGetFilenameParts, objGetFilenamesByDir
Dim aryFilenames, filename, newFullFilename
Dim tmpStr


'実行確認
If MsgBox("Shift_JISからUTF-8に変換しますか?", vbQuestion + vbYesNo, "実行確認") = vbNo Then
	WScript.Quit
End If



'変換対象ファイル取得
Set objGetFilenamesByDir = New clsGetFilenamesByDir

objGetFilenamesByDir.setDirectory = "C:\sampleMacro\2209charcode\"
objGetFilenamesByDir.setPattern = ".*\.txt"

aryFilenames = objGetFilenamesByDir.getFilenamesByDir()



'変換実行
Set objGetFilenameParts = New clsGetFilenameParts
Set objInput = CreateObject("ADODB.Stream")
Set objOutput = CreateObject("ADODB.Stream")


For Each filename In aryFilenames

	' 文字コード変換後のファイル名を取得
	objGetFilenameParts.setFullFilename = filename
	newFullFilename = objGetFilenameParts.getPath & "convFiles\" & objGetFilenameParts.getFilename
	
	Select Case MsgBox(filename & vbcrlf &"-->" & vbcrlf & newFullFilename, vbQuestion + vbYesNoCancel, "実行確認")
		Case vbYes
			'Shift_JIS形式でファイルを開いて、一時変数へ格納する
			objInput.Charset = "Shift_JIS"
			objInput.Open
			objInput.LoadFromFile filename
			tmpStr = objInput.ReadText
			objInput.Close

			'一時変数を、UTF-8形式でファイルに書き込む
			objOutput.Charset = "UTF-8"
			objOutput.Open
			objOutput.WriteText tmpStr
			objOutput.SaveTofile newFullFilename , 1	'1=同名ファイルがある場合保存しない。2=同名ファイルがある場合上書き 'https://ray88.hatenablog.com/entry/2021/09/19/094953
			objOutput.Close		
		
		Case vbNo
			'処理しないで次のファイルへ
	
		Case vbCancel
			msgbox "ちゅうし"
			WScript.Quit

	End Select

Next




