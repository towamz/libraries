Option Explicit

'Include("clsGetFilenames.vbs")
'Include("clsCompareFilesInFolders.vbs")

Class clsCompareFilesInFolders

private GFs1
private	GFs2
private executeDirectory
private result1textfile
private result2textfile
private resultBothtextfile
private resultLogfile

'検索ディレクトリを設定する
Public Property Let setDirectory1(argDirectory)
	GFs1.setDirectory = argDirectory
End Property

Public Property Let setDirectory2(argDirectory)
	GFs2.setDirectory = argDirectory
End Property

Public Property Let setPattern(argPattern)
	GFs1.setPattern = argPattern
	GFs2.setPattern = argPattern
End Property

Public Property Let setFilenameResult1(argFilename)
	result1textfile = argFilename
End Property

Public Property Let setFilenameResult2(argFilename)
	result2textfile = argFilename
End Property

Public Property Let setFilenameResultBoth(argFilename)
	resultBothtextfile = argFilename
End Property

Public Property Let setFilenameResultLog(argFilename)
	resultLogfile = argFilename
End Property

Private Sub Class_Initialize()
	Set GFs1 = New clsGetFilenames
	Set GFs2 = New clsGetFilenames

	'既定値はすべてのファイル / defalut is all files
	GFs1.setPattern = ".*"
	GFs2.setPattern = ".*"

	call setExecuteDirectory()

	'ログファイルの既定値 / dafault logfile
	result1textfile = executeDirectory & "/" & "result1.txt"
	result2textfile = executeDirectory & "/" & "result2.txt"
	resultBothtextfile = executeDirectory & "/" & "resultBoth.txt"
	resultLogfile = executeDirectory & "/" & "resultLog.txt"
End Sub

Private Sub Class_Terminate()
	Set GFs1 = Nothing
	Set GFs2 = Nothing	
End Sub

Private sub setExecuteDirectory()
	With Wscript.createObject("Scripting.FileSystemObject")
		executeDirectory =  .getParentFolderName(WScript.ScriptFullName)
	End With 
End Sub


Public Sub compareFilesInFolders()
	Dim ary2
	Dim dic1,dic2,dicBoth
	Dim filename

	Call checkParam

	'フォルダ1のディクショナリを生成する / generate dictionary for dir1
	set dic1 = GFs1.getFilenamesDictionary()
	'フォルダ2の配列を1つ生成する / generate array for dir2
	ary2 = GFs2.getFilenamesArray()

	'差分格納用ディクショナリ生成 / generate dictionaries for diffrences
	Set dic2 = CreateObject("Scripting.Dictionary")
	Set dicBoth = CreateObject("Scripting.Dictionary")

	'dir2のファイル名をdir1のdicに存在するか判定する / 
	'judge files in dir2 exist in dir1　
	For each filename in ary2
		'両方のフォルダにある場合 / both of dirs
		If(dic1.Exists(filename))Then
			dicBoth.Add filename,0
			dic1.Remove filename
		'フォルダ1にない・フォルダ2にある場合 / only in dir2
		Else
			dic2.Add filename,0
		End If
	Next

	call writeLog(dic1.count,dic2.count,dicBoth.count)
	call writeResult(result1textfile,dic1)
	call writeResult(result2textfile,dic2)
	call writeResult(resultBothtextfile,dicBoth)
End Sub

Private Sub writeLog(dic1cnt,dic2cnt,dicBothcnt)
	Dim TF

	With CreateObject("Scripting.FileSystemObject")
		If .FileExists(resultLogfile) Then
			Set TF = .OpenTextFile(resultLogfile, 2)
		Else
			Set TF = .CreateTextFile(resultLogfile)
		End If

		TF.WriteLine("-----実行ログ-----")
		TF.WriteLine("実行時間:" & Now)
		TF.WriteLine("フォルダ1ファイル数:" & dic1cnt)
		TF.WriteLine("フォルダ2ファイル数:" & dic2cnt)
		TF.WriteLine("フォルダ両方ファイル数:" & dicBothcnt)
		TF.Close
	End With
End Sub

Private Sub writeResult(filename,dic)
	Dim TF

	With CreateObject("Scripting.FileSystemObject")
		If .FileExists(filename) Then
			Set TF = .OpenTextFile(filename, 2)
		Else
			Set TF = .CreateTextFile(filename)
		End If

		TF.WriteLine(Join(dic.keys,vbCrLf))
		TF.Close
	End With
End Sub

Private Sub checkParam()
	Dim dir1,dir2

	dir1 = GFs1.getDirectory
	dir2 = GFs2.getDirectory

	If dir1 = dir2 Then
		If dir1 = executeDirectory Then
			Err.Description = "検索ディレクトリがカレントディレクトリに設定されています"
			Err.Number = 1000
			Err.raise 1000
		Else
			Err.Description = "検索ディレクトリが同じです"
			Err.Number = 1000
			Err.raise 1000
		End If
	End IF

	If dir1 = executeDirectory Or GFs2.getDirectory = executeDirectory Then
		If msgbox("検索ディレクトリがカレントディレクトリに設定されています。実行しますか",vbYesNO) = vbNo Then
			Err.Description = "検索ディレクトリがカレントディレクトリに設定されています"
			Err.Number = 1000
			Err.raise 1000
		End If
	End If
End Sub

'call execTest(dic1,dic2,dicBoth)
Private Sub execTest(d1,d2,d3)
	msgbox Join(d1.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d1.Items,vbCrLf)
	msgbox Join(d2.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d2.Items,vbCrLf)
	msgbox Join(d3.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d3.Items,vbCrLf)
End Sub

Private Sub execTestParam()
	msgbox GFs1.getDirectory
	msgbox GFs2.getDirectory
End Sub


End Class