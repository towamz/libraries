Option Explicit

'Include("clsGetFilenames.vbs")
'Include("clsCompareFilesInFolders.vbs")

Class clsCompareFilesInFolders

private directory1
private directory2
private pattern
private result1textfile
private result2textfile
private resultBothtextfile



'検索ディレクトリを設定する
Public Property Let setDirectory1(argDirectory)
	With CreateObject("Scripting.FileSystemObject")
		If Not .FolderExists(argDirectory) Then
			Err.Raise 1000
		End If
	End With

	directory1 = argDirectory
End Property

Public Property Let setDirectory2(argDirectory)
	With CreateObject("Scripting.FileSystemObject")
		If Not .FolderExists(argDirectory) Then
			Err.Raise 1000
		End If
	End With

	directory2 = argDirectory
End Property

Public Property Let setPattern(argPattern)
	'不正な正規表現であればエラー発生 / an error occure if invalid
	With CreateObject("VBScript.RegExp")
		.Pattern = argPattern
		.Test("testExec")
	End With
	pattern = argPattern
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

Private Sub Class_Initialize()
	'既定値はすべてのファイル / defalut is all files
	pattern = ".*"
End Sub

Public Sub compareFilesInFolders()
	Dim GFs1,GFs2
	Dim ary2
	Dim dic1,dic2,dicBoth
	Dim filename

	Set GFs1 = New clsGetFilenames
	Set GFs2 = New clsGetFilenames

	GFs1.setDirectory = directory1
	GFs2.setDirectory = directory2
	GFs1.setPattern = pattern
	GFs2.setPattern = pattern

	'フォルダ1のディクショナリを２つ生成する
	set dic1 = GFs1.getFilenamesDictionary()
	'フォルダ2の配列を1つ生成する
	ary2 = GFs2.getFilenamesArray()

	'ファイル一覧を取得したのでオブジェクトを開放する
	'release objects as file lists were made
	Set GFs1 = Nothing
	Set GFs2 = Nothing	

	'差分格納用ディクショナリ生成
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

	call writeResult(result1textfile,dic1)
	call writeResult(result2textfile,dic2)
	call writeResult(resultBothtextfile,dicBoth)
End Sub

Private Sub writeResult(filename,dic)
	Dim objFso
	Set objFso = Wscript.CreateObject("Scripting.FileSystemObject")

	With objFso.OpenTextFile(filename, 2)
		.WriteLine(Join(dic.keys,vbCrLf))
		.Close
	End With
End Sub

'call execTest(dic1,dic2,dicBoth)
Private Sub execTest(d1,d2,d3)
	msgbox Join(d1.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d1.Items,vbCrLf)
	msgbox Join(d2.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d2.Items,vbCrLf)
	msgbox Join(d3.keys,vbCrLf) & vbCrLf & vbCrLf & Join(d3.Items,vbCrLf)
End Sub


End Class