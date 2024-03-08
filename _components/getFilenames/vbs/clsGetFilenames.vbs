Option Explicit

Class clsGetFilenames


Private STR_Directory 
Private STR_Pattern 

'検索ディレクトリを設定する
Public Property Let setDirectory(argDirectory)
	With CreateObject("Scripting.FileSystemObject")
		If Not .FolderExists(argDirectory) Then
			Err.Raise 1000
		End If
	End With

	STR_Directory = argDirectory
End Property

'パターンを設定する
Public Property Let setPattern(argPattern)
	'不正な正規表現であればエラー発生 / an error occure if invalid
	With CreateObject("VBScript.RegExp")
		.Pattern = argPattern
		.Test("testExec")
	End With
	STR_Pattern = argPattern
End Property

'設定されている検索ディレクトリを返す
Public Property Get getDirectory()
	getDirectory = STR_Directory
End Property

'設定されている検索ディレクトリの文字列長を返す
Public Property Get getDirectoryLen()
	getDirectoryLen = Len(STR_Directory)
End Property

'設定されている検索ディレクトリを返す
Public Property Get getPattern()
	getPattern = STR_Pattern
End Property

Private Sub Class_Initialize()
	'既定値はカレントディレクトリ / defalut is the current directory
	With CreateObject("Scripting.FileSystemObject")
		STR_Directory = .getParentFolderName(WScript.ScriptFullName)
	End With

	'既定値はすべてのファイル / defalut is all files
	STR_Pattern = ".*"
End Sub


Public Function getFilesObj()
	With CreateObject("Scripting.FileSystemObject")
		Set getFilesObj = .GetFolder(STR_Directory).Files
	End With
End Function

Public Function getFilenamesArray()
	Dim objFiles, objFile
	Dim aryFileName() 
	Dim cnt
	
	Set objFiles = getFilesObj()


	cnt = 0
	'配列要素数をファイル数に設定する / set array index to fils count (possible max number) 
	ReDim Preserve aryFileName(objFiles.count)
	
	With CreateObject("VBScript.RegExp")
		.Pattern = STR_Pattern

		For Each objFile in objFiles
			If .Test(objFile.name) then
				aryFileName(cnt) = objFile.name
				cnt = cnt + 1		
			End If 
		Next
	End With

	If cnt = 0 then
		getFilenamesArray = ""
	else
		'配列要素数を再設定する(正規表現に一致しないファイルがある可能性がある) / 
		'reset array index (there may be files unmatch the regular expression)
		ReDim Preserve aryFileName(cnt - 1)
		getFilenamesArray = aryFileName
	End If
End Function

Public Function getFilenamesDictionary()
	Dim objFiles, objFile
	Dim dicFileNames

	Set objFiles = getFilesObj()

	Set dicFileNames = CreateObject("Scripting.Dictionary")

	With CreateObject("VBScript.RegExp")
		.Pattern = STR_Pattern
		For Each objFile in objFiles
			If .Test(objFile.name) then
				dicFileNames.Add objFile.name,0
			End If 
		Next
	End With

	set getFilenamesDictionary = dicFileNames

End Function

Public Function getFirstMatchFilename()
	Dim objFiles, objFile
	
	Set objFiles = getFilesObj()

	With CreateObject("VBScript.RegExp")
		.Pattern = STR_Pattern
		For Each objFile in objFiles
			If .Test(objFile.name) then
				'マッチした最初のファイルを返す
				getFirstMatchFilename = objFile.name
				Exit Function
			End If 
		Next
	End With

	'マッチするファイルが見つからなかったので、空白を返す
	getFirstMatchFilename = ""
End Function


End Class
