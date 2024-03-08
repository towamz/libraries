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
	Dim objRE, objFiles, objFile
	Dim aryFileName() 
	Dim cnt
	
	Set objFiles = getFilesObj()

	Set objRE = CreateObject("VBScript.RegExp")
	objRE.Pattern = STR_Pattern

	cnt = 0
	ReDim Preserve aryFileName(objFiles.count)
	For Each objFile in objFiles
		If objRE.Test(objFile.name) then
			aryFileName(cnt) = objFile.name
			cnt = cnt + 1		
		End If 
	Next

	If cnt = 0 then
		getFilenamesArray = ""
	else
		ReDim Preserve aryFileName(cnt - 1)
		getFilenamesArray = aryFileName
	End If
End Function

Public Function getFilenamesDictionary()
	Dim objRE, objFiles, objFile
	Dim dicFileNames

	Set objFiles = getFilesObj()

	Set dicFileNames = CreateObject("Scripting.Dictionary")

	Set objRE = CreateObject("VBScript.RegExp")
	objRE.Pattern = STR_Pattern

	For Each objFile in objFiles
		If objRE.Test(objFile.name) then
			dicFileNames.Add objFile.name,0
		End If 
	Next

	set getFilenamesDictionary = dicFileNames

End Function

Public Function getFirstMatchFilename()
	Dim objRE, objFiles, objFile
	
	Set objFiles = getFilesObj()

	Set objRE = CreateObject("VBScript.RegExp")
	objRE.Pattern = STR_Pattern

	For Each objFile in objFiles
		If objRE.Test(objFile.name) then
			'マッチした最初のファイルを返す
			getFirstMatchFilename = objFile.name
			Exit Function
		End If 
	Next

	'マッチするファイルが見つからなかったので、空白を返す
	getFirstMatchFilename = ""
End Function


End Class
