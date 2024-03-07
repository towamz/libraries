Option Explicit

Class clsGetFilenames


Private STR_Directory 
Private STR_Pattern 

'検索ディレクトリを設定する
Public Property Let setDirectory(argDirectory)
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

Public Function getFilenamesArray()
	Dim objRE, objFiles, objFile
	Dim aryFileName() 
	Dim cnt
	
	If STR_Directory="" Or STR_Pattern=""   then
		Err.Raise 1000
		Exit Function
	End If

	With CreateObject("Scripting.FileSystemObject")
		Set objFiles = .GetFolder(STR_Directory).Files
	End With

	Set objRE = CreateObject("VBScript.RegExp")
	objRE.Pattern = STR_Pattern

	cnt = 0

	For Each objFile in objFiles
		If objRE.Test(objFile.name) then
			ReDim Preserve aryFileName(cnt)
			aryFileName(cnt) = objFile.name
			cnt = cnt + 1		
		End If 
	Next

	getFilenamesArray = aryFileName

End Function

Public Function getFilenamesDictionary()
	Dim objRE, objFiles, objFile
	Dim dicFileNames

	Set dicFileNames = CreateObject("Scripting.Dictionary")

	With CreateObject("Scripting.FileSystemObject")
		Set objFiles = .GetFolder(STR_Directory).Files
	End With

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
	
	If STR_Directory="" Or STR_Pattern=""   then
		Err.Raise 1000
		Exit Function
	End If

	With CreateObject("Scripting.FileSystemObject")
		Set objFiles = .GetFolder(STR_Directory).Files
	End With

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
