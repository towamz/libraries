Option Explicit

Class clsGetFilenamesByFSO
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




	Public Function getFilenamesByFSO()
		Dim objWshShell, objRE, objFS, objFolder, objFiles
		Dim file
	    Dim aryFileName() 
	    Dim FileName 
	    Dim cnt
	    
	    If STR_Directory="" Or STR_Pattern=""   then
	    	Err.Raise 1000
	    	Exit Function
	    End If
	    
	    
	    'msgbox STR_Directory & STR_Pattern




		Set objWshShell = WScript.CreateObject("WScript.Shell")
		Set objFS = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFS.GetFolder(STR_Directory)
		Set objFiles = objFolder.Files
		
		Set objRE = CreateObject("VBScript.RegExp")
		objRE.Pattern = STR_Pattern
		
		
	    cnt = 0

		For Each file in objFiles
			If objRE.Test(file.name) then
		        ReDim Preserve aryFileName(cnt)
	        
		        aryFileName(cnt) = file
		        
		        cnt = cnt + 1		
				
			
			End If 
		Next

	    getFilenamesByFSO = aryFileName

	End Function


End Class









