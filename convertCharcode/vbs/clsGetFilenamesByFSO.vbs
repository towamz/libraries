Option Explicit

Class clsGetFilenamesByDir
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



	Public Function getFilenamesByDir()
		Dim objWshShell, objRE, objFS, objFolder, objFiles
		Dim file
	    Dim aryFileName() 
	    Dim FileName 
	    Dim cnt
	    
	    If STR_Directory="" Or STR_Pattern=""   then
	    	Err.Raise 1000
	    	Exit Function
	    End If
	    
	    
	    msgbox STR_Directory & STR_Pattern




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

	    getFilenamesByDir = aryFileName

	End Function


End Class











