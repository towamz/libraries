'http://wsh.style-mods.net/ref_filesystemobject/index.htm
Option Explicit

Sub getFilenameParts()
    Dim fullfilename
    Dim FL, FSO
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FL = FSO.CreateTextFile(FSO.GetTempName())

    fullfilename = "C:\test\script\testname.txt"

    Debug.Print FSO.GetDriveName(fullfilename)
    Debug.Print FSO.GetParentFolderName(fullfilename)
    Debug.Print FSO.GetFileName(fullfilename)
    Debug.Print FSO.GetBaseName(fullfilename)
    Debug.Print FSO.GetExtensionName(fullfilename)

    Debug.Print FSO.GetAbsolutePathName("absPath.txtx")
    Debug.Print FSO.GetTempName()


    FL.WriteLine (fullfilename)
    FL.WriteLine ("")
    FL.WriteLine (FSO.GetDriveName(fullfilename))
    FL.WriteLine (FSO.GetParentFolderName(fullfilename))
    FL.WriteLine (FSO.GetFileName(fullfilename))
    FL.WriteLine (FSO.GetBaseName(fullfilename))
    FL.WriteLine (FSO.GetExtensionName(fullfilename))
    FL.WriteLine (FSO.GetAbsolutePathName("absPath.txtx"))
    FL.WriteLine (FSO.GetTempName())
    FL.Close

End Sub

'-----テキストファイル-----
'C:
'C:\test\script
'testname.txt
'testname
'txt
'C:\Users\forwa\OneDrive\デスクトップ\getFilenamePart\absPath.txtx
'rad8E5B1.tmp