Option Explicit
Dim FSO

Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")


Function getDateString()
    Dim tmpStr

    tmpStr = Right("0" & Year(Now), 2) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) 
    getDateString = tmpStr
End Function

Function getFolderBase()
    getFolderBase = FSO.getParentFolderName(WScript.ScriptFullName)
End Function

Function getFoldername(folderBase)
    Dim i, tmpFoldername, tmpStr, tmpStr2

    tmpStr = getDateString()
    tmpStr = tmpStr & "-"

    i = 0
    Do 
        i = i + 1
        tmpStr2 = Right("0" & i, 2)

        tmpFoldername = folderBase & "\" & tmpStr & tmpStr2 

    Loop While FSO.FolderExists(tmpFoldername)

    getFoldername = tmpFoldername
End Function

Sub createFolders()
    Dim folderBase, subFoldername

    folderBase = getFolderBase()
    subFoldername = getFoldername(folderBase)

    If Not FSO.FolderExists(subFoldername) Then
        Call FSO.CreateFolder(subFoldername) 
    End If

End Sub

call createFolders()

WScript.Quit

