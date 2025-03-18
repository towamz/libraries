Option Explicit

Private FSO As Scripting.FileSystemObject

Private Sub Class_Initialize()
    Set FSO = New Scripting.FileSystemObject
End Sub


Public Function isFolderExist(argPath As String) As Boolean

    If FSO.FolderExists(argPath) Then
        isFolderExist = True
    Else
        isFolderExist = False
    End If

End Function


Public Sub makeFolder(argPath As String)
        FSO.CreateFolder (argPath)
End Sub


Public Function getPictureCount(argPath As String) As Long
        getPictureCount = FSO.GetFolder(argPath).Files.Count
End Function


Public Function getPictureArray(argPath As String) As Variant
    Dim objFile As Object
    Dim strFiles As String
    
    For Each objFile In FSO.GetFolder(argPath).Files
        strFiles = strFiles & objFile.Path & ","
    Next
    
    '最後のカンマを削除する
    strFiles = Left(strFiles, Len(strFiles) - 1)
    
    getPictureArray = Split(strFiles, ",")
    
End Function

