Option Explicit

Private FSO As Scripting.FileSystemObject

Private Sub Class_Initialize()
    Set FSO = New Scripting.FileSystemObject
End Sub

Public Function isFileExist(argFile As String) As Boolean
    isFileExist = FSO.FileExists(argFile)
End Function

Public Function getFileSize(argFile As String) As Long
    getFileSize = FSO.GetFile(argFile).Size
End Function

Public Function isFolderExist(argPath As String) As Boolean
    isFolderExist = FSO.FolderExists(argPath)
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
    
    If getPictureCount(argPath) = 0 Then
        err.Raise 9999, , "画像が1枚もありません"
    End If
    
    For Each objFile In FSO.GetFolder(argPath).Files
        strFiles = strFiles & objFile.Path & ","
    Next
    
    '最後のカンマを削除する
    strFiles = Left(strFiles, Len(strFiles) - 1)
    
    getPictureArray = Split(strFiles, ",")
    
End Function

