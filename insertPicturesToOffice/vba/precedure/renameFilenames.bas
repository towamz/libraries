Option Explicit

Sub renameFilenames()
    Dim FSO
    Dim baseDir As String
    Dim origFilename, destFilename As String
    Dim firstRng As Range
    Dim i As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    baseDir = Worksheets("設定").Range("B6").Value
    Set firstRng = Worksheets("ファイル名").Range("A1")


    i = 0
    Do Until firstRng.Offset(i, 0).Value = ""
        If firstRng.Offset(i, 1).Value = "" Then
            firstRng.Offset(i, 2).Value = "変更後の名前が設定されていません"
        ElseIf firstRng.Offset(i, 0).Value = firstRng.Offset(i, 1).Value Then
            firstRng.Offset(i, 2).Value = "変更前後の名前が同じです"
        Else
            origFilename = FSO.BuildPath(baseDir, firstRng.Offset(i, 0).Value)
            destFilename = FSO.BuildPath(baseDir, firstRng.Offset(i, 1).Value)
        
            On Error Resume Next
            Name origFilename As destFilename
            firstRng.Offset(i, 2).Value = Err.Description
            On Error GoTo 0
    
        End If
    
        i = i + 1
    Loop

End Sub
