Sub readTextFile(argSheetName As String, argInputFile As String)
    Dim buf_all As String
    
    With CreateObject("Scripting.FileSystemObject")
        With .GetFile(argInputFile).OpenAsTextStream
            buf_all = .ReadAll
            .Close
        End With
    End With

    Sheets(argSheetName).Activate
    Sheets(argSheetName).Range("A:A").Clear
    
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText buf_all
        .PutInClipboard
    End With
    
    Sheets(argSheetName).Paste Destination:=Sheets(argSheetName).Range("A1")
End Sub

