Option Explicit

Const baseDirectory As String = "C:\anki\"


Sub getMP3filename()
    Dim pattern As String
    Dim filename As String
    
    pattern = baseDirectory & "*" & Cells(ActiveCell.Row, 2).Value & "*.mp3"

    filename = Dir(pattern, vbNormal)

    If ActiveCell.Value <> "" Then
        'MsgBox "既に入力があります"
    ElseIf filename = "" Then
        'MsgBox "該当のファイルは見つかりませんでした"
    Else
        ActiveCell.Value = "[sound:" & filename & "]"
    End If

End Sub
