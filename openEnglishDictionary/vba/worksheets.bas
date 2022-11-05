Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If ActiveCell.Column = 2 Then
        Call openEnglishDictionary
    End If
End Sub