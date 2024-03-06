Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ed As New clsEnglishDictionary
    
    Select Case ActiveCell.column
        Case 2
            If Cells(ActiveCell.Row, 2).Value <> "" Then
                If Cells(ActiveCell.Row, 8).Value = "" Then
                    ed.setDriverOfPageBySelenium
                    ed.getMP3BySelenium
                    ed.getWordDefinitionBySelenium
                    
                End If
           End If
    End Select
    
    Set ed = Nothing
    
End Sub
