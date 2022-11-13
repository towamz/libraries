
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Select Case ActiveCell.Column
        Case 2
            Call getMP3BySelenium
            Call openEnglishDictionary
            
        'Case 8
            'Call getMP3filename
            
    End Select
    
End Sub
