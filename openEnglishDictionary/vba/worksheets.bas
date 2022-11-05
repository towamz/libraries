Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
    Select Case ActiveCell.Column
        Case 2
            Call openDic
            
        Case 8
            Call getMP3filename
    End Select
    
End Sub