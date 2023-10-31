Option Explicit

Sub main()
    Dim objMC As clsMakeCard


    Set objMC = New clsMakeCard
    
    objMC.showColumnsRowsNumber
    objMC.setPage
    objMC.setColumns
    objMC.setRows

    objMC.setCells
    objMC.setCellsReverse

    Set objMC = Nothing

End Sub




