Sub stvMain()
    Dim STV As New SalesTrandVisualizer

    STV.WorksheetData = Worksheets("Sheet1")
    STV.HeaderRange = Worksheets("Sheet1").Range("A1:D1")
    STV.FirstDayRange = Worksheets("Sheet1").Range("A2")
    
    STV.execution

End Sub
