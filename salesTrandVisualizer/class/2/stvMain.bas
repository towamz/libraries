Sub stvMain()
    Dim STV As New SalesTrandVisualizer

    Set STV.WorksheetData = Worksheets("Sheet1")
    Set STV.HeaderRange = Worksheets("Sheet1").Range("A1:D1")

    STV.execution
End Sub
