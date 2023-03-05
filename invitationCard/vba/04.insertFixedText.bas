Option Explicit

Sub insertFixedText(ByRef argText As String)

    Dim i, j, cnt As Long

    cnt = 0

    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step LNG_ROWS_NUMBER_IN_A_CELL
        For j = 1 To ActiveDocument.Tables(1).Columns.Count Step LNG_COLUMNS_NUMBER_IN_A_CELL

            ActiveDocument.Tables(1).Cell(i + 1, j).Range.Orientation = wdTextOrientationHorizontal
            ActiveDocument.Tables(1).Cell(i + 1, j).VerticalAlignment = wdCellAlignVerticalTop

            ActiveDocument.Tables(1).Cell(i + 1, j).Range.Text = argText
            
            cnt = cnt + 1
        Next
    Next

End Sub








