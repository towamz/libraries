Option Explicit

Sub insertText(ByRef argTextArray() As String)

    Dim i, j, cnt As Long
    Dim objIls As InlineShape
    Dim objS As Shape

    cnt = 0

    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step LNG_ROWS_NUMBER_IN_A_CELL
        For j = 1 To ActiveDocument.Tables(1).Columns.Count Step LNG_COLUMNS_NUMBER_IN_A_CELL
            ActiveDocument.Tables(1).Cell(i, j + 1).Range.Orientation = wdTextOrientationUpward
            ActiveDocument.Tables(1).Cell(i, j + 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft

            ActiveDocument.Tables(1).Cell(i, j + 1).Range.Text = argTextArray(cnt)
            
            cnt = cnt + 1
        Next
    Next

End Sub







