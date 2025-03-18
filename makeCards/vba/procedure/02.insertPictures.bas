Option Explicit

Sub insertPictures(ByRef argFilenameArray() As String)

    Dim i, j, cnt As Long
    Dim objIls As InlineShape
    Dim objS As Shape

    cnt = 0

    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step 2
        For j = 1 To ActiveDocument.Tables(1).Columns.Count Step 2

            Set objIls = ActiveDocument.InlineShapes.AddPicture( _
                FileName:=argFilenameArray(cnt), _
                Range:=ActiveDocument.Tables(1).Cell(i, j).Range)

            objIls.LockAspectRatio = msoTrue
            
            Debug.Print i & "," & j
            Debug.Print Int(objIls.Height * DBL_POINT_TO_MM) & "," & Int(objIls.Width * DBL_POINT_TO_MM)
            
             
            'Stop
            
            If Int(objIls.Height * DBL_POINT_TO_MM) > (LNG_ROW_ONE - LNG_PICTURE_CONVERT_MARGIN) Then
                objIls.Height = (LNG_ROW_ONE - LNG_PICTURE_CONVERT_MARGIN) * DBL_MM_TO_POINT
            End If
            
            If Int(objIls.Width * DBL_POINT_TO_MM) > (LNG_COLUMN_ONE - LNG_PICTURE_CONVERT_MARGIN) Then
                objIls.Width = (LNG_COLUMN_ONE - LNG_PICTURE_CONVERT_MARGIN) * DBL_MM_TO_POINT
            End If
            
            'inlineShapeのままだと回転できないので一旦Shapeに変更する。Shapeのままだとcell範囲外になるのでinlineShapeに戻す
            Set objS = objIls.ConvertToShape
            objS.rotation = 0
            objS.ConvertToInlineShape
            
            Debug.Print Int(objIls.Height * DBL_POINT_TO_MM) & "," & Int(objIls.Width * DBL_POINT_TO_MM)
            'Stop
            
            If Int(objIls.Height * DBL_POINT_TO_MM) > (LNG_ROW_ONE - LNG_PICTURE_CONVERT_MARGIN) Or Int(objIls.Width * DBL_POINT_TO_MM) > (LNG_COLUMN_ONE - LNG_PICTURE_CONVERT_MARGIN) Then
                Stop
            End If

            
            
            cnt = cnt + 1
        Next
    Next


End Sub



