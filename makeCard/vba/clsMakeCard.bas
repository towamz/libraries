Option Explicit


'グローバル変数
Private LNG_Columns_Number_In_A_Page As Long
Private ARYDBL_Columns_Length() As Double
Private LNG_Rows_Number_In_A_Page As Long
Private ARYDBL_Rows_Length() As Double

Private ARYDBL_Cells_Length() As Double


'定数(変更なし)
Const DBL_MM_TO_POINT As Double = 100 / 35.3
Const DBL_POINT_TO_MM As Double = 35.3 / 100
Const LNG_PICTURE_MARGIN As Long = 6


'定数(設定)
'縦
' Const LNG_PAGE_WIDTH As Long = 210
' Const LNG_PAGE_HEIGHT As Long = 297
'横
Const LNG_PAGE_WIDTH As Long = 297
Const LNG_PAGE_HEIGHT As Long = 210

'列設定
Const ARYSTR_COLUMNS_LENGTH As String = "45,46"

'行設定
Const ARYSTR_ROWS_LENGTH As String = "12,10,10,10,10"


Private Sub Class_Initialize()
    Dim cellWidth As Double
    Dim cellHeight As Double
    
    Dim l As Long
    Dim k As Long
    
    '1カードあたりの行列の数とその長さの文字列をdouble型配列へ格納する
    ARYDBL_Rows_Length = getAryFromAryStr(ARYSTR_ROWS_LENGTH, cellHeight)
    ARYDBL_Columns_Length = getAryFromAryStr(ARYSTR_COLUMNS_LENGTH, cellWidth)
    
    LNG_Columns_Number_In_A_Page = LNG_PAGE_WIDTH / cellWidth
    LNG_Rows_Number_In_A_Page = LNG_PAGE_HEIGHT / cellHeight
    
    If LNG_Columns_Number_In_A_Page = 0 Then
        
        Err.Raise 1000, , "1セルの幅・高さがページ範囲内になるように設定してください"
    
    ElseIf LNG_Rows_Number_In_A_Page = 0 Then
        Err.Raise 1000, , "1セルの幅・高さがページ範囲内になるように設定してください"
    
    End If
    
    
    
    'セルの長さを保持する変数を初期化
    ReDim ARYDBL_Cells_Length(UBound(ARYDBL_Columns_Length), UBound(ARYDBL_Rows_Length), 1)
        
    For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
        For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
            '列幅を格納
            ARYDBL_Cells_Length(l, k, 0) = ARYDBL_Columns_Length(l)
        Next

    Next
            
    For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
        For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
            '行高を格納
            ARYDBL_Cells_Length(l, k, 1) = ARYDBL_Rows_Length(k)
        Next

    Next

       
End Sub


'配列文字列から配列を生成するとstr型になるので、dbl型に変換する
'1セルあたりの幅・高さの合計を参照渡しの変数で返す
Private Function getAryFromAryStr(ByVal argAryStr As String, ByRef argLength As Double) As Double()

    Dim i As Long
    Dim aryTmp As Variant
    Dim aryTmp2() As Double
    Dim cellWidth As Double
    Dim cellHeight As Double

    aryTmp = Split(argAryStr, ",")
    
    ReDim aryTmp2(UBound(aryTmp))

    For i = 0 To UBound(aryTmp)
                        
        aryTmp2(i) = CDbl(aryTmp(i))
        argLength = argLength + CDbl(aryTmp(i))
        
    Next

    getAryFromAryStr = aryTmp2


End Function

Public Sub showColumnsRowsNumber()

    MsgBox "columns:" & LNG_Columns_Number_In_A_Page & vbCrLf & _
         "Rows:" & LNG_Rows_Number_In_A_Page
End Sub


Private Sub insertPicture(argI, argJ, argK, argL, argFilename)
    Dim objIls As InlineShape
    Dim objS As Shape
    Dim cntMargin As Long
            
    Set objIls = ActiveDocument.InlineShapes.AddPicture( _
        FileName:=argFilename, _
        Range:=ActiveDocument.Tables(1).Cell(argI + argK, argJ + argL).Range)

    objIls.LockAspectRatio = msoTrue
    
    'inlineShapeのままだと回転できないので一旦Shapeに変更する。Shapeのままだとcell範囲外になるのでinlineShapeに戻す
    Set objS = objIls.ConvertToShape
    objS.rotation = 0
    objS.ConvertToInlineShape
    
    
    Debug.Print argI + argK & "," & argJ + argL
    Debug.Print Int(objIls.Height * DBL_POINT_TO_MM) & "," & Int(objIls.Width * DBL_POINT_TO_MM)
    
     
    'Stop
    
    '大きさの調整
    cntMargin = -1
    
    Do
        cntMargin = cntMargin + 1
    
        If Int(objIls.Height * DBL_POINT_TO_MM) > (ARYDBL_Rows_Length(argK) - LNG_PICTURE_MARGIN) Then
            objIls.Height = (ARYDBL_Rows_Length(argK) - LNG_PICTURE_MARGIN - cntMargin) * DBL_MM_TO_POINT
        End If
        
        If Int(objIls.Width * DBL_POINT_TO_MM) > (ARYDBL_Columns_Length(argL) - LNG_PICTURE_MARGIN) Then
            objIls.Width = (ARYDBL_Columns_Length(argL) - LNG_PICTURE_MARGIN - cntMargin) * DBL_MM_TO_POINT
        End If
        
        
        Debug.Print Int(objIls.Height * DBL_POINT_TO_MM) & "," & Int(objIls.Width * DBL_POINT_TO_MM)
        'Stop

    Loop While Int(objIls.Height * DBL_POINT_TO_MM) > (ARYDBL_Rows_Length(argK) - LNG_PICTURE_MARGIN) Or _
               Int(objIls.Width * DBL_POINT_TO_MM) > (ARYDBL_Columns_Length(argL) - LNG_PICTURE_MARGIN)
            
End Sub



Public Sub setPage()
    'ページ設定
    With ActiveDocument.PageSetup
        '余白をゼロに設定
        .TopMargin = MillimetersToPoints(0)
        .BottomMargin = MillimetersToPoints(0)
        .LeftMargin = MillimetersToPoints(0)
        .RightMargin = MillimetersToPoints(0)
        
        '用紙サイズ指定
        .PageWidth = MillimetersToPoints(LNG_PAGE_WIDTH)
        .PageHeight = MillimetersToPoints(LNG_PAGE_HEIGHT)

    End With
    
    'テーブルを追加
    ActiveDocument.Tables.Add _
        Range:=ActiveDocument.Range(0, 0), _
        NumRows:=(UBound(ARYDBL_Rows_Length) + 1) * LNG_Rows_Number_In_A_Page, _
        NumColumns:=(UBound(ARYDBL_Columns_Length) + 1) * LNG_Columns_Number_In_A_Page, _
        DefaultTableBehavior:=wdWord8TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed
    
    '罫線なしに設定
    With ActiveDocument.Tables(1)
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    End With
End Sub


Public Sub setColumns()
    Dim j As Long
    Dim l As Long
    '列設定
    For j = 1 To ActiveDocument.Tables(1).Columns.Count Step (UBound(ARYDBL_Columns_Length) + 1)
        For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
            '配列に設定した列幅を設定する
            ActiveDocument.Tables(1).Columns(j + l).Width = ARYDBL_Columns_Length(l) * DBL_MM_TO_POINT

        Next
            
        '罫線を引く(切り取り線)
        '左罫線
        'ActiveDocument.Tables(1).Columns(i).Borders(wdBorderLeft).LineStyle = wdLineStyleDashDot
    
        '右罫線
        ActiveDocument.Tables(1).Columns(j + UBound(ARYDBL_Columns_Length)).Borders(wdBorderRight).LineStyle = wdLineStyleDashDot

    Next

End Sub


Public Sub setRows()
    Dim i As Long
    Dim k As Long
                        
    '行設定
    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step (UBound(ARYDBL_Rows_Length) + 1)
        For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
            '配列に設定した行高を設定する
            ActiveDocument.Tables(1).Rows(i + k).Height = ARYDBL_Rows_Length(k) * DBL_MM_TO_POINT
        Next
            
        '罫線を引く(切り取り線)
        '上罫線
        'ActiveDocument.Tables(1).Rows(i).Borders(wdBorderTop).LineStyle = wdLineStyleDashDot
                    
        '下罫線
        ActiveDocument.Tables(1).Rows(i + UBound(ARYDBL_Rows_Length)).Borders(wdBorderBottom).LineStyle = wdLineStyleDashDot
    
    Next




End Sub



Public Sub setCells()
    Dim i, j As Long
    Dim k, l As Long

    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step (UBound(ARYDBL_Rows_Length) + 1)
        For j = 1 To ActiveDocument.Tables(1).Columns.Count Step (UBound(ARYDBL_Columns_Length) + 1)
            For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
                For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
                    'デバッグ用
                    'Debug.Print i & "+" & k & "," & j & "+" & l & "(", i + k & "," & j + l
                    'ActiveDocument.Tables(1).Cell(i + k, j + l).Select
                    'ActiveDocument.Tables(1).Cell(i + k, j + l).Range.Text = "(" & k & "," & l & ")"

                Next
            Next
        Next
    Next


End Sub

Public Sub setCellsReverse()
    Dim i, j As Long
    Dim k, l As Long


    For i = ActiveDocument.Tables(1).Rows.Count - UBound(ARYDBL_Rows_Length) To 1 Step -(UBound(ARYDBL_Rows_Length) + 1)
        For j = ActiveDocument.Tables(1).Columns.Count - UBound(ARYDBL_Columns_Length) To 1 Step -(UBound(ARYDBL_Columns_Length) + 1)
            For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
                For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
                    'デバッグ用
                    'Debug.Print i & "+" & k & "," & j & "+" & l & "(", i + k & "," & j + l
                    'ActiveDocument.Tables(1).Cell(i + k, j + l).Select
                    'ActiveDocument.Tables(1).Cell(i + k, j + l).Range.Text = "(" & k & "," & l & ")"

                Next
            Next

        Next
    Next


End Sub

