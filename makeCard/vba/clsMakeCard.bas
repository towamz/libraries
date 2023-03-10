Option Explicit

Const DBL_MM_TO_POINT As Double = 100 / 35.3
Const DBL_POINT_TO_MM As Double = 35.3 / 100
Const LNG_PICTURE_MARGIN As Long = 6

'縦
' Const LNG_PAGE_WIDTH As Long = 210
' Const LNG_PAGE_HEIGHT As Long = 297
'横
Const LNG_PAGE_WIDTH As Long = 297
Const LNG_PAGE_HEIGHT As Long = 210

'列設定
Const LNG_COLUMNS_NUMBER_IN_A_PAGE As Long = 3
Const ARYSTR_COLUMNS_LENGTH As String = "45,46"
Private ARYDBL_Columns_Length() As Double

'行設定
Const LNG_ROWS_NUMBER_IN_A_PAGE As Long = 4
Const ARYSTR_ROWS_LENGTH As String = "12,10,10,10,10"
Private ARYDBL_Rows_Length() As Double

Private Sub Class_Initialize()
    Dim cellWidth As Double
    Dim cellHeight As Double
    
    '1カードあたりの行列の数とその長さの文字列をdouble型配列へ格納する
    ARYDBL_Rows_Length = getAryFromAryStr(ARYSTR_ROWS_LENGTH, cellHeight)
    ARYDBL_Columns_Length = getAryFromAryStr(ARYSTR_COLUMNS_LENGTH, cellWidth)
    
    'セルの長さ*セルの個数がページの長さを超えていないかチェック
    If cellHeight * LNG_ROWS_NUMBER_IN_A_PAGE > LNG_PAGE_HEIGHT Then
        Debug.Print cellHeight & "*" & LNG_ROWS_NUMBER_IN_A_PAGE & "=" & cellHeight * LNG_ROWS_NUMBER_IN_A_PAGE & "<-->" & LNG_PAGE_HEIGHT
        Err.Raise 1000, , "ページ範囲内になるようにセル幅、セル数を指定してください"
    ElseIf cellWidth * LNG_COLUMNS_NUMBER_IN_A_PAGE > LNG_PAGE_WIDTH Then
        Debug.Print cellWidth & "*" & LNG_COLUMNS_NUMBER_IN_A_PAGE & "=" & cellWidth * LNG_COLUMNS_NUMBER_IN_A_PAGE & "<-->" & LNG_PAGE_WIDTH
        Err.Raise 1000, , "ページ範囲内になるようにセル幅、セル数を指定してください"
    End If
    
    
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
        NumRows:=(UBound(ARYDBL_Rows_Length) + 1) * LNG_ROWS_NUMBER_IN_A_PAGE, _
        NumColumns:=(UBound(ARYDBL_Columns_Length) + 1) * LNG_COLUMNS_NUMBER_IN_A_PAGE, _
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
    Dim i, j As Long

    '列設定
    For i = 1 To ActiveDocument.Tables(1).Columns.Count Step (UBound(ARYDBL_Columns_Length) + 1)
        For j = 0 To UBound(ARYDBL_Columns_Length) Step 1
            '配列に設定した列幅を設定する
            ActiveDocument.Tables(1).Columns(i + j).Width = ARYDBL_Columns_Length(j) * DBL_MM_TO_POINT
    
            '罫線を引く(切り取り線)
            '左罫線
'            If j = 0 Then
'                ActiveDocument.Tables(1).Columns(i + j).Borders(wdBorderLeft).LineStyle = wdLineStyleDashDot
'            End If
        
            '右罫線
            If j = UBound(ARYDBL_Columns_Length) Then
                ActiveDocument.Tables(1).Columns(i + j).Borders(wdBorderRight).LineStyle = wdLineStyleDashDot
            
            End If


        Next

    Next

End Sub


Public Sub setRows()
    Dim i, j As Long

                        
    '行設定
    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step (UBound(ARYDBL_Rows_Length) + 1)
        For j = 0 To UBound(ARYDBL_Rows_Length) Step 1
            '配列に設定した行高を設定する
            ActiveDocument.Tables(1).Rows(i + j).Height = ARYDBL_Rows_Length(j) * DBL_MM_TO_POINT
            
            '罫線を引く(切り取り線)
            '上罫線
'            If j = 0 Then
'                ActiveDocument.Tables(1).Rows(i + j).Borders(wdBorderTop).LineStyle = wdLineStyleDashDot
'
'            End If
                    
            '下罫線
            If j = UBound(ARYDBL_Rows_Length) Then
                ActiveDocument.Tables(1).Rows(i + j).Borders(wdBorderBottom).LineStyle = wdLineStyleDashDot
            
            End If
        Next
    Next




End Sub



Public Sub setCells()


End Sub








