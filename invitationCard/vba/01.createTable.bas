Option Explicit

Sub setTable()
    Dim i, j As Long
    
    'ページ設定
    With ActiveDocument.PageSetup
        '余白をゼロに設定
        .TopMargin = MillimetersToPoints(0)
        .BottomMargin = MillimetersToPoints(0)
        .LeftMargin = MillimetersToPoints(0)
        .RightMargin = MillimetersToPoints(0)
        
        '用紙サイズ指定(A4=210*297)
        .PageWidth = MillimetersToPoints(210)
        .PageHeight = MillimetersToPoints(297)

    End With
    
    'テーブルを追加
    ActiveDocument.Tables.Add _
        Range:=ActiveDocument.Range(0, 0), _
        NumRows:=12, NumColumns:=6, _
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


    '列設定
    For j = 1 To ActiveDocument.Tables(1).Columns.Count Step 2
    
        '左側は50mm
        ActiveDocument.Tables(1).Columns(j).Width = 50 * DBL_POINT_TO_MM
        
        '一番左は罫線を引かない
        If j > 1 Then
            '罫線を引く(切り取り線)
            ActiveDocument.Tables(1).Columns(j).Borders(wdBorderLeft).LineStyle = wdLineStyleDashDot
        End If
        
        '右側は20mm
        ActiveDocument.Tables(1).Columns(j + 1).Width = 20 * DBL_POINT_TO_MM


    Next

                        
    '行設定
    For i = 1 To ActiveDocument.Tables(1).Rows.Count Step 2
        '上側は30mm
        ActiveDocument.Tables(1).Rows(i).Height = 30 * DBL_POINT_TO_MM
        '下側は17mm、罫線を引く(切り取り線)
        ActiveDocument.Tables(1).Rows(i + 1).Height = 18 * DBL_POINT_TO_MM
        ActiveDocument.Tables(1).Rows(i + 1).Borders(wdBorderBottom).LineStyle = wdLineStyleDashDot

    
    Next

End Sub
