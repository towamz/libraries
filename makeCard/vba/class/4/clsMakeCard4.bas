Option Explicit

'グローバル変数(変更なし)
Private LNG_Columns_Number_In_A_Page As Long
Private ARYDBL_Columns_Length() As Double
Private LNG_Rows_Number_In_A_Page As Long
Private ARYDBL_Rows_Length() As Double
Private TBL_Tbl As Word.Table

'最初のセルが回転しないので最後に回転させる
Private objIlsFirstCellPicture As InlineShape

'定数(変更なし)
Const DBL_MM_TO_POINT As Double = 100 / 35.3
Const DBL_POINT_TO_MM As Double = 35.3 / 100
Const LNG_PICTURE_MARGIN As Long = 5
'画像を回転させると上下に空白が入るのでマージンを別指定する
'Const LNG_PICTURE_MARGIN2 As Long = 37
'画像を回転:shape変換して空白分を上に移動させると丁度の大きさになる
Const LNG_PICTURE_OFFSET As Long = -21 '52 * -0.4


'定数(設定)
'縦
' Const LNG_PAGE_WIDTH As Long = 210
' Const LNG_PAGE_HEIGHT As Long = 297
'横
Const LNG_PAGE_WIDTH As Long = 297
Const LNG_PAGE_HEIGHT As Long = 210

'列設定
Const ARYSTR_COLUMNS_LENGTH As String = "91"

'行設定
Const ARYSTR_ROWS_LENGTH As String = "52"

'オブジェクト変数
Private OBJ_Fn1 As clsGetFilename
Private OBJ_ILH As clsInfiniteLoopHandler


Public Sub makeCard(num As Integer)
    Dim cnt As Integer
    
    showColumnsRowsNumber
    clearContents
    setPage
    
    For cnt = 1 To num Step 1
        '無限ループハンドラー呼び出し
        OBJ_ILH.infiniteLoopHandler
        
        
        addTable
        setColumns
        setRows
        setCells
        'setCellsReverse

        
        Selection.EndKey Unit:=wdStory
        Selection.InsertBreak Type:=wdPageBreak
        
    
    Next
End Sub

Private Sub Class_Terminate()
    'ファイル名参照オブジェクトを破棄
    Set OBJ_Fn1 = Nothing
    '無限ループハンドラーオブジェクトを破棄
    Set OBJ_ILH = Nothing

End Sub


Private Sub Class_Initialize()
    Dim cellWidth As Double
    Dim cellHeight As Double
        
    
    'ファイル名参照オブジェクトのインスタンス化
    Set OBJ_Fn1 = New clsGetFilename
    '無限ループハンドラーオブジェクトのインスタンス化
    Set OBJ_ILH = New clsInfiniteLoopHandler
    
    
    '1カードあたりの行列の数とその長さの文字列をdouble型配列へ格納する
    ARYDBL_Rows_Length = getAryFromAryStr(ARYSTR_ROWS_LENGTH, cellHeight)
    ARYDBL_Columns_Length = getAryFromAryStr(ARYSTR_COLUMNS_LENGTH, cellWidth)
    
    LNG_Columns_Number_In_A_Page = LNG_PAGE_WIDTH / cellWidth
    LNG_Rows_Number_In_A_Page = LNG_PAGE_HEIGHT / cellHeight
    
    If LNG_Columns_Number_In_A_Page = 0 Then
        
        err.Raise 1000, , "1セルの幅・高さがページ範囲内になるように設定してください"
    
    ElseIf LNG_Rows_Number_In_A_Page = 0 Then
        err.Raise 1000, , "1セルの幅・高さがページ範囲内になるように設定してください"
    
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



Private Sub insertPicture(argI, argJ, argK, argL, argFilename)
    Dim objIls As InlineShape
    Dim objS As Shape
    Dim aspectRatioCurrent As Double
    Dim aspectRatioSetting As Double
    
    Set objIls = ActiveDocument.InlineShapes.AddPicture( _
        FileName:=argFilename, _
        Range:=TBL_Tbl.Cell(argI + argK, argJ + argL).Range)
    
    If objIlsFirstCellPicture Is Nothing Then
        Set objIlsFirstCellPicture = objIls
    End If
    
    
    objIls.LockAspectRatio = msoTrue
    
    'objIls.Width = 50

    'inlineShapeのままだと回転できないので一旦Shapeに変更する。
    '横幅が広い画像でも90に回転して挿入される画像があるので、rotationを0に設定する必要がある
    '画像サイズ変更が次のブロックであるが、
    'それと統合するとifとelse両方にConvertToShape/ConvertToInlineShapeを書く必要があるので
    '回転と画像サイズ変更は別にした方がいい
    Set objS = objIls.ConvertToShape
        
    '幅が大きいときは0
    If objS.Height < objS.Width Then
        objS.rotation = 0
    '高さが大きいときは270(回転させる)
    Else
        objS.rotation = 270
    End If
    
    'Shapeのままだとcell範囲外になるのでinlineShapeに戻す
    Set objIls = objS.ConvertToInlineShape
    
    objIls.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    
    '回転させていないとき、セルの行幅=画像の高さ・セルの列幅=画像の幅
    '回転させているとき、セルの行幅=画像の幅・セルの列幅=画像の高さ
    '画像がセル内におさまらないときは、セルと画像のアスペクト比を比較し大きい方の画像の高さ・幅をセル幅(+マージン)に合わせる
    
    
    '幅の方が大きいとき
    If objIls.Height < objIls.Width Then
        'If objIls.Height * DBL_POINT_TO_MM > ARYDBL_Rows_Length(argK) Or objIls.Width * DBL_POINT_TO_MM > ARYDBL_Columns_Length(argL) Then
            'アスペクト比取得
            aspectRatioSetting = ARYDBL_Rows_Length(argK) / ARYDBL_Columns_Length(argL)
            aspectRatioCurrent = objIls.Height / objIls.Width
             
            If aspectRatioSetting > aspectRatioCurrent Then
                objIls.Width = (ARYDBL_Columns_Length(argL) * DBL_MM_TO_POINT) - LNG_PICTURE_MARGIN
            Else
                objIls.Height = (ARYDBL_Rows_Length(argK) * DBL_MM_TO_POINT) - LNG_PICTURE_MARGIN
            End If
        'End If
    '高さの方が大きいとき
    Else
        'If objIls.Height * DBL_POINT_TO_MM > ARYDBL_Columns_Length(argL) Or objIls.Width * DBL_POINT_TO_MM > ARYDBL_Rows_Length(argK) Then
            'アスペクト比取得
            aspectRatioSetting = ARYDBL_Rows_Length(argK) / ARYDBL_Columns_Length(argL)
            aspectRatioCurrent = objIls.Width / objIls.Height
            
            If aspectRatioSetting > aspectRatioCurrent Then
                objIls.Height = (ARYDBL_Columns_Length(argL) * DBL_MM_TO_POINT) - LNG_PICTURE_MARGIN
            Else
                objIls.Width = (ARYDBL_Rows_Length(argK) * DBL_MM_TO_POINT) - LNG_PICTURE_MARGIN
            End If
        'End If
        Set objS = objIls.ConvertToShape
        objS.Top = LNG_PICTURE_OFFSET
    End If
            
End Sub

'https://www.msofficeforums.com/word-vba/11055-word-vba-add-textboxs-table-cells.html

Private Sub insertTextbox(argI, argJ, argK, argL, argText)
    Dim lLeft As Long
    Dim Shp As Word.Shape
    Dim Rng As Word.Range
    

    Set Rng = TBL_Tbl.Cell(argI + argK, argJ + argL).Range
    Rng.Collapse wdCollapseStart
    
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    Set Shp = ActiveDocument.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=10 * DBL_MM_TO_POINT, Top:=35 * DBL_MM_TO_POINT, Width:=80 * DBL_MM_TO_POINT, Height:=15 * DBL_MM_TO_POINT, Anchor:=Rng)
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents
    DoEvents


    With Shp
      .RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
      .RelativeVerticalPosition = wdRelativeVerticalPositionPage
      .Line.Visible = msoFalse
      .TextFrame.TextRange.Text = argText
      .TextFrame.TextRange.Font.ColorIndex = wdBrightGreen
      .TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
    End With

End Sub



Public Sub showColumnsRowsNumber()

    MsgBox "columns:" & LNG_Columns_Number_In_A_Page & vbCrLf & _
         "Rows:" & LNG_Rows_Number_In_A_Page
End Sub

Private Sub clearContents()
    ActiveDocument.Content.Delete
End Sub

Private Sub setPage()
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
   

End Sub

Private Sub addTable()
    
    'テーブルを追加
    Set TBL_Tbl = ActiveDocument.Tables.Add( _
        Range:=Selection.Range, _
        NumRows:=(UBound(ARYDBL_Rows_Length) + 1) * LNG_Rows_Number_In_A_Page, _
        NumColumns:=(UBound(ARYDBL_Columns_Length) + 1) * LNG_Columns_Number_In_A_Page, _
        DefaultTableBehavior:=wdWord8TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed)
    
    '罫線なしに設定
    With TBL_Tbl
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
    For j = 1 To TBL_Tbl.Columns.Count Step (UBound(ARYDBL_Columns_Length) + 1)
        For l = 0 To UBound(ARYDBL_Columns_Length) Step 1
            '配列に設定した列幅を設定する
            TBL_Tbl.Columns(j + l).Width = ARYDBL_Columns_Length(l) * DBL_MM_TO_POINT

        Next
            
        '罫線を引く(切り取り線)
        '左罫線
        'TBL_Tbl.Columns(j).Borders(wdBorderLeft).LineStyle = wdLineStyleDashDot
    
        '右罫線
        'TBL_Tbl.Columns(j + UBound(ARYDBL_Columns_Length)).Borders(wdBorderRight).LineStyle = wdLineStyleDashDot

    Next

End Sub


Public Sub setRows()
    Dim i As Long
    Dim k As Long
                        
    '行設定
    For i = 1 To TBL_Tbl.Rows.Count Step (UBound(ARYDBL_Rows_Length) + 1)
        For k = 0 To UBound(ARYDBL_Rows_Length) Step 1
            '配列に設定した行高を設定する
            TBL_Tbl.Rows(i + k).Height = ARYDBL_Rows_Length(k) * DBL_MM_TO_POINT
        Next
            
        '罫線を引く(切り取り線)
        '上罫線
        'TBL_Tbl.Rows(i).Borders(wdBorderTop).LineStyle = wdLineStyleDashDot
                    
        '下罫線
        'TBL_Tbl.Rows(i + UBound(ARYDBL_Rows_Length)).Borders(wdBorderBottom).LineStyle = wdLineStyleDashDot
    
    Next




End Sub



Public Sub setCells()
    Dim i, j As Long
    Dim k, l As Long
    Dim varInfo As Variant
    Dim objS As Shape
    
    '最初のセルの写真のオブジェクトを破棄する
    Set objIlsFirstCellPicture = Nothing

    For i = 1 To TBL_Tbl.Rows.Count Step (UBound(ARYDBL_Rows_Length) + 1)
        For j = 1 To TBL_Tbl.Columns.Count Step (UBound(ARYDBL_Columns_Length) + 1)

            varInfo = OBJ_Fn1.getPictureInfo

            Call insertPicture(i, j, 0, 0, varInfo(0))
            
        Next
    Next


    Set objS = objIlsFirstCellPicture.ConvertToShape
    
    '回転が必要な時だけ実施する
    If objS.Height > objS.Width Then
        objS.rotation = 270
        objS.Top = LNG_PICTURE_OFFSET
    End If

End Sub

