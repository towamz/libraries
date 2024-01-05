Option Explicit

Private OBJ_EXCEL As Excel.Application
Private WB_FILENAME_DATA As Excel.Workbook

Private WH_FOREIGN_CITY As Excel.Worksheet
Private WH_FOREIGN_FOOD As Excel.Worksheet
Private WH_JAPAN_CITY As Excel.Worksheet
Private WH_JAPAN_FOOD As Excel.Worksheet
Private WH_CAT As Excel.Worksheet
Private WH_LANDSCAPE As Excel.Worksheet

Private DIC_FOREIGN_CITY_ROWS As Object
Private DIC_FOREIGN_FOOD_ROWS As Object
Private DIC_JAPAN_CITY_ROWS As Object
Private DIC_JAPAN_FOOD_ROWS As Object
Private DIC_CAT_ROWS As Object
Private DIC_LANDSCAPE_ROWS As Object

Const STR_DB_FILENAME As String = "C:\card\database\database.xlsx"
Const STR_FOREIGN_PATH As String = "C:\card\pic\f\"
Const STR_JAPAN_PATH As String = "C:card\pic\j\"
Const STR_CAT_PATH As String = "C:\card\pic\c\"
Const STR_LANDSCAPE_PATH As String = "C:\card\pic\l\"


Private Sub Class_Initialize()
    Call openFilenameDatabase
End Sub

Private Sub Class_Terminate()
    Call closeFilenameDatabase
End Sub


Private Sub openFilenameDatabase()

  'Excelを開いて名前付き範囲を2次元配列に格納
    Set OBJ_EXCEL = New Excel.Application
    Set WB_FILENAME_DATA = OBJ_EXCEL.Workbooks.Open(STR_DB_FILENAME)
    'OBJ_EXCEL.Windows(WB_FILENAME_DATA).Visible = True
    
    
    Set WH_FOREIGN_CITY = WB_FILENAME_DATA.Worksheets("foreign")
    Set WH_JAPAN_CITY = WB_FILENAME_DATA.Worksheets("japan")
    Set WH_LANDSCAPE = WB_FILENAME_DATA.Worksheets("landscape")

End Sub


Public Sub closeFilenameDatabase()
    
    WB_FILENAME_DATA.Close True
    Set WB_FILENAME_DATA = Nothing
    Set OBJ_EXCEL = Nothing

End Sub


Private Sub makeDic(argWh As Worksheet, argDic As Object)
    Dim cntOffsetRow As Long
    Dim dataStartCell As String
    Dim picOffsetCol As Long
    Dim cityOffsetCol As Long
    Dim flgOffsetCol As Long
    
    dataStartCell = "A2"
    picOffsetCol = 0
    cityOffsetCol = 1
    flgOffsetCol = 2
    
    
    If argWh.Range(dataStartCell).Offset(0, picOffsetCol).Value = "" Then
        Err.Raise 1000, , "データが存在しません"
    End If
    
    '既存のディクショナリは破棄する
    Set argDic = Nothing
    Set argDic = CreateObject("Scripting.Dictionary")
    
    Do
        cntOffsetRow = 0

        Do Until argWh.Range(dataStartCell).Offset(cntOffsetRow, picOffsetCol).Value = ""
            DoEvents
        
            If argWh.Range(dataStartCell).Offset(cntOffsetRow, flgOffsetCol).Value = "" Or _
               argWh.Range(dataStartCell).Offset(cntOffsetRow, flgOffsetCol).Value = 0 Then

                argDic.Add cntOffsetRow, 0
            Else
            
            
            End If
        

            cntOffsetRow = cntOffsetRow + 1
        Loop
        
        'すべて読み込みフラグが立っていたとき
        If argDic.Count = 0 Then
            'フラグを初期化して、ディクショナリの読み込みを再実行する
            argWh.Range( _
                    argWh.Range(dataStartCell).Offset(0, flgOffsetCol), _
                    argWh.Range(dataStartCell).Offset(cntOffsetRow - 1, flgOffsetCol) _
                    ).Value = 0

        End If
        
    Loop While argDic.Count = 0

End Sub


Private Function getInfo(argWh As Worksheet, argDic As Object, argPath As String) As Variant
    Dim aryTmp(1) As Variant
    Dim lngRndNum As Long
    Dim lngTgtOffsetRow As Long

    Dim dataStartCell As String
    Dim picOffsetCol As Long
    Dim cityOffsetCol As Long
    Dim flgOffsetCol As Long
    
    dataStartCell = "A2"
    picOffsetCol = 0
    cityOffsetCol = 1
    flgOffsetCol = 2


    If argDic Is Nothing Then
        Call makeDic(argWh, argDic)
    ElseIf argDic.Count = 0 Then
        Call makeDic(argWh, argDic)
    End If

    lngRndNum = Rnd * argDic.Count

    On Error GoTo errLabel
    lngTgtOffsetRow = argDic.Keys()(lngRndNum)
    On Error GoTo 0

    aryTmp(0) = argPath & argWh.Range(dataStartCell).Offset(lngTgtOffsetRow, picOffsetCol).Value
    aryTmp(1) = argWh.Range(dataStartCell).Offset(lngTgtOffsetRow, cityOffsetCol).Value
    'ファイル名を読み込んだのでflgを立てる
    argWh.Range(dataStartCell).Offset(lngTgtOffsetRow, flgOffsetCol).Value = 1
    argDic.Remove lngTgtOffsetRow
    
    getInfo = aryTmp

    Exit Function

errLabel:
    lngRndNum = Rnd * argDic.Count
    Resume
    
End Function



Public Function getForeignCityInfo() As Variant

    getForeignCityInfo = getInfo(WH_FOREIGN_CITY, DIC_FOREIGN_CITY_ROWS, STR_FOREIGN_PATH)

End Function


Public Function getJapanCityInfo() As Variant

    getJapanCityInfo = getInfo(WH_JAPAN_CITY, DIC_JAPAN_CITY_ROWS, STR_JAPAN_PATH)

End Function




Public Function getLandscapeInfo() As Variant

    getLandscapeInfo = getInfo(WH_LANDSCAPE, DIC_LANDSCAPE_ROWS, STR_LANDSCAPE_PATH)

End Function




Public Function getQRCodeInfo() As String

    getQRCodeInfo = "C:\card\pic\QR\businessCard-instaSide.PNG"

End Function
