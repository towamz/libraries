Private wb As Workbook
Private wsData As Worksheet
Private wsPivod As Worksheet
'Private wsDataName As String
'Private wsPivodName As String
Private rgData As Range
Private rgPivod As Range
Private rgPivodName As String
Private pc As PivotCache
Private pt As PivotTable

Private pivotFieldsNamesRows() As String
Private pivotFieldsNamesColumns() As String
Private pivotFieldsNamesPages() As String
Private pivotFieldsNamesDataField() As String

Private indexPivotFieldsNamesRows As Long
Private indexPivotFieldsNamesColumns As Long
Private indexPivotFieldsNamesPages As Long
Private indexPivotFieldsNamesDataField As Long

Public Property Let dataSheetName(sheetName)
    '存在しないシートをデータ元に指定できないのでオブジェクト変数に直接代入する
    Set wsData = wb.Sheets(sheetName)
End Property

Public Property Let pivodSheetName(sheetName)
    '出力先を指定するときはシートが存在することが前提なのでオブジェクト変数に直接代入する
    On Error Resume Next
    Set wsPivod = wb.Sheets(sheetName)
    On Error GoTo 0
    
    If wsPivod Is Nothing Then
        Set wsPivod = Sheets.add
        wsPivod.Name = sheetName    'Format(Now(), "yymmddhhmmss")
    End If
    
    
End Property

Public Property Let dataRangeName(rangeName)
    Set rgData = wsData.Range(rangeName)
End Property

Public Property Let pivodRangeName(rangeName)
    'ピボッドの開始セルはシート名を指定しない場合は新規シートを作成するのでstringで保持する
    rgPivodName = rangeName
End Property

Private Sub AddFieldName(ByRef pivotFieldsNames() As String, ByRef indexPivotFieldsNames As Long, fieldName As String)

    ' インデックスを増加
    indexPivotFieldsNames = indexPivotFieldsNames + 1

    ' 配列サイズがインデックスを超えた場合、次の2乗サイズでReDim
    If indexPivotFieldsNames > UBound(pivotFieldsNames) Then
        ReDim Preserve pivotFieldsNames(UBound(pivotFieldsNames) ^ 2)
    End If
    
    ' フィールド名を追加
    pivotFieldsNames(indexPivotFieldsNames) = fieldName
End Sub

Public Sub addPageFieldName(fieldName As String)
    Call AddFieldName(pivotFieldsNamesPages, indexPivotFieldsNamesPages, fieldName)
End Sub

Public Sub addRowFieldName(fieldName As String)
    Call AddFieldName(pivotFieldsNamesRows, indexPivotFieldsNamesRows, fieldName)
End Sub

Public Sub addColumnFieldName(fieldName As String)
    Call AddFieldName(pivotFieldsNamesColumns, indexPivotFieldsNamesColumns, fieldName)
End Sub


Public Sub addDataFieldName(fieldName As String, calculationType As XlConsolidationFunction)
'　コード検証用
'    ActiveSheet.PivotTables("ピボットテーブル5").AddDataField ActiveSheet.PivotTables( _
'        "ピボットテーブル5").PivotFields("金額"), "合計 / 金額", xlSum
    
    ' インデックスを増加
    indexPivotFieldsNamesDataField = indexPivotFieldsNamesDataField + 1

    ' 配列サイズがインデックスを超えた場合、次の2乗サイズでReDim
    If indexPivotFieldsNamesDataField > UBound(pivotFieldsNamesDataField) Then
        ReDim Preserve pivotFieldsNamesDataField(UBound(pivotFieldsNamesDataField) ^ 2)
    End If
    
    ' フィールド名を追加
    pivotFieldsNamesDataField(0, indexPivotFieldsNamesDataField) = fieldName
    pivotFieldsNamesDataField(1, indexPivotFieldsNamesDataField) = calculationType

End Sub



Private Sub Class_Initialize()
    Set wb = ThisWorkbook
    Set wsData = wb.Sheets(1)

    ReDim pivotFieldsNamesRows(8)
    ReDim pivotFieldsNamesColumns(8)
    ReDim pivotFieldsNamesPages(8)
    ReDim pivotFieldsNamesDataField(1, 8)

    indexPivotFieldsNamesRows = -1
    indexPivotFieldsNamesColumns = -1
    indexPivotFieldsNamesPages = -1
    indexPivotFieldsNamesDataField = -1
End Sub

Private Sub initializeBeforeExec()
    If wsPivod Is Nothing Then
        Me.pivodSheetName = Format(Now(), "yymmddhhmmss")
'        Set wsPivod = Sheets.add
'        wsPivod.Name = Format(Now(), "yymmddhhmmss")
    End If

    If rgData Is Nothing Then
        Set rgData = wsData.Range("A1").CurrentRegion
    ElseIf rgData.Rows.Count = 1 Or rgData.Columns.Count Then
        Set rgData = rgData.CurrentRegion
    End If


End Sub

Private Sub adjustPivotFieldsNamesIndex()

    If indexPivotFieldsNamesRows > -1 Then
        ReDim Preserve pivotFieldsNamesRows(indexPivotFieldsNamesRows)
    End If
    
    If indexPivotFieldsNamesColumns > -1 Then
        ReDim Preserve pivotFieldsNamesColumns(indexPivotFieldsNamesColumns)
    End If
    
    If indexPivotFieldsNamesPages > -1 Then
        ReDim Preserve pivotFieldsNamesPages(indexPivotFieldsNamesPages)
    End If

    If indexPivotFieldsNamesDataField > -1 Then
        ReDim Preserve pivotFieldsNamesDataField(1, indexPivotFieldsNamesDataField)
    End If

End Sub


Public Sub createPivodTable()
    'SourceData:= wsdata.Name & "!" & rgdata.Name _    '"Sheet1!R1C1:R8C5", _
    'SourceData:=wsData.Name & "!" & rgData.Address, _
'    Set pc = wb.PivotCaches.Create( _
'        SourceType:=xlDatabase, _
'        SourceData:=rgData, _
'        Version:=8)
'
'    Set pt = pc.CreatePivotTable( _
'        TableDestination:=wsPivod.Name & "!R3C1", _
'        TableName:=wsPivod.Name, _
'        DefaultVersion:=8)

    Call initializeBeforeExec

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rgData)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivod.Name & "!R3C1", _
        TableName:=wsPivod.Name)


End Sub


Public Sub addFields()
    Call adjustPivotFieldsNamesIndex

    Call AddField(pivotFieldsNamesRows, indexPivotFieldsNamesRows, xlRowField)
    Call AddField(pivotFieldsNamesColumns, indexPivotFieldsNamesColumns, xlColumnField)
    Call AddField(pivotFieldsNamesPages, indexPivotFieldsNamesPages, xlPageField)
    
    Call AddFieldDataField

End Sub


Private Sub AddField(ByRef pivotFieldsNames() As String, ByRef indexPivotFieldsNames As Long, fieldDataType As XlPivotFieldOrientation)
    Dim index As Long
    
    If indexPivotFieldsNames = -1 Then Exit Sub
    
    For index = LBound(pivotFieldsNames) To UBound(pivotFieldsNames)
        With pt.PivotFields(pivotFieldsNames(index))
            .Orientation = fieldDataType
            .Position = index + 1
        End With
    Next index

End Sub


Private Sub AddFieldDataField()
    Dim index As Long
    Dim fieldName As String
    Dim calculationType As XlConsolidationFunction
    
    If indexPivotFieldsNamesDataField = -1 Then Exit Sub

'  コード検証用に残す
'    pt.addDataField pt.PivotFields("金額"), "合計 / 金額", xlSum
    
    For index = LBound(pivotFieldsNamesDataField, 2) To UBound(pivotFieldsNamesDataField, 2)
        '配列を直接指定するとエラーとなるため一旦変数に格納する
        fieldName = pivotFieldsNamesDataField(0, index)
        calculationType = pivotFieldsNamesDataField(1, index)
        
        pt.addDataField pt.PivotFields(fieldName), _
                        "集計", _
                        calculationType
    Next index

End Sub

