Option Explicit

Private DEX As DataExtractor
Private GDR As getDataRows

Private WsData_ As Worksheet
Private ChartName_ As String

Private TargetDateDataMap_ As Object
Private TargetChartRange_ As Range

Private HeaderRange_ As Range
Private FirstDayRange_ As Range

Public Property Set WorksheetData(ws As Worksheet)
    Set WsData_ = ws
End Property

Public Property Let ChartName(nm As String)
    ChartName_ = nm
End Property

Public Property Set HeaderRange(rng As Range)
    Set HeaderRange_ = rng
End Property

Public Property Set FirstDayRange(rng As Range)
    Set FirstDayRange_ = rng
End Property

Private Sub Class_Initialize()

    Set DEX = New DataExtractor
    Set GDR = New getDataRows
    ChartName_ = "SalesTrandVisualizer"
    
End Sub

Public Sub execution()
    Call checkBeforeExec
    Call getDateDataMap
    Call applyDateDataMap
    Call getChartRange
    Call deleteChart
    Call createChart
End Sub


Private Sub checkBeforeExec()
    If WsData_ Is Nothing Then
        Err.Raise 9999, , "ワークシートが設定されていません"
    End If

    Set GDR.ws = WsData_

    If HeaderRange_ Is Nothing Then
        Err.Raise 9999, , "ヘッダーが設定されていません"
    End If

    'rangeが別シートの場合を想定して、指定ワークシートの同じセル範囲を再設定する
    Set HeaderRange_ = WsData_.Range(HeaderRange_.Address)

    '指定されていない場合は、ヘッダーの左上セルの１つ下を基準日セルと想定する
    If FirstDayRange_ Is Nothing Then
        Set FirstDayRange_ = HeaderRange_.Cells(1, 1).Offset(1, 0)
    '指定がある場合、指定ワークシートの同じセル範囲を再設定する
    Else
        Set FirstDayRange_ = WsData_.Range(FirstDayRange_.Address)
    End If

End Sub

'データソースから、対象日と対象データを取得
Private Sub getDateDataMap()
    Set TargetDateDataMap_ = DEX.getDateDataMap()
End Sub

Private Sub applyDateDataMap()
    Dim k As Variant
    Dim tDate As Date
    Dim tData As Variant
    Dim tRow As Range
    
    For Each k In TargetDateDataMap_.Keys
    
        tDate = DateSerial(Left(k, 4), Mid(k, 5, 2), Mid(k, 7, 2))
        tData = TargetDateDataMap_(k)
        Set tRow = getTargetDateRow(tDate)
        'シートに該当日付があるときのみデータを設定する
        If Not tRow Is Nothing Then
            Call setData(tRow, tData)
        End If
    Next k
End Sub

'取得した対象日を元に貼り付け行を特定する
Private Function getTargetDateRow(targetDate As Date) As Range
    Dim dayDiff As Long
    Dim tRange As Range
    
    ' 1. 日付の差分を計算
    dayDiff = targetDate - FirstDayRange_.Value
    
    ' 2. 一旦変数に格納
    Set tRange = FirstDayRange_.Offset(dayDiff, 0)
    
    ' 3. 算出したセルの値と、探したい日付が一致するかチェック
    If CDate(tRange.Value) <> targetDate Then
        Set tRange = FirstDayRange_.EntireColumn.Find(What:=targetDate, LookIn:=xlFormulas, LookAt:=xlWhole)
        If tRange Is Nothing Then
            Set tRange = FirstDayRange_.EntireColumn.Find(What:=targetDate, LookIn:=xlValues, LookAt:=xlWhole)
'            If tRange Is Nothing Then
'                Err.Raise 1001, "getTargetDateRow", "該当の日付が見つかりませんでした"
'            End If
        End If
    End If
    ' 4. 一致・検索出来ればプロパティにセット
    Set getTargetDateRow = tRange
End Function

Private Sub setData(targetRow, targetData)
    Dim cCnt As Long:   cCnt = HeaderRange_.Columns.Count - 1 'ヘッダーに日付列も含める場合は-1する
    Dim rStIdx As Long: rStIdx = LBound(targetData, 1)
    Dim rEnIdx As Long: rEnIdx = UBound(targetData, 1)
    Dim cStIdx As Long: cStIdx = LBound(targetData, 2)
    Dim cEnIdx As Long: cEnIdx = LBound(targetData, 2) + cCnt - 1
    
    ReDim Preserve targetData(rStIdx To rEnIdx, cStIdx To cEnIdx)
    
    ' シートへ貼り付け
    'targetRow=該当日付セルなので、1列右からデータを貼り付ける
    targetRow.Offset(0, 1).Resize(1, cCnt).Value = targetData
End Sub

Private Sub getChartRange()
    '日付を除いたヘッダー範囲を渡す
    Set GDR.TargetHeadersRange = HeaderRange_.Offset(0, 1).Resize(HeaderRange_.Rows.Count, HeaderRange_.Columns.Count - 1)
    Set TargetChartRange_ = GDR.getDataRange(True, FirstDayRange_.Column)
End Sub

Private Sub deleteChart()
    Dim shp As Shape
    For Each shp In WsData_.Shapes
        If shp.Name = ChartName_ Then
            shp.Delete
            Exit Sub
        End If
    Next shp
End Sub

Private Sub createChart()
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = TargetChartRange_.Parent
    
    Set shp = ws.Shapes.AddChart2(Style:=227, XlChartType:=xlLine)
    shp.Name = ChartName_
    
    With shp.Chart
        .SetSourceData Source:=TargetChartRange_
        .PlotBy = xlColumns
        .DisplayBlanksAs = xlInterpolated
        .HasTitle = False
'        .ChartTitle.Text = "データ推移"
    End With
End Sub
