Option Explicit

Private DEX As Object

Private WsData_ As Worksheet
Private ChartName_ As String

Private TargetDate_ As Date
Private TargetData_ As Variant
Private TargetRow_ As Range
Private TargetChartRange_ As Range

Private HeaderRange_ As Range
Private FirstDayRange_ As Range

Public Property Let WorksheetData(ws As Worksheet)
    Set WsData_ = ws
End Property

Public Property Let ChartName(nm As String)
    ChartName_ = nm
End Property

Public Property Let HeaderRange(rng As Range)
    Set HeaderRange_ = rng
End Property

Public Property Let FirstDayRange(rng As Range)
    Set FirstDayRange_ = rng
End Property

Private Sub Class_Initialize()

    Set DEX = New DataExtractor
    ChartName_ = "SalesTrandVisualizer"
    
End Sub

Public Sub execution()
    Call getDateData
    Call getTargetDateRow
    Call setData
    Call getChartRange
    Call deleteChart
    Call createChart
End Sub

'データソースから、対象日と対象データを取得
Private Sub getDateData()
    Call DEX.getDateData(TargetDate_, TargetData_)
End Sub

'取得した対象日を元に貼り付け行を特定する
Private Sub getTargetDateRow()
    Dim dayDiff As Long
    Dim tRange As Range
    
    ' 1. 日付の差分を計算
    dayDiff = TargetDate_ - FirstDayRange_.Value
    
    ' 2. 一旦変数に格納
    Set tRange = FirstDayRange_.Offset(dayDiff, 0)
    
    ' 3. 算出したセルの値と、探したい日付が一致するかチェック
    ' ※時刻が含まれる場合を考慮し、Date型で比較します
    If CDate(tRange.Value) <> TargetDate_ Then
        Err.Raise 1001, "getTargetDateRow", "該当の日付が見つかりませんでした"
        
        '検索ロジック挿入まで例外を投げる
        Exit Sub
    End If
    
    ' 4. 一致していればプロパティにセット
    Set TargetRow_ = tRange
End Sub


Private Sub setData()
    TargetRow_.Offset(0, 1).Resize(1, UBound(TargetData_, 1) - LBound(TargetData_, 1) + 1).Value = TargetData_
End Sub


Private Sub getChartRange()
    Set TargetChartRange_ = WsData_.Range(HeaderRange_.Cells(1, 1), TargetRow_.Offset(0, HeaderRange_.Columns.Count - 1))
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
        .HasTitle = False
'        .ChartTitle.Text = "データ推移"
    End With
End Sub
