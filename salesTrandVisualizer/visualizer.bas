Sub visualizerTest()
    Dim targetDate As Date
    Dim targetData As Variant
    Dim targetRow As Range
    Dim dr As Range
    
    '日付・データ取得
    Call getDateData(targetDate, targetData)

    '貼り付け位置特定
    Set targetRow = getTargetDateRange(Range("A2"), targetDate)

    '貼り付け
    targetRow.Offset(0, 1).Resize(1, UBound(targetData, 1) - LBound(targetData, 1) + 1).Value = targetData
    
    'グラフ描画範囲特定
    Set dr = Worksheets("sheet1").Range(Worksheets("sheet1").Range("A1"), targetRow.Offset(0, UBound(targetData, 1) + 1))
    
    'グラフ描画
    Call createChart(dr)

End Sub

'データソースから該当日の数値を取得する関数
'データの構成はシステムにより違うのでstubにする
Sub getDateData(ByRef tDate, ByRef tData)

    tDate = DateValue("2026/3/5")
    tDate = DateValue("2026/3/6")
    tDate = DateValue("2026/3/7")
    tData = Array(5, 7, 6)
    tData = Array(10, 17, 16)

End Sub

'貼り付け対象行の特定
Function getTargetDateRange(firstDayRange As Range, targetDate As Date) As Range
    Dim dayDiff As Long
    Dim tRange As Range
    
    ' 1. 日付の差分を計算
    dayDiff = targetDate - firstDayRange.Value
    
    ' 2. 一旦変数に格納
    Set tRange = firstDayRange.Offset(dayDiff, 0)
    
    ' 3. 算出したセルの値と、探したい日付が一致するかチェック
    ' ※時刻が含まれる場合を考慮し、Date型で比較します
    If CDate(tRange.Value) <> targetDate Then
        '検索ロジック挿入
        
        Exit Function
    End If
    
    ' 4. 一致していれば戻り値としてセット
    Set getTargetDateRange = tRange
End Function

'グラフ作成
Sub createChart(dataRange As Range)
    Dim ws As Worksheet
    Dim shp As Shape
    
    Set ws = dataRange.Parent
    
    Set shp = ws.Shapes.AddChart2(Style:=227, XlChartType:=xlLine)
    
    With shp.Chart
        .SetSourceData Source:=dataRange
        .PlotBy = xlColumns
        .HasTitle = False
'        .ChartTitle.Text = "データ推移"
    End With
End Sub
