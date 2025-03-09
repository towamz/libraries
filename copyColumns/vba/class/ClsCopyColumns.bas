Option Explicit

Private DataSheet As String
Private ResultSheet As String
Private TargetColumns() As String

'データシートを設定する
Public Property Let setDataSheet(arg1 As String)
    DataSheet = arg1
End Property

'処理結果シートを設定する(任意)
Public Property Let setResultSheet(arg1 As String)
    ResultSheet = arg1
End Property

'データシートの列名チェックを設定する(任意)
Public Property Let setTargetColumns(arg1 As String)
    Dim testRange As Range
    Dim errNumber As Long
    
    '引数で与えられた文字列を列全体形式に加工(例: A -> A:A)
    If InStr(1, arg1, ":", vbTextCompare) = 0 Then
        arg1 = arg1 & ":" & arg1
    End If
    
    'セルの有効性確認
    On Error Resume Next
    Set testRange = Range(arg1)
    
    'エラー処理無効にしないと例外を投げられないのでエラー番号を取得する
    errNumber = Err.Number
    On Error GoTo 0
    
    Select Case errNumber
        Case 0
            '何もしない
        Case 1004
            Err.Raise 1001, , "有効なセル番地を設定してください"
        Case Else
            Err.Raise 9999, , "不明なエラーが発生しました"
    End Select

    '列全体のセル番地か確認
    If testRange.Address <> testRange.EntireColumn.Address Then
        Err.Raise 1002, , "有効なセル番地(列全体)を設定してください"
    End If

    'セル番地(列全体)の文字列を保存する
    '動的配列が初期化されていないとき
    If (Not TargetColumns) = -1 Then
        ReDim TargetColumns(0)
        TargetColumns(0) = arg1
    
    '動的配列が初期化されているとき
    Else
        ReDim Preserve TargetColumns(UBound(TargetColumns) + 1)
        TargetColumns(UBound(TargetColumns)) = arg1
    End If

End Property

Private Sub Class_Initialize()
    '結果シートの既定値
    ResultSheet = "結果"
End Sub

'列コピー
Public Sub copyColumns()
    Dim wbData As Workbook
    Dim wbResult As Workbook
    
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    
    Dim i As Long

    '必須項目チェック / mandatory parameter check
    If DataSheet = "" Then
        Err.Raise 1051, , "データシート名を設定してください"
    End If
    
    'コピー対象列が設定されていないとき
    If (Not TargetColumns) = -1 Then
        Err.Raise 1011, , "コピー対象のセル番地(列全体)を設定してください"
    End If

    Set wbData = ThisWorkbook
    Set wbResult = ThisWorkbook

    Set wsData = wbData.Worksheets(DataSheet)
    
    On Error Resume Next
    Set wsResult = wbData.Worksheets(ResultSheet)
    On Error GoTo 0
    
    '結果シートが存在しない場合は、シート名を結果シート名にする
    If wsResult Is Nothing Then
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = ResultSheet
    '結果シートが存在する場合は、シート名を結果シート名+yymmdd-hhnnssにする
    Else
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = ResultSheet & Format(Now(), "yymmdd-hhnnss")
    End If

    i = 0
    Do
        wsData.Range(TargetColumns(i)).Copy Destination:=wsResult.Columns(i + 1)
        i = i + 1
    Loop Until UBound(TargetColumns) < i

End Sub
