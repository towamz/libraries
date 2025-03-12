Option Explicit

Private PDataSheet As String
Private PResultSheet As String
Private PTargetColumns() As String
Private PTargetTitles() As String
Private PTargetTitleRow As Long

'データシートを設定する
Public Property Let DataSheet(arg1 As String)
    PDataSheet = arg1
End Property

'処理結果シートを設定する(任意)
Public Property Let ResultSheet(arg1 As String)
    PResultSheet = arg1
End Property

'タイトル行番号
Public Property Let TargetTitleRow(arg1 As Long)
    If arg1 < 1 Or ActiveSheet.Rows.Count < arg1 Then
        Err.Raise 1012, , "有効な行番号を設定してください"
    End If
    PTargetTitleRow = arg1
End Property

Public Property Get TargetTitleRow() As Long
    TargetTitleRow = PTargetTitleRow
End Property

Public Property Get TargetTitles() As String
    TargetTitles = Join(PTargetTitles, vbCrLf)
End Property

Public Property Get TargetColumns() As String
    PTargetColumns = Join(PTargetColumns, vbCrLf)
End Property

'データシートの列名チェックを設定する(任意)
Public Property Let TargetTitles(arg1 As String)
    'セル番地(列全体)の文字列を保存する
    '動的配列が初期化されていないとき
    If (Not PTargetTitles) = -1 Then
        ReDim PTargetTitles(0)
        PTargetTitles(0) = arg1
    '動的配列が初期化されているとき
    Else
        ReDim Preserve PTargetTitles(UBound(PTargetTitles) + 1)
        PTargetTitles(UBound(PTargetTitles)) = arg1
    End If
End Property

'データシートの列名チェックを設定する(任意)
Public Property Let TargetColumns(arg1 As String)
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
    If (Not PTargetColumns) = -1 Then
        ReDim PTargetColumns(0)
        PTargetColumns(0) = arg1
    
    '動的配列が初期化されているとき
    Else
        ReDim Preserve PTargetColumns(UBound(PTargetColumns) + 1)
        PTargetColumns(UBound(PTargetColumns)) = arg1
    End If
End Property

Private Sub Class_Initialize()
    '結果シートの既定値
    PResultSheet = "結果"
    PTargetTitleRow = 1
End Sub

'列コピー
Public Sub copyColumns()
    Dim wbData As Workbook
    Dim wbResult As Workbook
    
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    
    Dim i As Long

    '必須項目チェック / mandatory parameter check
    If PDataSheet = "" Then
        Err.Raise 1051, , "データシート名を設定してください"
    End If
    
    'コピー対象列が設定されていないとき
    If (Not PTargetColumns) = -1 Then
        Err.Raise 1011, , "コピー対象のセル番地(列全体)を設定してください"
    End If

    Set wbData = ThisWorkbook
    Set wbResult = ThisWorkbook

    Set wsData = wbData.Worksheets(PDataSheet)
    
    On Error Resume Next
    Set wsResult = wbData.Worksheets(PResultSheet)
    On Error GoTo 0
    
    '結果シートが存在しない場合は、シート名を結果シート名にする
    If wsResult Is Nothing Then
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = PResultSheet
    '結果シートが存在する場合は、シート名を結果シート名+yymmdd-hhnnssにする
    Else
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = PResultSheet & Format(Now(), "yymmdd-hhnnss")
    End If

    i = 0
    Do
        wsData.Range(PTargetColumns(i)).Copy Destination:=wsResult.Columns(i + 1)
        i = i + 1
    Loop Until UBound(PTargetColumns) < i

End Sub

'列コピー
Public Sub copyColumnsByTitles()
    Dim wbData As Workbook
    Dim wbResult As Workbook
    
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    
    Dim testRange As Range
    Dim firstAddress As String
    
    Dim i As Long

    '必須項目チェック / mandatory parameter check
    If PDataSheet = "" Then
        Err.Raise 1051, , "データシート名を設定してください"
    End If
    
    'コピー対象列(タイトル文字列)が設定されていないとき
    If (Not PTargetTitles) = -1 Then
        Err.Raise 1012, , "コピー対象のタイトルを設定してください"
    End If

    'コピー対象列(列記号)が設定されているとき
    If (Not PTargetColumns) <> -1 Then
        If MsgBox("コピー対象の列がセル番地で入力されています。セル番地を消去してタイトル文字列で列コピーを実施してもいいですか。", vbYesNo) = vbYes Then
            Erase PTargetColumns
        Else
            Exit Sub
        End If
    End If

    Set wbData = ThisWorkbook
    Set wbResult = ThisWorkbook

    Set wsData = wbData.Worksheets(PDataSheet)
    
    '指定されたタイトルを検索して、列記号を取得する
    i = 0
    Do
        'タイトルが2つ以上ある場合は例外を投げるため2回検索
        Set testRange = wsData.Rows(PTargetTitleRow).Find(What:=PTargetTitles(i), LookIn:=xlValues, LookAt:=xlWhole)
        
        If testRange Is Nothing Then
            Debug.Print "指定されたタイトル文字列は見つかりませんでした:" & PTargetTitles(i)
        Else
            firstAddress = testRange.Address
            Set testRange = wsData.Rows(PTargetTitleRow).FindNext(testRange)
            
            If firstAddress = testRange.Address Then
                '登録
                Me.TargetColumns = testRange.EntireColumn.Address
                Debug.Print "登録します:" & PTargetTitles(i) & ":" & firstAddress
            Else
                'エラー
                Err.Raise 1021, , "指定されたタイトル文字列が2つ以上あります"
            End If
        End If
        i = i + 1
    Loop Until UBound(PTargetTitles) < i

    Me.copyColumns

End Sub
