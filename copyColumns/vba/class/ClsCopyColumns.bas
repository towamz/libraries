Option Explicit

Private DataSheet_ As String
Private ResultSheet_ As String
Private TargetColumns_() As String
Private TargetTitles_() As String
Private TargetTitleRow_ As Long

'データシートを設定する
Public Property Let DataSheet(arg1 As String)
    DataSheet_ = arg1
End Property

'処理結果シートを設定する(任意)
Public Property Let ResultSheet(arg1 As String)
    ResultSheet_ = arg1
End Property

'タイトル行番号
Public Property Let TargetTitleRow(arg1 As Long)
    If arg1 < 1 Or ActiveSheet.Rows.Count < arg1 Then
        Err.Raise 1006, , "有効な行番号を設定してください"
    End If
    TargetTitleRow_ = arg1
End Property

Public Property Get TargetTitleRow() As Long
    TargetTitleRow = TargetTitleRow_
End Property

Public Property Get TargetTitles() As String
    TargetTitles = Join(TargetTitles_, vbCrLf)
End Property

Public Property Get TargetColumns() As String
    TargetColumns = Join(TargetColumns_, vbCrLf)
End Property

'データシートの列名チェックを設定する(任意)
Public Property Let TargetTitles(arg1 As String)
    If (Not TargetTitles_) = -1 Then
        ReDim TargetTitles_(0)
        TargetTitles_(0) = arg1
    Else
        ReDim Preserve TargetTitles_(UBound(TargetTitles_) + 1)
        TargetTitles_(UBound(TargetTitles_)) = arg1
    End If
End Property

'データシートの列名チェックを設定する(任意)
Public Property Let TargetColumns(arg1 As String)
    Dim testRange As Range
    Dim errNumber As Long

    If InStr(1, arg1, ":", vbTextCompare) = 0 Then
        arg1 = arg1 & ":" & arg1
    End If

    On Error Resume Next
    Set testRange = Range(arg1)
    errNumber = Err.Number
    On Error GoTo 0

    Select Case errNumber
        Case 0
        Case 1004
            Err.Raise 1001, , "有効なセル番地を設定してください"
        Case Else
            Err.Raise 9999, , "不明なエラーが発生しました"
    End Select

    If testRange.Address <> testRange.EntireColumn.Address Then
        Err.Raise 1002, , "有効なセル番地(列全体)を設定してください"
    End If

    If (Not TargetColumns_) = -1 Then
        ReDim TargetColumns_(0)
        TargetColumns_(0) = arg1
    Else
        ReDim Preserve TargetColumns_(UBound(TargetColumns_) + 1)
        TargetColumns_(UBound(TargetColumns_)) = arg1
    End If

End Property

Private Sub Class_Initialize()
    ResultSheet_ = "結果"
    TargetTitleRow_ = 1
End Sub

'列コピー
Public Sub copyColumns()
    Dim wbData As Workbook
    Dim wbResult As Workbook
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim i As Long

    If DataSheet_ = "" Then
        Err.Raise 1051, , "データシート名を設定してください"
    End If

    If (Not TargetColumns_) = -1 Then
        Err.Raise 1011, , "コピー対象のセル番地(列全体)を設定してください"
    End If

    Set wbData = ThisWorkbook
    Set wbResult = ThisWorkbook

    Set wsData = wbData.Worksheets(DataSheet_)

    On Error Resume Next
    Set wsResult = wbData.Worksheets(ResultSheet_)
    On Error GoTo 0

    If wsResult Is Nothing Then
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = ResultSheet_
    Else
        Set wsResult = wbResult.Worksheets.Add
        wsResult.Name = ResultSheet_ & Format(Now(), "yymmdd-hhnnss")
    End If

    i = 0
    Do
        wsData.Range(TargetColumns_(i)).Copy Destination:=wsResult.Columns(i + 1)
        i = i + 1
    Loop Until UBound(TargetColumns_) < i

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
    If DataSheet_ = "" Then
        Err.Raise 1051, , "データシート名を設定してください"
    End If

    'コピー対象列(タイトル文字列)が設定されていないとき
    If (Not TargetTitles_) = -1 Then
        Err.Raise 1012, , "コピー対象のタイトル文字列を設定してください"
    End If

    'コピー対象列(列記号)が設定されているとき
    If (Not TargetColumns_) <> -1 Then
        MsgBox "コピー対象の列がセル番地とタイトル文字列で指定されています。セル番地で指定した列を先にコピーします", vbOKOnly + vbInformation
'        If MsgBox("コピー対象の列がセル番地で入力されています。セル番地を消去してタイトル文字列で列コピーを実施してもいいですか。", vbYesNo) = vbYes Then
'            Erase TargetColumns_
'        Else
'            Exit Sub
'        End If
    End If

    Set wbData = ThisWorkbook
    Set wbResult = ThisWorkbook

    Set wsData = wbData.Worksheets(DataSheet_)

    '指定されたタイトルを検索して、列記号を取得する
    i = 0
    Do
        'タイトルが2つ以上ある場合は例外を投げるため2回検索
        Set testRange = wsData.Rows(TargetTitleRow_).Find(What:=TargetTitles_(i), LookIn:=xlValues, LookAt:=xlWhole)

        If testRange Is Nothing Then
            Debug.Print "指定されたタイトル文字列は見つかりませんでした:" & TargetTitles_(i)
        Else
            firstAddress = testRange.Address
            Set testRange = wsData.Rows(TargetTitleRow_).FindNext(testRange)

            If firstAddress = testRange.Address Then
                '登録
                Me.TargetColumns = testRange.EntireColumn.Address
                Debug.Print "登録します:" & TargetTitles_(i) & ":" & firstAddress
            Else
                'エラー
                Err.Raise 1021, , "指定されたタイトル文字列が2つ以上あります"
            End If

        End If

        i = i + 1
    Loop Until UBound(TargetTitles_) < i

    Me.copyColumns

End Sub
