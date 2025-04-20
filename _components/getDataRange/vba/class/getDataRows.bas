Option Explicit

Private Ws_ As Worksheet

Private FirstRowNumber_ As Long
Private LastRowNumber_ As Long
Private LastRowDataNumber_ As Long
Private TargetHeadersRowNumber_ As Long
Private TargetHeaders_() As String

Public Property Set Ws(arg1 As Worksheet)
    Set Ws_ = arg1
End Property

Public Property Let FirstRowNumber(arg1 As Long)
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    Debug.Print Rows(arg1).Address
    
    If Err.Number <> 0 Then
        Debug.Print Err.Number & Err.Description
        Err.Raise 1001, , "行番号が正しくありません"
    End If
    On Error GoTo 0
    
    FirstRowNumber_ = arg1
End Property


Public Property Let LastRowNumber(arg1 As Long)
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    Debug.Print Rows(arg1).Address
    
    If Err.Number <> 0 Then
        Debug.Print Err.Number & Err.Description
        Err.Raise 1001, , "行番号が正しくありません"
    End If
    On Error GoTo 0
    
    LastRowNumber_ = arg1
End Property


Public Property Let TargetHeadersRowNumber(arg1 As Long)
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    Debug.Print Rows(arg1).Address
    
    If Err.Number <> 0 Then
        Debug.Print Err.Number & Err.Description
        Err.Raise 1001, , "行番号が正しくありません"
    End If
    On Error GoTo 0
    
    TargetHeadersRowNumber_ = arg1

End Property


Public Property Let TargetColumnLetter(arg1 As String)
    If (Not TargetHeaders_) = -1 Then
        ReDim TargetHeaders_(1, 0)
    Else
        ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
    End If
    
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    TargetHeaders_(0, UBound(TargetHeaders_, 2)) = Ws_.Columns(arg1).Column
    TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 0

End Property


Public Property Let TargetColumnNumber(arg1 As Long)
    If (Not TargetHeaders_) = -1 Then
        ReDim TargetHeaders_(1, 0)
    Else
        ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
    End If
    
    TargetHeaders_(0, UBound(TargetHeaders_, 2)) = arg1
    TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 0
End Property


Public Property Let TargetHeader(arg1 As String)
    If (Not TargetHeaders_) = -1 Then
        ReDim TargetHeaders_(1, 0)
    Else
        ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
    End If
    
    TargetHeaders_(0, UBound(TargetHeaders_, 2)) = arg1
    TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 1
End Property


Private Sub Class_Initialize()
    '検索対象行を設定(デフォルトでは全行) / set target rows (default is all rows)
    FirstRowNumber_ = 1
    LastRowNumber_ = Rows.Count

    '見出し行(デフォルト:1) / headerRow (default is 1)
    TargetHeadersRowNumber_ = 1
End Sub

Private Sub getColumnNumberFromHeader()
    Dim testRange As Range
    Dim firstAddress As String
    Dim isValid As Boolean
    Dim i As Long
    
    isValid = False
    
    For i = LBound(TargetHeaders_, 2) To UBound(TargetHeaders_, 2)
        '見出し文字列で指定されている場合は検索する
        If TargetHeaders_(1, i) = 1 Then
        
            'タイトルが2つ以上ある場合は例外を投げるため2回検索
            Set testRange = Ws_.Rows(TargetHeadersRowNumber_).Find(What:=TargetHeaders_(0, i), LookIn:=xlValues, LookAt:=xlWhole)
    
            If testRange Is Nothing Then
                Debug.Print "指定されたタイトル文字列は見つかりませんでした:" & TargetHeaders_(0, i)
                TargetHeaders_(0, i) = -1
                TargetHeaders_(1, i) = -1
                '一つでも有効なデータがあれば有効なので、isValid = Falseはしない
            Else
                firstAddress = testRange.Address
                Set testRange = Ws_.Rows(TargetHeadersRowNumber_).FindNext(testRange)
    
                If firstAddress = testRange.Address Then
                    '登録
                    Debug.Print "登録します:" & i & ":" & TargetHeaders_(0, i) & ":" & firstAddress
                    TargetHeaders_(0, i) = testRange.EntireColumn.Column
                    TargetHeaders_(1, i) = 0
                    '見出し文字列から列番号に変換できたので有効判定する
                    isValid = True
                Else
                    'エラー
                    Err.Raise 1021, , "指定されたタイトル文字列が2つ以上あります"
                End If
            End If
        Else
            '列番号で指定されているデータがあったので有効判定する
            isValid = True
        End If
    Next i

    If Not isValid Then
        Err.Raise 1004, , "有効な検索対象列が設定されていません。"
    End If

End Sub

'データ最終行番号を取得 / get the data last row number
Public Function getLastRow() As Long
    Dim currentLastRowDataNumber As Long
    Dim i As Long
    
    '--------------初期確認--------------
    If Ws_ Is Nothing Then
        Err.Raise 1001, , "検索対象シートが設定されていません。"
    End If
    
    If FirstRowNumber_ > LastRowNumber_ Then
        Err.Raise 1002, , "検索対象行(開始) < 検索対象行(終了)に設定してください。" & vbCrLf & "開始行:" & FirstRowNumber_ & vbCrLf & "終了行:" & LastRowNumber_
    End If

    If (Not TargetHeaders_) = -1 Then
        Err.Raise 1003, , "検索対象列が設定されていません。"
    End If
    '--------------初期確認終了--------------
    
    '見出し文字列から列番号に変換する
    Call getColumnNumberFromHeader

    '行番号を-1(データなし)に設定
    LastRowDataNumber_ = -1
    
    For i = LBound(TargetHeaders_, 2) To UBound(TargetHeaders_, 2)
        '列番号指定であれば検索対象
        If TargetHeaders_(1, i) = 0 Then
               
            'データ範囲最終にデータがあった場合は最終行を設定してループを抜ける(これ以上検索不要)
            'exit for if data exist in the last row(no need to further search)
            If Ws_.Cells(LastRowNumber_, CLng(TargetHeaders_(0, i))) <> "" Then
                LastRowDataNumber_ = LastRowNumber_
                Exit For
            End If
        
            currentLastRowDataNumber = Ws_.Cells(LastRowNumber_, CLng(TargetHeaders_(0, i))).End(xlUp).row
        
            '取得した行が検索開始行と同じときはセルに値があるか確認する
            'check the cell if the currentRow is equal to the first row
            If currentLastRowDataNumber = FirstRowNumber_ Then
                If Ws_.Cells(currentLastRowDataNumber, CLng(TargetHeaders_(0, i))) = "" Then
                    currentLastRowDataNumber = -1
                End If
            '取得した行が検索開始行より小さい場合はデータなしと判定する
            'judge no data if the currentRow is smaller than the first row
            ElseIf currentLastRowDataNumber < FirstRowNumber_ Then
                    currentLastRowDataNumber = -1
            End If
    
            '今回取得した最終行が今まで最終行より大きい場合は書き換え
            'overwrite if the currentRow is bigger than the previous Row
            If LastRowDataNumber_ < currentLastRowDataNumber Then
                LastRowDataNumber_ = currentLastRowDataNumber
            End If
        End If
    Next i
    
    getLastRow = LastRowDataNumber_

End Function


'データ範囲のrangeオブジェクトを取得(行全体) / get data range object(entire rows)
Public Function getDataRows() As Range
    
    Call getLastRow
    
    If LastRowDataNumber_ = -1 Then
        'データがないときはnothingを返す / set nothing if data does not exist
        Set getDataRows = Nothing
    Else
        Set getDataRows = Ws_.Rows(CStr(FirstRowNumber_) & ":" & CStr(LastRowDataNumber_))
    End If

End Function


'Err.Raise 1001, , "行番号が正しくありません"


