Option Explicit

Private Ws_ As Worksheet

Private FirstRowNumber_ As Long
Private LastRowNumber_ As Long
Private HeaderRowNumber_ As Long
Private LastRowDataNumber_ As Long
Private TargetHeaderTexts_ As Object
Private TargetColNumbers_ As Object

Public Property Set ws(arg1 As Worksheet)
    Set Ws_ = arg1
End Property

Public Property Let HeaderRowNumber(arg1 As Long)
    If Ws_ Is Nothing Then
        Err.Raise 1001, "HeaderRowNumber", "検索対象シートが設定されていません。"
    End If
    
    If arg1 < 1 Or Ws_.Rows.Count < arg1 Then
        Err.Raise 1011, "HeaderRowNumber", "行番号が正しくありません"
    End If
    
    If HeaderRowNumber_ <> -1 Then
        Err.Raise 1012, "HeaderRowNumber", "既に設定済みです:" & HeaderRowNumber_
    End If
    
    HeaderRowNumber_ = arg1
End Property

Public Property Let FirstRowNumber(arg1 As Long)
    If Ws_ Is Nothing Then
        Err.Raise 1001, "FirstRowNumber", "検索対象シートが設定されていません。"
    End If
    
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    If arg1 < 2 Or Ws_.Rows.Count < arg1 Then
        Err.Raise 1011, "FirstRowNumber", "行番号が正しくありません"
    End If
    
    If FirstRowNumber_ <> -1 Then
        Err.Raise 1012, "FirstRowNumber", "既に設定済みです:" & FirstRowNumber_
    End If
    
    FirstRowNumber_ = arg1
End Property

Public Property Let LastRowNumber(arg1 As Long)
    If Ws_ Is Nothing Then
        Err.Raise 1001, "LastRowNumber", "検索対象シートが設定されていません。"
    End If
    
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    If arg1 < 1 Or Ws_.Rows.Count < arg1 Then
        Err.Raise 1011, "LastRowNumber", "行番号が正しくありません"
    End If
    
    If LastRowNumber_ <> -1 Then
        Err.Raise 1012, "LastRowNumber", "既に設定済みです:" & LastRowNumber_
    End If
    
    LastRowNumber_ = arg1
End Property


Public Property Set TargetHeadersRange(arg1 As Range)
    Dim r As Range
    Dim rowNumber As Long
    
    If arg1.Rows.Count <> 1 Then
        Err.Raise 1031, "TargetHeadersRange", "見出しセル番地は1行で指定してください"
    End If
        
    '見出し行番号を設定
    rowNumber = arg1.Row
    
    If HeaderRowNumber_ = -1 Then
        HeaderRowNumber = rowNumber
    ElseIf HeaderRowNumber_ <> rowNumber Then
        Err.Raise 1032, "TargetHeadersRange", _
                        "見出しは同じ行を指定してください" & vbCrLf & _
                        "設定済行:" & HeaderRowNumber_ & vbCrLf & _
                        "指定行:" & rowNumber
    End If

    For Each r In arg1
        TargetColumnNumber = r.Column
    Next r

End Property

Public Property Let TargetColumnLetter(arg1 As String)
    On Error Resume Next
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    TargetColumnNumber = Ws_.Columns(arg1).Column

    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise 1026, "TargetColumnLetter", "列記号はA:" & Split(Ws_.Cells(1, Ws_.Columns.Count).Address, "$")(1) & "です"
    End If
End Property

Public Property Let TargetColumnNumber(arg1 As Long)
    If arg1 < 1 Or Ws_.Columns.Count < arg1 Then
        Err.Raise 1021, "TargetColumnNumber", "列番号は1~" & Ws_.Columns.Count & "です"
    End If

    If Not TargetColNumbers_.Exists(CLng(arg1)) Then
        TargetColNumbers_.Add CLng(arg1), ""
    End If
End Property

Public Property Let TargetHeaderText(arg1 As String)
    If Not TargetHeaderTexts_.Exists(arg1) Then
        TargetHeaderTexts_.Add arg1, ""
    End If
End Property


Private Sub Class_Initialize()
    Set TargetColNumbers_ = CreateObject("Scripting.Dictionary")
    Set TargetHeaderTexts_ = CreateObject("Scripting.Dictionary")
    
    '未設定(-1)を設定
    FirstRowNumber_ = -1
    LastRowNumber_ = -1
    HeaderRowNumber_ = -1
    LastRowDataNumber_ = -1

End Sub

Private Sub getColumnNumberFromHeaderTexts()
    Dim testRange1 As Range
    Dim testRange2 As Range
    Dim key As Variant
    
    For Each key In TargetHeaderTexts_.Keys
        'タイトルが2つ以上ある場合は例外を投げるため2回検索
        Set testRange1 = Ws_.Rows(HeaderRowNumber_).Find(What:=key, LookIn:=xlValues, LookAt:=xlWhole)

        If testRange1 Is Nothing Then
            Debug.Print "指定されたタイトル文字列は見つかりませんでした:" & key
        Else
            Set testRange2 = Ws_.Rows(HeaderRowNumber_).FindNext(testRange1)

            If testRange1.Address = testRange2.Address Then
                '登録
                TargetColumnNumber = testRange1.EntireColumn.Column
            Else
                'エラー
                Err.Raise 1033, "getColumnNumberFromHeaderTexts", "指定されたタイトル文字列が2つ以上あります"
            End If
        End If

    Next key
    
    TargetHeaderTexts_.RemoveAll

End Sub

'データ最終行番号を取得 / get the data last row number
Public Function getLastDataRowNumber() As Long
    Dim currentLastRowDataNumber As Long
    Dim key As Variant
    
    '--------------初期確認--------------
    If Ws_ Is Nothing Then
        Err.Raise 1001, "getLastDataRowNumber", "検索対象シートが設定されていません。"
    End If

    ''未設定値に既定の値を設定(FirstRowNumber_, HeaderRowNumber_)
    If FirstRowNumber_ = -1 And HeaderRowNumber_ = -1 Then
        HeaderRowNumber = 1
        FirstRowNumber = 2
    ElseIf FirstRowNumber_ = -1 Then
        FirstRowNumber = HeaderRowNumber_ + 1
    ElseIf HeaderRowNumber_ = -1 Then
        HeaderRowNumber = FirstRowNumber_ - 1
    End If
    
    ''未設定値に既定の値を設定(LastRowNumber_)
    If LastRowNumber_ = -1 Then
        LastRowNumber = Ws_.Rows.Count
    End If

    ''行設定値の整合性確認(FirstRowNumber_, HeaderRowNumber_,LastRowNumber_)
    If HeaderRowNumber_ >= FirstRowNumber_ Then
        Err.Raise 1041, "getLastDataRowNumber", "見出し行 < 検索対象行(開始)に設定してください。" & vbCrLf & "見出し行:" & HeaderRowNumber_ & vbCrLf & "開始行:" & FirstRowNumber_
    End If

    If FirstRowNumber_ > LastRowNumber_ Then
        Err.Raise 1042, "getLastDataRowNumber", "検索対象行(開始) < 検索対象行(終了)に設定してください。" & vbCrLf & "開始行:" & FirstRowNumber_ & vbCrLf & "終了行:" & LastRowNumber_
    End If

    '見出し文字列から列番号に変換する
    Call getColumnNumberFromHeaderTexts
    
    If TargetColNumbers_.Count = 0 Then
        Err.Raise 1022, "getLastDataRowNumber", "検索対象列が設定されていません。"
    End If
    '--------------初期確認終了--------------

    '行番号を-1(データなし)に設定
    LastRowDataNumber_ = -1
    
    For Each key In TargetColNumbers_.Keys

        'データ範囲最終にデータがあった場合は最終行を設定してループを抜ける(これ以上検索不要)
        'exit for if data exist in the last row(no need to further search)
        If Ws_.Cells(LastRowNumber_, CLng(key)) <> "" Then
            LastRowDataNumber_ = LastRowNumber_
            Exit For
        End If
    
        currentLastRowDataNumber = Ws_.Cells(LastRowNumber_, CLng(key)).End(xlUp).Row
    
        '取得した行が検索開始行と同じときはセルに値があるか確認する
        'check the cell if the currentRow is equal to the first row
        If currentLastRowDataNumber = FirstRowNumber_ Then
            If Ws_.Cells(currentLastRowDataNumber, CLng(key)) = "" Then
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
    
    Next key
    
    getLastDataRowNumber = LastRowDataNumber_

End Function


'データ範囲のrangeオブジェクトを取得(行全体) / get data range object(entire rows)
Public Function getDataRows(Optional withHeader As Boolean = False) As Range
    
    LastRowDataNumber_ = getLastDataRowNumber
    
    If LastRowDataNumber_ = -1 Then
        'データがないときはnothingを返す / set nothing if data does not exist
        Set getDataRows = Nothing
    Else
        If withHeader Then
            Set getDataRows = Ws_.Rows(CStr(HeaderRowNumber_) & ":" & CStr(LastRowDataNumber_))
        Else
            Set getDataRows = Ws_.Rows(CStr(FirstRowNumber_) & ":" & CStr(LastRowDataNumber_))
        End If
    End If

End Function

'データ範囲のrangeオブジェクトを取得(指定行列) / get data range object(specified Rows and Columns)
Public Function getDataRange(Optional withHeader As Boolean = False, Optional firstColNumber As Long = -1, Optional lastColNumber As Long = -1) As Range
    
    LastRowDataNumber_ = getLastDataRowNumber
    
    If LastRowDataNumber_ = -1 Then
        'データがないときはnothingを返す / set nothing if data does not exist
        Set getDataRange = Nothing
    Else
        If firstColNumber = -1 Then
            firstColNumber = Application.WorksheetFunction.Min(TargetColNumbers_.Keys)
        End If
        
        If lastColNumber = -1 Then
            lastColNumber = Application.WorksheetFunction.Max(TargetColNumbers_.Keys)
        End If
        
        If withHeader Then
            Set getDataRange = Ws_.Range(Ws_.Cells(HeaderRowNumber_, firstColNumber), Ws_.Cells(LastRowDataNumber_, lastColNumber))
        Else
            Set getDataRange = Ws_.Range(Ws_.Cells(FirstRowNumber_, firstColNumber), Ws_.Cells(LastRowDataNumber_, lastColNumber))
        End If
    
    End If

End Function


''シート関連
'Err.Raise 1001, "", "検索対象シートが設定されていません。"
''行関連
'Err.Raise 1011, "", "行番号が正しくありません"
'Err.Raise 1012, "", "既に設定済みです:"
''列関連
'Err.Raise 1021, "", "列番号は1~" & Ws_.Columns.Count & "です"
'Err.Raise 1022, "", "検索対象列が設定されていません。"
'Err.Raise 1026, "", "列記号はA:" & Split(Ws_.Cells(1, Ws_.Columns.Count).Address, "$")(1) & "です"
''見出し関連
'Err.Raise 1031, "", "見出しセル番地は1行で指定してください"
'Err.Raise 1032, "", "見出しは同じ行を指定してください"
'Err.Raise 1033, "", "指定されたタイトル文字列が2つ以上あります"
''整合性関連
'Err.Raise 1041, "", "見出し行 < 検索対象行(開始)に設定してください。"
'Err.Raise 1042, "", "検索対象行(開始) < 検索対象行(終了)に設定してください。"


