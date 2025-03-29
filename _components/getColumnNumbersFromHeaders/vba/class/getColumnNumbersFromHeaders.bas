Option Explicit

Private Ws_ As Worksheet

'タイトル行の行番号を格納
Private TargetHeadersRowNumber_ As Long

'(0,i)=列文字列/列番号/見出し文字列
'(1,i)=  0:列番号(列文字列は列番号に変換して格納)
'        1:見出し文字列
Private TargetHeaders_() As String

'列文字列/列番号/見出し文字列を列番号に変換した結果を格納
Private ResultColumnNumbers_() As Long
Private IsAllowDuplicateHeaders_  As Boolean


Public Property Set Ws(arg1 As Worksheet)
    Set Ws_ = arg1
End Property

Public Property Let TargetHeadersRowNumber(arg1 As Long)
    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    Debug.Print "見出し行番号有効性確認:" & Rows(arg1).Address
    
    If Err.Number <> 0 Then
        Debug.Print Err.Number & Err.Description
        Err.Raise 1001, , "行番号が正しくありません"
    End If
    On Error GoTo 0
    
    TargetHeadersRowNumber_ = arg1

End Property

Public Property Get TargetHeadersRowNumber() As Long
    TargetHeadersRowNumber = TargetHeadersRowNumber_
End Property


'列文字(A,AAなど)を指定 / specify column letter
Public Property Let TargetColumnLetter(arg1 As String)
    Dim i As Long
    
    Debug.Print "列記号:" & arg1
    
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    For i = Ws_.Columns(arg1).Column To Ws_.Columns(arg1).Columns.Count + Ws_.Columns(arg1).Column - 1
        If (Not TargetHeaders_) = -1 Then
            ReDim TargetHeaders_(1, 0)
        Else
            ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
        End If
        
        Debug.Print "->列番号:" & i
        TargetHeaders_(0, UBound(TargetHeaders_, 2)) = i
        TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 0
    Next
End Property

'列数字(1,2など)を指定 / specify column number
Public Property Let TargetColumnNumber(arg1 As Long)
    If (Not TargetHeaders_) = -1 Then
        ReDim TargetHeaders_(1, 0)
    Else
        ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
    End If
    
    Debug.Print "列番号:" & arg1
    TargetHeaders_(0, UBound(TargetHeaders_, 2)) = arg1
    TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 0
End Property


Public Property Let TargetHeader(arg1 As String)
    If (Not TargetHeaders_) = -1 Then
        ReDim TargetHeaders_(1, 0)
    Else
        ReDim Preserve TargetHeaders_(1, UBound(TargetHeaders_, 2) + 1)
    End If
    
    Debug.Print "見出し文字列:" & arg1
    TargetHeaders_(0, UBound(TargetHeaders_, 2)) = arg1
    TargetHeaders_(1, UBound(TargetHeaders_, 2)) = 1
End Property


Public Property Let IsAllowDuplicateHeaders(arg1 As Boolean)
    IsAllowDuplicateHeaders_ = arg1
End Property

Public Property Get IsAllowDuplicateHeaders() As Boolean
    AllowDuplicateHeaders = IsAllowDuplicateHeaders_
End Property


Private Sub Class_Initialize()
    '見出し行(既定値:1) / headerRow (default is 1)
    TargetHeadersRowNumber_ = 1
    '同じ見出し文字列を許可するか(既定値:false)
    IsAllowDuplicateHeaders_ = False
End Sub

Public Function getColumnNumberFromHeader() As Long()
    Dim testRange As Range
    Dim firstAddress As String

    Dim i As Long
    
    If (Not TargetHeaders_) = -1 Then
        Err.Raise 1004, , "有効な検索対象列が設定されていません。"
    End If
    
    For i = LBound(TargetHeaders_, 2) To UBound(TargetHeaders_, 2)
        '見出し文字列で指定されている場合は検索する
        If TargetHeaders_(1, i) = 1 Then
        
            '1回目の検索を実施
            Set testRange = Ws_.Rows(TargetHeadersRowNumber_).Find(What:=TargetHeaders_(0, i), LookIn:=xlValues, LookAt:=xlWhole)
    
            If testRange Is Nothing Then
                Debug.Print "指定されたタイトル文字列は見つかりませんでした:" & TargetHeaders_(0, i)
            Else
                '最初のアドレスを保存する
                firstAddress = testRange.Address
                
                '配列の初期化・インデックスを+1
                If (Not ResultColumnNumbers_) = -1 Then
                    ReDim ResultColumnNumbers_(0)
                Else
                    ReDim Preserve ResultColumnNumbers_(UBound(ResultColumnNumbers_) + 1)
                End If
                
                '1回目の列番号を格納
                ResultColumnNumbers_(UBound(ResultColumnNumbers_)) = testRange.EntireColumn.Column
                Debug.Print "タイトル文字列:" & TargetHeaders_(0, i) & "->列番号:" & ResultColumnNumbers_(UBound(ResultColumnNumbers_))

            
                '2回目の検索を実施
                Set testRange = Ws_.Rows(TargetHeadersRowNumber_).FindNext(testRange)

                'おなじ見出し文字列を許可していない場合で２つ以上見つかった場合は例外を投げる
                If Not IsAllowDuplicateHeaders_ Then
                    If firstAddress <> testRange.Address Then
                        Err.Raise 1021, , "指定されたタイトル文字列が2つ以上あります"
                    End If
                End If

                Do Until firstAddress = testRange.Address
                    '2回目以降の列番号を格納
                    ReDim Preserve ResultColumnNumbers_(UBound(ResultColumnNumbers_) + 1)
                    ResultColumnNumbers_(UBound(ResultColumnNumbers_)) = testRange.EntireColumn.Column
                    Debug.Print "タイトル文字列:" & TargetHeaders_(0, i) & "->列番号:" & ResultColumnNumbers_(UBound(ResultColumnNumbers_))
                
                    '3回目以降の検索を実施
                    Set testRange = Ws_.Rows(TargetHeadersRowNumber_).FindNext(testRange)
                Loop
            End If
        Else
            '配列の初期化・インデックスを+1
            If (Not ResultColumnNumbers_) = -1 Then
                ReDim ResultColumnNumbers_(0)
            Else
                ReDim Preserve ResultColumnNumbers_(UBound(ResultColumnNumbers_) + 1)
            End If
            
            '列番号をそのまま格納
            ResultColumnNumbers_(UBound(ResultColumnNumbers_)) = CLng(TargetHeaders_(0, i))
            
        End If
    Next i
    
    If (Not ResultColumnNumbers_) = -1 Then
        Err.Raise 1004, , "有効な検索対象列が設定されていません。"
    End If
    
    getColumnNumberFromHeader = ResultColumnNumbers_
End Function
