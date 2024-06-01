Option Explicit

Private searchWorksheet As Worksheet
Private searchFirstRow As Long
Private searchLastRow As Long
Private searchColumns As Object
Private dataLastRow As Long

Public Property Set setSearchWorksheet(worksheetObj As Worksheet)

    Set searchWorksheet = worksheetObj

    dataLastRow = -1

End Property

Public Function setSearchWorksheetString(worksheetString As String, Optional workbookString As String = "")
    
    If workbookString = "" Then
        workbookString = ThisWorkbook.Name
    End If
    
    Set searchWorksheet = Workbooks(workbookString).Worksheets(worksheetString)

    dataLastRow = -1

End Function

'Propertyだとoptinalが設定できない / optional cannot be used in property
'Public Property Let setSearchWorksheetString(worksheetString As String, workbookString As String)
'
'    Set searchWorksheet = Workbooks(workbookString).Worksheets(worksheetString)
'
'End Property

Public Property Get getSearchWorksheet()
    
    getSearchWorksheet = searchWorksheet.Name

End Property


Public Property Let setSearchFirstRow(rowNumber As Long)
    Dim rowAddress As String

    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    rowAddress = Rows(rowNumber).Address
    
    If Err.Number = 0 Then
        searchFirstRow = rowNumber
    End If

    dataLastRow = -1

End Property

Public Property Let setSearchLastRow(rowNumber As Long)
    Dim rowAddress As String

    '引数が行番号として有効か判定する / judge the argument is valid as a row number
    On Error Resume Next
    rowAddress = Rows(rowNumber).Address
    
    If Err.Number = 0 Then
        searchLastRow = rowNumber
    End If

    dataLastRow = -1

End Property

Public Property Get getSearchFirstRow()
    getSearchFirstRow = searchFirstRow
End Property

Public Property Get getSearchLastRow()
    getSearchLastRow = searchLastRow
End Property

Public Property Let setSearchColumn(ColumnString As String)
    Dim columnNumber As Long
    
    '列(英字)を列(数値)に変更 / get row number from row alphabet
    columnNumber = Columns(ColumnString).Column
    
    If Not searchColumns.Exists(columnNumber) Then
        searchColumns.Add columnNumber, 0
    End If

    dataLastRow = -1

End Property


Public Property Get getSearchColumns()
    Dim columnsString As String
    Dim searchColumn As Variant

    For Each searchColumn In searchColumns
        '列(数値)を列(英字)に変更して返す / return row alphabet from row number
        columnsString = columnsString & Split(Columns(searchColumn).Address, "$")(2) & vbCrLf
    Next
    
    '最後に空白行があるので配列最後の要素を削除する
    columnsString = Left(columnsString, Len(columnsString) - 2)
    getSearchColumns = columnsString

End Property

Public Sub clearSearchColumns()

    '検索対象列保存用ディクショナリを破棄・再作成する / destroy and recreate dic
    Set searchColumns = Nothing
    Set searchColumns = CreateObject("Scripting.Dictionary")
    
    dataLastRow = -1

End Sub



Private Sub Class_Initialize()
    '検索対象のシートを設定(デフォルトではアクティブシート)
    'set target sheet ( default is the activesheet
    Set searchWorksheet = ActiveSheet

    '検索対象行を設定(デフォルトでは全行) / set target rows (default is all rows)
    searchFirstRow = 1
    searchLastRow = Rows.Count

    '検索対象列保存用ディクショナリ
    Set searchColumns = CreateObject("Scripting.Dictionary")
    
    'データ最終行を-1(データなし)で初期化 / set dataLastRow is -1 (deta does not exsit)
    dataLastRow = -1

End Sub

Private Sub Class_Terminate()
    Set searchColumns = Nothing
End Sub


'データ最終行番号を取得 / get the data last row number
Public Function getLastRow() As Long
    Dim searchColumn As Variant
    Dim currentLastRow As Long
    
    '--------------初期確認--------------
    If searchWorksheet Is Nothing Then
        Err.Raise 1001, , "検索対象シートが設定されていません。"
    End If
    
    
    If searchFirstRow > searchLastRow Then
        Err.Raise 1002, , "検索対象行(開始) < 検索対象行(終了)に設定してください。" & vbCrLf & "開始行:" & searchFirstRow & vbCrLf & "終了行:" & searchLastRow
    End If

    If searchColumns.Count = 0 Then
        Err.Raise 1003, , "検索対象列が設定されていません。"
    End If
    '--------------初期確認終了--------------
      
    
    '行番号を-1(データなし)に設定
    dataLastRow = -1
    
    For Each searchColumn In searchColumns
        
        'データ範囲最終にデータがあった場合は最終行を設定してループを抜ける(これ以上検索不要)
        'exit for if data exist in the last row(no need to further search)
        If searchWorksheet.Cells(searchLastRow, searchColumn) <> "" Then
            dataLastRow = searchLastRow
            Exit For
        End If
        
        currentLastRow = searchWorksheet.Cells(searchLastRow, searchColumn).End(xlUp).row

        '取得した行が検索開始行と同じときはセルに値があるか確認する
        'check the cell if the currentRow is equal to the first row
        If currentLastRow = searchFirstRow Then
            If searchWorksheet.Cells(currentLastRow, searchColumn) = "" Then
                currentLastRow = -1
            End If
        '取得した行が検索開始行より小さい場合はデータなしと判定する
        'judge no data if the currentRow is smaller than the first row
        ElseIf currentLastRow < searchFirstRow Then
                currentLastRow = -1
        End If

        '今回取得した最終行が今まで最終行より大きい場合は書き換え
        'overwrite if the currentRow is bigger than the previous Row
        If dataLastRow < currentLastRow Then
            
            dataLastRow = currentLastRow
        
        End If
    
    Next

    getLastRow = dataLastRow

End Function


'データ範囲のrangeオブジェクトを取得(行全体) / get data range object(entire rows)
Public Function getDataRows() As Range
    
    Call getLastRow
    
    If dataLastRow = -1 Then
        'データがないときはnothingを返す / set nothing if data does not exist
        Set getDataRows = Nothing
    Else
        Set getDataRows = searchWorksheet.Rows(searchFirstRow & ":" & dataLastRow)
    End If

End Function
