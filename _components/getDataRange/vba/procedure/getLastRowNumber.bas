'rowType
'0=データのある行番号
'1=データのない行番号
Private Function getLastRowNumber(ws As Worksheet, rgName As String, Optional rowType As Long = 0) As Long
    Dim columnLetter As String
    Dim rowNumber As Long
    
    columnLetter = Split(Range(rgName).Address, "$")(1)
    rowNumber = CLng(Split(Range(rgName).Address, "$")(2))
    
    
    
    If ws.Range(rgName) = "" Then
        getLastRowNumber = rowNumber
    Else
        getLastRowNumber = ws.Range(columnLetter & Rows.Count).End(xlUp).Row
    End If

    If getLastRowNumber < rowNumber Then
        getLastRowNumber = rowNumber
    End If

    'データのない行番号が指定されたときは+1する
    getLastRowNumber = getLastRowNumber + rowType

End Function