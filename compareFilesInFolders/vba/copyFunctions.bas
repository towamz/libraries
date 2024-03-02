Option Explicit

Sub copyFunctions()
    Dim lastRow As Long
    
    '-----Sheet1用関数-----
    '最終行取得
    lastRow = Sheets("Sheet1").Range("A" & Rows.Count).End(xlUp).Row
    
    'B列
    '=FIND("●", SUBSTITUTE(A1, "\", "●", LEN(A1) - LEN(SUBSTITUTE(A1, "\", ""))))
    'C列
    '=TRIM(MID(A1,B1+1,1000))
    'D列
    '=VLOOKUP(C1,Sheet2!C:C,1,FALSE)

    Sheets("Sheet1").Range("B2:D" & Rows.Count).Clear
    Sheets("Sheet1").Range("B1:D1").Copy Sheets("Sheet1").Range("B2:B" & lastRow)


    '-----Sheet2用関数-----
    '最終行取得
    lastRow = Sheets("Sheet2").Range("A" & Rows.Count).End(xlUp).Row

    'B列
    '=FIND("●", SUBSTITUTE(A1, "\", "●", LEN(A1) - LEN(SUBSTITUTE(A1, "\", ""))))
    'C列
    '=TRIM(MID(A1,B1+1,1000))

    Sheets("Sheet2").Range("B2:C" & Rows.Count).Clear
    Sheets("Sheet2").Range("B1:C1").Copy Sheets("Sheet2").Range("B2:B" & lastRow)
End Sub
