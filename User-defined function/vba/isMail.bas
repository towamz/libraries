Function isMail(str1)
    Dim re As New RegExp
    Dim i, j As Long
    With re
        .Global = True '検索文字列全体について検索する(True)
        .IgnoreCase = True '検索するときに大文字と小文字を区別しない(True)
        .Pattern = "^[a-zA-Z0-9_.+-]+@([a-zA-Z0-9][a-zA-Z0-9-]*[a-zA-Z0-9]*\.)+[a-zA-Z]{2,}$" 
        isMail = .test(str1)
    End With
End Function