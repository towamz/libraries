'参照したサイト
'https://www.tipsfound.com/vba/18015


Sub main()

    Dim objCSVFile As Workbook
    Dim strFullFileName As String
    Dim strFileName As String
    
    
    If Not getFilenameByDialog(strFullFileName) Then
        Exit Sub
    End If

    strFileName = Mid(strFullFileName, InStrRev(strFullFileName, "\") + 1)
    
    Select Case MsgBox("選択したテキストファイルの文字コードを選択してください" & vbCrLf & "はい=sjis" & vbCrLf & "いいえ=utf8", vbYesNoCancel)
        Case vbYes
            'Set objCSVFile = Workbooks.OpenText(Filename:=strFileName, Origin:=932, Comma:=True)
            Call Workbooks.OpenText(Filename:=strFullFileName, Origin:=932, Comma:=True)
        Case vbNo
            'Set objCSVFile = Workbooks.OpenText(Filename:=strFileName, Origin:=65001, Comma:=True)
            Call Workbooks.OpenText(Filename:=strFullFileName, Origin:=65001, Comma:=True)
        Case vbCancel
            Exit Sub
    End Select
    
    
    
        
        
    Select Case MsgBox("保存する文字コードを選択してください" & vbCrLf & "はい=sjis" & vbCrLf & "いいえ=utf8", vbYesNoCancel)
        Case vbYes
            Workbooks(strFileName).SaveAs Filename:=strFullFileName & "cng.txt", FileFormat:=xlCSV
        Case vbNo
            Workbooks(strFileName).SaveAs Filename:=strFullFileName & "cng.txt", FileFormat:=xlCSVUTF8
        Case vbCancel
            '処理しない
    End Select
    
    Workbooks(strFileName & "cng.txt").Close

End Sub
