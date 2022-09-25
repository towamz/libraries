Sub execMarge()
    Dim ins As New margeClass

    If ins.readFileByDay Then
        MsgBox "処理が終了しました"
    Else
        MsgBox "処理を中断しました"
    End If

End Sub


Sub makeFile()

    Dim ins2 As New makeNewfileClass

    ins2.makeNewfile

    Set ins2 = Nothing

End Sub
