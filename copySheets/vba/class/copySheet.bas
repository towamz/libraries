Sub copySheet()
    Dim ins1 As New ClsCopySheet
    
    ins1.setDataSheetRename = "test用の名前"

    Call ins1.setDataSheetCheckKey("A1", "Date")
    Call ins1.setDataSheetCheckKey("B1", "ALL")
    Call ins1.setDataSheetCheckKey("C1", "Hokkaido")

    ins1.copySheet

End Sub