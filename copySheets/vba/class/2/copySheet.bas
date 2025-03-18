Sub copySheet()
    Dim ins1 As New ClsCopySheet
    
    ins1.OrigDialogFileFilter = "csv,*.csv"
    ins1.OrigFileName = "newly_confirmed_cases_daily.csv"
    ins1.OrigSheetName = "newly_confirmed_cases_daily"
    Call ins1.setOrigSheetCheckKey("A1", "Date")
    ins1.OrigSheetNameNew = "classでテスト"
    ins1.copySheet

End Sub