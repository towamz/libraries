Sub execMarge()
    Dim MS As margeData
    
    Set MS = New margeData

    MS.setTemplateWorksheetName = "hinagata"
    MS.setDataWorksheetName = "結果"

    MS.setDirectory = "C:\sampleMacro\getDataRange\data"
    MS.setPattern = "*.xlsx"

    MS.setSearchFirstRow = 4
    MS.setSearchLastRow = 103
    MS.setSearchColumn = "C"
    MS.setSearchColumn = "E"


    MS.setWorkbookNameColumn = "A"
    MS.setSerialNumbersColumn = "B"

    MS.printSettings

    MS.margeData

End Sub