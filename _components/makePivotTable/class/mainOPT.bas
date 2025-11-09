Sub clsPivodTest()

    Dim OPT As New OperatePivodTable

    OPT.dataSheetName = "datasrc"
    OPT.pivodSheetName = "test"
    OPT.addPageFieldName ("地区")
    OPT.addColumnFieldName ("性別")
    OPT.addRowFieldName ("都道府県")
    Call OPT.addDataFieldName("金額", xlAverage)
    OPT.createPivodTable
    OPT.addFields

End Sub
