Function convertWideToHalfSpace(targetString As String) As String

    '全角空白を半角空白に変換するコード(個数は変わらず)
    convertWideToHalfSpace = Replace(targetString, ChrW(&H3000), Chr(32))

End Function