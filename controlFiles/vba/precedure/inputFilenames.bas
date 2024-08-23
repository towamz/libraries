Option Explicit

Sub inputFilenames()
    Dim GFN As getFilenames
    Dim filenames() As String
    Dim i As Long
    Dim firstRng As Range
  
    '前回取得したファイル名を削除
    Worksheets("ファイル名").Columns(1).Clear
    '警告を削除
    Worksheets("ファイル名").Columns(3).Clear
  

    Set firstRng = Worksheets("ファイル名").Range("A1")

    Set GFN = New getFilenames
    GFN.setDirectory = Worksheets("設定").Range("B6").Value
    filenames = GFN.getFilenamesArray
    
    For i = LBound(filenames) To UBound(filenames)
    
        firstRng.Offset(i, 0).Value = filenames(i)
    
    Next

End Sub
