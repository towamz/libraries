Option Explicit

Const STD_ASPECT_RATIO = 400 / 900

Sub insertPicturesPPT()
    Dim PPT As Object
    Dim PPF As Object
    Dim PPS As Object
    Dim PIC As Object
    Dim FSO As Object
    Dim baseDir As String
    Dim documentFilename, baseDocumentFilename As String
    Dim pictureFilename As String
    Dim firstRng As Range
    Dim i As Long
    Dim cuurentAcpectRatio As Double
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    '警告を削除
    Worksheets("ファイル名2").Columns(3).Clear

    baseDir = Worksheets("設定").Range("B6").Value
    Set firstRng = Worksheets("ファイル名2").Range("A1")
    
    baseDocumentFilename = Worksheets("設定").Range("B8").Value
    
    If baseDocumentFilename = "" Then
        baseDocumentFilename = "insertPictures"
    End If
    
    documentFilename = FSO.BuildPath(baseDir, baseDocumentFilename & ".pptx")
    
    
    If Dir(documentFilename) <> "" Then
        If MsgBox("既にファイルが生成されています。再作成しますか", vbYesNo) = vbNo Then
            Exit Sub
        End If
  
        i = 0
        For i = 0 To 100
            documentFilename = FSO.BuildPath(baseDir, baseDocumentFilename)
            documentFilename = documentFilename & "-" & Format(i, "00") & ".pptx"
            
            If Dir(documentFilename) = "" Then
                Exit For
            End If
            
        Next
    
        If i >= 100 Then
            MsgBox "ファイルが大量に生成されています。処理を中断します"
            Exit Sub
        End If
    
    End If
    
    


    'https://vba.9ol.net/excel-excel-vba-error-462-howto-240225/
    On Error Resume Next
    Set PPT = GetObject(, "PowerPoint.Application")
    If Err.Number <> 0 Then
        Set PPT = CreateObject("PowerPoint.Application")
    End If
    On Error GoTo 0
    
    PPT.Visible = True

    'PPTファイル新規追加 / create a new ppt file
    Set PPF = PPT.Presentations.Add
    
    
    i = 0
    Do Until firstRng.Offset(i, 0).Value = ""
        pictureFilename = FSO.BuildPath(baseDir, firstRng.Offset(i, 0).Value)
        
        'エクセルに入力されたファイルが存在しない場合は何もしないで次のセルを読み込む
        If Dir(pictureFilename) = "" Then
            firstRng.Offset(i, 2).Value = "ファイルが存在しません"
        Else
            'PPTファイルに新規スライド追加 / add a new slide to the PPT file
            'ppLayoutBlank=12
            Set PPS = PPF.Slides.Add(PPF.Slides.Count + 1, 12)
            PPS.Select
            
            'msoTextOrientationHorizontal=1
            PPS.Shapes.AddTextbox(Orientation:=1, Left:=10, Top:=10, Width:=900, Height:=90).TextFrame.TextRange.Text _
                                = firstRng.Offset(i, 0).Value & vbCrLf & firstRng.Offset(i, 1).Value

            'msoFalse=0,不正解
            'msoTrue=-1, はい
            Set PIC = PPS.Shapes.AddPicture(filename:=pictureFilename, LinkToFile:=0, SaveWithDocument:=-1, Left:=10, Top:=100)
            
            Debug.Print PIC.Height & "---" & PIC.Width
            
            cuurentAcpectRatio = PIC.Height / PIC.Width
            
            If cuurentAcpectRatio > STD_ASPECT_RATIO Then
                PIC.Height = 400
            Else
                PIC.Width = 900
            End If
            
            '1秒待って次の画像を挿入する / wait 1 sec to insert next pic
            Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 1)
        End If
        
        i = i + 1
    Loop
    
    
    PPF.SaveAs (documentFilename)


End Sub

