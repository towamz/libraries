Option Explicit

Sub insertPictures()
    Dim WRD As Object
    Dim DOC As Object
    Dim FSO As Object
    Dim objIls As Object
    Dim baseDir As String
    Dim documentFilename, baseDocumentFilename As String
    Dim pictureFilename As String
    Dim firstRng As Range
    Dim i As Long
    
    Set FSO = CreateObject("Scripting.FileSystemObject")

    '警告を削除
    Worksheets("ファイル名2").Columns(3).Clear

    baseDir = Worksheets("設定").Range("B6").Value
    Set firstRng = Worksheets("ファイル名2").Range("A1")
    
    baseDocumentFilename = Worksheets("設定").Range("B8").Value
    
    If baseDocumentFilename = "" Then
        baseDocumentFilename = "insertPictures"
    End If
    
    documentFilename = FSO.BuildPath(baseDir, baseDocumentFilename & ".docx")
    
    
    If Dir(documentFilename) <> "" Then
        If MsgBox("既にファイルが生成されています。再作成しますか", vbYesNo) = vbNo Then
            Exit Sub
        End If
  
        i = 0
        For i = 0 To 100
            documentFilename = FSO.BuildPath(baseDir, baseDocumentFilename)
            documentFilename = documentFilename & "-" & Format(i, "00") & ".docx"
            
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
    Set WRD = GetObject(, "Word.Application")
    If Err.Number <> 0 Then
        Set WRD = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    WRD.Visible = True

    Set DOC = WRD.Documents.Add
    
    
    i = 0
    Do Until firstRng.Offset(i, 0).Value = ""
        pictureFilename = FSO.BuildPath(baseDir, firstRng.Offset(i, 0).Value)
        
        'エクセルに入力されたファイルが存在しない場合は何もしないで次のセルを読み込む
        If Dir(pictureFilename) = "" Then
            firstRng.Offset(i, 2).Value = "ファイルが存在しません"
        Else
            'wd定数が参照できないので数値で指定する
            DOC.Content.InsertAfter Text:=firstRng.Offset(i, 0).Value
            'WRD.Selection.EndKey Unit:=wdLine, Extend:=wdMove
            'WRD.Selection.InsertBreak Type:=wdLineBreak
            WRD.Selection.EndKey Unit:=5, Extend:=0
            WRD.Selection.InsertBreak Type:=6
            
            If firstRng.Offset(i, 1).Value <> "" Then
                'wdStory=6
                DOC.Content.InsertAfter Text:=firstRng.Offset(i, 1).Value
                WRD.Selection.EndKey Unit:=6, Extend:=0
                WRD.Selection.InsertBreak Type:=6
            End If
            
            
            Debug.Print pictureFilename ' & "-->" & Left(filename, InStrRev(filename, "-") - 1)
            
            Set objIls = DOC.InlineShapes.AddPicture(filename:=pictureFilename)
    
            'WRD.Selection.EndKey Unit:=wdStory
            'WRD.Selection.InsertBreak Type:=wdPageBreak
            WRD.Selection.EndKey Unit:=6
            WRD.Selection.InsertBreak Type:=7
            Application.Wait TimeSerial(Hour(Now()), Minute(Now()), Second(Now()) + 1)
        End If
        
        i = i + 1
    Loop
    
    
    DOC.SaveAs (documentFilename)


End Sub
