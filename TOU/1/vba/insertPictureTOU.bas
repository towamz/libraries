Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub insertPictureTOU()
    Dim FSO As Object
    Dim GFN As clsGetFilenames
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set GFN = New clsGetFilenames
    
    Dim currentLectureNumber
    Dim baseDirectory
    
    Dim filenamesArray As Variant
    Dim filename As Variant
    Dim fullfilename As String
    Dim lectureSublectureString As String

    baseDirectory = "C:\テクノロジーマーケティングⅠ\単位認定試験まとめ\"
    
    For currentLectureNumber = 1 To 8
        
        GFN.setDirectory = FSO.BuildPath(baseDirectory, currentLectureNumber)
        filenamesArray = GFN.getFilenamesArray

        For Each filename In filenamesArray
            lectureSublectureString = Left(filename, InStrRev(filename, "-") - 1)
            fullfilename = FSO.BuildPath(FSO.BuildPath(baseDirectory, currentLectureNumber), filename)
            
            ActiveDocument.Content.InsertAfter Text:=lectureSublectureString
            Selection.EndKey Unit:=wdLine
            Selection.InsertBreak Type:=wdLineBreak
            
            'Debug.Print fullfilename & "-->" & Left(filename, InStrRev(filename, "-") - 1)
            
            Set objIls = ActiveDocument.InlineShapes.AddPicture(filename:=fullfilename)

            Selection.EndKey Unit:=wdStory
            Selection.InsertBreak Type:=wdPageBreak
            
            Sleep 1000
        
        Next
        
        Call GFN.initVariables
    Next

End Sub


