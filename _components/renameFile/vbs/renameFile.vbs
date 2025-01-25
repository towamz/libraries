Option Explicit
Dim FSO
Dim FILENAME_BASE

FILENAME_BASE = "testFile.txt"

Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")


Function getInputString()
    Dim tmpStr, loopFlg

    loopFlg = False
    Do
        tmpStr = InputBox("入力してください。")
        If tmpStr = "" Then
            If MsgBox ("入力されませんでした。再入力しますか", 4 + 32, "MsgBox の例") = 6 Then
                loopFlg = False
            Else
                tmpStr = "未指定"
                loopFlg = True            
            End If
        Else
            loopFlg = True            
        End If
    Loop Until loopFlg

    getInputString = tmpStr
End Function


Function getDateString()
    Dim tmpStr

    tmpStr = Right("0" & Year(Now), 2) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) 
    tmpStr = tmpStr & "-"
    tmpStr = tmpStr & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2)
    getDateString = tmpStr
End Function

Function getFolderBase()
    getFolderBase = FSO.getParentFolderName(WScript.ScriptFullName)
End Function

Function getFilename(folderBase, filenameBase)
    Dim i, tmpFilename, tmpStr, tmpStr2

    tmpStr = getDateString()
    tmpStr2 = getInputString()
    tmpFilename = folderBase & "\" & tmpStr & "-" & tmpStr2 & "-" &filenameBase

    getFilename = tmpFilename
End Function

Sub renameFiles()
    Dim folderBase, filenameBase, filenameFrom, filenameTo

    filenameBase = FILENAME_BASE
    folderBase = getFolderBase()
    filenameFrom = folderBase & "\" & filenameBase
    filenameTo = getFilename(folderBase, filenameBase)

    If Not FSO.FileExists(filenameTo) Then
        Call FSO.CopyFile(filenameFrom, filenameTo, False) 
    End If
End Sub


call createFirenameFilesles()