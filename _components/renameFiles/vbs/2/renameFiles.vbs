'指定フォルダにあるファイルを1,2,3・・・に名前変更する
'拡張子の指定があればその拡張子のみ、なければ指定フォルダのすべてのファイルが対象
'change filenames to 1,2,3... in a folder
'target extention can be specified, or all files if it is not specified
Option Explicit
Dim FSO
Const TARGET_FOLDER = "C:\Users\forwa\OneDrive - 東京通信大学\articles\2025\ubuntu"
Const TARGET_EXTENTION = "png"

Set FSO = Wscript.CreateObject("Scripting.FileSystemObject")

Public Function renameFiles()
    Dim FLDR, FILES, FILE
    Dim filenameTo
    Dim paddingStrings
    Dim digit,i

    If Msgbox("処理を実行しますか->" & vbCRLF & _
              "フォルダ:"  & TARGET_FOLDER & vbCRLF & _
              "拡張子:"  & TARGET_EXTENTION ,4) = 7 Then
        WScript.Quit()
    End If


    Set FLDR = FSO.GetFolder(TARGET_FOLDER)
    Set FILES = FLDR.Files

    digit = Len(FILES.Count) + 1
    If digit < 2 Then digit = 2

	For i = 0 to digit
		paddingStrings = paddingStrings & "0"
	next


    i = 0
    For Each FILE in FILES
        If FSO.GetExtensionName(FILE.Name) = TARGET_EXTENTION Or TARGET_EXTENTION = "" Then
            filenameTo = right(paddingStrings & i, digit) & "." & FSO.GetExtensionName(FILE.Name)
            If Not FSO.FileExists(FSO.BuildPath(TARGET_FOLDER, filenameTo)) Then
                If i < 3 Then
                    If Msgbox(FILE.Name & "->" & vbCRLF & filenameTo ,4) = 7 Then
                        WScript.Quit()
                    End If
                End If
                i = i + 1
                FILE.Name = filenameTo
            End If
        End If
    Next

    Msgbox "処理が終了しました"
End Function

call renameFiles()
