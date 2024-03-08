'https://stackoverflow.com/questions/21559775/vbscript-to-open-a-dialog-to-select-a-filepath

Dim ShE
Dim fullFilename

With CreateObject("WScript.Shell")
    Set ShE= .Exec("mshta.exe ""about:<input type=file id=FILE><script>FILE.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>""")
End With
fullFilename = ShE.StdOut.ReadLine
wscript.echo fullFilename