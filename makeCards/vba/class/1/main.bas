Option Explicit

Sub main()
    Dim objMC As clsMakeCard
    

    Set objMC = New clsMakeCard
    
    objMC.makeCard (2)

    Set objMC = Nothing

End Sub
