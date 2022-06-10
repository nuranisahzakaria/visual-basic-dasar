Sub Button1_loop()
    Dim i As Integer
    i = 1
    For i = 1 To 10
        Range("A" & i) = i
        
    Next i
    
End Sub
Sub Button2_Clear()
    Dim i As Integer
    i = 1
    For i = 1 To 10
        Range("A" & i) = Clear
        
    Next i
End Sub
Sub Button3_Do_While()
    Dim i As Integer
    i = 1
    Do While i <= 10
        Range("A" & i) = i
        i = i + 1
    Loop
End Sub
