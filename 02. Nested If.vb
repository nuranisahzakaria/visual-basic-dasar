Sub Button1_CekNilai()
    Select Case Range("B1")
    Case 1
        Range("B7") = "Laptop"
    Case 2
        Range("B7") = "Android"
    Case 3
        Range("B7") = "Keyboard"
    End Select
End Sub
