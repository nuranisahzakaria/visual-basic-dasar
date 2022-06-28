''Pengecekan bilangan ganjil genap
Sub Button1_Cek()
    Dim angka As Integer
    
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        MsgBox ("Bilangan " & angka & " merupakan bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        MsgBox ("Bilangan " & angka & " merupakan bilangan genap")
    End If
    
End Sub

''Pengecekan bilangan prima