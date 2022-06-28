'' REVISI 01
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

'' REVISI 02
''Pengecekan bilangan prima
Sub Button1_Cek()
    Dim angka As Integer
    
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        MsgBox ("Bilangan " & angka & " merupakan bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        MsgBox ("Bilangan " & angka & " merupakan bilangan genap")
    End If
    
End Sub
Sub Button2_Prima()
    Dim angka_prima As Integer
    angka_prima = Range("C8")
    pembagi = 2
    
    If angka_prima / pembagi <> 1 And angka_prima / pembagi <> 0 Then
        pembagi = pembagi + 1
    ElseIf pembagi = angka_prima Then
        MsgBox ("Bilangan" & angka_prima & "adalah bilangan prima")
    Else
        MsgBox ("Bilangan" & angka_prima & "bukanlah bilangan prima")
    End If
End Sub