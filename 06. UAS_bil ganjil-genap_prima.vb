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

'' REVISI 02 => BELUM WORK INI
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


    '' REVISI 03 => BELUM WORK INI
Sub Button1_Cek()
    '' PENGECEKAN GANJIL-GENAP
    Dim angka As Integer
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan genap")
    End If
    
    ''PENGECEKAN BILANGAN PRIMA
    Dim angka_prima As Integer
    angka_prima = Range("C6")
    pembagi = 2
    
    If angka_prima Mod pembagi <> 1 And angka_prima Mod pembagi <> 0 Then
        pembagi = pembagi + 1
    ElseIf pembagi = angka_prima Then
        Range("C10") = ("Dan bilangan tersebut adalah bilangan prima")
    Else
        Range("C10") = ("Dan bilangan tersebut bukanlah bilangan prima")
    End If
    
End Sub


    ''REVISI 04 => BELUM WORK INI
Sub Button1_Cek()
    '' PENGECEKAN GANJIL-GENAP
    Dim angka As Integer
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan genap")
    End If
    
    ''PENGECEKAN BILANGAN PRIMA
    Dim angka_prima As Integer
    angka_prima = Range("C6")
    
    For pembagi = 2 To 100
        If angka_prima Mod pembagi <> 1 And angka_prima Mod pembagi <> 0 Then
        ElseIf pembagi = angka_prima Then
            Range("C10") = ("Dan bilangan " & angka & " adalah bilangan prima")
        Else
            Range("C10") = ("Dan bilangan " & angka & " bukanlah bilangan prima")
        End If
    Next pembagi
    
End Sub

    ''REVISI 05 => UDAH WORK, TAPI HARUS UBAH NAMA VARIABEL DULU
Sub Button1_Cek()
    '' PENGECEKAN GANJIL-GENAP
    Dim angka As Integer
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan genap")
    End If
    
    ''PENGECEKAN BILANGAN PRIMA
    Dim divisors As Integer, number As Long, i As Long
    divisors = 0
    number = Range("C6")
    
    For i = 1 To number
        If number Mod i = 0 Then
        divisors = divisors + 1
        End If
    Next i
    
    If divisors = 2 Then
        Range("C10") = number & " adalah bilangan prima"
    Else
        Range("C10") = number & " bukanlah bilangan prima"
    End If
End Sub


    ''REVISI 06 => UDAH WORK, TAPI KURANG RAPI
Sub Button1_Cek()
    '' PENGECEKAN GANJIL-GENAP
    Dim angka As Long
    angka = Range("C6")
    
    If angka Mod 2 = 1 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan genap")
    End If
    
    ''PENGECEKAN BILANGAN PRIMA
    Dim pembagi As Integer, number As Long, i As Long
    pembagi = 0
    angka = Range("C6")
    
    For i = 1 To angka
        If angka Mod i = 0 Then
        pembagi = pembagi + 1
        End If
    Next i
    
    If pembagi = 2 Then
        Range("C10") = ("Dan bilangan " & angka & " adalah bilangan prima")
    Else
        Range("C10") = ("Dan bilangan " & angka & " bukanlah bilangan prima")
    End If
End Sub


    ''REVISI 07 => KUMPUL KE BAPAK
Sub Button1_Cek()
    ''PENDEKLARASIAN VARIABEL
    Dim angka, i As Long
    Dim pembagi As Integer
    
    angka = Range("C6")
    
    '' PENGECEKAN GANJIL-GENAP
    If angka Mod 2 = 1 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("C9") = ("Bilangan " & angka & " adalah bilangan genap")
    End If
    
    ''PENGECEKAN BILANGAN PRIMA
    pembagi = 0
    
    For i = 1 To angka
        If angka Mod i = 0 Then
            pembagi = pembagi + 1
        End If
    Next i
    
    If pembagi = 2 Then
        Range("C10") = ("Dan bilangan " & angka & " adalah bilangan prima")
    Else
        Range("C10") = ("Dan bilangan " & angka & " bukanlah bilangan prima")
    End If
End Sub


    ''REVISI 08 => SHARE KE KAWAN WKWKWK
Sub Button1_Cek()
    Dim angka As Integer
    angka = Range("C3")
    
    If angka Mod 2 = 1 Then
        Range("A7") = ("Angka " & angka & " adalah angka ganjil")
    ElseIf angka Mod 2 = 0 Then
        Range("A7") = ("Angka " & angka & " adalah angka genap")
    End If
    
    
    Dim divisors As Integer, number As Long, i As Long
    divisors = 0
    number = Range("C3")
    
    For i = 1 To number
        If number Mod i = 0 Then
        divisors = divisors + 1
        End If
    Next i
    
    If divisors = 2 Then
        Range("A8") = "Dan angka tersebut adalah angka prima"
    Else
        Range("A8") = "Dan angka tersebut bukanlah angka prima"
    End If
End Sub