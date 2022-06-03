Sub Cek_Nilai()
    Dim Nilai As Integer
    Dim Nilai_Mutu As String
    Dim Angka_Mutu As Integer
    Dim mutu As String
    
    Nilai = Range("C1")
    
    If Nilai >= 85 Then
        Nilai_Mutu = "A"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "4"
        Range("C3") = Angka_Mutu
        mutu = "Istimewa"
        Range("C4") = mutu
        
    ElseIf Nilai >= 80 Then
        Nilai_Mutu = "A-"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "3,7"
        Range("C3") = Angka_Mutu
        mutu = "Sangat Memuaskan"
        Range("C4") = mutu
        
    ElseIf Nilai >= 75 Then
        Nilai_Mutu = "B+"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "3,3"
        Range("C3") = Angka_Mutu
        mutu = "Memuaskan"
        Range("C4") = mutu
        
    ElseIf Nilai >= 70 Then
        Nilai_Mutu = "B"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "3,0"
        Range("C3") = Angka_Mutu
        mutu = "Sangat Baik"
        Range("C4") = mutu
        
    ElseIf Nilai >= 65 Then
        Nilai_Mutu = "B-"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "2,7"
        Range("C3") = Angka_Mutu
        mutu = "Baik"
        Range("C4") = mutu
        
    ElseIf Nilai >= 60 Then
        Nilai_Mutu = "C+"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "2,3"
        Range("C3") = Angka_Mutu
        mutu = "Cukup Baik"
        
    ElseIf Nilai >= 55 Then
        Nilai_Mutu = "C"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "2,0"
        Range("C3") = Angka_Mutu
        mutu = "Cukup"
        Range("C4") = mutu
    
    ElseIf Nilai >= 50 Then
        Nilai_Mutu = "C-"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "1,7"
        Range("C3") = Angka_Mutu
        mutu = "Kurang"
        Range("C4") = mutu
        
    ElseIf Nilai >= 45 Then
        Nilai_Mutu = "D"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "1,0"
        Range("C3") = Angka_Mutu
        mutu = "Sangat Kurang"
        Range("C4") = mutu
        
    ElseIf Nilai >= 1 Then
        Nilai_Mutu = "E"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "0,0"
        Range("C3") = Angka_Mutu
        mutu = "Gagal"
        Range("C4") = mutu
        
    
    Else
        Nilai_Mutu = "F"
        Range("C2") = Nilai_Mutu
        Angka_Mutu = "0,0"
        Range("C3") = Angka_Mutu
        mutu = "Tunda"
        Range("C4") = mutu
        
    End If

End Sub
