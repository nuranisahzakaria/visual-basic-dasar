Sub bayar()
    Dim Jumlah_Beli, Harga, Jumlah_Diskon, Jumlah_Bayar As Long
    Dim Diskon As String
    
    
    Harga = Range("K5")
    Jumlah_Beli = Range("K6")
    Jumlah_Bayar = Harga * Jumlah_Beli
    
    
    If Jumlah_Beli = 10 Then
        Diskon = "10%"
        Range("K7") = Diskon
        Jumlah_Diskon = Harga * 10 / 100
        Range("K8") = Jumlah_Diskon
        Jumlah_Bayar = Jumlah_Bayar - Jumlah_Diskon
        Range("K9") = Jumlah_Bayar
        
    ElseIf Jumlah_Beli = 5 Then
        Diskon = "5%"
        Range("K7") = Diskon
        Jumlah_Diskon = Harga * 5 / 100
        Range("K8") = Jumlah_Diskon
        Jumlah_Bayar = Jumlah_Bayar - Jumlah_Diskon
        Range("K9") = Jumlah_Bayar
    Else
        Diskon = "0%"
        Range("K7") = Diskon
        Jumlah_Diskon = Harga * 0
        Range("K8") = Jumlah_Diskon
        Jumlah_Bayar = Jumlah_Bayar - Jumlah_Diskon
        Range("K9") = Jumlah_Bayar
    End If
End Sub
