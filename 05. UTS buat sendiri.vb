Sub Button1_Pembelian()
    Dim Jumlah_Beli, Harga, Jumlah_Diskon, Jumlah_Bayar As Long
    Dim Diskon As String

    Harga = Range("D5")
    Jumlah_Beli = Range("D6")
    Total_harga = Harga * Jumlah_Beli
    
    If Jumlah_Beli = 10 Then
        Diskon = "10%"
        Jumlah_Diskon = Total_harga * 10 / 100
    ElseIf Jumlah_Beli = 5 Then
        Diskon = "5%"
        Jumlah_Diskon = Total_harga * 5 / 100
    Else
        Diskon = "0%"
        Jumlah_Diskon = Total_harga * 0
    End If
    
    Range("D7") = Diskon
    Range("D8") = Jumlah_Diskon
    Jumlah_Bayar = Total_harga - Jumlah_Diskon
    Range("D9") = Jumlah_Bayar
End Sub


' Perbedaan dari koding sebelumnya yang punya kawan
' 1. Lebih sederhana dan singkat karena tidak menggunakan fungsi if untuk elemen yang tidak perlu
' 2. Diskon dihitung dari jumlah seluruh harga
