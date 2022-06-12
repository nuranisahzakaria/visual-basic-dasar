Sub Button1_Pembelian()
    Dim Jumlah_Beli, Harga, Jumlah_Diskon, Jumlah_Bayar As Long
    Dim Diskon As String
    
    
    Harga = Range("D5")
    Jumlah_Beli = Range("D6")
    Total_harga = Harga * Jumlah_Beli
    
    
    If Jumlah_Beli = 10 Then
        Diskon = "10%"
        Range("D7") = Diskon
        Jumlah_Diskon = Harga * 10 / 100
        Range("D8") = Jumlah_Diskon
        Jumlah_Bayar = Total_harga - Jumlah_Diskon
        Range("D9") = Jumlah_Bayar
        
    ElseIf Jumlah_Beli = 5 Then
        Diskon = "5%"
        Range("D7") = Diskon
        Jumlah_Diskon = Harga * 5 / 100
        Range("D8") = Jumlah_Diskon
        Jumlah_Bayar = Total_harga - Jumlah_Diskon
        Range("D9") = Jumlah_Bayar
    Else
        Diskon = "0%"
        Range("D7") = Diskon
        Jumlah_Diskon = Harga * 0
        Range("D8") = Jumlah_Diskon
        Jumlah_Bayar = Total_harga - Jumlah_Diskon
        Range("D9") = Jumlah_Bayar
    End If
End Sub

' PR BUAT DIRI SENDIRI, HARUS CARI TAU
' 1. Kenapa pas di cell C8 ngga mau bener perhitungannya, tapi kalo diubah ke D8 atau K8
' langsung work
' 2. Cari tau berapa batasan integer di VBA (kalo di python kan ga dibatasi, krn int sudah 
' cukup untuk digunakan dan tidak perlu pake long), kalo VBA gimana? belum tau => 4 DIGIT RUPANYA
' 3. Aku harus latih logika pemograman deh, kurang banget ini. Masak ngga kepikiran buat 
' begini sih kemaren. Kalo belajar, insya Allah bisa kok.
' 4. Coba selesaikan soalnya pake cara dan jalan logika lain nis, atau setidaknya coba 
' ketik ulang kodingnya tanpa liat script
