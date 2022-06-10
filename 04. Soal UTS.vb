Sub Button1_Hitung()
    Dim nama_barang As String
    Dim harga_barang As Integer
    Dim jumlah_beli As Integer
    Dim diskon As Integer
    jumlah_diskon As Integer
    Dim jumlah_bayar As Integer
    
    nama_barang = Range("C3")
    jumlah_beli = Range("C5")
    
    If nama_barang = "Chiki" Then
        harga_barang = 1000
        Range("D4") = harga_barang
        
        If jumlah_beli = 10 Then
            Range("C6") = 10
        ElseIf jumlah_beli = 5 Then
            Range("C6") = 5
        End If
        
        jumlah_diskon = harga_barang * jumlah_beli * 10 / 100
        Range("C7") = jumlah_diskon
        
    End If
    
    
End Sub




    If jumlah_beli = 10 Then
        diskon = harga_barang * 10 / 100
    ElseIf jumlah_beli = 5 Then
        diskon = harga_barang * 5 / 100
    End If
    Range("D6") = diskon
    
