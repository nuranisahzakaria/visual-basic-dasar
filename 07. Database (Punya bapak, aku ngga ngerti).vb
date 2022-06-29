Dim MyWorkbook As Workbook
Dim Form As Worksheet, DataBase As Worksheet

Sub tombol_cek()
    Set MyWorkbook = ActiveWorkbook
    Set Form = MyWorkbook.Sheets(1)
    Set DataBase = MyWorkbook.Sheets(2)
    
    Form.Range("D2") = ""
            
    If Form.Range("B2") = "" Then
        MsgBox ("Untuk Mengecek Data, Isi NIM terlebih Dahulu...!!")
        Form.Range("B2").Select
        Exit Sub
    ElseIf Len(Form.Range("B2")) <> 9 Then
        MsgBox ("Panjang Karakter harus 9 Karakter, Panjang Karakter yang anda masukkan: " & Len(Form.Range("B2")))
        Form.Range("B2").Select
        Application.SendKeys "{F2}"
        Exit Sub
    End If
    
    For i = 1 To 1000
        If DataBase.Range("B" & i) = "" Then
            Form.Range("D2") = i
            Form.Range("B3").Select
            Exit For
        ElseIf DataBase.Range("B" & i) = Form.Range("B2") Then
            Form.Range("D2") = i
            Form.Range("B3") = DataBase.Range("C" & i)
            Form.Range("B4") = DataBase.Range("D" & i)
            Form.Range("B3").Select
            Exit For
        End If
    Next i
End Sub
Sub tombol_clear()
    Set MyWorkbook = ActiveWorkbook
    Set Form = MyWorkbook.Sheets(1)
    Set DataBase = MyWorkbook.Sheets(2)
    
    Form.Range("B2:B4") = ""
    Form.Range("D2") = ""
    Form.Range("B2").Select
End Sub
Sub tombol_simpan()
    Set MyWorkbook = ActiveWorkbook
    Set Form = MyWorkbook.Sheets(1)
    Set DataBase = MyWorkbook.Sheets(2)
    
    If Form.Range("D2") = "" Then
        MsgBox ("Cek data terlebih dahulu...")
        Exit Sub
    End If
    
    If Form.Range("B2") = "" Then
        MsgBox ("Untuk Mengecek Data, Isi NIM terlebih Dahulu...!!")
        Form.Range("B2").Select
        Exit Sub
    ElseIf Len(Form.Range("B2")) <> 9 Then
        MsgBox ("Panjang Karakter harus 9 Karakter, Panjang Karakter yang anda masukkan: " & Len(Form.Range("B2")))
        Form.Range("B2").Select
        Application.SendKeys "{F2}"
        Exit Sub
    End If
    
    
    If Form.Range("B2") = "" Or Form.Range("B3") = "" Or Form.Range("B4") = "" Then
        MsgBox ("Form Isian Tidak Boleh Kosong...!!!!")
        Exit Sub
    Else
        DataBase.Range("B" & Form.Range("D2")) = Form.Range("B2")
        DataBase.Range("C" & Form.Range("D2")) = Form.Range("B3")
        DataBase.Range("D" & Form.Range("D2")) = Form.Range("B4")
        MsgBox ("Data Sudah Berhasil Disimpan..")
        tombol_clear
    End If
End Sub
Sub tombol_hapus()
    Set MyWorkbook = ActiveWorkbook
    Set Form = MyWorkbook.Sheets(1)
    Set DataBase = MyWorkbook.Sheets(2)
    
    If Form.Range("D2") = "" Then
        MsgBox ("Tidak ada data yang akan dihapus...")
        Exit Sub
    ElseIf Form.Range("D2") = 2 And DataBase.Range("B3") = "" Then
        DataBase.Range("B2") = ""
        DataBase.Range("C2") = ""
        DataBase.Range("D2") = ""
        MsgBox ("Data Sudah dihapus...")
        tombol_clear
    Else
        DataBase.Rows(Form.Range("D2")).EntireRow.Delete
        MsgBox ("Data Sudah dihapus...")
        tombol_clear
    End If
End Sub
