Attribute VB_Name = "modul_karyawan"
Public Function karyawan_isi()
With Karyawan_form
If Not .Data1.Recordset.BOF Then
    .Text1 = .Data1.Recordset!nik
    .Text2 = .Data1.Recordset!nama_karyawan
    .Text22 = .Data1.Recordset!jabatan
    .Text3 = .Data1.Recordset!no_ktp
    .Text4 = .Data1.Recordset!tempat_lahir
    .Text5(0) = Format(.Data1.Recordset!tanggal_lahir, "dd")
    .Text5(1) = Format(.Data1.Recordset!tanggal_lahir, "mm")
    .Text5(2) = Format(.Data1.Recordset!tanggal_lahir, "yyyy")
    .Text6 = .Data1.Recordset!agama
    .Text7 = .Data1.Recordset!kewarganegaraan
    .Text8 = .Data1.Recordset!status_nikah
    .Text9 = .Data1.Recordset!pendidikan_akhir
    .Text10 = .Data1.Recordset!jalan
    .Text11 = .Data1.Recordset!no_rumah
    .Text12 = .Data1.Recordset!rt
    .Text13 = .Data1.Recordset!rw
    .Text14 = .Data1.Recordset!kelurahan
    .Text15 = .Data1.Recordset!kecamatan
    .Text16 = .Data1.Recordset!kabupaten
    .Text17 = .Data1.Recordset!propinsi
    .Text18 = .Data1.Recordset!kode_pos
    .Text19 = .Data1.Recordset!telp_rumah
    .Text20 = .Data1.Recordset!hp
    .Text21 = .Data1.Recordset!email
    If .Data1.Recordset!kelamin = "LAKI-LAKI" Then
        .Option1(0).Value = True
    Else
        .Option1(1).Value = True
    End If
End If
End With
End Function

Public Function karyawan_kosong()
With Karyawan_form
'.Text1 = ""
.Text2 = ""
.Text22 = ""
.Text3 = ""
.Text4 = ""
.Text5(0) = ""
.Text5(1) = ""
.Text5(2) = ""
.Text6 = ""
.Text7 = ""
.Text8 = ""
.Text9 = ""
.Text10 = ""
.Text11 = ""
.Text12 = ""
.Text13 = ""
.Text14 = ""
.Text15 = ""
.Text16 = ""
.Text17 = ""
.Text18 = ""
.Text19 = ""
.Text20 = ""
.Text21 = ""
.Option1(0).Value = False
.Option1(1).Value = False
End With
End Function

Public Function karyawan_burem()
With Karyawan_form
.Text1.Enabled = False
.Text2.Enabled = False
.Text22.Enabled = False
.Text3.Enabled = False
.Text4.Enabled = False
.Text5(0).Enabled = False
.Text5(1).Enabled = False
.Text5(2).Enabled = False
.Text6.Enabled = False
.Text7.Enabled = False
.Text8.Enabled = False
.Text9.Enabled = False
.Text10.Enabled = False
.Text11.Enabled = False
.Text12.Enabled = False
.Text13.Enabled = False
.Text14.Enabled = False
.Text15.Enabled = False
.Text16.Enabled = False
.Text17.Enabled = False
.Text18.Enabled = False
.Text19.Enabled = False
.Text20.Enabled = False
.Text21.Enabled = False
.Option1(0).Enabled = False
.Option1(1).Enabled = False
End With
End Function

Public Function karyawan_terang()
With Karyawan_form
'.Text1.Enabled = True
.Text2.Enabled = True
.Text22.Enabled = True
.Text3.Enabled = True
.Text4.Enabled = True
.Text5(0).Enabled = True
.Text5(1).Enabled = True
.Text5(2).Enabled = True
.Text6.Enabled = True
.Text7.Enabled = True
.Text8.Enabled = True
.Text9.Enabled = True
.Text10.Enabled = True
.Text11.Enabled = True
.Text12.Enabled = True
.Text13.Enabled = True
.Text14.Enabled = True
.Text15.Enabled = True
.Text16.Enabled = True
.Text17.Enabled = True
.Text18.Enabled = True
.Text19.Enabled = True
.Text20.Enabled = True
.Text21.Enabled = True
.Option1(0).Enabled = True
.Option1(1).Enabled = True
End With
End Function

Public Function karyawan_simpan()
Dim lahir As Date
With Karyawan_form
    lahir = .Text5(1) & "/" & .Text5(0) & "/" & .Text5(2)
    .Data1.Recordset!nik = .Text1
    .Data1.Recordset!nama_karyawan = .Text2
    .Data1.Recordset!jabatan = .Text22
    .Data1.Recordset!no_ktp = .Text3
    .Data1.Recordset!tempat_lahir = .Text4
    .Data1.Recordset!tanggal_lahir = lahir
    .Data1.Recordset!agama = .Text6
    .Data1.Recordset!kewarganegaraan = .Text7
    .Data1.Recordset!status_nikah = .Text8
    .Data1.Recordset!pendidikan_akhir = .Text9
    .Data1.Recordset!jalan = .Text10
    .Data1.Recordset!no_rumah = .Text11
    .Data1.Recordset!rt = .Text12
    .Data1.Recordset!rw = .Text13
    .Data1.Recordset!kelurahan = .Text14
    .Data1.Recordset!kecamatan = .Text15
    .Data1.Recordset!kabupaten = .Text16
    .Data1.Recordset!propinsi = .Text17
    .Data1.Recordset!kode_pos = .Text18
    .Data1.Recordset!telp_rumah = .Text19
    .Data1.Recordset!hp = .Text20
    .Data1.Recordset!email = .Text21
    If .Option1(0).Value = True Then
        .Data1.Recordset!kelamin = "LAKI-LAKI"
    Else
        .Data1.Recordset!kelamin = "PEREMPUAN"
    End If
End With
End Function

Public Function karyawan_validasi1()
With Karyawan_form
'If .Text1 = "" Then
'    .Text1.SetFocus
If .Text2 = "" Then
    .Text2.SetFocus
ElseIf .Text22 = "" Then
    .Text22.SetFocus
ElseIf .Text3 = "" Then
    .Text3.SetFocus
ElseIf .Text4 = "" Then
    .Text4.SetFocus
ElseIf .Text5(0) = "" Then
    .Text5(0).SetFocus
ElseIf .Text5(1) = "" Then
    .Text5(1).SetFocus
ElseIf .Text5(2) = "" Then
    .Text5(2).SetFocus
ElseIf .Text6 = "" Then
    .Text6.SetFocus
ElseIf .Text7 = "" Then
    .Text7.SetFocus
ElseIf .Text8 = "" Then
    .Text8.SetFocus
ElseIf .Text9 = "" Then
    .Text9.SetFocus
ElseIf .Text10 = "" Then
    .Text10.SetFocus
ElseIf .Text11 = "" Then
    .Text11.SetFocus
ElseIf .Text12 = "" Then
    .Text12.SetFocus
ElseIf .Text13 = "" Then
    .Text13.SetFocus
ElseIf .Text14 = "" Then
    .Text14.SetFocus
ElseIf .Text15 = "" Then
    .Text15.SetFocus
ElseIf .Text16 = "" Then
    .Text16.SetFocus
ElseIf .Text17 = "" Then
    .Text17.SetFocus
ElseIf .Text18 = "" Then
    .Text18.SetFocus
ElseIf .Text19 = "" Then
    .Text19.SetFocus
ElseIf Text20 = "" Then
    .Text20.SetFocus
ElseIf .Text21 = "" Then
    .Text21.SetFocus
ElseIf .Option1(0).Value = False And .Option1(1).Value = False Then
    .Option1(0).Value = True
End If
End With
End Function

Public Function karyawan_auto()
Dim urutan As String * 7
Dim hitung As Single
With Karyawan_form.Data1.Recordset
    If .RecordCount = 0 Then
        urutan = "PGW" & "0001"
    Else
        .MoveLast
        If Val(Left(.Fields("nik"), 4)) <> "0000" Then
            urutan = "0000" & "0001"
        Else
            hitung = Val(Right(.Fields("nik"), 4)) + 1
            urutan = "PGW" & Right("0000" & hitung, 4)
        End If
    End If
    Karyawan_form.Text1.Text = urutan
End With
End Function



