Attribute VB_Name = "BukaDB"
Option Explicit
Public lakuntbl, lneraca, lbukubesar, lrugilaba, lnasabah, ljurnal, lsaldoakun, data, rpt, bdjurnal, saldo_harian, dbsaldo_akun, tabel_akun, tabel_nasabah, tabungan_nasabah, transjurnal, transpinjaman, transtabungan As String

Private Sub Basic_DB()
    data = App.Path & "\dataBMT.mdb"
    rpt = App.Path & "" '\BAS 100\report"
    bdjurnal = "select * from bdjurnal"
    saldo_harian = "select * from saldo_harian"
    dbsaldo_akun = "select * from dbsaldo_akun order by sandi_akun asc"
    tabel_akun = "select * from tabel_akun order by sandi_akun asc"
    tabel_nasabah = "select * from tabel_nasabah"
    tabungan_nasabah = "select * from tabungan_nasabah"
    transjurnal = "select * from transjurnal"
    transpinjaman = "select * from transpinjaman"
    transtabungan = "select * from transtabungan"
    lakuntbl = rpt & "\tabel akun.rpt"
    lbukubesar = rpt & "\Buku besar.rpt"
    lnasabah = rpt & "\data nasabah.rpt"
    'E:\bmt project1\report\LAPORAN NERACA.rpt
    lneraca = rpt & "\laporan neraca2.rpt"
    lrugilaba = rpt & "\laporan rugi laba.rpt"
    ljurnal = rpt & "\data transaksi jurnal.rpt"
    lsaldoakun = rpt & "\saldo akun.rpt"
End Sub

Public Sub utama()
Basic_DB
With main_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data2.RecordSource = dbsaldo_akun
    .Data1.RecordSource = saldo_harian
    .CrystalReport1.ReportFileName = lneraca
    .CrystalReport2.ReportFileName = lrugilaba
End With
End Sub

Public Sub cetak_neraca()
Basic_DB
With CetakNeraca_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = saldo_harian
    .CrystalReport1.ReportFileName = lneraca
    .CrystalReport2.ReportFileName = lrugilaba
End With
End Sub

Public Sub cetak_jurnal()
Basic_DB
With CetakDBJurnal_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = bdjurnal
    .CrystalReport1.ReportFileName = ljurnal
End With
End Sub
Public Sub Buka_DBAkunTbl()
    Basic_DB
    AkunTbl_form.Data1.DatabaseName = data
    AkunTbl_form.Data1.RecordSource = tabel_akun
    AkunTbl_form.CrystalReport1.ReportFileName = lakuntbl
End Sub

Public Sub Buka_DBBukuBesar()
    Basic_DB
    BukuBesar_form.Data1.DatabaseName = data
    BukuBesar_form.Data2.DatabaseName = data
    BukuBesar_form.Data3.DatabaseName = data
    BukuBesar_form.Data2.RecordSource = tabel_akun
    BukuBesar_form.Data1.RecordSource = dbsaldo_akun
    BukuBesar_form.Data3.RecordSource = "select dbsaldo_akun.sandi_akun,nama_akun,saldo_akun,Jenis_akun,Klasifikasi_akun from dbsaldo_akun,tabel_akun where dbsaldo_akun.sandi_akun=tabel_akun.sandi_akun order by dbsaldo_akun.sandi_akun asc"
    BukuBesar_form.CrystalReport1.ReportFileName = lbukubesar
End Sub

Public Sub Buka_DBcari_nsb()
    Basic_DB
    With cari_nsb
    .Data1.DatabaseName = data
    .Data1.RecordSource = tabel_nasabah
    End With
End Sub

Public Sub Buka_DBJurnal()
    Basic_DB
    With DBJurnal_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = bdjurnal
    .CrystalReport1.ReportFileName = ljurnal
    End With
End Sub

Public Sub Buka_DBSaldoTabungan()
    Basic_DB
    With SaldoTabungan_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = tabungan_nasabah
    End With
End Sub

Public Sub Buka_DBTransNasabah()
    Basic_DB
    With DBTransNasabah_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data1.RecordSource = transtabungan
    .Data2.RecordSource = transpinjaman
    End With
End Sub

Public Sub Buka_DBNasabah()
    Basic_DB
    With Nasabah_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = tabel_nasabah
    .CrystalReport1.ReportFileName = lnasabah
    End With
End Sub

Public Sub Buka_DBsubTransNasabah()
    Basic_DB
    With SubTransNasabah_form
    .Data1.DatabaseName = data
    .Data1.RecordSource = tabel_nasabah
    End With
End Sub

Public Sub Buka_DBTransAngsuran()
    Basic_DB
    With TransAngsuran_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data1.RecordSource = transpinjaman
    .Data2.RecordSource = dbsaldo_akun
    End With
End Sub

Public Sub Buka_DBTransJurnal()
    Basic_DB
    With TransJurnal_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data3.DatabaseName = data
    .Data1.RecordSource = tabel_akun
    .Data2.RecordSource = transjurnal
    .Data3.RecordSource = transjurnal
    End With
End Sub

Public Sub Buka_DBTransPinjamBaru()
    Basic_DB
    With TransPinjamBaru_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data1.RecordSource = transpinjaman
    .Data2.RecordSource = dbsaldo_akun
    End With
End Sub

Public Sub Buka_DBTransTabungan()
    Basic_DB
    With TransTabungan_form
    .Data1.DatabaseName = data
    .Data2.DatabaseName = data
    .Data3.DatabaseName = data
    .Data1.RecordSource = tabungan_nasabah
    .Data2.RecordSource = transtabungan
    .Data3.RecordSource = dbsaldo_akun
    End With
End Sub

