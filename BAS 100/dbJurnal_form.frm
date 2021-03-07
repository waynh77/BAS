VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form AkunTbl_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL AKUNTANSI"
   ClientHeight    =   6765
   ClientLeft      =   1005
   ClientTop       =   1380
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6000
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Text            =   "Combo3"
      Top             =   1920
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "dbJurnal_form.frx":0000
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "dbJurnal_form.frx":0014
      TabIndex        =   18
      Top             =   3720
      Width           =   6735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Last >|"
      Height          =   495
      Index           =   4
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Next >"
      Height          =   495
      Index           =   3
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CARI"
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "< Prev"
      Height          =   495
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "|< First"
      Height          =   495
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   4
      Left            =   5880
      TabIndex        =   9
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   495
      Index           =   3
      Left            =   4560
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS"
      Height          =   495
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EDIT"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "TAMBAH"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   600
      Width           =   5055
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1560
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SALDO NORMAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   1920
      Width           =   1470
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   6840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   6840
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "KLASIFIKASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SANDI AKUN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1155
   End
End
Attribute VB_Name = "AkunTbl_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
    isi_Klas
    Combo2.SetFocus
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
auto_sandi
Text1.SetFocus
SendKeys "{End}"
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(0).SetFocus
Else
    KeyAscii = 0
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "TAMBAH" Then
        Command1(0).Caption = "SIMPAN"
        Command1(1).Caption = "BATAL"
        kosong
        terang
        isi_jenis
        Combo1.SetFocus
        DBGrid1.Enabled = False
        Data1.Recordset.AddNew
        Command2(0).Enabled = False
        Command2(1).Enabled = False
        Command2(2).Enabled = False
        Command2(3).Enabled = False
        Command2(4).Enabled = False
        Command1(2).Enabled = False
        Command1(3).Enabled = False
        Command1(4).Enabled = False
    ElseIf Command1(0).Caption = "SIMPAN" Then
        If Text1 = "" Or Len(Text1) <> 6 Or Combo1 = "" Or Combo2 = "" Or Text2 = "" Then
            x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi Data")
            validasi_isi
        Else
            validasi_sandi
        End If
    ElseIf Command1(0).Caption = "UPDATE" Then
        If Text1 = "" Or Len(Text1) <> 6 Or Combo1 = "" Or Combo2 = "" Or Text2 = "" Then
            x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi Data")
            validasi_isi
        Else
            trans
            Data1.Recordset.Update
            Command1(0).Caption = "TAMBAH"
            Command1(1).Caption = "EDIT"
            Command1(2).Enabled = True
            Command1(3).Enabled = True
            Command1(4).Enabled = True
            Command2(0).Enabled = True
            Command2(1).Enabled = True
            Command2(2).Enabled = True
            Command2(3).Enabled = True
            Command2(4).Enabled = True
            Data1.Refresh
            isi
            burem
        End If
    End If
Case 1
    If Command1(1).Caption = "EDIT" Then
        Command1(1).Caption = "BATAL"
        terang
        Combo1.Enabled = False
        Combo2.Enabled = False
        Text1.Enabled = False
        If Not Data1.Recordset.BOF Then
            Command1(0).Caption = "UPDATE"
            Data1.Recordset.Edit
            Text2.SetFocus
        Else
            x = MsgBox("Data masih kosong silahkan di isi terlebih dahulu", vbOKOnly, "DATA KOSONG")
            Command1(0).Caption = "SIMPAN"
            kosong
            terang
            Data1.Recordset.AddNew
            DBGrid1.Enabled = False
        End If
        Command2(0).Enabled = False
        Command2(1).Enabled = False
        Command2(2).Enabled = False
        Command2(3).Enabled = False
        Command2(4).Enabled = False
        Command1(2).Enabled = False
        Command1(3).Enabled = False
        Command1(4).Enabled = False
    Else
        Command1(0).Caption = "TAMBAH"
        Command1(1).Caption = "EDIT"
        Command1(2).Enabled = True
        Command1(3).Enabled = True
        Command1(4).Enabled = True
        Command2(0).Enabled = True
        Command2(1).Enabled = True
        Command2(2).Enabled = True
        Command2(3).Enabled = True
        Command2(4).Enabled = True
        Data1.Refresh
        isi
        burem
        DBGrid1.Enabled = True
    End If
Case 2
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin menghapus data ini...???", vbOKCancel, "HAPUS DATA")
        If x = vbOK Then
            Data1.Recordset.Delete
            Data1.Refresh
            isi
        End If
    Else
    x = MsgBox("Data masih kosong silahkan di isi terlebih dahulu", vbOKOnly, "DATA KOSONG")
    Command1(0).Caption = "SIMPAN"
    Command1(1).Caption = "BATAL"
    kosong
    terang
    isi_jenis
    Combo1.SetFocus
    Data1.Recordset.AddNew
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Command1(4).Enabled = False
    End If
Case 3
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
Case 4
    Me.Hide
    main_form.Show
    main_form.Enabled = True
End Select
End Sub

Private Sub validasi_isi()
If Combo1 = "" Then
    Combo1.SetFocus
ElseIf Combo2 = "" Then
    Combo2.SetFocus
ElseIf Text1 = "" Or Len(Text1) <> 6 Then
    Text1.SetFocus
ElseIf Text2 = "" Then
    Text2.SetFocus
End If
End Sub

Private Sub Command2_Click(Index As Integer)
If Not Data1.Recordset.BOF Then
    Select Case Index
        Case 0
            Data1.Recordset.MoveFirst
            isi
        Case 1
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then
                x = MsgBox("Anda sudah di data pertama", vbOKOnly, "Data Pertama")
                Data1.Recordset.MoveFirst
            End If
            isi
        Case 2
            Call uncons
        Case 3
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then
                x = MsgBox("Anda sudah berada di data terakhir", vbOKOnly, "Data Terakhir")
                Data1.Recordset.MoveLast
            End If
            isi
        Case 4
            Data1.Recordset.MoveLast
            isi
    End Select
Else
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
End If

End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Command1(0).Caption = "SIMPAN" Then
    isi
End If
End Sub

Private Sub Form_Activate()
If Not Data1.Recordset.BOF Then
    isi
    isi_combo3
    burem
    Command2(0).Enabled = True
    Command2(1).Enabled = True
    Command2(2).Enabled = True
    Command2(3).Enabled = True
    Command2(4).Enabled = True
Else
    x = MsgBox("Data masih kosong silahkan di isi terlebih dahulu", vbOKOnly, "DATA KOSONG")
    Command1(0).Caption = "SIMPAN"
    Command1(1).Caption = "BATAL"
    kosong
    terang
    isi_jenis
    isi_combo3
    Combo1.SetFocus
    Data1.Recordset.AddNew
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
    Command1(2).Enabled = False
    Command1(3).Enabled = False
    Command1(4).Enabled = False
End If
End Sub

Private Sub isi_combo3()
Combo3.Clear
Combo3.AddItem ("DEBET")
Combo3.AddItem ("KREDIT")
End Sub
Private Sub Form_Load()
Call Buka_DBAkunTbl
kosong
limiter
End Sub

Private Sub kosong()
Text1 = ""
Combo1 = ""
Combo2 = ""
Text2 = ""
End Sub

Private Sub trans()
With Data1.Recordset
!sandi_akun = Text1
!nama_akun = Text2
!Jenis_akun = Combo1
!klasifikasi_akun = Combo2
If Combo3 = "DEBET" Then
    !saldo_normal = True
Else
    !saldo_normal = False
End If
End With
End Sub

Private Sub isi()
With Data1.Recordset
If Not .BOF Then
    Text1 = !sandi_akun
    Combo1 = !Jenis_akun
    Combo2 = !klasifikasi_akun
    Text2 = !nama_akun
    If !saldo_normal = True Then
        Combo3 = "DEBET"
    Else
        Combo3 = "KREDIT"
    End If
Else
    Text1 = ""
    Combo1 = ""
    Combo2 = ""
    Text2 = ""
    Combo3 = ""
End If
End With
End Sub

Private Sub limiter()
Text1.MaxLength = 6
Text2.MaxLength = 50
End Sub

Private Sub burem()
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
Combo3.Enabled = False
End Sub

Private Sub terang()
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
End Sub

Private Sub isi_jenis()
Combo1.Clear
Combo1.AddItem ("AKTIVA")
Combo1.AddItem ("KEWAJIBAN")
Combo1.AddItem ("INVESTASI TIDAK TERIKAT")
Combo1.AddItem ("EKUITAS")
Combo1.AddItem ("PENDAPATAN OPERASI UTAMA")
Combo1.AddItem ("HAK PIHAK KE-3 ATAS BG HSL INVES. TDK TERIKAT")
Combo1.AddItem ("PENDAPATAN OPERASI LAINNYA")
Combo1.AddItem ("BEBAN OPERASIONAL LAINNYA")
Combo1.AddItem ("PENDAPATAN NON-OPERASI")
Combo1.AddItem ("BEBAN NON-OPERASI")
End Sub

Private Sub isi_Klas()
Combo2.Clear
If Combo1 = "AKTIVA" Then
    Combo2.AddItem ("KAS")
    Combo2.AddItem ("TABUNGAN PADA BANK LAIN")
    Combo2.AddItem ("PIUTANG")
    Combo2.AddItem ("PEMBIAYAAN MUDHARABAH")
    Combo2.AddItem ("PEMBIAYAAN MUSYARAKAH")
    Combo2.AddItem ("PINJAMAN QARDH")
    Combo2.AddItem ("PENYISIHAN KERUGIAN PENGHAPUSBUKUAN AKTIVA")
    Combo2.AddItem ("PERSEDIAAN")
    Combo2.AddItem ("IJARAH")
    Combo2.AddItem ("AKTIVA ISTISHNA DALAM PENYELESAIAN")
    Combo2.AddItem ("AKTIVA TETAP DAN AKUMULASI PENYUSUTAN")
    Combo2.AddItem ("PIUTANG PENDAPATAN BAGI HASIL")
    Combo2.AddItem ("PIUTANG PENDAPATAN IJARAH")
    Combo2.AddItem ("AKTIVA LAINNYA")
ElseIf Combo1 = "KEWAJIBAN" Then
    Combo2.AddItem ("KEWAJIBAN SEGERA")
    Combo2.AddItem ("BAGI HASIL YANG BELUM DIBAGIKAN")
    Combo2.AddItem ("SIMPANAN WADIAH")
    Combo2.AddItem ("HUTANG")
    Combo2.AddItem ("KEWAJIBAN LAIN-LAIN")
    Combo2.AddItem ("HUTANG PAJAK")
    Combo2.AddItem ("PINJAMAN YANG DITERIMA")
ElseIf Combo1 = "INVESTASI TIDAK TERIKAT" Then
    Combo2.AddItem ("TABUNGAN MUDHARABAH")
    Combo2.AddItem ("DEPOSITO MUDHARABAH")
ElseIf Combo1 = "EKUITAS" Then
    Combo2.AddItem ("MODAL DISETOR")
    Combo2.AddItem ("SALDO LABA")
ElseIf Combo1 = "PENDAPATAN OPERASI UTAMA" Then
    Combo2.AddItem ("PENDAPATAN DARI JUAL BELI")
    Combo2.AddItem ("PENDAPATAN DARI BAGI HASIL")
    Combo2.AddItem ("PENDAPATAN DARI SEWA")
    Combo2.AddItem ("PENDAPATAN OPERASI UTAMA LAINNYA")
ElseIf Combo1 = "HAK PIHAK KE-3 ATAS BG HSL INVES. TDK TERIKAT" Then
    Combo2.AddItem ("BAGI HASIL TABUNGAN")
    Combo2.AddItem ("BAGI HASIL DEPOSITO")
    Combo2.AddItem ("BAGI HASIL PENEMPATAN DANA")
ElseIf Combo1 = "PENDAPATAN OPERASI LAINNYA" Then
    Combo2.AddItem ("PENDAPATAN FEE RAHN")
    Combo2.AddItem ("PENDAPATAN FEE JASA-JASA")
    Combo2.AddItem ("PENDAPATAN FEE INVESTASI TERIKAT")
    Combo2.AddItem ("PENDAPATAN FEE LAINNYA")
    Combo2.AddItem ("PENDAPATAN ADMINISTRASI")
ElseIf Combo1 = "BEBAN OPERASIONAL LAINNYA" Then
    Combo2.AddItem ("BEBAN BONUS WADIAH")
    Combo2.AddItem ("BEBAN PENYISIHAN KERUGIAN AKTIVA PRODUKTIF")
    Combo2.AddItem ("BEBAN PENYUSUTAN AKTIVA TETAP")
    Combo2.AddItem ("BEBAN PREMI DALAM RANGKA PENJAMINAN")
    Combo2.AddItem ("BEBAN SEWA")
    Combo2.AddItem ("BEBAN PROMOSI")
    Combo2.AddItem ("BEBAN TENAGA KERJA")
    Combo2.AddItem ("BEBAN ADMINISTRASI DAN UMUM")
ElseIf Combo1 = "PENDAPATAN NON-OPERASI" Then
    Combo2.AddItem ("PENDAPATAN NON-OPERASI")
ElseIf Combo1 = "BEBAN NON-OPERASI" Then
    Combo2.AddItem ("BEBAN NON-OPERASI")
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Text2.SetFocus
    End If
End Sub

Private Sub auto_sandi()
Dim a As String
Dim b As String
Dim urutan As String * 6
Dim hitung As Single
Select Case Combo1
Case "AKTIVA"
    a = "01"
Case "KEWAJIBAN"
    a = "02"
Case "INVESTASI TIDAK TERIKAT"
    a = "03"
Case "EKUITAS"
    a = "04"
Case "PENDAPATAN OPERASI UTAMA"
    a = "05"
Case "HAK PIHAK KE-3 ATAS BG HSL INVES. TDK TERIKAT"
    a = "06"
Case "PENDAPATAN OPERASI LAINNYA"
    a = "07"
Case "BEBAN OPERASIONAL LAINNYA"
    a = "08"
Case "PENDAPATAN NON-OPERASI"
    a = "09"
Case "BEBAN NON-OPERASI"
    a = "10"
End Select
Select Case Combo2
Case "KAS", "KEWAJIBAN SEGERA", "TABUNGAN MUDHARABAH", "MODAL DISETOR", "PENDAPATAN DARI JUAL BELI", "BAGI HASIL TABUNGAN", "PENDAPATAN FEE RAHN", "BEBAN BONUS WADIAH", "PENDAPATAN NON-OPERASI", "BEBAN NON-OPERASI"
    b = "01"
Case "TABUNGAN PADA BANK LAIN", "BAGI HASIL YANG BELUM DIBAGIKAN", "DEPOSITO MUDHARABAH", "SALDO LABA", "PENDAPATAN DARI BAGI HASIL", "BAGI HASIL DEPOSITO", "PENDAPATAN FEE JASA-JASA", "BEBAN PENYISIHAN KERUGIAN AKTIVA PRODUKTIF"
    b = "02"
Case "PIUTANG", "SIMPANAN WADIAH", "PENDAPATAN DARI SEWA", "BAGI HASIL PENEMPATAN DANA", "PENDAPATAN FEE INVESTASI TERIKAT", "BEBAN PENYUSUTAN AKTIVA TETAP"
    b = "03"
Case "PEMBIAYAAN MUDHARABAH", "HUTANG", "PENDAPATAN OPERASI UTAMA LAINNYA", "PENDAPATAN FEE LAINNYA", "BEBAN PREMI DALAM RANGKA PENJAMINAN"
    b = "04"
Case "PEMBIAYAAN MUSYARAKAH", "KEWAJIBAN LAIN-LAIN", "PENDAPATAN ADMINISTRASI", "BEBAN SEWA"
    b = "05"
Case "PINJAMAN QARDH", "HUTANG PAJAK", "BEBAN PROMOSI"
    b = "06"
Case "PENYISIHAN KERUGIAN PENGHAPUSBUKUAN AKTIVA", "PINJAMAN YANG DITERIMA", "BEBAN TENAGA KERJA"
    b = "07"
Case "PERSEDIAAN", "BEBAN ADMINISTRASI DAN UMUM"
    b = "08"
Case "IJARAH"
    b = "09"
Case "AKTIVA ISTISHNA DALAM PENYELESAIAN"
    b = "10"
Case "AKTIVA TETAP DAN AKUMULASI PENYUSUTAN"
    b = "11"
Case "PIUTANG PENDAPATAN BAGI HASIL"
    b = "12"
Case "PIUTANG PENDAPATAN IJARAH"
    b = "13"
Case "AKTIVA LAINNYA"
    b = "14"
End Select
Text1 = a & b
End Sub

Private Sub validasi_sandi()
Dim a As Byte
With Data1.Recordset
If Not .BOF Then
    a = 0
    .MoveFirst
    Do While Not .EOF
        If Text1 = !sandi_akun Then
            .MoveLast
'            .MovePrevious
            a = 1
        End If
        .MoveNext
        If .BOF Then
            .MoveLast
        End If
    Loop
End If
If a = 0 Then
    .AddNew
    trans
    .Update
    Command1(0).Caption = "TAMBAH"
    Command1(1).Caption = "EDIT"
    Command1(2).Enabled = True
    Command1(3).Enabled = True
    Command1(4).Enabled = True
    Command2(0).Enabled = True
    Command2(1).Enabled = True
    Command2(2).Enabled = True
    Command2(3).Enabled = True
    Command2(4).Enabled = True
    Data1.Refresh
    isi
    burem
    DBGrid1.Enabled = True
Else
    x = MsgBox("Sandi akun sudah ada, silahkan isi yang lain", vbOKOnly, "Data Sudah Ada")
    Text1.SetFocus
    SendKeys "{end}"
End If
End With
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo3.SetFocus
End If
End Sub
