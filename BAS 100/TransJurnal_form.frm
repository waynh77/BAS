VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form TransJurnal_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Jurnal"
   ClientHeight    =   8490
   ClientLeft      =   285
   ClientTop       =   915
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   2
      Left            =   10200
      TabIndex        =   20
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "SIMPAN"
      Height          =   495
      Index           =   0
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "TransJurnal_form.frx":0000
      Height          =   4215
      Left            =   5760
      OleObjectBlob   =   "TransJurnal_form.frx":0014
      TabIndex        =   19
      Top             =   3000
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PROSES"
      Height          =   735
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   7680
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      TabIndex        =   2
      Text            =   "Text6"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "TransJurnal_form.frx":09F7
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      DataField       =   "nama_akun"
      DataSource      =   "Data1"
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      TabIndex        =   18
      Text            =   "Text4"
      Top             =   1320
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000007&
      DataField       =   "klasifikasi_akun"
      DataSource      =   "Data1"
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      TabIndex        =   17
      Text            =   "Text3"
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      DataField       =   "jenis_akun"
      DataSource      =   "Data1"
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   600
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      DataField       =   "sandi_akun"
      DataSource      =   "Data1"
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   7680
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "TransJurnal_form.frx":09FD
      Height          =   7335
      Left            =   120
      OleObjectBlob   =   "TransJurnal_form.frx":0A11
      TabIndex        =   7
      Top             =   600
      Width           =   5415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   5760
      X2              =   11760
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "DEBET/KREDIT "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   6
      Left            =   5760
      TabIndex        =   14
      Top             =   2520
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   5
      Left            =   5760
      TabIndex        =   13
      Top             =   2160
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "KETERANGAN "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   4
      Left            =   5760
      TabIndex        =   12
      Top             =   1800
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA AKUN "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   3
      Left            =   5760
      TabIndex        =   11
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "KLASIFIKASI"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   2
      Left            =   5760
      TabIndex        =   10
      Top             =   960
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   1
      Left            =   5760
      TabIndex        =   9
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SANDI AKUN "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   0
      Left            =   5760
      TabIndex        =   8
      Top             =   240
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SILAHKAN PILIH SANDI AKUN :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Index           =   17
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3240
   End
End
Attribute VB_Name = "TransJurnal_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click()
If Text5 = "" Or Text6 = "" Then
    If Text5 = "" Then
        x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi")
        Text5.SetFocus
    ElseIf Text6 = "" Then
        x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi")
        Text6.SetFocus
    End If
Else
    simpan_temp
End If
End Sub

Private Sub simpan_temp()
With Data3.Recordset
    .AddNew
    !tanggal = Format(Date, "dd/mm/yyyy")
    !jam = Format(Time, "hh:mm:ss")
    !sandi_akun = Text1
    !keterangan = Text5
    If Combo1 = "DEBET" Then
        !debet = Val(Text6)
        !kredit = 0
    Else
        !debet = 0
        !kredit = Val(Text6)
    End If
    .Update
End With
Data3.Refresh
Data2.RecordSource = "select transjurnal.sandi_akun,nama_akun,keterangan,debet,kredit from transjurnal,tabel_akun where transjurnal.sandi_akun=tabel_akun.sandi_akun"
Data2.Refresh
kosong
End Sub


Private Sub simpan_tetap()
Dim nilai_saldo, saldo As Single
With DBJurnal_form.Data1.Recordset
Data3.Recordset.MoveFirst
Do While Not Data3.Recordset.EOF
    .AddNew
    !tanggal = Data3.Recordset!tanggal
    !jam = Data3.Recordset!jam
    !sandi_akun = Data3.Recordset!sandi_akun
    !keterangan = Data3.Recordset!keterangan
    !debet = Data3.Recordset!debet
    !kredit = Data3.Recordset!kredit
    nilai_saldo = 0
    nilai_saldo = Data3.Recordset!debet - Data3.Recordset!kredit
                    With BukuBesar_form
                    .Data1.Recordset.MoveFirst
                    Do While Not .Data1.Recordset.EOF
                        saldo = 0
                        If Data3.Recordset!sandi_akun = .Data1.Recordset!sandi_akun Then
                            .Data1.Recordset.Edit
                            saldo = .Data1.Recordset!saldo_akun
                            .Data2.Recordset.MoveFirst
                            Do While Not .Data2.Recordset.EOF
                                If Data3.Recordset!sandi_akun = .Data2.Recordset!sandi_akun Then
                                    If .Data2.Recordset!saldo_normal = True Then
                                        BukuBesar_form.Data1.Recordset!saldo_akun = saldo + nilai_saldo
                                    Else
                                        BukuBesar_form.Data1.Recordset!saldo_akun = saldo - nilai_saldo
                                    End If
                                    .Data2.Recordset.MoveLast
                                Else
                                    .Data2.Recordset.MoveNext
                                End If
                            Loop
                            .Data1.Recordset.Update
                            .Data1.Recordset.MoveLast
                        Else
                            .Data1.Recordset.MoveNext
                        End If
                    Loop
                    .Data2.Refresh
                    BukuBesar_form.Data1.Refresh
                    End With
    .Update
    Data3.Recordset.Delete
    Data3.Recordset.MoveNext
Loop
End With
DBJurnal_form.Data1.Refresh
Data3.Refresh
Data2.RecordSource = "select transjurnal.sandi_akun,nama_akun,keterangan,debet,kredit from transjurnal,tabel_akun where transjurnal.sandi_akun=tabel_akun.sandi_akun"
Data2.Refresh
kosong
End Sub

Private Sub hitung_saldo()
Dim saldo As Single
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0
    validasi_data
Case 1
    kosong
    If Data3.Recordset.RecordCount <> 0 Then
        Data3.Recordset.MoveFirst
        Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
        Loop
    Data3.Refresh
    End If
    Data2.RecordSource = "select transjurnal.sandi_akun,nama_akun,keterangan,debet,kredit from transjurnal,tabel_akun"
    Data2.Refresh
Case 2
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub validasi_data()
If Data3.Recordset.BOF Then
    x = MsgBox("Data belum diisi!!! harap mengisi data terlebih dahulu", vbOKOnly, "Validasi Data")
    kosong
    Text5.SetFocus
Else
    validasi_nilai
End If
End Sub

Private Sub validasi_nilai()
Dim dbt, krd As Single
With Data3.Recordset
.MoveFirst
dbt = 0
krd = 0
Do While Not .EOF
    dbt = dbt + !debet
    krd = krd + !kredit
    .MoveNext
Loop
If dbt <> krd Then
    x = MsgBox("Data belum balance/seimbang Debet = " & dbt & " kredit = " & krd, vbOKOnly, "Validasi Balance")
Else
    x = MsgBox("Data sudah balance/seimbang Debet = " & dbt & " kredit = " & krd, vbOKOnly, "Validasi Balance")
    simpan_tetap
End If
End With
End Sub

Private Sub Form_Activate()
Data2.RecordSource = "select transjurnal.sandi_akun,nama_akun,keterangan,debet,kredit from transjurnal,tabel_akun"
Data2.Refresh
Data1.RecordSource = "select sandi_akun,nama_akun,Jenis_akun,klasifikasi_akun from tabel_akun order by sandi_akun asc"
Data1.Refresh
limiter
kosong
isi_combo
tutup
End Sub

Private Sub tutup()
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
End Sub

Private Sub isi_combo()
Combo1.Clear
Combo1.AddItem ("DEBET")
Combo1.AddItem ("KREDIT")
Combo1.ListIndex = 0
End Sub

Private Sub Form_Load()
Call Buka_DBTransJurnal
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub kosong()
Text5 = ""
Text6 = ""
Text5.SetFocus
End Sub

Private Sub limiter()
Text5.MaxLength = 50
Text6.MaxLength = 20
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Combo1.SetFocus
    End If
End Sub
