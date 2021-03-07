VERSION 5.00
Begin VB.Form TransPinjamBaru_form 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pinjaman Baru"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      Height          =   495
      Index           =   3
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2640
      Width           =   975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   2
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   960
      Width           =   1200
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Kali"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   2160
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "PERIODE ANGSURAN"
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
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL PINJAMAN"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS PINJAMAN"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA NASABAH"
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
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH"
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
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NO REK PINJAMAN "
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "TransPinjamBaru_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
'Text2.SetFocus
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Text1 = "" Or Text2 = "" Or Text3 = "" Then
        x = MsgBox("Data belum lengkap", vbOKOnly, "Validasi Data")
        If Text1 = "" Then
            Text1.SetFocus
        ElseIf Text2 = "" Then
            Text2.SetFocus
        ElseIf Text3 = "" Then
            Text3.SetFocus
        End If
    Else
        validasi_data
    End If
Case 1
    kosong
    limiter
    isi_combo1
    tutup_tombol
Case 2
    Me.Hide
    SubTransNasabah_form.Show
Case 3
    Call uncons
End Select
End Sub

Private Sub validasi_data()
Dim a As Boolean
With Data1.Recordset
If .BOF Then
    .AddNew
    !rekpinjaman = Text1
    !id_nasabah = lbl1(0)
    !jenis_pinjaman = Combo1
    !jumlah_pinjaman = Text2
    !jumlah_periode = Text3
    !sisa_pinjaman = Text2
    !angsuran_ke = 0
    !nominal_angsuran = 0
    !nominal_bagihasil = 0
    !tanggal = Date
    !jam = Time
    .Update
    Data1.Refresh
    update_akun
Else
    .MoveFirst
    a = False
    Do While Not .EOF
        If !rekpinjaman = Text1 Then
            x = MsgBox("No Rekening sudah ada... Silahkan masukan No Rekening yang lain", vbOKOnly, "Data Sudah Ada")
            Text1 = ""
            Text1.SetFocus
            a = True
            .MoveLast
        End If
        .MoveNext
    Loop
    Data1.Refresh
    If a = False Then
        .AddNew
        !rekpinjaman = Text1
        !id_nasabah = lbl1(0)
        !jenis_pinjaman = Combo1
        !jumlah_pinjaman = Text2
        !jumlah_periode = Text3
        !sisa_pinjaman = Text2
        !angsuran_ke = 0
        !nominal_angsuran = 0
        !nominal_bagihasil = 0
        !tanggal = Date
        !jam = Time
        .Update
        Data1.Refresh
        update_akun
    End If
End If
End With
End Sub

Private Sub update_akun()
With Data2.Recordset
    .MoveFirst
    Do While Not .EOF
        If !sandi_akun = "010101" Then 'sandi akun = kas
            .Edit
            !saldo_akun = !saldo_akun - Val(Text2)
            .Update
            Data2.Refresh
            .MoveLast
        End If
        .MoveNext
    Loop
    .MoveFirst
    Do While Not .EOF
        If Combo1 = "MUDHARABAH" Then
    '    Do While Not .EOF
            If !sandi_akun = "010401" Then 'sandi akun = pembiayaan mudharabah
                .Edit
                !saldo_akun = !saldo_akun + Val(Text2)
                .Update
                Data2.Refresh
                .MoveLast
            End If
        ElseIf Combo1 = "MUSYARAKAH" Then
            If !sandi_akun = "010501" Then 'sandi akun = pembiayaan musyarakah
                .Edit
                !saldo_akun = !saldo_akun + Val(Text2)
                .Update
                Data2.Refresh
                .MoveLast
            End If
        End If
        .MoveNext
    Loop
End With
x = MsgBox("TRANSAKSI TELAH DIPROSES..., LAKUKAN TRANSAKSI PINJAMAN LAGI?", vbOKCancel, "PROSES TRANSAKSI")
If x = vbCancel Then
    Me.Hide
    SubTransNasabah_form.Show
Else
    kosong
End If
End Sub

Private Sub kosong()
Text1 = ""
Text3 = ""
Text2 = ""
Text1.SetFocus
End Sub

Private Sub tutup_tombol()
    Command1(0).Enabled = False
    Command1(1).Enabled = False
End Sub

Private Sub buka_tombol()
    Command1(0).Enabled = True
    Command1(1).Enabled = True
End Sub

Private Sub Form_Activate()
kosong
limiter
isi_combo1
tutup_tombol
If Text1 <> "" And Text2 <> "" And Text3 <> "" Then
    buka_tombol
End If
End Sub

Private Sub limiter()
Text1.MaxLength = 10
Text2.MaxLength = 12
Text3.MaxLength = 2
End Sub

Private Sub isi_combo1()
    Combo1.Clear
    Combo1.AddItem ("MUDHARABAH")
    Combo1.AddItem ("MUSYARAKAH")
'    Combo1.AddItem ("MURABAHAH")
    Combo1.ListIndex = 0
End Sub

Private Sub Form_Load()
Call Buka_DBTransPinjamBaru
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        x = MsgBox("Text masih kosong, silahkan diisi dahulu...", vbOKOnly, "Text Kosong")
        Text1.SetFocus
    Else
        Data1.RecordSource = "select * from transpinjaman where rekpinjaman='" & Text1 & "'" ' and id_nasabah='" & lbl1(0) & "'"
        Data1.Refresh
        With Data1.Recordset
        If .BOF Then
            Combo1.SetFocus
        Else
            x = MsgBox("No Rekening sudah ada... Silahkan masukan No Rekening yang lain", vbOKOnly, "Data Sudah Ada")
            Text1 = ""
            Text1.SetFocus
        End If
        End With
    End If
End If
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text2 <> "" Then
            Text3.SetFocus
        Else
            x = MsgBox("Data masih kosong", vbOKOnly, "Data Kosong")
            Text2.SetFocus
        End If
    End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text3 <> "" Then
            buka_tombol
            Command1(0).SetFocus
        Else
            x = MsgBox("Data masih kosong", vbOKOnly, "Data Kosong")
            Text3.SetFocus
        End If
    End If
End Sub
