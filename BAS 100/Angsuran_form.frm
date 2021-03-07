VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form TransAngsuran_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANGSURAN PINJAMAN"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   5085
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Angsuran_form.frx":0000
      Height          =   1215
      Left            =   120
      OleObjectBlob   =   "Angsuran_form.frx":0014
      TabIndex        =   26
      Top             =   5760
      Width           =   4815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   3
      Left            =   3720
      TabIndex        =   7
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   495
      Index           =   2
      Left            =   2520
      TabIndex        =   6
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox text3 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "text3"
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox text2 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Text            =   "text2"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Sisa"
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
      Index           =   7
      Left            =   2640
      TabIndex        =   25
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Total"
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
      Left            =   2640
      TabIndex        =   24
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Angsuran"
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
      Left            =   2640
      TabIndex        =   23
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Periode"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   2400
      Width           =   1440
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Jumlah"
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
      Left            =   2640
      TabIndex        =   21
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Jenis"
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
      Left            =   2640
      TabIndex        =   20
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Nama"
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
      Left            =   2640
      TabIndex        =   19
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Auto Id Nasabah"
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
      Left            =   2640
      TabIndex        =   18
      Top             =   600
      Width           =   1800
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1575
      Left            =   120
      Top             =   3480
      Width           =   4815
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1695
      Left            =   120
      Top             =   1560
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1215
      Left            =   120
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SISA PINJAMAN Rp"
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
      Index           =   10
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   1920
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TOTAL ANGSURAN"
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
      Index           =   9
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL BAGI HASIL"
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
      Index           =   8
      Left            =   240
      TabIndex        =   15
      Top             =   4320
      Width           =   2160
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL ANGSURAN Rp"
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
      Index           =   7
      Left            =   240
      TabIndex        =   14
      Top             =   3960
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ANGSURAN KE-"
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
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JUMLAH PERIODE "
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
      Left            =   240
      TabIndex        =   12
      Top             =   2400
      Width           =   1800
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
      Index           =   4
      Left            =   240
      TabIndex        =   11
      Top             =   1680
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JUMLAH PINJAMAN"
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
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA NASABAH "
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
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   1560
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH "
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
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   1320
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "TransAngsuran_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    With Data1.Recordset
        .AddNew
        !rekpinjaman = Text1
        !id_nasabah = lbl1(0)
        !jenis_pinjaman = lbl1(2)
        !jumlah_pinjaman = lbl1(3)
        !jumlah_periode = lbl1(4)
        !sisa_pinjaman = Val(lbl1(7)) - Val(Text2)
        !angsuran_ke = lbl1(5)
        !nominal_angsuran = Val(Text2)
        !nominal_bagihasil = Val(Text3)
        !tanggal = Date
        !jam = Time
        .Update
        Data1.Refresh
        update_akun
    End With
Case 1
    kosong
    tutup_text
    tutup_tombol
Case 2
    Call uncons
Case 3
    Me.Hide
    SubTransNasabah_form.Show
End Select
End Sub

Private Sub update_akun()
With Data2.Recordset
    .MoveFirst
    Do While Not .EOF
        If !sandi_akun = "010101" Then 'sandi akun = kas
            .Edit
            !saldo_akun = !saldo_akun + Val(Text2) + Val(Text3)
            .Update
            Data2.Refresh
            .MoveLast
        End If
        .MoveNext
    Loop
    .MoveFirst
    Do While Not .EOF
        If lbl1(2) = "MUDHARABAH" Then
    '    Do While Not .EOF
            If !sandi_akun = "010401" Or !sandi_akun = "050202" Then 'sandi akun = pembiayaan mudharabah
                .Edit
                If !sandi_akun = "010401" Then
                    !saldo_akun = !saldo_akun - Val(Text2)
                Else
                    !saldo_akun = !saldo_akun + Val(Text3)
                End If
                .Update
            End If
        ElseIf lbl1(2) = "MUSYARAKAH" Then
            If !sandi_akun = "010501" Or !sandi_akun = "050201" Then 'sandi akun = pembiayaan musyarakah
                .Edit
                If !sandi_akun = "010501" Then
                    !saldo_akun = !saldo_akun - Val(Text2)
                Else
                    !saldo_akun = !saldo_akun + Val(Text3)
                End If
                .Update
            End If
        End If
        .MoveNext
    Loop
    Data2.Refresh
End With
x = MsgBox("TRANSAKSI TELAH DIPROSES..., LAKUKAN TRANSAKSI ANGSURAN LAGI?", vbOKCancel, "PROSES TRANSAKSI")
If x = vbCancel Then
    Me.Hide
    SubTransNasabah_form.Show
Else
    kosong
    tutup_text
    tutup_tombol
    Data1.RecordSource = "select * from transpinjaman where id_nasabah='" & lbl1(0) & "'"
    Data1.Refresh
End If
End Sub

Private Sub kosong()
Text1 = ""
Text2 = ""
Text3 = ""
lbl1(2) = ""
lbl1(3) = ""
lbl1(4) = ""
lbl1(5) = ""
lbl1(6) = ""
lbl1(7) = ""
Text1.Enabled = True
Text1.SetFocus
End Sub

Private Sub tutup_text()
Text2.Enabled = False
Text3.Enabled = False
End Sub

Private Sub buka_text()
Text2.Enabled = True
Text3.Enabled = True
End Sub

Private Sub tutup_tombol()
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(2).Enabled = False
End Sub

Private Sub buka_tombol()
Command1(0).Enabled = True
Command1(1).Enabled = True
Command1(2).Enabled = True
End Sub

Private Sub limiter()
Text1.MaxLength = 10
Text2.MaxLength = 12
Text3.MaxLength = 12
End Sub

Private Sub Form_Activate()
kosong
limiter
tutup_text
tutup_tombol
Data1.RecordSource = "select * from transpinjaman where id_nasabah='" & lbl1(0) & "'"
Data1.Refresh
End Sub

Private Sub Form_Load()
Call Buka_DBTransAngsuran
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Text1 = "" Then
        x = MsgBox("Text masih kosong...", vbOKOnly, "Text Kosong")
        Text1.SetFocus
    Else
        Data1.RecordSource = "select * from transpinjaman where rekpinjaman='" & Text1 & "' and id_nasabah='" & lbl1(0) & "'"
        Data1.Refresh
        If Data1.Recordset.BOF Then
            x = MsgBox("No rekening tidak diketemukan,silahkan masukan no rekening yang lainnya...", vbOKOnly, "No Rekening Tidak Ada")
            Text1 = ""
            Text1.SetFocus
        Else
            With Data1.Recordset
                .MoveLast
                lbl1(2) = !jenis_pinjaman
                lbl1(3) = !jumlah_pinjaman
                lbl1(4) = !jumlah_periode
                lbl1(5) = .RecordCount
                lbl1(7) = !sisa_pinjaman
                buka_text
                Text2.SetFocus
                Text3.Enabled = False
                Text1.Enabled = False
            End With
        End If
    End If
End If
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text2 = "" Then
            x = MsgBox("Text masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Text kosong")
            Text2.SetFocus
        Else
            Text2.Enabled = False
            Text3.Enabled = True
            Text3.SetFocus
        End If
    End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text2 = "" Then
            x = MsgBox("Text masih kosong, silahkan diisi terlebih dahulu...", vbOKOnly, "Text kosong")
            Text2.SetFocus
        Else
            buka_tombol
            Text3.Enabled = False
            Command1(0).SetFocus
            lbl1(6) = Val(Text2) + Val(Text3)
        End If
    End If
End Sub
