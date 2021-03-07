VERSION 5.00
Begin VB.Form TransTabungan_form 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABUNGAN"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CETAK"
      Height          =   495
      Index           =   3
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000006&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   2520
      TabIndex        =   14
      Text            =   "Combo2"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   2520
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   2520
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO SALDO"
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
      Left            =   3120
      TabIndex        =   21
      Top             =   2040
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "DEPOSITO MUDHARABAH "
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
      Left            =   360
      TabIndex        =   20
      Top             =   2040
      Width           =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO SALDO"
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
      Left            =   3120
      TabIndex        =   18
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO SALDO"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label2 
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
      Index           =   1
      Left            =   2520
      TabIndex        =   16
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH   "
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
      Left            =   2520
      TabIndex        =   15
      Top             =   120
      Width           =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL TRANSAKSI"
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
      Left            =   360
      TabIndex        =   8
      Top             =   3600
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS TRANSAKSI"
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
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS TABUNGAN"
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
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI BARU"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "WADIAH"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TABUNGAN MUDHARABAH "
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
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2400
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SALDO TABUNGAN"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1680
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH   "
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
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "TransTabungan_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Combo2.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
Else
KeyAscii = 0
End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim tb, sal As Single
Select Case Index
Case 0
If Text1 <> "" Then
    With Data2.Recordset
        .AddNew
        !id_nasabah = Label2(0)
        !jenis_tabungan = Combo1
        !jenis_transaksi = Combo2
        !nominal_transaksi = Text1
        !tanggal = Date
        !jam = Time
        .Update
        Data2.Refresh
    End With
    With Data1.Recordset
    .MoveFirst
    tb = 0
    tb = Text1
    Do While Not .EOF
        If !id_nasabah = Label2(0) And !jenis_tabungan = Combo1 Then
            sal = !saldo
            .Edit
            If Combo2 = "SETORAN" Then
                !saldo = sal + tb
                If Combo1 = "WADIAH" Then
                    Label2(2) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "020301" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun + tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                ElseIf Combo1 = "TABUNGAN MUDHARABAH" Then
                    Label2(3) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "030101" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun + tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                ElseIf Combo1 = "DEPOSITO MUDHARABAH" Then
                    Label2(4) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "030201" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun + tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                End If
                Data3.Recordset.MoveFirst
                Do While Not Data3.Recordset.EOF
                    If Data3.Recordset!sandi_akun = "010101" Then
                        Data3.Recordset.Edit
                        Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun + tb
                        Data3.Recordset.Update
                        Data3.Recordset.MoveLast
                    End If
                    Data3.Recordset.MoveNext
                Loop
                Data3.Refresh
            Else
                !saldo = sal - tb
                If Combo1 = "WADIAH" Then
                    Label2(2) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "020301" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun - tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                ElseIf Combo1 = "TABUNGAN MUDHARABAH" Then
                    Label2(3) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "030101" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun - tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                ElseIf Combo1 = "DEPOSITO MUDHARABAH" Then
                    Label2(4) = !saldo
                    Data3.Recordset.MoveFirst
                    Do While Not Data3.Recordset.EOF
                        If Data3.Recordset!sandi_akun = "030201" Then
                            Data3.Recordset.Edit
                            Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun - tb
                            Data3.Recordset.Update
                            Data3.Recordset.MoveLast
                        End If
                        Data3.Recordset.MoveNext
                    Loop
                    Data3.Refresh
                End If
                Data3.Recordset.MoveFirst
                Do While Not Data3.Recordset.EOF
                    If Data3.Recordset!sandi_akun = "010101" Then
                        Data3.Recordset.Edit
                        Data3.Recordset!saldo_akun = Data3.Recordset!saldo_akun - tb
                        Data3.Recordset.Update
                        Data3.Recordset.MoveLast
                    End If
                    Data3.Recordset.MoveNext
                Loop
                Data3.Refresh
            End If
            .Update
            .MoveLast
        End If
        .MoveNext
    Loop
    Data1.Refresh
    End With
    kosong
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Text1.SetFocus
Else
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Text1.SetFocus
End If
Case 1
    Command1(0).Enabled = False
    Command1(1).Enabled = False
    Text1 = ""
    Text1.SetFocus
Case 2
    Me.Hide
    SubTransNasabah_form.Show
Case 3
    Call uncons
End Select
End Sub

Private Sub Form_Activate()
Dim a As Boolean
kosong
Label2(2) = 0
Label2(3) = 0
Label2(4) = 0
limiter
isi_combo1
isi_combo2
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = False
Data3.RecordSource = "select * from dbsaldo_akun order by sandi_akun asc"
Data3.Refresh
With Data1.Recordset
If Not .BOF Then
    a = False
    .MoveFirst
    Do While Not .EOF
        If !id_nasabah = Label2(0) And !jenis_tabungan = "WADIAH" Then
            Label2(2) = !saldo
        ElseIf !id_nasabah = Label2(0) And !jenis_tabungan = "TABUNGAN MUDHARABAH" Then
            Label2(3) = !saldo
        ElseIf !id_nasabah = Label2(0) And !jenis_tabungan = "DEPOSITO MUDHARABAH" Then
            Label2(4) = !saldo
            a = True
        End If
        .MoveNext
    Loop
Else
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "WADIAH"
    !saldo = 0
    .Update
    Data1.Refresh
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "TABUNGAN MUDHARABAH"
    !saldo = 0
    .Update
    Data1.Refresh
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "DEPOSITO MUDHARABAH"
    !saldo = 0
    .Update
    Data1.Refresh
    a = True
End If
If a = False Then
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "WADIAH"
    !saldo = 0
    .Update
    Data1.Refresh
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "TABUNGAN MUDHARABAH"
    !saldo = 0
    .Update
    Data1.Refresh
    .AddNew
    !id_nasabah = Label2(0)
    !jenis_tabungan = "DEPOSITO MUDHARABAH"
    !saldo = 0
    .Update
    Data1.Refresh
End If
End With
Text1.SetFocus
End Sub

Private Sub kosong()
Combo1 = ""
Combo2 = ""
Text1 = ""
End Sub

Private Sub limiter()
Text1.MaxLength = 10
End Sub

Private Sub isi_combo1()
Combo1.Clear
Combo1.AddItem ("WADIAH")
Combo1.AddItem ("TABUNGAN MUDHARABAH")
Combo1.AddItem ("DEPOSITO MUDHARABAH")
Combo1.ListIndex = 0
End Sub

Private Sub isi_combo2()
Combo2.Clear
Combo2.AddItem ("SETORAN")
Combo2.AddItem ("PENARIKAN")
Combo2.ListIndex = 0
End Sub

Private Sub Form_Load()
Call Buka_DBTransTabungan
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Command1(0).Enabled = True
        Command1(1).Enabled = True
        Command1(0).SetFocus
    End If
End Sub
