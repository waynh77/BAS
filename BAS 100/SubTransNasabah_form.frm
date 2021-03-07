VERSION 5.00
Begin VB.Form SubTransNasabah_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub Transaksi Nasabah"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "ANGSURAN PINJAMAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "PINJAMAN BARU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "TABUNGAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2640
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   120
      X2              =   4080
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NSB"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   4080
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SILAHKAN PILIH JENIS TRANSAKSI :"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO NAMA "
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
      Left            =   2160
      TabIndex        =   7
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA NASABAH :"
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
      TabIndex        =   6
      Top             =   600
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH   :"
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
      Width           =   1680
   End
End
Attribute VB_Name = "SubTransNasabah_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Me.Hide
    TransTabungan_form.Show
    With TransTabungan_form
        .Label2(0) = "NSB" & Text1
        .Label2(1) = Label1(2)
    End With
Case 1
    Me.Hide
    TransPinjamBaru_form.Show
    With TransPinjamBaru_form
        .lbl1(0) = "NSB" & Text1
        .lbl1(1) = Label1(2)
    End With
Case 2
    Me.Hide
    main_form.Enabled = True
    main_form.Show
'End
Case 3
    Me.Hide
    TransAngsuran_form.Show
    With TransAngsuran_form
        .lbl1(0) = "NSB" & Text1
        .lbl1(1) = Label1(2)
    End With
End Select
End Sub

Private Sub Form_Activate()
Text1 = ""
Text1.SetFocus
Text1.MaxLength = 5
Label1(2) = "AUTO NAMA"
Command1(0).Enabled = False
Command1(1).Enabled = False
Command1(3).Enabled = False
End Sub

Private Sub Form_Load()
Call Buka_DBsubTransNasabah
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        cari_data
    End If
End Sub

Private Sub cari_data()
Dim a As Boolean
With Data1.Recordset
.MoveFirst
a = False
Do While Not .EOF
    If !id_nasabah = Label1(4) & Text1.Text Then
        Label1(2) = !nama_nasabah
        a = True
        .MoveLast
    End If
    .MoveNext
Loop
If a = True Then
    Command1(0).Enabled = True
    Command1(1).Enabled = True
    Command1(3).Enabled = True
    Command1(0).SetFocus
Else
    x = MsgBox("Id Nasabah yang anda masukan tidak diketemukan, silahkan masukan Id Nasabah lainnya", vbOKOnly, "Nasabah Tidak Ketemu")
    Text1 = ""
    Label1(2) = "AUTO NAMA"
    Text1.SetFocus
End If
Data1.Refresh
End With
End Sub
