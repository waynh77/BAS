VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Begin VB.Form DBTransNasabah_form 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Transaksi Nasabah"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "CETAK"
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CETAK"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "DBTransNasabah_form.frx":0000
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "DBTransNasabah_form.frx":0014
      TabIndex        =   6
      Top             =   4920
      Width           =   11535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "HAPUS"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Data Data2 
      Caption         =   "PINJAMAN"
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
      Top             =   4440
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBTransNasabah_form.frx":09E7
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "DBTransNasabah_form.frx":09FB
      TabIndex        =   3
      Top             =   960
      Width           =   11535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   375
      Left            =   10200
      TabIndex        =   2
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "TABUNGAN"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   120
      X2              =   11640
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI PINJAMAN"
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
      TabIndex        =   4
      Top             =   4080
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI TABUNGAN "
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
      Width           =   2280
   End
End
Attribute VB_Name = "DBTransNasabah_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
Else
    x = MsgBox("Data masih kosong...", vbOKOnly, "Data Kosong")
End If
End Sub

Private Sub Command2_Click()
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Sub

Private Sub Command3_Click()
If Not Data2.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data2.Recordset.Delete
        Data2.Refresh
    End If
Else
    x = MsgBox("Data masih kosong...", vbOKOnly, "Data Kosong")
End If
End Sub

Private Sub Command4_Click()
Call uncons
End Sub

Private Sub Command5_Click()
Call uncons
End Sub

Private Sub Form_Activate()
Data1.Refresh
Data2.Refresh
End Sub

Private Sub Form_Load()
Call Buka_DBTransNasabah
End Sub
