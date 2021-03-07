VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form SaldoTabungan_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALDO TABUNGAN NASABAH"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "SaldoTabungan_form.frx":0000
      Height          =   4095
      Left            =   120
      OleObjectBlob   =   "SaldoTabungan_form.frx":0014
      TabIndex        =   3
      Top             =   600
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   375
      Index           =   2
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   375
      Index           =   1
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS"
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Data Data1 
      BackColor       =   &H80000006&
      Caption         =   "SALDO TABUNGAN"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "SaldoTabungan_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data1.Recordset.Delete
            Data1.Refresh
        End If
    Else
        x = MsgBox("Data masih kosong...", vbOKOnly, "Data Kosong")
    End If
Case 1
    Call uncons
Case 2
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub Form_Load()
Call Buka_DBSaldoTabungan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Sub
