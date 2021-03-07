VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form DBJurnal_form 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Data Base Jurnal"
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "HAPUS"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CETAK"
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.Data Data1 
      Caption         =   "JURNAL"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "DBJurnal_form2.frx":0000
      Height          =   7455
      Left            =   120
      OleObjectBlob   =   "DBJurnal_form2.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   11535
   End
End
Attribute VB_Name = "DBJurnal_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    Me.Enabled = False
    CetakDBJurnal_form.Show
Case 1
If Not Data1.Recordset.BOF Then
    x = MsgBox("Apakah anda yakin data akan dihapus?", vbOKCancel, "Hapus Data")
    If x = vbOK Then
        Data1.Recordset.Delete
        Data1.Refresh
    End If
Else
    x = MsgBox("Data masih kosong...", vbOKOnly, "Data Kosong")
End If
Case 2
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub Form_Load()
Call Buka_DBJurnal
End Sub
