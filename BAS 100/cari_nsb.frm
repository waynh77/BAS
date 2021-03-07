VERSION 5.00
Begin VB.Form cari_nsb 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI DATA NASABAH"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   3000
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CARI"
      Height          =   495
      Index           =   0
      Left            =   720
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SILAHKAN MASUKAN ID ATAU NAMA  NASABAH:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4680
   End
End
Attribute VB_Name = "cari_nsb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If cari_nsb.Caption = "CARI DATA NASABAH" Then
        Data1.RecordSource = "select * from tabel_nasabah where id_nasabah = '" & Text1 & "' OR nama_nasabah like '*" & Text1 & "*'"
        Data1.Refresh
        If Not Data1.Recordset.BOF Then
            Nasabah_form.Data1.RecordSource = Data1.RecordSource
            Nasabah_form.Data1.Refresh
            Call nasabah_isi
            Unload Me
            Nasabah_form.Show
            Nasabah_form.Enabled = True
        Else
            x = MsgBox("Data tidak diketemukan, silahkan masukan id/nama lainnya", vbOKOnly, "DATA TIDAK KETEMU")
            Text1 = ""
            Text1.SetFocus
        End If
    Else
    
    End If
Case 1
    Unload Me
    Nasabah_form.Show
    Nasabah_form.Enabled = True
End Select
End Sub

Private Sub Form_Load()
Call Buka_DBcari_nsb
Text1 = ""
Text1.MaxLength = 50
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(0).SetFocus
End If
End Sub
