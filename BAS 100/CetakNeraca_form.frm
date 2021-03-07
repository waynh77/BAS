VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form CetakNeraca_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CETAK NERACA"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3840
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   2640
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "BATAL"
      Height          =   495
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   1140
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
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   495
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
      Index           =   2
      Left            =   2760
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   735
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3120
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SILAHKAN MASUKAN PERIODE "
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
      TabIndex        =   9
      Top             =   240
      Width           =   3000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "FORMAT (DD/MM/YYYY:23/04/1999)"
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
      TabIndex        =   8
      Top             =   480
      Width           =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Tgl:"
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
      TabIndex        =   7
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Bln:"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Thn:"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "CetakNeraca_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
'Dim totext As Date
Select Case Index
    Case 0
        If Text1(0) = "" Or Val(Text1(0)) > 31 Or Val(Text1(0)) < 1 Or Text1(1) = "" Or Val(Text1(1)) > 12 Or Val(Text1(1)) < 1 Or Text1(2) = "" Then
            x = MsgBox("Anda belum masukan data periode secara lengkap dan benar...", vbOKOnly, "Text Kosong")
            If Text1(0) = "" Or Val(Text1(0)) > 31 Or Val(Text1(0)) < 1 Then
                Text1(0).SetFocus
            ElseIf Text1(1) = "" Or Val(Text1(1)) > 12 Or Val(Text1(1)) < 1 Then
                Text1(1).SetFocus
            ElseIf Text1(2) = "" Then
                Text1(2).SetFocus
            End If
        Else
            If CetakNeraca_form.Caption = "CETAK NERACA" Then
                CrystalReport1.SelectionFormula = "{saldo_harian.tanggal}= date(" & Text1(2) & "," & Text1(1) & "," & Text1(0) & ")"
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.RetrieveDataFiles
                CrystalReport1.Action = 1
            Else
                CrystalReport1.SelectionFormula = "{saldo_harian.tanggal}= date(" & Text1(2) & "," & Text1(1) & "," & Text1(0) & ")"
                CrystalReport2.WindowState = crptMaximized
                CrystalReport2.RetrieveDataFiles
                CrystalReport2.Action = 1
            End If
            Me.Hide
            main_form.Enabled = True
        End If
    Case 1
        Me.Hide
        main_form.Enabled = True
        main_form.Show
End Select
End Sub

Private Sub Form_Activate()
Call cetak_neraca
Text1(0) = ""
Text1(1) = ""
Text1(2) = "2007"
Text1(0).SetFocus
Text1(0).MaxLength = 2
Text1(1).MaxLength = 2
Text1(2).MaxLength = 4
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 0
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text1(0) = "" Or ((Val(Text1(0)) > 31 Or Val(Text1(0)) < 1)) Then
            x = MsgBox("Anda belum masukan data tanggal/data tanggal error(31<data<0)...", vbOKOnly, "Text Kosong/Error")
            Text1(0).SetFocus
        Else
            Text1(1).SetFocus
        End If
    End If
    Case 1
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13 Or ((Val(Text1(0)) < 13 Or Val(Text1(0)) > 0))) Then
        Beep
        KeyAscii = 0
    ElseIf (Val(Text1(1)) < 1 And Val(Text1(1)) > 12) Then
        x = MsgBox("Anda belum masukan data tanggal/data tanggal error(31<data<0)...", vbOKOnly, "Text Kosong/Error")
        Text1(0).SetFocus
    End If
    If KeyAscii = 13 Then
        If Text1(1) = "" Or Val((Text1(1)) > 12 Or Val(Text1(1) < 1)) Then
            x = MsgBox("Anda belum masukan data bulan/data bulan error(31<data<0)...", vbOKOnly, "Text Kosong/Error")
        Else
            Text1(2).SetFocus
        End If
    End If
    Case 2
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        If Text1(2) = "" Then
            x = MsgBox("Anda belum masukan data tahun...", vbOKOnly, "Text Kosong")
        Else
            Command1(0).SetFocus
        End If
    End If
    End Select
End Sub


