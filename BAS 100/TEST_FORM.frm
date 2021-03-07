VERSION 5.00
Begin VB.Form TEST_FORM 
   BackColor       =   &H80000012&
   Caption         =   "TEST"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Caption         =   "HASIL TEST"
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TEST COLOR"
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   240
      MouseIcon       =   "TEST_FORM.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   480
      Width           =   1815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "TEST_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As Byte
a = Val(Text1)
Label2.ForeColor = QBColor(a)
Text1 = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
End
End Sub


Private Sub Form_Load()
Text1 = ""
'Text1.SetFocus
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = QBColor(12)
End Sub

