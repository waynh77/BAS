VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form main_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BMT Accounting System (BAS) ver.1.0.0"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "main_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   10320
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9720
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXIT"
      Height          =   855
      Index           =   1
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LOGIN"
      Height          =   855
      Index           =   0
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5400
      PasswordChar    =   "*"
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4680
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   8880
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3480
         Top             =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Jam :"
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
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Tanggal :"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Label2"
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
         Left            =   1560
         TabIndex        =   2
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "Label1"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TABUNGAN NSB"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   7
      Left            =   9120
      MouseIcon       =   "main_form.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   40
      Top             =   3360
      Width           =   1800
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "WaynhSoft"
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
      Left            =   360
      TabIndex        =   13
      Top             =   600
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Amanah Ummah"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   450
      Index           =   2
      Left            =   360
      TabIndex        =   39
      Top             =   840
      Width           =   2880
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "1. Menolong kaum dhuafa dari jeratan rentenir dan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   3
      Left            =   360
      TabIndex        =   29
      Top             =   3960
      Width           =   5880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "M I S I "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   19
      Left            =   2880
      TabIndex        =   38
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "V I S I"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Index           =   20
      Left            =   2880
      TabIndex        =   37
      Top             =   1800
      Width           =   1155
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AKUNTANSI"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   18
      Left            =   7320
      MouseIcon       =   "main_form.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   36
      ToolTipText     =   "tabel akun"
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   17
      Left            =   360
      MouseIcon       =   "main_form.frx":0EDE
      MousePointer    =   99  'Custom
      TabIndex        =   35
      ToolTipText     =   "keluar dari program"
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "jujur, dan berakhlaqul karimah"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   8
      Left            =   720
      TabIndex        =   34
      Top             =   5760
      Width           =   3600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "3. BMT dikelola oleh tenaga yang profesional, amanah"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   7
      Left            =   360
      TabIndex        =   33
      Top             =   5400
      Width           =   6240
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "   Islam dan lingkungan sekitarnya"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   6
      Left            =   360
      TabIndex        =   32
      Top             =   5040
      Width           =   4080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "2. Meningkatkan pendapatan dan kesejahteraan Ummat"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   5
      Left            =   360
      TabIndex        =   31
      Top             =   4680
      Width           =   6000
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "   lintah darat"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   4
      Left            =   360
      TabIndex        =   30
      Top             =   4320
      Width           =   1800
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Allah dan Kemaslahatan Ummat"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   2
      Left            =   1680
      TabIndex        =   28
      Top             =   2880
      Width           =   3360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Utama Mitra Usaha dalam rangka menggapai ridlo"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   1
      Left            =   600
      TabIndex        =   27
      Top             =   2520
      Width           =   5520
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Menjadi BMT yang dipercaya ummat Islam dan pilihan"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   2160
      Width           =   6000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   6015
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "IKHTISAR KEUANGAN"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   16
      Left            =   7320
      MouseIcon       =   "main_form.frx":11E8
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6720
      Width           =   2550
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NERACA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   15
      Left            =   7320
      MouseIcon       =   "main_form.frx":14F2
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   5760
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "LAPORAN RUGI-LABA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   14
      Left            =   7320
      MouseIcon       =   "main_form.frx":17FC
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   6240
      Width           =   2550
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI JURNAL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   12
      Left            =   7320
      MouseIcon       =   "main_form.frx":1B06
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   4800
      Width           =   2400
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI NASABAH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   11
      Left            =   7320
      MouseIcon       =   "main_form.frx":1E10
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   4320
      Width           =   2550
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "DATA TRANS.NSB"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   10
      Left            =   9120
      MouseIcon       =   "main_form.frx":211A
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   2880
      Width           =   2100
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "BUKU BESAR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   9
      Left            =   9120
      MouseIcon       =   "main_form.frx":2424
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JURNAL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   8
      Left            =   7320
      MouseIcon       =   "main_form.frx":272E
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NASABAH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   6
      Left            =   7320
      MouseIcon       =   "main_form.frx":2A38
      MousePointer    =   99  'Custom
      TabIndex        =   17
      ToolTipText     =   "data base nasabah"
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "LAPORAN"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   5
      Left            =   6960
      TabIndex        =   16
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "TRANSAKSI"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   4
      Left            =   6960
      TabIndex        =   15
      Top             =   3840
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "DATA BASE"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   270
      Index           =   3
      Left            =   6960
      TabIndex        =   14
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "BMT ACCOUNTING SYSTEM VER.1.0.0"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   450
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   240
      Width           =   7440
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "PLEASE ENTER YOUR USER NAME AND PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   1
      Left            =   3600
      MouseIcon       =   "main_form.frx":2D42
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   120
      Picture         =   "main_form.frx":304C
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   11775
   End
   Begin VB.Menu data_mnu 
      Caption         =   "&Data Base"
      Begin VB.Menu nasabah_mnu 
         Caption         =   "&Nasabah"
      End
      Begin VB.Menu akun_mnu 
         Caption         =   "&Tabel Akun"
      End
      Begin VB.Menu jurnal_mnu 
         Caption         =   "&Jurnal"
      End
      Begin VB.Menu BB_mnu 
         Caption         =   "&Buku Besar"
      End
      Begin VB.Menu SubBB_mnu 
         Caption         =   "&Sub Buku Besar"
      End
   End
   Begin VB.Menu transaksi_mnu 
      Caption         =   "&Transaksi"
      Begin VB.Menu TNasabah_mnu 
         Caption         =   "Transaksi &Nasabah "
      End
      Begin VB.Menu TJurnal_mnu 
         Caption         =   "Transaksi &Jurnal"
      End
   End
   Begin VB.Menu lap_mnu 
      Caption         =   "&Laporan"
      Begin VB.Menu rl_mnu 
         Caption         =   "&Rugi Laba"
      End
      Begin VB.Menu neraca_mnu 
         Caption         =   "&Neraca"
      End
      Begin VB.Menu analisis_mnu 
         Caption         =   "&Ikhtisar Keuangan"
      End
   End
   Begin VB.Menu calk_mnu 
      Caption         =   "Kalkulator"
   End
   Begin VB.Menu setting_mnu 
      Caption         =   "&Setting"
      Begin VB.Menu back_mnu 
         Caption         =   "Background"
         Begin VB.Menu yback_mnu 
            Caption         =   "Aktif"
         End
         Begin VB.Menu noback_mnu 
            Caption         =   "Tidak Aktif"
         End
      End
   End
   Begin VB.Menu kunci_mnu 
      Caption         =   "Kunci"
   End
   Begin VB.Menu keluar_mnu 
      Caption         =   "&Keluar"
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub Absen_mnu_Click()
Call uncons
End Sub

Private Sub akun_mnu_Click()
    Me.Enabled = False
    AkunTbl_form.Show
End Sub

Private Sub analisis_mnu_Click()
Call uncons
End Sub

Private Sub BB_mnu_Click()
Me.Enabled = False
BukuBesar_form.Show
End Sub

Private Sub calk_mnu_Click()
    AppActivate Shell("calc.exe", 1)
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    'check for correct password
    If Text1 = "" And Text2 = "" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        login_benar
        buka_menu
    Else
        MsgBox "Invalid user name or Password, try again!", , "Login"
        Text2.SetFocus
        SendKeys "{Home}+{End}"
    End If
Case 1
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    End
End Select
End Sub

Private Sub tutup_menu()
data_mnu.Enabled = False
transaksi_mnu.Enabled = False
lap_mnu.Enabled = False
setting_mnu.Enabled = False
calk_mnu.Enabled = False
kunci_mnu.Enabled = False
Label4(3).Visible = False
Label4(4).Visible = False
Label4(5).Visible = False
Label4(6).Visible = False
Label4(7).Visible = False
Label4(8).Visible = False
Label4(9).Visible = False
Label4(10).Visible = False
Label4(11).Visible = False
Label4(12).Visible = False
Label4(14).Visible = False
Label4(15).Visible = False
Label4(16).Visible = False
Label4(17).Visible = False
Label4(18).Visible = False
Label4(19).Visible = False
Label4(20).Visible = False
Shape1.Visible = False
Label6(0).Visible = False
Label6(1).Visible = False
Label6(2).Visible = False
Label6(3).Visible = False
Label6(4).Visible = False
Label6(5).Visible = False
Label6(6).Visible = False
Label6(7).Visible = False
Label6(8).Visible = False
End Sub
Private Sub buka_menu()
data_mnu.Enabled = True
transaksi_mnu.Enabled = True
lap_mnu.Enabled = True
calk_mnu.Enabled = True
setting_mnu.Enabled = True
Shape1.Visible = True
kunci_mnu.Enabled = True
Label4(3).Visible = True
Label4(4).Visible = True
Label4(5).Visible = True
Label4(6).Visible = True
Label4(7).Visible = True
Label4(8).Visible = True
Label4(9).Visible = True
Label4(10).Visible = True
Label4(11).Visible = True
Label4(12).Visible = True
Label4(14).Visible = True
Label4(15).Visible = True
'Label4(16).Visible = True
Label4(17).Visible = True
Label4(18).Visible = True
Label4(19).Visible = True
Label4(20).Visible = True
Label6(0).Visible = True
Label6(1).Visible = True
Label6(2).Visible = True
Label6(3).Visible = True
Label6(4).Visible = True
Label6(5).Visible = True
Label6(6).Visible = True
Label6(7).Visible = True
Label6(8).Visible = True
Image1.Visible = False
End Sub


Private Sub Form_Load()
'Me.WindowState = 2
Label1 = Format(Date, "d/m/yyyy")
Label2 = Format(Time, "hh:mm:ss")
Text1 = ""
Text2 = ""
tutup_menu
noback_mnu.Enabled = False
Call utama
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub jurnal_mnu_Click()
Me.Enabled = False
DBJurnal_form.Show
End Sub

Private Sub keluar_mnu_Click()
    With Data1.Recordset
    If Not Data1.Recordset.BOF Then
        .MoveFirst
        Do While Not .EOF
            If !tanggal = Date Then
                .Delete
            End If
            .MoveNext
        Loop
        Data1.Refresh
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            .AddNew
            !tanggal = Date
            !sandi_akun = Data2.Recordset!sandi_akun
            !saldo_akun = Data2.Recordset!saldo_akun
            .Update
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
     Else
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            .AddNew
            !tanggal = Date
            !sandi_akun = Data2.Recordset!sandi_akun
            !saldo_akun = Data2.Recordset!saldo_akun
            .Update
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
    End If
    Data1.Refresh
    End With
    End
End Sub

Private Sub kunci_mnu_Click()
Text1 = ""
Text2 = ""
tutup_menu
noback_mnu.Enabled = False
Text1.Visible = True
Text2.Visible = True
Image1.Visible = True
Label4(0).Visible = True
Label4(1).Visible = True
Label4(2).Visible = True
Command1(0).Visible = True
Command1(1).Visible = True
End Sub

Private Sub Label4_Click(Index As Integer)
Select Case Index
Case 7
    Me.Enabled = False
    SaldoTabungan_form.Show
Case 6
    Nasabah_form.Show
    Me.Enabled = False
Case 8
    Me.Enabled = False
    DBJurnal_form.Show
Case 9
    BukuBesar_form.Show
    main_form.Enabled = False
Case 10
    DBTransNasabah_form.Show
    Me.Enabled = False
Case 11
    SubTransNasabah_form.Show
    Me.Enabled = False
Case 12
    TransJurnal_form.Show
    Me.Enabled = False
Case 13
    Call uncons
Case 14
    Me.Enabled = False
    CetakNeraca_form.Caption = "CETAK RUGI/LABA"
    CetakNeraca_form.Show
'    CrystalReport2.WindowState = crptMaximized
'    CrystalReport2.RetrieveDataFiles
'    CrystalReport2.Action = 1
Case 15
    Me.Enabled = False
    CetakNeraca_form.Caption = "CETAK NERACA"
    CetakNeraca_form.Show
'    CrystalReport1.WindowState = crptMaximized
'    CrystalReport1.RetrieveDataFiles
'    CrystalReport1.Action = 1
Case 16
    Call uncons
Case 17
    With Data1.Recordset
    If Not Data1.Recordset.BOF Then
        .MoveFirst
        Do While Not .EOF
            If !tanggal = Date Then
                .Delete
            End If
            .MoveNext
        Loop
        Data1.Refresh
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            .AddNew
            !tanggal = Date
            !sandi_akun = Data2.Recordset!sandi_akun
            !saldo_akun = Data2.Recordset!saldo_akun
            .Update
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
     Else
        Data2.Recordset.MoveFirst
        Do While Not Data2.Recordset.EOF
            .AddNew
            !tanggal = Date
            !sandi_akun = Data2.Recordset!sandi_akun
            !saldo_akun = Data2.Recordset!saldo_akun
            .Update
            Data2.Recordset.MoveNext
        Loop
        Data2.Refresh
    End If
    Data1.Refresh
    End With
    End
Case 18
    Me.Enabled = False
    AkunTbl_form.Show
End Select
End Sub


Private Sub nasabah_mnu_Click()
Nasabah_form.Show
Me.Enabled = False
End Sub

Private Sub neraca_mnu_Click()
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
End Sub

Private Sub noback_mnu_Click()
noback_mnu.Enabled = False
yback_mnu.Enabled = True
Image1.Visible = False
End Sub

Private Sub rl_mnu_Click()
    CrystalReport2.WindowState = crptMaximized
    CrystalReport2.RetrieveDataFiles
    CrystalReport2.Action = 1
End Sub

Private Sub SubBB_mnu_Click()
Me.Enabled = False
DBTransNasabah_form.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Text2.SetFocus
End If
End Sub

Private Sub text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Command1(0).SetFocus
End If
End Sub

Private Sub Timer1_Timer()
Label2 = Format(Time, "hh:mm:ss")
End Sub

Private Sub login_benar()
Label4(0).Visible = False
Label4(1).Visible = False
Label4(2).Visible = False
Text1.Visible = False
Text2.Visible = False
Command1(0).Visible = False
Command1(1).Visible = False
End Sub


Private Sub TJurnal_mnu_Click()
Me.Enabled = False
TransJurnal_form.Show
End Sub

Private Sub TNasabah_mnu_Click()
    SubTransNasabah_form.Show
    Me.Enabled = False
End Sub

Private Sub yback_mnu_Click()
yback_mnu.Enabled = False
noback_mnu.Enabled = True
Image1.Visible = True
End Sub
