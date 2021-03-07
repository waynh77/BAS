VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Trans_Nasabah_form 
   BackColor       =   &H80000007&
   Caption         =   "TRANSAKSI NASABAH"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "E:\bmt project1\database\dataBMT.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "trans_nasabah"
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10320
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&KELUAR"
      Height          =   615
      Index           =   2
      Left            =   10080
      TabIndex        =   11
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "TRANSAKSI &BATAL"
      Height          =   615
      Index           =   1
      Left            =   8400
      TabIndex        =   10
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&SIMPAN TRANSAKSI"
      Height          =   615
      Index           =   0
      Left            =   6720
      TabIndex        =   9
      Top             =   7200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   4560
      Width           =   2535
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   3120
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "E:\bmt project1\database\dataBMT.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tabel_nasabah"
      Top             =   7680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&CETAK"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   8
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&HAPUS"
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&EDIT"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&PROSES"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   7560
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Trans_Nasabah_form.frx":0000
      Height          =   2055
      Left            =   360
      OleObjectBlob   =   "Trans_Nasabah_form.frx":0014
      TabIndex        =   30
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\bmt project1\database\dataBMT.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "temp_Tnasabah"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   31
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label3 
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
      Index           =   22
      Left            =   3000
      TabIndex        =   51
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AKUN TRANSASI"
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
      Index           =   21
      Left            =   6960
      TabIndex        =   50
      Top             =   6480
      Width           =   1560
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO DEBET/KREDIT"
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
      Index           =   11
      Left            =   9480
      TabIndex        =   49
      Top             =   6480
      Width           =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA TRANSAKSI"
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
      Index           =   20
      Left            =   6960
      TabIndex        =   48
      Top             =   5760
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS TRANS"
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
      Left            =   9480
      TabIndex        =   47
      Top             =   5760
      Width           =   1320
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   5295
      Index           =   2
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   6375
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1215
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO DEBET/KREDIT"
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
      Left            =   3120
      TabIndex        =   46
      Top             =   5040
      Width           =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA TRANSAKSI"
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
      Index           =   18
      Left            =   720
      TabIndex        =   45
      Top             =   4080
      Width           =   1680
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO NUMBER"
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
      Left            =   3120
      TabIndex        =   44
      Top             =   3120
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AUTO NAMA"
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
      Left            =   3000
      TabIndex        =   43
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NOMINAL"
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
      Left            =   9480
      TabIndex        =   42
      Top             =   6120
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NILAI TRANSAKSI Rp"
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
      Index           =   19
      Left            =   6960
      TabIndex        =   41
      Top             =   6120
      Width           =   2160
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS TRANS"
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
      Left            =   9480
      TabIndex        =   40
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "DATE"
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
      Left            =   9480
      TabIndex        =   39
      Top             =   5040
      Width           =   480
   End
   Begin VB.Label Label4 
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
      Left            =   9480
      TabIndex        =   38
      Top             =   4080
      Width           =   1200
   End
   Begin VB.Label Label4 
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
      Left            =   9480
      TabIndex        =   37
      Top             =   3720
      Width           =   1200
   End
   Begin VB.Label Label4 
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
      Index           =   1
      Left            =   9480
      TabIndex        =   36
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Label Label4 
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
      Index           =   0
      Left            =   9480
      TabIndex        =   35
      Top             =   2400
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NO TRANSAKSI :"
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
      Index           =   17
      Left            =   720
      TabIndex        =   34
      Top             =   3120
      Width           =   1680
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
      TabIndex        =   33
      Top             =   360
      Width           =   7440
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
      TabIndex        =   32
      Top             =   960
      Width           =   2880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "AKUN TRANSAKSI"
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
      Index           =   16
      Left            =   720
      TabIndex        =   29
      Top             =   5040
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NILAI TRANSAKSI Rp"
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
      Index           =   15
      Left            =   720
      TabIndex        =   28
      Top             =   4560
      Width           =   2160
   End
   Begin VB.Label Label3 
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
      Index           =   14
      Left            =   720
      TabIndex        =   27
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   13
      Left            =   360
      TabIndex        =   26
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "PINJAMAN MUSYARAKAH"
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
      Index           =   12
      Left            =   6960
      TabIndex        =   25
      Top             =   4080
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "PINJAMAN MUDHARABAH"
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
      Index           =   11
      Left            =   6960
      TabIndex        =   24
      Top             =   3720
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SALDO PINJAMAN"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   10
      Left            =   6360
      TabIndex        =   23
      Top             =   3360
      Width           =   1680
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TABUNGAN MUDHARABAH"
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
      Left            =   6960
      TabIndex        =   22
      Top             =   2760
      Width           =   2280
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TABUNGAN WADIAH"
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
      Left            =   6960
      TabIndex        =   21
      Top             =   2400
      Width           =   1800
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   7
      Left            =   6360
      TabIndex        =   20
      Top             =   2040
      Width           =   1680
   End
   Begin VB.Label Label3 
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
      Index           =   6
      Left            =   6960
      TabIndex        =   19
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TANGGAL"
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
      Left            =   6960
      TabIndex        =   18
      Top             =   5040
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "TRANSAKSI TERAKHIR"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Index           =   4
      Left            =   6360
      TabIndex        =   17
      Top             =   4680
      Width           =   2160
   End
   Begin VB.Label Label3 
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
      Index           =   3
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "ID NASABAH"
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
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Width           =   1200
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
      Left            =   10440
      TabIndex        =   14
      Top             =   480
      Width           =   720
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
      Left            =   10440
      TabIndex        =   13
      Top             =   840
      Width           =   720
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
      Left            =   9120
      TabIndex        =   12
      Top             =   480
      Width           =   1080
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
      Left            =   9600
      TabIndex        =   0
      Top             =   840
      Width           =   600
   End
End
Attribute VB_Name = "Trans_Nasabah_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tutup()
Combo1.Enabled = False
Combo2.Enabled = False
Text2.Enabled = False
End Sub

Private Sub buka()
Combo1.Enabled = True
Combo2.Enabled = True
Text2.Enabled = True
End Sub

Private Sub Combo1_Click()
nama_trans
dk
End Sub

Private Sub dk()
If Combo1 = "Setoran Tabungan" Or Combo1 = "Pembayaran Pinjaman" Then
    Label4(9).Caption = "KREDIT"
Else
    Label4(9).Caption = "DEBET"
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Combo2_Click()
Text2.SetFocus
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0

Case 1

Case 2
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub Form_Activate()
Label1 = Format(Date, "dd/mm/yyyy")
Label2 = Format(Time, "hh:mm:ss")
Text1.MaxLength = 5
Text2.MaxLength = 10
isi_jenis
kosong
tutup
Text1.SetFocus
End Sub

Private Sub isi_jenis()
Combo1.AddItem ("Setoran Tabungan")
Combo1.AddItem ("Pengambilan Tabungan")
Combo1.AddItem ("Penarikan Pinjaman")
Combo1.AddItem ("Pembayaran Pinjaman")
End Sub

Private Sub nama_trans()
Combo2.Clear
If Combo1 = "Setoran Tabungan" Or Combo1 = "Pengambilan Tabungan" Then
    Combo2.AddItem ("Tabungan Wadiah")
    Combo2.AddItem ("Tabungan Mudharabah")
Else
    Combo2.AddItem ("Pinjaman Mudharabah")
    Combo2.AddItem ("Pinjaman Musyarakah")
End If
Combo2.ListIndex = 0
Combo2.SetFocus
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        cek_nasabah
        Combo1.TabIndex = 0
    End If
End Sub

Private Sub cek_nasabah()
Dim a As Boolean
Dim b As String
With Data2.Recordset
If Not Data2.Recordset.BOF Then
a = False
b = "NSB"
.MoveFirst
    Do While Not .EOF
        If b & Text1 = !id_nasabah Then
            a = True
            Label4(7).Caption = !nama_nasabah
            .MoveLast
            buka
            Combo1.SetFocus
            Combo1.ListIndex = 0
            Combo2.ListIndex = 0
        End If
        .MoveNext
    Loop
If a = False Then
    X = MsgBox("Id Nasabah tidak diketemukan...!!! silahkan masukan id yang lain", vbOKOnly, "Warning")
    Text1 = ""
    Text1.SetFocus
End If
End If
End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = 13) Then
        Beep
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        Command1(0).SetFocus
    End If
End Sub

Private Sub Timer1_Timer()
Label2 = Format(Time, "hh:mm:ss")
End Sub

Private Sub kosong()
Text1 = ""
Combo1 = ""
Combo2 = ""
Text2 = ""
End Sub
