VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form Nasabah_form 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DATA BASE NASABAH"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Bindings        =   "Nasabah_form.frx":0000
      Left            =   5520
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh"
      Height          =   495
      Index           =   5
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "cari data"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Find"
      Height          =   495
      Index           =   4
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "cari data"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Last =>|"
      Height          =   495
      Index           =   3
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "last data"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Next =>"
      Height          =   495
      Index           =   2
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "next data"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<= Prev"
      Height          =   495
      Index           =   1
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "previus data"
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "|<= First"
      Height          =   495
      Index           =   0
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "first data"
      Top             =   5160
      Width           =   855
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Nasabah_form.frx":0014
      Height          =   2295
      Left            =   240
      OleObjectBlob   =   "Nasabah_form.frx":0028
      TabIndex        =   60
      Top             =   5760
      Width           =   11535
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   4
      Left            =   4440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "keluar data base nasabah"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "HAPUS"
      Height          =   495
      Index           =   3
      Left            =   3480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "hapus data nasabah"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CETAK"
      Height          =   495
      Index           =   2
      Left            =   2520
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "cetak data nasabah"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EDIT"
      Height          =   495
      Index           =   1
      Left            =   1560
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "ubah data nasabah"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TAMBAH"
      Height          =   495
      Index           =   0
      Left            =   600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "tambah data nasabah"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000008&
      Caption         =   "DATA PRIBADI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   120
      TabIndex        =   48
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Text5 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   13
         Text            =   "Text9"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   12
         Text            =   "Text8"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Text            =   "Text7"
         Top             =   3480
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Text            =   "Text6"
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   7
         Text            =   "Text5"
         Top             =   2760
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "Perempuan"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000012&
         Caption         =   "Laki-laki"
         ForeColor       =   &H8000000E&
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Text            =   "Text4"
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   3
         Text            =   "Text3"
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Text            =   "Text2"
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label23 
         BackColor       =   &H80000012&
         Caption         =   "Tgl/Bln/Thn (01/01/1900)"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   2520
         TabIndex        =   59
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "PENDIDIKAN TERAKHIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   4200
         Width           =   2160
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "STATUS NIKAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Top             =   3840
         Width           =   1380
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "KEWARGANEGARAAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   56
         Top             =   3480
         Width           =   1965
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "AGAMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "TANGGAL LAHIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   54
         Top             =   2400
         Width           =   1500
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "TEMPAT LAHIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   2040
         Width           =   1380
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "JENIS KELAMIN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "NOMOR KTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "NAMA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   50
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "ID NASABAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4815
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.TextBox Text21 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   25
         Text            =   "Text21"
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox Text20 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Text            =   "Text20"
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text19 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Text            =   "Text19"
         Top             =   3600
         Width           =   1935
      End
      Begin VB.TextBox Text18 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Text            =   "Text18"
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text17 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   21
         Text            =   "Text17"
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Text16 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   20
         Text            =   "Text16"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Text            =   "Text15"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox Text14 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Text            =   "Text14"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Text            =   "Text13"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Text            =   "Text12"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   15
         Text            =   "Text11"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   285
         Left            =   2160
         TabIndex        =   14
         Text            =   "Text10"
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "NAMA JALAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "NOMOR RUMAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   46
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "RT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   45
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "RW"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   330
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "KELURAHAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   43
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "KECAMATAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         Width           =   1140
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "KABUPATEN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   2520
         Width           =   1125
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "PROPINSI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   2880
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "TELEPON RUMAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   3600
         Width           =   1605
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "HAND PHONE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   3960
         Width           =   1245
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "KODE POS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   3240
         Width           =   960
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         Caption         =   "EMAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   4320
         Width           =   570
      End
   End
End
Attribute VB_Name = "Nasabah_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "SIMPAN" Then
        If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5(0) = "" Or Text5(1) = "" Or Text5(2) = "" Or Text6 = "" Or Text7 = "" Or Text8 = "" Or Text9 = "" Or Text10 = "" Or Text11 = "" Or Text12 = "" Or Text13 = "" Or Text14 = "" Or Text15 = "" Or Text16 = "" Or Text17 = "" Or Text18 = "" Or Text19 = "" Or Text20 = "" Or Text21 = "" Then
            x = MsgBox("Data belum lengkap", vbOKOnly, "Peringatan...!!!")
            Call nasabah_validasi1
        Else
            Call nasabah_simpan
            Data1.Recordset.Update
            Data1.Refresh
            Call nasabah_isi
            Call nasabah_burem
            Command1(0).Caption = "TAMBAH"
            Command1(1).Caption = "EDIT"
            DBGrid1.Enabled = True
            Command1(2).Visible = True
            Command1(3).Visible = True
            Command1(4).Visible = True
            Command2(0).Enabled = True
            Command2(1).Enabled = True
            Command2(2).Enabled = True
            Command2(3).Enabled = True
            Command2(4).Enabled = True
        End If
    Else
        Command2(0).Enabled = False
        Command2(1).Enabled = False
        Command2(2).Enabled = False
        Command2(3).Enabled = False
        Command2(4).Enabled = False
        Call nasabah_terang
        Call nasabah_kosong
        Call nasabah_auto
        Data1.Recordset.AddNew
        Text2.SetFocus
        Command1(0).Caption = "SIMPAN"
        Command1(1).Caption = "BATAL"
        DBGrid1.Enabled = False
        Command1(2).Visible = False
        Command1(3).Visible = False
        Command1(4).Visible = False
    End If
Case 1
    If Command1(1).Caption = "BATAL" Then
        If Data1.Recordset.BOF Then
            x = MsgBox("Data masih kosong,apakah anda yakin batal...?", vbOKCancel, "Data Kosong...!!!")
            If x = vbOK Then
                Call nasabah_kosong
                Data1.Refresh
                Call nasabah_isi
                Call nasabah_burem
                Command1(0).Caption = "TAMBAH"
                Command1(1).Caption = "EDIT"
                DBGrid1.Enabled = True
                Command1(2).Visible = True
                Command1(3).Visible = True
                Command1(4).Visible = True
            Else
                Text2.SetFocus
            End If
        Else
            Command2(0).Enabled = True
            Command2(1).Enabled = True
            Command2(2).Enabled = True
            Command2(3).Enabled = True
            Command2(4).Enabled = True
            Data1.Refresh
            Call nasabah_isi
            Call nasabah_burem
            Command1(0).Caption = "TAMBAH"
            Command1(1).Caption = "EDIT"
            DBGrid1.Enabled = True
            Command1(2).Visible = True
            Command1(3).Visible = True
            Command1(4).Visible = True
        End If
    Else
        Command2(0).Enabled = False
        Command2(1).Enabled = False
        Command2(2).Enabled = False
        Command2(3).Enabled = False
        Command2(4).Enabled = False
        If Not Data1.Recordset.BOF Then
            Data1.Recordset.Edit
            Call nasabah_terang
            Text2.SetFocus
            Command1(0).Caption = "SIMPAN"
            Command1(1).Caption = "BATAL"
            DBGrid1.Enabled = False
            Command1(2).Visible = False
            Command1(3).Visible = False
            Command1(4).Visible = False
        Else
            Call nasabah_kosong
            Call nasabah_terang
            x = MsgBox("Silahkan isi data terlebih dahulu", vbOKOnly, "Data Kosong")
            Command1(0).Caption = "SIMPAN"
            Command1(1).Caption = "BATAL"
            DBGrid1.Enabled = False
            Command1(2).Visible = False
            Command1(3).Visible = False
            Command1(4).Visible = False
            Data1.Recordset.AddNew
            Call nasabah_auto
            Text1.Enabled = False
            Text2.SetFocus
        End If
    End If
Case 2
    CrystalReport1.SelectionFormula = "{tabel_nasabah.id_nasabah}= '" & Text1 & "'"
    CrystalReport1.WindowState = crptMaximized
    CrystalReport1.RetrieveDataFiles
    CrystalReport1.Action = 1
Case 3
    If Not Data1.Recordset.BOF Then
        x = MsgBox("Apakah anda yakin data akan dihapus", vbOKCancel, "Hapus Data")
        If x = vbOK Then
            Data1.Recordset.Delete
            Data1.Refresh
            Call nasabah_isi
        Else
            Data1.Refresh
            Call nasabah_isi
        End If
    Else
        Call nasabah_kosong
        Call nasabah_terang
        x = MsgBox("Silahkan isi data terlebih dahulu", vbOKOnly, "Data Kosong")
        Command1(0).Caption = "SIMPAN"
        Command1(1).Caption = "BATAL"
        DBGrid1.Enabled = False
        Command1(2).Visible = False
        Command1(3).Visible = False
        Command1(4).Visible = False
        Data1.Recordset.AddNew
        Call nasabah_auto
        Text1.Enabled = False
        Text2.SetFocus
    End If
Case 4
    Me.Hide
    Unload Me
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
If Not Data1.Recordset.BOF Then
    Select Case Index
        Case 0
            Data1.Recordset.MoveFirst
            Call nasabah_isi
        Case 1
            Data1.Recordset.MovePrevious
            If Data1.Recordset.BOF Then
                x = MsgBox("Anda sudah di data pertama", vbOKOnly, "Data Pertama")
                Data1.Recordset.MoveFirst
            End If
            Call nasabah_isi
        Case 2
            Data1.Recordset.MoveNext
            If Data1.Recordset.EOF Then
                x = MsgBox("Anda sudah berada di data terakhir", vbOKOnly, "Data Terakhir")
                Data1.Recordset.MoveLast
            End If
            Call nasabah_isi
        Case 3
            Data1.Recordset.MoveLast
            Call nasabah_isi
        Case 4
            Me.Enabled = False
            cari_nsb.Show
            cari_nsb.Label1.Caption = "SILAHKAN MASUKAN ID ATAU NAMA NASABAH:"
            cari_nsb.Caption = "CARI DATA NASABAH"
            'blm_ada
        Case 5
            Data1.RecordSource = "select * from tabel_nasabah"
            Data1.Refresh
            Call nasabah_isi
    End Select
Else
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
End If
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Command1(0).Caption = "SIMPAN" Then
    Call nasabah_isi
End If
End Sub

Private Sub Form_Activate()
limiter
If Data1.Recordset.BOF Then
    Call nasabah_kosong
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
    x = MsgBox("Silahkan isi data terlebih dahulu", vbOKOnly, "Data Kosong")
    Command1(0).Caption = "SIMPAN"
    Command1(1).Caption = "BATAL"
    DBGrid1.Enabled = False
    Command1(2).Visible = False
    Command1(3).Visible = False
    Command1(4).Visible = False
    Data1.Recordset.AddNew
    Call nasabah_auto
    Text1.Enabled = False
    Text2.SetFocus
Else
    Call nasabah_burem
    Call nasabah_isi
    Command2(0).Enabled = True
    Command2(1).Enabled = True
    Command2(2).Enabled = True
    Command2(3).Enabled = True
    Command2(4).Enabled = True
End If
End Sub

Private Sub Form_Load()
Call Buka_DBNasabah
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
Unload Me
main_form.Enabled = True
main_form.Show
End Sub

Private Sub blm_ada()
    Beep
    x = MsgBox("MAAF FITUR INI MASIH DALAM PENGEMBANGAN", vbOKOnly, "DALAM PROSES")
End Sub

Private Sub limiter()
Text1.MaxLength = 10
Text2.MaxLength = 30
Text3.MaxLength = 20
Text4.MaxLength = 30
Text5(0).MaxLength = 2
Text5(1).MaxLength = 2
Text5(2).MaxLength = 4
Text6.MaxLength = 10
Text7.MaxLength = 20
Text8.MaxLength = 20
Text9.MaxLength = 20
Text10.MaxLength = 50
Text11.MaxLength = 10
Text12.MaxLength = 5
Text13.MaxLength = 5
Text14.MaxLength = 20
Text15.MaxLength = 20
Text16.MaxLength = 20
Text17.MaxLength = 30
Text18.MaxLength = 10
Text19.MaxLength = 20
Text20.MaxLength = 20
Text21.MaxLength = 50
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub


Private Sub Text19_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub


Private Sub Text20_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub

Private Sub Text5_Change(Index As Integer)
Select Case Index
Case 0
    If Val(Text5(0)) > 31 Or Val(Text5(0)) < 0 Then
        Beep
        x = MsgBox("Tanggal tidak lebih dari 31...", vbOKOnly, "Tanggal Error...!!!")
        Text5(0) = ""
        Text5(0).SetFocus
    End If
Case 1
    If Val(Text5(1)) > 12 Or Val(Text5(1)) < 0 Then
        Beep
        x = MsgBox("Bulan tidak lebih dari 12...", vbOKOnly, "Bulan Error...!!!")
        Text5(1) = ""
        Text5(1).SetFocus
    End If
Case 2

End Select
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub
