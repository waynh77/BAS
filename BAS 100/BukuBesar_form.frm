VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form BukuBesar_form 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buku Besar"
   ClientHeight    =   6630
   ClientLeft      =   1515
   ClientTop       =   1170
   ClientWidth     =   6795
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KELUAR"
      Height          =   495
      Index           =   1
      Left            =   4080
      TabIndex        =   17
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EDIT"
      Height          =   495
      Index           =   0
      Left            =   5520
      TabIndex        =   16
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000007&
      DataField       =   " "
      DataSource      =   " "
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   315
      ItemData        =   "BukuBesar_form.frx":0000
      Left            =   1920
      List            =   "BukuBesar_form.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000007&
      DataField       =   "saldo_akun"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000007&
      DataField       =   " "
      DataSource      =   " "
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      DataField       =   " "
      DataSource      =   " "
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000007&
      DataField       =   " "
      DataSource      =   " "
      Enabled         =   0   'False
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Width           =   4695
   End
   Begin VB.Data Data2 
      Caption         =   "tabel_akun"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "|< First"
      Height          =   495
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "< Prev"
      Height          =   495
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CETAK"
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Next >"
      Height          =   495
      Index           =   3
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Last >|"
      Height          =   495
      Index           =   4
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2760
      Width           =   855
   End
   Begin VB.Data Data1 
      Caption         =   "saldo_akun"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "BukuBesar_form.frx":0004
      Height          =   3135
      Left            =   120
      OleObjectBlob   =   "BukuBesar_form.frx":0018
      TabIndex        =   5
      Top             =   3360
      Width           =   6495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SALDO AKUN "
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
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "NAMA AKUN"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "KLASIFIKASI"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "JENIS AKUN"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "SANDI AKUN "
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
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
End
Attribute VB_Name = "BukuBesar_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tutup_navigate()
Command2(0).Enabled = False
Command2(1).Enabled = False
Command2(3).Enabled = False
Command2(4).Enabled = False
End Sub

Private Sub buka_navigate()
Command2(0).Enabled = True
Command2(1).Enabled = True
Command2(3).Enabled = True
Command2(4).Enabled = True
End Sub


Private Sub Combo1_Click()
With Data3.Recordset
.MoveFirst
Do While Not .EOF Or .BOF
    If Combo1 = !sandi_akun Then
'        Combo1 = !sandi_akun
        Text1 = !Jenis_akun
        Text2 = !klasifikasi_akun
        Text3 = !nama_akun
        Text4 = !saldo_akun
        .MoveLast
'        .MovePrevious
    End If
    .MoveNext
Loop
'Data3.RecordSource = "select saldo_akun.sandi_akun,nama_akun,Jenis_akun,Klasifikasi_akun,saldo_akun from saldo_akun,tabel_akun where saldo_akun.sandi_akun= 'nsb" & Combo1 & "'" 'tabel_akun.sandi_akun order by saldo_akun.sandi_akun asc"
Data3.Refresh
End With
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    If Command1(0).Caption = "EDIT" Then
        Text4.Enabled = True
        Text4.SetFocus
        Command1(0).Caption = "SIMPAN"
        Command2(2).Caption = "BATAL"
        tutup_navigate
    ElseIf Command1(0).Caption = "SIMPAN" Then
        If Text4 = "" Then
            x = MsgBox("Text masih kosong", vbOKOnly, "Text Kosong")
            Text4.SetFocus
        Else
            With Data1.Recordset
                .MoveFirst
                Do While Not .EOF
                    If !sandi_akun = Combo1 Then
                        .Edit
                        !sandi_akun = Combo1
                        !saldo_akun = Text4
                        .Update
                        Data1.Refresh
                        Data3.Refresh
                        .MoveLast
                    End If
                    .MoveNext
                Loop
            End With
            Data1.Refresh
            Data3.Refresh
            Command1(0).Caption = "EDIT"
            Command2(2).Caption = "CETAK"
            Text4.Enabled = False
            buka_navigate
        End If
    End If
Case 1
    Me.Hide
    main_form.Enabled = True
    main_form.Show
End Select
End Sub

Private Sub DBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    isi
End Sub

Private Sub Form_Activate()
If Data1.Recordset.BOF Then
    With Data2.Recordset
    .MoveFirst
    Do While Not .EOF
        Data1.Recordset.AddNew
        Data1.Recordset!sandi_akun = !sandi_akun
        Data1.Recordset!saldo_akun = 0
        Data1.Recordset.Update
        Data1.Refresh
        .MoveNext
    Loop
    End With
End If
'Data1.RecordSource = "select * from dbsaldo_akun" ' order by sandi_akun"
Data1.Refresh
Data3.RecordSource = "select dbsaldo_akun.sandi_akun,nama_akun,saldo_akun,Jenis_akun,Klasifikasi_akun from dbsaldo_akun,tabel_akun where dbsaldo_akun.sandi_akun=tabel_akun.sandi_akun order by dbsaldo_akun.sandi_akun asc"
Data3.Refresh
isi_combo
Data3.Refresh
End Sub

Private Sub isi_combo()
Combo1.Clear
Data2.Recordset.MoveFirst
Do While Not Data2.Recordset.EOF
    Combo1.AddItem (Data2.Recordset!sandi_akun)
    Data2.Recordset.MoveNext
Loop
Data2.Refresh
Combo1.ListIndex = (0)
End Sub
Private Sub Command2_Click(Index As Integer)
If Not Data3.Recordset.BOF Then
    Select Case Index
        Case 0
            Data1.Recordset.MoveFirst
            Data3.Recordset.MoveFirst
            isi
        Case 1
            Data1.Recordset.MovePrevious
            Data3.Recordset.MovePrevious
            If Data3.Recordset.BOF Then
                x = MsgBox("Anda sudah di data pertama", vbOKOnly, "Data Pertama")
                Data3.Recordset.MoveFirst
                Data1.Recordset.MoveFirst
            End If
            isi
        Case 2
            If Command2(2).Caption = "CETAK" Then
                CrystalReport1.WindowState = crptMaximized
                CrystalReport1.RetrieveDataFiles
                CrystalReport1.Action = 1
            Else
                buka_navigate
                Text4.Enabled = False
                Command1(0).Caption = "EDIT"
                Command2(2).Caption = "CETAK"
            End If
        Case 3
            Data1.Recordset.MoveNext
            Data3.Recordset.MoveNext
            If Data3.Recordset.EOF Or Data3.Recordset.BOF Then
                x = MsgBox("Anda sudah berada di data terakhir", vbOKOnly, "Data Terakhir")
                Data3.Recordset.MoveLast
                Data1.Recordset.MoveLast
            End If
            isi
        Case 4
            Data3.Recordset.MoveLast
            Data1.Recordset.MoveLast
            isi
    End Select
Else
    Command2(0).Enabled = False
    Command2(1).Enabled = False
    Command2(2).Enabled = False
    Command2(3).Enabled = False
    Command2(4).Enabled = False
End If

End Sub

Private Sub isi()
With Data3.Recordset
'.MoveFirst
'Do While Not .EOF Or .BOF
'    If Combo1 = !sandi_akun Then
        Combo1 = !sandi_akun
        Text1 = !Jenis_akun
        Text2 = !klasifikasi_akun
        Text3 = !nama_akun
        Text4 = !saldo_akun
'.MoveLast
'        .MovePrevious
'    End If
'    .MoveNext
'Loop
End With
End Sub

Private Sub Form_Load()
Call Buka_DBBukuBesar
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
main_form.Enabled = True
main_form.Show
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
        Beep
        KeyAscii = 0
    End If
End Sub
