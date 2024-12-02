VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form VC05 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABEL BARANG"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "MUTASI PER BARANG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   2767
      TabIndex        =   7
      Top             =   1695
      Width           =   4290
      Begin Crystal.CrystalReport crpt 
         Left            =   90
         Top             =   315
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton Command3 
         Caption         =   "BATAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2220
         TabIndex        =   9
         Top             =   2430
         Width           =   1890
      End
      Begin VB.CommandButton cmdCTK 
         Caption         =   "CETAK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   225
         TabIndex        =   8
         Top             =   2445
         Width           =   1890
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   420
         Left            =   990
         TabIndex        =   10
         Top             =   675
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   49152
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   16777088
         Format          =   58851329
         CurrentDate     =   39286
         MinDate         =   39083
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   240
         Left            =   75
         OleObjectBlob   =   "VC05.frx":0000
         TabIndex        =   11
         Top             =   345
         Width           =   4140
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   240
         Left            =   75
         OleObjectBlob   =   "VC05.frx":0086
         TabIndex        =   12
         Top             =   1335
         Width           =   4140
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   420
         Left            =   990
         TabIndex        =   13
         Top             =   1665
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   49152
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   16777088
         Format          =   58851329
         CurrentDate     =   39286
         MinDate         =   39083
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEMUA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7890
      TabIndex        =   6
      Top             =   105
      Width           =   1890
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CARI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   105
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1575
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   105
      Width           =   4155
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3967
      TabIndex        =   0
      Top             =   5775
      Width           =   1890
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7140
      OleObjectBlob   =   "VC05.frx":010E
      Top             =   4830
   End
   Begin VB.PictureBox Picture1 
      Height          =   1230
      Left            =   -735
      ScaleHeight     =   1170
      ScaleWidth      =   11025
      TabIndex        =   1
      Top             =   5565
      Width           =   11085
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   165
      Left            =   105
      OleObjectBlob   =   "VC05.frx":0342
      TabIndex        =   2
      Top             =   180
      Width           =   1410
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   4860
      Left            =   30
      TabIndex        =   3
      Top             =   525
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   8573
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      AllowUserResizing=   3
   End
End
Attribute VB_Name = "VC05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi, Pusing As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String
Private T, M, D, T2, M2, D2


Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdCTK_Click()
Call Seleksi

Dim tanya
tanya = MsgBox("CETAK TANGGAL", vbOKCancel, "KONFIRMASI")
    If tanya = vbOK Then
        Call CetakBarang
    Else
        Exit Sub
    End If
End Sub

Private Sub CetakBarang()
crpt.ReportFileName = App.Path + "\ReportSELULER\MUTASIBARANG.rpt"
crpt.SelectionFormula = "{B005.KODE_JNS} = '" + Trim(KB) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub Seleksi()
SDel = "Delete * From B005CTK"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing

SCari1 = "Select * From B005 where Tgl_Fak >= datevalue('" + Trim(DTPicker1) + "') and Tgl_Fak <= datevalue('" + Trim(DTPicker2) + "')"
Set RCari1 = RDCO.OpenResultset(SCari1, rdOpenDynamic, rdConcurRowVer)
RCari1.MoveFirst
Do While Not RCari1.EOF

    SCari2 = "Select * From B005CTK"
    Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenKeyset, rdConcurRowVer)
    RCari2.AddNew
        RCari2("Status") = 1
        RCari2("KODE_TRANS") = RCari1("Kode_Trans")
        RCari2("KODE_JNS") = RCari1("KODE_JNS")
        RCari2("NAMA_JNS") = RCari1("NAMA_JNS")
        RCari2("NO_FAKTUR") = RCari1("NO_FAKTUR")
        RCari2("NO_BUKTI") = RCari1("NO_BUKTI")
        RCari2("KETERANGAN") = RCari1("KETERANGAN")
        RCari2("JML_DBT") = RCari1("JML_DBT")
        RCari2("JML_CRD") = RCari1("JML_CRD")
        RCari2("JML_AKHIR") = RCari1("JML_AKHIR")
        RCari2("MUTASI_DBT") = RCari1("MUTASI_DBT")
        RCari2("MUTASI_CRT") = RCari1("MUTASI_CRT")
        RCari2("SALDO_AKHIR") = RCari1("SALDO_AKHIR")
        RCari2("H_POKOK") = RCari1("H_POKOK")
        RCari2("NOMDISC") = RCari1("NOMDISC")
        RCari2("SPCDISC") = RCari1("SPCDISC")
        RCari2("LABA") = RCari1("LABA")
        RCari2("KAS") = RCari1("KAS")
        RCari2("TGL_S") = RCari1("TGL_S")
        RCari2("TGL_FAK") = RCari1("TGL_FAK")
    RCari2.Update
    RCari2.Close
    Set RCari2 = Nothing

RCari1.MoveNext
Loop
RCari1.Close
Set RCari1 = Nothing

End Sub

Private Sub Command1_Click()
grid.Clear
Call SiapkanGrid2
Call IsiGrid2
End Sub

Private Sub Command2_Click()
Call SiapkanGrid
Call IsiGrid
End Sub

Private Sub Command3_Click()
Unload Me
VC05.Show 1
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me
Combo1 = ""

Call SiapkanGrid
Call IsiGrid
grid.Refresh

SSPL = "Select NamaBR From B003"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
If RSPL.RowCount <> 0 Then
    RSPL.MoveFirst
    Do While Not RSPL.EOF
        Combo1.AddItem RSPL("NamaBR")
    RSPL.MoveNext
    Loop
    RSPL.Close
    Set RSPL = Nothing
    Combo1.ListIndex = 0
End If

Frame1.Visible = False
DTPicker1 = Date
DTPicker2 = Date

End Sub

Private Sub TGL()
T = Year(DTPicker1)
M = Month(DTPicker1)
D = Day(DTPicker1)

T2 = Year(DTPicker2)
M2 = Month(DTPicker2)
D2 = Day(DTPicker2)
End Sub

Private Sub SiapkanGrid()
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "BELI": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUAL": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "SALDO": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "H.JUAL": .CellAlignment = 4
End With
End Sub

Private Sub SiapkanGrid2()
grid.Rows = 2
With grid
    .Cols = 6
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2500: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1000: .Text = "BELI": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1000: .Text = "JUAL": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1250: .Text = "SALDO": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 1500: .Text = "H.JUAL": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SKTG = "Select * From B003 order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR"): .CellAlignment = 1
              .Col = 2: .Text = RKTG("JD"): .CellAlignment = 4
              .Col = 3: .Text = RKTG("JC"): .CellAlignment = 4
              .Col = 4: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 5: .Text = Format(RKTG("HJual"), "##,###")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub IsiGrid2()
SKTG = "Select * From B003 where NamaBR = '" + Trim(Combo1) + "'"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGrid
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR"): .CellAlignment = 1
              .Col = 2: .Text = RKTG("JD"): .CellAlignment = 4
              .Col = 3: .Text = RKTG("JC"): .CellAlignment = 4
              .Col = 4: .Text = RKTG("JAkhir"): .CellAlignment = 4
              .Col = 5: .Text = Format(RKTG("HJual"), "##,###")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub grid_dblClick()
KB = ""
KB = (grid.TextMatrix(grid.Row, 0))

Frame1.Visible = True
Frame1.ZOrder
grid.Visible = False

Frame1.Caption = "MUTASI " + KB

End Sub


Private Sub Combo1_Change()
Static ChangeFlag As Boolean
Dim cboText As String
Dim lencboText As Integer
Dim tmpLen As Integer
Dim tmp As Integer

If Not ChangeFlag Then
cboText = Combo1.Text
lencboText = Len(Combo1.Text)
If Not cekKey Then
For tmp = 0 To Combo1.ListCount - 1
If UCase(Left(Combo1.Text, Combo1.SelStart)) = UCase _
(Combo1.List(tmp)) Then
ChangeFlag = True
Combo1.Text = Combo1.List(tmp)
Combo1.SelStart = Len(Combo1.Text)
ChangeFlag = False
cekKey = False
Exit Sub
End If
Next tmp

If lencboText > 0 Then
For tmp = 0 To Combo1.ListCount - 1
If UCase(Left(Combo1.List(tmp), _
lencboText)) = UCase(cboText) Then
tmpLen = lencboText
ChangeFlag = True
Combo1.Text = Combo1.List(tmp)
Combo1.SelStart = tmpLen
Combo1.SelLength = Len(Combo1.List( _
tmp)) - tmpLen
ChangeFlag = False
Exit For
End If
Next tmp
End If
End If
cekKey = False
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack) Then
cekKey = True
End If
End Sub

