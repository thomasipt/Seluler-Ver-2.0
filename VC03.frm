VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VC03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENJUALAN PULSA ELETRONIK"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4680
      TabIndex        =   27
      Text            =   "8"
      Top             =   6345
      Width           =   1545
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2115
      TabIndex        =   2
      Text            =   "7"
      Top             =   1710
      Width           =   2580
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3060
      TabIndex        =   25
      Text            =   "3"
      Top             =   6300
      Width           =   1545
   End
   Begin VB.TextBox Text21 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   900
      TabIndex        =   21
      Text            =   "21"
      Top             =   5220
      Width           =   1740
   End
   Begin VB.TextBox Text22 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3660
      TabIndex        =   20
      Text            =   "22"
      Top             =   5220
      Width           =   1740
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6300
      TabIndex        =   19
      Text            =   "23"
      Top             =   5220
      Width           =   1740
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   6
      Text            =   "2"
      Top             =   3825
      Width           =   5820
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   5
      Text            =   "1"
      Top             =   3420
      Width           =   2130
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   0
      Text            =   "6"
      Top             =   90
      Width           =   1545
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   4
      Text            =   "5"
      Top             =   3030
      Width           =   5820
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   3
      Text            =   "4"
      Top             =   2640
      Width           =   5820
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
      Left            =   5494
      TabIndex        =   8
      Top             =   4410
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2085
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   540
      Width           =   2040
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4725
      OleObjectBlob   =   "VC03.frx":0000
      Top             =   4590
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":0234
      TabIndex        =   9
      Top             =   600
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   3135
      OleObjectBlob   =   "VC03.frx":029A
      TabIndex        =   10
      Top             =   1035
      Width           =   1965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SIMPAN"
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
      Left            =   731
      TabIndex        =   7
      Top             =   4410
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   720
      Left            =   -315
      ScaleHeight     =   660
      ScaleWidth      =   9270
      TabIndex        =   11
      Top             =   4320
      Width           =   9330
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   225
      Left            =   6150
      OleObjectBlob   =   "VC03.frx":0300
      TabIndex        =   12
      Top             =   1035
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":0368
      TabIndex        =   13
      Top             =   2685
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":03D6
      TabIndex        =   14
      Top             =   3090
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":0440
      TabIndex        =   15
      Top             =   150
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":04A6
      TabIndex        =   16
      Top             =   3480
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":0510
      TabIndex        =   17
      Top             =   3885
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   225
      Left            =   3150
      OleObjectBlob   =   "VC03.frx":0582
      TabIndex        =   18
      Top             =   1350
      Width           =   1965
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   240
      Left            =   75
      OleObjectBlob   =   "VC03.frx":05E8
      TabIndex        =   22
      Top             =   5220
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   240
      Left            =   2715
      OleObjectBlob   =   "VC03.frx":0650
      TabIndex        =   23
      Top             =   5220
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   240
      Left            =   5475
      OleObjectBlob   =   "VC03.frx":06BA
      TabIndex        =   24
      Top             =   5220
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   240
      Left            =   210
      OleObjectBlob   =   "VC03.frx":0722
      TabIndex        =   26
      Top             =   1770
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   225
      Left            =   2205
      OleObjectBlob   =   "VC03.frx":0794
      TabIndex        =   28
      Top             =   1035
      Width           =   885
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   225
      Left            =   2205
      OleObjectBlob   =   "VC03.frx":07FA
      TabIndex        =   29
      Top             =   1350
      Width           =   885
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   225
      Left            =   5175
      OleObjectBlob   =   "VC03.frx":0862
      TabIndex        =   30
      Top             =   1035
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   -360
      TabIndex        =   31
      Top             =   855
      Width           =   9015
   End
End
Attribute VB_Name = "VC03"
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


Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Or Text4 = "" Or Text5 = "" Or Text6 = "" Then
    MsgBox "MASIH ADA DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Dim tanya
tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI PENJUALAN", vbSystemModal, "KONFIRMASI")
If tanya = vbOK Then
    Call Simpan
End If

Unload Me
VC03.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC01 where INDUK = '" + Trim(KB) + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave("SALDO") < CCur(SkinLabel6) Then
    MsgBox "SALDO SUDAH HABIS / TIDAK CUKUP", vbCritical, "TRANSAKSI GAGAL"
    Combo1.SetFocus
    Exit Sub
Else
    RSave.Edit
        RSave("CREDIT") = RSave("CREDIT") + CCur(SkinLabel6)
        RSave("SALDO") = RSave("SALDO") - CCur(SkinLabel6)
        RSave("TANGGAL") = Date
    
        SSave2 = "Select * From VC03"
        Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
        RSave2.AddNew
            RSave2("NOTA") = Trim(Text6)
            RSave2("INDUK") = Trim(KB)
            RSave2("KODE") = Trim(Combo1)
            RSave2("NAMA") = Trim(SkinLabel3)
            RSave2("CREDIT") = CCur(SkinLabel6)
            RSave2("SALDO") = RSave("SALDO")
            RSave2("LABA") = CCur(Text8)
            RSave2("HJUAL") = CCur(Text7)
            RSave2("HBELI") = RSave("HARGA_BELI")
            RSave2("CUSTOMER") = Trim(Text4)
            RSave2("ALAMAT") = Trim(Text5)
            RSave2("HP") = Trim(Text1)
            RSave2("KETERANGAN") = Trim(Text2)
            RSave2("TANGGAL") = Date
                
            SSave4 = "Select * From G003 where CodeSL = '101001'"
            Set RSave4 = RDCO.OpenResultset(SSave4, rdOpenDynamic, rdConcurRowVer)
            If RSave4.RowCount <> 0 Then
                DEBET = RSave4("MutasiD")
                SALDO = RSave4("Saldo")
                RSave4.Edit
                RSave4("MutasiD") = CCur(DEBET) + CCur(SkinLabel7)
                RSave4("Saldo") = CCur(SALDO) + CCur(SkinLabel7)
            
                    SSave3 = "Select * From G005 ORDER BY NOURUT"
                    Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenKeyset, rdConcurRowVer)
                    RSave3.AddNew
                        RSave3("CodeCab") = CodeCab
                        RSave3("CodeSl") = "VOUCHER"
                        RSave3("NamaSl") = "VOUCHER"
                        RSave3("Nobukti") = Trim(Text6)
                        RSave3("Keterangan") = "JL.VC. " + Trim(Text1)
                        RSave3("NominalD") = CCur(SkinLabel7)
                        RSave3("NominalC") = 0
                        RSave3("Saldo") = CCur(SALDO) + CCur(SkinLabel7)
                        RSave3("Laba") = CCur(Text8)
                        RSave3("Tanggal") = Date
                        RSave3("Jam") = Time
                        RSave3("UserCode") = Operator
                    RSave3.Update
                    RSave3.Close
                    Set RSave3 = Nothing
                
                    SJual8 = "Select * From B005 ORDER BY NO_URUT"
                    Set RJual8 = RDCO.OpenResultset(SJual8, rdOpenKeyset, rdConcurRowVer)
                    RJual8.AddNew
                        RJual8("Status") = 1
                        RJual8("KODE_TRANS") = "JL"
                        RJual8("KODE_JNS") = Trim(Combo1)
                        RJual8("NAMA_JNS") = Trim(SkinLabel3)
                        RJual8("NO_FAKTUR") = Text6
                        RJual8("NO_BUKTI") = Text6
                        RJual8("KETERANGAN") = "JL.VC.NO. " + Trim(Text6)
                        RJual8("JML_DBT") = 0
                        RJual8("JML_CRD") = CCur(SkinLabel6)
                        RJual8("JML_AKHIR") = RSave("SALDO")
                        RJual8("MUTASI_DBT") = 0
                        RJual8("MUTASI_CRT") = CCur(SkinLabel7)
                        RJual8("SALDO_AKHIR") = CCur(SALDO) + CCur(SkinLabel7)
                        RJual8("H_POKOK") = CCur(SkinLabel7)
                        RJual8("NOMDISC") = 0
                        RJual8("SPCDISC") = 0
                        RJual8("LABA") = CCur(Text8)
                        RJual8("KAS") = 0
                        RJual8("TGL_S") = Date
                        RJual8("TGL_FAK") = Date
                    RJual8.Update
                    RJual8.Close
                    Set RJual8 = Nothing
                
                RSave4.Update
                RSave4.Close
                Set RSave4 = Nothing
            End If
                    
        RSave2.Update
        RSave2.Close
        Set RSave2 = Nothing
        
    RSave.Update
    RSave.Close
    Set RSave = Nothing
End If
End Sub

Private Sub EditVC01()
Dim Stock As String
Dim HBeli As String

SSave2 = "Select * From VC01 where Kode = '" + Trim(Combo1) + "'"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    Stock = RSave2("Stokbel")
    HBeli = RSave2("Satuan")
RSave2.Edit
    RSave2("Stokbel") = CCur(Stock) - (CCur(Text1) * CCur(HBeli))
    RSave2("Jumlah") = CCur(Pusing) - CCur(Text1)
RSave2.Update
RSave2.Close
Set RSave2 = Nothing
End Sub

Private Sub EditNoBukti()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
    TOGEL = RCari9("NoJual")
    RCari9.Edit
        RCari9("NoJual") = TOGEL + 1
    RCari9.Update
    RCari9.Close
    Set RCari9 = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
SCari2 = "Select * From VC02 where KODE = '" + Combo1 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel3 = RCari2("NAMA")
    SkinLabel7 = Format(RCari2("HARGA"), "##,###")
    Text7 = Format(RCari2("HARGA"), "##,###")
    SkinLabel6 = Format(RCari2("UNIT"), "##,###")
    KB = RCari2("INDUK")
    Text8 = Format(RCari2("LABA"), "##,###")

    SCari3 = "Select * From VC00 where INDUK = '" + KB + "' order by NO Desc"
    Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenDynamic, rdConcurRowVer)
    If RCari3.RowCount <> 0 Then
        RCari3.MoveFirst
        Do While Not RCari3.EOF
            Text3 = CCur(SkinLabel6) * RCari3("PERSEN") / 100
        RCari3.MoveNext
        Loop
        RCari3.Close
        Set RCari3 = Nothing
    End If
    
    'If CCur(SkinLabel7) > CCur(SkinLabel6) Then
    '    Text3 = CCur(SkinLabel7) - CCur(SkinLabel6)
    'ElseIf CCur(SkinLabel7) < CCur(SkinLabel6) Then
    '    Text3 = CCur(SkinLabel6) - CCur(SkinLabel7)
    'ElseIf CCur(SkinLabel7) < CCur(SkinLabel6) Then
    '    Text3 = 0
    'End If
Else
    MsgBox "KODE INDUK BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

Call CekFIFO

Combo1 = ""

SSPL = "Select KODE From VC02 order by KODE"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
If RSPL.RowCount <> 0 Then
    RSPL.MoveFirst
    Do While Not RSPL.EOF
        Combo1.AddItem RSPL("KODE")
    RSPL.MoveNext
    Loop
    RSPL.Close
    Set RSPL = Nothing
    Combo1.ListIndex = 0
End If

SkinLabel3 = ""
SkinLabel7 = ""
SkinLabel6 = ""

SCari = "Select * From G003 where CodeSL = '101001'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Text21 = Format(RCari("MutasiD"), "##,###")
    Text22 = Format(RCari("MutasiC"), "##,###")
    Text23 = Format(RCari("Saldo"), "##,###")
End If
RCari.Close
Set RCari = Nothing

End Sub

Private Sub CekFIFO()
SCari6 = "Select * From VC00 order by NO Desc"
Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenKeyset)
RCari6.MoveFirst
Do While Not RCari6.EOF
    NNO = RCari6("NO")
    SSALDO = CCur(RCari6("SALDO"))
    SSISA = CCur(RCari6("SISA"))
    
    If CCur(SSALDO) = CCur(SSISA) Then
        SDel = "Delete * From VC00 where SALDO = SISA"
        Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
        RDEl.Close
        Set RDEl = Nothing
    End If
RCari6.MoveNext
Loop
RCari6.Close
Set RCari6 = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text2_LostFocus()
    Text2 = Format(Text2, ">")
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text4_LostFocus()
    Text4 = Format(Text4, ">")
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text5_LostFocus()
    Text5 = Format(Text5, ">")
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text6_LostFocus()
Text6 = Format(Text6, ">")
Call CekData
End Sub

Private Sub CekData()
If Text6.Text = "" Then Exit Sub

SCari = "Select * From G005 where NOBUKTI = '" + Trim(Text6) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox "NO NOTA TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text6 = ""
        Text6.SetFocus
    Exit Sub
    End If

RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text7_LostFocus()
    Text7 = Format(CCur(Text7), "##,###")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
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

