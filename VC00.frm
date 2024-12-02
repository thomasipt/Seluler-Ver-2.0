VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VC00 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DEPOSIT PULSA"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3735
      TabIndex        =   10
      Text            =   "3"
      Top             =   1965
      Width           =   870
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1942
      TabIndex        =   2
      Text            =   "2"
      Top             =   1470
      Width           =   3075
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
      Left            =   198
      TabIndex        =   3
      Top             =   2730
      Width           =   1890
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
      Left            =   3116
      TabIndex        =   4
      Top             =   2730
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1942
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1942
      TabIndex        =   1
      Text            =   "1"
      Top             =   945
      Width           =   3075
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   187
      OleObjectBlob   =   "VC00.frx":0000
      TabIndex        =   5
      Top             =   165
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   225
      Left            =   187
      OleObjectBlob   =   "VC00.frx":0076
      TabIndex        =   6
      Top             =   990
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   300
      Left            =   1942
      OleObjectBlob   =   "VC00.frx":00EC
      TabIndex        =   7
      Top             =   540
      Width           =   3075
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6750
      OleObjectBlob   =   "VC00.frx":015E
      Top             =   810
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   -90
      ScaleHeight     =   1515
      ScaleWidth      =   8145
      TabIndex        =   8
      Top             =   2565
      Width           =   8205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   180
      OleObjectBlob   =   "VC00.frx":0392
      TabIndex        =   9
      Top             =   1515
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   270
      Left            =   1942
      OleObjectBlob   =   "VC00.frx":0404
      TabIndex        =   11
      Top             =   1980
      Width           =   3075
   End
End
Attribute VB_Name = "VC00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

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

Private RLB As rdoResultset
Private SLB As String

Private SqlNo As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA KOSONG", vbCritical, "KONFIRMASI"
    Exit Sub
End If

Dim tanya
tanya = MsgBox("ANDA YAKIN MELAKUKAN TRANSAKSI DEPOSIT PULSA", vbSystemModal, "KONFIRMASI")
If tanya = vbOK Then
    Text3 = CCur(Text2) / CCur(Text1) * 100
    Call Simpan
    Call Simpan2
End If

Unload Me
VC00.Show 1

End Sub

Private Sub Simpan()
SSave = "Select * From VC01 where INDUK = '" + Combo1 + "'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.Edit
    RSave("DEBET") = RSave("DEBET") + CCur(Text1)
    RSave("SALDO") = RSave("SALDO") + CCur(Text1)
    RSave("TANGGAL") = Date
    RSave("PERSEN") = Trim(Text3)
    RSave("HARGA_BELI") = CCur(Text2)

    SSave2 = "Select * From VC03"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenDynamic, rdConcurRowVer)
    RSave2.AddNew
        RSave2("NOTA") = "DEPOSIT"
        RSave2("INDUK") = Trim(Combo1)
        RSave2("KODE") = Trim(Combo1)
        RSave2("NAMA") = Trim(SkinLabel5)
        RSave2("DEBET") = CCur(Text1)
        RSave2("SALDO") = RSave("SALDO")
        RSave2("TANGGAL") = Date
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing
            
        SJual8 = "Select * From B005 ORDER BY NO_URUT"
        Set RJual8 = RDCO.OpenResultset(SJual8, rdOpenKeyset, rdConcurRowVer)
        RJual8.AddNew
            RJual8("Status") = 1
            RJual8("KODE_TRANS") = "BL"
            RJual8("KODE_JNS") = Trim(Combo1)
            RJual8("NAMA_JNS") = Trim(SkinLabel5)
            RJual8("NO_FAKTUR") = "--"
            RJual8("NO_BUKTI") = "--"
            RJual8("KETERANGAN") = "DEPOSIT"
            RJual8("JML_DBT") = CCur(Text1)
            RJual8("JML_CRD") = 0
            RJual8("JML_AKHIR") = RSave("SALDO")
            RJual8("MUTASI_DBT") = CCur(Text2)
            RJual8("MUTASI_CRT") = 0
            RJual8("SALDO_AKHIR") = CCur(SALDO) + CCur(Text2)
            RJual8("H_POKOK") = CCur(Text2)
            RJual8("NOMDISC") = 0
            RJual8("SPCDISC") = 0
            RJual8("LABA") = CCur(Text3)
            RJual8("KAS") = 0
            RJual8("TGL_S") = Date
            RJual8("TGL_FAK") = Date
        RJual8.Update
        RJual8.Close
        Set RJual8 = Nothing
                    
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Simpan2()
SSave3 = "Select * From G003 where CODESL = '101001'"
Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
If RSave3("SALDO") <= 0 Then
    MsgBox "SALDO KAS Rp 0.00", vbCritical, "KONFIRMASI"
    Exit Sub
Else
        CREDIT = RSave3("MutasiC") + CCur(Text2)
        SALDO = RSave3("Saldo") - CCur(Text2)

            SJual8 = "Select * From G005 ORDER BY NOURUT"
            Set RJual8 = RDCO.OpenResultset(SJual8, rdOpenKeyset, rdConcurRowVer)
            RJual8.AddNew
                RJual8("CodeCab") = CodeCab
                RJual8("CodeSl") = Trim(Combo1)
                RJual8("NamaSl") = SkinLabel5
                RJual8("Nobukti") = Trim(Text1)
                RJual8("Keterangan") = "DEPOSIT PULSA"
                RJual8("NominalD") = 0
                RJual8("NominalC") = CCur(Text2)
                RJual8("Saldo") = SALDO
                RJual8("Tanggal") = Date
                RJual8("Jam") = Time
                RJual8("UserCode") = Operator
            RJual8.Update
            RJual8.Close
            Set RJual8 = Nothing
        
    RSave3.Edit
        RSave3("MutasiC") = CREDIT
        RSave3("Saldo") = SALDO
    
    RSave3.Update
    RSave3.Close
    Set RSave3 = Nothing
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SCari2 = "Select * From VC01 where INDUK = '" + Combo1 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel5 = RCari2("NAMA_INDUK")
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

Combo1 = ""

SSPL = "Select INDUK From VC01 order by INDUK"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo1.AddItem RSPL("INDUK")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing

SkinLabel5 = ""

SSave3 = "Select * From G003 where CODESL = '101001'"
Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenDynamic, rdConcurRowVer)
If RSave3("SALDO") <= 0 Then
    MsgBox "SALDO KAS Rp 0.00", vbCritical, "KONFIRMASI"
    Text1.Enabled = False
    Text2.Enabled = False
    Picture1.ZOrder
    cmdCLOSE.ZOrder
End If
RSave3.Close
Set RSave3 = Nothing

Text1 = 0
Text2 = 0


End Sub

Private Sub Text1_GotFocus()
Text1 = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_Lostfocus()
If Text1 = "" Then
    Text1 = 0
    Exit Sub
End If
If Text1 <> 0 Then
    Text3 = CCur(Text2) / CCur(Text1) * 100
End If
Text1 = Format(Text1, "##,###")
End Sub

Private Sub Text2_GotFocus()
Text2 = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
If Text2 = "" Then
    Text2 = 0
    Exit Sub
End If
If Text2 <> 0 Then
    Text3 = CCur(Text2) / CCur(Text1) * 100
End If
    Text2 = Format(Text2, "##,###")
End Sub


