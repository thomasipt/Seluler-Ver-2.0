VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LOGIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   525
      OleObjectBlob   =   "LOGIN.frx":0000
      Top             =   3990
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
      Left            =   2175
      TabIndex        =   3
      Top             =   1350
      Width           =   1890
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "MASUK"
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
      Left            =   150
      TabIndex        =   2
      Top             =   1350
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   1057
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   765
      Width           =   2370
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1057
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   2370
   End
End
Attribute VB_Name = "LOGIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RDOE As rdoEnvironment
Private RDCO As rdoConnection

Private SqlPass As String
Private tUser As rdoResultset
Private tMasuk As rdoResultset

Private RTgl, RHapus, RDEl, RSave2, RSave3, RSave4, RCari, RCari2, RSLNO, rscs3 As rdoResultset
Private STgl, SHapus, SDel, SSave2, SSave3, SSave4, SCari, SCari2, SqlNo, sqlcs3, Kode As String

Private Sub cmdCLOSE_Click()
MsgBox "TERIMA KASIH ANDA TELAH MENGGUNAKAN PROGRAM KAMI", vbCritical, "ADI JAYA SARANA"
End
End Sub

Private Sub Masuk2()
SCari = "Select * From C013 where UserCode = '" + Text1 + "' and Password = '" + Text2 + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Call Masuk
    Unload Me
Else
    LOGIN.Hide
    MsgBox "ANDA TIDAK BERHAK LOG IN KE SYSTEM", vbCritical, "KONFIRMASI"
    LOGIN.Show
    Text1 = ""
    Text2 = ""
    Text1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Masuk()
SqlPass = "Select * from C013 where UserCode =  '" + Trim(Text1) + "' "
Set tMasuk = RDCO.OpenResultset(SqlPass, rdOpenDynamic, rdConcurRowVer)
If tMasuk.RowCount <> 0 Then
    If tMasuk("MAIN") = "01" Then
        Operator = Trim(tMasuk("Nama"))
        CodeCab = Trim(tMasuk("CodeCab"))
        If TglS <> Date Then
            Call G004
        End If
        MAINMENU.Show
    ElseIf tMasuk("MAIN") = "02" Then
        Operator = Trim(tMasuk("Nama"))
        CodeCab = Trim(tMasuk("CodeCab"))
        If TglS <> Date Then
            Call G004
        End If
        MAINKASIR.Show
    End If
End If
tMasuk.Close
Set tMasuk = Nothing
End Sub

Private Sub G004()
SSave = "Select * From G003 where CodeSL = '101001'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave.RowCount <> 0 Then
    
    SSave2 = "Select * From G004"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    RSave2.AddNew
        RSave2("codecab") = RSave("CODECAB")
        RSave2("codesgl") = RSave("CODESGL")
        RSave2("codesl") = RSave("CODESL")
        RSave2("namasl") = RSave("NAMASL")
        RSave2("SaldoAwal") = RSave("SALDO")
        RSave2("mutasid") = RSave("MUTASID")
        RSave2("mutasic") = RSave("MUTASIC")
        RSave2("saldo") = RSave("SALDO")
        RSave2("posisi") = RSave("POSISI")
        RSave2("Tanggal") = TglS
        RSave2("Jam") = Time
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing

        SSave3 = "Select Tanggal From A001"
        Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenKeyset, rdConcurRowVer)
        RSave3.Edit
            RSave3("tanggal") = Date
        RSave3.Update
        RSave3.Close
        Set RSave3 = Nothing
        
    SSave4 = "Select * From G005 ORDER BY NOURUT"
    Set RSave4 = RDCO.OpenResultset(SSave4, rdOpenKeyset, rdConcurRowVer)
    RSave4.AddNew
        RSave4("CodeCab") = CodeCab
        RSave4("CodeSl") = "101001"
        RSave4("NamaSl") = "KAS"
        RSave4("Nobukti") = 0
        RSave4("Keterangan") = "PINDAH BUKU"
        RSave4("NominalD") = 0
        RSave4("NominalC") = 0
        RSave4("Saldo") = RSave("SALDO")
        RSave4("Laba") = 0
        RSave4("Tanggal") = Date
        RSave4("Jam") = Time
        RSave4("UserCode") = "SYSTEM"
    RSave4.Update
    RSave4.Close
    Set RSave4 = Nothing
    
    RSave.Edit
        RSave("SaldoAwal") = RSave("SALDO")
        RSave("MUTASID") = 0
        RSave("MUTASIC") = 0
    RSave.Update
    
End If
RSave.Close
Set RSave = Nothing
End Sub

Private Sub cmdOK_Click()
'Call Trial
Call Masuk2
End Sub

Private Sub Trial()
Dim Da, Mo, Ye As Integer
Dim M
Da = Day(Date)
Mo = Month(Date)
Ye = Year(Date)
If Mo > 7 Or Ye <> 2008 Then
M = MsgBox("MAAF VERSI TRIAL TELAH HABIS", vbOKCancel + vbCritical, "THANK'S ADI JAYA SARANA")
    If M = vbOK Then
        MsgBox "HUBUNGI ADI JAYA SARANA 024 7673 9586", vbInformation + vbOKOnly, "ADI JAYA SARANA"
            SDel = "Delete * From C013"
            Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
            RDEl.Close
            Set RDEl = Nothing
        Unload Me
    Else
        Unload Me
    End If
Exit Sub
End If
End Sub

Private Sub TGL()
STgl = "Select * from A001"
Set RTgl = RDCO.OpenResultset(STgl, rdOpenDynamic, rdConcurRowVer)
If RTgl.RowCount <> 0 Then
    TglS = RTgl("Tanggal")
    Skin = RTgl("S")
    NTOKO = RTgl("TOKO")
    NAlamat = RTgl("ALamat")
    NMOtto = RTgl("Motto")
    NTelepon = RTgl("Telepon")
Else
End If
RTgl.Close
Set RTgl = Nothing
End Sub

Private Sub Form_Load()
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)
Text1 = ""
Text2 = ""
Call TGL
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_Lostfocus()
Text1 = Format(Text1, ">")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, ">")
End Sub



