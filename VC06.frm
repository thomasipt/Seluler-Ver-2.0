VERSION 5.00
Begin VB.Form VC06 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORMASI"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "VC06"
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
End Sub
