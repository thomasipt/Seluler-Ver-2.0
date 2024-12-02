VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "crystl32.ocx"
Begin VB.Form JL001 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI PENJUALAN"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9795
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7680
      TabIndex        =   53
      Text            =   "23"
      Top             =   7605
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
      Left            =   4496
      TabIndex        =   52
      Text            =   "22"
      Top             =   7605
      Width           =   1740
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
      Left            =   1200
      TabIndex        =   51
      Text            =   "21"
      Top             =   7605
      Width           =   1740
   End
   Begin VB.TextBox Text19 
      Alignment       =   1  'Right Justify
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
      Left            =   1800
      TabIndex        =   0
      Text            =   "19"
      Top             =   45
      Width           =   2625
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   90
      TabIndex        =   37
      Text            =   "Text2"
      Top             =   525
      Width           =   4110
   End
   Begin VB.TextBox Text15 
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Text            =   "Text15"
      Top             =   8175
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4095
      TabIndex        =   13
      Text            =   "1,000,000.00"
      Top             =   525
      Width           =   5610
   End
   Begin VB.CommandButton cmdbatal 
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
      Height          =   465
      Left            =   244
      TabIndex        =   10
      Top             =   6795
      Width           =   1725
   End
   Begin VB.CommandButton cmdtutup 
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
      Height          =   465
      Left            =   7826
      TabIndex        =   11
      Top             =   6795
      Width           =   1725
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1575
      OleObjectBlob   =   "JL001.frx":0000
      Top             =   11625
   End
   Begin Crystal.CrystalReport crpt 
      Left            =   375
      Top             =   11850
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   150
      OleObjectBlob   =   "JL001.frx":0234
      TabIndex        =   12
      Top             =   105
      Width           =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   240
      Left            =   105
      OleObjectBlob   =   "JL001.frx":02AA
      TabIndex        =   15
      Top             =   6225
      Width           =   3540
   End
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   -300
      TabIndex        =   29
      Top             =   375
      Width           =   10290
   End
   Begin VB.PictureBox Picture1 
      Height          =   825
      Left            =   -623
      ScaleHeight     =   765
      ScaleWidth      =   10980
      TabIndex        =   49
      Top             =   6615
      Width           =   11040
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   375
      OleObjectBlob   =   "JL001.frx":033E
      TabIndex        =   54
      Top             =   7605
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   240
      Left            =   3555
      OleObjectBlob   =   "JL001.frx":03A6
      TabIndex        =   55
      Top             =   7605
      Width           =   960
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   240
      Left            =   6855
      OleObjectBlob   =   "JL001.frx":0410
      TabIndex        =   56
      Top             =   7605
      Width           =   960
   End
   Begin VB.Frame PENJUALAN 
      Height          =   4815
      Left            =   37
      TabIndex        =   16
      Top             =   1260
      Width           =   9720
      Begin VB.Frame Frame2 
         Caption         =   "DAFTAR BARANG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3750
         Left            =   1905
         TabIndex        =   45
         Top             =   1035
         Width           =   5865
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
            Height          =   195
            Left            =   1440
            TabIndex        =   46
            Top             =   3420
            Width           =   2985
         End
         Begin MSFlexGridLib.MSFlexGrid gridDF 
            Height          =   3075
            Left            =   90
            TabIndex        =   47
            Top             =   270
            Width           =   5685
            _ExtentX        =   10028
            _ExtentY        =   5424
            _Version        =   393216
            Cols            =   1
            FixedCols       =   0
            BackColor       =   16777215
            BackColorFixed  =   65280
            BackColorBkg    =   16777152
            GridColor       =   0
            TextStyle       =   3
            TextStyleFixed  =   3
            Appearance      =   0
         End
      End
      Begin VB.TextBox Text24 
         Height          =   285
         Left            =   255
         MaxLength       =   50
         TabIndex        =   7
         Text            =   "24"
         Top             =   1440
         Width           =   4635
      End
      Begin MSFlexGridLib.MSFlexGrid grid 
         Height          =   2595
         Left            =   195
         TabIndex        =   36
         Top             =   1770
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4577
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   65280
         BackColorBkg    =   16777152
         GridColor       =   0
         TextStyle       =   3
         TextStyleFixed  =   3
      End
      Begin VB.CommandButton cmdBL003 
         Caption         =   "TAMBAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8670
         TabIndex        =   8
         Top             =   1305
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   4245
         TabIndex        =   3
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7725
         TabIndex        =   44
         Text            =   "18"
         Top             =   4425
         Width           =   1500
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   6225
         TabIndex        =   43
         Text            =   "17"
         Top             =   4425
         Width           =   1500
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   3750
         TabIndex        =   42
         Text            =   "8"
         Top             =   4425
         Width           =   465
      End
      Begin VB.TextBox Text7 
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
         Left            =   6345
         TabIndex        =   5
         Text            =   "Text7"
         Top             =   555
         Width           =   2265
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "JL001.frx":0478
         Left            =   1905
         List            =   "JL001.frx":047A
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   300
         Width           =   2310
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   1905
         TabIndex        =   2
         Text            =   "Combo2"
         Top             =   675
         Width           =   2310
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
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
         Left            =   6345
         TabIndex        =   4
         Text            =   "Text6"
         Top             =   180
         Width           =   690
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
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
         Left            =   6345
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   930
         Width           =   690
      End
      Begin VB.TextBox Text14 
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
         Left            =   6345
         TabIndex        =   9
         Text            =   "Text14"
         Top             =   1305
         Width           =   2265
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   165
         Left            =   255
         OleObjectBlob   =   "JL001.frx":047C
         TabIndex        =   30
         Top             =   375
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   165
         Left            =   255
         OleObjectBlob   =   "JL001.frx":04F0
         TabIndex        =   31
         Top             =   750
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   195
         Left            =   5070
         OleObjectBlob   =   "JL001.frx":0564
         TabIndex        =   32
         Top             =   645
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   195
         Left            =   5070
         OleObjectBlob   =   "JL001.frx":05D8
         TabIndex        =   33
         Top             =   270
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   195
         Left            =   5070
         OleObjectBlob   =   "JL001.frx":0642
         TabIndex        =   34
         Top             =   1020
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
         Height          =   195
         Left            =   5070
         OleObjectBlob   =   "JL001.frx":06B8
         TabIndex        =   35
         Top             =   1395
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   165
         Left            =   255
         OleObjectBlob   =   "JL001.frx":0728
         TabIndex        =   57
         Top             =   1125
         Width           =   1560
      End
   End
   Begin VB.Frame BAYAR 
      Height          =   4815
      Left            =   37
      TabIndex        =   17
      Top             =   1275
      Width           =   9720
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Left            =   -3195
         Locked          =   -1  'True
         TabIndex        =   50
         Text            =   "20"
         Top             =   3420
         Width           =   2310
      End
      Begin VB.CommandButton cmdsimpan 
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
         Height          =   465
         Left            =   6510
         TabIndex        =   48
         Top             =   2070
         Width           =   2760
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
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
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "Text10"
         Top             =   1200
         Width           =   2085
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "Text5"
         Top             =   450
         Width           =   2085
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "Text9"
         Top             =   825
         Width           =   2085
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
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
         Left            =   2355
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "Text16"
         Top             =   2100
         Width           =   2310
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
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
         Left            =   6510
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Text13"
         Top             =   1425
         Width           =   2760
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Text12"
         Top             =   1650
         Width           =   1785
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1530
         TabIndex        =   19
         Text            =   "Text11"
         Top             =   1710
         Width           =   735
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
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
         Left            =   6510
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   450
         Width           =   2760
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   195
         Left            =   225
         OleObjectBlob   =   "JL001.frx":079A
         TabIndex        =   22
         Top             =   510
         Width           =   2160
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   240
         Left            =   4830
         OleObjectBlob   =   "JL001.frx":0830
         TabIndex        =   23
         Top             =   510
         Width           =   1560
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   195
         Left            =   225
         OleObjectBlob   =   "JL001.frx":0898
         TabIndex        =   24
         Top             =   885
         Width           =   2160
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
         Height          =   195
         Left            =   225
         OleObjectBlob   =   "JL001.frx":0930
         TabIndex        =   25
         Top             =   1320
         Width           =   2160
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   195
         Left            =   225
         OleObjectBlob   =   "JL001.frx":09C4
         TabIndex        =   26
         Top             =   1770
         Width           =   2760
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
         Height          =   195
         Left            =   225
         OleObjectBlob   =   "JL001.frx":0A5E
         TabIndex        =   27
         Top             =   2220
         Width           =   2160
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   165
         Left            =   4830
         OleObjectBlob   =   "JL001.frx":0AF8
         TabIndex        =   28
         Top             =   1530
         Width           =   1935
      End
   End
End
Attribute VB_Name = "JL001"
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

Private SqlNo As String
Private TTL

Private Sub cmdbatal_Click()
PENJUALAN.Visible = True
BAYAR.Visible = False
Combo1.SetFocus
End Sub

Private Sub cmdsimpan_Click()
Dim tanya
tanya = MsgBox("TRANSAKSI SELESAI", vbCritical, "KONFIRMASI")
    If tanya = vbOK Then
        Call NoBukti2
        Call PERSEDIAAN_BAHAN
        Call SimpanG005
    Else
        MsgBox "CANCEL", vbCritical, "KONFIRMASI"
    End If
    
'NOTAFAK = Trim(Text19)
'Call VALIDASI

Unload Me
JL001.Show 1
End Sub

Private Sub SimpanG005()
SSave = "Select * From G003 where CodeSL = '101001'"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
If RSave.RowCount <> 0 Then
    DEBET = RSave("MutasiD")
    SALDO = RSave("Saldo")
    RSave.Edit
    RSave("MutasiD") = CCur(DEBET) + CCur(Text16)
    RSave("Saldo") = CCur(SALDO) + CCur(Text16)

    SSave2 = "Select * From G005 ORDER BY NOURUT"
    Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    RSave2.AddNew
        RSave2("CodeCab") = CodeCab
        RSave2("CodeSl") = "JUAL"
        RSave2("NamaSl") = "JUAL"
        RSave2("Nobukti") = Trim(Text19)
        RSave2("Keterangan") = "PENJUALAN FAK." + Trim(Text19)
        RSave2("NominalD") = CCur(Text16)
        RSave2("NominalC") = 0
        RSave2("Saldo") = CCur(SALDO) + CCur(Text16)
        RSave2("Laba") = CCur(Text20)
        RSave2("Tanggal") = Date
        RSave2("Jam") = Time
        RSave2("UserCode") = Operator
    RSave2.Update
    RSave2.Close
    Set RSave2 = Nothing
    
    RSave.Update
    RSave.Close
    Set RSave = Nothing
End If
End Sub

Private Sub VALIDASI()
MsgBox "SIAPKAN VALIDASI KE PRINTER"
crpt.ReportFileName = App.Path + "\ReportSELULER\Nota.rpt"
crpt.SelectionFormula = "{B005.No_Bukti} = '" + Trim(NOTAFAK) + "'"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Private Sub DelJL001()
SDel = "Delete * From JL001"
Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
RDEl.Close
Set RDEl = Nothing
End Sub

Private Sub Command1_Click()
Frame2.Visible = True
Frame2.ZOrder
End Sub

Private Sub Command3_Click()
Frame2.Visible = False
Combo1.SetFocus
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

PENJUALAN.Visible = True
BAYAR.Visible = False

Call DelJL001
Call KOSONG
'Call NoBukti
Call IsiCombo
Call IsiText
Call SiapkanGrid
Call IsiGrid
Text15 = 1

Frame2.Visible = False
    Call SiapkanGridDF
    Call IsiGridDF
    
grid.Visible = False

SCari = "Select * From G003 where CodeSL = '101001'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset)
If RCari.RowCount <> 0 Then
    Text21 = Format(RCari("MutasiD"), "##,###")
    Text22 = Format(RCari("MutasiC"), "##,###")
    Text23 = Format(RCari("Saldo"), "##,###")
End If
RCari.Close
Set RCari = Nothing

Text20 = 0

End Sub

Private Sub KOSONG()
ClearTextBoxes JL001
Combo1 = ""
Combo2 = ""
End Sub

Private Sub IsiText()
Text7 = 0
Text1 = 0
Text14 = 0
End Sub

'Private Sub NoBukti()
'Dim No As Double
'SqlNo = "Select * from C013 where nama = '" + Operator + "'"
'Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)
'No = Val(RSLNO("NoBeli")) + 1
'NoStr = Digit(7, No)
'Text19 = "1." + NoStr
'RSLNO.Close
'Set RSLNO = Nothing
'End Sub

Private Sub NoBukti2()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
    TOGEL = RCari9("NoJual")
    RCari9.Edit
        RCari9("NoJual") = TOGEL + 1
RCari9.Update
RCari9.Close
Set RCari9 = Nothing
End Sub

Private Sub SiapkanGridDF()
With gridDF
    .Cols = 2
    .Row = 0
    .Col = 0: .ColWidth(0) = 1500: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2700: .Text = "NAMA": .CellAlignment = 4
End With
End Sub

Private Sub IsiGridDF()
SKTG = "Select * From B003 order by KodeBR Asc"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenKeyset, rdConcurReadOnly)
If RKTG.RowCount <> 0 Then
   Call SiapkanGridDF
   RKTG.MoveFirst
   B = 1
   Do Until RKTG.EOF
      gridDF.Rows = B + 1
      gridDF.Row = B
         With gridDF
              .Col = 0: .Text = RKTG("KodeBR"): .CellAlignment = 4
              .Col = 1: .Text = RKTG("NamaBR")
         End With
      B = B + 1
      RKTG.MoveNext
   Loop
End If
RKTG.Close
Set RKTG = Nothing
End Sub

Private Sub IsiCombo()
SKTG = "Select * From B003 order by kodebr"
Set RKTG = RDCO.OpenResultset(SKTG, rdOpenDynamic, rdOpenKeyset)
RKTG.MoveFirst
Do While Not RKTG.EOF
    Combo1.AddItem RKTG("KodeBR")
RKTG.MoveNext
Loop
RKTG.Close
Set RKTG = Nothing

SSTN = "Select * From B003 order by namabr"
Set RSTN = RDCO.OpenResultset(SSTN, rdOpenDynamic, rdOpenKeyset)
RSTN.MoveFirst
Do While Not RSTN.EOF
    Combo2.AddItem RSTN("NamaBR")
RSTN.MoveNext
Loop
RSTN.Close
Set RSTN = Nothing

End Sub

Private Sub Combo1_GotFocus()
Combo1.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyDelete) Or (KeyCode = vbKeyBack) Then
cekKey = True
End If
Select Case KeyCode
    Case vbKeyF5
        PENJUALAN.Visible = False
        BAYAR.Visible = True
        Text11.SetFocus
End Select
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
Combo1.BackColor = RGB(255, 255, 255)

If Combo1 = "" Then Exit Sub
SCari = "Select * From B003 where KodeBR='" + Trim(Combo1) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
If RCari.RowCount <> 0 Then
    Combo2 = RCari("NamaBR")
    Text7 = Format(RCari("HJUAL"), "##,###")
    Text6.SetFocus
Else
    MsgBox "KODE BARANG BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1 = ""
    Combo2 = ""
    Combo1.SetFocus
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Combo2_GotFocus()
Combo2.BackColor = RGB(255, 255, 0)
End Sub

Private Sub combo2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        PENJUALAN.Visible = False
        BAYAR.Visible = True
        Text11.SetFocus
End Select
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo2_LostFocus()
Combo2.BackColor = RGB(255, 255, 255)

If Combo2 = "" Then Exit Sub
SCari2 = "Select * From B003 where NamaBR='" + Trim(Combo2) + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    Combo1 = RCari2("KodeBR")
    Text7 = Format(RCari2("HJUAL"), "##,###")
    Text6.SetFocus
Else
    MsgBox "NAMA BARANG BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1 = ""
    Combo2 = ""
    Combo1.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
Combo2 = Format(Combo2, ">")
End Sub

Private Sub grid_dblClick()
Dim tanya

KB = (grid.TextMatrix(grid.Row, 1))

    tanya = MsgBox("HAPUS DAFTAR " + KB, vbCritical, "KONFIRMASI")
    If tanya = vbOK Then
        SDel = "Delete * From JL001 where KODE_JNS = '" + Trim(KB) + "'"
        Set RDEl = RDCO.OpenResultset(SDel, rdOpenDynamic, rdConcurRowVer)
        RDEl.Close
        Set RDEl = Nothing
    Else
        Exit Sub
    End If

Call SiapkanGrid
Call IsiGrid
Call IsiText
Call TOTAL
Text6 = ""
Combo1 = ""
Combo2 = ""
Combo1.SetFocus

Call CEKGRID

MsgBox "DATA TELAH DIHAPUS", vbCritical, "KONFIRMASI"

SJual10 = "Select * From JL001A"
Set RJual10 = RDCO.OpenResultset(SJual10, rdOpenKeyset, rdConcurReadOnly)
If RJual10.RowCount > 0 Then
    Text8 = RJual10("SumOfJML_BAHAN")
    Text17 = Format(RJual10("SumOfTOTAL_JUAL"), "##,###")
    Text18 = Format(RJual10("SumOfTOTAL_DISCOUNT"), "##,###")
Else
    Text8 = "0"
    Text17 = "0.00"
    Text18 = "0.00"
End If
RJual10.Close
Set RJual10 = Nothing

End Sub

Private Sub CEKGRID()
SCari = "Select * From JL001"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount = 0 Then
        grid.Visible = False
        Exit Sub
    End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub gridDF_dblClick()
Combo1 = (gridDF.TextMatrix(gridDF.Row, 0))
Combo2 = (gridDF.TextMatrix(gridDF.Row, 1))
Frame2.Visible = False
Call Combo1_LostFocus
Call Combo2_LostFocus
Text6.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        PENJUALAN.Visible = False
        BAYAR.Visible = True
        Text11.SetFocus
End Select
End Sub

Private Sub Text1_GotFocus()
Text1.BackColor = RGB(255, 255, 0)
Text1 = ""
End Sub

Private Sub Text1_Lostfocus()
Text1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text1 = "" Then
        Text1 = 0
        Text14 = Format((CCur((Text6) * (Text7)) - (CCur((Text6) * (Text7) * (Text1) / 100))), "##,###")
    Else
        Text14 = Format((CCur((Text6) * (Text7)) - (CCur((Text6) * (Text7) * (Text1) / 100))), "##,###")
        SendKeys "{TAB}"
    End If
End If
End Sub

Private Sub cmdBL003_GotFocus()
Text14.BackColor = RGB(255, 255, 0)
End Sub

Private Sub cmdBL003_LostFocus()
Text14.BackColor = RGB(255, 255, 255)
End Sub

Private Sub cmdBL003_Click()
grid.Visible = True

SJS = "Select * From B003 where KodeBR = '" + Trim(Combo1) + "'"
Set RJS = RDCO.OpenResultset(SJS, rdOpenKeyset, rdConcurRowVer)

TTL = CCur(RJS("JAkhir"))

If CCur(Text6) > RJS("JAkhir") Then
    MsgBox "JUMLAH STOCK BARANG TERSEDIA " + Trim(RJS("JAkhir")) + " PCS", vbCritical, "KONFIRMASI"
    Text6 = ""
    Text6.SetFocus
    
Else

Dim tanya
    If Combo1 = "" Or Combo2 = "" Or Text6 = "" Or Text7 = "" Or Text1 = "" Or Text14 = ",00" Or Text24 = "" Then
        MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
        Combo1.SetFocus
        Exit Sub
    Else
        tanya = MsgBox("MASUKAN DATA PENJUALAN", vbCritical, "KONFIRMASI")
        If tanya = vbOK Then
            
            Call SimpanJL001
                SJual10 = "Select * From JL001A"
                Set RJual10 = RDCO.OpenResultset(SJual10, rdOpenKeyset, rdConcurReadOnly)
                    Text8 = RJual10("SumOfJML_BAHAN")
                    Text17 = Format(RJual10("SumOfTOTAL_JUAL"), "##,###")
                    Text18 = Format(RJual10("SumOfTOTAL_DISCOUNT"), "##,###")
                RJual10.Close
                Set RJual10 = Nothing
            Call SiapkanGrid
            Call IsiGrid
            Call IsiText
            Call TOTAL
                Text6 = ""
                Text24 = ""
                Combo1 = ""
                Combo2 = ""
                Combo1.SetFocus
            Exit Sub
        End If
    End If
End If

RJS.Close
Set RJS = Nothing

End Sub

Private Sub TOTAL()
SCari3 = "Select * From JL001A"
Set RCari3 = RDCO.OpenResultset(SCari3, rdOpenKeyset, rdConcurReadOnly)
If RCari3.RowCount <> 0 Then
    TOTALP = RCari3("sumofTOTAL_JUAL")
    TOTALD = RCari3("SumOfNOMDISC")
    Text3 = Format(TOTALP, "##,###")
Else
    Text3 = 0
End If

RCari3.Close
Set RCari3 = Nothing
Text2.Text = "TOTAL BAYAR"
End Sub

Private Sub SimpanJL001()
SCari6 = "Select * From JL001 where Kode_JNS = '" + Trim(Combo1) + "'"
Set RCari6 = RDCO.OpenResultset(SCari6, rdOpenDynamic, rdConcurRowVer)
If RCari6.RowCount <> 0 Then
    JUMLAH = RCari6("JML_BAHAN") + CCur(Text6)
    HARGA = CCur(Text7) * JUMLAH
    NNOMDISC = CCur(Text1) / 100 * CCur(Text7)

'CEK JUMLAH BARANG B003 DENGAN JL001

    If CCur(TTL) < CCur(JUMLAH) Then
        MsgBox "JUMLAH STOCK BARANG TERSEDIA " + Trim(RJS("JAkhir")) + " PCS", vbCritical, "KONFIRMASI"
        Text6 = ""
        Text6.SetFocus
        Exit Sub
    End If
    
    RCari6.Edit
    RCari6("JML_BAHAN") = CCur(JUMLAH)
    RCari6("HJUAL_PCS") = CCur(Text7)
    RCari6("HARGA_JUAL") = CCur(HARGA)
    RCari6("DISCOUNT") = CCur(Text1)
    RCari6("NOMINAL") = 0
    RCari6("NOMDISC") = CCur(NNOMDISC)
    RCari6("TOTAL_JUAL") = CCur(Text7) * CCur(JUMLAH)
    RCari6("TOTAL_DISCOUNT") = CCur(HARGA) - CCur(NNOMDISC) * CCur(JUMLAH)
    RCari6.Update
    RCari6.Close
    Set RCari6 = Nothing

Else

    SCari5 = "Select * From B003 where KodeBR = '" + Trim(Combo1) + "'"
    Set RCari5 = RDCO.OpenResultset(SCari5, rdOpenDynamic, rdConcurRowVer)
        INDUK = RCari5("KodeInd")
        HJUAL = RCari5("HJual")
        
        SSave = "Select * From JL001"
        Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
        RSave.AddNew
            RSave("No_Trans") = Text19
            RSave("No_Urut") = Text15
            RSave("Kode_Ind") = INDUK
            RSave("Kode_JNS") = Combo1
            RSave("Nama_JNS") = Combo2
            RSave("KETERANGAN") = Trim(Text24)
            RSave("Jml_BAHAN") = CCur(Text6)
            RSave("HJual_PCS") = CCur(Text7)
            RSave("Harga_JUAL") = HJUAL * CCur(Text6)
            RSave("Nominal") = 0
            RSave("Laba") = 0
            RSave("User_Code") = Operator
            RSave("Discount") = CCur(Text1)
            RSave("NomDisc") = CCur((Text7) * (Text1) / 100)
            RSave("TOTAL_JUAL") = CCur(Text7) * CCur(Text6)
            RSave("TOTAL_DISCOUNT") = (CCur(Text7) - (CCur((Text7) * (Text1) / 100))) * CCur(Text6)
        RSave.Update
        RSave.Close
        Set RSave = Nothing
        Text15 = Text15 + 1
    RCari5.Close
    Set RCari5 = Nothing

End If
End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 9
    .Col = 0: .ColWidth(0) = 500: .Text = "NO": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 2000: .Text = "NAMA BARANG": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 500: .Text = "JML": .CellAlignment = 4
    .Col = 4: .ColWidth(4) = 1500: .Text = "HARGA PCS": .CellAlignment = 4
    .Col = 5: .ColWidth(5) = 500: .Text = "%": .CellAlignment = 4
    .Col = 6: .ColWidth(6) = 1500: .Text = "JUMLAH HARGA": .CellAlignment = 4
    .Col = 7: .ColWidth(7) = 1500: .Text = "HARGA BERSIH": .CellAlignment = 4
    .Col = 8: .ColWidth(8) = 2000: .Text = "KETERANGAN": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari4 = "Select * From JL001"
Set RCari4 = RDCO.OpenResultset(SCari4, rdOpenKeyset, rdConcurReadOnly)
If RCari4.RowCount <> 0 Then
   RCari4.MoveFirst
   B = 1
   Do Until RCari4.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari4("No_Urut"): .CellAlignment = 4
              .Col = 1: .Text = RCari4("Kode_JNS"): .CellAlignment = 4
              .Col = 2: .Text = RCari4("Nama_JNS")
              .Col = 3: .Text = RCari4("Jml_BAHAN"): .CellAlignment = 4
              .Col = 4: .Text = Format(RCari4("HJual_PCS"), "##,###")
              .Col = 5: .Text = RCari4("Discount"): .CellAlignment = 4
              .Col = 6: .Text = Format(RCari4("TOTAL_JUAL"), "##,###")
              .Col = 7: .Text = Format((RCari4("HJUAL_PCS") - RCari4("NomDisc")) * RCari4("Jml_BAHAN"), "##,###"): .CellFontBold = True
              .Col = 8: .Text = Format(RCari4("KETERANGAN"), ">")
         End With
      B = B + 1
      RCari4.MoveNext
   Loop
End If
RCari4.Close
Set RCari4 = Nothing

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Combo1.SetFocus
End If
End Sub

Private Sub Text19_LostFocus()
Text19 = Format(Text19, ">")
Call CekData
End Sub

Private Sub CekData()
If Text19.Text = "" Then Exit Sub

SCari = "Select * From G005 where NOBUKTI = '" + Trim(Text19) + "'"
Set RCari = RDCO.OpenResultset(SCari, rdOpenDynamic, rdConcurRowVer)
    If RCari.RowCount <> 0 Then
        MsgBox " NO FAKTUR NOTA TELAH DIGUNAKAN", vbCritical, "KONFIRMASI"
        Text19 = ""
        Text19.SetFocus
    Exit Sub
    End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBL003.SetFocus
End If
End Sub

Private Sub Text24_LostFocus()
Text24 = Format(Text24, ">")
End Sub

Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        PENJUALAN.Visible = False
        BAYAR.Visible = True
        Text11.SetFocus
End Select
End Sub

Private Sub text6_GotFocus()
Text6.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text6_LostFocus()
Text6.BackColor = RGB(255, 255, 255)
If Not IsNumeric(Text6) Then
    Text6 = 0
    Exit Sub
Else
    If Text6 = "" Or Text6 = 0 Or Text7 = "" Or Text7 = 0 Then Exit Sub
    Text14 = Format(CCur(Text6) * CCur(Text7), "##,###")
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF5
        PENJUALAN.Visible = False
        BAYAR.Visible = True
        Text11.SetFocus
End Select
End Sub

Private Sub Text7_GotFocus()
Text7.BackColor = RGB(255, 255, 0)
End Sub

Private Sub Text7_LostFocus()
Text7.BackColor = RGB(255, 255, 255)
    Text7 = Format(CCur(Text7), "##,###")
    Text14 = Format(CCur(Text6) * CCur(Text7), "##,###")
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Text11_GotFocus()

Text12 = 0
SSave3 = "Select * From JL001A"
Set RSave3 = RDCO.OpenResultset(SSave3, rdOpenKeyset, rdConcurReadOnly)
    Text5 = Format(RSave3("sumofTOTAL_JUAL"), "##,###")
    Text9 = RSave3("SumOfTOTAL_JUAL") - RSave3("SumOfTOTAL_DISCOUNT")
    Text10 = Format(CCur(Text5) - CCur(Text9), "##,###")
    Text9 = Format(Text9, "##,###")
RSave3.Close
Set RSave3 = Nothing
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text11 = "" Then
        Text11 = 0
        Text12 = CCur(Text10) * CCur(Text11 / 100)
        Text16 = CCur(Text10) - CCur(Text12)
    Else
        Text12 = CCur(Text10) * CCur(Text11) / 100
        Text16 = CCur(Text10) - CCur(Text12)
        Text4.SetFocus
    End If
End If
End Sub

Private Sub text4_gotfocus()
Text3 = Format(CCur(Text16), "##,###")
Text2.Text = "TOTAL BAYAR"
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
If KeyAscii = 13 Then
    If Text4 = "" Then
        Text4.SetFocus
    Else
        Text4 = CCur(Text4)
        Text13 = CCur(Text4 - Text16)
        Text2.Text = "KEMBALI :"
        Text3 = CCur(Text13)
        Text13.SetFocus
    End If
End If
End Sub

Private Sub text13_gotfocus()
If Text4 < Text16 Then
    MsgBox "NOMINAL PEMBAYARAN KURANG", vbCritical, "KONFIRMASI"
    Text4.SetFocus
Else
    cmdsimpan.SetFocus
End If
End Sub

Private Sub PERSEDIAAN_BAHAN()
SPCDISC = CCur(Text12) / CCur(Text8)

SJual4 = "Select * From JL001 where NO_TRANS = '" + Trim(Text19) + "' ORDER BY NO_URUT"
Set RJual4 = RDCO.OpenResultset(SJual4, rdOpenKeyset, rdConcurRowVer)
RJual4.MoveFirst
Do While Not RJual4.EOF
    NOURUT = RJual4("NO_URUT")
    KODEJNS = RJual4("KODE_JNS")
    NAMAJNS = RJual4("NAMA_JNS")
    JMLBAHAN = RJual4("JML_BAHAN")
    HJUALPCS = RJual4("HJUAL_PCS")
    HARGAJUAL = RJual4("HARGA_JUAL")
    NNOMDISC = RJual4("NOMDISC") * RJual4("JML_BAHAN")
    
'EDIT MUTASIPRODUKSI BERDASARKAN METODE STOCK'
    SJual5 = "Select * From B003 where KODEBR = '" + Trim(KODEJNS) + "'"
    Set RJual5 = RDCO.OpenResultset(SJual5, rdOpenKeyset, rdConcurRowVer)
        JMLBAHAN1 = CCur(JMLBAHAN)
        NOMINAL = 0
        NOMINAL1 = 0
        HPOKOK = 0
        SDISC = 0
        MTSTOCK = RJual5("MSTOCK")

        SJual6 = "Select * From B004 where KODE_JNS = '" + Trim(KODEJNS) + "' ORDER BY NO_URUT"
        Set RJual6 = RDCO.OpenResultset(SJual6, rdOpenKeyset, rdConcurRowVer)
        RJual6.MoveFirst
        Do While Not RJual6.EOF
            NO4 = RJual6("NO_URUT")
            HBELIPCS = RJual6("HBELI_PCS")
            JMLSALDO = RJual6("JML_SALDO")
            NOMSALDO = RJual6("NOM_SALDO")
            HPPCS = RJual6("HJUAL_PCS")
            
            If JMLBAHAN1 >= JMLSALDO Then
            JMLBAHAN1 = JMLBAHAN1 - JMLSALDO
            NOMINAL1 = NOMINAL1 + NOMSALDO
            HPOKOK = HPOKOK + SALDOHP
            
                RJual6.Edit
                RJual6("JML_SALDO") = 0
                RJual6("NOM_SALDO") = 0
                RJual6.Update
                
            ElseIf JMLBAHAN1 < JMLSALDO And JMLBAHAN1 <> 0 Then
            JMLSALDO = JMLSALDO - JMLBAHAN1
            NOMSALDO = NOMSALDO - (HBELIPCS * JMLBAHAN1)
            NOMINAL1 = NOMINAL1 + (HBELIPCS * JMLBAHAN1)
            SALDOHP = SALDOHP - (HPPCS * JMLBAHAN1)
            HPOKOK = HPOKOK + (HPPCS * JMLBAHAN1)
                
                RJual6.Edit
                RJual6("JML_SALDO") = CCur(JMLSALDO)
                RJual6("NOM_SALDO") = CCur(NOMSALDO)
                RJual6.Update
                
                JMLBAHAN1 = 0
            End If
            
            SSDISC = CCur(SPCDISC) * CCur(JMLBAHAN)
            
            JMLSTOCK = CCur(RJual5("JAKHIR"))
    
                If JMLBAHAN > JMLSTOCK And JMLSTOCK > 0 Then
                JMLBHNLEBIH = CCur(JMLBAHAN) - CCur(JMLSTOCK)
                NOMLEBIH = CCur(HJUALPCS) * CCur(JMLBHNLEBIH)
                NOMINAL1 = CCur(NOMINAL1) + CCur(NOMLEBIH)
                End If
    
            laba = RJual4("HARGA_JUAL") - (NOMINAL1 + NNOMDISC + SSDISC)
            RJual4.Edit
            RJual4("NOMINAL") = CCur(NOMINAL1)
            RJual4("LABA") = CCur(laba)
            Text20 = CCur(Text20) + CCur(laba)
            RJual4.Update
        
        RJual6.MoveNext
        Loop
        RJual6.Close
        Set RJual6 = Nothing
        
                    SJual7 = "Delete * From B004 where JML_SALDO = 0 AND NOM_SALDO < 1"
                    Set RJual7 = RDCO.OpenResultset(SJual7, rdOpenDynamic, rdConcurRowVer)
                    RJual7.Close
                    Set RJual7 = Nothing
    
    JMLCRD = RJual5("JC")
    JMLAKHIR = RJual5("JAKHIR")
    MUTASICRT = RJual5("MUTASIC")
    SALDOAKHIR = RJual5("SALDO")

                If SALDOAKHIR <= 0 Then
                SDISC = SPCDISC * JMLBAHAN
                HARGAJUAL = RJual5("HJUAL")
                NNOMDISC = RJual4("NOMDISC")
                NOMINAL1 = (HARGAJUAL - NOMDISC) - SDISC
                laba = 0
                End If
        
        JMLCRD = JMLCRD + JMLBAHAN
        JMLAKHIR = JMLAKHIR - JMLBAHAN
        MUTASICRT = MUTASICRT + NOMINAL1
        SALDOAKHIR = SALDOAKHIR - NOMINAL1
        
    RJual5.Edit
    RJual5("JC") = CCur(JMLCRD)
    RJual5("JAKHIR") = CCur(JMLAKHIR)
    RJual5("MUTASIC") = CCur(MUTASICRT)
    RJual5("SALDO") = CCur(SALDOAKHIR)
        
'UPDATE HISTORY TRANSAKSI BAHAN BAKU'
                    SJual8 = "Select * From B005 ORDER BY NO_URUT"
                    Set RJual8 = RDCO.OpenResultset(SJual8, rdOpenKeyset, rdConcurRowVer)
                    RJual8.AddNew
                        RJual8("Status") = 1
                        RJual8("KODE_TRANS") = "JL"
                        RJual8("KODE_JNS") = KODEJNS
                        RJual8("NAMA_JNS") = NAMAJNS
                        RJual8("NO_FAKTUR") = Text19
                        RJual8("NO_BUKTI") = Text19
                        RJual8("KETERANGAN") = "PENJUALAN NO." + RJual4("NO_TRANS") + " ( " + RJual4("KETERANGAN") + " ) "
                        RJual8("JML_DBT") = 0
                        RJual8("JML_CRD") = JMLBAHAN
                        RJual8("JML_AKHIR") = JMLAKHIR
                        RJual8("MUTASI_DBT") = 0
                        RJual8("MUTASI_CRT") = NOMINAL1
                        RJual8("SALDO_AKHIR") = SALDOAKHIR
                        RJual8("H_POKOK") = HARGAJUAL
                        RJual8("NOMDISC") = NNOMDISC
                        RJual8("SPCDISC") = SDISC
                        RJual8("LABA") = laba
                        RJual8("KAS") = 0
                        RJual8("TGL_S") = Date
                        RJual8("TGL_FAK") = Date
                    RJual8.Update
                    RJual8.Close
                    Set RJual8 = Nothing
   
    RJual5.Update
    RJual5.Close
    Set RJual5 = Nothing

RJual4.MoveNext
Loop
RJual4.Close
Set RJual4 = Nothing
        
End Sub


Private cekKey As Boolean

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



