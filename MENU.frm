VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form MENU 
   Caption         =   "PROGRAM PENGGAJIAN KARYAWAN"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "MENU.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4275
      Width           =   5640
      _ExtentX        =   9948
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "03/08/20"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "11:29"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   3480
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5AD7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5B097
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5B3B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5B6CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5B9E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MENU.frx":5BCFF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MNFILE 
      Caption         =   "FILE"
      Begin VB.Menu MNPEGAWAI 
         Caption         =   "PEGAWAI"
      End
      Begin VB.Menu MNGOL 
         Caption         =   "GOLONGAN"
      End
      Begin VB.Menu MNJABATAN 
         Caption         =   "JABATAN"
      End
   End
   Begin VB.Menu MNTRANSAKSI 
      Caption         =   "TRANSAKSI"
      Begin VB.Menu MNABSEN 
         Caption         =   "ABSEN CARA PERTAMA"
      End
      Begin VB.Menu MNENTRIDATA 
         Caption         =   "ABSEN CARA KEDUA"
      End
   End
   Begin VB.Menu MNHITUNG 
      Caption         =   "PERHITUNGAN"
   End
   Begin VB.Menu MNLAPORAN 
      Caption         =   "LAPORAN"
      Begin VB.Menu MNLAPPEGAWAI 
         Caption         =   "DATA PEGAWAI"
      End
      Begin VB.Menu MNLAPGOLONGAN 
         Caption         =   "DATA GOLONGAN"
      End
      Begin VB.Menu MNLAPJABATAN 
         Caption         =   "DATA JABATAN"
      End
      Begin VB.Menu MNLAPTRANSAKSI 
         Caption         =   "DATA TRANSAKSI"
      End
      Begin VB.Menu MNSLIP 
         Caption         =   "SLIP GAJI"
      End
   End
   Begin VB.Menu MNKELUAR 
      Caption         =   "KELUAR"
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub


Private Sub MNABSEN_Click()
ENTRIDATA.Show vbModal
End Sub

Private Sub MNENTRIDATA_Click()
ABSEN1.Show vbModal
End Sub

Private Sub MNGOL_Click()
TABELGOLONGAN.Show vbModal
End Sub

Private Sub MNHITUNG_Click()
HITUNGGAJI.Show
End Sub

Private Sub MNJABATAN_Click()
TABELJABATAN.Show vbModal
End Sub

Private Sub MNKELUAR_Click()
End
End Sub

Private Sub MNLAPABSEN_Click()
    CR.ReportFileName = App.Path & "\Lap ABSEN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPGOLONGAN_Click()
    CR.ReportFileName = App.Path & "\Lap GOLONGAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPJABATAN_Click()
    CR.ReportFileName = App.Path & "\Lap JABATAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPLEMBUR_Click()
    CR.ReportFileName = App.Path & "\Lap LEMBUR.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPPEGAWAI_Click()
    CR.ReportFileName = App.Path & "\Lap PEGAWAI.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPPOTONGAN_Click()
    CR.ReportFileName = App.Path & "\Lap POTONGAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub MNLAPTRANSAKSI_Click()
LAPORAN.Show vbModal
End Sub

Private Sub MNPEGAWAI_Click()
PEGAWAI.Show vbModal
End Sub

Private Sub MNSLIP_Click()
SLIPGAJI.Show
End Sub

Private Sub MNUJISQL_Click()
UjiSQL.Show vbModal
End Sub
