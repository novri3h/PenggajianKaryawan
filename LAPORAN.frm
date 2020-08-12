VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form LAPORAN 
   Caption         =   "LAPORAN TRANSAKSI"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form12"
   ScaleHeight     =   6135
   ScaleWidth      =   3510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Data Penggajian"
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   4560
      Width           =   3000
      Begin VB.ComboBox Combo8 
         Height          =   345
         Left            =   1320
         TabIndex        =   17
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo7 
         Height          =   345
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1250
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Absen"
      Height          =   1335
      Left            =   240
      TabIndex        =   12
      Top             =   240
      Width           =   3000
      Begin VB.ComboBox Combo2 
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1250
      End
   End
   Begin Crystal.CrystalReport CR 
      Left            =   1440
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Caption         =   "Data Potongan"
      Height          =   1335
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   3000
      Begin VB.ComboBox Combo5 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   1500
      End
      Begin VB.ComboBox Combo6 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1250
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1250
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Data Lembur"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   3000
      Begin VB.ComboBox Combo4 
         Height          =   345
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   1500
      End
      Begin VB.ComboBox Combo3 
         Height          =   345
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Bulan"
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1250
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Tahun"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1250
      End
   End
End
Attribute VB_Name = "LAPORAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'pada saat form dipanggil
Private Sub Form_Load()
'definisikan variabel string TGLkomputer sebagai Bulan
Dim TglKomputer As Date
'ambil Bulan komputer
TglKomputer = Date
'jika formatnya bukan Bulan/bulan/tahun maka (ini format taggal indonesia)
If BULAN <> Format(Date, "DD/MM/YY") Then
    'ubah dulu formatnya ke tgl/bulan/tahun
    BULAN = Format(Date, "DD/MM/YY")
End If

'buka database
Call BukaDB
'definisikan recordset (tabel) baru
Dim RSTGL As New ADODB.Recordset
'cari Bulan di tabel Master dan tampilkan angka bulannya saja di combo1
RSTGL.Open "select distinct month(Bulan) as Bulan from Master", Conn
Do While Not RSTGL.EOF
    Combo1.AddItem RSTGL!BULAN & Space(5) & MonthName(RSTGL!BULAN)
    Combo3.AddItem RSTGL!BULAN & Space(5) & MonthName(RSTGL!BULAN)
    Combo5.AddItem RSTGL!BULAN & Space(5) & MonthName(RSTGL!BULAN)
    Combo7.AddItem RSTGL!BULAN & Space(5) & MonthName(RSTGL!BULAN)
    RSTGL.MoveNext
Loop
'tutup database
Conn.Close

'buka lagi database
Call BukaDB
Dim RSTHN As New ADODB.Recordset
'cari Bulan di tabel Master.
'yang ditampilkan hanya angka tahunnya saja
RSTHN.Open "select distinct year(Bulan)  as Tahun from Master", Conn
Do While Not RSTHN.EOF
    Combo2.AddItem RSTHN!TAHUN
    Combo4.AddItem RSTHN!TAHUN
    Combo6.AddItem RSTHN!TAHUN
    Combo8.AddItem RSTHN!TAHUN
    RSTHN.MoveNext
Loop
Conn.Close

End Sub

'Lap Harian
'CR adalah nama objek Crystal Report yang telah diubah
'nama asli objek saat dibuat dalam form adalah CrystalReport1
'digantinya nama CrystalReport1 jadi CR hanya untuk memperpendek penulisan objek saja



'Lap Bulanan
Private Sub Combo2_Click()
    'buka database
    Call BukaDB
    'cari data yang cocok dengan pilihan di combo1 dan 5 (bulan dan tahun)
    RSMASTER.Open "select * from Master where month(Bulan)='" & Val(Left(Combo1, 2)) & "' and year(Bulan)='" & (Combo2) & "'", Conn
    'jika data tidak ditemukan, laporan tidak ditampilkan dulu
    If RSMASTER.EOF Then
        'tampilkan pesan bahwa data tidak ditemukan
        MsgBox "Data tidak ditemukan"
        'kembali ke combo1
        Combo1.SetFocus
        'program dibawahnya dihentikan (tidak dibaca)
        Exit Sub
    End If
    'saring laporan yang bulannya dipilih di combo1 dan tahunnya dipilih di combo2
    CR.SelectionFormula = "Month({Master.Bulan})=" & Val(Left(Combo1, 2)) & " and Year({Master.Bulan})=" & Val(Combo2.Text)
    'panggil laporan
    CR.ReportFileName = App.Path & "\lap absen.rpt"
    'tampilkan satu layar penuh
    CR.WindowState = crptMaximized
    'lakukan updating jika tabel telah berubah
    CR.RetrieveDataFiles
    'tampilkan ke layar
    CR.Action = 1
End Sub


'Lap Bulanan
Private Sub combo4_Click()
    Call BukaDB
    RSMASTER.Open "select * from Master where month(Bulan)='" & Val(Left(Combo3, 2)) & "' and year(Bulan)='" & (Combo4) & "'", Conn
    If RSMASTER.EOF Then
        MsgBox "Data tidak ditemukan"
        Combo3.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "Month({Master.Bulan})=" & Val(Left(Combo3, 2)) & " and Year({Master.Bulan})=" & Val(Combo4.Text)
    CR.ReportFileName = App.Path & "\lap lembur.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub



Private Sub COMBO6_Click()
    Call BukaDB
    RSMASTER.Open "select * from Master where month(Bulan)='" & Val(Left(Combo5, 2)) & "' and year(Bulan)='" & (Combo6) & "'", Conn
    If RSMASTER.EOF Then
        MsgBox "Data tidak ditemukan"
        Combo5.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "Month({Master.Bulan})=" & Val(Left(Combo5, 2)) & " and Year({Master.Bulan})=" & Val(Combo6.Text)
    CR.ReportFileName = App.Path & "\lap POTONGAN.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub


Private Sub COMBO8_Click()
    Call BukaDB
    RSMASTER.Open "select * from gaji where month(Bulan)='" & Val(Left(Combo7, 2)) & "' and year(Bulan)='" & (Combo8) & "'", Conn
    If RSMASTER.EOF Then
        MsgBox "Data tidak ditemukan"
        Combo7.SetFocus
        Exit Sub
    End If
    CR.SelectionFormula = "Month({gaji.Bulan})=" & Val(Left(Combo7, 2)) & " and Year({gaji.Bulan})=" & Val(Combo8.Text)
    CR.ReportFileName = App.Path & "\lap gaji.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

