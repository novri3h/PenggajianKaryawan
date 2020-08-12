VERSION 5.00
Begin VB.Form SLIPGAJI 
   Caption         =   "Cetak  Slip Gaji"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2910
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
   ScaleHeight     =   1035
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   345
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " NIP"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1000
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Bulan"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1000
   End
End
Attribute VB_Name = "SLIPGAJI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub COMBO2_Keypress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
Call BukaDB

RSGAJI.Open "SELECT DISTINCT BULAN FROM GAJI", Conn
Combo1.Clear
Do While Not RSGAJI.EOF
    Combo1.AddItem RSGAJI!BULAN
    RSGAJI.MoveNext
Loop

RSPegawai.Open "SELECT * FROM PEGAWAI ORDER BY NIP", Conn
Combo2.Clear
Do While Not RSPegawai.EOF
    Combo2.AddItem RSPegawai!NIP
    RSPegawai.MoveNext
Loop
End Sub


Private Sub Combo2_Click()
Call BukaDB
Dim RSSLIP As New ADODB.Recordset
RSSLIP.Open "SELECT PEGAWAI.NIP,NAMA,JABATAN.NMJABATAN AS JABATAN,PEGAWAI.GOL,STATUS,JMLANAK,MASUK,MASTER.LEMBUR,GAPOK,JABATAN.TJJABATAN,IIF(PEGAWAI.STATUS='MENIKAH',TJSUAMIISTRI,0)  AS TJKELUARGA,IIF (STATUS='MENIKAH',TJANAK*JMLANAK,0)  AS TJANAK,UMAKAN*MASUK AS UANGMAKAN,MASTER.LEMBUR*GOLONGAN.LEMBUR AS UANGLEMBUR,ASKES, (GAPOK+TJJABATAN+TJKELUARGA+TJANAK+UANGMAKAN+ASKES+UANGLEMBUR) AS PENDAPATAN,  POTONGAN,PENDAPATAN-POTONGAN AS TOTALGAJI  FROM PEGAWAI,GOLONGAN,JABATAN,MASTER WHERE PEGAWAI.GOL=GOLONGAN.GOL AND PEGAWAI.KOJAB=JABATAN.KOJAB AND PEGAWAI.NIP=MASTER.NIP AND PEGAWAI.NIP='" & Combo2 & "'", Conn
If RSSLIP.EOF Then
    MsgBox "DATA TIDAK DITEMUKAN"
Else
    Call CetakGAJI
End If
End Sub

Function CetakGAJI()
Call BukaDB
Dim RSGAJI As New ADODB.Recordset
RSGAJI.Open "SELECT PEGAWAI.NIP,NAMA,JABATAN.NMJABATAN AS JABATAN,PEGAWAI.GOL,STATUS,JMLANAK,MASUK,MASTER.LEMBUR,GAPOK,JABATAN.TJJABATAN,IIF(PEGAWAI.STATUS='MENIKAH',TJSUAMIISTRI,0)  AS TJKELUARGA,IIF (STATUS='MENIKAH',TJANAK*JMLANAK,0)  AS TJANAK,UMAKAN*MASUK AS UANGMAKAN,MASTER.LEMBUR*GOLONGAN.LEMBUR AS UANGLEMBUR,ASKES, (GAPOK+TJJABATAN+TJKELUARGA+TJANAK+UANGMAKAN+ASKES+UANGLEMBUR) AS PENDAPATAN,  POTONGAN,PENDAPATAN-POTONGAN AS TOTALGAJI  FROM PEGAWAI,GOLONGAN,JABATAN,MASTER WHERE PEGAWAI.GOL=GOLONGAN.GOL AND PEGAWAI.KOJAB=JABATAN.KOJAB AND PEGAWAI.NIP=MASTER.NIP AND PEGAWAI.NIP='" & Combo2 & "'", Conn
LAYAR.Show
Dim MGrs As String
MGrs = String$(40, "-")
LAYAR.Font = "Courier New"
LAYAR.Print
LAYAR.Print
'LAYAR.Print Tab(5); MGrs
LAYAR.Print Tab(5); "TANGGAL                :   "; Format(Date, "DD-MMM-YYYY")
LAYAR.Print
LAYAR.Print Tab(5); "NIP                    :   "; RSGAJI!NIP
LAYAR.Print Tab(5); "NAMA PEGAWAI           :   "; RSGAJI!NAMA
LAYAR.Print Tab(5); "JABATAN                :   "; RSGAJI!JABATAN
LAYAR.Print Tab(5); "GOLONGAN               :   "; RSGAJI!GOL
LAYAR.Print Tab(5); "STATUS                 :   "; RSGAJI!Status
LAYAR.Print Tab(5); "JUMLAH ANAK            :  "; RSGAJI!JMLANAK
LAYAR.Print Tab(5); "JUMLAH HARI KERJA      :  "; RSGAJI!masuk
LAYAR.Print Tab(5); "JUMLAH JAM LEMBUR      :  "; RSGAJI!LEMBUR
LAYAR.Print Tab(5); MGrs
LAYAR.Print Tab(5); "GAJI POKOK             :   "; RKanan(RSGAJI!GAPOK, "##,###,###")
LAYAR.Print Tab(5); "TUNJANGAN JABATAN      :   "; RKanan(RSGAJI!TJJABATAN, "##,###,###")
LAYAR.Print Tab(5); "TUNJANGAN SUAMI/ISTRI  :   ";


If RSGAJI!TJKELUARGA = 0 Then
    LAYAR.Print Tab(40); RSGAJI!TJKELUARGA
Else
    LAYAR.Print Tab(32); RKanan(RSGAJI!TJKELUARGA, "##,###,###")
End If

LAYAR.Print Tab(5); "TUNJANGAN ANAK         :   ";
If RSGAJI!TJANAK = 0 Then
    LAYAR.Print Tab(40); RSGAJI!TJANAK
Else
    LAYAR.Print Tab(32); RKanan(RSGAJI!TJANAK, "##,###,###")
End If

LAYAR.Print Tab(5); "UANG MAKAN             :   "; RKanan(RSGAJI!UANGMAKAN, "##,###,###")

LAYAR.Print Tab(5); "UANG LEMBUR            :   ";
If RSGAJI!UANGLEMBUR = 0 Then
    LAYAR.Print Tab(40); RSGAJI!UANGLEMBUR
Else
    LAYAR.Print Tab(32); RKanan(RSGAJI!UANGLEMBUR, "##,###,###");
End If

LAYAR.Print Tab(5); "ASKES                  :   "; RKanan(RSGAJI!ASKES, "##,###,###")
LAYAR.Print Tab(5); MGrs
LAYAR.Print Tab(5); "TOTAL PENDAPATAN       :   "; RKanan(RSGAJI!PENDAPATAN, "##,###,###")
LAYAR.Print Tab(5); MGrs

LAYAR.Print Tab(5); "POTONGAN               :   ";
If RSGAJI!POTONGAN = 0 Then
    LAYAR.Print Tab(40); RSGAJI!POTONGAN
Else
    LAYAR.Print Tab(32); RKanan(RSGAJI!POTONGAN, "##,###,###")
End If

LAYAR.Print Tab(5); MGrs
LAYAR.Print Tab(5); "TOTAL GAJI             :   "; RKanan(RSGAJI!TotalGAJI, "##,###,###")
LAYAR.Print Tab(5); MGrs
LAYAR.Print
LAYAR.Print
Conn.Close
End Function

Private Function RKanan(NData, CFormat) As String
    RKanan = Format(NData, CFormat)
    RKanan = Space(Len(CFormat) - Len(RKanan)) + RKanan
End Function

